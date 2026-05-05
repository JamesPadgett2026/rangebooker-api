// RangeBooker API
// Version: 2026-05-04 EVENTS API ADDED
// File: src/functions/GetLocations.js

const { app } = require("@azure/functions");

const API_VERSION = "2026-05-04 EVENTS API ADDED";

async function getAccessToken() {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    const response = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({
                client_id: clientId,
                client_secret: clientSecret,
                scope: "https://graph.microsoft.com/.default",
                grant_type: "client_credentials"
            })
        }
    );

    const data = await response.json();

    if (!response.ok) {
        throw new Error("Token request failed: " + JSON.stringify(data));
    }

    return data.access_token;
}

async function getRangeBookerSite(token) {
    const res = await fetch(
        "https://graph.microsoft.com/v1.0/sites/tropicaltech.sharepoint.com:/sites/RangeBooker",
        {
            headers: { Authorization: `Bearer ${token}` }
        }
    );

    const data = await res.json();

    if (!res.ok) {
        throw new Error("Site lookup failed: " + JSON.stringify(data));
    }

    return data;
}

async function getListItems(token, siteId, listName) {
    const res = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}/items?expand=fields&$top=5000`,
        {
            headers: { Authorization: `Bearer ${token}` }
        }
    );

    const data = await res.json();

    if (!res.ok) {
        throw new Error(`${listName} lookup failed: ` + JSON.stringify(data));
    }

    return data.value || [];
}

function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
}

function splitPhone(phone) {
    const digits = String(phone || "").replace(/\D/g, "");

    return {
        areaCode: digits.length >= 3 ? digits.substring(0, 3) : "",
        phone3: digits.length >= 6 ? digits.substring(3, 6) : "",
        phone4: digits.length >= 10 ? digits.substring(6, 10) : ""
    };
}

function isDuplicateEmailError(data) {
    const message = String(data?.error?.message || "").toLowerCase();

    return (
        message.includes("unique constraints") ||
        message.includes("duplicate") ||
        message.includes("already has the provided value")
    );
}

function buildImageDataUrl(base64Value) {
    const cleanBase64 = String(base64Value || "")
        .replace(/^data:image\/[a-zA-Z0-9.+-]+;base64,/, "")
        .replace(/^data:image;application\/octet-stream;base64,/, "")
        .trim();

    if (!cleanBase64) {
        return "";
    }

    let mimeType = "image/png";

    if (cleanBase64.startsWith("/9j/")) {
        mimeType = "image/jpeg";
    }

    return `data:${mimeType};base64,${cleanBase64}`;
}

function getImageUrlFromGraphImageColumn(imageValue) {
    if (!imageValue) {
        return "";
    }

    try {
        let img = imageValue;

        if (typeof imageValue === "string") {
            img = JSON.parse(imageValue);
        }

        if (img.serverUrl && img.serverRelativeUrl) {
            return img.serverUrl + img.serverRelativeUrl;
        }

        return img.url || img.Url || "";
    } catch {
        return "";
    }
}

function formatEventDate(value) {
    if (!value) {
        return "";
    }

    const date = new Date(value);

    if (isNaN(date.getTime())) {
        return String(value);
    }

    return date.toLocaleDateString("en-US", {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric"
    });
}

async function getMemberItems(token, siteId) {
    return await getListItems(token, siteId, "MemberListSP");
}

async function getPendingRequestItems(token, siteId) {
    return await getListItems(token, siteId, "PendingRequestsListSP");
}

async function updateMemberLastLogin(token, siteId, memberId, lastLoginValue) {
    const res = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/MemberListSP/items/${memberId}/fields`,
        {
            method: "PATCH",
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                LastLogin: lastLoginValue
            })
        }
    );

    if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error("LastLogin update failed: " + JSON.stringify(data));
    }
}

//
// GET LOCATIONS / CALENDAR
//
app.http("GetLocations", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`GetLocations called. Version: ${API_VERSION}`);

        try {
            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            const items = await getListItems(token, site.id, "CalendarNewSPList");

            const locations = items.map((item, index) => {
                const f = item.fields || {};

                return {
                    id: item.id,
                    name: f.Title || `Item ${index + 1}`,
                    dateTimeToSchedule: f.DateTimeToSchedule || "",
                    startTime: f.StartTimeTextColSP || "",
                    endTime: f.EndTimeTextColSP || "",
                    startTimeId: f.StartTimeTextIDColSP || "",
                    endTimeId: f.EndTimeTextIDColSP || "",
                    availableOrBooked: f.AvailableOrBooked || "",
                    displayText:
                        `${f.StartTimeTextColSP || ""} - ${f.EndTimeTextColSP || ""}: ${f.AvailableOrBooked || ""}`
                };
            });

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    locations
                }
            };

        } catch (err) {
            return {
                status: 500,
                jsonBody: {
                    success: false,
                    version: API_VERSION,
                    error: err.message
                }
            };
        }
    }
});

//
// GET EVENTS (BASE64 ONLY)
//
app.http("GetEvents", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`GetEvents called. Version: ${API_VERSION}`);

        try {
            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            const events = await getListItems(token, site.id, "EventListMaster");
            const base64Photos = await getListItems(token, site.id, "EventPhotosBase64");

            const results = events.map(item => {
                const f = item.fields || {};
                const eventId = Number(f.ID || item.id);

                // 🔥 GET ALL PHOTOS FOR EVENT
                const photos = base64Photos.filter(photo => {
                    const pf = photo.fields || {};
                    return Number(pf.EventLockInIDColSP || 0) === eventId;
                });

                // 🔥 BUILD IMAGE ARRAY
                const images = photos
                    .map(p => buildImageDataUrl(p.fields.Base64ColSP))
                    .filter(x => x);

                const eventDate =
                    f.EventDate ||
                    f.WhenCreated ||
                    "";

                return {
                    id: eventId,

                    title:
                        f.EventNameColSP ||
                        f.Title ||
                        "Untitled Event",

                    description:
                        f.Description ||
                        "",

                    eventDate,
                    eventDateText: formatEventDate(eventDate),

                    createdDate:
                        f.WhenCreated ||
                        "",

                    createdDateText:
                        formatEventDate(f.WhenCreated || ""),

                    createdByName:
                        f.CreatedName ||
                        "",

                    // 🔥 NEW
                    images: images,

                    // 🔥 FIRST IMAGE FOR PREVIEW
                    imageDataUrl: images.length > 0 ? images[0] : ""
                };
            });

            results.sort((a, b) => {
                const dateA = new Date(a.eventDate || "2100-01-01");
                const dateB = new Date(b.eventDate || "2100-01-01");
                return dateA - dateB;
            });

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    count: results.length,
                    events: results
                }
            };

        } catch (err) {
            context.error(err);

            return {
                status: 500,
                jsonBody: {
                    success: false,
                    version: API_VERSION,
                    error: err.message
                }
            };
        }
    }
});
//
// REGISTER MEMBER
//
app.http("RegisterMember", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`RegisterMember called. Version: ${API_VERSION}`);

        if (request.method === "GET") {
            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "RegisterMember API is reachable."
                }
            };
        }

        try {
            const body = await request.json();

            const firstName = String(body.firstName || "").trim();
            const lastName = String(body.lastName || "").trim();
            const email = normalizeEmail(body.email);
            const phone = String(body.phone || "").trim();
            const password = String(body.password || "");
            const notes = String(body.notes || "").trim();

            if (!firstName || !lastName || !email || !password) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "First name, last name, email, and password are required."
                    }
                };
            }

            const nowIso = new Date().toISOString();
            const phoneParts = splitPhone(phone);
            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            const fields = {
                Title: `${firstName} ${lastName}`,
                FirstNameColSP: firstName,
                LastNameColSP: lastName,
                email: email,
                loginemail: email,
                PasswordColSP: password,
                AreaCodeColSP: phoneParts.areaCode ? Number(phoneParts.areaCode) : 0,
                Phone3ColSP: phoneParts.phone3 ? Number(phoneParts.phone3) : 0,
                Phone4ColSP: phoneParts.phone4 ? Number(phoneParts.phone4) : 0,
                MemberType: 1,
                Active: "Yes",
                MembershipRequestApproved: "No",
                DateJoined: nowIso,
                LastLogin: nowIso,
                Notes: notes || ""
            };

            const res = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/MemberListSP/items`,
                {
                    method: "POST",
                    headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({ fields })
                }
            );

            const data = await res.json();

            if (!res.ok) {
                if (isDuplicateEmailError(data)) {
                    return {
                        status: 409,
                        jsonBody: {
                            success: false,
                            version: API_VERSION,
                            error: "An account with this email already exists."
                        }
                    };
                }

                throw new Error("Member create failed: " + JSON.stringify(data));
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Member created in SharePoint. Account is pending approval.",
                    itemId: data.id
                }
            };

        } catch (err) {
            return {
                status: 500,
                jsonBody: {
                    success: false,
                    version: API_VERSION,
                    error: err.message
                }
            };
        }
    }
});

//
// LOGIN MEMBER
//
app.http("LoginMember", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`LoginMember called. Version: ${API_VERSION}`);

        if (request.method === "GET") {
            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "LoginMember API is reachable."
                }
            };
        }

        try {
            const body = await request.json();

            const email = normalizeEmail(body.email);
            const password = String(body.password || "");

            if (!email || !password) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Email and password are required."
                    }
                };
            }

            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);
            const members = await getMemberItems(token, site.id);

            const member = members.find(item => {
                const f = item.fields || {};

                return (
                    normalizeEmail(f.email) === email ||
                    normalizeEmail(f.loginemail) === email ||
                    normalizeEmail(f.EmailColSP) === email
                );
            });

            if (!member) {
                return {
                    status: 401,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Invalid email or password."
                    }
                };
            }

            const f = member.fields || {};
            const savedPassword = String(f.PasswordColSP || "");
            const approvedValue = String(f.MembershipRequestApproved || "").trim().toLowerCase();
            const activeValue = String(f.Active || "").trim().toLowerCase();

            if (savedPassword !== password) {
                return {
                    status: 401,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Invalid email or password."
                    }
                };
            }

            if (approvedValue !== "yes") {
                return {
                    status: 403,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Your account has not been approved yet."
                    }
                };
            }

            if (
                activeValue &&
                activeValue !== "yes" &&
                activeValue !== "true" &&
                activeValue !== "active"
            ) {
                return {
                    status: 403,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "This account is not active yet."
                    }
                };
            }

            const lastLoginNow = new Date().toISOString();

            await updateMemberLastLogin(token, site.id, member.id, lastLoginNow);

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Login successful.",
                    member: {
                        id: member.id,
                        firstName: f.FirstNameColSP || "",
                        lastName: f.LastNameColSP || "",
                        email: email,
                        title: f.Title || "",
                        memberType: f.MemberType || 1,
                        membershipRequestApproved: f.MembershipRequestApproved || "",
                        dateJoined: f.DateJoined || "",
                        lastLogin: lastLoginNow
                    }
                }
            };

        } catch (err) {
            return {
                status: 500,
                jsonBody: {
                    success: false,
                    version: API_VERSION,
                    error: err.message
                }
            };
        }
    }
});

//
// GET MY REQUESTS
//
app.http("GetMyRequests", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`GetMyRequests called. Version: ${API_VERSION}`);

        try {
            const url = new URL(request.url);
            const memberId = Number(url.searchParams.get("memberId") || 0);
            const email = normalizeEmail(url.searchParams.get("email"));

            if (!memberId && !email) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Email or memberId is required."
                    }
                };
            }

            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);
            const items = await getPendingRequestItems(token, site.id);

            const requests = items
                .filter(item => {
                    const f = item.fields || {};
                    const requestMemberId = Number(f.MemberIDLOckInColSP || 0);

                    return Boolean(memberId && requestMemberId === memberId);
                })
                .map(item => {
                    const f = item.fields || {};

                    return {
                        id: item.id,
                        dateId: f.DateIDLOckInColSP || "",
                        requestedDate: f.DateActualColSP || "",
                        status: f.Approved || "Requesting",
                        userLevel: f.UserLevelColSP || "",
                        requestedAt: f.DateRequestWasAdded || ""
                    };
                })
                .sort((a, b) => {
                    const levelA = Number(a.userLevel || 0);
                    const levelB = Number(b.userLevel || 0);

                    if (levelA !== levelB) {
                        return levelA - levelB;
                    }

                    return new Date(a.requestedAt) - new Date(b.requestedAt);
                });

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    requests
                }
            };

        } catch (err) {
            return {
                status: 500,
                jsonBody: {
                    success: false,
                    version: API_VERSION,
                    error: err.message
                }
            };
        }
    }
});

//
// REQUEST BOOKING
//
app.http("RequestBooking", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`RequestBooking called. Version: ${API_VERSION}`);

        if (request.method === "GET") {
            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "RequestBooking API is reachable."
                }
            };
        }

        try {
            const body = await request.json();

            const memberId = Number(body.memberId || 0);
            const userLevel = Number(body.userLevel || 1);
            const memberName = String(body.memberName || "").trim();
            const dateActual = String(body.dateActual || "").trim();
            const dateId = Number(body.dateId || 0);

            if (!memberId || !memberName || !dateActual || !dateId) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error:
                            "Missing member or date information. " +
                            `memberId=${memberId}, ` +
                            `memberName=${memberName || "blank"}, ` +
                            `dateActual=${dateActual || "blank"}, ` +
                            `dateId=${dateId || "blank"}`
                    }
                };
            }

            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            const fields = {
                Title: new Date().toISOString(),
                DateIDLOckInColSP: dateId,
                DateActualColSP: dateActual,
                MemberIDLOckInColSP: memberId,
                MemberNameCombinedColSP: memberName,
                DateRequestWasAdded: new Date().toISOString(),
                Approved: "Requesting",
                UserLevelColSP: userLevel
            };

            context.log("RequestBooking fields:", JSON.stringify(fields));

            const res = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/PendingRequestsListSP/items`,
                {
                    method: "POST",
                    headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({ fields })
                }
            );

            const data = await res.json();

            if (!res.ok) {
                throw new Error("Booking request failed: " + JSON.stringify(data));
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Booking request submitted.",
                    itemId: data.id
                }
            };

        } catch (err) {
            return {
                status: 500,
                jsonBody: {
                    success: false,
                    version: API_VERSION,
                    error: err.message
                }
            };
        }
    }
});

//
// DELETE REQUEST
//
app.http("DeleteRequest", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`DeleteRequest called. Version: ${API_VERSION}`);

        if (request.method === "GET") {
            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "DeleteRequest API is reachable."
                }
            };
        }

        try {
            const body = await request.json();
            const requestId = String(body.requestId || "").trim();

            if (!requestId) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Missing requestId."
                    }
                };
            }

            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            const res = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/PendingRequestsListSP/items/${requestId}`,
                {
                    method: "DELETE",
                    headers: {
                        Authorization: `Bearer ${token}`
                    }
                }
            );

            if (!res.ok) {
                const text = await res.text();
                throw new Error("Delete failed: " + text);
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Request deleted."
                }
            };

        } catch (err) {
            return {
                status: 500,
                jsonBody: {
                    success: false,
                    version: API_VERSION,
                    error: err.message
                }
            };
        }
    }
});

//
// GET SPLASH PAGE PASSWORD + BASE64 IMAGE
//
app.http("GetSplashPagePassword", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`GetSplashPagePassword called. Version: ${API_VERSION}`);

        try {
            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            const items = await getListItems(token, site.id, "SplashPagePassword");

            const item = items[0];
            const f = item?.fields || {};

            const base64Image = String(f.Base64ColSP || "")
                .replace(/^data:image\/[a-zA-Z0-9.+-]+;base64,/, "")
                .replace(/^data:image;application\/octet-stream;base64,/, "")
                .trim();

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    password: f.SplashPagePasswordColSP || "",
                    imageBase64: base64Image
                }
            };

        } catch (err) {
            return {
                status: 500,
                jsonBody: {
                    success: false,
                    version: API_VERSION,
                    error: err.message
                }
            };
        }
    }
});
