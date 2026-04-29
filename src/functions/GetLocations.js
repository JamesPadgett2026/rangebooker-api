// RangeBooker API
// Version: 2026-04-29 RequestBooking Debug Update
// File: src/functions/GetLocations.js

const { app } = require("@azure/functions");

const API_VERSION = "2026-04-29 RequestBooking Debug Update";

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
    const siteRes = await fetch(
        "https://graph.microsoft.com/v1.0/sites/tropicaltech.sharepoint.com:/sites/RangeBooker",
        {
            headers: { Authorization: `Bearer ${token}` }
        }
    );

    const siteData = await siteRes.json();

    if (!siteRes.ok) {
        throw new Error("SharePoint site lookup failed: " + JSON.stringify(siteData));
    }

    return siteData;
}

function splitPhone(phone) {
    const digits = String(phone || "").replace(/\D/g, "");

    return {
        areaCode: digits.length >= 3 ? digits.substring(0, 3) : "",
        phone3: digits.length >= 6 ? digits.substring(3, 6) : "",
        phone4: digits.length >= 10 ? digits.substring(6, 10) : ""
    };
}

function isDuplicateEmailError(createData) {
    const message = String(createData?.error?.message || "").toLowerCase();

    return (
        message.includes("unique constraints") ||
        message.includes("duplicate") ||
        message.includes("already has the provided value")
    );
}

function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
}

function formatSafeDate(value) {
    return value || "";
}

async function getMemberItems(token, siteId) {
    const listRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/MemberListSP/items?expand=fields&$top=5000`,
        {
            headers: { Authorization: `Bearer ${token}` }
        }
    );

    const listData = await listRes.json();

    if (!listRes.ok) {
        throw new Error("Member lookup failed: " + JSON.stringify(listData));
    }

    return listData.value || [];
}

async function getPendingRequestItems(token, siteId) {
    const listRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/PendingRequestsListSP/items?expand=fields&$top=5000`,
        {
            headers: { Authorization: `Bearer ${token}` }
        }
    );

    const listData = await listRes.json();

    if (!listRes.ok) {
        throw new Error("Pending request lookup failed: " + JSON.stringify(listData));
    }

    return listData.value || [];
}

app.http("GetLocations", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`GetLocations called. Version: ${API_VERSION}`);

        try {
            const token = await getAccessToken();
            const siteData = await getRangeBookerSite(token);

            const listRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists/CalendarNewSPList/items?expand=fields&$top=5000`,
                {
                    headers: { Authorization: `Bearer ${token}` }
                }
            );

            const listData = await listRes.json();

            if (!listRes.ok) {
                throw new Error("Calendar lookup failed: " + JSON.stringify(listData));
            }

            const locations = (listData.value || []).map((item, index) => {
                const fields = item.fields || {};

                return {
                    id: item.id,
                    name: fields.Title || `Item ${index + 1}`,
                    dateTimeToSchedule: formatSafeDate(fields.DateTimeToSchedule),
                    startTime: fields.StartTimeTextColSP || "",
                    endTime: fields.EndTimeTextColSP || "",
                    startTimeId: fields.StartTimeTextIDColSP || "",
                    endTimeId: fields.EndTimeTextIDColSP || "",
                    availableOrBooked: fields.AvailableOrBooked || "",
                    displayText:
                        `${fields.StartTimeTextColSP || ""} - ${fields.EndTimeTextColSP || ""}: ${fields.AvailableOrBooked || ""}`
                };
            });

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    locations: locations
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

app.http("GetMyRequests", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`GetMyRequests called. Version: ${API_VERSION}`);

        try {
            const url = new URL(request.url);
            const email = normalizeEmail(url.searchParams.get("email"));
            const memberId = Number(url.searchParams.get("memberId") || 0);

            if (!email && !memberId) {
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
            const siteData = await getRangeBookerSite(token);
            const items = await getPendingRequestItems(token, siteData.id);

            const myRequests = items
                .filter(item => {
                    const fields = item.fields || {};

                    const requestMemberId = Number(fields.MemberIDLOckInColSP || 0);
                    const requestEmail = normalizeEmail(fields.MemberEmailColSP || fields.EmailColSP || "");

                    if (memberId && requestMemberId === memberId) {
                        return true;
                    }

                    if (email && requestEmail && requestEmail === email) {
                        return true;
                    }

                    return false;
                })
                .map(item => {
                    const fields = item.fields || {};

                    return {
                        id: item.id,
                        dateId: fields.DateIDLOckInColSP || "",
                        requestedDate: fields.DateActualColSP || "",
                        status: fields.Approved || "Requesting",
                        userLevel: fields.UserLevelColSP || "",
                        requestedAt: fields.DateRequestWasAdded || ""
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
                    requests: myRequests
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

            const phoneParts = splitPhone(phone);
            const token = await getAccessToken();
            const siteData = await getRangeBookerSite(token);

            const fieldsToCreate = {
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
                DateJoined: new Date().toISOString(),
                Notes: notes || ""
            };

            const createRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists/MemberListSP/items`,
                {
                    method: "POST",
                    headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({ fields: fieldsToCreate })
                }
            );

            const createData = await createRes.json();

            if (!createRes.ok) {
                if (isDuplicateEmailError(createData)) {
                    return {
                        status: 409,
                        jsonBody: {
                            success: false,
                            version: API_VERSION,
                            error: "An account with this email already exists."
                        }
                    };
                }

                throw new Error("Member create failed: " + JSON.stringify(createData));
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Member created in SharePoint. Account is pending approval.",
                    itemId: createData.id
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
            const siteData = await getRangeBookerSite(token);
            const members = await getMemberItems(token, siteData.id);

            const matchingMember = members.find(item => {
                const fields = item.fields || {};

                const email1 = normalizeEmail(fields.email);
                const email2 = normalizeEmail(fields.loginemail);
                const email3 = normalizeEmail(fields.EmailColSP);

                return email1 === email || email2 === email || email3 === email;
            });

            if (!matchingMember) {
                return {
                    status: 401,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Invalid email or password."
                    }
                };
            }

            const fields = matchingMember.fields || {};
            const savedPassword = String(fields.PasswordColSP || "");
            const activeValue = String(fields.Active || "").trim().toLowerCase();
            const approvedValue = String(fields.MembershipRequestApproved || "").trim().toLowerCase();

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

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Login successful.",
                    member: {
                        id: matchingMember.id,
                        firstName: fields.FirstNameColSP || "",
                        lastName: fields.LastNameColSP || "",
                        email: email,
                        title: fields.Title || "",
                        memberType: fields.MemberType || 1,
                        membershipRequestApproved: fields.MembershipRequestApproved || ""
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
            const memberEmail = normalizeEmail(body.memberEmail || body.email);
            const dateActual = String(body.dateActual || "").trim();
            const dateId = Number(body.dateId || 0);

            if (!memberId || !memberName || !memberEmail || !dateActual || !dateId) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error:
                            "Missing member or date information. " +
                            `memberId=${memberId}, ` +
                            `memberName=${memberName || "blank"}, ` +
                            `memberEmail=${memberEmail || "blank"}, ` +
                            `dateActual=${dateActual || "blank"}, ` +
                            `dateId=${dateId || "blank"}`
                    }
                };
            }

            const token = await getAccessToken();
            const siteData = await getRangeBookerSite(token);

            const fieldsToCreate = {
                Title: new Date().toISOString(),
                DateIDLOckInColSP: dateId,
                DateActualColSP: dateActual,
                MemberIDLOckInColSP: memberId,
                MemberNameCombinedColSP: memberName,
                DateRequestWasAdded: new Date().toISOString(),
                Approved: "Requesting",
                UserLevelColSP: userLevel,
            };

            context.log("RequestBooking fieldsToCreate:", JSON.stringify(fieldsToCreate));

            const createRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists/PendingRequestsListSP/items`,
                {
                    method: "POST",
                    headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({ fields: fieldsToCreate })
                }
            );

            const createData = await createRes.json();

            if (!createRes.ok) {
                throw new Error(
                    "Booking request failed: " +
                    JSON.stringify(createData)
                );
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Booking request submitted.",
                    itemId: createData.id
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
