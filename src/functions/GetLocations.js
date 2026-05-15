// RangeBooker API
// File: src/functions/GetLocations.js
// Version: 2026-05-14 WEBSITE PHOTOS ADDED

const { app } = require("@azure/functions");

const API_VERSION = "2026-05-14 WEBSITE PHOTOS ADDED";

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

    if (cleanBase64.startsWith("R0lGOD")) {
        mimeType = "image/gif";
    }

    if (cleanBase64.startsWith("UklGR")) {
        mimeType = "image/webp";
    }

    return `data:${mimeType};base64,${cleanBase64}`;
}

function getCleanBase64(base64Value) {
    return String(base64Value || "")
        .replace(/^data:image\/[a-zA-Z0-9.+-]+;base64,/, "")
        .replace(/^data:image;application\/octet-stream;base64,/, "")
        .trim();
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

            const today = new Date();
            today.setHours(0, 0, 0, 0);

            const locations = items
                .map((item, index) => {
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
                })
                .filter(loc => {
                    if (!loc.dateTimeToSchedule) {
                        return false;
                    }

                    const d = new Date(loc.dateTimeToSchedule);

                    if (isNaN(d.getTime())) {
                        return false;
                    }

                    d.setHours(0, 0, 0, 0);

                    return d >= today;
                })
                .sort((a, b) => {
                    return new Date(a.dateTimeToSchedule) - new Date(b.dateTimeToSchedule);
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
//
// GET WEBSITE PHOTOS
// Use this before login / before GetEvents.
// Pulls photos from WebsitePhotoListSP, splash image, bulletin board photos, and event photos.
//
app.http("GetWebsitePhotos", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`GetWebsitePhotos called. Version: ${API_VERSION}`);

        try {
            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            let websitePhotoItems = [];
            let splashItems = [];
            let bulletinPosts = [];
            let bulletinPhotos = [];
            let eventItems = [];
            let eventPhotos = [];

            try {
                websitePhotoItems = await getListItems(token, site.id, "WebsitePhotoListSP");
            } catch (err) {
                context.warn("WebsitePhotoListSP load skipped: " + err.message);
            }

            try {
                splashItems = await getListItems(token, site.id, "SplashPagePassword");
            } catch (err) {
                context.warn("SplashPagePassword photo load skipped: " + err.message);
            }

            try {
                bulletinPosts = await getListItems(token, site.id, "BulletinBoardPostsSP");
            } catch (err) {
                context.warn("BulletinBoardPostsSP load skipped: " + err.message);
            }

            try {
                bulletinPhotos = await getListItems(token, site.id, "BulletinBoardPhotos");
            } catch (err) {
                context.warn("BulletinBoardPhotos load skipped: " + err.message);
            }

            try {
                eventItems = await getListItems(token, site.id, "EventListMaster");
            } catch (err) {
                context.warn("EventListMaster load skipped: " + err.message);
            }

            try {
                eventPhotos = await getListItems(token, site.id, "EventPhotosBase64");
            } catch (err) {
                context.warn("EventPhotosBase64 load skipped: " + err.message);
            }

            const photos = [];

            //
            // WEBSITE PHOTOS LIST
            //
            websitePhotoItems.forEach(item => {
                const f = item.fields || {};

                const image = buildImageDataUrl(
                    f.PhotoBase64ColSP || ""
                );

                if (image) {
                    photos.push({
                        id: `website-${item.id}`,
                        source: "WebsitePhotoListSP",
                        type: "website",
                        title: f.PhotoNameColSP || f.Title || "",
                        photoName: f.PhotoNameColSP || "",
                        relatedId: item.id,
                        notes: f.NotesColSP || "",
                        date: f.DateTimeUpdated || "",
                        image: image
                    });
                }
            });

            //
            // SPLASH PHOTO
            //
            const splashFields = splashItems[0]?.fields || {};
            const splashImage = buildImageDataUrl(
                splashFields.Base64ColSP || ""
            );

            if (splashImage) {
                photos.push({
                    id: "splash-main",
                    source: "SplashPagePassword",
                    type: "splash",
                    title: "Splash Image",
                    photoName: "Splash Image",
                    relatedId: "",
                    image: splashImage
                });
            }

            //
            // BULLETIN BOARD PHOTOS
            //
            bulletinPhotos.forEach(photo => {
                const pf = photo.fields || {};
                const postId = Number(pf.BBPostIDLockInColSP || 0);

                const matchingPost = bulletinPosts.find(post => {
                    const postFields = post.fields || {};
                    return Number(postFields.ID || post.id || 0) === postId;
                });

                const matchingPostFields = matchingPost?.fields || {};
                const image = buildImageDataUrl(pf.Base64ColSP || "");

                if (image) {
                    photos.push({
                        id: `bulletin-${photo.id}`,
                        source: "BulletinBoardPhotos",
                        type: "bulletin",
                        title:
                            pf.PhotoTitleColSP ||
                            matchingPostFields.PostTitleColSP ||
                            matchingPostFields.Title ||
                            "Bulletin Board Photo",
                        photoName: pf.PhotoTitleColSP || "",
                        relatedId: postId,
                        postTitle:
                            matchingPostFields.PostTitleColSP ||
                            matchingPostFields.Title ||
                            "",
                        category: matchingPostFields.CategoryColSP || "",
                        date:
                            matchingPostFields.DatePostInformation ||
                            matchingPostFields.DateAddedColSP ||
                            "",
                        image: image
                    });
                }
            });

            //
            // EVENT PHOTOS
            //
            eventPhotos.forEach(photo => {
                const pf = photo.fields || {};
                const eventId = Number(pf.EventLockInIDColSP || 0);

                const matchingEvent = eventItems.find(event => {
                    const eventFields = event.fields || {};
                    return Number(eventFields.ID || event.id || 0) === eventId;
                });

                const matchingEventFields = matchingEvent?.fields || {};
                const image = buildImageDataUrl(
                    pf.Base64ColSP || pf.Base64 || ""
                );

                if (image) {
                    photos.push({
                        id: `event-${photo.id}`,
                        source: "EventPhotosBase64",
                        type: "event",
                        title:
                            matchingEventFields.EventNameColSP ||
                            matchingEventFields.Title ||
                            "Event Photo",
                        photoName:
                            matchingEventFields.EventNameColSP ||
                            matchingEventFields.Title ||
                            "",
                        relatedId: eventId,
                        eventDate:
                            matchingEventFields.EventDate ||
                            matchingEventFields.WhenCreated ||
                            "",
                        eventDateText: formatEventDate(
                            matchingEventFields.EventDate ||
                            matchingEventFields.WhenCreated ||
                            ""
                        ),
                        image: image
                    });
                }
            });

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    count: photos.length,
                    photos: photos
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
// GET BULLETIN BOARD POSTS WITH PHOTOS
//
app.http("GetBulletinBoardPosts", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`GetBulletinBoardPosts called. Version: ${API_VERSION}`);

        try {
            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            const items = await getListItems(token, site.id, "BulletinBoardPostsSP");
            const photoItems = await getListItems(token, site.id, "BulletinBoardPhotos");

            const posts = items
                .map(item => {
                    const f = item.fields || {};
                    const postId = Number(f.ID || item.id);

                    const matchingPhotos = photoItems
                        .filter(photo => {
                            const pf = photo.fields || {};
                            return Number(pf.BBPostIDLockInColSP || 0) === postId;
                        })
                        .map(photo => {
                            const pf = photo.fields || {};

                            return {
                                id: photo.id,
                                title: pf.PhotoTitleColSP || "",
                                image: buildImageDataUrl(pf.Base64ColSP || "")
                            };
                        })
                        .filter(photo => photo.image);

                    return {
                        id: item.id,
                        title: f.PostTitleColSP || f.Title || "Untitled Post",
                        information: f.PostInformationColSP || "",
                        category: f.CategoryColSP || "",
                        datePostInformation: f.DatePostInformation || "",
                        dateAdded: f.DateAddedColSP || "",
                        photos: matchingPhotos
                    };
                })
                .sort((a, b) => {
                    return (
                        new Date(b.datePostInformation || b.dateAdded || 0) -
                        new Date(a.datePostInformation || a.dateAdded || 0)
                    );
                });

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    count: posts.length,
                    posts: posts
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
// GET EVENTS
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

            const firstPhotoFields = base64Photos[0]?.fields || {};

            const results = events.map(item => {
                const f = item.fields || {};
                const eventId = Number(f.ID || item.id);

                const photos = base64Photos.filter(photo => {
                    const pf = photo.fields || {};
                    return Number(pf.EventLockInIDColSP || 0) === eventId;
                });

                const images = photos
                    .map(photo => {
                        const pf = photo.fields || {};
                        return buildImageDataUrl(pf.Base64ColSP || pf.Base64 || "");
                    })
                    .filter(src => src);

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
                    eventDate: eventDate,
                    eventDateText: formatEventDate(eventDate),
                    createdDate:
                        f.WhenCreated ||
                        "",
                    createdDateText:
                        formatEventDate(f.WhenCreated || ""),
                    createdByName:
                        f.CreatedName ||
                        "",
                    images: images,
                    imageDataUrl: images.length > 0 ? images[0] : "",
                    debugPhotoInfo: {
                        eventId: eventId,
                        totalBase64Photos: base64Photos.length,
                        matchedPhotos: photos.length,
                        firstPhotoFields: firstPhotoFields
                    }
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
                    totalBase64Photos: base64Photos.length,
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
                Notes: notes || "",
                PreferredContactChoice: "None"
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
                        preferredContactChoice: f.PreferredContactChoice || "",
                        dateJoined: f.DateJoined || "",
                        lastLogin: lastLoginNow
                    }
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

            const base64Image = getCleanBase64(f.Base64ColSP || "");

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    password: f.SplashPagePasswordColSP || "",
                    imageBase64: base64Image,
                    imageDataUrl: buildImageDataUrl(base64Image)
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
// UPDATE MY SETTINGS
//
app.http("UpdateMySettings", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        context.log(`UpdateMySettings called. Version: ${API_VERSION}`);

        if (request.method === "GET") {
            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "UpdateMySettings API is reachable."
                }
            };
        }

        try {
            const body = await request.json();

            const memberId = String(body.memberId || "").trim();
            const preferredContactChoice = String(body.preferredContactChoice || "").trim();
            const newPassword = String(body.newPassword || "").trim();

            if (!memberId) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Missing memberId."
                    }
                };
            }

            if (!preferredContactChoice) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Missing preferred contact choice."
                    }
                };
            }

            const allowedChoices = ["Text", "Email", "None"];

            if (!allowedChoices.includes(preferredContactChoice)) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Invalid preferred contact choice."
                    }
                };
            }

            const token = await getAccessToken();
            const site = await getRangeBookerSite(token);

            const fieldsToUpdate = {
                PreferredContactChoice: preferredContactChoice
            };

            if (newPassword) {
                fieldsToUpdate.PasswordColSP = newPassword;
            }

            const res = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/MemberListSP/items/${memberId}/fields`,
                {
                    method: "PATCH",
                    headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify(fieldsToUpdate)
                }
            );

            if (!res.ok) {
                const text = await res.text();

                return {
                    status: res.status,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "SharePoint update failed.",
                        details: text
                    }
                };
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Settings updated successfully.",
                    preferredContactChoice: preferredContactChoice
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
