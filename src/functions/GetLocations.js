// RangeBooker API
// Version: 2026-04-27 07:20 PM Eastern
// File: src/functions/GetLocations.js
// Notes:
// - Keeps GetLocations working
// - RegisterMember writes to MemberListSP
// - LoginMember checks MemberListSP for email/password
// - RequestBooking writes to PendingRequestsListSP
// - GetLocations now returns real SharePoint item.id for calendar dates

const { app } = require("@azure/functions");

const API_VERSION = "2026-04-27 07:20 PM Eastern";

async function getAccessToken() {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            scope: "https://graph.microsoft.com/.default",
            grant_type: "client_credentials"
        })
    });

    const data = await response.json();

    if (!response.ok) {
        throw new Error(JSON.stringify(data));
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
        throw new Error(`Site lookup failed: ${JSON.stringify(siteData)}`);
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

async function getMemberItems(token, siteId) {
    const listRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/MemberListSP/items?expand=fields&$top=5000`,
        {
            headers: { Authorization: `Bearer ${token}` }
        }
    );

    const listData = await listRes.json();

    if (!listRes.ok) {
        throw new Error(`Member lookup failed: ${JSON.stringify(listData)}`);
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
                `https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists/CalendarNewSPList/items?expand=fields`,
                {
                    headers: { Authorization: `Bearer ${token}` }
                }
            );

            const listData = await listRes.json();

            if (!listRes.ok) {
                throw new Error(`List lookup failed: ${JSON.stringify(listData)}`);
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    locations: (listData.value || []).map((item, index) => ({
                        id: item.id,
                        name: item.fields?.Title || `Item ${index + 1}`,
                        status: "Active",
                        fields: item.fields
                    }))
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
            const email = String(body.email || "").trim().toLowerCase();
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

                throw new Error(`Member create failed: ${JSON.stringify(createData)}`);
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    version: API_VERSION,
                    message: "Member created in SharePoint.",
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

            const email = String(body.email || "").trim().toLowerCase();
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

                const email1 = String(fields.email || "").trim().toLowerCase();
                const email2 = String(fields.loginemail || "").trim().toLowerCase();
                const email3 = String(fields.EmailColSP || "").trim().toLowerCase();

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

            if (activeValue && activeValue !== "yes" && activeValue !== "true" && activeValue !== "active") {
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
                        memberType: fields.MemberType || 1
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
            const dateActual = String(body.dateActual || "").trim();
            const dateId = Number(body.dateId || 0);

            if (!memberId || !memberName || !dateActual || !dateId) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        version: API_VERSION,
                        error: "Missing member or date information."
                    }
                };
            }

            const token = await getAccessToken();
            const siteData = await getRangeBookerSite(token);

            const fieldsToCreate = {
                Title: new Date().toISOString(),
                MemberIDLOckInColSP: memberId,
                DateRequestWasAdded: new Date().toISOString(),
                Approved: "Requesting",
                UserLevelColSP: userLevel,
                MemberNameCombinedColSP: memberName,
                DateActualColSP: dateActual,
                DateIDLOckInColSP: dateId
            };

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
                throw new Error(`Booking request failed: ${JSON.stringify(createData)}`);
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
