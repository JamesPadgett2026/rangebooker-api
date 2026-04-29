// RangeBooker API
// Version: 2026-04-29 FINAL (Booking + Delete Request)
// File: src/functions/GetLocations.js

const { app } = require("@azure/functions");

const API_VERSION = "2026-04-29 FINAL";

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

function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
}

async function getMemberItems(token, siteId) {
    const res = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/MemberListSP/items?expand=fields&$top=5000`,
        { headers: { Authorization: `Bearer ${token}` } }
    );

    const data = await res.json();

    if (!res.ok) throw new Error(JSON.stringify(data));

    return data.value || [];
}

async function getPendingRequestItems(token, siteId) {
    const res = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/PendingRequestsListSP/items?expand=fields&$top=5000`,
        { headers: { Authorization: `Bearer ${token}` } }
    );

    const data = await res.json();

    if (!res.ok) throw new Error(JSON.stringify(data));

    return data.value || [];
}

//
// 🔐 LOGIN
//
app.http("LoginMember", {
    methods: ["POST"],
    authLevel: "anonymous",
    handler: async (request) => {
        const body = await request.json();

        const email = normalizeEmail(body.email);
        const password = String(body.password || "");

        const token = await getAccessToken();
        const site = await getRangeBookerSite(token);
        const members = await getMemberItems(token, site.id);

        const member = members.find(m => {
            const f = m.fields || {};
            return (
                normalizeEmail(f.email) === email ||
                normalizeEmail(f.loginemail) === email ||
                normalizeEmail(f.EmailColSP) === email
            );
        });

        if (!member) {
            return { status: 401, jsonBody: { success: false, error: "Invalid login" } };
        }

        const f = member.fields || {};

        if (String(f.PasswordColSP) !== password) {
            return { status: 401, jsonBody: { success: false, error: "Invalid login" } };
        }

        if (String(f.MembershipRequestApproved).toLowerCase() !== "yes") {
            return { status: 403, jsonBody: { success: false, error: "Account not approved" } };
        }

        return {
            status: 200,
            jsonBody: {
                success: true,
                member: {
                    id: member.id,
                    firstName: f.FirstNameColSP,
                    lastName: f.LastNameColSP,
                    email: email,
                    memberType: f.MemberType || 1
                }
            }
        };
    }
});

//
// 📅 GET MY REQUESTS
//
app.http("GetMyRequests", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: async (request) => {

        const url = new URL(request.url);
        const memberId = Number(url.searchParams.get("memberId") || 0);

        const token = await getAccessToken();
        const site = await getRangeBookerSite(token);
        const items = await getPendingRequestItems(token, site.id);

        const requests = items
            .filter(i => Number(i.fields?.MemberIDLOckInColSP || 0) === memberId)
            .map(i => ({
                id: i.id,
                requestedDate: i.fields.DateActualColSP,
                status: i.fields.Approved || "Requesting",
                requestedAt: i.fields.DateRequestWasAdded
            }));

        return {
            status: 200,
            jsonBody: { success: true, requests }
        };
    }
});

//
// 📩 CREATE REQUEST
//
app.http("RequestBooking", {
    methods: ["POST"],
    authLevel: "anonymous",
    handler: async (request, context) => {

        const body = await request.json();

        const memberId = Number(body.memberId || 0);
        const memberName = String(body.memberName || "");
        const dateActual = String(body.dateActual || "");
        const dateId = Number(body.dateId || 0);

        if (!memberId || !memberName || !dateActual || !dateId) {
            return {
                status: 400,
                jsonBody: {
                    success: false,
                    error: "Missing required fields"
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
            UserLevelColSP: 1
        };

        context.log("Create Request:", JSON.stringify(fields));

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
            throw new Error(JSON.stringify(data));
        }

        return {
            status: 200,
            jsonBody: { success: true }
        };
    }
});

//
// ❌ DELETE REQUEST (NEW)
//
app.http("DeleteRequest", {
    methods: ["POST"],
    authLevel: "anonymous",
    handler: async (request) => {

        const body = await request.json();
        const requestId = String(body.requestId || "");

        const token = await getAccessToken();
        const site = await getRangeBookerSite(token);

        const res = await fetch(
            `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/PendingRequestsListSP/items/${requestId}`,
            {
                method: "DELETE",
                headers: { Authorization: `Bearer ${token}` }
            }
        );

        if (!res.ok) {
            throw new Error("Delete failed");
        }

        return {
            status: 200,
            jsonBody: { success: true }
        };
    }
});
