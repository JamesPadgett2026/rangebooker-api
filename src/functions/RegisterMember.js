const { app } = require("@azure/functions");

async function getAccessToken() {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: "POST",
        headers: {
            "Content-Type": "application/x-www-form-urlencoded"
        },
        body: new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            scope: "https://graph.microsoft.com/.default",
            grant_type: "client_credentials"
        })
    });

    const data = await response.json();

    if (!response.ok) {
        throw new Error("Token failed: " + JSON.stringify(data));
    }

    return data.access_token;
}

async function getSite(token) {
    const res = await fetch(
        "https://graph.microsoft.com/v1.0/sites/tropicaltech.sharepoint.com:/sites/RangeBooker",
        {
            headers: {
                Authorization: `Bearer ${token}`
            }
        }
    );

    const data = await res.json();

    if (!res.ok) {
        throw new Error("Site lookup failed: " + JSON.stringify(data));
    }

    return data;
}

async function readResponseBody(response) {
    const text = await response.text();

    if (!text) {
        return {};
    }

    try {
        return JSON.parse(text);
    } catch {
        return {
            rawResponse: text
        };
    }
}

app.http("RegisterMember", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",

    handler: async (request, context) => {
        try {
            if (request.method === "GET") {
                return {
                    status: 200,
                    jsonBody: {
                        success: true,
                        message: "RegisterMember API is reachable."
                    }
                };
            }

            const body = await request.json();

            const firstName = String(body.firstName || "").trim();
            const lastName = String(body.lastName || "").trim();

            const title =
                `${firstName} ${lastName}`.trim() ||
                "Test Member";

            const token = await getAccessToken();
            const site = await getSite(token);

            const fields = {
                Title: title
            };

            context.log("TITLE ONLY FIELDS BEING SENT:");
            context.log(JSON.stringify(fields, null, 2));

            const createRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/MemberListSP/items`,
                {
                    method: "POST",
                    headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({
                        fields
                    })
                }
            );

            const createData = await readResponseBody(createRes);

            context.log("TITLE ONLY CREATE RESPONSE:");
            context.log(JSON.stringify(createData, null, 2));

            if (!createRes.ok) {
                return {
                    status: 500,
                    jsonBody: {
                        success: false,
                        error: "Title-only member create failed.",
                        sentFields: fields,
                        graphResponse: createData
                    }
                };
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    message: "Title-only member created successfully.",
                    id: createData.id
                }
            };

        } catch (err) {
            context.log("REGISTER MEMBER ERROR:");
            context.log(err.message);

            return {
                status: 500,
                jsonBody: {
                    success: false,
                    error: err.message
                }
            };
        }
    }
});
