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

function isDuplicateEmailError(data) {
    const message = String(data?.error?.message || "").toLowerCase();

    return (
        message.includes("unique constraints") ||
        message.includes("duplicate") ||
        message.includes("already has the provided value")
    );
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
            const email = String(body.email || "").trim().toLowerCase();
            const password = String(body.password || "");

            if (!firstName || !lastName || !email || !password) {
                return {
                    status: 400,
                    jsonBody: {
                        success: false,
                        error: "First name, last name, email, and password are required."
                    }
                };
            }

            const token = await getAccessToken();
            const site = await getSite(token);

            const duplicateRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/MemberListSP/items?expand=fields&$top=5000`,
                {
                    headers: {
                        Authorization: `Bearer ${token}`
                    }
                }
            );

            const duplicateData = await readResponseBody(duplicateRes);

            if (!duplicateRes.ok) {
                throw new Error("Duplicate check failed: " + JSON.stringify(duplicateData));
            }

            const emailAlreadyExists = (duplicateData.value || []).some(item => {
                const existingEmailColSP =
                    String(item.fields?.EmailColSP || "").trim().toLowerCase();

                const existingLoginEmail =
                    String(item.fields?.loginemail || "").trim().toLowerCase();

                const existingEmailLowercase =
                    String(item.fields?.email || "").trim().toLowerCase();

                return (
                    existingEmailColSP === email ||
                    existingLoginEmail === email ||
                    existingEmailLowercase === email
                );
            });

            if (emailAlreadyExists) {
                return {
                    status: 409,
                    jsonBody: {
                        success: false,
                        error: "An account with this email already exists."
                    }
                };
            }

            const fields = {
                Title: `${firstName} ${lastName}`,
                FirstNameColSP: firstName,
                LastNameColSP: lastName,
                EmailColSP: email,
                loginemail: email,
                PasswordColSP: password
            };

            context.log("FIELDS BEING SENT:");
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

            context.log("RAW CREATE RESPONSE:");
            context.log(JSON.stringify(createData, null, 2));

            if (!createRes.ok) {
                if (isDuplicateEmailError(createData)) {
                    return {
                        status: 409,
                        jsonBody: {
                            success: false,
                            error: "An account with this email already exists."
                        }
                    };
                }

                return {
                    status: 500,
                    jsonBody: {
                        success: false,
                        error: "Member create failed.",
                        sentFields: fields,
                        graphResponse: createData
                    }
                };
            }

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    message: "Member created successfully",
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
