const { app } = require("@azure/functions");

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

async function getSite(token) {
    const res = await fetch(
        "https://graph.microsoft.com/v1.0/sites/tropicaltech.sharepoint.com:/sites/RangeBooker",
        {
            headers: { Authorization: `Bearer ${token}` }
        }
    );

    const data = await res.json();

    if (!res.ok) {
        throw new Error(`Site lookup failed: ${JSON.stringify(data)}`);
    }

    return data;
}

function splitPhone(phone) {
    const digits = String(phone || "").replace(/\D/g, "");

    return {
        areaCode: digits.substring(0, 3) || "",
        phone3: digits.substring(3, 6) || "",
        phone4: digits.substring(6, 10) || ""
    };
}

app.http("RegisterMember", {
    methods: ["GET", "POST"],
    authLevel: "anonymous",
    handler: async (request, context) => {
        try {
            // --- Test endpoint ---
            if (request.method === "GET") {
                return {
                    status: 200,
                    jsonBody: {
                        success: true,
                        message: "RegisterMember API is reachable."
                    }
                };
            }

            // --- Get data ---
            const body = await request.json();

            const firstName = (body.firstName || "").trim();
            const lastName = (body.lastName || "").trim();
            const email = (body.email || "").trim().toLowerCase();
            const phone = (body.phone || "").trim();
            const password = body.password || "";
            const notes = (body.notes || "").trim();

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

            // --- CHECK FOR DUPLICATE EMAIL ---
            const safeEmail = email.replace(/'/g, "''");

            const duplicateRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/MemberListSP/items?expand=fields&$filter=fields/EmailColSP eq '${safeEmail}'`,
                {
                    headers: {
                        Authorization: `Bearer ${token}`
                    }
                }
            );

            const duplicateData = await duplicateRes.json();

            if (!duplicateRes.ok) {
                throw new Error(`Duplicate check failed: ${JSON.stringify(duplicateData)}`);
            }

            if ((duplicateData.value || []).length > 0) {
                return {
                    status: 409,
                    jsonBody: {
                        success: false,
                        error: "An account with this email already exists."
                    }
                };
            }

            // --- PREP DATA ---
            const phoneParts = splitPhone(phone);

            const fields = {
                Title: `${firstName} ${lastName}`,

                FirstNameColSP: firstName,
                LastNameColSP: lastName,

                EmailColSP: email,
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

            // --- CREATE MEMBER ---
            const createRes = await fetch(
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

            const createData = await createRes.json();

            if (!createRes.ok) {
                throw new Error(`Create failed: ${JSON.stringify(createData)}`);
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
