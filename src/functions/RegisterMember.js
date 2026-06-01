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

async function getSite(token) {
    const response = await fetch(
        "https://graph.microsoft.com/v1.0/sites/tropicaltech.sharepoint.com:/sites/RangeBooker",
        {
            headers: {
                Authorization: `Bearer ${token}`
            }
        }
    );

    const data = await readResponseBody(response);

    if (!response.ok) {
        throw new Error("Site lookup failed: " + JSON.stringify(data));
    }

    return data;
}

async function getList(token, siteId) {
    const response = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
        {
            headers: {
                Authorization: `Bearer ${token}`
            }
        }
    );

    const data = await readResponseBody(response);

    if (!response.ok) {
        throw new Error("List lookup failed: " + JSON.stringify(data));
    }

    const lists = data.value || [];

    const memberList = lists.find(list => {
        const name = String(list.name || "").trim().toLowerCase();
        const displayName = String(list.displayName || "").trim().toLowerCase();

        return (
            name === "memberlistsp" ||
            displayName === "memberlistsp" ||
            displayName === "member list sp" ||
            displayName === "memberlist"
        );
    });

    if (!memberList) {
        throw new Error(
            "Could not find MemberListSP. Available lists: " +
            lists.map(list => `${list.displayName || list.name} (${list.id})`).join(", ")
        );
    }

    return memberList;
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

            const title =
                `${firstName} ${lastName}`.trim() ||
                "Test Member";

            const token = await getAccessToken();
            const site = await getSite(token);
            const memberList = await getList(token, site.id);

            context.log("USING LIST:");
            context.log(JSON.stringify(memberList, null, 2));

            const fields = {
                Title: title
            };

            context.log("FIELDS BEING SENT:");
            context.log(JSON.stringify(fields, null, 2));

            const createResponse = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/${memberList.id}/items`,
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

            const createData = await readResponseBody(createResponse);

            context.log("CREATE RESPONSE:");
            context.log(JSON.stringify(createData, null, 2));

            if (!createResponse.ok) {
                if (isDuplicateEmailError(createData)) {
                    return {
                        status: 409,
                        jsonBody: {
                            success: false,
                            error: "An account with this email already exists.",
                            graphResponse: createData
                        }
                    };
                }

                return {
                    status: 500,
                    jsonBody: {
                        success: false,
                        error: "Member create failed.",
                        listUsed: memberList,
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
                    id: createData.id,
                    listUsed: {
                        id: memberList.id,
                        name: memberList.name,
                        displayName: memberList.displayName
                    }
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
