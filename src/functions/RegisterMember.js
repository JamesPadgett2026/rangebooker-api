const { app } = require("@azure/functions");

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

            return {
                status: 200,
                jsonBody: {
                    success: true,
                    message: "RegisterMember API is working.",
                    received: body
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
