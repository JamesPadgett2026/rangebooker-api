const { app } = require("@azure/functions");

app.http("GetLocations", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    context.log("GetLocations called");

    return {
      status: 200,
      jsonBody: {
        success: true,
        locations: [
          {
            id: 1,
            name: "Sample Location 1",
            status: "Active"
          },
          {
            id: 2,
            name: "Sample Location 2",
            status: "Active"
          }
        ]
      }
    };
  }
});

app.http("RegisterMember", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    context.log("RegisterMember called");

    let body;

    try {
      body = await request.json();
    } catch (err) {
      return {
        status: 400,
        jsonBody: {
          error: "Invalid JSON body."
        }
      };
    }

    const firstName = (body.firstName || "").trim();
    const lastName = (body.lastName || "").trim();
    const email = (body.email || "").trim();
    const phone = (body.phone || "").trim();
    const password = body.password || "";
    const notes = (body.notes || "").trim();

    if (!firstName || !lastName || !email || !password) {
      return {
        status: 400,
        jsonBody: {
          error: "First name, last name, email, and password are required."
        }
      };
    }

    return {
      status: 200,
      jsonBody: {
        success: true,
        message: "RegisterMember API is working.",
        received: {
          firstName,
          lastName,
          email,
          phone,
          notes
        }
      }
    };
  }
});
