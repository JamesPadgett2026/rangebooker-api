const { app } = require("@azure/functions");

app.http("GetPendingRequests", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {

    const data = [
      {
        MemberEmailColSP: "james@example.com",
        DateIDLOckInColSP: "2026-05-01",
        DateRequestWasAdded: "2026-04-28T20:14:00",
        StatusColSP: "Pending"
      }
    ];

    return {
      status: 200,
      jsonBody: data
    };
  }
});
