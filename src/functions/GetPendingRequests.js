const { app } = require("@azure/functions");

app.http("GetPendingRequests", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    const data = [
      {
        MemberEmailColSP: "acousticjohnson.aj@gmail.com",
        DateIDLOckInColSP: "2026-05-01",
        DateRequestWasAdded: "2026-04-28T20:14:00",
        StatusColSP: "Pending"
      },
      {
        MemberEmailColSP: "acousticjohnson.aj@gmail.com",
        DateIDLOckInColSP: "2026-05-03",
        DateRequestWasAdded: "2026-04-27T18:10:00",
        StatusColSP: "Approved"
      }
    ];

    return {
      status: 200,
      jsonBody: data
    };
  }
});
