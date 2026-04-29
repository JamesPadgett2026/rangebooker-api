module.exports = async function (context, req) {

    const data = [
        {
            MemberEmailColSP: "james@example.com",
            DateIDLOckInColSP: "2026-05-01",
            DateRequestWasAdded: "2026-04-28T20:14:00",
            StatusColSP: "Pending"
        }
    ];

    context.res = {
        status: 200,
        body: data
    };
};
