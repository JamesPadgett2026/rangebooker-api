const body = await request.json();

const email = body.email;
const password = body.password;

// Look up SharePoint user here...

if (member.MembershipRequestApproved !== "Yes") {
  return {
    status: 403,
    jsonBody: {
      success: false,
      message: "Your account has not been approved yet."
    }
  };
}
