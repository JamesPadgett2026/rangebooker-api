const { app } = require('@azure/functions');

async function getAccessToken() {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            scope: 'https://graph.microsoft.com/.default',
            grant_type: 'client_credentials'
        })
    });

    const data = await response.json();
    return data.access_token;
}

app.http('GetLocations', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: async (request, context) => {

        try {
            const token = await getAccessToken();

            // Get SharePoint site
            const siteRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/tropicaltech.sharepoint.com:/sites/CharleyToppinoAndSons`,
                { headers: { Authorization: `Bearer ${token}` } }
            );

            const site = await siteRes.json();

            // Get list items
            const listRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/MCPDriverList/items?expand=fields`,
                { headers: { Authorization: `Bearer ${token}` } }
            );

            const listData = await listRes.json();

            const locations = listData.value.map((item, index) => ({
                id: index + 1,
                name: item.fields.DriverName || "Unknown",
                status: "Active"
            }));

            return {
                jsonBody: {
                    success: true,
                    locations: locations
                }
            };

        } catch (err) {
            return {
                status: 500,
                jsonBody: {
                    error: err.message
                }
            };
        }
    }
});
