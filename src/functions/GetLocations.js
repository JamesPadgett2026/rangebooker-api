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

    if (!response.ok) {
        throw new Error(JSON.stringify(data));
    }

    return data.access_token;
}

app.http('GetLocations', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        try {
            const token = await getAccessToken();

            const siteRes = await fetch(
                'https://graph.microsoft.com/v1.0/sites/tropicaltech.sharepoint.com:/sites/RangeBooker',
                {
                    headers: {
                        Authorization: `Bearer ${token}`
                    }
                }
            );

            const siteData = await siteRes.json();

            if (!siteRes.ok) {
                throw new Error(`Site lookup failed: ${JSON.stringify(siteData)}`);
            }

            const listRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists/CalendarNewSPList/items?expand=fields`,
                {
                    headers: {
                        Authorization: `Bearer ${token}`
                    }
                }
            );

            const listData = await listRes.json();

            if (!listRes.ok) {
                throw new Error(`List lookup failed: ${JSON.stringify(listData)}`);
            }

            return {
                jsonBody: {
                    success: true,
                    locations: (listData.value || []).map((item, index) => ({
                        id: index + 1,
                        name: item.fields?.Title || `Item ${index + 1}`,
                        status: 'Active',
                        fields: item.fields
                    }))
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
