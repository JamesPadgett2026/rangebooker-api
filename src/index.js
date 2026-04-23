<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>RangeBooker</title>
</head>
<body>
  <h1>RangeBooker</h1>
const { app } = require('@azure/functions');

  <h2>Locations</h2>
  <button id="loadBtn">Load Locations</button>
async function getAccessToken() {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

  <div id="status" style="margin-top: 12px; font-weight: bold;"></div>
  <div id="locations" style="margin-top: 12px;"></div>
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

  <script>
    async function loadLocations() {
      const status = document.getElementById('status');
      const container = document.getElementById('locations');
    const data = await response.json();
    return data.access_token;
}

      status.textContent = 'Loading...';
      container.innerHTML = '';
app.http('GetLocations', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: async (request, context) => {

      try {
        const response = await fetch('/api/GetLocations');
        status.textContent = 'HTTP Status: ' + response.status;
        try {
            const token = await getAccessToken();

        const text = await response.text();
            // Get SharePoint site
            const siteRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/tropicaltech.sharepoint.com:/sites/CharleyToppinoAndSons`,
                { headers: { Authorization: `Bearer ${token}` } }
            );

        let data;
        try {
          data = JSON.parse(text);
        } catch (parseError) {
          container.innerHTML = '<pre>' + text + '</pre>';
          return;
        }
            const site = await siteRes.json();

        if (!data.locations || !Array.isArray(data.locations)) {
          container.innerHTML = '<pre>' + JSON.stringify(data, null, 2) + '</pre>';
          return;
        }
            // Get list items
            const listRes = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/MCPDriverList/items?expand=fields`,
                { headers: { Authorization: `Bearer ${token}` } }
            );

        container.innerHTML = data.locations.map(location => `
          <div style="border:1px solid #ccc; padding:10px; margin:10px 0; border-radius:8px;">
            <strong>${location.name}</strong><br>
            Status: ${location.status}<br>
            ID: ${location.id}
          </div>
        `).join('');
      } catch (error) {
        status.textContent = 'JavaScript error';
        container.innerHTML = '<pre>' + error.message + '</pre>';
      }
    }
            const listData = await listRes.json();

    document.getElementById('loadBtn').addEventListener('click', loadLocations);
  </script>
</body>
</html>
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
