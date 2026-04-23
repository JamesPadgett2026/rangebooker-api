<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>RangeBooker</title>
</head>
<body>
  <h1>RangeBooker</h1>

  <h2>Locations</h2>
  <button id="loadBtn">Load Locations</button>

  <div id="status" style="margin-top: 12px; font-weight: bold;"></div>
  <div id="locations" style="margin-top: 12px;"></div>

  <script>
    async function loadLocations() {
      const status = document.getElementById('status');
      const container = document.getElementById('locations');

      status.textContent = 'Loading...';
      container.innerHTML = '';

      try {
        const response = await fetch('/api/GetLocations');
        status.textContent = 'HTTP Status: ' + response.status;

        const text = await response.text();

        let data;
        try {
          data = JSON.parse(text);
        } catch (parseError) {
          container.innerHTML = '<pre>' + text + '</pre>';
          return;
        }

        if (!data.locations || !Array.isArray(data.locations)) {
          container.innerHTML = '<pre>' + JSON.stringify(data, null, 2) + '</pre>';
          return;
        }

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

    document.getElementById('loadBtn').addEventListener('click', loadLocations);
  </script>
</body>
</html>
