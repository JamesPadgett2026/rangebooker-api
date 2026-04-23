<body>

<h1>RangeBooker</h1>

<h2>Test Locations</h2>
<button onclick="loadLocations()">Load Locations</button>
<div id="locations"></div>

<script>
async function loadLocations() {
  const container = document.getElementById('locations');
  container.innerHTML = 'Loading...';

  try {
    const response = await fetch('/api/GetLocations');
    const data = await response.json();

    container.innerHTML = data.locations.map(location => `
      <div style="border:1px solid #ccc; padding:10px; margin:10px 0; border-radius:8px;">
        <strong>${location.name}</strong><br>
        Status: ${location.status}
      </div>
    `).join('');
  } catch (error) {
    container.innerHTML = 'Error loading locations.';
  }
}
</script>

</body>
