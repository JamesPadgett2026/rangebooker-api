<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>RangeBooker</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 0;
      padding: 0;
      background: #f5f7fb;
      color: #1f2937;
    }

    .wrap {
      max-width: 1100px;
      margin: 0 auto;
      padding: 24px;
    }

    h1 {
      margin: 0 0 8px 0;
      font-size: 32px;
    }

    .subtext {
      margin-bottom: 20px;
      color: #6b7280;
    }

    .toolbar {
      margin-bottom: 20px;
    }

    button {
      background: #2563eb;
      color: white;
      border: none;
      padding: 10px 16px;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
    }

    button:hover {
      background: #1d4ed8;
    }

    #status {
      margin-bottom: 18px;
      font-weight: bold;
    }

    .calendar-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
      gap: 16px;
    }

    .day-card {
      background: white;
      border-radius: 14px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
      padding: 16px;
      border-left: 8px solid #d1d5db;
    }

    .day-card.available {
      border-left-color: #16a34a;
    }

    .day-card.booked {
      border-left-color: #dc2626;
    }

    .date-text {
      font-size: 18px;
      font-weight: 700;
      margin-bottom: 10px;
    }

    .time-text {
      font-size: 15px;
      margin-bottom: 12px;
      color: #374151;
    }

    .badge {
      display: inline-block;
      padding: 6px 10px;
      border-radius: 999px;
      font-size: 13px;
      font-weight: 700;
    }

    .badge.available {
      background: #dcfce7;
      color: #166534;
    }

    .badge.booked {
      background: #fee2e2;
      color: #991b1b;
    }

    .empty-message {
      background: white;
      padding: 20px;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
    }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>RangeBooker</h1>
    <div class="subtext">View days and whether they are available.</div>

    <div class="toolbar">
      <button onclick="loadCalendar()">Load Calendar</button>
    </div>

    <div id="status"></div>
    <div id="calendar" class="calendar-grid"></div>
  </div>

  <script>
    function formatDateText(dateString) {
      const date = new Date(dateString);
      return date.toLocaleDateString('en-US', {
        weekday: 'long',
        month: 'long',
        day: 'numeric',
        year: 'numeric'
      });
    }

    function normalizeStatus(value) {
      return String(value || '').trim().toLowerCase();
    }

    async function loadCalendar() {
      const statusEl = document.getElementById('status');
      const calendarEl = document.getElementById('calendar');

      statusEl.textContent = 'Loading...';
      calendarEl.innerHTML = '';

      try {
        const response = await fetch('/api/GetLocations');
        const data = await response.json();

        if (!response.ok || !data.success) {
          statusEl.textContent = 'Could not load calendar data.';
          calendarEl.innerHTML = `<div class="empty-message"><pre>${JSON.stringify(data, null, 2)}</pre></div>`;
          return;
        }

        const items = (data.locations || [])
          .map(item => ({
            date: item.fields?.DateTimeToSchedule,
            start: item.fields?.StartTimeTextColSP,
            end: item.fields?.EndTimeTextColSP,
            status: item.fields?.AvailableOrBooked
          }))
          .filter(item => item.date && item.status)
          .sort((a, b) => new Date(a.date) - new Date(b.date));

        if (!items.length) {
          statusEl.textContent = 'No calendar items found.';
          calendarEl.innerHTML = `<div class="empty-message">There are no days to display.</div>`;
          return;
        }

        statusEl.textContent = `${items.length} day(s) loaded`;

        calendarEl.innerHTML = items.map(item => {
          const status = normalizeStatus(item.status);
          const isAvailable = status === 'available';

          return `
            <div class="day-card ${isAvailable ? 'available' : 'booked'}">
              <div class="date-text">${formatDateText(item.date)}</div>
              <div class="time-text">${item.start || ''}${item.start && item.end ? ' - ' : ''}${item.end || ''}</div>
              <div class="badge ${isAvailable ? 'available' : 'booked'}">
                ${isAvailable ? 'Available' : 'Booked'}
              </div>
            </div>
          `;
        }).join('');
      } catch (error) {
        statusEl.textContent = 'Error loading calendar.';
        calendarEl.innerHTML = `<div class="empty-message">${error.message}</div>`;
      }
    }

    loadCalendar();
  </script>
</body>
</html>
