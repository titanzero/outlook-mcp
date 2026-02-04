/**
 * List events functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Use calendarView endpoint to properly include all-day events
    // The me/events endpoint with dateTime filter excludes all-day events
    // because they use date-only format (e.g., "2026-02-04") instead of ISO timestamps
    const now = new Date();
    const startDateTime = now.toISOString();
    // Look ahead 90 days by default
    const endDate = new Date(now);
    endDate.setDate(endDate.getDate() + 90);
    const endDateTime = endDate.toISOString();

    let endpoint = `me/calendarView?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}`;

    // Add query parameters
    const queryParams = {
      $top: count,
      $orderby: 'start/dateTime',
      $select: config.CALENDAR_SELECT_FIELDS
    };

    // Make API call
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No calendar events found."
        }]
      };
    }
    
    // Format results
    const eventList = response.value.map((event, index) => {
      let startDate, endDate;

      // Handle all-day events differently - they use date-only format
      if (event.isAllDay) {
        // All-day events: just show the date(s)
        startDate = event.start.dateTime.split('T')[0];
        endDate = event.end.dateTime.split('T')[0];
        // For single-day all-day events, end date is next day in Graph API
        const startD = new Date(startDate);
        const endD = new Date(endDate);
        if ((endD - startD) / (1000 * 60 * 60 * 24) === 1) {
          // Single day all-day event
          startDate = `${startDate} (All Day)`;
          endDate = startDate;
        } else {
          startDate = `${startDate} (All Day)`;
          endDate = `${event.end.dateTime.split('T')[0]} (All Day)`;
        }
      } else {
        startDate = new Date(event.start.dateTime).toLocaleString(event.start.timeZone);
        endDate = new Date(event.end.dateTime).toLocaleString(event.end.timeZone);
      }

      const location = event.location?.displayName || 'No location';

      // Format attendees list
      let attendeesStr = '';
      if (event.attendees && event.attendees.length > 0) {
        const attendeeNames = event.attendees.map(a => {
          const name = a.emailAddress?.name || a.emailAddress?.address || 'Unknown';
          const status = a.status?.response || '';
          return status ? `${name} (${status})` : name;
        });
        attendeesStr = `\nAttendees: ${attendeeNames.join(', ')}`;
      }

      return `${index + 1}. ${event.subject} - Location: ${location}\nStart: ${startDate}\nEnd: ${endDate}${attendeesStr}\nSummary: ${event.bodyPreview}\nID: ${event.id}\n`;
    }).join("\n");
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} events:\n\n${eventList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error listing events: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEvents;
