/**
 * List events functionality
 * Uses the Microsoft Graph JS SDK.
 */
const config = require('../config');
const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Resolve human-readable start/end strings for a calendar event.
 * @param {object} event - Graph API event object
 * @returns {{ startDate: string, endDate: string }}
 */
function resolveEventDates(event) {
  if (event.isAllDay) {
    const startStr = event.start.dateTime.split('T')[0];
    const endStr = event.end.dateTime.split('T')[0];
    const startD = new Date(startStr);
    const endD = new Date(endStr);
    if ((endD - startD) / (1000 * 60 * 60 * 24) === 1) {
      return { startDate: `${startStr} (All Day)`, endDate: `${startStr} (All Day)` };
    }
    return { startDate: `${startStr} (All Day)`, endDate: `${endStr} (All Day)` };
  }
  return {
    startDate: new Date(event.start.dateTime).toLocaleString(event.start.timeZone),
    endDate: new Date(event.end.dateTime).toLocaleString(event.end.timeZone),
  };
}

/**
 * Build a compact attendees string.
 * @param {Array} attendees
 * @returns {string}
 */
function formatAttendees(attendees) {
  if (!attendees || attendees.length === 0) return '';
  return attendees.map(a => {
    const name = a.emailAddress?.name || a.emailAddress?.address || 'Unknown';
    const status = a.status?.response || '';
    return status ? `${name} (${status})` : name;
  }).join(', ');
}

/**
 * Build structured rows for TOON encoding.
 * @param {Array} events - Raw Graph API event objects
 * @returns {Array<object>}
 */
function buildEventRows(events) {
  return events.map((event, index) => {
    const { startDate, endDate } = resolveEventDates(event);
    return {
      n: index + 1,
      subject: event.subject,
      start: startDate,
      end: endDate,
      location: event.location?.displayName || '',
      attendees: formatAttendees(event.attendees),
      summary: event.bodyPreview || '',
      id: event.id,
    };
  });
}

/**
 * Build human-readable plain-text output for an event list.
 * @param {Array} events - Raw Graph API event objects
 * @returns {string}
 */
function buildEventText(events) {
  const eventList = events.map((event, index) => {
    const { startDate, endDate } = resolveEventDates(event);
    const location = event.location?.displayName || 'No location';

    let attendeesStr = '';
    const atStr = formatAttendees(event.attendees);
    if (atStr) attendeesStr = `\nAttendees: ${atStr}`;

    return `${index + 1}. ${event.subject} - Location: ${location}\nStart: ${startDate}\nEnd: ${endDate}${attendeesStr}\nSummary: ${event.bodyPreview}\nID: ${event.id}\n`;
  }).join("\n");

  return `Found ${events.length} events:\n\n${eventList}`;
}

/**
 * List events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);

  try {
    const client = await getGraphClient();

    // Use calendarView endpoint to properly include all-day events
    const now = new Date();
    const startDateTime = now.toISOString();
    const endDate = new Date(now);
    endDate.setDate(endDate.getDate() + 90);
    const endDateTime = endDate.toISOString();

    // Make API call
    const response = await client.api('me/calendarView')
      .query({
        startDateTime: startDateTime,
        endDateTime: endDateTime
      })
      .top(count)
      .orderby('start/dateTime')
      .select(config.CALENDAR_SELECT_FIELDS)
      .get();
    
    if (!response.value || response.value.length === 0) {
      return makeResponse('No calendar events found.');
    }

    const text = formatResponse(
      { events: buildEventRows(response.value) },
      buildEventText(response.value)
    );

    return makeResponse(text);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error listing events: ${error.message}`);
  }
}

module.exports = handleListEvents;
