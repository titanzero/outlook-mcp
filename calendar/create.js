/**
 * Create event functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { DEFAULT_TIMEZONE } = require('../config');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Create event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateEvent(args) {
  const { subject, start, end, attendees, body } = args;

  if (!subject || !start || !end) {
    return makeErrorResponse('Subject, start, and end times are required to create an event.');
  }

  try {
    const client = await getGraphClient();

    // Request body
    const bodyContent = {
      subject,
      start: { dateTime: start.dateTime || start, timeZone: start.timeZone || DEFAULT_TIMEZONE },
      end: { dateTime: end.dateTime || end, timeZone: end.timeZone || DEFAULT_TIMEZONE },
      attendees: attendees?.map(email => ({ emailAddress: { address: email }, type: "required" })),
      body: { contentType: "HTML", content: body || "" }
    };

    // Make API call
    await client.api('me/events').post(bodyContent);

    return makeResponse(`Event '${subject}' has been successfully created.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }

    return makeErrorResponse(`Error creating event: ${error.message}`);
  }
}

module.exports = handleCreateEvent;