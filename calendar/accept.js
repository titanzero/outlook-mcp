/**
 * Accept event functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Accept event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAcceptEvent(args) {
  const { eventId, comment } = args;

  if (!eventId) {
    return makeErrorResponse('Event ID is required to accept an event.');
  }

  try {
    const client = await getGraphClient();

    // Request body
    const body = {
      comment: comment || "Accepted via API"
    };

    // Make API call
    await client.api(`me/events/${eventId}/accept`).post(body);

    return makeResponse(`Event with ID ${eventId} has been successfully accepted.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }

    return makeErrorResponse(`Error accepting event: ${error.message}`);
  }
}

module.exports = handleAcceptEvent;