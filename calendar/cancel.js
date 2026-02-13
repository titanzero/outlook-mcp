/**
 * Cancel event functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Cancel event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCancelEvent(args) {
  const { eventId, comment } = args;

  if (!eventId) {
    return makeErrorResponse('Event ID is required to cancel an event.');
  }

  try {
    const client = await getGraphClient();

    // Request body
    const body = {
      comment: comment || "Cancelled via API"
    };

    // Make API call
    await client.api(`me/events/${eventId}/cancel`).post(body);

    return makeResponse(`Event with ID ${eventId} has been successfully cancelled.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }

    return makeErrorResponse(`Error cancelling event: ${error.message}`);
  }
}

module.exports = handleCancelEvent;