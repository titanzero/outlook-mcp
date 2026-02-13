/**
 * Decline event functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Decline event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeclineEvent(args) {
  const { eventId, comment } = args;

  if (!eventId) {
    return makeErrorResponse('Event ID is required to decline an event.');
  }

  try {
    const client = await getGraphClient();

    // Request body
    const body = {
      comment: comment || "Declined via API"
    };

    // Make API call
    await client.api(`me/events/${eventId}/decline`).post(body);

    return makeResponse(`Event with ID ${eventId} has been successfully declined.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }

    return makeErrorResponse(`Error declining event: ${error.message}`);
  }
}

module.exports = handleDeclineEvent;