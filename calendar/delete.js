/**
 * Delete event functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Delete event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteEvent(args) {
  const { eventId } = args;

  if (!eventId) {
    return makeErrorResponse('Event ID is required to delete an event.');
  }

  try {
    const client = await getGraphClient();

    // Make API call
    await client.api(`me/events/${eventId}`).delete();

    return makeResponse(`Event with ID ${eventId} has been successfully deleted.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }

    return makeErrorResponse(`Error deleting event: ${error.message}`);
  }
}

module.exports = handleDeleteEvent;