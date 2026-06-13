const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleTentativelyAcceptEvent(args) {
  const { eventId, comment = '' } = args;

  if (!eventId) return makeErrorResponse('Event ID is required.');

  try {
    const client = await getGraphClient();
    await client.api(`me/events/${eventId}/tentativelyAccept`).post({ comment });
    return makeResponse(`Event tentatively accepted.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error tentatively accepting event: ${error.message}`);
  }
}

module.exports = handleTentativelyAcceptEvent;
