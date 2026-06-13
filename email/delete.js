const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleDeleteEmail(args) {
  const { id } = args;

  if (!id) return makeErrorResponse('Email ID is required.');

  try {
    const client = await getGraphClient();
    await client.api(`me/messages/${id}`).delete();
    return makeResponse('Email deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error deleting email: ${error.message}`);
  }
}

module.exports = handleDeleteEmail;
