const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleDeleteContact(args) {
  const { id } = args;

  if (!id) return makeErrorResponse('Contact ID is required.');

  try {
    const client = await getGraphClient();
    await client.api(`me/contacts/${id}`).delete();
    return makeResponse('Contact deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error deleting contact: ${error.message}`);
  }
}

module.exports = handleDeleteContact;
