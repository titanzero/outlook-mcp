/**
 * Delete attachment from email/draft functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Delete attachment handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteAttachment(args) {
  const { id, attachmentId } = args;

  if (!id) {
    return makeErrorResponse('Email/Draft ID is required.');
  }
  if (!attachmentId) {
    return makeErrorResponse('Attachment ID is required.');
  }

  try {
    const client = await getGraphClient();

    await client.api(`me/messages/${id}/attachments/${attachmentId}`).delete();

    return makeResponse('Attachment deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting attachment: ${error.message}`);
  }
}

module.exports = handleDeleteAttachment;
