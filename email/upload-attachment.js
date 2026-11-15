/**
 * Upload attachment to email/draft functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Upload attachment handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUploadAttachment(args) {
  const { id, name, contentBytes, contentType, isInline } = args;

  if (!id) return makeErrorResponse('Email/Draft ID is required.');
  if (!name) return makeErrorResponse('Attachment name is required.');
  if (!contentBytes) return makeErrorResponse('Attachment content (base64) is required.');

  try {
    const client = await getGraphClient();

    const attachment = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name,
      contentBytes,
      contentType: contentType || 'application/octet-stream',
      isInline: isInline || false
    };

    const response = await client.api(`me/messages/${id}/attachments`).post(attachment);

    if (response && response.id) {
      return makeResponse(`Attachment "${name}" uploaded successfully.\nID: ${response.id}\nSize: ${response.size} bytes`);
    } else {
      return makeErrorResponse('Failed to upload attachment.');
    }
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error uploading attachment: ${error.message}`);
  }
}

module.exports = handleUploadAttachment;