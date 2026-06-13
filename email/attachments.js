const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleListAttachments(args) {
  const { id } = args;

  if (!id) return makeErrorResponse('Email ID is required.');

  try {
    const client = await getGraphClient();
    const response = await client
      .api(`me/messages/${id}/attachments`)
      .select('id,name,size,contentType,isInline')
      .get();

    const attachments = response.value || [];

    if (attachments.length === 0) {
      return makeResponse('This email has no attachments.');
    }

    const structured = attachments.map(a => ({
      id: a.id,
      name: a.name,
      contentType: a.contentType,
      size: a.size,
      isInline: a.isInline,
    }));

    const textFallback = attachments
      .map((a, i) => `${i + 1}. ${a.name} (${a.contentType}, ${Math.round(a.size / 1024)} KB) [id: ${a.id}]`)
      .join('\n');

    return makeResponse(formatResponse(structured, `Attachments (${attachments.length}):\n${textFallback}`));
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error listing attachments: ${error.message}`);
  }
}

async function handleDownloadAttachment(args) {
  const { id, attachmentId } = args;

  if (!id) return makeErrorResponse('Email ID is required.');
  if (!attachmentId) return makeErrorResponse('Attachment ID is required.');

  try {
    const client = await getGraphClient();
    const attachment = await client
      .api(`me/messages/${id}/attachments/${attachmentId}`)
      .get();

    const structured = {
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      content: attachment.contentBytes || null,
    };

    const textFallback = [
      `Name: ${attachment.name}`,
      `Type: ${attachment.contentType}`,
      `Size: ${Math.round(attachment.size / 1024)} KB`,
      attachment.contentBytes ? `Content (base64):\n${attachment.contentBytes}` : '',
    ].filter(Boolean).join('\n');

    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error downloading attachment: ${error.message}`);
  }
}

module.exports = { handleListAttachments, handleDownloadAttachment };
