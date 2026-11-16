/**
 * Task attachments tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List attachments for a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListTaskAttachments(args) {
  const { listId, taskId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await graphGetPaginated(client, `me/todo/lists/${listId}/tasks/${taskId}/attachments`, {
      $select: 'id,name,contentType,size,lastModifiedDateTime,isInline',
    });
    
    const attachments = response.value || [];
    
    const structured = attachments.map(a => ({
      id: a.id,
      name: a.name,
      contentType: a.contentType,
      size: a.size,
      lastModifiedDateTime: a.lastModifiedDateTime,
      isInline: a.isInline,
    }));
    
    const textFallback = [
      `Task Attachments (${attachments.length}):`,
      '',
      attachments.length === 0
        ? '  No attachments found.'
        : attachments.map((a, i) => 
            `${i + 1}. ${a.name} (${a.contentType}, ${a.size} bytes) ${a.isInline ? '[inline]' : ''} [id: ${a.id}]`
          ).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ attachments: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing task attachments: ${error.message}`);
  }
}

/**
 * Upload attachment to a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUploadTaskAttachment(args) {
  const { listId, taskId, name, contentBytes, contentType, isInline = false } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  if (!name) {
    return makeErrorResponse('Attachment name is required.');
  }
  if (!contentBytes) {
    return makeErrorResponse('Base64-encoded content is required.');
  }
  
  // Decode base64
  let contentBuffer;
  try {
    contentBuffer = Buffer.from(contentBytes, 'base64');
  } catch (e) {
    return makeErrorResponse('Invalid base64 content.');
  }
  
  // Check size (max 25MB for task attachments)
  if (contentBuffer.length > 25 * 1024 * 1024) {
    return makeErrorResponse('Attachment too large. Maximum size is 25MB.');
  }
  
  try {
    const client = await getGraphClient();
    
    const attachment = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name,
      contentBytes: contentBytes,
      contentType: contentType || 'application/octet-stream',
      isInline,
    };
    
    const created = await client.api(`me/todo/lists/${listId}/tasks/${taskId}/attachments`).post(attachment);
    
    return makeResponse(`Task attachment uploaded successfully.\nID: ${created.id}\nName: ${created.name}\nSize: ${created.size} bytes`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error uploading task attachment: ${error.message}`);
  }
}

/**
 * Download task attachment
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDownloadTaskAttachment(args) {
  const { listId, taskId, attachmentId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }
  if (!attachmentId) {
    return makeErrorResponse('attachmentId is required (from list-task-attachments).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get attachment metadata
    const attachment = await client.api(`me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}`).get();
    
    // Get content
    const contentResponse = await client.api(`me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}/$value`).get();
    
    const base64Content = Buffer.from(contentResponse).toString('base64');
    
    const structured = {
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      lastModifiedDateTime: attachment.lastModifiedDateTime,
      isInline: attachment.isInline,
      contentBytes: base64Content,
    };
    
    const textFallback = [
      `Task Attachment: ${attachment.name}`,
      `ID: ${attachment.id}`,
      `Type: ${attachment.contentType}`,
      `Size: ${attachment.size} bytes`,
      `Modified: ${attachment.lastModifiedDateTime}`,
      `Inline: ${attachment.isInline ? 'Yes' : 'No'}`,
      '',
      `Base64 Content: ${base64Content.substring(0, 100)}... (truncated)`,
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Attachment not found.');
    }
    return makeErrorResponse(`Error downloading task attachment: ${error.message}`);
  }
}

/**
 * Delete task attachment
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteTaskAttachment(args) {
  const { listId, taskId, attachmentId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }
  if (!attachmentId) {
    return makeErrorResponse('attachmentId is required (from list-task-attachments).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}`).delete();
    
    return makeResponse('Task attachment deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting task attachment: ${error.message}`);
  }
}

module.exports = {
  handleListTaskAttachments,
  handleUploadTaskAttachment,
  handleDownloadTaskAttachment,
  handleDeleteTaskAttachment,
};