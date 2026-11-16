/**
 * Message recall tool for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Recall a sent message
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRecallMessage(args) {
  const { messageId } = args;
  
  if (!messageId) {
    return makeErrorResponse('Message ID is required (from sent items).');
  }
  
  try {
    const client = await getGraphClient();
    
    // First, verify the message exists and is in sent items
    const message = await client
      .api(`me/messages/${messageId}`)
      .select('id,subject,sentDateTime,isDraft,toRecipients')
      .get();
    
    if (message.isDraft) {
      return makeErrorResponse('Cannot recall a draft message. Only sent messages can be recalled.');
    }
    
    if (!message.sentDateTime) {
      return makeErrorResponse('Message does not appear to be sent yet.');
    }
    
    // Attempt to recall the message
    // Note: Message recall only works for messages sent within the same organization
    // and only if the recipient hasn't read it yet
    await client.api(`me/messages/${messageId}/recall`).post({});
    
    const structured = {
      messageId,
      subject: message.subject,
      sentDateTime: message.sentDateTime,
      status: 'recall_initiated',
      note: 'Recall request sent. Success depends on recipient\'s mailbox settings and whether they have read the message. You will receive a recall status notification.',
    };
    
    const textFallback = [
      `Message Recall Initiated`,
      `Message ID: ${messageId}`,
      `Subject: ${message.subject}`,
      `Sent: ${message.sentDateTime}`,
      '',
      'Recall request has been sent to the recipient\'s mailbox.',
      'Note: Recall only works for messages sent within the same organization.',
      'The recipient must not have read the message yet.',
      'You will receive a notification about the recall status.',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 400) {
      return makeErrorResponse(`Cannot recall this message. Recall is only available for messages sent within your organization to recipients who haven't read them yet.`);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse(`Message not found. The message may have been moved or deleted.`);
    }
    return makeErrorResponse(`Error recalling message: ${error.message}`);
  }
}

/**
 * Get recall status for a message
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetRecallStatus(args) {
  const { messageId } = args;
  
  if (!messageId) {
    return makeErrorResponse('Message ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Check for recall status notifications in the inbox
    const response = await client
      .api('me/messages')
      .filter(`subject eq 'Message Recall: ${messageId}' or contains(subject, 'Recall')`)
      .select('id,subject,receivedDateTime,bodyPreview,isRead')
      .orderby('receivedDateTime desc')
      .top(10)
      .get();
    
    const notifications = response.value || [];
    
    const structured = {
      messageId,
      recallNotifications: notifications.map(n => ({
        id: n.id,
        subject: n.subject,
        receivedDateTime: n.receivedDateTime,
        preview: n.bodyPreview,
        isRead: n.isRead,
      })),
    };
    
    const textFallback = [
      `Recall Status for Message: ${messageId}`,
      '',
      notifications.length === 0
        ? 'No recall status notifications found yet.'
        : notifications.map(n => `- ${n.receivedDateTime}: ${n.subject}\n   ${n.bodyPreview}`).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting recall status: ${error.message}`);
  }
}

module.exports = {
  handleRecallMessage,
  handleGetRecallStatus,
};