/**
 * Mark email as read functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Mark email as read handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMarkAsRead(args) {
  const emailId = args.id;
  const isRead = args.isRead !== undefined ? args.isRead : true; // Default to true
  
  if (!emailId) {
    return makeErrorResponse('Email ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    const updateData = {
      isRead: isRead
    };
    
    try {
      await client.api(`me/messages/${emailId}`).patch(updateData);
      
      const status = isRead ? 'read' : 'unread';
      
      return makeResponse(`Email successfully marked as ${status}.`);
    } catch (error) {
      console.error(`Error marking email as ${isRead ? 'read' : 'unread'}: ${error.message}`);
      
      // Improved error handling with more specific messages
      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return makeErrorResponse('The email ID seems invalid or doesn\'t belong to your mailbox. Please try with a different email ID.');
      } else if (error.message.includes("UNAUTHORIZED")) {
        return makeErrorResponse('Authentication failed. Please re-authenticate and try again.');
      }
      return makeErrorResponse(`Failed to mark email as ${isRead ? 'read' : 'unread'}: ${error.message}`);
    }
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error accessing email: ${error.message}`);
  }
}

module.exports = handleMarkAsRead;
