/**
 * Flag/unflag email functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Flag email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleFlagEmail(args) {
  const { id, flagStatus, dueDateTime } = args;

  if (!id) return makeErrorResponse('Email ID is required.');

  try {
    const client = await getGraphClient();

    const patch = {};
    
    if (flagStatus === 'flagged') {
      patch.flag = {
        flagStatus: 'flagged',
        startDateTime: dueDateTime ? { dateTime: dueDateTime, timeZone: 'UTC' } : { dateTime: new Date().toISOString(), timeZone: 'UTC' },
        dueDateTime: dueDateTime ? { dateTime: dueDateTime, timeZone: 'UTC' } : { dateTime: new Date().toISOString(), timeZone: 'UTC' },
        completedDateTime: null
      };
    } else if (flagStatus === 'completed') {
      patch.flag = {
        flagStatus: 'completed',
        completedDateTime: { dateTime: new Date().toISOString(), timeZone: 'UTC' }
      };
    } else if (flagStatus === 'cleared') {
      patch.flag = {
        flagStatus: 'notFlagged'
      };
    } else {
      return makeErrorResponse('Invalid flagStatus. Use: "flagged", "completed", or "cleared".');
    }

    await client.api(`me/messages/${id}`).patch(patch);

    const statusText = flagStatus === 'cleared' ? 'cleared' : flagStatus;
    return makeResponse(`Email flag ${statusText} successfully.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error updating email flag: ${error.message}`);
  }
}

module.exports = handleFlagEmail;