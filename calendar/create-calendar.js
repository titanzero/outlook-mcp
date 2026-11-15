/**
 * Create calendar functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Create calendar handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateCalendar(args) {
  const { name, color, isDefault } = args;

  if (!name) {
    return makeErrorResponse('Calendar name is required.');
  }

  try {
    const client = await getGraphClient();

    const calendarData = { name };
    if (color) calendarData.color = color;
    if (isDefault !== undefined) calendarData.isDefaultCalendar = isDefault;

    const response = await client.api('me/calendars').post(calendarData);

    if (response && response.id) {
      return makeResponse(`Successfully created calendar "${name}".\nID: ${response.id}`);
    } else {
      return makeErrorResponse('Failed to create calendar. The server didn\'t return a calendar ID.');
    }
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }

    console.error(`Error creating calendar "${name}": ${error.message}`);
    return makeErrorResponse(`Error creating calendar: ${error.message}`);
  }
}

module.exports = handleCreateCalendar;