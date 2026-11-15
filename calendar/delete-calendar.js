/**
 * Delete calendar functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Delete calendar handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteCalendar(args) {
  const { name } = args;

  if (!name) {
    return makeErrorResponse('Calendar name is required.');
  }

  try {
    const client = await getGraphClient();

    // Find the calendar by name
    const response = await client.api('me/calendars').select('id,name,isDefaultCalendar').get();
    const calendars = response.value || [];

    const calendar = calendars.find(c => c.name === name);
    if (!calendar) {
      return makeErrorResponse(
        `Calendar "${name}" not found. Use 'list-calendars' to see available calendars.`
      );
    }

    // Prevent deletion of default calendar
    if (calendar.isDefaultCalendar) {
      return makeErrorResponse(
        `Cannot delete the default calendar "${name}". You must first set another calendar as default.`
      );
    }

    // Delete the calendar
    await client.api(`me/calendars/${calendar.id}`).delete();

    return makeResponse(`Successfully deleted calendar "${name}".`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }

    console.error(`Error deleting calendar "${name}": ${error.message}`);
    return makeErrorResponse(`Error deleting calendar: ${error.message}`);
  }
}

module.exports = handleDeleteCalendar;
