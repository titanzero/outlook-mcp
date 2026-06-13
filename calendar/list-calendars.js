const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleListCalendars(args) {
  try {
    const client = await getGraphClient();
    const response = await client
      .api('me/calendars')
      .select('id,name,color,isDefaultCalendar,canEdit,owner')
      .get();

    const calendars = response.value || [];

    if (calendars.length === 0) return makeResponse('No calendars found.');

    const structured = calendars.map(c => ({
      id: c.id,
      name: c.name,
      color: c.color,
      isDefault: c.isDefaultCalendar,
      canEdit: c.canEdit,
      owner: c.owner?.address || null,
    }));

    const textFallback = calendars
      .map(c => `- ${c.name}${c.isDefaultCalendar ? ' [default]' : ''}${c.canEdit ? '' : ' [read-only]'} (id: ${c.id})`)
      .join('\n');

    return makeResponse(formatResponse(structured, `Calendars (${calendars.length}):\n${textFallback}`));
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error listing calendars: ${error.message}`);
  }
}

module.exports = handleListCalendars;
