const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const config = require('../config');

async function handleUpdateEvent(args) {
  const { eventId, subject, start, end, location, body, attendees, isAllDay } = args;

  if (!eventId) return makeErrorResponse('Event ID is required.');

  try {
    const client = await getGraphClient();

    const patch = {};
    if (subject !== undefined) patch.subject = subject;
    if (location !== undefined) patch.location = { displayName: location };
    if (body !== undefined) patch.body = { contentType: 'text', content: body };
    if (isAllDay !== undefined) patch.isAllDay = isAllDay;

    if (start !== undefined) {
      patch.start = { dateTime: start, timeZone: config.DEFAULT_TIMEZONE };
    }
    if (end !== undefined) {
      patch.end = { dateTime: end, timeZone: config.DEFAULT_TIMEZONE };
    }
    if (attendees !== undefined) {
      patch.attendees = attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      }));
    }

    await client.api(`me/events/${eventId}`).patch(patch);

    return makeResponse(`Event updated successfully.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error updating event: ${error.message}`);
  }
}

module.exports = handleUpdateEvent;
