const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const config = require('../config');

async function handleGetFreeBusy(args) {
  const { emails, startTime, endTime, intervalMinutes = 30 } = args;

  if (!emails) return makeErrorResponse('At least one email address is required.');
  if (!startTime) return makeErrorResponse('startTime is required (ISO 8601).');
  if (!endTime) return makeErrorResponse('endTime is required (ISO 8601).');

  try {
    const client = await getGraphClient();

    const schedules = emails.split(',').map(e => e.trim());

    const response = await client.api('me/calendar/getSchedule').post({
      schedules,
      startTime: { dateTime: startTime, timeZone: config.DEFAULT_TIMEZONE },
      endTime: { dateTime: endTime, timeZone: config.DEFAULT_TIMEZONE },
      availabilityViewInterval: intervalMinutes,
    });

    const results = response.value || [];

    const structured = results.map(r => ({
      email: r.scheduleId,
      availabilityView: r.availabilityView,
      items: (r.scheduleItems || []).map(item => ({
        status: item.status,
        start: item.start?.dateTime,
        end: item.end?.dateTime,
        subject: item.subject || null,
      })),
    }));

    const textFallback = results.map(r => {
      const busy = (r.scheduleItems || []).filter(i => i.status !== 'free');
      return `${r.scheduleId}: ${busy.length === 0 ? 'free' : busy.map(i => `${i.status} ${i.start?.dateTime} – ${i.end?.dateTime}`).join('; ')}`;
    }).join('\n');

    return makeResponse(formatResponse(structured, `Free/Busy:\n${textFallback}`));
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error fetching free/busy: ${error.message}`);
  }
}

module.exports = handleGetFreeBusy;
