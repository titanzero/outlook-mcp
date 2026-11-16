/**
 * Event reminders tools for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Get reminders for events
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetReminders(args) {
  const { startTime, endTime } = args;
  
  if (!startTime || !endTime) {
    return makeErrorResponse('startTime and endTime are required (ISO 8601 format).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get events with reminders in the time range
    const response = await client.api('me/events').get({
      $filter: `start/dateTime ge '${startTime}' and end/dateTime le '${endTime}' and isReminderOn eq true`,
      $select: 'id,subject,start,end,isReminderOn,reminderMinutesBeforeStart,location,organizer',
      $orderby: 'start/dateTime asc',
    });
    
    const events = response.value || [];
    
    const structured = events.map(e => ({
      id: e.id,
      subject: e.subject,
      start: e.start?.dateTime,
      startTimeZone: e.start?.timeZone,
      end: e.end?.dateTime,
      endTimeZone: e.end?.timeZone,
      location: e.location?.displayName,
      reminderMinutesBeforeStart: e.reminderMinutesBeforeStart,
      reminderTime: e.start?.dateTime ? new Date(new Date(e.start.dateTime).getTime() - (e.reminderMinutesBeforeStart || 15) * 60000).toISOString() : null,
      organizer: e.organizer?.emailAddress?.address,
    }));
    
    const textFallback = [
      `Event Reminders (${events.length})`,
      `Range: ${startTime} to ${endTime}`,
      '',
      events.length === 0
        ? '  No events with reminders in this time range.'
        : events.map((e, i) => {
            const reminderStr = e.reminderMinutesBeforeStart
              ? `${e.reminderMinutesBeforeStart} min before`
              : 'default';
            return `${i + 1}. ${e.subject}\n   When: ${e.start?.dateTime} (${e.start?.timeZone})\n   Reminder: ${reminderStr} (at ${e.reminderTime || 'N/A'})\n   Where: ${e.location?.displayName || 'No location'}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ reminders: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting reminders: ${error.message}`);
  }
}

/**
 * Snooze/dismiss reminder for an event
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSnoozeReminder(args) {
  const { eventId, snoozeMinutes = 5 } = args;
  
  if (!eventId) {
    return makeErrorResponse('Event ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // For snoozing, we need to update the event's reminder time
    // This is a bit tricky - we can set reminderMinutesBeforeStart
    // But snoozing typically means "remind me again in X minutes from now"
    // We'll update the reminder to fire at (now + snoozeMinutes)
    
    const event = await client.api(`me/events/${eventId}`).select('start,reminderMinutesBeforeStart').get();
    
    if (!event.start?.dateTime) {
      return makeErrorResponse('Event has no start time.');
    }
    
    const now = new Date();
    const eventStart = new Date(event.start.dateTime);
    const snoozeTime = new Date(now.getTime() + snoozeMinutes * 60000);
    
    // If snooze time is after event start, just dismiss
    if (snoozeTime >= eventStart) {
      await client.api(`me/events/${eventId}`).patch({ isReminderOn: false });
      return makeResponse(`Reminder dismissed (snooze time would be after event start).`);
    }
    
    // Calculate minutes before start
    const minutesBeforeStart = Math.round((eventStart - snoozeTime) / 60000);
    
    await client.api(`me/events/${eventId}`).patch({
      isReminderOn: true,
      reminderMinutesBeforeStart: Math.max(1, minutesBeforeStart),
    });
    
    return makeResponse(`Reminder snoozed for ${snoozeMinutes} minutes (will fire ${minutesBeforeStart} minutes before event start).`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error snoozing reminder: ${error.message}`);
  }
}

/**
 * Dismiss reminder for an event
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDismissReminder(args) {
  const { eventId } = args;
  
  if (!eventId) {
    return makeErrorResponse('Event ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/events/${eventId}`).patch({ isReminderOn: false });
    
    return makeResponse('Reminder dismissed successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error dismissing reminder: ${error.message}`);
  }
}

/**
 * Set/update reminder for an event
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSetReminder(args) {
  const { eventId, minutesBeforeStart } = args;
  
  if (!eventId) {
    return makeErrorResponse('Event ID is required.');
  }
  if (!minutesBeforeStart || minutesBeforeStart < 0) {
    return makeErrorResponse('minutesBeforeStart is required (0 to disable, positive integer for minutes).');
  }
  
  try {
    const client = await getGraphClient();
    
    const patch = {
      isReminderOn: minutesBeforeStart > 0,
    };
    
    if (minutesBeforeStart > 0) {
      patch.reminderMinutesBeforeStart = minutesBeforeStart;
    }
    
    await client.api(`me/events/${eventId}`).patch(patch);
    
    return makeResponse(minutesBeforeStart > 0
      ? `Reminder set to ${minutesBeforeStart} minutes before event start.`
      : 'Reminder disabled for this event.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error setting reminder: ${error.message}`);
  }
}

module.exports = {
  handleGetReminders,
  handleSnoozeReminder,
  handleDismissReminder,
  handleSetReminder,
};