/**
 * Calendar view tools for Outlook MCP server
 * Provides expanded recurring events view
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Get calendar view (expanded recurring events) for a date range
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetCalendarView(args) {
  const { startTime, endTime, calendarId, count = 100 } = args;
  
  if (!startTime || !endTime) {
    return makeErrorResponse('startTime and endTime are required (ISO 8601 format).');
  }
  
  try {
    const client = await getGraphClient();
    
    let endpoint;
    if (calendarId) {
      endpoint = `me/calendars/${calendarId}/calendarView`;
    } else {
      endpoint = 'me/calendarView';
    }
    
    const response = await graphGetPaginated(client, endpoint, {
      startTime,
      endTime,
      $select: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled,importance,categories,recurrence,onlineMeeting,showAs,sensitivity',
      $orderby: 'start/dateTime asc',
      $top: Math.min(count, 1000),
    });
    
    const events = response.value || [];
    
    const structured = events.map(e => ({
      id: e.id,
      subject: e.subject,
      bodyPreview: e.bodyPreview,
      start: e.start?.dateTime,
      startTimeZone: e.start?.timeZone,
      end: e.end?.dateTime,
      endTimeZone: e.end?.timeZone,
      location: e.location?.displayName,
      isAllDay: e.isAllDay,
      isCancelled: e.isCancelled,
      importance: e.importance,
      categories: e.categories || [],
      showAs: e.showAs,
      sensitivity: e.sensitivity,
      organizer: e.organizer?.emailAddress?.address,
      attendees: e.attendees?.map(a => ({
        address: a.emailAddress?.address,
        name: a.emailAddress?.name,
        status: a.status?.response,
        type: a.type,
      })) || [],
      recurrence: e.recurrence ? {
        pattern: e.recurrence.pattern,
        range: e.recurrence.range,
      } : null,
      onlineMeeting: e.onlineMeeting ? {
        joinUrl: e.onlineMeeting.joinUrl,
        conferenceId: e.onlineMeeting.conferenceId,
      } : null,
    }));
    
    const textFallback = [
      `Calendar View (${events.length} events)`,
      `Range: ${startTime} to ${endTime}${calendarId ? `, Calendar: ${calendarId}` : ''}`,
      '',
      events.length === 0
        ? '  No events in this time range.'
        : events.map((e, i) => {
            const allDay = e.isAllDay ? ' [ALL DAY]' : '';
            const cancelled = e.isCancelled ? ' [CANCELLED]' : '';
            const recur = e.recurrence ? ' 🔄' : '';
            const online = e.onlineMeeting ? ' 📹' : '';
            const cats = e.categories?.length ? ` [${e.categories.join(', ')}]` : '';
            return `${i + 1}. ${e.start?.dateTime} - ${e.end?.dateTime}${allDay}${cancelled}${recur}${online}${cats}\n   ${e.subject}\n   ${e.location?.displayName || 'No location'}\n   ${e.organizer?.emailAddress?.address ? `Organizer: ${e.organizer.emailAddress.address}` : ''}${e.attendees?.length ? ` | Attendees: ${e.attendees.length}` : ''}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ events: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting calendar view: ${error.message}`);
  }
}

/**
 * Get calendar view for multiple calendars (merged)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetMergedCalendarView(args) {
  const { startTime, endTime, calendarIds, count = 100 } = args;
  
  if (!startTime || !endTime) {
    return makeErrorResponse('startTime and endTime are required (ISO 8601 format).');
  }
  
  try {
    const client = await getGraphClient();
    
    let calendarsToQuery = [];
    
    if (calendarIds && calendarIds.length > 0) {
      calendarsToQuery = calendarIds;
    } else {
      // Get all calendars the user can access
      const response = await client.api('me/calendars').select('id').get();
      calendarsToQuery = (response.value || []).map(c => c.id);
    }
    
    // Fetch events from each calendar and merge
    const allEvents = [];
    
    for (const calId of calendarsToQuery) {
      try {
        const response = await graphGetPaginated(client, `me/calendars/${calId}/calendarView`, {
          startTime,
          endTime,
          $select: 'id,subject,bodyPreview,start,end,location,organizer,isAllDay,isCancelled,importance,categories,recurrence,onlineMeeting,calendar',
          $orderby: 'start/dateTime asc',
          $top: Math.min(count, 1000),
        });
        
        const events = response.value || [];
        // Add calendar info to each event
        const calInfo = await client.api(`me/calendars/${calId}`).select('id,name,color').get();
        events.forEach(e => {
          e.calendar = { id: calInfo.id, name: calInfo.name, color: calInfo.color };
        });
        
        allEvents.push(...events);
      } catch (e) {
        // Skip calendars we can't access
      }
    }
    
    // Sort by start time
    allEvents.sort((a, b) => new Date(a.start?.dateTime) - new Date(b.start?.dateTime));
    
    const structured = allEvents.slice(0, count).map(e => ({
      id: e.id,
      subject: e.subject,
      bodyPreview: e.bodyPreview,
      start: e.start?.dateTime,
      startTimeZone: e.start?.timeZone,
      end: e.end?.dateTime,
      endTimeZone: e.end?.timeZone,
      location: e.location?.displayName,
      isAllDay: e.isAllDay,
      isCancelled: e.isCancelled,
      importance: e.importance,
      categories: e.categories || [],
      calendar: e.calendar,
      recurrence: e.recurrence ? {
        pattern: e.recurrence.pattern,
        range: e.recurrence.range,
      } : null,
      onlineMeeting: e.onlineMeeting ? {
        joinUrl: e.onlineMeeting.joinUrl,
        conferenceId: e.onlineMeeting.conferenceId,
      } : null,
    }));
    
    const textFallback = [
      `Merged Calendar View (${structured.length} events from ${calendarsToQuery.length} calendars)`,
      `Range: ${startTime} to ${endTime}`,
      '',
      structured.length === 0
        ? '  No events in this time range.'
        : structured.map((e, i) => {
            const allDay = e.isAllDay ? ' [ALL DAY]' : '';
            const cancelled = e.isCancelled ? ' [CANCELLED]' : '';
            const recur = e.recurrence ? ' 🔄' : '';
            const online = e.onlineMeeting ? ' 📹' : '';
            const cats = e.categories?.length ? ` [${e.categories.join(', ')}]` : '';
            const cal = e.calendar ? ` (${e.calendar.name})` : '';
            return `${i + 1}. ${e.start?.dateTime} - ${e.end?.dateTime}${allDay}${cancelled}${recur}${online}${cats}${cal}\n   ${e.subject}\n   ${e.location?.displayName || 'No location'}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ events: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting merged calendar view: ${error.message}`);
  }
}

module.exports = {
  handleGetCalendarView,
  handleGetMergedCalendarView,
};