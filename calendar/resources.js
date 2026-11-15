/**
 * Calendar resources for MCP
 * Exposes calendars and events as addressable resources
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError } = require('../utils/response-helpers');

/**
 * Register calendar resources with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerCalendarResources(server) {
  // Resource for list of calendars
  server.resource(
    "outlook-calendars",
    "outlook://calendars",
    {
      title: "Calendars",
      description: "List all calendars in the account"
    },
    async (uri) => {
      try {
        const client = await getGraphClient();
        const response = await client.api('me/calendars')
          .select('id,name,color,isDefaultCalendar,canEdit,canShare,canViewPrivateItems,owner')
          .get();
        
        const calendars = response.value || [];
        
        const structured = calendars.map(c => ({
          id: c.id,
          name: c.name,
          color: c.color,
          isDefault: c.isDefaultCalendar,
          canEdit: c.canEdit,
          canShare: c.canShare,
          canViewPrivateItems: c.canViewPrivateItems,
          owner: c.owner?.address || c.owner?.name
        }));
        
        const textFallback = calendars.length === 0
          ? 'No calendars found.'
          : `Calendars (${calendars.length}):\n\n` +
            calendars.map(c => 
              `- ${c.name} ${c.isDefaultCalendar ? '(DEFAULT)' : ''} [${c.color}] [id: ${c.id}]`
            ).join('\n');
        
        return {
          contents: [{
            uri: uri.toString(),
            mimeType: "application/json",
            text: JSON.stringify({ calendars: structured }, null, 2)
          }]
        };
      } catch (error) {
        if (isAuthError(error)) throw error;
        throw new Error(`Error reading calendars: ${error.message}`);
      }
    }
  );

  // Resource for individual calendar
  server.resource(
    "outlook-calendar",
    "outlook://calendars/{calendarId}",
    {
      title: "Calendar",
      description: "Access a specific calendar with its events"
    },
    async (uri, { calendarId }) => {
      try {
        const client = await getGraphClient();
        
        // Get calendar info
        const calendar = await client.api(`me/calendars/${calendarId}`).get();
        
        // Get upcoming events
        const response = await graphGetPaginated(client, `me/calendars/${calendarId}/events`, {
          $select: 'id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,importance',
          $orderby: 'start/dateTime asc',
          $filter: `end/dateTime ge '${new Date().toISOString()}'`
        });
        
        const events = response.value || [];
        
        const structured = {
          id: calendar.id,
          name: calendar.name,
          color: calendar.color,
          isDefault: calendar.isDefaultCalendar,
          canEdit: calendar.canEdit,
          canShare: calendar.canShare,
          canViewPrivateItems: calendar.canViewPrivateItems,
          owner: calendar.owner?.address || calendar.owner?.name,
          upcomingEvents: events.map(e => ({
            id: e.id,
            subject: e.subject,
            start: e.start?.dateTime,
            startTimeZone: e.start?.timeZone,
            end: e.end?.dateTime,
            endTimeZone: e.end?.timeZone,
            location: e.location?.displayName,
            isAllDay: e.isAllDay,
            isCancelled: e.isCancelled,
            importance: e.importance,
            organizer: e.organizer?.emailAddress?.address,
            attendees: e.attendees?.map(a => a.emailAddress?.address).filter(Boolean) || []
          }))
        };
        
        const textFallback = [
          `Calendar: ${calendar.name}`,
          `ID: ${calendar.id}`,
          calendar.isDefaultCalendar ? '(DEFAULT CALENDAR)' : '',
          `Color: ${calendar.color}`,
          `Upcoming Events: ${events.length}`,
          '',
          events.map(e => 
            `- ${e.subject} | ${e.start?.dateTime} to ${e.end?.dateTime} | ${e.location?.displayName || 'no location'} [id: ${e.id}]`
          ).join('\n')
        ].filter(Boolean).join('\n');
        
        return {
          contents: [{
            uri: uri.toString(),
            mimeType: "application/json",
            text: JSON.stringify(structured, null, 2)
          }]
        };
      } catch (error) {
        if (isAuthError(error)) throw error;
        throw new Error(`Error reading calendar: ${error.message}`);
      }
    }
  );

  // Resource for events in date range
  server.resource(
    "outlook-events",
    "outlook://events",
    {
      title: "Events (date range)",
      description: "Get events across all calendars for a date range"
    },
    async (uri) => {
      try {
        const client = await getGraphClient();
        
        // Get events from default calendar for next 30 days
        const start = new Date().toISOString();
        const end = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
        
        const response = await graphGetPaginated(client, 'me/events', {
          $select: 'id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,importance',
          $orderby: 'start/dateTime asc',
          $filter: `start/dateTime ge '${start}' and end/dateTime le '${end}'`
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
          isAllDay: e.isAllDay,
          isCancelled: e.isCancelled,
          importance: e.importance,
          organizer: e.organizer?.emailAddress?.address,
          attendees: e.attendees?.map(a => a.emailAddress?.address).filter(Boolean) || []
        }));
        
        const textFallback = events.length === 0
          ? 'No events in the next 30 days.'
          : `Events (next 30 days, ${events.length}):\n\n` +
            events.map(e => 
              `- ${e.subject} | ${e.start?.dateTime} to ${e.end?.dateTime} | ${e.location?.displayName || 'no location'} [id: ${e.id}]`
            ).join('\n');
        
        return {
          contents: [{
            uri: uri.toString(),
            mimeType: "application/json",
            text: JSON.stringify({ events: structured }, null, 2)
          }]
        };
      } catch (error) {
        if (isAuthError(error)) throw error;
        throw new Error(`Error reading events: ${error.message}`);
      }
    }
  );

  // Resource for individual event
  server.resource(
    "outlook-event",
    "outlook://events/{eventId}",
    {
      title: "Event",
      description: "Access a specific event by ID"
    },
    async (uri, { eventId }) => {
      try {
        const client = await getGraphClient();
        const event = await client.api(`me/events/${eventId}`)
          .select('id,subject,body,start,end,location,organizer,attendees,isAllDay,isCancelled,importance,categories,recurrence,onlineMeeting')
          .get();
        
        const structured = {
          id: event.id,
          subject: event.subject,
          body: event.body?.content,
          bodyPreview: event.bodyPreview,
          start: event.start?.dateTime,
          startTimeZone: event.start?.timeZone,
          end: event.end?.dateTime,
          endTimeZone: event.end?.timeZone,
          location: event.location?.displayName,
          isAllDay: event.isAllDay,
          isCancelled: event.isCancelled,
          importance: event.importance,
          categories: event.categories || [],
          organizer: event.organizer?.emailAddress?.address,
          attendees: event.attendees?.map(a => ({
            address: a.emailAddress?.address,
            name: a.emailAddress?.name,
            status: a.status?.response,
            type: a.type
          })) || [],
          recurrence: event.recurrence,
          onlineMeeting: event.onlineMeeting
        };
        
        const textFallback = [
          `Event: ${event.subject}`,
          `ID: ${event.id}`,
          `When: ${event.start?.dateTime} (${event.start?.timeZone}) to ${event.end?.dateTime} (${event.end?.timeZone})`,
          event.location?.displayName ? `Where: ${event.location.displayName}` : '',
          event.isAllDay ? 'All day' : '',
          event.isCancelled ? 'CANCELLED' : '',
          `Importance: ${event.importance}`,
          event.categories?.length ? `Categories: ${event.categories.join(', ')}` : '',
          `Organizer: ${event.organizer?.emailAddress?.address || 'unknown'}`,
          event.attendees?.length ? `Attendees: ${event.attendees.map(a => `${a.emailAddress?.address} (${a.status?.response || 'none'})`).join(', ')}` : '',
          event.bodyPreview ? `\nPreview: ${event.bodyPreview}` : ''
        ].filter(Boolean).join('\n');
        
        return {
          contents: [{
            uri: uri.toString(),
            mimeType: "application/json",
            text: JSON.stringify(structured, null, 2)
          }]
        };
      } catch (error) {
        if (isAuthError(error)) throw error;
        throw new Error(`Error reading event: ${error.message}`);
      }
    }
  );
}

module.exports = { registerCalendarResources };