const handleListEvents = require('./list');
const handleAcceptEvent = require('./accept');
const handleDeclineEvent = require('./decline');
const handleCreateEvent = require('./create');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');
const handleUpdateEvent = require('./update');
const handleTentativelyAcceptEvent = require('./tentative-accept');
const handleListCalendars = require('./list-calendars');
const handleGetFreeBusy = require('./free-busy');
const handleCreateCalendar = require('./create-calendar');
const handleDeleteCalendar = require('./delete-calendar');
const { registerCalendarResources } = require('./resources');
const { z } = require('zod');

/**
 * Register calendar tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerCalendarTools(server) {
  // Read-only tools
  server.tool(
    'list-events',
    'Lists upcoming events from your calendar',
    {
      count: z
        .number()
        .int()
        .positive()
        .max(50)
        .optional()
        .describe('Number of events to retrieve (default: 10, max: 50)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListEvents
  );

  server.tool(
    'list-calendars',
    'Lists all calendars in your Outlook account',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListCalendars
  );

  server.tool(
    'get-free-busy',
    'Returns the free/busy schedule for one or more users',
    {
      emails: z.string().describe('Comma-separated list of email addresses to check'),
      startTime: z.string().describe('Start of the time range in ISO 8601 format'),
      endTime: z.string().describe('End of the time range in ISO 8601 format'),
      intervalMinutes: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('Slot size in minutes for availabilityView (default: 30)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: true,
      },
    },
    handleGetFreeBusy
  );

  // Calendar management tools
  server.tool(
    'create-calendar',
    'Creates a new calendar',
    {
      name: z.string().describe('Name of the calendar to create'),
      color: z
        .string()
        .optional()
        .describe(
          "Calendar color theme (e.g., 'auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'lightRed', 'maxColor')"
        ),
      isDefault: z.boolean().optional().describe('Whether to set this as the default calendar'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateCalendar
  );

  server.tool(
    'delete-calendar',
    'Deletes a calendar. Cannot delete the default calendar.',
    {
      name: z.string().describe('Name of the calendar to delete'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteCalendar
  );

  // Write tools
  server.tool(
    'create-event',
    'Creates a new calendar event',
    {
      subject: z.string().describe('The subject of the event'),
      start: z.string().describe('The start time of the event in ISO 8601 format'),
      end: z.string().describe('The end time of the event in ISO 8601 format'),
      attendees: z.array(z.string()).optional().describe('List of attendee email addresses'),
      body: z.string().optional().describe('Optional body content for the event'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleCreateEvent
  );

  server.tool(
    'update-event',
    'Updates an existing calendar event',
    {
      eventId: z.string().describe('ID of the event to update'),
      subject: z.string().optional().describe('New subject'),
      start: z.string().optional().describe('New start time in ISO 8601 format'),
      end: z.string().optional().describe('New end time in ISO 8601 format'),
      location: z.string().optional().describe('New location display name'),
      body: z.string().optional().describe('New body content'),
      attendees: z
        .array(z.string())
        .optional()
        .describe('Replacement list of attendee email addresses'),
      isAllDay: z.boolean().optional().describe('Whether the event is an all-day event'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateEvent
  );

  server.tool(
    'accept-event',
    'Accepts a calendar event invitation',
    {
      eventId: z.string().describe('The ID of the event to accept'),
      comment: z.string().optional().describe('Optional comment'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    handleAcceptEvent
  );

  server.tool(
    'tentatively-accept-event',
    'Tentatively accepts a calendar event invitation',
    {
      eventId: z.string().describe('The ID of the event to tentatively accept'),
      comment: z.string().optional().describe('Optional comment'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    handleTentativelyAcceptEvent
  );

  server.tool(
    'decline-event',
    'Declines a calendar event',
    {
      eventId: z.string().describe('The ID of the event to decline'),
      comment: z.string().optional().describe('Optional comment'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    handleDeclineEvent
  );

  server.tool(
    'cancel-event',
    'Cancels a calendar event',
    {
      eventId: z.string().describe('The ID of the event to cancel'),
      comment: z.string().optional().describe('Optional comment'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    handleCancelEvent
  );

  server.tool(
    'delete-event',
    'Deletes a calendar event',
    {
      eventId: z.string().describe('The ID of the event to delete'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteEvent
  );
}

module.exports = {
  registerCalendarTools,
  registerCalendarResources,
  handleListEvents,
  handleAcceptEvent,
  handleDeclineEvent,
  handleCreateEvent,
  handleUpdateEvent,
  handleTentativelyAcceptEvent,
  handleCancelEvent,
  handleDeleteEvent,
  handleListCalendars,
  handleGetFreeBusy,
  handleCreateCalendar,
  handleDeleteCalendar,
};
