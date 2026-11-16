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
const { handleShareCalendar, handleListCalendarPermissions, handleUpdateCalendarPermission, handleRemoveCalendarPermission, handleGetSharedCalendars } = require('./sharing');
const { handleListRooms, handleListRoomLists, handleCheckRoomAvailability, handleFindAvailableRooms } = require('./rooms');
const { handleListCalendarGroups, handleCreateCalendarGroup, handleDeleteCalendarGroup, handleAddCalendarToGroup, handleRemoveCalendarFromGroup } = require('./groups');
const { handleGetCalendarView, handleGetMergedCalendarView } = require('./view');
const { handleGetReminders, handleSnoozeReminder, handleDismissReminder, handleSetReminder } = require('./reminders');
const { handleListOnlineMeetingProviders, handleCreateOnlineMeeting, handleGetOnlineMeeting } = require('./online-meetings');
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

  // Calendar Sharing tools (requires Calendars.ReadWrite.Shared)
  server.tool(
    'share-calendar',
    'Shares a calendar with another user (requires Calendars.ReadWrite.Shared scope)',
    {
      calendarId: z.string().describe('Calendar ID (from list-calendars)'),
      emailAddress: z.string().describe('Email address of the person to share with'),
      role: z.enum(['freeBusyRead', 'limitedRead', 'read', 'write', 'delegateWithoutPrivate', 'delegateWithPrivate']).optional().describe('Permission level (default: read)'),
      message: z.string().optional().describe('Optional message to include with the sharing invitation'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleShareCalendar
  );

  server.tool(
    'list-calendar-permissions',
    'Lists permissions for a calendar (requires Calendars.ReadWrite.Shared scope)',
    {
      calendarId: z.string().describe('Calendar ID (from list-calendars)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListCalendarPermissions
  );

  server.tool(
    'update-calendar-permission',
    'Updates a calendar permission (requires Calendars.ReadWrite.Shared scope)',
    {
      calendarId: z.string().describe('Calendar ID'),
      permissionId: z.string().describe('Permission ID (from list-calendar-permissions)'),
      role: z.enum(['freeBusyRead', 'limitedRead', 'read', 'write', 'delegateWithoutPrivate', 'delegateWithPrivate']).describe('New permission level'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateCalendarPermission
  );

  server.tool(
    'remove-calendar-permission',
    'Removes a calendar permission (requires Calendars.ReadWrite.Shared scope)',
    {
      calendarId: z.string().describe('Calendar ID'),
      permissionId: z.string().describe('Permission ID (from list-calendar-permissions)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleRemoveCalendarPermission
  );

  server.tool(
    'get-shared-calendars',
    'Gets calendars shared with you and calendars you own (requires Calendars.Read.Shared scope)',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetSharedCalendars
  );

  // Meeting Rooms / Places tools
  server.tool(
    'list-rooms',
    'Lists meeting rooms in the organization (requires Calendars.Read scope)',
    {
      roomList: z.string().optional().describe('Optional room list ID to filter rooms (from list-room-lists)'),
      count: z.number().int().positive().max(100).optional().describe('Number of rooms to retrieve (default: 100)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListRooms
  );

  server.tool(
    'list-room-lists',
    'Lists room lists (groups of meeting rooms) in the organization',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListRoomLists
  );

  server.tool(
    'check-room-availability',
    'Checks availability of a specific room for a time range',
    {
      roomEmail: z.string().describe('Room email address (from list-rooms)'),
      startTime: z.string().describe('Start time in ISO 8601 format (UTC)'),
      endTime: z.string().describe('End time in ISO 8601 format (UTC)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleCheckRoomAvailability
  );

  server.tool(
    'find-available-rooms',
    'Finds available rooms for a time range with optional capacity filter',
    {
      startTime: z.string().describe('Start time in ISO 8601 format (UTC)'),
      endTime: z.string().describe('End time in ISO 8601 format (UTC)'),
      capacity: z.number().int().positive().optional().describe('Minimum room capacity'),
      roomList: z.string().optional().describe('Optional room list ID to search within'),
      count: z.number().int().positive().max(50).optional().describe('Max results to return (default: 20)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleFindAvailableRooms
  );

  // Calendar Groups tools
  server.tool(
    'list-calendar-groups',
    'Lists calendar groups with their calendars',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListCalendarGroups
  );

  server.tool(
    'create-calendar-group',
    'Creates a new calendar group',
    {
      name: z.string().describe('Name of the calendar group to create'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateCalendarGroup
  );

  server.tool(
    'delete-calendar-group',
    'Deletes a calendar group (calendars are not deleted)',
    {
      groupId: z.string().describe('Calendar group ID (from list-calendar-groups)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteCalendarGroup
  );

  server.tool(
    'add-calendar-to-group',
    'Adds a calendar to a calendar group',
    {
      groupId: z.string().describe('Calendar group ID (from list-calendar-groups)'),
      calendarId: z.string().describe('Calendar ID (from list-calendars)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleAddCalendarToGroup
  );

  server.tool(
    'remove-calendar-from-group',
    'Removes a calendar from a calendar group',
    {
      groupId: z.string().describe('Calendar group ID'),
      calendarId: z.string().describe('Calendar ID'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleRemoveCalendarFromGroup
  );

  // Calendar View (expanded recurring events)
  server.tool(
    'get-calendar-view',
    'Gets expanded calendar view (recurring events expanded) for a date range',
    {
      startTime: z.string().describe('Start time in ISO 8601 format'),
      endTime: z.string().describe('End time in ISO 8601 format'),
      calendarId: z.string().optional().describe('Optional specific calendar ID (default: all calendars)'),
      count: z.number().int().positive().max(1000).optional().describe('Max events to return (default: 100)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetCalendarView
  );

  server.tool(
    'get-merged-calendar-view',
    'Gets merged calendar view across multiple calendars for a date range',
    {
      startTime: z.string().describe('Start time in ISO 8601 format'),
      endTime: z.string().describe('End time in ISO 8601 format'),
      calendarIds: z.array(z.string()).optional().describe('Optional specific calendar IDs (default: all accessible calendars)'),
      count: z.number().int().positive().max(1000).optional().describe('Max events to return (default: 100)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetMergedCalendarView
  );

  // Event Reminders tools
  server.tool(
    'get-reminders',
    'Gets event reminders in a date range',
    {
      startTime: z.string().describe('Start time in ISO 8601 format'),
      endTime: z.string().describe('End time in ISO 8601 format'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetReminders
  );

  server.tool(
    'snooze-reminder',
    'Snoozes an event reminder for specified minutes',
    {
      eventId: z.string().describe('Event ID'),
      snoozeMinutes: z.number().int().positive().max(1440).optional().describe('Minutes to snooze (default: 5, max: 1440)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleSnoozeReminder
  );

  server.tool(
    'dismiss-reminder',
    'Dismisses an event reminder (disables it)',
    {
      eventId: z.string().describe('Event ID'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDismissReminder
  );

  server.tool(
    'set-reminder',
    'Sets or updates an event reminder',
    {
      eventId: z.string().describe('Event ID'),
      minutesBeforeStart: z.number().int().min(0).max(10080).describe('Minutes before start (0 to disable, max: 10080 = 1 week)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetReminder
  );

  // Online Meetings tools
  server.tool(
    'list-online-meeting-providers',
    'Lists configured online meeting providers (Teams, Zoom, etc.)',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListOnlineMeetingProviders
  );

  server.tool(
    'create-online-meeting',
    'Creates an online meeting (Teams, etc.)',
    {
      subject: z.string().describe('Meeting subject'),
      startTime: z.string().describe('Start time in ISO 8601 format (UTC)'),
      endTime: z.string().describe('End time in ISO 8601 format (UTC)'),
      attendees: z.array(z.string()).optional().describe('Attendee email addresses'),
      isEntryExitAnnounced: z.boolean().optional().describe('Announce when people join/leave (default: true)'),
      allowedPresenters: z.enum(['everyone', 'organization', 'specificPeople', 'onlyMe']).optional().describe('Who can present (default: organization)'),
      allowAttendeeToEnableCamera: z.boolean().optional().describe('Allow attendees to turn on camera (default: true)'),
      allowAttendeeToEnableMic: z.boolean().optional().describe('Allow attendees to unmute (default: true)'),
      allowMeetingChat: z.boolean().optional().describe('Allow meeting chat (default: true)'),
      autoAdmittedUsers: z.enum(['everyone', 'organization', 'invited', 'onlyMe']).optional().describe('Who bypasses lobby (default: everyone)'),
      description: z.string().optional().describe('Meeting description'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleCreateOnlineMeeting
  );

  server.tool(
    'get-online-meeting',
    'Gets online meeting details by ID',
    {
      meetingId: z.string().describe('Online meeting ID'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetOnlineMeeting
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
  handleShareCalendar,
  handleListCalendarPermissions,
  handleUpdateCalendarPermission,
  handleRemoveCalendarPermission,
  handleGetSharedCalendars,
  handleListRooms,
  handleListRoomLists,
  handleCheckRoomAvailability,
  handleFindAvailableRooms,
  handleListCalendarGroups,
  handleCreateCalendarGroup,
  handleDeleteCalendarGroup,
  handleAddCalendarToGroup,
  handleRemoveCalendarFromGroup,
  handleGetCalendarView,
  handleGetMergedCalendarView,
  handleGetReminders,
  handleSnoozeReminder,
  handleDismissReminder,
  handleSetReminder,
  handleListOnlineMeetingProviders,
  handleCreateOnlineMeeting,
  handleGetOnlineMeeting,
};
