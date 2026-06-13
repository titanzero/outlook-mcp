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

const calendarTools = [
  {
    name: "list-events",
    description: "Lists upcoming events from your calendar",
    inputSchema: {
      type: "object",
      properties: {
        count: { type: "number", description: "Number of events to retrieve (default: 10, max: 50)" }
      },
      required: []
    },
    handler: handleListEvents
  },
  {
    name: "list-calendars",
    description: "Lists all calendars in your Outlook account",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleListCalendars
  },
  {
    name: "create-event",
    description: "Creates a new calendar event",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "The subject of the event" },
        start: { type: "string", description: "The start time of the event in ISO 8601 format" },
        end: { type: "string", description: "The end time of the event in ISO 8601 format" },
        attendees: { type: "array", items: { type: "string" }, description: "List of attendee email addresses" },
        body: { type: "string", description: "Optional body content for the event" }
      },
      required: ["subject", "start", "end"]
    },
    handler: handleCreateEvent
  },
  {
    name: "update-event",
    description: "Updates an existing calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "ID of the event to update" },
        subject: { type: "string", description: "New subject" },
        start: { type: "string", description: "New start time in ISO 8601 format" },
        end: { type: "string", description: "New end time in ISO 8601 format" },
        location: { type: "string", description: "New location display name" },
        body: { type: "string", description: "New body content" },
        attendees: { type: "array", items: { type: "string" }, description: "Replacement list of attendee email addresses" },
        isAllDay: { type: "boolean", description: "Whether the event is an all-day event" }
      },
      required: ["eventId"]
    },
    handler: handleUpdateEvent
  },
  {
    name: "accept-event",
    description: "Accepts a calendar event invitation",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "The ID of the event to accept" },
        comment: { type: "string", description: "Optional comment" }
      },
      required: ["eventId"]
    },
    handler: handleAcceptEvent
  },
  {
    name: "tentatively-accept-event",
    description: "Tentatively accepts a calendar event invitation",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "The ID of the event to tentatively accept" },
        comment: { type: "string", description: "Optional comment" }
      },
      required: ["eventId"]
    },
    handler: handleTentativelyAcceptEvent
  },
  {
    name: "decline-event",
    description: "Declines a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "The ID of the event to decline" },
        comment: { type: "string", description: "Optional comment" }
      },
      required: ["eventId"]
    },
    handler: handleDeclineEvent
  },
  {
    name: "cancel-event",
    description: "Cancels a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "The ID of the event to cancel" },
        comment: { type: "string", description: "Optional comment" }
      },
      required: ["eventId"]
    },
    handler: handleCancelEvent
  },
  {
    name: "delete-event",
    description: "Deletes a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: { type: "string", description: "The ID of the event to delete" }
      },
      required: ["eventId"]
    },
    handler: handleDeleteEvent
  },
  {
    name: "get-free-busy",
    description: "Returns the free/busy schedule for one or more users",
    inputSchema: {
      type: "object",
      properties: {
        emails: { type: "string", description: "Comma-separated list of email addresses to check" },
        startTime: { type: "string", description: "Start of the time range in ISO 8601 format" },
        endTime: { type: "string", description: "End of the time range in ISO 8601 format" },
        intervalMinutes: { type: "number", description: "Slot size in minutes for availabilityView (default: 30)" }
      },
      required: ["emails", "startTime", "endTime"]
    },
    handler: handleGetFreeBusy
  }
];

module.exports = {
  calendarTools,
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
};
