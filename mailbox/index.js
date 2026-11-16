const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { registerMailboxResources } = require('./resources');
const { handleGetWorkingHours, handleSetWorkingHours, handleClearWorkingHours } = require('./working-hours');
const { handleListDelegates, handleAddDelegate, handleUpdateDelegate, handleRemoveDelegate } = require('./delegates');
const { handleGetLanguageRegion, handleSetLanguage, handleSetDateFormat, handleSetTimeFormat, handleListLocales, handleListDateFormats, handleListTimeFormats } = require('./language-region');
const { handleGetMailTips, handleGetMailboxUsage, handleGetUserProfile } = require('./mail-tips');
const { z } = require('zod');

async function handleGetMailboxSettings(_args) {
  try {
    const client = await getGraphClient();
    const settings = await client.api('me/mailboxSettings').get();

    const structured = {
      timezone: settings.timeZone,
      language: settings.language?.displayName || null,
      dateFormat: settings.dateFormat || null,
      timeFormat: settings.timeFormat || null,
      automaticReplies: {
        status: settings.automaticRepliesSetting?.status,
        scheduledStart: settings.automaticRepliesSetting?.scheduledStartDateTime?.dateTime || null,
        scheduledEnd: settings.automaticRepliesSetting?.scheduledEndDateTime?.dateTime || null,
        internalMessage: settings.automaticRepliesSetting?.internalReplyMessage || null,
        externalMessage: settings.automaticRepliesSetting?.externalReplyMessage || null,
      },
    };

    const ars = structured.automaticReplies;
    const textFallback = [
      `Timezone: ${structured.timezone}`,
      `Language: ${structured.language || 'not set'}`,
      `Auto-replies: ${ars.status}`,
      ars.scheduledStart ? `  From: ${ars.scheduledStart}` : '',
      ars.scheduledEnd ? `  To:   ${ars.scheduledEnd}` : '',
      ars.internalMessage ? `  Internal: ${ars.internalMessage}` : '',
      ars.externalMessage ? `  External: ${ars.externalMessage}` : '',
    ]
      .filter(Boolean)
      .join('\n');

    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error fetching mailbox settings: ${error.message}`);
  }
}

async function handleSetOutOfOffice(_args) {
  const { status, internalMessage, externalMessage, startTime, endTime } = _args;

  if (!status) {
    return makeErrorResponse('status is required: "enabled", "scheduled", or "disabled".');
  }

  const validStatuses = ['enabled', 'scheduled', 'disabled'];
  if (!validStatuses.includes(status)) {
    return makeErrorResponse(`Invalid status. Use one of: ${validStatuses.join(', ')}.`);
  }

  if (status === 'scheduled' && (!startTime || !endTime)) {
    return makeErrorResponse('startTime and endTime are required when status is "scheduled".');
  }

  try {
    const client = await getGraphClient();

    const automaticRepliesSetting = { status };
    if (internalMessage !== undefined) {
      automaticRepliesSetting.internalReplyMessage = internalMessage;
    }
    if (externalMessage !== undefined) {
      automaticRepliesSetting.externalReplyMessage = externalMessage;
    }
    if (startTime) {
      automaticRepliesSetting.scheduledStartDateTime = { dateTime: startTime, timeZone: 'UTC' };
    }
    if (endTime) {
      automaticRepliesSetting.scheduledEndDateTime = { dateTime: endTime, timeZone: 'UTC' };
    }

    await client.api('me/mailboxSettings').patch({ automaticRepliesSetting });

    return makeResponse(`Out-of-office set to "${status}" successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error setting out-of-office: ${error.message}`);
  }
}

/**
 * Register mailbox tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerMailboxTools(server) {
  server.tool(
    'get-mailbox-settings',
    'Returns mailbox settings including timezone, language, and auto-reply (out-of-office) status',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetMailboxSettings
  );

  server.tool(
    'set-out-of-office',
    'Enables, schedules, or disables automatic out-of-office replies',
    {
      status: z.enum(['enabled', 'scheduled', 'disabled']).describe('Auto-reply status'),
      internalMessage: z
        .string()
        .optional()
        .describe('Reply message for people inside your organisation'),
      externalMessage: z
        .string()
        .optional()
        .describe('Reply message for people outside your organisation'),
      startTime: z
        .string()
        .optional()
        .describe("ISO 8601 start time (required when status is 'scheduled')"),
      endTime: z
        .string()
        .optional()
        .describe("ISO 8601 end time (required when status is 'scheduled')"),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetOutOfOffice
  );

  // Working Hours tools
  server.tool(
    'get-working-hours',
    'Gets configured working hours',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetWorkingHours
  );

  server.tool(
    'set-working-hours',
    'Sets working hours (days, time range, time zone)',
    {
      daysOfWeek: z.array(z.enum(['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'])).describe('Working days'),
      startTime: z.string().describe('Start time in HH:mm format (e.g., "09:00")'),
      endTime: z.string().describe('End time in HH:mm format (e.g., "17:00")'),
      timeZone: z.string().optional().describe('Time zone (default: UTC)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetWorkingHours
  );

  server.tool(
    'clear-working-hours',
    'Clears working hours configuration',
    {},
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleClearWorkingHours
  );

  // Delegates tools (requires MailboxSettings.ReadWrite)
  server.tool(
    'list-delegates',
    'Lists all delegates with their permissions',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListDelegates
  );

  server.tool(
    'add-delegate',
    'Adds a delegate with specified permissions',
    {
      emailAddress: z.string().describe('Delegate email address'),
      calendarRole: z.enum(['none', 'reviewer', 'author', 'editor', 'custom']).optional().describe('Calendar permission (default: reviewer)'),
      inboxRole: z.enum(['none', 'reviewer', 'author', 'editor', 'custom']).optional().describe('Inbox permission (default: none)'),
      contactsRole: z.enum(['none', 'reviewer', 'author', 'editor', 'custom']).optional().describe('Contacts permission (default: none)'),
      canViewPrivateItems: z.boolean().optional().describe('Can view private items (default: false)'),
      receiveCopiesOfMeetingMessages: z.boolean().optional().describe('Receive meeting copies (default: false)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleAddDelegate
  );

  server.tool(
    'update-delegate',
    'Updates delegate permissions',
    {
      emailAddress: z.string().describe('Delegate email address'),
      calendarRole: z.enum(['none', 'reviewer', 'author', 'editor', 'custom']).optional().describe('Calendar permission'),
      inboxRole: z.enum(['none', 'reviewer', 'author', 'editor', 'custom']).optional().describe('Inbox permission'),
      contactsRole: z.enum(['none', 'reviewer', 'author', 'editor', 'custom']).optional().describe('Contacts permission'),
      canViewPrivateItems: z.boolean().optional().describe('Can view private items'),
      receiveCopiesOfMeetingMessages: z.boolean().optional().describe('Receive meeting copies'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateDelegate
  );

  server.tool(
    'remove-delegate',
    'Removes a delegate',
    {
      emailAddress: z.string().describe('Delegate email address'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleRemoveDelegate
  );

  // Language & Region tools
  server.tool(
    'get-language-region',
    'Gets language, date format, and time format settings',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetLanguageRegion
  );

  server.tool(
    'set-language',
    'Sets the mailbox language/locale',
    {
      locale: z.string().describe('Locale code (e.g., "en-US", "fr-FR", "de-DE")'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetLanguage
  );

  server.tool(
    'set-date-format',
    'Sets the date format',
    {
      format: z.string().describe('Date format (e.g., "MM/dd/yyyy", "dd/MM/yyyy", "yyyy-MM-dd")'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetDateFormat
  );

  server.tool(
    'set-time-format',
    'Sets the time format',
    {
      format: z.string().describe('Time format (e.g., "H:mm", "h:mm tt", "HH:mm")'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetTimeFormat
  );

  server.tool(
    'list-locales',
    'Lists common locale codes for language setting',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListLocales
  );

  server.tool(
    'list-date-formats',
    'Lists common date format strings',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListDateFormats
  );

  server.tool(
    'list-time-formats',
    'Lists common time format strings',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListTimeFormats
  );

  // Mail Tips & Mailbox Info tools
  server.tool(
    'get-mail-tips',
    'Gets mail tips for recipients (auto-reply, mailbox full, delivery restrictions, etc.)',
    {
      emailAddresses: z.array(z.string()).describe('Array of recipient email addresses (max 20)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: true,
      },
    },
    handleGetMailTips
  );

  server.tool(
    'get-mailbox-usage',
    'Gets mailbox usage statistics and summary',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetMailboxUsage
  );

  server.tool(
    'get-user-profile',
    'Gets extended user profile information',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetUserProfile
  );
}

module.exports = {
  registerMailboxTools,
  registerMailboxResources,
  handleGetMailboxSettings,
  handleSetOutOfOffice,
  handleGetWorkingHours,
  handleSetWorkingHours,
  handleClearWorkingHours,
  handleListDelegates,
  handleAddDelegate,
  handleUpdateDelegate,
  handleRemoveDelegate,
  handleGetLanguageRegion,
  handleSetLanguage,
  handleSetDateFormat,
  handleSetTimeFormat,
  handleListLocales,
  handleListDateFormats,
  handleListTimeFormats,
  handleGetMailTips,
  handleGetMailboxUsage,
  handleGetUserProfile,
};
