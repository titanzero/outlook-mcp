const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleGetMailboxSettings(args) {
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
    ].filter(Boolean).join('\n');

    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error fetching mailbox settings: ${error.message}`);
  }
}

async function handleSetOutOfOffice(args) {
  const { status, internalMessage, externalMessage, startTime, endTime } = args;

  if (!status) return makeErrorResponse('status is required: "enabled", "scheduled", or "disabled".');

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
    if (internalMessage !== undefined) automaticRepliesSetting.internalReplyMessage = internalMessage;
    if (externalMessage !== undefined) automaticRepliesSetting.externalReplyMessage = externalMessage;
    if (startTime) automaticRepliesSetting.scheduledStartDateTime = { dateTime: startTime, timeZone: 'UTC' };
    if (endTime) automaticRepliesSetting.scheduledEndDateTime = { dateTime: endTime, timeZone: 'UTC' };

    await client.api('me/mailboxSettings').patch({ automaticRepliesSetting });

    return makeResponse(`Out-of-office set to "${status}" successfully.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error setting out-of-office: ${error.message}`);
  }
}

const mailboxTools = [
  {
    name: "get-mailbox-settings",
    description: "Returns mailbox settings including timezone, language, and auto-reply (out-of-office) status",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleGetMailboxSettings
  },
  {
    name: "set-out-of-office",
    description: "Enables, schedules, or disables automatic out-of-office replies",
    inputSchema: {
      type: "object",
      properties: {
        status: {
          type: "string",
          description: "Auto-reply status",
          enum: ["enabled", "scheduled", "disabled"]
        },
        internalMessage: { type: "string", description: "Reply message for people inside your organisation" },
        externalMessage: { type: "string", description: "Reply message for people outside your organisation" },
        startTime: { type: "string", description: "ISO 8601 start time (required when status is 'scheduled')" },
        endTime: { type: "string", description: "ISO 8601 end time (required when status is 'scheduled')" }
      },
      required: ["status"]
    },
    handler: handleSetOutOfOffice
  }
];

module.exports = { mailboxTools };
