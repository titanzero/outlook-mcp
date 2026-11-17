const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

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

module.exports = {
  handleGetMailboxSettings,
};
