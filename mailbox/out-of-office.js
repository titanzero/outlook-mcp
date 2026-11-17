const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

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

module.exports = {
  handleSetOutOfOffice,
};
