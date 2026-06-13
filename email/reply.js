const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleReplyEmail(args) {
  const { id, comment = '', replyAll = false } = args;

  if (!id) return makeErrorResponse('Email ID is required.');

  try {
    const client = await getGraphClient();
    const action = replyAll ? 'replyAll' : 'reply';
    await client.api(`me/messages/${id}/${action}`).post({ comment });
    return makeResponse(`Email ${replyAll ? 'replied to all' : 'replied'} successfully.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error replying to email: ${error.message}`);
  }
}

module.exports = handleReplyEmail;
