const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleForwardEmail(args) {
  const { id, to, comment = '' } = args;

  if (!id) return makeErrorResponse('Email ID is required.');
  if (!to) return makeErrorResponse('At least one recipient (to) is required.');

  try {
    const client = await getGraphClient();
    const toRecipients = to.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    }));

    await client.api(`me/messages/${id}/forward`).post({ toRecipients, comment });

    return makeResponse(`Email forwarded successfully to ${toRecipients.map(r => r.emailAddress.address).join(', ')}.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error forwarding email: ${error.message}`);
  }
}

module.exports = handleForwardEmail;
