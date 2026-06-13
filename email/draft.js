const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

function buildRecipients(csv) {
  if (!csv) return undefined;
  return csv.split(',').map(email => ({ emailAddress: { address: email.trim() } }));
}

async function handleCreateDraft(args) {
  const { to, cc, bcc, subject, body, importance = 'normal' } = args;

  try {
    const client = await getGraphClient();

    const message = {
      subject: subject || '',
      importance,
      body: {
        contentType: body && body.includes('<html') ? 'html' : 'text',
        content: body || '',
      },
    };

    const toRecipients = buildRecipients(to);
    if (toRecipients) message.toRecipients = toRecipients;
    const ccRecipients = buildRecipients(cc);
    if (ccRecipients) message.ccRecipients = ccRecipients;
    const bccRecipients = buildRecipients(bcc);
    if (bccRecipients) message.bccRecipients = bccRecipients;

    const draft = await client.api('me/messages').post(message);

    return makeResponse(`Draft created successfully.\nID: ${draft.id}\nSubject: ${draft.subject || '(no subject)'}`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error creating draft: ${error.message}`);
  }
}

async function handleUpdateDraft(args) {
  const { id, to, cc, bcc, subject, body, importance } = args;

  if (!id) return makeErrorResponse('Draft ID is required.');

  try {
    const client = await getGraphClient();

    const patch = {};
    if (subject !== undefined) patch.subject = subject;
    if (importance !== undefined) patch.importance = importance;
    if (body !== undefined) {
      patch.body = {
        contentType: body.includes('<html') ? 'html' : 'text',
        content: body,
      };
    }
    const toRecipients = buildRecipients(to);
    if (toRecipients) patch.toRecipients = toRecipients;
    const ccRecipients = buildRecipients(cc);
    if (ccRecipients) patch.ccRecipients = ccRecipients;
    const bccRecipients = buildRecipients(bcc);
    if (bccRecipients) patch.bccRecipients = bccRecipients;

    await client.api(`me/messages/${id}`).patch(patch);

    return makeResponse('Draft updated successfully.');
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error updating draft: ${error.message}`);
  }
}

async function handleSendDraft(args) {
  const { id } = args;

  if (!id) return makeErrorResponse('Draft ID is required.');

  try {
    const client = await getGraphClient();
    await client.api(`me/messages/${id}/send`).post({});
    return makeResponse('Draft sent successfully.');
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error sending draft: ${error.message}`);
  }
}

module.exports = { handleCreateDraft, handleUpdateDraft, handleSendDraft };
