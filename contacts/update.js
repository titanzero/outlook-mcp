const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleUpdateContact(args) {
  const { id, displayName, email, phone, jobTitle, companyName } = args;

  if (!id) return makeErrorResponse('Contact ID is required.');

  try {
    const client = await getGraphClient();

    const patch = {};
    if (displayName !== undefined) patch.displayName = displayName;
    if (email !== undefined) patch.emailAddresses = [{ address: email, name: displayName || email }];
    if (phone !== undefined) patch.mobilePhone = phone;
    if (jobTitle !== undefined) patch.jobTitle = jobTitle;
    if (companyName !== undefined) patch.companyName = companyName;

    await client.api(`me/contacts/${id}`).patch(patch);

    return makeResponse('Contact updated successfully.');
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error updating contact: ${error.message}`);
  }
}

module.exports = handleUpdateContact;
