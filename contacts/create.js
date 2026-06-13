const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleCreateContact(args) {
  const { displayName, email, phone, jobTitle, companyName } = args;

  if (!displayName) return makeErrorResponse('displayName is required.');

  try {
    const client = await getGraphClient();

    const contact = { displayName };
    if (email) contact.emailAddresses = [{ address: email, name: displayName }];
    if (phone) contact.mobilePhone = phone;
    if (jobTitle) contact.jobTitle = jobTitle;
    if (companyName) contact.companyName = companyName;

    const created = await client.api('me/contacts').post(contact);

    return makeResponse(`Contact created successfully.\nID: ${created.id}\nName: ${created.displayName}`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error creating contact: ${error.message}`);
  }
}

module.exports = handleCreateContact;
