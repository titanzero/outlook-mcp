const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const config = require('../config');

const CONTACT_SELECT = 'id,displayName,emailAddresses,mobilePhone,businessPhones,jobTitle,companyName';

async function handleListContacts(args) {
  const count = Math.min(args.count || 25, config.MAX_RESULT_COUNT);
  const query = args.query || '';

  try {
    const client = await getGraphClient();

    let request = client.api('me/contacts').select(CONTACT_SELECT).top(count);
    if (query) {
      request = request.filter(
        `startsWith(displayName,'${query}') or emailAddresses/any(e:startsWith(e/address,'${query}'))`
      );
    }

    const response = await request.get();
    const contacts = response.value || [];

    if (contacts.length === 0) return makeResponse('No contacts found.');

    const structured = contacts.map(c => ({
      id: c.id,
      name: c.displayName,
      emails: (c.emailAddresses || []).map(e => e.address),
      phone: c.mobilePhone || (c.businessPhones || [])[0] || null,
      title: c.jobTitle || null,
      company: c.companyName || null,
    }));

    const textFallback = contacts.map(c => {
      const email = (c.emailAddresses || [])[0]?.address || 'no email';
      return `- ${c.displayName} <${email}>${c.jobTitle ? ` — ${c.jobTitle}` : ''}`;
    }).join('\n');

    return makeResponse(formatResponse(structured, `Contacts (${contacts.length}):\n${textFallback}`));
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error listing contacts: ${error.message}`);
  }
}

module.exports = handleListContacts;
