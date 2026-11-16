/**
 * Advanced contact search tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Search contacts with advanced filters
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchContacts(args) {
  const { 
    query, 
    folderId, 
    companyName, 
    jobTitle, 
    emailDomain,
    hasPhoto,
    count = 50 
  } = args;
  
  try {
    const client = await getGraphClient();
    
    let endpoint = 'me/contacts';
    if (folderId) {
      endpoint = `me/contactFolders/${folderId}/contacts`;
    }
    
    // Build filter
    const filters = [];
    
    if (query) {
      filters.push(`contains(displayName,'${query}') or contains(givenName,'${query}') or contains(surname,'${query}') or emailAddresses/any(e:contains(e/address,'${query}'))`);
    }
    if (companyName) {
      filters.push(`contains(companyName,'${companyName}')`);
    }
    if (jobTitle) {
      filters.push(`contains(jobTitle,'${jobTitle}')`);
    }
    if (emailDomain) {
      filters.push(`emailAddresses/any(e:endswith(e/address,'@${emailDomain}'))`);
    }
    if (hasPhoto === true) {
      // Check if photo exists - this is tricky with Graph API
      // We'll skip this filter for now as it requires separate API calls
    }
    
    const filterString = filters.length > 0 ? filters.join(' and ') : undefined;
    
    const response = await graphGetPaginated(client, endpoint, {
      $filter: filterString,
      $select: 'id,displayName,givenName,surname,emailAddresses,businessPhones,homePhones,mobilePhone,jobTitle,companyName,department,officeLocation,businessAddress,homeAddress,hasPhoto',
      $orderby: 'displayName asc',
      $top: Math.min(count, 1000),
    });
    
    const contacts = response.value || [];
    
    const structured = contacts.map(c => ({
      id: c.id,
      displayName: c.displayName,
      givenName: c.givenName,
      surname: c.surname,
      emails: c.emailAddresses?.map(e => ({ address: e.address, name: e.name })) || [],
      businessPhones: c.businessPhones || [],
      homePhones: c.homePhones || [],
      mobilePhone: c.mobilePhone,
      jobTitle: c.jobTitle,
      companyName: c.companyName,
      department: c.department,
      officeLocation: c.officeLocation,
      businessAddress: c.businessAddress,
      homeAddress: c.homeAddress,
      hasPhoto: c.hasPhoto || false,
    }));
    
    const textFallback = [
      `Contact Search Results (${contacts.length})`,
      query ? `Query: ${query}` : '',
      companyName ? `Company: ${companyName}` : '',
      jobTitle ? `Job Title: ${jobTitle}` : '',
      emailDomain ? `Email Domain: @${emailDomain}` : '',
      '',
      contacts.length === 0
        ? '  No contacts found matching criteria.'
        : contacts.map((c, i) => {
            const emails = c.emails.map(e => e.address).join(', ') || 'no email';
            const phones = [];
            if (c.mobilePhone) phones.push(`Mobile: ${c.mobilePhone}`);
            if (c.businessPhones?.length) phones.push(`Business: ${c.businessPhones.join(', ')}`);
            const photo = c.hasPhoto ? ' 📷' : '';
            return `${i + 1}. ${c.displayName}${photo} (${emails})${phones.length ? ` - ${phones.join(', ')}` : ''}${c.jobTitle ? `, ${c.jobTitle}` : ''}${c.companyName ? ` @ ${c.companyName}` : ''}${c.department ? ` (${c.department})` : ''}`;
          }).join('\n'),
    ].filter(Boolean).join('\n');
    
    return makeResponse(formatResponse({ contacts: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error searching contacts: ${error.message}`);
  }
}

/**
 * Get contact statistics
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetContactStats(_args) {
  try {
    const client = await getGraphClient();
    
    // Get all contacts (paginated)
    let totalContacts = 0;
    let withEmail = 0;
    let withPhone = 0;
    let withPhoto = 0;
    let withCompany = 0;
    let withJobTitle = 0;
    const companies = new Map();
    const domains = new Map();
    
    const response = await graphGetPaginated(client, 'me/contacts', {
      $select: 'emailAddresses,businessPhones,homePhones,mobilePhone,companyName,jobTitle,hasPhoto',
      $top: 1000,
    });
    
    const contacts = response.value || [];
    totalContacts = contacts.length;
    
    contacts.forEach(c => {
      if (c.emailAddresses?.length) {
        withEmail++;
        c.emailAddresses.forEach(e => {
          const domain = e.address.split('@')[1];
          if (domain) {
            domains.set(domain, (domains.get(domain) || 0) + 1);
          }
        });
      }
      if (c.businessPhones?.length || c.homePhones?.length || c.mobilePhone) {
        withPhone++;
      }
      if (c.hasPhoto) {
        withPhoto++;
      }
      if (c.companyName) {
        withCompany++;
        companies.set(c.companyName, (companies.get(c.companyName) || 0) + 1);
      }
      if (c.jobTitle) {
        withJobTitle++;
      }
    });
    
    // Get folder counts
    const foldersResponse = await client.api('me/contactFolders').select('displayName,contactCount').get();
    const folders = foldersResponse.value || [];
    
    const structured = {
      totalContacts,
      withEmail,
      withPhone,
      withPhoto,
      withCompany,
      withJobTitle,
      folders: folders.map(f => ({
        name: f.displayName,
        count: f.contactCount || 0,
      })),
      topCompanies: Array.from(companies.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([name, count]) => ({ name, count })),
      topDomains: Array.from(domains.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([domain, count]) => ({ domain, count })),
    };
    
    const textFallback = [
      `Contact Statistics`,
      `Total Contacts: ${totalContacts}`,
      `With Email: ${withEmail} (${((withEmail/totalContacts)*100).toFixed(1)}%)`,
      `With Phone: ${withPhone} (${((withPhone/totalContacts)*100).toFixed(1)}%)`,
      `With Photo: ${withPhoto} (${((withPhoto/totalContacts)*100).toFixed(1)}%)`,
      `With Company: ${withCompany} (${((withCompany/totalContacts)*100).toFixed(1)}%)`,
      `With Job Title: ${withJobTitle} (${((withJobTitle/totalContacts)*100).toFixed(1)}%)`,
      '',
      `Folders (${folders.length}):`,
      folders.map(f => `  - ${f.displayName}: ${f.contactCount || 0}`).join('\n') || '  None',
      '',
      `Top Companies:`,
      structured.topCompanies.length > 0
        ? structured.topCompanies.map((c, i) => `  ${i + 1}. ${c.name}: ${c.count}`).join('\n')
        : '  None',
      '',
      `Top Email Domains:`,
      structured.topDomains.length > 0
        ? structured.topDomains.map((d, i) => `  ${i + 1}. @${d.domain}: ${d.count}`).join('\n')
        : '  None',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting contact stats: ${error.message}`);
  }
}

module.exports = {
  handleSearchContacts,
  handleGetContactStats,
};