/**
 * List emails functionality
 * Uses the Microsoft Graph JS SDK.
 */
const config = require('../config');
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { resolveFolderPath } = require('./folder-utils');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Build a structured array of email rows suitable for TOON encoding.
 * @param {Array} emails - Raw Graph API email objects
 * @returns {Array<object>}
 */
function buildEmailRows(emails) {
  return emails.map((email, index) => {
    const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
    return {
      n: index + 1,
      unread: !email.isRead,
      date: email.receivedDateTime,
      from: sender.name,
      email: sender.address,
      subject: email.subject,
      id: email.id,
    };
  });
}

/**
 * Build human-readable plain-text output for an email list.
 * @param {Array} emails - Raw Graph API email objects
 * @param {string} folder - Folder name for the header
 * @returns {string}
 */
function buildEmailText(emails, folder) {
  const emailList = emails.map((email, index) => {
    const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
    const date = new Date(email.receivedDateTime).toLocaleString();
    const readStatus = email.isRead ? '' : '[UNREAD] ';
    return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
  }).join("\n");

  return `Found ${emails.length} emails in ${folder}:\n\n${emailList}`;
}

/**
 * List emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = Math.min(args.count || 10, config.MAX_RESULT_COUNT);
  
  try {
    const client = await getGraphClient();

    // Resolve the folder path
    const endpoint = await resolveFolderPath(folder);
    
    // Query parameters
    const queryParams = {
      $top: Math.min(config.DEFAULT_PAGE_SIZE, requestedCount),
      $orderby: 'receivedDateTime desc',
      $select: config.EMAIL_SELECT_FIELDS
    };
    
    // Make API call with pagination support
    const response = await graphGetPaginated(client, endpoint, queryParams, requestedCount);
    
    if (!response.value || response.value.length === 0) {
      return makeResponse(`No emails found in ${folder}.`);
    }
    
    const text = formatResponse(
      { emails: buildEmailRows(response.value) },
      buildEmailText(response.value, folder)
    );

    return makeResponse(text);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error listing emails: ${error.message}`);
  }
}

module.exports = handleListEmails;
