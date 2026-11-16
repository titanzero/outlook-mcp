/**
 * KQL (Keyword Query Language) search tool for Outlook MCP server
 * Provides advanced search capabilities using Microsoft Graph KQL syntax
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Search emails using KQL syntax
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleKqlSearch(args) {
  const { query, folder, count = 50 } = args;
  
  if (!query) {
    return makeErrorResponse('KQL query is required. See https://learn.microsoft.com/en-us/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference for syntax.');
  }
  
  try {
    const client = await getGraphClient();
    
    let endpoint;
    if (folder) {
      endpoint = `me/mailFolders/${folder}/messages`;
    } else {
      endpoint = 'me/messages';
    }
    
    const response = await graphGetPaginated(client, endpoint, {
      $search: `"${query}"`,
      $select: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead,categories,flag',
      $orderby: 'receivedDateTime desc',
      $top: Math.min(count, 500),
    });
    
    const emails = response.value || [];
    
    const structured = {
      query,
      folder: folder || 'all',
      resultCount: emails.length,
      emails: emails.map(e => {
        const sender = e.from ? e.from.emailAddress : { name: 'Unknown', address: 'unknown' };
        const toRecipients = e.toRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        const ccRecipients = e.ccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        
        return {
          id: e.id,
          subject: e.subject,
          from: { name: sender.name, address: sender.address },
          to: toRecipients,
          cc: ccRecipients,
          receivedDateTime: e.receivedDateTime,
          bodyPreview: e.bodyPreview,
          hasAttachments: e.hasAttachments,
          importance: e.importance,
          isRead: e.isRead,
          categories: e.categories || [],
          flag: e.flag ? {
            status: e.flag.flagStatus,
            dueDateTime: e.flag.dueDateTime?.dateTime,
          } : null,
        };
      }),
    };
    
    const textFallback = [
      `KQL Search Results`,
      `Query: ${query}`,
      `Folder: ${folder || 'all folders'}`,
      `Results: ${emails.length}`,
      '',
      emails.length === 0
        ? '  No emails found matching the query.'
        : emails.map((e, i) => {
            const sender = e.from ? e.from.emailAddress : { name: 'Unknown', address: 'unknown' };
            const readStatus = e.isRead ? '' : '[UNREAD] ';
            const cats = e.categories?.length ? ` [${e.categories.join(', ')}]` : '';
            const flag = e.flag ? ` [${e.flag.flagStatus}]` : '';
            return `${i + 1}. ${readStatus}${e.receivedDateTime} - ${sender.name} <${sender.address}>${cats}${flag}\n   Subject: ${e.subject}\n   Preview: ${e.bodyPreview || '(empty)'}\n   ID: ${e.id}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error executing KQL search: ${error.message}`);
  }
}

/**
 * Get KQL search syntax help
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleKqlHelp(_args) {
  const help = `
KQL (Keyword Query Language) Syntax Reference for Outlook Search:

BASIC TERMS:
  keyword                 Search in subject, body, sender, recipients
  "exact phrase"          Exact phrase match
  word*                   Wildcard (prefix only)

PROPERTY RESTRICTIONS:
  from:"name"             Sender name or email
  to:"name"               To recipients
  cc:"name"               CC recipients
  subject:"text"          Subject line
  body:"text"             Message body
  attachment:"name"       Attachment filename
  hasattachment:true      Messages with attachments
  importance:high         High importance
  importance:low          Low importance
  isread:true             Read messages
  isread:false            Unread messages
  isdraft:true            Draft messages
  category:"name"         Category name
  flagged:true            Flagged messages
  sent:>=2024-01-01       Sent after date
  sent:<=2024-12-31       Sent before date
  received:>=2024-01-01   Received after date
  size:>10000             Size in bytes
  kind:email              Email messages only
  kind:calendar           Calendar items
  kind:contacts           Contacts

LOGICAL OPERATORS:
  AND                     Both terms must match (default)
  OR                      Either term matches
  NOT                     Exclude term
  ( )                     Grouping

EXAMPLES:
  from:john subject:report                    Emails from John with "report" in subject
  subject:"quarterly report" hasattachment:true  Emails with "quarterly report" in subject and attachments
  from:john OR from:jane                      Emails from John or Jane
  received:>=2024-01-01 received:<=2024-01-31  Emails received in January 2024
  importance:high isread:false                Unread high-importance emails
  category:"Project Alpha"                    Emails categorized as "Project Alpha"
  body:budget NOT draft                       Emails with "budget" in body but not "draft"

DATE FORMATS:
  YYYY-MM-DD              2024-01-15
  YYYY-MM-DDTHH:MM:SS     2024-01-15T14:30:00
  relative dates          today, yesterday, this week, last month

NOTE: KQL searches across ALL mailbox content (Inbox, Sent, Archive, etc.)
Use the 'folder' parameter to limit scope to a specific folder.
`;

  return makeResponse(help.trim());
}

module.exports = {
  handleKqlSearch,
  handleKqlHelp,
};