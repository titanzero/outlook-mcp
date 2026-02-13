/**
 * Read email functionality
 * Uses the Microsoft Graph JS SDK.
 */
const config = require('../config');
const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Convert HTML to plain text with improved handling.
 * Strips tags, decodes common HTML entities, and normalizes whitespace.
 * @param {string} html - HTML content
 * @returns {string} - Plain text
 */
function htmlToText(html) {
  return html
    // Remove style and script blocks entirely
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    // Add newlines for block-level elements
    .replace(/<\/(p|div|h[1-6]|li|tr|br\s*\/?)>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    // Strip remaining tags
    .replace(/<[^>]*>/g, '')
    // Decode common HTML entities
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&#x27;/gi, "'")
    // Normalize whitespace: collapse multiple spaces/tabs within lines
    .replace(/[ \t]+/g, ' ')
    // Collapse 3+ consecutive newlines into 2
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Read email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleReadEmail(args) {
  const emailId = args.id;
  const fullBody = args.fullBody === true;
  
  if (!emailId) {
    return makeErrorResponse('Email ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Use full body fields only when explicitly requested
    const selectFields = fullBody ? config.EMAIL_FULL_BODY_FIELDS : config.EMAIL_DETAIL_FIELDS;
    
    try {
      const email = await client.api(`me/messages/${emailId}`)
        .select(selectFields)
        .get();
      
      if (!email) {
        return makeErrorResponse(`Email with ID ${emailId} not found.`);
      }
      
      // Format sender, recipients, etc.
      const sender = email.from ? `${email.from.emailAddress.name} (${email.from.emailAddress.address})` : 'Unknown';
      const to = email.toRecipients ? email.toRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const cc = email.ccRecipients && email.ccRecipients.length > 0 ? email.ccRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const bcc = email.bccRecipients && email.bccRecipients.length > 0 ? email.bccRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const date = new Date(email.receivedDateTime).toLocaleString();
      
      // Extract body content
      let body = '';
      if (fullBody && email.body) {
        // Full body requested: convert HTML to text or use plain text
        body = email.body.contentType === 'html'
          ? htmlToText(email.body.content)
          : email.body.content;
      } else {
        // Default: use bodyPreview (max 255 chars, clean plain text from Graph API)
        body = email.bodyPreview || 'No content';
      }

      // Build structured object for TOON (key-value)
      const structured = {
        from: sender,
        to,
        subject: email.subject,
        date,
        importance: email.importance || 'normal',
        hasAttachments: !!email.hasAttachments,
        body,
      };
      if (cc !== 'None') structured.cc = cc;
      if (bcc !== 'None') structured.bcc = bcc;
      if (!fullBody) structured.note = 'Preview — use fullBody=true for complete content';

      // Build plain-text fallback
      const formattedEmail = `From: ${sender}
To: ${to}
${cc !== 'None' ? `CC: ${cc}\n` : ''}${bcc !== 'None' ? `BCC: ${bcc}\n` : ''}Subject: ${email.subject}
Date: ${date}
Importance: ${email.importance || 'normal'}
Has Attachments: ${email.hasAttachments ? 'Yes' : 'No'}
${!fullBody ? '(Preview — use fullBody=true for complete content)\n' : ''}
${body}`;

      const text = formatResponse(structured, formattedEmail);
      
      return makeResponse(text);
    } catch (error) {
      console.error(`Error reading email: ${error.message}`);
      
      // Improved error handling with more specific messages
      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return makeErrorResponse('The email ID seems invalid or doesn\'t belong to your mailbox. Please try with a different email ID.');
      }
      return makeErrorResponse(`Failed to read email: ${error.message}`);
    }
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error accessing email: ${error.message}`);
  }
}

module.exports = handleReadEmail;
