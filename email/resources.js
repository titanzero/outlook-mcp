/**
 * Email resources for MCP
 * Exposes emails and drafts as addressable resources
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { resolveFolderPath } = require('./folder-utils');
const { isAuthError } = require('../utils/response-helpers');
const { ResourceTemplate } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { getEmailFolderCompletions, getEventCompletions } = require('../completion');

/**
 * Register email resources with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerEmailResources(server) {
  // Resource for emails in a folder
  server.resource(
    'outlook-emails',
    'outlook://emails',
    {
      title: 'Emails',
      description: 'List emails in a folder (use query params: folder, count, unreadOnly)',
    },
    async (uri, vars) => {
      try {
        const client = await getGraphClient();

        const folder = vars.folder || 'inbox';
        const count = parseInt(vars.count) || 20;
        const unreadOnly = vars.unreadOnly === 'true';

        const endpoint = await resolveFolderPath(folder);

        let request = client
          .api(endpoint)
          .top(Math.min(count, 100))
          .select(
            'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead'
          )
          .orderby('receivedDateTime desc');

        if (unreadOnly) {
          request = request.filter('isRead eq false');
        }

        const response = await request.get();
        const emails = response.value || [];

        const structured = emails.map(e => {
          const sender = e.from ? e.from.emailAddress : { name: 'Unknown', address: 'unknown' };
          const toRecipients =
            e.toRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
          const ccRecipients =
            e.ccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];

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
          };
        });

        const textFallback =
          emails.length === 0
            ? `No emails found in ${folder}.`
            : `Emails in ${folder} (${emails.length}):\n\n` +
              emails
                .map((e, i) => {
                  const sender = e.from
                    ? e.from.emailAddress
                    : { name: 'Unknown', address: 'unknown' };
                  const readStatus = e.isRead ? '' : '[UNREAD] ';
                  return `${i + 1}. ${readStatus}${e.receivedDateTime} - ${sender.name} <${sender.address}>\n   Subject: ${e.subject}\n   ID: ${e.id}`;
                })
                .join('\n\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify({ emails: structured }, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading emails: ${error.message}`);
      }
    }
  );

  // Resource for individual email with completion
  const emailTemplate = new ResourceTemplate('outlook://emails/{emailId}', {
    complete: {
      emailId: getEventCompletions,
    },
  });

  server.resource(
    'outlook-email',
    emailTemplate,
    {
      title: 'Email',
      description: 'Access a specific email by ID',
    },
    async (uri, { emailId }) => {
      try {
        const client = await getGraphClient();
        const email = await client
          .api(`me/messages/${emailId}`)
          .select(
            'id,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,receivedDateTime,sentDateTime,bodyPreview,body,hasAttachments,importance,isRead,isDraft,conversationId,conversationIndex,internetMessageId,categories,flag'
          )
          .get();

        const sender = email.from
          ? email.from.emailAddress
          : { name: 'Unknown', address: 'unknown' };
        const toRecipients =
          email.toRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        const ccRecipients =
          email.ccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        const bccRecipients =
          email.bccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        const replyTo = email.replyTo?.map(r => r.emailAddress?.address).filter(Boolean) || [];

        const structured = {
          id: email.id,
          subject: email.subject,
          from: { name: sender.name, address: sender.address },
          to: toRecipients,
          cc: ccRecipients,
          bcc: bccRecipients,
          replyTo,
          receivedDateTime: email.receivedDateTime,
          sentDateTime: email.sentDateTime,
          bodyPreview: email.bodyPreview,
          body: email.body?.content,
          bodyContentType: email.body?.contentType,
          hasAttachments: email.hasAttachments,
          importance: email.importance,
          isRead: email.isRead,
          isDraft: email.isDraft,
          conversationId: email.conversationId,
          conversationIndex: email.conversationIndex,
          internetMessageId: email.internetMessageId,
          categories: email.categories || [],
          flag: email.flag
            ? {
              status: email.flag.flagStatus,
              startDateTime: email.flag.startDateTime?.dateTime,
              dueDateTime: email.flag.dueDateTime?.dateTime,
              completedDateTime: email.flag.completedDateTime?.dateTime,
            }
            : null,
        };

        const textFallback = [
          `Email: ${email.subject}`,
          `ID: ${email.id}`,
          `From: ${sender.name} <${sender.address}>`,
          `To: ${toRecipients.join(', ') || 'none'}`,
          ccRecipients.length ? `CC: ${ccRecipients.join(', ')}` : '',
          bccRecipients.length ? `BCC: ${bccRecipients.join(', ')}` : '',
          `Received: ${email.receivedDateTime}`,
          email.sentDateTime ? `Sent: ${email.sentDateTime}` : '',
          `Read: ${email.isRead ? 'Yes' : 'No'}`,
          `Draft: ${email.isDraft ? 'Yes' : 'No'}`,
          `Importance: ${email.importance}`,
          email.hasAttachments ? 'Has Attachments: Yes' : '',
          email.categories?.length ? `Categories: ${email.categories.join(', ')}` : '',
          email.flag ? `Flag: ${email.flag.flagStatus}` : '',
          email.conversationId ? `Conversation: ${email.conversationId}` : '',
          '',
          `Preview: ${email.bodyPreview || '(empty)'}`,
          email.body?.content
            ? `\nBody (${email.body.contentType}):\n${email.body.content.substring(0, 500)}${email.body.content.length > 500 ? '...' : ''}`
            : '',
        ]
          .filter(Boolean)
          .join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading email: ${error.message}`);
      }
    }
  );

  // Resource for drafts
  server.resource(
    'outlook-drafts',
    'outlook://drafts',
    {
      title: 'Drafts',
      description: 'List all email drafts',
    },
    async uri => {
      try {
        const client = await getGraphClient();
        const response = await graphGetPaginated(client, 'me/messages', {
          $filter: 'isDraft eq true',
          $select:
            'id,subject,from,toRecipients,ccRecipients,bccRecipients,bodyPreview,body,hasAttachments,importance,createdDateTime,lastModifiedDateTime',
          $orderby: 'lastModifiedDateTime desc',
        });

        const drafts = response.value || [];

        const structured = drafts.map(d => ({
          id: d.id,
          subject: d.subject,
          to: d.toRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [],
          cc: d.ccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [],
          bcc: d.bccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [],
          bodyPreview: d.bodyPreview,
          hasAttachments: d.hasAttachments,
          importance: d.importance,
          createdDateTime: d.createdDateTime,
          lastModifiedDateTime: d.lastModifiedDateTime,
        }));

        const textFallback =
          drafts.length === 0
            ? 'No drafts found.'
            : `Drafts (${drafts.length}):\n\n` +
              drafts
                .map(
                  (d, i) =>
                    `${i + 1}. ${d.subject || '(no subject)'}\n   To: ${d.toRecipients?.map(r => r.emailAddress?.address).join(', ') || 'none'}\n   Modified: ${d.lastModifiedDateTime}\n   ID: ${d.id}`
                )
                .join('\n\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify({ drafts: structured }, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading drafts: ${error.message}`);
      }
    }
  );

  // Resource for individual draft
  server.resource(
    'outlook-draft',
    'outlook://drafts/{draftId}',
    {
      title: 'Draft',
      description: 'Access a specific draft by ID',
    },
    async (uri, { draftId }) => {
      try {
        const client = await getGraphClient();
        const draft = await client
          .api(`me/messages/${draftId}`)
          .select(
            'id,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,bodyPreview,body,hasAttachments,importance,createdDateTime,lastModifiedDateTime,categories'
          )
          .get();

        const structured = {
          id: draft.id,
          subject: draft.subject,
          to: draft.toRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [],
          cc: draft.ccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [],
          bcc: draft.bccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [],
          bodyPreview: draft.bodyPreview,
          body: draft.body?.content,
          bodyContentType: draft.body?.contentType,
          hasAttachments: draft.hasAttachments,
          importance: draft.importance,
          createdDateTime: draft.createdDateTime,
          lastModifiedDateTime: draft.lastModifiedDateTime,
          categories: draft.categories || [],
        };

        const textFallback = [
          `Draft: ${draft.subject || '(no subject)'}`,
          `ID: ${draft.id}`,
          `To: ${structured.to.join(', ') || 'none'}`,
          structured.cc.length ? `CC: ${structured.cc.join(', ')}` : '',
          structured.bcc.length ? `BCC: ${structured.bcc.join(', ')}` : '',
          `Created: ${draft.createdDateTime}`,
          `Modified: ${draft.lastModifiedDateTime}`,
          `Importance: ${draft.importance}`,
          draft.hasAttachments ? 'Has Attachments: Yes' : '',
          draft.categories?.length ? `Categories: ${draft.categories.join(', ')}` : '',
          '',
          `Preview: ${draft.bodyPreview || '(empty)'}`,
          draft.body?.content
            ? `\nBody (${draft.body.contentType}):\n${draft.body.content.substring(0, 500)}${draft.body.content.length > 500 ? '...' : ''}`
            : '',
        ]
          .filter(Boolean)
          .join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading draft: ${error.message}`);
      }
    }
  );
}

module.exports = { registerEmailResources };
