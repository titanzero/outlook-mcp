const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleMarkAsRead = require('./mark-as-read');
const handleReplyEmail = require('./reply');
const handleForwardEmail = require('./forward');
const handleDeleteEmail = require('./delete');
const { handleListAttachments, handleDownloadAttachment } = require('./attachments');
const { handleCreateDraft, handleUpdateDraft, handleSendDraft } = require('./draft');
const handleFlagEmail = require('./flag');
const handleCategorizeEmail = require('./categorize');
const handleUploadAttachment = require('./upload-attachment');
const handleDeleteAttachment = require('./delete-attachment');
const { registerEmailResources } = require('./resources');
const { handleListSharedMailboxes, handleAccessSharedMailbox, handleListSharedMailboxFolders } = require('./shared-mailbox');
const { handleListGroups, handleGetGroup, handleListGroupMembers, handleSendEmailToGroup } = require('./groups');
const { handleListConversations, handleGetConversation } = require('./conversations');
const { handleRecallMessage, handleGetRecallStatus } = require('./recall');
const { handleKqlSearch, handleKqlHelp } = require('./kql-search');
const { z } = require('zod');

/**
 * Register email tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerEmailTools(server) {
  // Read-only tools
  server.tool(
    'list-emails',
    'Lists recent emails from your inbox',
    {
      folder: z
        .string()
        .optional()
        .describe("Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"),
      count: z
        .number()
        .int()
        .positive()
        .max(500)
        .optional()
        .describe(
          'Number of emails to retrieve (default: 10, max: 500). Pagination is handled automatically.'
        ),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListEmails
  );

  server.tool(
    'search-emails',
    'Search for emails using various criteria. Searches across ALL folders (inbox, archive, sent, etc.) by default. Specify folder to limit scope.',
    {
      query: z.string().optional().describe('Search query text to find in emails'),
      folder: z
        .string()
        .optional()
        .describe(
          "Email folder to search in (default: all folders). Use 'inbox', 'archive', 'sent', etc. to limit scope."
        ),
      from: z.string().optional().describe('Filter by sender email address or name'),
      to: z.string().optional().describe('Filter by recipient email address or name'),
      subject: z.string().optional().describe('Filter by email subject'),
      hasAttachments: z.boolean().optional().describe('Filter to only emails with attachments'),
      unreadOnly: z.boolean().optional().describe('Filter to only unread emails'),
      count: z
        .number()
        .int()
        .positive()
        .max(500)
        .optional()
        .describe(
          'Number of results to return (default: 10, max: 500). Pagination is handled automatically.'
        ),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleSearchEmails
  );

  server.tool(
    'read-email',
    'Reads the content of a specific email. Returns a short preview (255 chars) by default. Set fullBody=true to fetch the complete email body.',
    {
      id: z.string().describe('ID of the email to read'),
      fullBody: z
        .boolean()
        .optional()
        .describe(
          'If true, fetches the complete email body instead of the 255-char preview. Use for emails where the preview is insufficient (default: false).'
        ),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleReadEmail
  );

  server.tool(
    'list-attachments',
    'Lists attachments of a specific email',
    {
      id: z.string().describe('ID of the email'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListAttachments
  );

  server.tool(
    'download-attachment',
    'Downloads an attachment from an email. Returns file metadata and base64-encoded content.',
    {
      id: z.string().describe('ID of the email'),
      attachmentId: z.string().describe('ID of the attachment (from list-attachments)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleDownloadAttachment
  );

  // Attachment management tools (for drafts)
  server.tool(
    'upload-attachment',
    'Uploads an attachment to an email or draft. Provide base64-encoded content.',
    {
      id: z.string().describe('ID of the email or draft'),
      name: z.string().describe('Attachment file name'),
      contentBytes: z.string().describe('Base64-encoded file content'),
      contentType: z
        .string()
        .optional()
        .describe("MIME type (e.g., 'application/pdf', 'image/png'). Auto-detected if omitted."),
      isInline: z
        .boolean()
        .optional()
        .describe('Whether the attachment is inline (default: false)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleUploadAttachment
  );

  server.tool(
    'delete-attachment',
    'Deletes an attachment from an email or draft',
    {
      id: z.string().describe('ID of the email or draft'),
      attachmentId: z.string().describe('ID of the attachment (from list-attachments)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteAttachment
  );

  // Write tools - mutating operations
  server.tool(
    'send-email',
    'Composes and sends a new email',
    {
      to: z.string().describe('Comma-separated list of recipient email addresses'),
      cc: z.string().optional().describe('Comma-separated list of CC recipient email addresses'),
      bcc: z.string().optional().describe('Comma-separated list of BCC recipient email addresses'),
      subject: z.string().describe('Email subject'),
      body: z.string().describe('Email body content (can be plain text or HTML)'),
      importance: z
        .enum(['normal', 'high', 'low'])
        .optional()
        .describe('Email importance (normal, high, low)'),
      saveToSentItems: z.boolean().optional().describe('Whether to save the email to sent items'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleSendEmail
  );

  server.tool(
    'reply-email',
    'Replies to an email. Set replyAll=true to reply to all recipients.',
    {
      id: z.string().describe('ID of the email to reply to'),
      comment: z.string().describe('Reply message body'),
      replyAll: z
        .boolean()
        .optional()
        .describe('If true, replies to all recipients (default: false)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleReplyEmail
  );

  server.tool(
    'forward-email',
    'Forwards an email to one or more recipients',
    {
      id: z.string().describe('ID of the email to forward'),
      to: z.string().describe('Comma-separated list of recipient email addresses'),
      comment: z.string().optional().describe('Optional message to include with the forward'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleForwardEmail
  );

  server.tool(
    'delete-email',
    'Permanently deletes an email',
    {
      id: z.string().describe('ID of the email to delete'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteEmail
  );

  server.tool(
    'flag-email',
    'Flags, completes, or clears the flag on an email for follow-up',
    {
      id: z.string().describe('ID of the email to flag'),
      flagStatus: z
        .enum(['flagged', 'completed', 'cleared'])
        .describe(
          "Flag status: 'flagged' (for follow-up), 'completed' (done), 'cleared' (remove flag)"
        ),
      dueDateTime: z
        .string()
        .optional()
        .describe(
          "Due date/time for follow-up in ISO 8601 format (UTC). Only used when flagStatus is 'flagged'."
        ),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleFlagEmail
  );

  server.tool(
    'categorize-email',
    'Adds or removes categories from an email',
    {
      id: z.string().describe('ID of the email to categorize'),
      categories: z
        .string()
        .optional()
        .describe('Comma-separated list of category names. Empty string clears all categories.'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleCategorizeEmail
  );

  server.tool(
    'mark-as-read',
    'Marks an email as read or unread',
    {
      id: z.string().describe('ID of the email to mark as read/unread'),
      isRead: z
        .boolean()
        .optional()
        .describe('Whether to mark as read (true) or unread (false). Default: true'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleMarkAsRead
  );

  // Draft tools
  server.tool(
    'create-draft',
    'Creates a new email draft without sending it',
    {
      to: z.string().optional().describe('Comma-separated list of recipient email addresses'),
      cc: z.string().optional().describe('Comma-separated list of CC recipient email addresses'),
      bcc: z.string().optional().describe('Comma-separated list of BCC recipient email addresses'),
      subject: z.string().optional().describe('Email subject'),
      body: z.string().optional().describe('Email body content (plain text or HTML)'),
      importance: z
        .enum(['normal', 'high', 'low'])
        .optional()
        .describe('Email importance (normal, high, low)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateDraft
  );

  server.tool(
    'update-draft',
    'Updates an existing email draft',
    {
      id: z.string().describe('ID of the draft to update'),
      to: z.string().optional().describe('Comma-separated list of recipient email addresses'),
      cc: z.string().optional().describe('Comma-separated list of CC recipient email addresses'),
      bcc: z.string().optional().describe('Comma-separated list of BCC recipient email addresses'),
      subject: z.string().optional().describe('Email subject'),
      body: z.string().optional().describe('Email body content (plain text or HTML)'),
      importance: z
        .enum(['normal', 'high', 'low'])
        .optional()
        .describe('Email importance (normal, high, low)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateDraft
  );

  server.tool(
    'send-draft',
    'Sends an existing email draft',
    {
      id: z.string().describe('ID of the draft to send'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleSendDraft
  );

  // Shared Mailbox tools
  server.tool(
    'list-shared-mailboxes',
    'Lists mailboxes you have delegate/shared access to (requires Mail.Read.Shared scope)',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListSharedMailboxes
  );

  server.tool(
    'access-shared-mailbox',
    'Access emails from a shared/delegate mailbox by email address (requires Mail.Read.Shared scope)',
    {
      emailAddress: z.string().describe('Email address of the shared mailbox'),
      folder: z.string().optional().describe("Folder to access (default: 'inbox')"),
      count: z.number().int().positive().max(100).optional().describe('Number of emails to retrieve (default: 20, max: 100)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleAccessSharedMailbox
  );

  server.tool(
    'list-shared-mailbox-folders',
    'Lists folders in a shared/delegate mailbox (requires Mail.Read.Shared scope)',
    {
      emailAddress: z.string().describe('Email address of the shared mailbox'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListSharedMailboxFolders
  );

  // Microsoft 365 Groups / Distribution Lists tools
  server.tool(
    'list-groups',
    'Lists all Microsoft 365 Groups and Distribution Lists you are a member of (requires Group.Read.All scope)',
    {
      count: z.number().int().positive().max(100).optional().describe('Number of groups to retrieve (default: 50, max: 100)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListGroups
  );

  server.tool(
    'get-group',
    'Gets detailed information about a group including members and owners (requires Group.Read.All scope)',
    {
      groupId: z.string().describe('ID of the group (from list-groups)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetGroup
  );

  server.tool(
    'list-group-members',
    'Lists members of a Microsoft 365 Group or Distribution List (requires Group.Read.All scope)',
    {
      groupId: z.string().describe('ID of the group (from list-groups)'),
      count: z.number().int().positive().max(999).optional().describe('Number of members to retrieve (default: 100, max: 999)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListGroupMembers
  );

  server.tool(
    'send-email-to-group',
    'Sends an email to a Microsoft 365 Group or Distribution List (requires Group.Read.All scope)',
    {
      groupId: z.string().describe('ID of the group (from list-groups)'),
      subject: z.string().describe('Email subject'),
      body: z.string().describe('Email body content (HTML)'),
      importance: z.enum(['normal', 'high', 'low']).optional().describe('Email importance (normal, high, low)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleSendEmailToGroup
  );

  // Conversation/Thread view tools
  server.tool(
    'list-conversations',
    'Lists email conversations (threaded view) grouped by conversation ID',
    {
      folder: z.string().optional().describe("Folder to list conversations from (default: 'inbox')"),
      count: z.number().int().positive().max(100).optional().describe('Number of conversations to retrieve (default: 20, max: 100)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListConversations
  );

  server.tool(
    'get-conversation',
    'Gets the full conversation thread by conversation ID',
    {
      conversationId: z.string().describe('Conversation ID (from list-conversations)'),
      count: z.number().int().positive().max(200).optional().describe('Number of messages in thread to retrieve (default: 50, max: 200)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetConversation
  );

  // Message Recall tools
  server.tool(
    'recall-message',
    'Attempts to recall a sent message (only works within same organization, unread messages)',
    {
      messageId: z.string().describe('ID of the sent message to recall (from sent items)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleRecallMessage
  );

  server.tool(
    'get-recall-status',
    'Checks the recall status for a message by looking for recall notifications',
    {
      messageId: z.string().describe('ID of the recalled message'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetRecallStatus
  );

  // KQL Search tools
  server.tool(
    'search-emails-kql',
    'Searches emails using Microsoft Graph KQL (Keyword Query Language) syntax for advanced queries',
    {
      query: z.string().describe('KQL query string (e.g., from:john subject:report hasattachment:true)'),
      folder: z.string().optional().describe("Optional folder to limit search scope (default: all folders)"),
      count: z.number().int().positive().max(500).optional().describe('Number of results to return (default: 50, max: 500)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleKqlSearch
  );

  server.tool(
    'kql-help',
    'Returns KQL (Keyword Query Language) syntax reference and examples for email search',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleKqlHelp
  );
}

module.exports = {
  registerEmailTools,
  registerEmailResources,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleMarkAsRead,
  handleReplyEmail,
  handleForwardEmail,
  handleDeleteEmail,
  handleListAttachments,
  handleDownloadAttachment,
  handleCreateDraft,
  handleUpdateDraft,
  handleSendDraft,
  handleFlagEmail,
  handleCategorizeEmail,
  handleUploadAttachment,
  handleDeleteAttachment,
  handleListSharedMailboxes,
  handleAccessSharedMailbox,
  handleListSharedMailboxFolders,
  handleListGroups,
  handleGetGroup,
  handleListGroupMembers,
  handleSendEmailToGroup,
  handleListConversations,
  handleGetConversation,
  handleRecallMessage,
  handleGetRecallStatus,
  handleKqlSearch,
  handleKqlHelp,
};
