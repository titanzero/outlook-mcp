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
const { z } = require('zod');

/**
 * Register email tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerEmailTools(server) {
  // Read-only tools
  server.tool(
    "list-emails",
    "Lists recent emails from your inbox",
    {
      folder: z.string().optional().describe("Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"),
      count: z.number().int().positive().max(500).optional().describe("Number of emails to retrieve (default: 10, max: 500). Pagination is handled automatically.")
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleListEmails
  );

  server.tool(
    "search-emails",
    "Search for emails using various criteria. Searches across ALL folders (inbox, archive, sent, etc.) by default. Specify folder to limit scope.",
    {
      query: z.string().optional().describe("Search query text to find in emails"),
      folder: z.string().optional().describe("Email folder to search in (default: all folders). Use 'inbox', 'archive', 'sent', etc. to limit scope."),
      from: z.string().optional().describe("Filter by sender email address or name"),
      to: z.string().optional().describe("Filter by recipient email address or name"),
      subject: z.string().optional().describe("Filter by email subject"),
      hasAttachments: z.boolean().optional().describe("Filter to only emails with attachments"),
      unreadOnly: z.boolean().optional().describe("Filter to only unread emails"),
      count: z.number().int().positive().max(500).optional().describe("Number of results to return (default: 10, max: 500). Pagination is handled automatically.")
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleSearchEmails
  );

  server.tool(
    "read-email",
    "Reads the content of a specific email. Returns a short preview (255 chars) by default. Set fullBody=true to fetch the complete email body.",
    {
      id: z.string().describe("ID of the email to read"),
      fullBody: z.boolean().optional().describe("If true, fetches the complete email body instead of the 255-char preview. Use for emails where the preview is insufficient (default: false).")
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleReadEmail
  );

  server.tool(
    "list-attachments",
    "Lists attachments of a specific email",
    {
      id: z.string().describe("ID of the email")
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleListAttachments
  );

  server.tool(
    "download-attachment",
    "Downloads an attachment from an email. Returns file metadata and base64-encoded content.",
    {
      id: z.string().describe("ID of the email"),
      attachmentId: z.string().describe("ID of the attachment (from list-attachments)")
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleDownloadAttachment
  );

  // Write tools - mutating operations
  server.tool(
    "send-email",
    "Composes and sends a new email",
    {
      to: z.string().describe("Comma-separated list of recipient email addresses"),
      cc: z.string().optional().describe("Comma-separated list of CC recipient email addresses"),
      bcc: z.string().optional().describe("Comma-separated list of BCC recipient email addresses"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content (can be plain text or HTML)"),
      importance: z.enum(["normal", "high", "low"]).optional().describe("Email importance (normal, high, low)"),
      saveToSentItems: z.boolean().optional().describe("Whether to save the email to sent items")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true
      }
    },
    handleSendEmail
  );

  server.tool(
    "reply-email",
    "Replies to an email. Set replyAll=true to reply to all recipients.",
    {
      id: z.string().describe("ID of the email to reply to"),
      comment: z.string().describe("Reply message body"),
      replyAll: z.boolean().optional().describe("If true, replies to all recipients (default: false)")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true
      }
    },
    handleReplyEmail
  );

  server.tool(
    "forward-email",
    "Forwards an email to one or more recipients",
    {
      id: z.string().describe("ID of the email to forward"),
      to: z.string().describe("Comma-separated list of recipient email addresses"),
      comment: z.string().optional().describe("Optional message to include with the forward")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true
      }
    },
    handleForwardEmail
  );

  server.tool(
    "delete-email",
    "Permanently deletes an email",
    {
      id: z.string().describe("ID of the email to delete")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false
      }
    },
    handleDeleteEmail
  );

  server.tool(
    "mark-as-read",
    "Marks an email as read or unread",
    {
      id: z.string().describe("ID of the email to mark as read/unread"),
      isRead: z.boolean().optional().describe("Whether to mark as read (true) or unread (false). Default: true")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false
      }
    },
    handleMarkAsRead
  );

  // Draft tools
  server.tool(
    "create-draft",
    "Creates a new email draft without sending it",
    {
      to: z.string().optional().describe("Comma-separated list of recipient email addresses"),
      cc: z.string().optional().describe("Comma-separated list of CC recipient email addresses"),
      bcc: z.string().optional().describe("Comma-separated list of BCC recipient email addresses"),
      subject: z.string().optional().describe("Email subject"),
      body: z.string().optional().describe("Email body content (plain text or HTML)"),
      importance: z.enum(["normal", "high", "low"]).optional().describe("Email importance (normal, high, low)")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false
      }
    },
    handleCreateDraft
  );

  server.tool(
    "update-draft",
    "Updates an existing email draft",
    {
      id: z.string().describe("ID of the draft to update"),
      to: z.string().optional().describe("Comma-separated list of recipient email addresses"),
      cc: z.string().optional().describe("Comma-separated list of CC recipient email addresses"),
      bcc: z.string().optional().describe("Comma-separated list of BCC recipient email addresses"),
      subject: z.string().optional().describe("Email subject"),
      body: z.string().optional().describe("Email body content (plain text or HTML)"),
      importance: z.enum(["normal", "high", "low"]).optional().describe("Email importance (normal, high, low)")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false
      }
    },
    handleUpdateDraft
  );

  server.tool(
    "send-draft",
    "Sends an existing email draft",
    {
      id: z.string().describe("ID of the draft to send")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true
      }
    },
    handleSendDraft
  );
}

module.exports = {
  registerEmailTools,
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
};