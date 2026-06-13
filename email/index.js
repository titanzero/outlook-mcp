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

const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: { type: "string", description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')" },
        count: { type: "number", description: "Number of emails to retrieve (default: 10, max: 500). Pagination is handled automatically." }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria. Searches across ALL folders (inbox, archive, sent, etc.) by default. Specify folder to limit scope.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search query text to find in emails" },
        folder: { type: "string", description: "Email folder to search in (default: all folders). Use 'inbox', 'archive', 'sent', etc. to limit scope." },
        from: { type: "string", description: "Filter by sender email address or name" },
        to: { type: "string", description: "Filter by recipient email address or name" },
        subject: { type: "string", description: "Filter by email subject" },
        hasAttachments: { type: "boolean", description: "Filter to only emails with attachments" },
        unreadOnly: { type: "boolean", description: "Filter to only unread emails" },
        count: { type: "number", description: "Number of results to return (default: 10, max: 500). Pagination is handled automatically." }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email. Returns a short preview (255 chars) by default. Set fullBody=true to fetch the complete email body.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the email to read" },
        fullBody: { type: "boolean", description: "If true, fetches the complete email body instead of the 255-char preview. Use for emails where the preview is insufficient (default: false)." }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email",
    inputSchema: {
      type: "object",
      properties: {
        to: { type: "string", description: "Comma-separated list of recipient email addresses" },
        cc: { type: "string", description: "Comma-separated list of CC recipient email addresses" },
        bcc: { type: "string", description: "Comma-separated list of BCC recipient email addresses" },
        subject: { type: "string", description: "Email subject" },
        body: { type: "string", description: "Email body content (can be plain text or HTML)" },
        importance: { type: "string", description: "Email importance (normal, high, low)", enum: ["normal", "high", "low"] },
        saveToSentItems: { type: "boolean", description: "Whether to save the email to sent items" }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "reply-email",
    description: "Replies to an email. Set replyAll=true to reply to all recipients.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the email to reply to" },
        comment: { type: "string", description: "Reply message body" },
        replyAll: { type: "boolean", description: "If true, replies to all recipients (default: false)" }
      },
      required: ["id"]
    },
    handler: handleReplyEmail
  },
  {
    name: "forward-email",
    description: "Forwards an email to one or more recipients",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the email to forward" },
        to: { type: "string", description: "Comma-separated list of recipient email addresses" },
        comment: { type: "string", description: "Optional message to include with the forward" }
      },
      required: ["id", "to"]
    },
    handler: handleForwardEmail
  },
  {
    name: "delete-email",
    description: "Permanently deletes an email",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the email to delete" }
      },
      required: ["id"]
    },
    handler: handleDeleteEmail
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the email to mark as read/unread" },
        isRead: { type: "boolean", description: "Whether to mark as read (true) or unread (false). Default: true" }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "list-attachments",
    description: "Lists attachments of a specific email",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the email" }
      },
      required: ["id"]
    },
    handler: handleListAttachments
  },
  {
    name: "download-attachment",
    description: "Downloads an attachment from an email. Returns file metadata and base64-encoded content.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the email" },
        attachmentId: { type: "string", description: "ID of the attachment (from list-attachments)" }
      },
      required: ["id", "attachmentId"]
    },
    handler: handleDownloadAttachment
  },
  {
    name: "create-draft",
    description: "Creates a new email draft without sending it",
    inputSchema: {
      type: "object",
      properties: {
        to: { type: "string", description: "Comma-separated list of recipient email addresses" },
        cc: { type: "string", description: "Comma-separated list of CC recipient email addresses" },
        bcc: { type: "string", description: "Comma-separated list of BCC recipient email addresses" },
        subject: { type: "string", description: "Email subject" },
        body: { type: "string", description: "Email body content (plain text or HTML)" },
        importance: { type: "string", description: "Email importance (normal, high, low)", enum: ["normal", "high", "low"] }
      },
      required: []
    },
    handler: handleCreateDraft
  },
  {
    name: "update-draft",
    description: "Updates an existing email draft",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the draft to update" },
        to: { type: "string", description: "Comma-separated list of recipient email addresses" },
        cc: { type: "string", description: "Comma-separated list of CC recipient email addresses" },
        bcc: { type: "string", description: "Comma-separated list of BCC recipient email addresses" },
        subject: { type: "string", description: "Email subject" },
        body: { type: "string", description: "Email body content (plain text or HTML)" },
        importance: { type: "string", description: "Email importance (normal, high, low)", enum: ["normal", "high", "low"] }
      },
      required: ["id"]
    },
    handler: handleUpdateDraft
  },
  {
    name: "send-draft",
    description: "Sends an existing email draft",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the draft to send" }
      },
      required: ["id"]
    },
    handler: handleSendDraft
  }
];

module.exports = {
  emailTools,
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
