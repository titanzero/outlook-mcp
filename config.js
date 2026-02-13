/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

/**
 * Resolve the shared path where OAuth tokens are stored.
 * This function is used both by the MCP server and the standalone
 * auth server so they always read/write the same JSON file.
 */
function getTokenStorePath() {
  return path.join(homeDir, '.outlook-mcp-tokens.json');
}

const AUTH_ENDPOINT = process.env.OUTLOOK_AUTH_ENDPOINT || 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
const TOKEN_ENDPOINT = process.env.OUTLOOK_TOKEN_ENDPOINT || 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const DEFAULT_REDIRECT_URI = process.env.OUTLOOK_REDIRECT_URI || 'http://localhost:3333/auth/callback';
const DEFAULT_AUTH_SERVER_URL = process.env.OUTLOOK_AUTH_SERVER_URL || 'http://localhost:3333';
const TOKEN_REFRESH_BUFFER_MS = 20 * 60 * 1000;

const DEFAULT_SCOPES = [
  'offline_access', 'email', 'openid',
  'User.Read',
  'Mail.Read', 'Mail.ReadWrite', 'Mail.Send',
  'Calendars.Read', 'Calendars.ReadWrite',
  'Contacts.Read',
  'MailboxSettings.ReadWrite', 'MailboxSettings.Read',
  'MailboxFolder.ReadWrite', 'MailboxFolder.Read',
];
const parsedScopes = (process.env.OUTLOOK_SCOPES || '').split(/\s+/).filter(Boolean);

const FOLDER_CACHE_TTL_MS = 5 * 60 * 1000;
const FOLDER_PAGE_SIZE = 100;
const FOLDER_PATH_MAX_DEPTH = 10;
const FOLDER_TRAVERSAL_MAX_DEPTH = 5;
const DEFAULT_RULE_SEQUENCE = 100;

module.exports = {
  // Server information
  SERVER_NAME: "outlook-assistant",
  SERVER_VERSION: "1.0.0",
  
  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.OUTLOOK_CLIENT_ID || '',
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET || '',
    redirectUri: DEFAULT_REDIRECT_URI,
    scopes: parsedScopes.length > 0 ? parsedScopes : DEFAULT_SCOPES,
    tokenStorePath: getTokenStorePath(),
    authServerUrl: DEFAULT_AUTH_SERVER_URL
  },
  AUTH_ENDPOINT,
  TOKEN_ENDPOINT,
  TOKEN_REFRESH_BUFFER_MS,
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled,recurrence',

  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,receivedDateTime,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance',
  EMAIL_FULL_BODY_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance',
  FOLDER_SELECT_FIELDS: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 50,
  MAX_RESULT_COUNT: 1000,

  // Folder traversal configuration
  FOLDER_CACHE_TTL_MS,
  FOLDER_PAGE_SIZE,
  FOLDER_PATH_MAX_DEPTH,
  FOLDER_TRAVERSAL_MAX_DEPTH,

  // Rules configuration
  DEFAULT_RULE_SEQUENCE,

  // Response format: 'toon' for token-optimized, 'text' for human-readable plain text
  RESPONSE_FORMAT: process.env.OUTLOOK_RESPONSE_FORMAT || 'toon',

  // Timezone
  DEFAULT_TIMEZONE: "Central European Standard Time",
  
  // Shared helpers
  getTokenStorePath,
};
