/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const { makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return makeResponse(
    `📧 MODULAR Outlook Assistant MCP Server v${config.SERVER_VERSION} 📧\n\nProvides access to Microsoft Outlook email, calendar, and contacts through Microsoft Graph API.\nPowered by the official Microsoft Graph JavaScript SDK.`
  );
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;

  try {
    if (force) {
      await tokenManager.clearTokens();
    }

    if (!config.AUTH_CONFIG.clientId) {
      return makeErrorResponse(
        'Authentication configuration is missing. Set OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET, then restart Claude and try again.'
      );
    }

    const authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${encodeURIComponent(config.AUTH_CONFIG.clientId)}`;
    const prefix = force ? 'Existing tokens were cleared. ' : '';
    return makeResponse(
      `${prefix}Authentication required. Please visit the following URL to authenticate with Microsoft: ${authUrl}\n\nAfter authentication, you will be redirected back to this application.`
    );
  } catch (error) {
    return makeErrorResponse(`Error starting authentication flow: ${error.message}`);
  }
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');
  
  const tokens = tokenManager.loadTokenCacheSync();
  
  console.error(`[CHECK-AUTH-STATUS] Tokens loaded: ${tokens ? 'YES' : 'NO'}`);
  
  if (!tokens || !tokens.access_token) {
    console.error('[CHECK-AUTH-STATUS] No valid access token found');
    return makeResponse('Not authenticated');
  }
  
  console.error('[CHECK-AUTH-STATUS] Access token present');
  console.error(`[CHECK-AUTH-STATUS] Token expires at: ${tokens.expires_at}`);
  console.error(`[CHECK-AUTH-STATUS] Current time: ${Date.now()}`);
  
  return makeResponse('Authenticated and ready');
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this Outlook Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
