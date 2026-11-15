/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const { makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { z } = require('zod');

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
        'Authentication configuration is missing. Set OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET, then restart and try again.'
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

/**
 * Register authentication tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerAuthTools(server) {
  // About tool - read-only, no side effects
  server.tool(
    "about",
    "Returns information about this Outlook Assistant server",
    {},
    { 
      annotations: { 
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleAbout
  );

  // Authenticate tool - may trigger OAuth flow
  server.tool(
    "authenticate",
    "Authenticate with Microsoft Graph API to access Outlook data",
    {
      force: z.boolean().optional().describe("Force re-authentication even if already authenticated")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true
      }
    },
    handleAuthenticate
  );

  // Check auth status - read-only
  server.tool(
    "check-auth-status",
    "Check the current authentication status with Microsoft Graph API",
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleCheckAuthStatus
  );
}

module.exports = {
  registerAuthTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};