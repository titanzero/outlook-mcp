/**
 * Authentication module for Outlook MCP server
 */
const tokenManager = require('./token-manager');
const { authTools } = require('./tools');

/**
 * Build a rich authentication error that callers (tools) can inspect
 * to provide more helpful feedback to the user.
 * @returns {Error}
 */
function buildAuthError() {
  const reason = tokenManager.getLastErrorReason();
  let message = 'Authentication required. Please use the \'authenticate\' tool first.';

  if (reason && reason.code) {
    switch (reason.code) {
      case 'TOKEN_FILE_MISSING':
        message = `Authentication required: token file not found at ${reason.path}. ` +
          `Start the auth server with \`npm run auth-server\` and use the 'authenticate' tool to sign in.`;
        break;
      case 'TOKEN_FILE_INVALID_JSON':
        message = `Authentication failed: your token file at ${reason.path} contains invalid JSON. ` +
          `Delete the file and run the 'authenticate' tool again to regenerate it.`;
        break;
      case 'TOKEN_FILE_INVALID_SHAPE':
        message = `Authentication failed: your token file at ${reason.path} is missing required fields (like access_token). ` +
          `Delete the file and re-run the 'authenticate' tool.`;
        break;
      case 'CLIENT_CONFIG_MISSING':
        message = 'Authentication configuration is missing: OUTLOOK_CLIENT_ID and/or OUTLOOK_CLIENT_SECRET are not set. ' +
          'Update your Claude Desktop config to include these env vars (using the Azure client ID and secret VALUE), restart Claude, then authenticate again.';
        break;
      case 'REFRESH_FAILED_INVALID_CLIENT':
      case 'CODE_EXCHANGE_INVALID_CLIENT':
        message = 'Authentication failed: Microsoft rejected the client credentials (invalid_client). ' +
          'Double-check that OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET are correctly configured (using the Azure client secret VALUE, not the Secret ID), then re-authenticate.';
        break;
      case 'REFRESH_NETWORK_ERROR':
      case 'CODE_EXCHANGE_NETWORK_ERROR':
        message = `Authentication failed due to a network error talking to Microsoft: ${reason.message}. ` +
          'Please check your internet connection and try the \'authenticate\' tool again.';
        break;
      default:
        if (reason.message) {
          message = `Authentication failed: ${reason.message}`;
        }
        break;
    }
  }

  const error = new Error(message);
  error.isAuthError = true;
  error.authReason = (reason && reason.code) || 'AUTH_REQUIRED';
  if (reason) {
    error.authDetails = reason;
  }
  return error;
}

/**
 * Ensures the user is authenticated and returns an access token.
 * Automatically refreshes expired tokens using the refresh token.
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    throw buildAuthError();
  }

  const accessToken = await tokenManager.getAccessToken();
  if (!accessToken) {
    throw buildAuthError();
  }

  return accessToken;
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated
};
