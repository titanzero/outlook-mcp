const querystring = require('querystring');
const crypto = require('crypto');
const tokenManager = require('./token-manager');
const config = require('../config');

// HTML templates
function escapeHtml(unsafe) {
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

const templates = {
  authError: (error, errorDescription) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">❌ Authorization Failed</h1>
        <p><strong>Error:</strong> ${escapeHtml(error)}</p>
        ${errorDescription ? `<p><strong>Description:</strong> ${escapeHtml(errorDescription)}</p>` : ''}
        <p>You can close this window and try again.</p>
      </body>
    </html>`,
  authSuccess: `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #2ecc71;">✅ Authentication Successful</h1>
        <p>You have successfully authenticated with Microsoft Graph API.</p>
        <p>You can close this window.</p>
      </body>
    </html>`,
  tokenExchangeError: (error) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">❌ Token Exchange Failed</h1>
        <p>Failed to exchange authorization code for access token.</p>
        <p><strong>Error:</strong> ${escapeHtml(error instanceof Error ? error.message : String(error))}</p>
        <p>You can close this window and try again.</p>
      </body>
    </html>`,
  tokenStatus: (status) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1>🔐 Token Status</h1>
        <p>${escapeHtml(status)}</p>
      </body>
    </html>`
};

function createAuthConfig() {
  const parsedScopes = (process.env.OUTLOOK_SCOPES || '').split(/\s+/).filter(Boolean);
  const envScopes = parsedScopes.length > 0 ? parsedScopes : null;

  return {
    clientId: process.env.OUTLOOK_CLIENT_ID || config.AUTH_CONFIG.clientId || '',
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET || config.AUTH_CONFIG.clientSecret || '',
    redirectUri: process.env.OUTLOOK_REDIRECT_URI || config.AUTH_CONFIG.redirectUri,
    scopes: envScopes || [...config.AUTH_CONFIG.scopes],
    tokenEndpoint: config.TOKEN_ENDPOINT,
    authEndpoint: config.AUTH_ENDPOINT
  };
}

function setupOAuthRoutes(app, authConfig) {
  if (!authConfig) {
    authConfig = createAuthConfig();
  }

  app.get('/auth', (req, res) => {
    if (!authConfig.clientId) {
      return res.status(500).send(templates.authError('Configuration Error', 'Client ID is not configured.'));
    }
    const state = crypto.randomBytes(16).toString('hex');

    const authorizationUrl = `${authConfig.authEndpoint}?` +
      querystring.stringify({
        client_id: authConfig.clientId,
        response_type: 'code',
        redirect_uri: authConfig.redirectUri,
        scope: authConfig.scopes.join(' '),
        response_mode: 'query',
        state: state
      });
    res.redirect(authorizationUrl);
  });

  app.get('/auth/callback', async (req, res) => {
    const { code, error, error_description, state } = req.query;

    if (!state) {
        console.error("OAuth callback received without a 'state' parameter. Rejecting request to prevent potential CSRF attack.");
        return res.status(400).send(templates.authError('Missing State Parameter', 'The state parameter was missing from the OAuth callback. This is a security risk. Please try authenticating again.'));
    }

    if (error) {
      return res.status(400).send(templates.authError(error, error_description));
    }

    if (!code) {
      return res.status(400).send(templates.authError('Missing Authorization Code', 'No authorization code was provided in the callback.'));
    }

    try {
      await tokenManager.exchangeCodeForTokens(code);
      res.send(templates.authSuccess);
    } catch (exchangeError) {
      console.error('Token exchange error:', exchangeError);
      res.status(500).send(templates.tokenExchangeError(exchangeError));
    }
  });

  app.get('/token-status', async (req, res) => {
    try {
      const token = await tokenManager.getAccessToken();
      if (token) {
        const expiryDate = new Date(tokenManager.getExpiryTime());
        res.send(templates.tokenStatus(`Access token is valid. Expires at: ${expiryDate.toLocaleString()}`));
      } else {
        res.send(templates.tokenStatus('No valid access token found. Please authenticate.'));
      }
    } catch (err) {
      res.status(500).send(templates.tokenStatus(`Error checking token status: ${err.message}`));
    }
  });
}

module.exports = {
  setupOAuthRoutes,
  createAuthConfig,
};
