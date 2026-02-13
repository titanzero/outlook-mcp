/**
 * Unified token management for Microsoft Graph API authentication.
 * 
 * Consolidates the previous token-manager.js and token-storage.js into
 * a single async module with:
 *   - Async file I/O (fs.promises)
 *   - Proactive refresh with 5-minute buffer
 *   - Concurrency guard for simultaneous refresh requests
 *   - Code exchange for OAuth flow
 *   - Token clearing for logout/re-auth
 */
const fs = require('fs').promises;
const fsSync = require('fs');
const https = require('https');
const querystring = require('querystring');
const config = require('../config');

// In-memory token cache
let cachedTokens = null;

// Concurrency guards
let _loadPromise = null;
let _refreshPromise = null;

// Last authentication/token error (for diagnostics in tools)
let lastErrorReason = null;

/**
 * Record the last token/auth related error in a structured way.
 * This is later used by ensureAuthenticated() and tools to provide
 * more helpful error messages to the user.
 * 
 * @param {string} code - Short machine-readable error code
 * @param {string} message - Human-readable description
 * @param {object} [extra] - Optional extra diagnostic data
 */
function setLastError(code, message, extra = {}) {
  lastErrorReason = { code, message, ...extra };
}

/**
 * Get the last recorded token/auth error.
 * @returns {{code: string, message: string, [key: string]: any} | null}
 */
function getLastErrorReason() {
  return lastErrorReason;
}

/**
 * Load tokens from the on-disk token file.
 * Uses a deduplication promise to prevent concurrent reads.
 * @returns {Promise<object|null>} Parsed tokens or null
 */
async function loadTokens() {
  if (cachedTokens) return cachedTokens;

  if (_loadPromise) return _loadPromise;

  _loadPromise = (async () => {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    try {
      const data = await fs.readFile(tokenPath, 'utf8');
      
      let tokens;
      try {
        tokens = JSON.parse(data);
      } catch (parseError) {
        console.error(`[token-manager] Invalid token JSON at ${tokenPath}: ${parseError.message}`);
        setLastError('TOKEN_FILE_INVALID_JSON', `Invalid token JSON at ${tokenPath}: ${parseError.message}`, {
          path: tokenPath,
        });
        return null;
      }

      if (!tokens || typeof tokens !== 'object') {
        console.error(`[token-manager] Token file at ${tokenPath} does not contain a JSON object`);
        setLastError(
          'TOKEN_FILE_INVALID_SHAPE',
          `Token file at ${tokenPath} does not contain a valid JSON object`,
          { path: tokenPath }
        );
        return null;
      }

      if (!tokens.access_token) {
        console.error(`[token-manager] Token file at ${tokenPath} has no access_token`);
        setLastError(
          'TOKEN_FILE_INVALID_SHAPE',
          `Token file at ${tokenPath} is missing access_token`,
          { path: tokenPath }
        );
        return null;
      }

      cachedTokens = tokens;
      // Successful load clears any previous error
      lastErrorReason = null;
      return tokens;
    } catch (error) {
      if (error.code === 'ENOENT') {
        console.error(`[token-manager] Token file not found at ${tokenPath}`);
        setLastError(
          'TOKEN_FILE_MISSING',
          `Token file not found at ${tokenPath}`,
          { path: tokenPath }
        );
      } else {
        console.error('[token-manager] Error loading tokens:', error.message);
        setLastError(
          'TOKEN_FILE_READ_ERROR',
          `Error loading token file at ${tokenPath}: ${error.message}`,
          { path: tokenPath }
        );
      }
      return null;
    }
  })().finally(() => {
    _loadPromise = null;
  });

  return _loadPromise;
}

/**
 * Save tokens to the on-disk token file and update in-memory cache.
 * @param {object} tokens - Token data to persist
 */
async function saveTokens(tokens) {
  if (!tokens) {
    console.warn('[token-manager] No tokens to save');
    return;
  }

  const tokenPath = config.AUTH_CONFIG.tokenStorePath;
  await fs.writeFile(tokenPath, JSON.stringify(tokens, null, 2));
  cachedTokens = tokens;
  // Successful save clears any previous error
  lastErrorReason = null;
  console.error('[token-manager] Tokens saved successfully');
}

/**
 * Check whether the current access token is expired or about to expire.
 * @param {object} tokens - Token object with expires_at field
 * @returns {boolean} True if expired or within the refresh buffer
 */
function isExpired(tokens) {
  if (!tokens || !tokens.expires_at) return true;
  return Date.now() >= (tokens.expires_at - config.TOKEN_REFRESH_BUFFER_MS);
}

/**
 * Refresh the access token using the stored refresh_token.
 * Deduplicates concurrent refresh calls (returns the same promise).
 * @returns {Promise<object>} New token object
 * @throws {Error} If refresh fails
 */
async function refreshAccessToken() {
  // Concurrency guard: reuse in-flight promise
  if (_refreshPromise) {
    console.error('[token-manager] Refresh already in progress, waiting…');
    return _refreshPromise;
  }

  // Ensure client credentials are configured
  if (!config.AUTH_CONFIG.clientId || !config.AUTH_CONFIG.clientSecret) {
    const message = 'Client ID or Client Secret is not configured. Set OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET in the environment.';
    console.error(`[token-manager] ${message}`);
    setLastError('CLIENT_CONFIG_MISSING', message);
    throw new Error(message);
  }

  // Ensure tokens are loaded
  const tokens = await loadTokens();

  // Also read directly from disk in case the in-memory cache is expired but disk has refresh_token
  let refreshToken = tokens?.refresh_token;
  if (!refreshToken) {
    try {
      const diskData = await fs.readFile(config.AUTH_CONFIG.tokenStorePath, 'utf8');
      const diskTokens = JSON.parse(diskData);
      refreshToken = diskTokens.refresh_token;
    } catch (_e) {
      // ignore
    }
  }

  if (!refreshToken) {
    throw new Error('No refresh token available');
  }

  console.error('[token-manager] Refreshing access token…');

  const postData = querystring.stringify({
    client_id: config.AUTH_CONFIG.clientId,
    client_secret: config.AUTH_CONFIG.clientSecret,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope: config.AUTH_CONFIG.scopes.join(' ')
  });

  _refreshPromise = new Promise((resolve, reject) => {
    const tokenUrl = new URL(config.TOKEN_ENDPOINT);
    const options = {
      protocol: tokenUrl.protocol,
      hostname: tokenUrl.hostname,
      port: tokenUrl.port || undefined,
      path: `${tokenUrl.pathname}${tokenUrl.search || ''}`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', async () => {
        try {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            const newTokens = JSON.parse(data);
            newTokens.expires_at = Date.now() + (newTokens.expires_in * 1000);

            // Preserve old refresh token if new one not provided
            if (!newTokens.refresh_token && refreshToken) {
              newTokens.refresh_token = refreshToken;
            }

            await saveTokens(newTokens);
            console.error('[token-manager] Token refresh successful, expires:', new Date(newTokens.expires_at).toISOString());
            resolve(newTokens);
          } else {
            let parsedBody = null;
            try {
              parsedBody = JSON.parse(data);
            } catch {
              // Non-JSON error body; keep as raw text
            }

            const errorDescription =
              parsedBody?.error_description ||
              parsedBody?.error ||
              data;

            const code = /invalid_client/i.test(data) || /AADSTS7000215/.test(data)
              ? 'REFRESH_FAILED_INVALID_CLIENT'
              : 'REFRESH_FAILED';

            console.error(`[token-manager] Refresh failed (${res.statusCode}): ${errorDescription}`);
            setLastError(code, `Token refresh failed with status ${res.statusCode}: ${errorDescription}`, {
              statusCode: res.statusCode,
              rawBody: data,
            });

            reject(new Error(`Token refresh failed with status ${res.statusCode}: ${errorDescription}`));
          }
        } catch (e) {
          reject(e);
        }
      });
    });

    req.on('error', (error) => {
      console.error('[token-manager] Network error during refresh:', error.message);
      setLastError('REFRESH_NETWORK_ERROR', `Network error during token refresh: ${error.message}`);
      reject(error);
    });

    req.write(postData);
    req.end();
  }).finally(() => {
    _refreshPromise = null;
  });

  return _refreshPromise;
}

/**
 * Get a valid access token, refreshing automatically if expired.
 * This is the primary method called by ensureAuthenticated().
 * @returns {Promise<string|null>} Access token string, or null if unavailable
 */
async function getAccessToken() {
  // Check in-memory cache first
  if (cachedTokens && cachedTokens.access_token && !isExpired(cachedTokens)) {
    // Cached token is valid; clear any previous error
    lastErrorReason = null;
    return cachedTokens.access_token;
  }

  // Try loading from disk
  const tokens = await loadTokens();
  if (tokens && tokens.access_token && !isExpired(tokens)) {
    // loadTokens already updated cachedTokens and cleared lastErrorReason
    return tokens.access_token;
  }

  // Token expired or missing — try refresh
  console.error('[token-manager] Token expired or missing, attempting refresh…');
  try {
    const refreshed = await refreshAccessToken();
    return refreshed.access_token;
  } catch (error) {
    console.error('[token-manager] Refresh failed:', error.message);
    // refreshAccessToken will have already set a detailed lastErrorReason
    return null;
  }
}

/**
 * Exchange an authorization code for tokens (OAuth code flow).
 * Used by the auth server callback.
 * @param {string} authCode - The authorization code from OAuth redirect
 * @returns {Promise<object>} Token response
 */
async function exchangeCodeForTokens(authCode) {
  if (!config.AUTH_CONFIG.clientId || !config.AUTH_CONFIG.clientSecret) {
    const message = 'Client ID or Client Secret is not configured. Cannot exchange code for tokens.';
    setLastError('CLIENT_CONFIG_MISSING', message);
    throw new Error(message);
  }

  console.error('[token-manager] Exchanging authorization code for tokens…');

  const postData = querystring.stringify({
    client_id: config.AUTH_CONFIG.clientId,
    client_secret: config.AUTH_CONFIG.clientSecret,
    grant_type: 'authorization_code',
    code: authCode,
    redirect_uri: config.AUTH_CONFIG.redirectUri,
    scope: config.AUTH_CONFIG.scopes.join(' ')
  });

  return new Promise((resolve, reject) => {
    const tokenUrl = new URL(config.TOKEN_ENDPOINT);
    const options = {
      protocol: tokenUrl.protocol,
      hostname: tokenUrl.hostname,
      port: tokenUrl.port || undefined,
      path: `${tokenUrl.pathname}${tokenUrl.search || ''}`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', async () => {
        try {
          const responseBody = JSON.parse(data);
          if (res.statusCode >= 200 && res.statusCode < 300) {
            const tokens = {
              access_token: responseBody.access_token,
              refresh_token: responseBody.refresh_token,
              expires_in: responseBody.expires_in,
              expires_at: Date.now() + (responseBody.expires_in * 1000),
              scope: responseBody.scope,
              token_type: responseBody.token_type
            };
            await saveTokens(tokens);
            console.error('[token-manager] Tokens exchanged and saved successfully');
            resolve(tokens);
          } else {
            const errorDescription =
              responseBody.error_description ||
              responseBody.error ||
              `Token exchange failed with status ${res.statusCode}`;

            const code = /invalid_client/i.test(errorDescription) || /AADSTS7000215/.test(errorDescription)
              ? 'CODE_EXCHANGE_INVALID_CLIENT'
              : 'CODE_EXCHANGE_FAILED';

            console.error('[token-manager] Code exchange failed:', responseBody);
            setLastError(code, `Token exchange failed: ${errorDescription}`, {
              statusCode: res.statusCode,
              body: responseBody,
            });

            reject(new Error(errorDescription));
          }
        } catch (e) {
          reject(new Error(`Error processing token response: ${e.message}`));
        }
      });
    });

    req.on('error', (error) => {
      console.error('[token-manager] Network error during code exchange:', error.message);
      setLastError('CODE_EXCHANGE_NETWORK_ERROR', `Network error during code exchange: ${error.message}`);
      reject(error);
    });

    req.write(postData);
    req.end();
  });
}

/**
 * Clear all tokens (in-memory and on-disk). Used for logout / force re-auth.
 */
async function clearTokens() {
  cachedTokens = null;
  try {
    await fs.unlink(config.AUTH_CONFIG.tokenStorePath);
    console.error('[token-manager] Token file deleted');
  } catch (error) {
    if (error.code === 'ENOENT') {
      console.error('[token-manager] Token file not found, nothing to delete');
    } else {
      console.error('[token-manager] Error deleting token file:', error.message);
    }
  }
}

/**
 * Synchronous token load — used only by checkAuthStatus where we need
 * a quick look at the cached state without triggering refresh.
 * @returns {object|null} Token object or null
 */
function loadTokenCacheSync() {
  if (cachedTokens) return cachedTokens;

  const tokenPath = config.AUTH_CONFIG.tokenStorePath;

  try {
    if (!fsSync.existsSync(tokenPath)) {
      setLastError('TOKEN_FILE_MISSING', `Token file not found at ${tokenPath}`, { path: tokenPath });
      return null;
    }

    const data = fsSync.readFileSync(tokenPath, 'utf8');

    let tokens;
    try {
      tokens = JSON.parse(data);
    } catch (parseError) {
      console.error(`[token-manager] Invalid token JSON at ${tokenPath}: ${parseError.message}`);
      setLastError('TOKEN_FILE_INVALID_JSON', `Invalid token JSON at ${tokenPath}: ${parseError.message}`, {
        path: tokenPath,
      });
      return null;
    }

    if (!tokens || typeof tokens !== 'object' || !tokens.access_token) {
      setLastError(
        'TOKEN_FILE_INVALID_SHAPE',
        `Token file at ${tokenPath} is missing access_token or has invalid structure`,
        { path: tokenPath }
      );
      return null;
    }

    cachedTokens = tokens;
    lastErrorReason = null;
    return tokens;
  } catch (_error) {
    // For sync path (used only in check-auth-status), avoid throwing; lastErrorReason
    // will already reflect more detailed async load failures when they occur.
    return null;
  }
}

/**
 * Get the expiry time of the currently cached tokens.
 * @returns {number} Unix timestamp (ms) or 0
 */
function getExpiryTime() {
  return cachedTokens?.expires_at || 0;
}

module.exports = {
  getAccessToken,
  refreshAccessToken,
  exchangeCodeForTokens,
  clearTokens,
  loadTokens,
  saveTokens,
  loadTokenCacheSync,
  getExpiryTime,
  isExpired,
  getLastErrorReason
};
