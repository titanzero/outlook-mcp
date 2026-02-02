/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const https = require('https');
const querystring = require('querystring');
const config = require('../config');

// Global variable to store tokens
let cachedTokens = null;

// Flag to prevent concurrent refresh attempts
let isRefreshing = false;

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`[DEBUG] Attempting to load tokens from: ${tokenPath}`);
    console.error(`[DEBUG] HOME directory: ${process.env.HOME}`);
    console.error(`[DEBUG] Full resolved path: ${tokenPath}`);
    
    // Log file existence and details
    if (!fs.existsSync(tokenPath)) {
      console.error('[DEBUG] Token file does not exist');
      return null;
    }
    
    const stats = fs.statSync(tokenPath);
    console.error(`[DEBUG] Token file stats:
      Size: ${stats.size} bytes
      Created: ${stats.birthtime}
      Modified: ${stats.mtime}`);
    
    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    console.error('[DEBUG] Token file contents length:', tokenData.length);
    console.error('[DEBUG] Token file first 200 characters:', tokenData.slice(0, 200));
    
    try {
      const tokens = JSON.parse(tokenData);
      console.error('[DEBUG] Parsed tokens keys:', Object.keys(tokens));
      
      // Log each key's value to see what's present
      Object.keys(tokens).forEach(key => {
        console.error(`[DEBUG] ${key}: ${typeof tokens[key]}`);
      });
      
      // Check for access token presence
      if (!tokens.access_token) {
        console.error('[DEBUG] No access_token found in tokens');
        return null;
      }
      
      // Check token expiration
      const now = Date.now();
      const expiresAt = tokens.expires_at || 0;
      
      console.error(`[DEBUG] Current time: ${now}`);
      console.error(`[DEBUG] Token expires at: ${expiresAt}`);
      
      if (now > expiresAt) {
        console.error('[DEBUG] Token has expired, attempting refresh...');
        // Don't return null - try to refresh the token
        if (tokens.refresh_token) {
          // Return expired tokens so caller knows refresh_token is available
          cachedTokens = tokens;
          return null; // Signal that refresh is needed
        }
        return null;
      }
      
      // Update the cache
      cachedTokens = tokens;
      return tokens;
    } catch (parseError) {
      console.error('[DEBUG] Error parsing token JSON:', parseError);
      return null;
    }
  } catch (error) {
    console.error('[DEBUG] Error loading token cache:', error);
    return null;
  }
}

/**
 * Saves authentication tokens to the token file
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`Saving tokens to: ${tokenPath}`);
    
    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    console.error('Tokens saved successfully');
    
    // Update the cache
    cachedTokens = tokens;
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error);
    return false;
  }
}

/**
 * Refreshes the access token using the refresh token
 * @returns {Promise<object|null>} - New tokens or null if refresh fails
 */
async function refreshAccessToken() {
  if (isRefreshing) {
    console.error('[DEBUG] Refresh already in progress, waiting...');
    // Wait for the other refresh to complete
    await new Promise(resolve => setTimeout(resolve, 1000));
    return cachedTokens;
  }

  // Load tokens to get refresh_token
  const tokenPath = config.AUTH_CONFIG.tokenStorePath;
  let tokens;
  try {
    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    tokens = JSON.parse(tokenData);
  } catch (error) {
    console.error('[DEBUG] Could not read tokens for refresh:', error.message);
    return null;
  }

  if (!tokens.refresh_token) {
    console.error('[DEBUG] No refresh token available');
    return null;
  }

  isRefreshing = true;
  console.error('[DEBUG] Refreshing access token...');

  return new Promise((resolve) => {
    const postData = querystring.stringify({
      client_id: config.AUTH_CONFIG.clientId,
      client_secret: config.AUTH_CONFIG.clientSecret,
      refresh_token: tokens.refresh_token,
      grant_type: 'refresh_token',
      scope: config.AUTH_CONFIG.scopes.join(' ')
    });

    const options = {
      hostname: 'login.microsoftonline.com',
      path: '/common/oauth2/v2.0/token',
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {
        isRefreshing = false;

        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const newTokens = JSON.parse(data);

            // Calculate expiration time
            const expiresAt = Date.now() + (newTokens.expires_in * 1000);
            newTokens.expires_at = expiresAt;

            // Save new tokens
            saveTokenCache(newTokens);
            console.error('[DEBUG] Token refresh successful, new expiry:', new Date(expiresAt).toISOString());

            resolve(newTokens);
          } catch (error) {
            console.error('[DEBUG] Error parsing refresh response:', error.message);
            resolve(null);
          }
        } else {
          console.error(`[DEBUG] Token refresh failed with status ${res.statusCode}: ${data}`);
          resolve(null);
        }
      });
    });

    req.on('error', (error) => {
      isRefreshing = false;
      console.error('[DEBUG] Network error during token refresh:', error.message);
      resolve(null);
    });

    req.write(postData);
    req.end();
  });
}

/**
 * Gets the current access token, loading from cache if necessary
 * Automatically refreshes the token if expired
 * @returns {Promise<string|null>} - The access token or null if not available
 */
async function getAccessToken() {
  // Check cache first
  if (cachedTokens && cachedTokens.access_token) {
    // Verify it's not expired (with 5 min buffer)
    const now = Date.now();
    const expiresAt = cachedTokens.expires_at || 0;

    if (now < expiresAt - (5 * 60 * 1000)) {
      return cachedTokens.access_token;
    }

    // Token is expired or about to expire, try to refresh
    console.error('[DEBUG] Cached token expired, attempting refresh...');
    const newTokens = await refreshAccessToken();
    if (newTokens) {
      return newTokens.access_token;
    }
  }

  // Try loading from file
  const tokens = loadTokenCache();
  if (tokens) {
    return tokens.access_token;
  }

  // No valid tokens, try refresh if we have a refresh token
  console.error('[DEBUG] No valid tokens, attempting refresh...');
  const refreshedTokens = await refreshAccessToken();
  return refreshedTokens ? refreshedTokens.access_token : null;
}

/**
 * Synchronous version of getAccessToken for backwards compatibility
 * @returns {string|null} - The access token or null if not available
 * @deprecated Use getAccessToken() (async) instead for automatic refresh
 */
function getAccessTokenSync() {
  if (cachedTokens && cachedTokens.access_token) {
    return cachedTokens.access_token;
  }

  const tokens = loadTokenCache();
  return tokens ? tokens.access_token : null;
}

/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + (3600 * 1000) // 1 hour
  };
  
  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,
  getAccessTokenSync,
  refreshAccessToken,
  createTestTokens
};
