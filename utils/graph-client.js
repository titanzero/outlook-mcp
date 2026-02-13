/**
 * Microsoft Graph SDK client wrapper
 * 
 * Replaces the old custom HTTP client (graph-api.js) with the official
 * @microsoft/microsoft-graph-client SDK.
 */
const { Client } = require('@microsoft/microsoft-graph-client');
const { ensureAuthenticated } = require('../auth');

// Singleton client instance (lazily initialized)
let _clientInstance = null;
let _lastToken = null;

/**
 * Custom AuthenticationProvider that integrates with our token manager.
 * Implements the interface required by @microsoft/microsoft-graph-client.
 */
class OutlookAuthProvider {
  /**
   * Called by the SDK before each request to get a valid access token.
   * @returns {Promise<string>} A valid access token
   */
  async getAccessToken() {
    const token = await ensureAuthenticated();
    return token;
  }
}

/**
 * Returns a configured Microsoft Graph SDK client.
 * The client uses our OutlookAuthProvider to handle token acquisition and refresh.
 * 
 * @returns {Promise<import('@microsoft/microsoft-graph-client').Client>}
 */
async function getGraphClient() {
  // Verify we can get a token (this also triggers refresh if needed)
  const token = await ensureAuthenticated();

  // Re-create client if token changed (e.g. after refresh) or first call
  if (!_clientInstance || _lastToken !== token) {
    _clientInstance = Client.initWithMiddleware({
      authProvider: new OutlookAuthProvider(),
    });
    _lastToken = token;
  }

  return _clientInstance;
}

/**
 * Paginated GET helper using the Graph SDK.
 * Follows @odata.nextLink automatically to fetch up to maxCount items.
 * 
 * @param {import('@microsoft/microsoft-graph-client').Client} client - Graph SDK client
 * @param {string} path - API path (e.g. 'me/messages')
 * @param {object} queryParams - OData query parameters ($top, $select, $filter, etc.)
 * @param {number} maxCount - Maximum items to retrieve (0 = unlimited)
 * @returns {Promise<object>} - { value: Array, '@odata.count': number }
 */
async function graphGetPaginated(client, path, queryParams = {}, maxCount = 0) {
  const allItems = [];
  let nextLink = null;

  // Build the initial request
  let request = client.api(path);

  // Apply query parameters
  for (const [key, value] of Object.entries(queryParams)) {
    switch (key) {
      case '$top':
        request = request.top(value);
        break;
      case '$select':
        request = request.select(value);
        break;
      case '$filter':
        request = request.filter(value);
        break;
      case '$orderby':
        request = request.orderby(value);
        break;
      case '$search':
        request = request.search(value);
        break;
      case '$count':
        request = request.count(value);
        break;
      case '$expand':
        request = request.expand(value);
        break;
      case '$skip':
        request = request.skip(value);
        break;
      default:
        // For non-standard OData params (e.g. startDateTime, endDateTime)
        request = request.query({ [key]: value });
        break;
    }
  }

  try {
    // First page
    const response = await request.get();

    if (response.value && Array.isArray(response.value)) {
      allItems.push(...response.value);
      console.error(`Pagination: Retrieved ${response.value.length} items, total so far: ${allItems.length}`);
    }

    nextLink = response['@odata.nextLink'];

    // Follow pagination
    while (nextLink) {
      if (maxCount > 0 && allItems.length >= maxCount) {
        console.error(`Pagination: Reached max count of ${maxCount}, stopping`);
        break;
      }

      const nextResponse = await client.api(nextLink).get();

      if (nextResponse.value && Array.isArray(nextResponse.value)) {
        allItems.push(...nextResponse.value);
        console.error(`Pagination: Retrieved ${nextResponse.value.length} items, total so far: ${allItems.length}`);
      }

      nextLink = nextResponse['@odata.nextLink'];
    }

    // Trim to exact count if needed
    const finalItems = maxCount > 0 ? allItems.slice(0, maxCount) : allItems;

    console.error(`Pagination complete: Retrieved ${finalItems.length} total items`);

    return {
      value: finalItems,
      '@odata.count': finalItems.length
    };
  } catch (error) {
    console.error('Error during paginated request:', error);
    throw error;
  }
}

/**
 * Reset the cached client instance (useful for testing).
 */
function resetClient() {
  _clientInstance = null;
  _lastToken = null;
}

module.exports = {
  getGraphClient,
  graphGetPaginated,
  OutlookAuthProvider,
  resetClient
};
