/**
 * Shared response helpers for MCP tool handlers.
 *
 * Centralises the auth-error check and the MCP response envelope so
 * every handler can use one-liners instead of duplicated boilerplate.
 */

/**
 * Determine whether an error represents an authentication failure.
 * Handles the various shapes produced by the token-manager / Graph SDK.
 *
 * @param {Error} error
 * @returns {boolean}
 */
function isAuthError(error) {
  if (!error || typeof error !== 'object') {
    return false;
  }

  const message = typeof error.message === 'string' ? error.message : '';
  return !!(
    error.isAuthError ||
    error.authReason ||
    message === 'Authentication required' ||
    message.startsWith('Authentication required')
  );
}

/**
 * Build an MCP error response (isError: true).
 *
 * @param {string} message - Human-readable error description
 * @returns {object} MCP-compliant response object
 */
function makeErrorResponse(message) {
  return {
    isError: true,
    content: [{ type: 'text', text: message }],
  };
}

/**
 * Build a successful MCP response.
 *
 * @param {string} text - Response body text
 * @returns {object} MCP-compliant response object
 */
function makeResponse(text) {
  return {
    content: [{ type: 'text', text }],
  };
}

module.exports = {
  isAuthError,
  makeErrorResponse,
  makeResponse,
};
