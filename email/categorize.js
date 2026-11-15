/**
 * Categorize email functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Categorize email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCategorizeEmail(args) {
  const { id, categories } = args;

  if (!id) {
    return makeErrorResponse('Email ID is required.');
  }

  try {
    const client = await getGraphClient();

    const categoryList = categories
      ? categories
        .split(',')
        .map(c => c.trim())
        .filter(c => c)
      : [];

    const patch = {
      categories: categoryList,
    };

    await client.api(`me/messages/${id}`).patch(patch);

    if (categoryList.length === 0) {
      return makeResponse('Categories cleared from email.');
    } else {
      return makeResponse(`Categories updated: ${categoryList.join(', ')}`);
    }
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating email categories: ${error.message}`);
  }
}

module.exports = handleCategorizeEmail;
