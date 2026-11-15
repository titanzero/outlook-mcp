/**
 * Update task functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Update task handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateTask(args) {
  const { listId, taskId, title, dueDateTime, importance, note, status } = args;

  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }

  try {
    const client = await getGraphClient();

    const patch = {};
    if (title !== undefined) {
      patch.title = title;
    }
    if (importance !== undefined) {
      patch.importance = importance;
    }
    if (dueDateTime !== undefined) {
      patch.dueDateTime = dueDateTime ? { dateTime: dueDateTime, timeZone: 'UTC' } : null;
    }
    if (note !== undefined) {
      patch.body = note ? { contentType: 'text', content: note } : null;
    }
    if (status !== undefined) {
      patch.status = status;
      if (status === 'completed') {
        patch.completedDateTime = { dateTime: new Date().toISOString(), timeZone: 'UTC' };
      }
    }

    // If no fields to update, return early
    if (Object.keys(patch).length === 0) {
      return makeResponse('No changes provided.');
    }

    const updated = await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).patch(patch);

    return makeResponse(
      `Task updated successfully.\nID: ${updated.id}\nTitle: ${updated.title}\nStatus: ${updated.status}`
    );
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating task: ${error.message}`);
  }
}

module.exports = handleUpdateTask;
