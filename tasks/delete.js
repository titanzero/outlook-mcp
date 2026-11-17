const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleDeleteTask(args) {
  const { listId, taskId } = args;

  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }

  try {
    const client = await getGraphClient();
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).delete();
    return makeResponse('Task deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting task: ${error.message}`);
  }
}

module.exports = {
  handleDeleteTask,
};
