const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleCompleteTask(args) {
  const { listId, taskId } = args;

  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }

  try {
    const client = await getGraphClient();
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).patch({
      status: 'completed',
      completedDateTime: { dateTime: new Date().toISOString(), timeZone: 'UTC' },
    });
    return makeResponse('Task marked as completed.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error completing task: ${error.message}`);
  }
}

module.exports = {
  handleCompleteTask,
};
