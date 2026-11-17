const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleCreateTask(args) {
  const { listId, title, dueDateTime, importance = 'normal', note } = args;

  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!title) {
    return makeErrorResponse('title is required.');
  }

  try {
    const client = await getGraphClient();

    const task = { title, importance };
    if (dueDateTime) {
      task.dueDateTime = { dateTime: dueDateTime, timeZone: 'UTC' };
    }
    if (note) {
      task.body = { contentType: 'text', content: note };
    }

    const created = await client.api(`me/todo/lists/${listId}/tasks`).post(task);

    return makeResponse(`Task created successfully.\nID: ${created.id}\nTitle: ${created.title}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating task: ${error.message}`);
  }
}

module.exports = {
  handleCreateTask,
};
