const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleListTaskLists(_args) {
  try {
    const client = await getGraphClient();
    const response = await client
      .api('me/todo/lists')
      .select('id,displayName,isOwner,isShared')
      .get();
    const lists = response.value || [];

    if (lists.length === 0) {
      return makeResponse('No task lists found.');
    }

    const structured = lists.map(l => ({
      id: l.id,
      name: l.displayName,
      isOwner: l.isOwner,
      isShared: l.isShared,
    }));
    const textFallback = lists.map(l => `- ${l.displayName} (id: ${l.id})`).join('\n');

    return makeResponse(
      formatResponse(structured, `Task lists (${lists.length}):\n${textFallback}`)
    );
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing task lists: ${error.message}`);
  }
}

async function handleListTasks(_args) {
  const { listId, includeCompleted = false } = _args;

  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }

  try {
    const client = await getGraphClient();

    let request = client
      .api(`me/todo/lists/${listId}/tasks`)
      .select('id,title,status,importance,dueDateTime,body');

    if (!includeCompleted) {
      request = request.filter("status ne 'completed'");
    }

    const response = await request.get();
    const tasks = response.value || [];

    if (tasks.length === 0) {
      return makeResponse('No tasks found.');
    }

    const structured = tasks.map(t => ({
      id: t.id,
      title: t.title,
      status: t.status,
      importance: t.importance,
      due: t.dueDateTime?.dateTime || null,
      note: t.body?.content || null,
    }));

    const textFallback = tasks
      .map(t => {
        const due = t.dueDateTime?.dateTime ? ` (due: ${t.dueDateTime.dateTime})` : '';
        const done = t.status === 'completed' ? ' ✓' : '';
        return `- [${t.importance}] ${t.title}${due}${done}`;
      })
      .join('\n');

    return makeResponse(formatResponse(structured, `Tasks (${tasks.length}):\n${textFallback}`));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing tasks: ${error.message}`);
  }
}

module.exports = {
  handleListTaskLists,
  handleListTasks,
};
