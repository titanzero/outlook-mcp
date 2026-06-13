const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleListTaskLists(args) {
  try {
    const client = await getGraphClient();
    const response = await client.api('me/todo/lists').select('id,displayName,isOwner,isShared').get();
    const lists = response.value || [];

    if (lists.length === 0) return makeResponse('No task lists found.');

    const structured = lists.map(l => ({ id: l.id, name: l.displayName, isOwner: l.isOwner, isShared: l.isShared }));
    const textFallback = lists.map(l => `- ${l.displayName} (id: ${l.id})`).join('\n');

    return makeResponse(formatResponse(structured, `Task lists (${lists.length}):\n${textFallback}`));
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error listing task lists: ${error.message}`);
  }
}

async function handleListTasks(args) {
  const { listId, includeCompleted = false } = args;

  if (!listId) return makeErrorResponse('listId is required (from list-task-lists).');

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

    if (tasks.length === 0) return makeResponse('No tasks found.');

    const structured = tasks.map(t => ({
      id: t.id,
      title: t.title,
      status: t.status,
      importance: t.importance,
      due: t.dueDateTime?.dateTime || null,
      note: t.body?.content || null,
    }));

    const textFallback = tasks.map(t => {
      const due = t.dueDateTime?.dateTime ? ` (due: ${t.dueDateTime.dateTime})` : '';
      const done = t.status === 'completed' ? ' ✓' : '';
      return `- [${t.importance}] ${t.title}${due}${done}`;
    }).join('\n');

    return makeResponse(formatResponse(structured, `Tasks (${tasks.length}):\n${textFallback}`));
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error listing tasks: ${error.message}`);
  }
}

async function handleCreateTask(args) {
  const { listId, title, dueDateTime, importance = 'normal', note } = args;

  if (!listId) return makeErrorResponse('listId is required (from list-task-lists).');
  if (!title) return makeErrorResponse('title is required.');

  try {
    const client = await getGraphClient();

    const task = { title, importance };
    if (dueDateTime) task.dueDateTime = { dateTime: dueDateTime, timeZone: 'UTC' };
    if (note) task.body = { contentType: 'text', content: note };

    const created = await client.api(`me/todo/lists/${listId}/tasks`).post(task);

    return makeResponse(`Task created successfully.\nID: ${created.id}\nTitle: ${created.title}`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error creating task: ${error.message}`);
  }
}

async function handleCompleteTask(args) {
  const { listId, taskId } = args;

  if (!listId) return makeErrorResponse('listId is required.');
  if (!taskId) return makeErrorResponse('taskId is required.');

  try {
    const client = await getGraphClient();
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).patch({
      status: 'completed',
      completedDateTime: { dateTime: new Date().toISOString(), timeZone: 'UTC' },
    });
    return makeResponse('Task marked as completed.');
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error completing task: ${error.message}`);
  }
}

async function handleDeleteTask(args) {
  const { listId, taskId } = args;

  if (!listId) return makeErrorResponse('listId is required.');
  if (!taskId) return makeErrorResponse('taskId is required.');

  try {
    const client = await getGraphClient();
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).delete();
    return makeResponse('Task deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error deleting task: ${error.message}`);
  }
}

const tasksTools = [
  {
    name: "list-task-lists",
    description: "Lists all Microsoft To Do task lists",
    inputSchema: { type: "object", properties: {}, required: [] },
    handler: handleListTaskLists
  },
  {
    name: "list-tasks",
    description: "Lists tasks in a specific To Do list. By default excludes completed tasks.",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "ID of the task list (from list-task-lists)" },
        includeCompleted: { type: "boolean", description: "Include completed tasks (default: false)" }
      },
      required: ["listId"]
    },
    handler: handleListTasks
  },
  {
    name: "create-task",
    description: "Creates a new task in a To Do list",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "ID of the task list (from list-task-lists)" },
        title: { type: "string", description: "Task title" },
        dueDateTime: { type: "string", description: "Due date/time in ISO 8601 format (UTC)" },
        importance: { type: "string", description: "Task importance", enum: ["low", "normal", "high"] },
        note: { type: "string", description: "Optional note/body for the task" }
      },
      required: ["listId", "title"]
    },
    handler: handleCreateTask
  },
  {
    name: "complete-task",
    description: "Marks a task as completed",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "ID of the task list" },
        taskId: { type: "string", description: "ID of the task to complete (from list-tasks)" }
      },
      required: ["listId", "taskId"]
    },
    handler: handleCompleteTask
  },
  {
    name: "delete-task",
    description: "Deletes a task from a To Do list",
    inputSchema: {
      type: "object",
      properties: {
        listId: { type: "string", description: "ID of the task list" },
        taskId: { type: "string", description: "ID of the task to delete (from list-tasks)" }
      },
      required: ["listId", "taskId"]
    },
    handler: handleDeleteTask
  }
];

module.exports = { tasksTools };
