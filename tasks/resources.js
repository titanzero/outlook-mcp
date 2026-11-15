/**
 * Tasks resources for MCP
 * Exposes task lists and tasks as addressable resources
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError } = require('../utils/response-helpers');
const { ResourceTemplate } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { getTaskListCompletions } = require('../completion');

/**
 * Register tasks resources with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerTasksResources(server) {
  // Resource for list of task lists
  server.resource(
    'outlook-task-lists',
    'outlook://tasks/lists',
    {
      title: 'Task Lists',
      description: 'List all Microsoft To Do task lists',
    },
    async uri => {
      try {
        const client = await getGraphClient();
        const response = await client
          .api('me/todo/lists')
          .select('id,displayName,isOwner,isShared,wellKnownListName')
          .get();

        const lists = response.value || [];

        const structured = lists.map(l => ({
          id: l.id,
          name: l.displayName,
          isOwner: l.isOwner,
          isShared: l.isShared,
          wellKnownName: l.wellKnownListName,
        }));

        const _textFallback =
          lists.length === 0
            ? 'No task lists found.'
            : `Task Lists (${lists.length}):\n\n` +
              lists
                .map(
                  l =>
                    `- ${l.displayName} ${l.wellKnownListName ? `(${l.wellKnownListName})` : ''} [id: ${l.id}]`
                )
                .join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify({ lists: structured }, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading task lists: ${error.message}`);
      }
    }
  );

  // Resource for individual task list with its tasks - with completion
  const taskListTemplate = new ResourceTemplate('outlook://tasks/lists/{listId}', {
    complete: {
      listId: getTaskListCompletions,
    },
  });

  server.resource(
    'outlook-task-list',
    taskListTemplate,
    {
      title: 'Task List',
      description: 'Access a specific task list with its tasks',
    },
    async (uri, { listId }) => {
      try {
        const client = await getGraphClient();

        // Get list info
        const list = await client.api(`me/todo/lists/${listId}`).get();

        // Get tasks
        const response = await graphGetPaginated(client, `me/todo/lists/${listId}/tasks`, {
          $select:
            'id,title,status,importance,dueDateTime,body,completedDateTime,categories,recurrence',
          $orderby: 'importance desc,status asc,dueDateTime/dateTime asc',
        });

        const tasks = response.value || [];

        const structured = {
          id: list.id,
          name: list.displayName,
          isOwner: list.isOwner,
          isShared: list.isShared,
          wellKnownName: list.wellKnownListName,
          tasks: tasks.map(t => ({
            id: t.id,
            title: t.title,
            status: t.status,
            importance: t.importance,
            dueDateTime: t.dueDateTime?.dateTime,
            dueTimeZone: t.dueDateTime?.timeZone,
            completedDateTime: t.completedDateTime?.dateTime,
            note: t.body?.content,
            categories: t.categories || [],
            recurrence: t.recurrence,
          })),
        };

        const _textFallback = [
          `Task List: ${list.displayName}`,
          `ID: ${list.id}`,
          `Tasks: ${tasks.length}`,
          '',
          tasks
            .map(t => {
              const due = t.dueDateTime?.dateTime ? ` (due: ${t.dueDateTime.dateTime})` : '';
              const done = t.status === 'completed' ? ' ✓' : '';
              const imp = t.importance !== 'normal' ? ` [${t.importance}]` : '';
              return `- ${t.title}${due}${imp}${done}`;
            })
            .join('\n'),
        ].join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading task list: ${error.message}`);
      }
    }
  );

  // Resource for individual task
  server.resource(
    'outlook-task',
    'outlook://tasks/{taskId}',
    {
      title: 'Task',
      description: 'Access a specific task by ID (requires listId in query)',
    },
    async (uri, vars) => {
      try {
        const client = await getGraphClient();
        const listId = vars.listId;

        if (!listId) {
          throw new Error('listId query parameter is required');
        }

        const task = await client
          .api(`me/todo/lists/${listId}/tasks/${vars.taskId}`)
          .select(
            'id,title,status,importance,dueDateTime,body,completedDateTime,categories,recurrence,linkedResources'
          )
          .get();

        const structured = {
          id: task.id,
          title: task.title,
          status: task.status,
          importance: task.importance,
          dueDateTime: task.dueDateTime?.dateTime,
          dueTimeZone: task.dueDateTime?.timeZone,
          completedDateTime: task.completedDateTime?.dateTime,
          note: task.body?.content,
          categories: task.categories || [],
          recurrence: task.recurrence,
          linkedResources: task.linkedResources || [],
        };

        const _textFallback = [
          `Task: ${task.title}`,
          `ID: ${task.id}`,
          `Status: ${task.status}`,
          `Importance: ${task.importance}`,
          task.dueDateTime?.dateTime
            ? `Due: ${task.dueDateTime.dateTime} (${task.dueDateTime.timeZone})`
            : '',
          task.completedDateTime?.dateTime ? `Completed: ${task.completedDateTime.dateTime}` : '',
          task.categories?.length ? `Categories: ${task.categories.join(', ')}` : '',
          task.note ? `\nNote: ${task.note}` : '',
        ]
          .filter(Boolean)
          .join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading task: ${error.message}`);
      }
    }
  );
}

module.exports = { registerTasksResources };
