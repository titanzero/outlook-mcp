const { getGraphClient } = require('../utils/graph-client');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { registerTasksResources } = require('./resources');
const { handleListTaskSteps, handleCreateTaskStep, handleUpdateTaskStep, handleDeleteTaskStep, handleCompleteTaskStep } = require('./steps');
const { handleSetTaskRecurrence, handleGetTaskRecurrence, handleRemoveTaskRecurrence, handleRecurrenceExamples } = require('./recurrence');
const { handleListTaskAttachments, handleUploadTaskAttachment, handleDownloadTaskAttachment, handleDeleteTaskAttachment } = require('./attachments');
const { handleListTaskLinkedResources, handleAddTaskLinkedResource, handleRemoveTaskLinkedResource, handleLinkEmailToTask, handleLinkEventToTask } = require('./linked-resources');
const { handleListCategories, handleCreateCategory, handleUpdateCategory, handleDeleteCategory, handleGetTaskCategories, handleSetTaskCategories } = require('./categories');
const { z } = require('zod');

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

const handleUpdateTask = require('./update');

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

/**
 * Register tasks tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerTasksTools(server) {
  server.tool(
    'list-task-lists',
    'Lists all Microsoft To Do task lists',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListTaskLists
  );

  server.tool(
    'list-tasks',
    'Lists tasks in a specific To Do list. By default excludes completed tasks.',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      includeCompleted: z.boolean().optional().describe('Include completed tasks (default: false)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListTasks
  );

  server.tool(
    'create-task',
    'Creates a new task in a To Do list',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      title: z.string().describe('Task title'),
      dueDateTime: z.string().optional().describe('Due date/time in ISO 8601 format (UTC)'),
      importance: z.enum(['low', 'normal', 'high']).optional().describe('Task importance'),
      note: z.string().optional().describe('Optional note/body for the task'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateTask
  );

  server.tool(
    'update-task',
    'Updates an existing task in a To Do list. Can modify title, due date, importance, note, or status.',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task to update (from list-tasks)'),
      title: z.string().optional().describe('New task title'),
      dueDateTime: z
        .string()
        .optional()
        .describe('New due date/time in ISO 8601 format (UTC). Set to empty string to clear.'),
      importance: z.enum(['low', 'normal', 'high']).optional().describe('New task importance'),
      note: z
        .string()
        .optional()
        .describe('New note/body for the task. Set to empty string to clear.'),
      status: z
        .enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'])
        .optional()
        .describe('New task status'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateTask
  );

  server.tool(
    'complete-task',
    'Marks a task as completed',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task to complete (from list-tasks)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleCompleteTask
  );

  server.tool(
    'delete-task',
    'Deletes a task from a To Do list',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task to delete (from list-tasks)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteTask
  );

  // Task Steps / Checklist tools (requires Tasks.ReadWrite)
  server.tool(
    'list-task-steps',
    'Lists checklist steps for a task',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListTaskSteps
  );

  server.tool(
    'create-task-step',
    'Creates a new checklist step for a task',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
      title: z.string().describe('Step title'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateTaskStep
  );

  server.tool(
    'update-task-step',
    'Updates a task step (title or status)',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task'),
      stepId: z.string().describe('ID of the step (from list-task-steps)'),
      title: z.string().optional().describe('New step title'),
      status: z.enum(['notStarted', 'completed']).optional().describe('Step status'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateTaskStep
  );

  server.tool(
    'delete-task-step',
    'Deletes a task step',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task'),
      stepId: z.string().describe('ID of the step (from list-task-steps)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteTaskStep
  );

  server.tool(
    'complete-task-step',
    'Marks a task step as completed',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task'),
      stepId: z.string().describe('ID of the step (from list-task-steps)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleCompleteTaskStep
  );

  // Task Recurrence tools (requires Tasks.ReadWrite)
  server.tool(
    'set-task-recurrence',
    'Sets recurrence pattern on a task (daily, weekly, monthly, yearly)',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
      pattern: z.object({
        type: z.enum(['daily', 'weekly', 'monthly', 'yearly', 'absoluteMonthly', 'relativeMonthly']).describe('Recurrence type'),
        interval: z.number().int().positive().max(999).optional().describe('Interval (default: 1)'),
        dayOfMonth: z.number().int().min(1).max(31).optional().describe('Day of month (for monthly/yearly)'),
        daysOfWeek: z.array(z.enum(['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'])).optional().describe('Days of week (for weekly/relativeMonthly)'),
        month: z.number().int().min(1).max(12).optional().describe('Month (for yearly)'),
        index: z.enum(['first', 'second', 'third', 'fourth', 'last']).optional().describe('Week index (for relativeMonthly)'),
        firstDayOfWeek: z.enum(['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']).optional().describe('First day of week'),
      }).describe('Recurrence pattern object'),
      range: z.object({
        type: z.enum(['noEnd', 'endDate', 'numbered']).describe('Range type'),
        startDate: z.string().describe('Start date (YYYY-MM-DD)'),
        endDate: z.string().optional().describe('End date (YYYY-MM-DD, for endDate type)'),
        recurrenceCount: z.number().int().positive().optional().describe('Number of occurrences (for numbered type)'),
      }).describe('Recurrence range object'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetTaskRecurrence
  );

  server.tool(
    'get-task-recurrence',
    'Gets the recurrence pattern for a task',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetTaskRecurrence
  );

  server.tool(
    'remove-task-recurrence',
    'Removes recurrence from a task',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleRemoveTaskRecurrence
  );

  server.tool(
    'recurrence-examples',
    'Returns examples of recurrence patterns and ranges for tasks',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleRecurrenceExamples
  );

  // Task Attachments tools (requires Tasks.ReadWrite)
  server.tool(
    'list-task-attachments',
    'Lists attachments for a task',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListTaskAttachments
  );

  server.tool(
    'upload-task-attachment',
    'Uploads an attachment to a task (base64-encoded, max 25MB)',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
      name: z.string().describe('Attachment file name'),
      contentBytes: z.string().describe('Base64-encoded file content'),
      contentType: z.string().optional().describe("MIME type (e.g., 'application/pdf', 'image/png'). Auto-detected if omitted."),
      isInline: z.boolean().optional().describe('Whether the attachment is inline (default: false)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleUploadTaskAttachment
  );

  server.tool(
    'download-task-attachment',
    'Downloads an attachment from a task (returns base64 content)',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task'),
      attachmentId: z.string().describe('ID of the attachment (from list-task-attachments)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleDownloadTaskAttachment
  );

  server.tool(
    'delete-task-attachment',
    'Deletes an attachment from a task',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task'),
      attachmentId: z.string().describe('ID of the attachment (from list-task-attachments)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteTaskAttachment
  );

  // Task Linked Resources tools (requires Tasks.ReadWrite)
  server.tool(
    'list-task-linked-resources',
    'Lists linked resources for a task (emails, events, files, etc.)',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListTaskLinkedResources
  );

  server.tool(
    'add-task-linked-resource',
    'Adds a linked resource to a task (URL, app name, display name, external ID)',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
      webUrl: z.string().describe('URL of the linked resource'),
      applicationName: z.string().describe('Source application name (e.g., "Outlook", "Teams", "SharePoint")'),
      displayName: z.string().describe('Display name for the link'),
      externalId: z.string().describe('Unique identifier in the source system'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleAddTaskLinkedResource
  );

  server.tool(
    'remove-task-linked-resource',
    'Removes a linked resource from a task',
    {
      listId: z.string().describe('ID of the task list'),
      taskId: z.string().describe('ID of the task'),
      linkedResourceId: z.string().describe('ID of the linked resource (from list-task-linked-resources)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleRemoveTaskLinkedResource
  );

  server.tool(
    'link-email-to-task',
    'Links an Outlook email to a task',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
      emailId: z.string().describe('ID of the email (from list-emails or search-emails)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleLinkEmailToTask
  );

  server.tool(
    'link-event-to-task',
    'Links an Outlook calendar event to a task',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
      eventId: z.string().describe('ID of the event (from list-events)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleLinkEventToTask
  );

  // Task Categories tools (requires Tasks.ReadWrite)
  server.tool(
    'list-categories',
    'Lists all master categories (Outlook category list)',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListCategories
  );

  server.tool(
    'create-category',
    'Creates a new master category',
    {
      name: z.string().describe('Category name'),
      color: z.enum([
        'none', 'preset0', 'preset1', 'preset2', 'preset3', 'preset4', 'preset5',
        'preset6', 'preset7', 'preset8', 'preset9', 'preset10', 'preset11',
        'preset12', 'preset13', 'preset14', 'preset15', 'preset16', 'preset17',
        'preset18', 'preset19', 'preset20', 'preset21', 'preset22', 'preset23', 'preset24'
      ]).optional().describe('Color preset (default: preset0)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateCategory
  );

  server.tool(
    'update-category',
    'Updates a master category (rename or change color)',
    {
      name: z.string().describe('Current category name'),
      newName: z.string().optional().describe('New category name'),
      color: z.enum([
        'none', 'preset0', 'preset1', 'preset2', 'preset3', 'preset4', 'preset5',
        'preset6', 'preset7', 'preset8', 'preset9', 'preset10', 'preset11',
        'preset12', 'preset13', 'preset14', 'preset15', 'preset16', 'preset17',
        'preset18', 'preset19', 'preset20', 'preset21', 'preset22', 'preset23', 'preset24'
      ]).optional().describe('New color preset'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateCategory
  );

  server.tool(
    'delete-category',
    'Deletes a master category',
    {
      name: z.string().describe('Category name to delete'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteCategory
  );

  server.tool(
    'get-task-categories',
    'Gets categories assigned to a task',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetTaskCategories
  );

  server.tool(
    'set-task-categories',
    'Sets categories on a task (replaces existing)',
    {
      listId: z.string().describe('ID of the task list (from list-task-lists)'),
      taskId: z.string().describe('ID of the task (from list-tasks)'),
      categories: z.array(z.string()).describe('Array of category names (empty to clear)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetTaskCategories
  );
}

module.exports = {
  registerTasksTools,
  registerTasksResources,
  handleListTaskLists,
  handleListTasks,
  handleCreateTask,
  handleUpdateTask,
  handleCompleteTask,
  handleDeleteTask,
  handleListTaskSteps,
  handleCreateTaskStep,
  handleUpdateTaskStep,
  handleDeleteTaskStep,
  handleCompleteTaskStep,
  handleSetTaskRecurrence,
  handleGetTaskRecurrence,
  handleRemoveTaskRecurrence,
  handleRecurrenceExamples,
  handleListTaskAttachments,
  handleUploadTaskAttachment,
  handleDownloadTaskAttachment,
  handleDeleteTaskAttachment,
  handleListTaskLinkedResources,
  handleAddTaskLinkedResource,
  handleRemoveTaskLinkedResource,
  handleLinkEmailToTask,
  handleLinkEventToTask,
  handleListCategories,
  handleCreateCategory,
  handleUpdateCategory,
  handleDeleteCategory,
  handleGetTaskCategories,
  handleSetTaskCategories,
};
