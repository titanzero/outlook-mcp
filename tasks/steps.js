/**
 * Task steps/checklist tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List steps for a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListTaskSteps(args) {
  const { listId, taskId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await graphGetPaginated(client, `me/todo/lists/${listId}/tasks/${taskId}/steps`, {
      $select: 'id,title,status,createdDateTime,lastModifiedDateTime',
      $orderby: 'createdDateTime asc',
    });
    
    const steps = response.value || [];
    
    const structured = steps.map(s => ({
      id: s.id,
      title: s.title,
      status: s.status,
      createdDateTime: s.createdDateTime,
      lastModifiedDateTime: s.lastModifiedDateTime,
    }));
    
    const textFallback = [
      `Task Steps (${steps.length}):`,
      '',
      steps.length === 0
        ? '  No steps found.'
        : steps.map((s, i) => {
            const done = s.status === 'completed' ? ' ✓' : ' ☐';
            return `${i + 1}. ${done} ${s.title} [id: ${s.id}]`;
          }).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ steps: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing task steps: ${error.message}`);
  }
}

/**
 * Create a new step for a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateTaskStep(args) {
  const { listId, taskId, title } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  if (!title) {
    return makeErrorResponse('Step title is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const step = { title };
    const created = await client.api(`me/todo/lists/${listId}/tasks/${taskId}/steps`).post(step);
    
    return makeResponse(`Task step created successfully.\nID: ${created.id}\nTitle: ${created.title}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating task step: ${error.message}`);
  }
}

/**
 * Update a task step
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateTaskStep(args) {
  const { listId, taskId, stepId, title, status } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }
  if (!stepId) {
    return makeErrorResponse('stepId is required (from list-task-steps).');
  }
  
  try {
    const client = await getGraphClient();
    
    const patch = {};
    if (title !== undefined) patch.title = title;
    if (status !== undefined) patch.status = status;
    
    if (Object.keys(patch).length === 0) {
      return makeErrorResponse('At least one of title or status is required.');
    }
    
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}/steps/${stepId}`).patch(patch);
    
    return makeResponse('Task step updated successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating task step: ${error.message}`);
  }
}

/**
 * Delete a task step
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteTaskStep(args) {
  const { listId, taskId, stepId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }
  if (!stepId) {
    return makeErrorResponse('stepId is required (from list-task-steps).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}/steps/${stepId}`).delete();
    
    return makeResponse('Task step deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting task step: ${error.message}`);
  }
}

/**
 * Complete a task step
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCompleteTaskStep(args) {
  const { listId, taskId, stepId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }
  if (!stepId) {
    return makeErrorResponse('stepId is required (from list-task-steps).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}/steps/${stepId}`).patch({
      status: 'completed',
    });
    
    return makeResponse('Task step marked as completed.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error completing task step: ${error.message}`);
  }
}

module.exports = {
  handleListTaskSteps,
  handleCreateTaskStep,
  handleUpdateTaskStep,
  handleDeleteTaskStep,
  handleCompleteTaskStep,
};