/**
 * Task categories tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List all categories (master category list)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListCategories(_args) {
  try {
    const client = await getGraphClient();
    
    const response = await client.api('me/outlook/masterCategories').get();
    const categories = response.value || [];
    
    const structured = categories.map(c => ({
      name: c.name,
      color: c.color,
      predefined: c.predefined,
    }));
    
    const colorLabels = {
      none: 'None',
      preset0: 'Red',
      preset1: 'Orange',
      preset2: 'Brown',
      preset3: 'Yellow',
      preset4: 'Green',
      preset5: 'Teal',
      preset6: 'Blue',
      preset7: 'Purple',
      preset8: 'Pink',
      preset9: 'Crimson',
      preset10: 'Steel',
      preset11: 'Dark Steel',
      preset12: 'Dark Red',
      preset13: 'Dark Orange',
      preset14: 'Dark Brown',
      preset15: 'Dark Yellow',
      preset16: 'Dark Green',
      preset17: 'Dark Teal',
      preset18: 'Dark Blue',
      preset19: 'Dark Purple',
      preset20: 'Dark Pink',
      preset21: 'Dark Crimson',
      preset22: 'Charcoal',
      preset23: 'Dark Steel',
      preset24: 'Gray',
    };
    
    const textFallback = [
      `Master Categories (${categories.length}):`,
      '',
      categories.length === 0
        ? '  No categories defined.'
        : categories.map((c, i) => {
            const colorName = colorLabels[c.color] || c.color;
            const pre = c.predefined ? ' (built-in)' : '';
            return `${i + 1}. ${c.name} [${colorName}]${pre}`;
          }).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ categories: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Master categories API not available.');
    }
    return makeErrorResponse(`Error listing categories: ${error.message}`);
  }
}

/**
 * Create a new category
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateCategory(args) {
  const { name, color = 'preset0' } = args;
  
  if (!name) {
    return makeErrorResponse('Category name is required.');
  }
  
  const validColors = [
    'none', 'preset0', 'preset1', 'preset2', 'preset3', 'preset4', 'preset5',
    'preset6', 'preset7', 'preset8', 'preset9', 'preset10', 'preset11',
    'preset12', 'preset13', 'preset14', 'preset15', 'preset16', 'preset17',
    'preset18', 'preset19', 'preset20', 'preset21', 'preset22', 'preset23', 'preset24'
  ];
  
  if (!validColors.includes(color)) {
    return makeErrorResponse(`Invalid color. Valid colors: ${validColors.join(', ')}`);
  }
  
  try {
    const client = await getGraphClient();
    
    const category = { name, color };
    const created = await client.api('me/outlook/masterCategories').post(category);
    
    return makeResponse(`Category "${name}" created successfully with color ${color}.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 400 && error.message.includes('already exists')) {
      return makeErrorResponse(`Category "${name}" already exists.`);
    }
    return makeErrorResponse(`Error creating category: ${error.message}`);
  }
}

/**
 * Update a category
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateCategory(args) {
  const { name, newName, color } = args;
  
  if (!name) {
    return makeErrorResponse('Current category name is required.');
  }
  if (!newName && !color) {
    return makeErrorResponse('At least one of newName or color is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const patch = {};
    if (newName) patch.name = newName;
    if (color) patch.color = color;
    
    await client.api(`me/outlook/masterCategories/${name}`).patch(patch);
    
    return makeResponse(`Category updated successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating category: ${error.message}`);
  }
}

/**
 * Delete a category
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteCategory(args) {
  const { name } = args;
  
  if (!name) {
    return makeErrorResponse('Category name is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/outlook/masterCategories/${name}`).delete();
    
    return makeResponse(`Category "${name}" deleted successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting category: ${error.message}`);
  }
}

/**
 * Get task categories (assigned to a task)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetTaskCategories(args) {
  const { listId, taskId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  
  try {
    const client = await getGraphClient();
    
    const task = await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).select('id,title,categories').get();
    
    const structured = {
      taskId: task.id,
      taskTitle: task.title,
      categories: task.categories || [],
    };
    
    const textFallback = [
      `Task Categories: ${task.title}`,
      `Task ID: ${task.id}`,
      '',
      task.categories?.length > 0
        ? `Categories: ${task.categories.join(', ')}`
        : 'No categories assigned.',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting task categories: ${error.message}`);
  }
}

/**
 * Set categories on a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSetTaskCategories(args) {
  const { listId, taskId, categories } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  if (!categories || !Array.isArray(categories)) {
    return makeErrorResponse('categories array is required (can be empty to clear).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).patch({ categories });
    
    return makeResponse(categories.length > 0
      ? `Task categories set to: ${categories.join(', ')}`
      : 'Task categories cleared.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error setting task categories: ${error.message}`);
  }
}

module.exports = {
  handleListCategories,
  handleCreateCategory,
  handleUpdateCategory,
  handleDeleteCategory,
  handleGetTaskCategories,
  handleSetTaskCategories,
};