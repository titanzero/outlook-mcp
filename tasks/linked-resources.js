/**
 * Task linked resources tools for Outlook MCP server
 * Links tasks to emails, events, files, etc.
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List linked resources for a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListTaskLinkedResources(args) {
  const { listId, taskId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await graphGetPaginated(client, `me/todo/lists/${listId}/tasks/${taskId}/linkedResources`, {
      $select: 'id,webUrl,applicationName,displayName,externalId',
    });
    
    const resources = response.value || [];
    
    const structured = resources.map(r => ({
      id: r.id,
      webUrl: r.webUrl,
      applicationName: r.applicationName,
      displayName: r.displayName,
      externalId: r.externalId,
    }));
    
    const textFallback = [
      `Task Linked Resources (${resources.length}):`,
      '',
      resources.length === 0
        ? '  No linked resources found.'
        : resources.map((r, i) => 
            `${i + 1}. ${r.displayName} (${r.applicationName})\n   URL: ${r.webUrl}\n   External ID: ${r.externalId}`
          ).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ linkedResources: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing linked resources: ${error.message}`);
  }
}

/**
 * Add linked resource to a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAddTaskLinkedResource(args) {
  const { listId, taskId, webUrl, applicationName, displayName, externalId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  if (!webUrl) {
    return makeErrorResponse('webUrl is required.');
  }
  if (!applicationName) {
    return makeErrorResponse('applicationName is required (e.g., "Outlook", "Teams", "SharePoint").');
  }
  if (!displayName) {
    return makeErrorResponse('displayName is required.');
  }
  if (!externalId) {
    return makeErrorResponse('externalId is required (unique identifier in the source system).');
  }
  
  try {
    const client = await getGraphClient();
    
    const resource = {
      webUrl,
      applicationName,
      displayName,
      externalId,
    };
    
    const created = await client.api(`me/todo/lists/${listId}/tasks/${taskId}/linkedResources`).post(resource);
    
    return makeResponse(`Linked resource added successfully.\nID: ${created.id}\nDisplay Name: ${created.displayName}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error adding linked resource: ${error.message}`);
  }
}

/**
 * Remove linked resource from a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRemoveTaskLinkedResource(args) {
  const { listId, taskId, linkedResourceId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }
  if (!linkedResourceId) {
    return makeErrorResponse('linkedResourceId is required (from list-task-linked-resources).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}/linkedResources/${linkedResourceId}`).delete();
    
    return makeResponse('Linked resource removed successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error removing linked resource: ${error.message}`);
  }
}

/**
 * Link email to task (helper)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleLinkEmailToTask(args) {
  const { listId, taskId, emailId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  if (!emailId) {
    return makeErrorResponse('emailId is required (from list-emails or search-emails).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get email details for display name
    const email = await client.api(`me/messages/${emailId}`).select('id,subject,webLink').get();
    
    const resource = {
      webUrl: email.webLink || `https://outlook.office.com/mail/inbox/id/${emailId}`,
      applicationName: 'Outlook',
      displayName: email.subject || 'Email',
      externalId: `message:${emailId}`,
    };
    
    const created = await client.api(`me/todo/lists/${listId}/tasks/${taskId}/linkedResources`).post(resource);
    
    return makeResponse(`Email linked to task successfully.\nEmail: ${email.subject}\nResource ID: ${created.id}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error linking email to task: ${error.message}`);
  }
}

/**
 * Link event to task (helper)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleLinkEventToTask(args) {
  const { listId, taskId, eventId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  if (!eventId) {
    return makeErrorResponse('eventId is required (from list-events).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get event details
    const event = await client.api(`me/events/${eventId}`).select('id,subject,webLink').get();
    
    const resource = {
      webUrl: event.webLink || `https://outlook.office.com/calendar/id/${eventId}`,
      applicationName: 'Outlook',
      displayName: event.subject || 'Calendar Event',
      externalId: `event:${eventId}`,
    };
    
    const created = await client.api(`me/todo/lists/${listId}/tasks/${taskId}/linkedResources`).post(resource);
    
    return makeResponse(`Event linked to task successfully.\nEvent: ${event.subject}\nResource ID: ${created.id}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error linking event to task: ${error.message}`);
  }
}

module.exports = {
  handleListTaskLinkedResources,
  handleAddTaskLinkedResource,
  handleRemoveTaskLinkedResource,
  handleLinkEmailToTask,
  handleLinkEventToTask,
};