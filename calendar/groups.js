/**
 * Calendar groups tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List calendar groups
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListCalendarGroups(_args) {
  try {
    const client = await getGraphClient();
    
    const response = await client.api('me/calendarGroups').get();
    const groups = response.value || [];
    
    // For each group, get its calendars
    const groupsWithCalendars = [];
    for (const group of groups) {
      try {
        const calendarsResponse = await client.api(`me/calendarGroups/${group.id}/calendars`).get();
        const calendars = calendarsResponse.value || [];
        groupsWithCalendars.push({
          ...group,
          calendars: calendars.map(c => ({
            id: c.id,
            name: c.name,
            color: c.color,
            isDefaultCalendar: c.isDefaultCalendar,
          })),
        });
      } catch (e) {
        groupsWithCalendars.push({ ...group, calendars: [] });
      }
    }
    
    const structured = groupsWithCalendars.map(g => ({
      id: g.id,
      name: g.name,
      classId: g.classId,
      calendars: g.calendars || [],
    }));
    
    const textFallback = [
      `Calendar Groups (${groups.length}):`,
      '',
      groupsWithCalendars.length === 0
        ? '  No calendar groups found.'
        : groupsWithCalendars.map((g, i) => {
            const calList = g.calendars.length > 0
              ? g.calendars.map(c => `    - ${c.name} [${c.color}]${c.isDefaultCalendar ? ' (DEFAULT)' : ''}`).join('\n')
              : '    (no calendars or access denied)';
            return `${i + 1}. ${g.name} [id: ${g.id}]\n${calList}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ calendarGroups: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing calendar groups: ${error.message}`);
  }
}

/**
 * Create a new calendar group
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateCalendarGroup(args) {
  const { name } = args;
  
  if (!name) {
    return makeErrorResponse('Calendar group name is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const group = await client.api('me/calendarGroups').post({ name });
    
    return makeResponse(`Calendar group "${name}" created successfully.\nID: ${group.id}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating calendar group: ${error.message}`);
  }
}

/**
 * Delete a calendar group
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteCalendarGroup(args) {
  const { groupId } = args;
  
  if (!groupId) {
    return makeErrorResponse('Calendar group ID is required (from list-calendar-groups).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/calendarGroups/${groupId}`).delete();
    
    return makeResponse('Calendar group deleted successfully. Note: Calendars in the group are not deleted, only the group itself.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting calendar group: ${error.message}`);
  }
}

/**
 * Add calendar to group
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAddCalendarToGroup(args) {
  const { groupId, calendarId } = args;
  
  if (!groupId) {
    return makeErrorResponse('Calendar group ID is required (from list-calendar-groups).');
  }
  if (!calendarId) {
    return makeErrorResponse('Calendar ID is required (from list-calendars).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get the calendar reference
    const calendar = await client.api(`me/calendars/${calendarId}`).select('id').get();
    
    // Add to group
    await client.api(`me/calendarGroups/${groupId}/calendars/$ref`).post({
      '@odata.id': `https://graph.microsoft.com/v1.0/me/calendars/${calendar.id}`,
    });
    
    return makeResponse(`Calendar added to group successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error adding calendar to group: ${error.message}`);
  }
}

/**
 * Remove calendar from group
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRemoveCalendarFromGroup(args) {
  const { groupId, calendarId } = args;
  
  if (!groupId) {
    return makeErrorResponse('Calendar group ID is required.');
  }
  if (!calendarId) {
    return makeErrorResponse('Calendar ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/calendarGroups/${groupId}/calendars/${calendarId}/$ref`).delete();
    
    return makeResponse('Calendar removed from group successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error removing calendar from group: ${error.message}`);
  }
}

module.exports = {
  handleListCalendarGroups,
  handleCreateCalendarGroup,
  handleDeleteCalendarGroup,
  handleAddCalendarToGroup,
  handleRemoveCalendarFromGroup,
};