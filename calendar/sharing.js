/**
 * Calendar sharing tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Share a calendar with another user
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleShareCalendar(args) {
  const { calendarId, emailAddress, role = 'read', message } = args;
  
  if (!calendarId) {
    return makeErrorResponse('Calendar ID is required (from list-calendars).');
  }
  if (!emailAddress) {
    return makeErrorResponse('Email address of the person to share with is required.');
  }
  
  const validRoles = ['freeBusyRead', 'limitedRead', 'read', 'write', 'delegateWithoutPrivate', 'delegateWithPrivate'];
  if (!validRoles.includes(role)) {
    return makeErrorResponse(`Invalid role. Valid roles: ${validRoles.join(', ')}`);
  }
  
  try {
    const client = await getGraphClient();
    
    const permission = {
      emailAddress: { address: emailAddress },
      role,
    };
    
    if (message) {
      permission.message = message;
    }
    
    await client.api(`me/calendars/${calendarId}/permissions`).post(permission);
    
    const roleLabels = {
      freeBusyRead: 'Availability only (free/busy)',
      limitedRead: 'Limited details',
      read: 'Full details (read)',
      write: 'Editor (read/write)',
      delegateWithoutPrivate: 'Delegate (no private items)',
      delegateWithPrivate: 'Delegate (with private items)',
    };
    
    return makeResponse(`Calendar shared successfully with ${emailAddress} as "${roleLabels[role]}".`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 403) {
      return makeErrorResponse('Permission denied. Ensure you have Calendars.ReadWrite.Shared scope and owner permissions on the calendar.');
    }
    return makeErrorResponse(`Error sharing calendar: ${error.message}`);
  }
}

/**
 * List calendar permissions
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListCalendarPermissions(args) {
  const { calendarId } = args;
  
  if (!calendarId) {
    return makeErrorResponse('Calendar ID is required (from list-calendars).');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await client.api(`me/calendars/${calendarId}/permissions`).get();
    const permissions = response.value || [];
    
    const structured = permissions.map(p => ({
      id: p.id,
      emailAddress: p.emailAddress?.address,
      name: p.emailAddress?.name,
      role: p.role,
      isInsideOrganization: p.isInsideOrganization,
      isRemovable: p.isRemovable,
    }));
    
    const roleLabels = {
      freeBusyRead: 'Availability only',
      limitedRead: 'Limited details',
      read: 'Full details (read)',
      write: 'Editor (read/write)',
      delegateWithoutPrivate: 'Delegate (no private items)',
      delegateWithPrivate: 'Delegate (with private items)',
    };
    
    const textFallback = [
      `Calendar Permissions (${permissions.length}):`,
      '',
      permissions.length === 0
        ? '  No explicit permissions (only default).'
        : permissions.map(p => `- ${p.emailAddress?.name || 'Unknown'} <${p.emailAddress?.address || 'N/A'}>: ${roleLabels[p.role] || p.role} ${p.isInsideOrganization ? '(internal)' : '(external)'} ${p.isRemovable ? '[removable]' : ''}`).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ permissions: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing calendar permissions: ${error.message}`);
  }
}

/**
 * Update calendar permission
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateCalendarPermission(args) {
  const { calendarId, permissionId, role } = args;
  
  if (!calendarId) {
    return makeErrorResponse('Calendar ID is required.');
  }
  if (!permissionId) {
    return makeErrorResponse('Permission ID is required (from list-calendar-permissions).');
  }
  if (!role) {
    return makeErrorResponse('New role is required.');
  }
  
  const validRoles = ['freeBusyRead', 'limitedRead', 'read', 'write', 'delegateWithoutPrivate', 'delegateWithPrivate'];
  if (!validRoles.includes(role)) {
    return makeErrorResponse(`Invalid role. Valid roles: ${validRoles.join(', ')}`);
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/calendars/${calendarId}/permissions/${permissionId}`).patch({ role });
    
    return makeResponse(`Calendar permission updated to "${role}" successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating calendar permission: ${error.message}`);
  }
}

/**
 * Remove calendar permission
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRemoveCalendarPermission(args) {
  const { calendarId, permissionId } = args;
  
  if (!calendarId) {
    return makeErrorResponse('Calendar ID is required.');
  }
  if (!permissionId) {
    return makeErrorResponse('Permission ID is required (from list-calendar-permissions).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/calendars/${calendarId}/permissions/${permissionId}`).delete();
    
    return makeResponse('Calendar permission removed successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error removing calendar permission: ${error.message}`);
  }
}

/**
 * Get shared calendars (calendars shared with me)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetSharedCalendars(_args) {
  try {
    const client = await getGraphClient();
    
    // Get calendars shared with the user
    const response = await client
      .api('me/calendars')
      .filter("canShare eq true and isDefaultCalendar eq false")
      .select('id,name,color,isDefaultCalendar,canEdit,canShare,canViewPrivateItems,owner')
      .get();
    
    const calendars = response.value || [];
    
    // Also try to get shared calendars via the shared endpoint
    let sharedCalendars = [];
    try {
      const sharedResponse = await client.api('me/calendarGroups').get();
      if (sharedResponse.value) {
        for (const group of sharedResponse.value) {
          if (group.name === 'Other calendars' || group.name === 'Shared calendars') {
            const groupCalendars = await client.api(`me/calendarGroups/${group.id}/calendars`).get();
            sharedCalendars = groupCalendars.value || [];
            break;
          }
        }
      }
    } catch (e) {
      // Ignore
    }
    
    const structured = {
      ownedCalendars: calendars.map(c => ({
        id: c.id,
        name: c.name,
        color: c.color,
        isDefault: c.isDefaultCalendar,
        canEdit: c.canEdit,
        canShare: c.canShare,
        canViewPrivateItems: c.canViewPrivateItems,
        owner: c.owner?.address || c.owner?.name,
      })),
      sharedCalendars: sharedCalendars.map(c => ({
        id: c.id,
        name: c.name,
        color: c.color,
        owner: c.owner?.address || c.owner?.name,
      })),
    };
    
    const textFallback = [
      `Owned Calendars (${calendars.length}):`,
      calendars.map(c => `- ${c.name} ${c.isDefaultCalendar ? '(DEFAULT)' : ''} [${c.color}] ${c.canShare ? '📤' : ''} [id: ${c.id}]`).join('\n') || '  None',
      '',
      `Shared Calendars (${sharedCalendars.length}):`,
      sharedCalendars.length > 0
        ? sharedCalendars.map(c => `- ${c.name} (owner: ${c.owner}) [${c.color}] [id: ${c.id}]`).join('\n')
        : '  None found',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting shared calendars: ${error.message}`);
  }
}

module.exports = {
  handleShareCalendar,
  handleListCalendarPermissions,
  handleUpdateCalendarPermission,
  handleRemoveCalendarPermission,
  handleGetSharedCalendars,
};