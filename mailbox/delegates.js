/**
 * Delegates management tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List delegates
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListDelegates(_args) {
  try {
    const client = await getGraphClient();
    
    const response = await client.api('me/mailboxSettings/delegates').get();
    const delegates = response.value || [];
    
    const permissionLabels = {
      none: 'None',
      editor: 'Editor (read/write)',
      reviewer: 'Reviewer (read only)',
      author: 'Author (read/create)',
      custom: 'Custom',
    };
    
    const structured = delegates.map(d => ({
      emailAddress: d.emailAddress?.address,
      name: d.emailAddress?.name,
      permissions: {
        calendar: permissionLabels[d.permissions?.calendar?.role] || d.permissions?.calendar?.role,
        inbox: permissionLabels[d.permissions?.inbox?.role] || d.permissions?.inbox?.role,
        contacts: permissionLabels[d.permissions?.contacts?.role] || d.permissions?.contacts?.role,
        notes: permissionLabels[d.permissions?.notes?.role] || d.permissions?.notes?.role,
        tasks: permissionLabels[d.permissions?.tasks?.role] || d.permissions?.tasks?.role,
        journal: permissionLabels[d.permissions?.journal?.role] || d.permissions?.journal?.role,
      },
      canViewPrivateItems: d.canViewPrivateItems,
      receiveCopiesOfMeetingMessages: d.receiveCopiesOfMeetingMessages,
    }));
    
    const textFallback = [
      `Delegates (${delegates.length}):`,
      '',
      delegates.length === 0
        ? '  No delegates configured.'
        : delegates.map((d, i) => {
            const perms = [];
            if (d.permissions?.calendar?.role) perms.push(`Calendar: ${permissionLabels[d.permissions.calendar.role] || d.permissions.calendar.role}`);
            if (d.permissions?.inbox?.role) perms.push(`Inbox: ${permissionLabels[d.permissions.inbox.role] || d.permissions.inbox.role}`);
            if (d.permissions?.contacts?.role) perms.push(`Contacts: ${permissionLabels[d.permissions.contacts.role] || d.permissions.contacts.role}`);
            return `${i + 1}. ${d.emailAddress?.name || 'Unknown'} <${d.emailAddress?.address || 'N/A'}>\n   ${perms.join(' | ')}\n   Private items: ${d.canViewPrivateItems ? 'Yes' : 'No'} | Meeting copies: ${d.receiveCopiesOfMeetingMessages ? 'Yes' : 'No'}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ delegates: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Delegates API not available or no delegates configured.');
    }
    return makeErrorResponse(`Error listing delegates: ${error.message}`);
  }
}

/**
 * Add delegate
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAddDelegate(args) {
  const { 
    emailAddress, 
    calendarRole = 'reviewer',
    inboxRole = 'none',
    contactsRole = 'none',
    canViewPrivateItems = false,
    receiveCopiesOfMeetingMessages = false,
  } = args;
  
  if (!emailAddress) {
    return makeErrorResponse('Email address of delegate is required.');
  }
  
  const validRoles = ['none', 'reviewer', 'author', 'editor', 'custom'];
  for (const role of [calendarRole, inboxRole, contactsRole]) {
    if (!validRoles.includes(role)) {
      return makeErrorResponse(`Invalid role: ${role}. Valid roles: ${validRoles.join(', ')}`);
    }
  }
  
  try {
    const client = await getGraphClient();
    
    // Get current delegates
    const currentResponse = await client.api('me/mailboxSettings/delegates').get();
    const currentDelegates = currentResponse.value || [];
    
    // Check if delegate already exists
    const exists = currentDelegates.some(d => d.emailAddress?.address?.toLowerCase() === emailAddress.toLowerCase());
    if (exists) {
      return makeErrorResponse(`Delegate ${emailAddress} already exists. Use update-delegate to modify.`);
    }
    
    const newDelegate = {
      emailAddress: { address: emailAddress },
      permissions: {
        calendar: { role: calendarRole },
        inbox: { role: inboxRole },
        contacts: { role: contactsRole },
      },
      canViewPrivateItems,
      receiveCopiesOfMeetingMessages,
    };
    
    const updatedDelegates = [...currentDelegates, newDelegate];
    
    await client.api('me/mailboxSettings/delegates').patch({ delegates: updatedDelegates });
    
    return makeResponse(`Delegate ${emailAddress} added successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error adding delegate: ${error.message}`);
  }
}

/**
 * Update delegate permissions
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateDelegate(args) {
  const { 
    emailAddress, 
    calendarRole,
    inboxRole,
    contactsRole,
    canViewPrivateItems,
    receiveCopiesOfMeetingMessages,
  } = args;
  
  if (!emailAddress) {
    return makeErrorResponse('Email address of delegate is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get current delegates
    const currentResponse = await client.api('me/mailboxSettings/delegates').get();
    const currentDelegates = currentResponse.value || [];
    
    const delegateIndex = currentDelegates.findIndex(d => d.emailAddress?.address?.toLowerCase() === emailAddress.toLowerCase());
    if (delegateIndex === -1) {
      return makeErrorResponse(`Delegate ${emailAddress} not found.`);
    }
    
    const delegate = currentDelegates[delegateIndex];
    
    // Update permissions
    if (calendarRole !== undefined) delegate.permissions.calendar = { role: calendarRole };
    if (inboxRole !== undefined) delegate.permissions.inbox = { role: inboxRole };
    if (contactsRole !== undefined) delegate.permissions.contacts = { role: contactsRole };
    if (canViewPrivateItems !== undefined) delegate.canViewPrivateItems = canViewPrivateItems;
    if (receiveCopiesOfMeetingMessages !== undefined) delegate.receiveCopiesOfMeetingMessages = receiveCopiesOfMeetingMessages;
    
    await client.api('me/mailboxSettings/delegates').patch({ delegates: currentDelegates });
    
    return makeResponse(`Delegate ${emailAddress} updated successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating delegate: ${error.message}`);
  }
}

/**
 * Remove delegate
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRemoveDelegate(args) {
  const { emailAddress } = args;
  
  if (!emailAddress) {
    return makeErrorResponse('Email address of delegate is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get current delegates
    const currentResponse = await client.api('me/mailboxSettings/delegates').get();
    const currentDelegates = currentResponse.value || [];
    
    const filteredDelegates = currentDelegates.filter(d => d.emailAddress?.address?.toLowerCase() !== emailAddress.toLowerCase());
    
    if (filteredDelegates.length === currentDelegates.length) {
      return makeErrorResponse(`Delegate ${emailAddress} not found.`);
    }
    
    await client.api('me/mailboxSettings/delegates').patch({ delegates: filteredDelegates });
    
    return makeResponse(`Delegate ${emailAddress} removed successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error removing delegate: ${error.message}`);
  }
}

module.exports = {
  handleListDelegates,
  handleAddDelegate,
  handleUpdateDelegate,
  handleRemoveDelegate,
};