/**
 * Folder permissions tools for Outlook MCP server
 * Requires Mail.ReadWrite.Shared scope
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List folder permissions
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFolderPermissions(args) {
  const { folderId } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required (from list-folders).');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await client.api(`me/mailFolders/${folderId}/permissions`).get();
    const permissions = response.value || [];
    
    const roleLabels = {
      none: 'None',
      owner: 'Owner',
      publishingEditor: 'Publishing Editor',
      editor: 'Editor',
      publishingAuthor: 'Publishing Author',
      author: 'Author',
      noneditingAuthor: 'Non-editing Author',
      reviewer: 'Reviewer',
      contributor: 'Contributor',
      custom: 'Custom',
    };
    
    const structured = permissions.map(p => ({
      id: p.id,
      emailAddress: p.emailAddress?.address,
      name: p.emailAddress?.name,
      role: p.role,
      roleLabel: roleLabels[p.role] || p.role,
      isInsideOrganization: p.isInsideOrganization,
      isRemovable: p.isRemovable,
    }));
    
    const textFallback = [
      `Folder Permissions (${permissions.length}):`,
      '',
      permissions.length === 0
        ? '  No explicit permissions (only default).'
        : permissions.map(p => 
            `- ${p.emailAddress?.name || 'Unknown'} <${p.emailAddress?.address || 'N/A'}>: ${roleLabels[p.role] || p.role} ${p.isInsideOrganization ? '(internal)' : '(external)'} ${p.isRemovable ? '[removable]' : ''}`
          ).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ permissions: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 403) {
      return makeErrorResponse('Permission denied. Ensure you have Mail.ReadWrite.Shared scope and owner permissions on the folder.');
    }
    return makeErrorResponse(`Error listing folder permissions: ${error.message}`);
  }
}

/**
 * Add folder permission
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAddFolderPermission(args) {
  const { folderId, emailAddress, role = 'reviewer' } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required (from list-folders).');
  }
  if (!emailAddress) {
    return makeErrorResponse('Email address is required.');
  }
  
  const validRoles = ['none', 'owner', 'publishingEditor', 'editor', 'publishingAuthor', 'author', 'noneditingAuthor', 'reviewer', 'contributor'];
  if (!validRoles.includes(role)) {
    return makeErrorResponse(`Invalid role. Valid roles: ${validRoles.join(', ')}`);
  }
  
  try {
    const client = await getGraphClient();
    
    const permission = {
      emailAddress: { address: emailAddress },
      role,
    };
    
    await client.api(`me/mailFolders/${folderId}/permissions`).post(permission);
    
    return makeResponse(`Folder permission added for ${emailAddress} as "${role}".`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 403) {
      return makeErrorResponse('Permission denied. Ensure you have Mail.ReadWrite.Shared scope and owner permissions.');
    }
    return makeErrorResponse(`Error adding folder permission: ${error.message}`);
  }
}

/**
 * Update folder permission
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateFolderPermission(args) {
  const { folderId, permissionId, role } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required.');
  }
  if (!permissionId) {
    return makeErrorResponse('Permission ID is required (from list-folder-permissions).');
  }
  if (!role) {
    return makeErrorResponse('New role is required.');
  }
  
  const validRoles = ['none', 'owner', 'publishingEditor', 'editor', 'publishingAuthor', 'author', 'noneditingAuthor', 'reviewer', 'contributor'];
  if (!validRoles.includes(role)) {
    return makeErrorResponse(`Invalid role. Valid roles: ${validRoles.join(', ')}`);
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/mailFolders/${folderId}/permissions/${permissionId}`).patch({ role });
    
    return makeResponse(`Folder permission updated to "${role}" successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating folder permission: ${error.message}`);
  }
}

/**
 * Remove folder permission
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRemoveFolderPermission(args) {
  const { folderId, permissionId } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required.');
  }
  if (!permissionId) {
    return makeErrorResponse('Permission ID is required (from list-folder-permissions).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/mailFolders/${folderId}/permissions/${permissionId}`).delete();
    
    return makeResponse('Folder permission removed successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error removing folder permission: ${error.message}`);
  }
}

module.exports = {
  handleListFolderPermissions,
  handleAddFolderPermission,
  handleUpdateFolderPermission,
  handleRemoveFolderPermission,
};