/**
 * Folder hierarchy management tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Get full folder hierarchy tree
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetFolderTree(_args) {
  try {
    const client = await getGraphClient();
    
    // Get all folders
    const response = await graphGetPaginated(client, 'me/mailFolders', {
      $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,wellKnownName',
    });
    
    const folders = response.value || [];
    
    // Build hierarchy
    const folderMap = new Map();
    const rootFolders = [];
    
    folders.forEach(folder => {
      folderMap.set(folder.id, { ...folder, children: [] });
    });
    
    folders.forEach(folder => {
      if (folder.parentFolderId && folderMap.has(folder.parentFolderId)) {
        folderMap.get(folder.parentFolderId).children.push(folder.id);
      } else {
        rootFolders.push(folder.id);
      }
    });
    
    function buildTree(folderId, level = 0) {
      const folder = folderMap.get(folderId);
      if (!folder) return null;
      
      const node = {
        id: folder.id,
        name: folder.displayName,
        parentFolderId: folder.parentFolderId,
        totalItems: folder.totalItemCount || 0,
        unreadItems: folder.unreadItemCount || 0,
        childFolderCount: folder.childFolderCount || 0,
        wellKnownName: folder.wellKnownName || null,
        level,
      };
      
      if (folder.children.length > 0) {
        node.children = folder.children.map(id => buildTree(id, level + 1)).filter(Boolean);
      }
      
      return node;
    }
    
    const tree = rootFolders.map(id => buildTree(id)).filter(Boolean);
    
    const structured = { folders: tree };
    
    function formatTree(folders, indent = 0) {
      return folders
        .map(f => {
          const prefix = '  '.repeat(indent);
          const unread = f.unreadItems > 0 ? ` (${f.unreadItems} unread)` : '';
          const wellKnown = f.wellKnownName ? ` [${f.wellKnownName}]` : '';
          let line = `${prefix}📁 ${f.name}${wellKnown} [${f.totalItems} items${unread}]`;
          if (f.children) {
            line += '\n' + formatTree(f.children, indent + 1);
          }
          return line;
        })
        .join('\n');
    }
    
    const textFallback = `Mail Folder Hierarchy:\n\n${formatTree(tree)}`;
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting folder tree: ${error.message}`);
  }
}

/**
 * Move folder to new parent
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveFolder(args) {
  const { folderId, newParentFolderId } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required (from list-folders).');
  }
  if (!newParentFolderId) {
    return makeErrorResponse('New parent folder ID is required (from list-folders).');
  }
  
  if (folderId === newParentFolderId) {
    return makeErrorResponse('Cannot move folder into itself.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Check if new parent is a descendant of the folder (would create cycle)
    // Get all descendants of folderId
    const response = await graphGetPaginated(client, 'me/mailFolders', {
      $select: 'id,parentFolderId',
    });
    const allFolders = response.value || [];
    
    function isDescendant(childId, ancestorId) {
      let current = allFolders.find(f => f.id === childId);
      while (current && current.parentFolderId) {
        if (current.parentFolderId === ancestorId) return true;
        current = allFolders.find(f => f.id === current.parentFolderId);
      }
      return false;
    }
    
    if (isDescendant(newParentFolderId, folderId)) {
      return makeErrorResponse('Cannot move folder into its own descendant (would create a cycle).');
    }
    
    // Move the folder by updating parentFolderId
    await client.api(`me/mailFolders/${folderId}`).patch({
      parentFolderId: newParentFolderId,
    });
    
    return makeResponse('Folder moved successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 400) {
      return makeErrorResponse('Invalid operation. Cannot move well-known folders or create cycles.');
    }
    return makeErrorResponse(`Error moving folder: ${error.message}`);
  }
}

/**
 * Create nested folder path (creates intermediate folders)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateFolderPath(args) {
  const { path, parentFolder } = args;
  
  if (!path) {
    return makeErrorResponse('Folder path is required (e.g., "Projects/ClientA/2024").');
  }
  
  try {
    const client = await getGraphClient();
    
    // Split path into segments
    const segments = path.split('/').map(s => s.trim()).filter(Boolean);
    if (segments.length === 0) {
      return makeErrorResponse('Invalid folder path.');
    }
    
    let currentParentId = parentFolder || null;
    let createdFolders = [];
    let existingFolders = [];
    
    for (const segment of segments) {
      // Check if folder already exists under current parent
      const listResponse = await client.api(currentParentId ? `me/mailFolders/${currentParentId}/childFolders` : 'me/mailFolders').get();
      const existing = listResponse.value?.find(f => f.displayName.toLowerCase() === segment.toLowerCase());
      
      if (existing) {
        currentParentId = existing.id;
        existingFolders.push({ name: segment, id: existing.id });
      } else {
        // Create new folder
        const newFolder = { displayName: segment };
        if (currentParentId) {
          newFolder.parentFolderId = currentParentId;
        }
        const created = await client.api(currentParentId ? `me/mailFolders/${currentParentId}/childFolders` : 'me/mailFolders').post(newFolder);
        currentParentId = created.id;
        createdFolders.push({ name: segment, id: created.id });
      }
    }
    
    const structured = {
      path,
      created: createdFolders,
      existing: existingFolders,
      finalFolderId: currentParentId,
    };
    
    const textFallback = [
      `Folder Path: ${path}`,
      '',
      `Created Folders (${createdFolders.length}):`,
      createdFolders.length > 0
        ? createdFolders.map(f => `  ✓ ${f.name} [id: ${f.id}]`).join('\n')
        : '  None',
      '',
      `Existing Folders (${existingFolders.length}):`,
      existingFolders.length > 0
        ? existingFolders.map(f => `  → ${f.name} [id: ${f.id}]`).join('\n')
        : '  None',
      '',
      `Final Folder ID: ${currentParentId}`,
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating folder path: ${error.message}`);
  }
}

/**
 * Empty folder (move all emails to Deleted Items)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleEmptyFolder(args) {
  const { folderId, permanent = false } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required (from list-folders).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get all emails in folder
    const response = await graphGetPaginated(client, `me/mailFolders/${folderId}/messages`, {
      $select: 'id',
      $top: 1000,
    });
    
    const emails = response.value || [];
    
    if (emails.length === 0) {
      return makeResponse('Folder is already empty.');
    }
    
    let success = 0;
    let failed = 0;
    const errors = [];
    
    for (const email of emails) {
      try {
        if (permanent) {
          await client.api(`me/messages/${email.id}`).delete();
        } else {
          await client.api(`me/messages/${email.id}/move`).post({
            destinationId: 'deleteditems',
          });
        }
        success++;
      } catch (e) {
        failed++;
        errors.push({ emailId: email.id, error: e.message });
      }
    }
    
    const action = permanent ? 'Permanently deleted' : 'Moved to Deleted Items';
    
    const structured = {
      folderId,
      totalEmails: emails.length,
      success,
      failed,
      action,
      errors: errors.slice(0, 10),
    };
    
    const textFallback = [
      `Empty Folder: ${action}`,
      `Total Emails: ${emails.length}`,
      `Successful: ${success}`,
      `Failed: ${failed}`,
      '',
      errors.length > 0 ? `Errors (first 10):\n${errors.map(e => `  ${e.emailId}: ${e.error}`).join('\n')}` : 'No errors.',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error emptying folder: ${error.message}`);
  }
}

/**
 * Get folder path from root
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetFolderPath(args) {
  const { folderId } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required (from list-folders).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get all folders
    const response = await graphGetPaginated(client, 'me/mailFolders', {
      $select: 'id,displayName,parentFolderId',
    });
    
    const folders = response.value || [];
    const folderMap = new Map();
    folders.forEach(f => folderMap.set(f.id, f));
    
    // Build path
    const path = [];
    let current = folderMap.get(folderId);
    
    while (current) {
      path.unshift(current.displayName);
      if (current.parentFolderId) {
        current = folderMap.get(current.parentFolderId);
      } else {
        current = null;
      }
    }
    
    const structured = {
      folderId,
      path: path.join(' / '),
      depth: path.length - 1,
      segments: path,
    };
    
    const textFallback = [
      `Folder Path: ${path.join(' / ')}`,
      `Folder ID: ${folderId}`,
      `Depth: ${path.length - 1}`,
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting folder path: ${error.message}`);
  }
}

module.exports = {
  handleGetFolderTree,
  handleMoveFolder,
  handleCreateFolderPath,
  handleEmptyFolder,
  handleGetFolderPath,
};