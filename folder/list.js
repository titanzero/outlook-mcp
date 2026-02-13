/**
 * List folders functionality
 * Uses getAllFolders from email/folder-utils for recursive folder discovery.
 */
const { getAllFolders } = require('../email/folder-utils');
const { formatResponse } = require('../utils/response-formatter');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Build structured rows for TOON encoding.
 * @param {Array} folders - Folder objects from getAllFolders()
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @returns {Array<object>}
 */
function buildFolderRows(folders, includeItemCounts) {
  return folders.map(folder => {
    const row = {
      name: folder.displayName,
      path: folder.path || folder.displayName,
    };
    if (includeItemCounts) {
      row.total = folder.totalItemCount || 0;
      row.unread = folder.unreadItemCount || 0;
    }
    return row;
  });
}

/**
 * List folders handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFolders(args) {
  const includeItemCounts = args.includeItemCounts === true;
  const includeChildren = args.includeChildren === true;
  
  try {
    // Get all mail folders recursively (with path info)
    const folders = await getAllFolders();

    if (!folders || folders.length === 0) {
      return makeResponse('No folders found.');
    }

    const textFallback = includeChildren
      ? formatFolderHierarchy(folders, includeItemCounts)
      : formatFolderList(folders, includeItemCounts);

    const text = formatResponse(
      { folders: buildFolderRows(folders, includeItemCounts) },
      textFallback
    );
    return makeResponse(text);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error listing folders: ${error.message}`);
  }
}

/**
 * Format folders as a flat list
 * @param {Array} folders - Array of folder objects
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @returns {string} - Formatted list
 */
function formatFolderList(folders, includeItemCounts) {
  if (!folders || folders.length === 0) {
    return "No folders found.";
  }
  
  // Sort folders alphabetically, with well-known folders first
  const wellKnownFolderNames = ['Inbox', 'Drafts', 'Sent Items', 'Deleted Items', 'Junk Email', 'Archive'];
  
  const sortedFolders = [...folders].sort((a, b) => {
    // Well-known folders come first
    const aIsWellKnown = wellKnownFolderNames.includes(a.displayName);
    const bIsWellKnown = wellKnownFolderNames.includes(b.displayName);
    
    if (aIsWellKnown && !bIsWellKnown) return -1;
    if (!aIsWellKnown && bIsWellKnown) return 1;
    
    if (aIsWellKnown && bIsWellKnown) {
      // Sort well-known folders by their index in the array
      return wellKnownFolderNames.indexOf(a.displayName) - wellKnownFolderNames.indexOf(b.displayName);
    }
    
    // Sort other folders alphabetically
    return a.displayName.localeCompare(b.displayName);
  });
  
  // Format each folder
  const folderLines = sortedFolders.map(folder => {
    let folderInfo = folder.displayName;
    
    // Show path for nested folders
    if (folder.path && folder.path.includes('/')) {
      folderInfo = folder.path;
    }
    
    // Add item counts if requested
    if (includeItemCounts) {
      const unreadCount = folder.unreadItemCount || 0;
      const totalCount = folder.totalItemCount || 0;
      folderInfo += ` - ${totalCount} items`;
      
      if (unreadCount > 0) {
        folderInfo += ` (${unreadCount} unread)`;
      }
    }
    
    return folderInfo;
  });
  
  return `Found ${folders.length} folders:\n\n${folderLines.join('\n')}`;
}

/**
 * Format folders as a hierarchical tree
 * @param {Array} folders - Array of folder objects
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @returns {string} - Formatted hierarchy
 */
function formatFolderHierarchy(folders, includeItemCounts) {
  if (!folders || folders.length === 0) {
    return "No folders found.";
  }
  
  // Build folder hierarchy
  const folderMap = new Map();
  const rootFolders = [];
  
  // First pass: create map of all folders
  folders.forEach(folder => {
    folderMap.set(folder.id, {
      ...folder,
      children: []
    });
    
    // Top-level folders have no '/' in their path
    if (!folder.path || !folder.path.includes('/')) {
      rootFolders.push(folder.id);
    }
  });
  
  // Second pass: build hierarchy
  folders.forEach(folder => {
    if (folder.path && folder.path.includes('/') && folder.parentFolderId) {
      const parent = folderMap.get(folder.parentFolderId);
      if (parent) {
        parent.children.push(folder.id);
      } else {
        // Fallback for orphaned folders
        rootFolders.push(folder.id);
      }
    }
  });
  
  // Format hierarchy recursively
  function formatSubtree(folderId, level = 0) {
    const folder = folderMap.get(folderId);
    if (!folder) return '';
    
    const indent = '  '.repeat(level);
    let line = `${indent}${folder.displayName}`;
    
    // Show usable path for nested folders
    if (folder.path && folder.path.includes('/')) {
      line += `  [path: ${folder.path}]`;
    }
    
    // Add item counts if requested
    if (includeItemCounts) {
      const unreadCount = folder.unreadItemCount || 0;
      const totalCount = folder.totalItemCount || 0;
      line += ` - ${totalCount} items`;
      
      if (unreadCount > 0) {
        line += ` (${unreadCount} unread)`;
      }
    }
    
    // Add children
    const childLines = folder.children
      .map(childId => formatSubtree(childId, level + 1))
      .filter(line => line.length > 0)
      .join('\n');
    
    return childLines.length > 0 ? `${line}\n${childLines}` : line;
  }
  
  // Format all root folders
  const formattedHierarchy = rootFolders
    .map(folderId => formatSubtree(folderId))
    .join('\n');
  
  return `Folder Hierarchy:\n\n${formattedHierarchy}`;
}

module.exports = handleListFolders;
