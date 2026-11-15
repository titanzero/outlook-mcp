/**
 * Rename folder functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { getFolderIdByName, invalidateFolderCache } = require('../email/folder-utils');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Rename folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRenameFolder(args) {
  const folderName = args.name;
  const newName = args.newName;
  
  if (!folderName) {
    return makeErrorResponse('Folder name is required.');
  }
  
  if (!newName) {
    return makeErrorResponse('New folder name is required.');
  }
  
  if (folderName === newName) {
    return makeErrorResponse('New name must be different from the current name.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Find the folder to rename
    const folderId = await getFolderIdByName(folderName);
    if (!folderId) {
      return makeErrorResponse(`Folder "${folderName}" not found. Use 'list-folders' to see available folders.`);
    }
    
    // Check if a folder with the new name already exists
    const existingFolder = await getFolderIdByName(newName);
    if (existingFolder) {
      return makeErrorResponse(`A folder named "${newName}" already exists.`);
    }
    
    // Rename the folder using PATCH
    const response = await client.api(`me/mailFolders/${folderId}`).patch({
      displayName: newName
    });
    
    // Bust the folder cache
    invalidateFolderCache();
    
    if (response && response.id) {
      return makeResponse(`Successfully renamed folder "${folderName}" to "${newName}".`);
    } else {
      return makeErrorResponse('Failed to rename folder. The server didn\'t return a folder ID.');
    }
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    console.error(`Error renaming folder "${folderName}": ${error.message}`);
    return makeErrorResponse(`Error renaming folder: ${error.message}`);
  }
}

module.exports = handleRenameFolder;