/**
 * Delete folder functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { getFolderIdByName, invalidateFolderCache } = require('../email/folder-utils');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Delete folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteFolder(args) {
  const folderName = args.name;

  if (!folderName) {
    return makeErrorResponse('Folder name is required.');
  }

  try {
    const client = await getGraphClient();

    // Find the folder to delete
    const folderId = await getFolderIdByName(folderName);
    if (!folderId) {
      return makeErrorResponse(
        `Folder "${folderName}" not found. Use 'list-folders' to see available folders.`
      );
    }

    // Check if it's a well-known folder that shouldn't be deleted
    const wellKnownFolders = [
      'Inbox',
      'Drafts',
      'Sent Items',
      'Deleted Items',
      'Junk Email',
      'Archive',
      'Outbox',
      'Conversation History',
    ];
    if (wellKnownFolders.includes(folderName)) {
      return makeErrorResponse(
        `Cannot delete well-known folder "${folderName}". These folders are system-managed.`
      );
    }

    // Delete the folder
    await client.api(`me/mailFolders/${folderId}`).delete();

    // Bust the folder cache
    invalidateFolderCache();

    return makeResponse(`Successfully deleted folder "${folderName}".`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }

    console.error(`Error deleting folder "${folderName}": ${error.message}`);

    // Check for specific error about non-empty folder
    if (error.message.includes('FolderNotEmpty') || error.message.includes('folder is not empty')) {
      return makeErrorResponse(
        `Folder "${folderName}" is not empty. Move or delete its contents first, or delete subfolders first.`
      );
    }

    return makeErrorResponse(`Error deleting folder: ${error.message}`);
  }
}

module.exports = handleDeleteFolder;
