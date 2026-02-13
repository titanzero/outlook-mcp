/**
 * Create folder functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { getFolderIdByName, invalidateFolderCache } = require('../email/folder-utils');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Create folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateFolder(args) {
  const folderName = args.name;
  const parentFolder = args.parentFolder || '';
  
  if (!folderName) {
    return makeErrorResponse('Folder name is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Create folder with appropriate parent
    const result = await createMailFolder(client, folderName, parentFolder);
    
    return makeResponse(result.message);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error creating folder: ${error.message}`);
  }
}

/**
 * Create a new mail folder
 * @param {object} client - Microsoft Graph SDK client
 * @param {string} folderName - Name of the folder to create
 * @param {string} parentFolderName - Name of the parent folder (optional)
 * @returns {Promise<object>} - Result object with status and message
 */
async function createMailFolder(client, folderName, parentFolderName) {
  try {
    // Check if a folder with this name already exists
    const existingFolder = await getFolderIdByName(folderName);
    if (existingFolder) {
      return {
        success: false,
        message: `A folder named "${folderName}" already exists.`
      };
    }
    
    // If parent folder specified, find its ID
    let endpoint = 'me/mailFolders';
    if (parentFolderName) {
      const parentId = await getFolderIdByName(parentFolderName);
      if (!parentId) {
        return {
          success: false,
          message: `Parent folder "${parentFolderName}" not found. Use 'list-folders' to see available folders. For nested parents, use a path like 'ParentFolder/SubFolder'.`
        };
      }
      
      endpoint = `me/mailFolders/${parentId}/childFolders`;
    }
    
    // Create the folder
    const folderData = {
      displayName: folderName
    };
    
    const response = await client.api(endpoint).post(folderData);

    // Bust the folder cache so subsequent look-ups see the new folder
    invalidateFolderCache();

    if (response && response.id) {
      const locationInfo = parentFolderName 
        ? `inside "${parentFolderName}"` 
        : "at the root level";
        
      return {
        success: true,
        message: `Successfully created folder "${folderName}" ${locationInfo}.`,
        folderId: response.id
      };
    } else {
      return {
        success: false,
        message: "Failed to create folder. The server didn't return a folder ID."
      };
    }
  } catch (error) {
    console.error(`Error creating folder "${folderName}": ${error.message}`);
    throw error;
  }
}

module.exports = handleCreateFolder;
