/**
 * Move emails functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { getFolderIdByName, invalidateFolderCache } = require('../email/folder-utils');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Move emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveEmails(args) {
  const emailIds = args.emailIds || '';
  const targetFolder = args.targetFolder || '';
  const sourceFolder = args.sourceFolder || '';
  
  if (!emailIds) {
    return makeErrorResponse('Email IDs are required. Please provide a comma-separated list of email IDs to move.');
  }
  
  if (!targetFolder) {
    return makeErrorResponse('Target folder name is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Parse email IDs
    const ids = emailIds.split(',').map(id => id.trim()).filter(id => id);
    
    if (ids.length === 0) {
      return makeErrorResponse('No valid email IDs provided.');
    }
    
    // Move emails
    const result = await moveEmailsToFolder(client, ids, targetFolder, sourceFolder);
    
    return makeResponse(result.message);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error moving emails: ${error.message}`);
  }
}

/**
 * Move emails to a folder
 * @param {object} client - Microsoft Graph SDK client
 * @param {Array<string>} emailIds - Array of email IDs to move
 * @param {string} targetFolderName - Name of the target folder
 * @param {string} sourceFolderName - Name of the source folder (optional)
 * @returns {Promise<object>} - Result object with status and message
 */
async function moveEmailsToFolder(client, emailIds, targetFolderName, sourceFolderName) {
  try {
    // Get the target folder ID
    const targetFolderId = await getFolderIdByName(targetFolderName);
    if (!targetFolderId) {
      return {
        success: false,
        message: `Target folder "${targetFolderName}" not found. Use 'list-folders' to see available folders. For nested folders, use a path like 'ParentFolder/SubFolder'.`
      };
    }
    
    // Track successful and failed moves
    const results = {
      successful: [],
      failed: []
    };
    
    // Process each email one by one to handle errors independently
    for (const emailId of emailIds) {
      try {
        // Move the email
        await client.api(`me/messages/${emailId}/move`).post({
          destinationId: targetFolderId
        });
        
        results.successful.push(emailId);
      } catch (error) {
        console.error(`Error moving email ${emailId}: ${error.message}`);
        results.failed.push({
          id: emailId,
          error: error.message
        });
      }
    }
    
    // Generate result message
    let message = '';
    
    if (results.successful.length > 0) {
      message += `Successfully moved ${results.successful.length} email(s) to "${targetFolderName}".`;
    }
    
    if (results.failed.length > 0) {
      if (message) message += '\n\n';
      message += `Failed to move ${results.failed.length} email(s). Errors:`;
      
      // Show first few errors with details
      const maxErrors = Math.min(results.failed.length, 3);
      for (let i = 0; i < maxErrors; i++) {
        const failure = results.failed[i];
        message += `\n- Email ${i+1}: ${failure.error}`;
      }
      
      // If there are more errors, just mention the count
      if (results.failed.length > maxErrors) {
        message += `\n...and ${results.failed.length - maxErrors} more.`;
      }
    }
    
    // Invalidate folder cache – item counts have changed
    if (results.successful.length > 0) {
      invalidateFolderCache();
    }

    return {
      success: results.successful.length > 0,
      message,
      results
    };
  } catch (error) {
    console.error(`Error in moveEmailsToFolder: ${error.message}`);
    throw error;
  }
}

module.exports = handleMoveEmails;
