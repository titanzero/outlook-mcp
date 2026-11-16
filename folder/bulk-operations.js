/**
 * Folder bulk operations tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Copy emails to another folder (instead of moving)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCopyEmails(args) {
  const { emailIds, targetFolder, sourceFolder = 'inbox' } = args;
  
  if (!emailIds) {
    return makeErrorResponse('Email IDs are required (comma-separated).');
  }
  if (!targetFolder) {
    return makeErrorResponse('Target folder name is required.');
  }
  
  const emailIdList = emailIds.split(',').map(id => id.trim()).filter(Boolean);
  if (emailIdList.length === 0) {
    return makeErrorResponse('At least one email ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const results = [];
    const errors = [];
    
    for (const emailId of emailIdList) {
      try {
        await client.api(`me/messages/${emailId}/copy`).post({
          destinationId: targetFolder,
        });
        results.push({ emailId, status: 'success' });
      } catch (e) {
        errors.push({ emailId, error: e.message });
      }
    }
    
    const structured = {
      copied: results.length,
      failed: errors.length,
      results,
      errors,
    };
    
    const textFallback = [
      `Copy Emails Results`,
      `Target Folder: ${targetFolder}`,
      `Source Folder: ${sourceFolder}`,
      `Successful: ${results.length}`,
      `Failed: ${errors.length}`,
      '',
      results.length > 0 ? 'Copied:' : '',
      ...results.map(r => `  ✓ ${r.emailId}`),
      errors.length > 0 ? 'Failed:' : '',
      ...errors.map(e => `  ✗ ${e.emailId}: ${e.error}`),
    ].filter(Boolean).join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error copying emails: ${error.message}`);
  }
}

/**
 * Bulk delete emails
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleBulkDeleteEmails(args) {
  const { emailIds, sourceFolder = 'inbox', permanent = false } = args;
  
  if (!emailIds) {
    return makeErrorResponse('Email IDs are required (comma-separated).');
  }
  
  const emailIdList = emailIds.split(',').map(id => id.trim()).filter(Boolean);
  if (emailIdList.length === 0) {
    return makeErrorResponse('At least one email ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const results = [];
    const errors = [];
    
    for (const emailId of emailIdList) {
      try {
        if (permanent) {
          await client.api(`me/messages/${emailId}`).delete();
        } else {
          // Move to Deleted Items
          await client.api(`me/messages/${emailId}/move`).post({
            destinationId: 'deleteditems',
          });
        }
        results.push({ emailId, status: 'success' });
      } catch (e) {
        errors.push({ emailId, error: e.message });
      }
    }
    
    const action = permanent ? 'Permanently deleted' : 'Moved to Deleted Items';
    
    const structured = {
      processed: results.length,
      failed: errors.length,
      action,
      results,
      errors,
    };
    
    const textFallback = [
      `${action} Results`,
      `Source Folder: ${sourceFolder}`,
      `Successful: ${results.length}`,
      `Failed: ${errors.length}`,
      '',
      results.length > 0 ? 'Success:' : '',
      ...results.map(r => `  ✓ ${r.emailId}`),
      errors.length > 0 ? 'Failed:' : '',
      ...errors.map(e => `  ✗ ${e.emailId}: ${e.error}`),
    ].filter(Boolean).join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error bulk deleting emails: ${error.message}`);
  }
}

/**
 * Bulk mark emails as read/unread
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleBulkMarkRead(args) {
  const { emailIds, isRead = true, sourceFolder = 'inbox' } = args;
  
  if (!emailIds) {
    return makeErrorResponse('Email IDs are required (comma-separated).');
  }
  
  const emailIdList = emailIds.split(',').map(id => id.trim()).filter(Boolean);
  if (emailIdList.length === 0) {
    return makeErrorResponse('At least one email ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const results = [];
    const errors = [];
    
    for (const emailId of emailIdList) {
      try {
        await client.api(`me/messages/${emailId}`).patch({ isRead });
        results.push({ emailId, status: 'success' });
      } catch (e) {
        errors.push({ emailId, error: e.message });
      }
    }
    
    const action = isRead ? 'Marked as read' : 'Marked as unread';
    
    const structured = {
      processed: results.length,
      failed: errors.length,
      action,
      results,
      errors,
    };
    
    const textFallback = [
      `${action} Results`,
      `Source Folder: ${sourceFolder}`,
      `Successful: ${results.length}`,
      `Failed: errors.length`,
      '',
      results.length > 0 ? 'Success:' : '',
      ...results.map(r => `  ✓ ${r.emailId}`),
      errors.length > 0 ? 'Failed:' : '',
      ...errors.map(e => `  ✗ ${e.emailId}: ${e.error}`),
    ].filter(Boolean).join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error bulk marking emails: ${error.message}`);
  }
}

/**
 * Bulk flag emails
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleBulkFlagEmails(args) {
  const { emailIds, flagStatus, sourceFolder = 'inbox' } = args;
  
  if (!emailIds) {
    return makeErrorResponse('Email IDs are required (comma-separated).');
  }
  if (!flagStatus) {
    return makeErrorResponse('Flag status is required: flagged, completed, or cleared.');
  }
  
  const validFlags = ['flagged', 'completed', 'cleared'];
  if (!validFlags.includes(flagStatus)) {
    return makeErrorResponse(`Invalid flag status. Must be one of: ${validFlags.join(', ')}`);
  }
  
  const emailIdList = emailIds.split(',').map(id => id.trim()).filter(Boolean);
  if (emailIdList.length === 0) {
    return makeErrorResponse('At least one email ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const results = [];
    const errors = [];
    
    for (const emailId of emailIdList) {
      try {
        if (flagStatus === 'cleared') {
          await client.api(`me/messages/${emailId}`).patch({
            flag: { flagStatus: 'notFlagged' },
          });
        } else {
          await client.api(`me/messages/${emailId}`).patch({
            flag: { flagStatus, completedDateTime: flagStatus === 'completed' ? { dateTime: new Date().toISOString(), timeZone: 'UTC' } : null },
          });
        }
        results.push({ emailId, status: 'success' });
      } catch (e) {
        errors.push({ emailId, error: e.message });
      }
    }
    
    const structured = {
      processed: results.length,
      failed: errors.length,
      flagStatus,
      results,
      errors,
    };
    
    const textFallback = [
      `Bulk Flag Emails (${flagStatus})`,
      `Source Folder: ${sourceFolder}`,
      `Successful: ${results.length}`,
      `Failed: ${errors.length}`,
      '',
      results.length > 0 ? 'Success:' : '',
      ...results.map(r => `  ✓ ${r.emailId}`),
      errors.length > 0 ? 'Failed:' : '',
      ...errors.map(e => `  ✗ ${e.emailId}: ${e.error}`),
    ].filter(Boolean).join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error bulk flagging emails: ${error.message}`);
  }
}

/**
 * Bulk categorize emails
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleBulkCategorizeEmails(args) {
  const { emailIds, categories, sourceFolder = 'inbox' } = args;
  
  if (!emailIds) {
    return makeErrorResponse('Email IDs are required (comma-separated).')
  }
  if (!categories) {
    return makeErrorResponse('Categories are required (comma-separated, empty to clear).');
  }
  
  const emailIdList = emailIds.split(',').map(id => id.trim()).filter(Boolean);
  if (emailIdList.length === 0) {
    return makeErrorResponse('At least one email ID is required.');
  }
  
  const categoryList = categories === '' ? [] : categories.split(',').map(c => c.trim()).filter(Boolean);
  
  try {
    const client = await getGraphClient();
    
    const results = [];
    const errors = [];
    
    for (const emailId of emailIdList) {
      try {
        await client.api(`me/messages/${emailId}`).patch({ categories: categoryList });
        results.push({ emailId, status: 'success' });
      } catch (e) {
        errors.push({ emailId, error: e.message });
      }
    }
    
    const structured = {
      processed: results.length,
      failed: errors.length,
      categories: categoryList,
      results,
      errors,
    };
    
    const textFallback = [
      `Bulk Categorize Emails`,
      `Categories: ${categoryList.length > 0 ? categoryList.join(', ') : '(cleared)'}`,
      `Source Folder: ${sourceFolder}`,
      `Successful: ${results.length}`,
      `Failed: ${errors.length}`,
      '',
      results.length > 0 ? 'Success:' : '',
      ...results.map(r => `  ✓ ${r.emailId}`),
      errors.length > 0 ? 'Failed:' : '',
      ...errors.map(e => `  ✗ ${e.emailId}: ${e.error}`),
    ].filter(Boolean).join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error bulk categorizing emails: ${error.message}`);
  }
}

module.exports = {
  handleCopyEmails,
  handleBulkDeleteEmails,
  handleBulkMarkRead,
  handleBulkFlagEmails,
  handleBulkCategorizeEmails,
};