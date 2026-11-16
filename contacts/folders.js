/**
 * Contact folders tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List contact folders
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListContactFolders(_args) {
  try {
    const client = await getGraphClient();
    
    const response = await graphGetPaginated(client, 'me/contactFolders', {
      $select: 'id,displayName,parentFolderId,contactCount',
    });
    
    const folders = response.value || [];
    
    const structured = folders.map(f => ({
      id: f.id,
      name: f.displayName,
      parentFolderId: f.parentFolderId,
      contactCount: f.contactCount || 0,
    }));
    
    const textFallback = [
      `Contact Folders (${folders.length}):`,
      '',
      folders.length === 0
        ? '  No contact folders found.'
        : folders.map((f, i) => 
            `${i + 1}. ${f.displayName} (${f.contactCount || 0} contacts)${f.parentFolderId ? ` [parent: ${f.parentFolderId}]` : ' [root]'} [id: ${f.id}]`
          ).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ folders: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing contact folders: ${error.message}`);
  }
}

/**
 * Get contact folder details with contacts
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetContactFolder(args) {
  const { folderId, count = 50 } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required (from list-contact-folders).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get folder info
    const folder = await client.api(`me/contactFolders/${folderId}`).get();
    
    // Get contacts in this folder
    const response = await graphGetPaginated(client, `me/contactFolders/${folderId}/contacts`, {
      $select: 'id,displayName,emailAddresses,mobilePhone,jobTitle,companyName,businessPhones,homePhones',
      $top: Math.min(count, 1000),
    });
    
    const contacts = response.value || [];
    
    const structured = {
      id: folder.id,
      name: folder.displayName,
      parentFolderId: folder.parentFolderId,
      contactCount: folder.contactCount || 0,
      contacts: contacts.map(c => ({
        id: c.id,
        displayName: c.displayName,
        emails: c.emailAddresses?.map(e => ({ address: e.address, name: e.name })) || [],
        mobilePhone: c.mobilePhone,
        businessPhones: c.businessPhones || [],
        homePhones: c.homePhones || [],
        jobTitle: c.jobTitle,
        companyName: c.companyName,
      })),
    };
    
    const textFallback = [
      `Contact Folder: ${folder.displayName}`,
      `ID: ${folder.id}`,
      `Contacts: ${contacts.length} (of ${folder.contactCount || 0} total)`,
      '',
      contacts.length === 0
        ? '  No contacts in this folder.'
        : contacts.map((c, i) => {
            const emails = c.emailAddresses?.map(e => e.address).join(', ') || 'no email';
            const phones = [];
            if (c.mobilePhone) phones.push(`Mobile: ${c.mobilePhone}`);
            if (c.businessPhones?.length) phones.push(`Business: ${c.businessPhones.join(', ')}`);
            return `${i + 1}. ${c.displayName} (${emails})${phones.length ? ` - ${phones.join(', ')}` : ''}${c.jobTitle ? `, ${c.jobTitle}` : ''}${c.companyName ? ` @ ${c.companyName}` : ''}`;
          }).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting contact folder: ${error.message}`);
  }
}

/**
 * Create a new contact folder
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateContactFolder(args) {
  const { name, parentFolderId } = args;
  
  if (!name) {
    return makeErrorResponse('Folder name is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const folder = { displayName: name };
    if (parentFolderId) {
      folder.parentFolderId = parentFolderId;
    }
    
    const created = await client.api('me/contactFolders').post(folder);
    
    return makeResponse(`Contact folder "${name}" created successfully.\nID: ${created.id}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating contact folder: ${error.message}`);
  }
}

/**
 * Update contact folder
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateContactFolder(args) {
  const { folderId, name, parentFolderId } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required.');
  }
  if (!name && !parentFolderId) {
    return makeErrorResponse('At least one of name or parentFolderId is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const patch = {};
    if (name) patch.displayName = name;
    if (parentFolderId) patch.parentFolderId = parentFolderId;
    
    await client.api(`me/contactFolders/${folderId}`).patch(patch);
    
    return makeResponse('Contact folder updated successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating contact folder: ${error.message}`);
  }
}

/**
 * Delete contact folder
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteContactFolder(args) {
  const { folderId } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Check if folder has contacts
    const folder = await client.api(`me/contactFolders/${folderId}`).select('contactCount').get();
    if (folder.contactCount > 0) {
      return makeErrorResponse(`Cannot delete folder with ${folder.contactCount} contacts. Move or delete contacts first.`);
    }
    
    await client.api(`me/contactFolders/${folderId}`).delete();
    
    return makeResponse('Contact folder deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting contact folder: ${error.message}`);
  }
}

/**
 * Move contact to folder
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveContactToFolder(args) {
  const { contactId, folderId } = args;
  
  if (!contactId) {
    return makeErrorResponse('Contact ID is required (from list-contacts).');
  }
  if (!folderId) {
    return makeErrorResponse('Destination folder ID is required (from list-contact-folders).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Move contact by updating its parentFolderId
    // Note: Graph API doesn't have a direct move operation for contacts
    // We need to recreate the contact in the target folder and delete the original
    // Or use the /move endpoint if available
    
    // Try the move action first
    try {
      await client.api(`me/contacts/${contactId}/move`).post({
        destinationId: folderId,
      });
    } catch (e) {
      // Fallback: copy and delete
      const contact = await client.api(`me/contacts/${contactId}`).get();
      
      // Create in new folder
      const { '@odata.context': _, '@odata.etag': __, id: ___, createdDateTime: ____, lastModifiedDateTime: _____, ...contactData } = contact;
      contactData.parentFolderId = folderId;
      
      await client.api(`me/contactFolders/${folderId}/contacts`).post(contactData);
      
      // Delete original
      await client.api(`me/contacts/${contactId}`).delete();
    }
    
    return makeResponse('Contact moved to folder successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error moving contact: ${error.message}`);
  }
}

module.exports = {
  handleListContactFolders,
  handleGetContactFolder,
  handleCreateContactFolder,
  handleUpdateContactFolder,
  handleDeleteContactFolder,
  handleMoveContactToFolder,
};