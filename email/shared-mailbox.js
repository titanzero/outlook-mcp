/**
 * Shared mailbox access tool for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List shared mailboxes the user has access to
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListSharedMailboxes(_args) {
  try {
    const client = await getGraphClient();
    
    // Get user's mailbox settings to find shared mailboxes
    // Also check for delegate access
    const response = await client
      .api('me/mailFolders')
      .select('id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,wellKnownName')
      .get();
    
    const folders = response.value || [];
    
    // Also try to get shared mailboxes via /users endpoint if available
    // This requires Mail.Read.Shared scope
    let sharedMailboxes = [];
    try {
      // Get mailboxes user has delegate access to
      const delegatesResponse = await client.api('me/mailboxSettings/delegates').get();
      if (delegatesResponse.value) {
        sharedMailboxes = delegatesResponse.value;
      }
    } catch (e) {
      // Delegates endpoint might not be available, continue with folder approach
    }
    
    // Get well-known folders that might be shared
    const wellKnownFolders = folders.filter(f => f.wellKnownName);
    
    const structured = {
      ownFolders: folders.map(f => ({
        id: f.id,
        name: f.displayName,
        path: f.displayName,
        parentFolderId: f.parentFolderId,
        totalItems: f.totalItemCount || 0,
        unreadItems: f.unreadItemCount || 0,
        childFolderCount: f.childFolderCount || 0,
        wellKnownName: f.wellKnownName || null,
      })),
      sharedMailboxes: sharedMailboxes.map(m => ({
        emailAddress: m.emailAddress?.address,
        displayName: m.displayName,
        permissions: m.permissions,
      })),
    };
    
    const textFallback = [
      `Own Mail Folders (${folders.length}):`,
      folders.map(f => `- ${f.displayName} (${f.totalItemCount || 0} items, ${f.unreadItemCount || 0} unread) [${f.wellKnownName || 'custom'}]`).join('\n'),
      '',
      `Shared Mailboxes/Delegates (${sharedMailboxes.length}):`,
      sharedMailboxes.length > 0 
        ? sharedMailboxes.map(m => `- ${m.displayName} <${m.emailAddress?.address}> (${m.permissions?.join(', ') || 'unknown'})`).join('\n')
        : '  No shared mailboxes found or delegates endpoint not available',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing shared mailboxes: ${error.message}`);
  }
}

/**
 * Access a shared mailbox by email address
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAccessSharedMailbox(args) {
  const { emailAddress, folder = 'inbox', count = 20 } = args;
  
  if (!emailAddress) {
    return makeErrorResponse('Email address of the shared mailbox is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Try to access the shared mailbox
    const endpoint = `users/${encodeURIComponent(emailAddress)}/mailFolders/${folder}/messages`;
    
    const response = await client
      .api(endpoint)
      .top(Math.min(count, 100))
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead')
      .orderby('receivedDateTime desc')
      .get();
    
    const emails = response.value || [];
    
    const structured = {
      mailbox: emailAddress,
      folder,
      emails: emails.map(e => {
        const sender = e.from ? e.from.emailAddress : { name: 'Unknown', address: 'unknown' };
        const toRecipients = e.toRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        const ccRecipients = e.ccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        
        return {
          id: e.id,
          subject: e.subject,
          from: { name: sender.name, address: sender.address },
          to: toRecipients,
          cc: ccRecipients,
          receivedDateTime: e.receivedDateTime,
          bodyPreview: e.bodyPreview,
          hasAttachments: e.hasAttachments,
          importance: e.importance,
          isRead: e.isRead,
        };
      }),
    };
    
    const textFallback = [
      `Shared Mailbox: ${emailAddress}`,
      `Folder: ${folder}`,
      `Emails (${emails.length}):`,
      '',
      emails.length === 0 
        ? '  No emails found.'
        : emails.map((e, i) => {
            const sender = e.from ? e.from.emailAddress : { name: 'Unknown', address: 'unknown' };
            const readStatus = e.isRead ? '' : '[UNREAD] ';
            return `${i + 1}. ${readStatus}${e.receivedDateTime} - ${sender.name} <${sender.address}>\n   Subject: ${e.subject}\n   ID: ${e.id}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 403 || error.statusCode === 404) {
      return makeErrorResponse(`Cannot access shared mailbox "${emailAddress}". Ensure you have delegate permissions and the Mail.Read.Shared scope is granted.`);
    }
    return makeErrorResponse(`Error accessing shared mailbox: ${error.message}`);
  }
}

/**
 * List folders in a shared mailbox
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListSharedMailboxFolders(args) {
  const { emailAddress } = args;
  
  if (!emailAddress) {
    return makeErrorResponse('Email address of the shared mailbox is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await client
      .api(`users/${encodeURIComponent(emailAddress)}/mailFolders`)
      .select('id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,wellKnownName')
      .get();
    
    const folders = response.value || [];
    
    const structured = {
      mailbox: emailAddress,
      folders: folders.map(f => ({
        id: f.id,
        name: f.displayName,
        parentFolderId: f.parentFolderId,
        totalItems: f.totalItemCount || 0,
        unreadItems: f.unreadItemCount || 0,
        childFolderCount: f.childFolderCount || 0,
        wellKnownName: f.wellKnownName || null,
      })),
    };
    
    const textFallback = [
      `Folders in ${emailAddress} (${folders.length}):`,
      '',
      folders.length === 0 
        ? '  No folders found.'
        : folders.map(f => `- ${f.displayName} (${f.totalItemCount || 0} items, ${f.unreadItemCount || 0} unread) [${f.wellKnownName || 'custom'}]`).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 403 || error.statusCode === 404) {
      return makeErrorResponse(`Cannot access shared mailbox "${emailAddress}". Ensure you have delegate permissions.`);
    }
    return makeErrorResponse(`Error listing shared mailbox folders: ${error.message}`);
  }
}

module.exports = {
  handleListSharedMailboxes,
  handleAccessSharedMailbox,
  handleListSharedMailboxFolders,
};