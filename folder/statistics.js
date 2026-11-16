/**
 * Folder statistics and search tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Get folder statistics (counts, sizes, etc.)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetFolderStatistics(args) {
  const { folderId } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required (from list-folders).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get folder info
    const folder = await client.api(`me/mailFolders/${folderId}`).select('id,displayName,totalItemCount,unreadItemCount,childFolderCount').get();
    
    // Get message details for more stats
    const messagesResponse = await graphGetPaginated(client, `me/mailFolders/${folderId}/messages`, {
      $select: 'id,size,hasAttachments,importance,receivedDateTime,isRead,categories,flag',
      $top: 1000,
    });
    
    const messages = messagesResponse.value || [];
    
    // Calculate statistics
    let totalSize = 0;
    let withAttachments = 0;
    let highImportance = 0;
    let lowImportance = 0;
    let flagged = 0;
    const categories = new Map();
    const dateRange = { earliest: null, latest: null };
    
    messages.forEach(m => {
      totalSize += m.size || 0;
      if (m.hasAttachments) withAttachments++;
      if (m.importance === 'high') highImportance++;
      if (m.importance === 'low') lowImportance++;
      if (m.flag?.flagStatus === 'flagged') flagged++;
      m.categories?.forEach(c => categories.set(c, (categories.get(c) || 0) + 1));
      if (m.receivedDateTime) {
        const date = new Date(m.receivedDateTime);
        if (!dateRange.earliest || date < dateRange.earliest) dateRange.earliest = date;
        if (!dateRange.latest || date > dateRange.latest) dateRange.latest = date;
      }
    });
    
    const structured = {
      folderId: folder.id,
      folderName: folder.displayName,
      totalItems: folder.totalItemCount || 0,
      unreadItems: folder.unreadItemCount || 0,
      childFolders: folder.childFolderCount || 0,
      scannedMessages: messages.length,
      totalSizeBytes: totalSize,
      totalSizeMB: (totalSize / (1024 * 1024)).toFixed(2),
      withAttachments,
      highImportance,
      lowImportance,
      normalImportance: messages.length - highImportance - lowImportance,
      flagged,
      categories: Array.from(categories.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 20)
        .map(([name, count]) => ({ name, count })),
      dateRange: {
        earliest: dateRange.earliest?.toISOString(),
        latest: dateRange.latest?.toISOString(),
      },
    };
    
    const formatBytes = (bytes) => {
      if (bytes < 1024) return `${bytes} B`;
      if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
      if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
      return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
    };
    
    const textFallback = [
      `Folder Statistics: ${folder.displayName}`,
      `Folder ID: ${folder.id}`,
      '',
      `Total Items: ${folder.totalItemCount || 0}`,
      `Unread Items: ${folder.unreadItemCount || 0}`,
      `Child Folders: ${folder.childFolderCount || 0}`,
      `Scanned Messages: ${messages.length}`,
      `Total Size: ${formatBytes(totalSize)}`,
      '',
      `With Attachments: ${withAttachments}`,
      `High Importance: ${highImportance}`,
      `Normal Importance: ${messages.length - highImportance - lowImportance}`,
      `Low Importance: ${lowImportance}`,
      `Flagged: ${flagged}`,
      '',
      `Date Range: ${dateRange.earliest ? dateRange.earliest.toISOString().split('T')[0] : 'N/A'} to ${dateRange.latest ? dateRange.latest.toISOString().split('T')[0] : 'N/A'}`,
      '',
      `Top Categories:`,
      structured.categories.length > 0
        ? structured.categories.map((c, i) => `  ${i + 1}. ${c.name}: ${c.count}`).join('\n')
        : '  None',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting folder statistics: ${error.message}`);
  }
}

/**
 * Get all folders statistics summary
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetAllFoldersStatistics(_args) {
  try {
    const client = await getGraphClient();
    
    const response = await client.api('me/mailFolders').select('id,displayName,totalItemCount,unreadItemCount,childFolderCount,wellKnownName').get();
    const folders = response.value || [];
    
    let totalItems = 0;
    let totalUnread = 0;
    let totalChildFolders = 0;
    
    const folderStats = folders.map(f => {
      totalItems += f.totalItemCount || 0;
      totalUnread += f.unreadItemCount || 0;
      totalChildFolders += f.childFolderCount || 0;
      return {
        id: f.id,
        name: f.displayName,
        wellKnownName: f.wellKnownName,
        totalItems: f.totalItemCount || 0,
        unreadItems: f.unreadItemCount || 0,
        childFolders: f.childFolderCount || 0,
      };
    });
    
    // Sort by total items descending
    folderStats.sort((a, b) => b.totalItems - a.totalItems);
    
    const structured = {
      totalFolders: folders.length,
      totalItems,
      totalUnread,
      totalChildFolders,
      folders: folderStats,
    };
    
    const textFallback = [
      `All Folders Statistics Summary`,
      `Total Folders: ${folders.length}`,
      `Total Items: ${totalItems}`,
      `Total Unread: ${totalUnread}`,
      `Total Child Folders: ${totalChildFolders}`,
      '',
      `Top Folders by Item Count:`,
      folderStats.slice(0, 20).map((f, i) => {
        const wellKnown = f.wellKnownName ? ` [${f.wellKnownName}]` : '';
        return `${i + 1}. ${f.name}${wellKnown}: ${f.totalItems} items (${f.unreadItems} unread)`;
      }).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting all folders statistics: ${error.message}`);
  }
}

/**
 * Search within a specific folder
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchInFolder(args) {
  const { folderId, query, count = 50 } = args;
  
  if (!folderId) {
    return makeErrorResponse('Folder ID is required (from list-folders).');
  }
  if (!query) {
    return makeErrorResponse('Search query is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await graphGetPaginated(client, `me/mailFolders/${folderId}/messages`, {
      $search: `"${query}"`,
      $select: 'id,subject,from,toRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead,categories',
      $orderby: 'receivedDateTime desc',
      $top: Math.min(count, 500),
    });
    
    const emails = response.value || [];
    
    const structured = {
      folderId,
      query,
      resultCount: emails.length,
      emails: emails.map(e => {
        const sender = e.from ? e.from.emailAddress : { name: 'Unknown', address: 'unknown' };
        const toRecipients = e.toRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        return {
          id: e.id,
          subject: e.subject,
          from: { name: sender.name, address: sender.address },
          to: toRecipients,
          receivedDateTime: e.receivedDateTime,
          bodyPreview: e.bodyPreview,
          hasAttachments: e.hasAttachments,
          importance: e.importance,
          isRead: e.isRead,
          categories: e.categories || [],
        };
      }),
    };
    
    const textFallback = [
      `Search in Folder`,
      `Folder ID: ${folderId}`,
      `Query: ${query}`,
      `Results: ${emails.length}`,
      '',
      emails.length === 0
        ? '  No emails found matching the query.'
        : emails.map((e, i) => {
            const sender = e.from ? e.from.emailAddress : { name: 'Unknown', address: 'unknown' };
            const readStatus = e.isRead ? '' : '[UNREAD] ';
            const cats = e.categories?.length ? ` [${e.categories.join(', ')}]` : '';
            return `${i + 1}. ${readStatus}${e.receivedDateTime} - ${sender.name} <${sender.address}>${cats}\n   Subject: ${e.subject}\n   Preview: ${e.bodyPreview || '(empty)'}\n   ID: ${e.id}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error searching in folder: ${error.message}`);
  }
}

module.exports = {
  handleGetFolderStatistics,
  handleGetAllFoldersStatistics,
  handleSearchInFolder,
};