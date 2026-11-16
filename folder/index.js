/**
 * Folder management module for Outlook MCP server
 */
const handleListFolders = require('./list');
const handleCreateFolder = require('./create');
const handleMoveEmails = require('./move');
const handleRenameFolder = require('./rename');
const handleDeleteFolder = require('./delete');
const { registerFolderResources } = require('./resources');
const { handleCopyEmails, handleBulkDeleteEmails, handleBulkMarkRead, handleBulkFlagEmails, handleBulkCategorizeEmails } = require('./bulk-operations');
const { handleListFolderPermissions, handleAddFolderPermission, handleUpdateFolderPermission, handleRemoveFolderPermission } = require('./permissions');
const { handleGetFolderStatistics, handleGetAllFoldersStatistics, handleSearchInFolder } = require('./statistics');
const { handleGetFolderTree, handleMoveFolder, handleCreateFolderPath, handleEmptyFolder, handleGetFolderPath } = require('./hierarchy');
const { z } = require('zod');

/**
 * Register folder tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerFolderTools(server) {
  server.tool(
    'list-folders',
    'Lists mail folders in your Outlook account',
    {
      includeItemCounts: z
        .boolean()
        .optional()
        .describe('Include counts of total and unread items'),
      includeChildren: z.boolean().optional().describe('Include child folders in hierarchy'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListFolders
  );

  server.tool(
    'create-folder',
    'Creates a new mail folder',
    {
      name: z.string().describe('Name of the folder to create'),
      parentFolder: z.string().optional().describe('Optional parent folder name (default is root)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateFolder
  );

  server.tool(
    'move-emails',
    'Moves emails from one folder to another',
    {
      emailIds: z.string().describe('Comma-separated list of email IDs to move'),
      targetFolder: z.string().describe('Name of the folder to move emails to'),
      sourceFolder: z
        .string()
        .optional()
        .describe('Optional name of the source folder (default is inbox)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleMoveEmails
  );

  server.tool(
    'rename-folder',
    'Renames an existing mail folder',
    {
      name: z.string().describe('Current name of the folder to rename'),
      newName: z.string().describe('New name for the folder'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleRenameFolder
  );

  server.tool(
    'delete-folder',
    'Deletes a mail folder. Cannot delete well-known folders (Inbox, Drafts, Sent Items, etc.) or non-empty folders.',
    {
      name: z.string().describe('Name of the folder to delete'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteFolder
  );

  // Bulk Email Operations
  server.tool(
    'copy-emails',
    'Copies emails to another folder (instead of moving)',
    {
      emailIds: z.string().describe('Comma-separated list of email IDs to copy'),
      targetFolder: z.string().describe('Name of the folder to copy emails to'),
      sourceFolder: z.string().optional().describe('Optional name of the source folder (default is inbox)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleCopyEmails
  );

  server.tool(
    'bulk-delete-emails',
    'Bulk deletes or moves emails to Deleted Items',
    {
      emailIds: z.string().describe('Comma-separated list of email IDs to delete'),
      sourceFolder: z.string().optional().describe('Optional source folder (default: inbox)'),
      permanent: z.boolean().optional().describe('If true, permanently delete instead of moving to Deleted Items (default: false)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleBulkDeleteEmails
  );

  server.tool(
    'bulk-mark-read',
    'Bulk marks emails as read or unread',
    {
      emailIds: z.string().describe('Comma-separated list of email IDs'),
      isRead: z.boolean().optional().describe('Whether to mark as read (true) or unread (false). Default: true'),
      sourceFolder: z.string().optional().describe('Optional source folder (default: inbox)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleBulkMarkRead
  );

  server.tool(
    'bulk-flag-emails',
    'Bulk flags, completes, or clears flags on emails',
    {
      emailIds: z.string().describe('Comma-separated list of email IDs'),
      flagStatus: z.enum(['flagged', 'completed', 'cleared']).describe('Flag status'),
      sourceFolder: z.string().optional().describe('Optional source folder (default: inbox)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleBulkFlagEmails
  );

  server.tool(
    'bulk-categorize-emails',
    'Bulk adds or removes categories from emails',
    {
      emailIds: z.string().describe('Comma-separated list of email IDs'),
      categories: z.string().describe('Comma-separated list of category names. Empty string clears all categories.'),
      sourceFolder: z.string().optional().describe('Optional source folder (default: inbox)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleBulkCategorizeEmails
  );

  // Folder Permissions (requires Mail.ReadWrite.Shared)
  server.tool(
    'list-folder-permissions',
    'Lists permissions for a mail folder (requires Mail.ReadWrite.Shared scope)',
    {
      folderId: z.string().describe('Folder ID (from list-folders)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListFolderPermissions
  );

  server.tool(
    'add-folder-permission',
    'Adds permission for a user to access a folder (requires Mail.ReadWrite.Shared scope)',
    {
      folderId: z.string().describe('Folder ID (from list-folders)'),
      emailAddress: z.string().describe('Email address of the user'),
      role: z.enum(['none', 'owner', 'publishingEditor', 'editor', 'publishingAuthor', 'author', 'noneditingAuthor', 'reviewer', 'contributor']).optional().describe('Permission role (default: reviewer)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleAddFolderPermission
  );

  server.tool(
    'update-folder-permission',
    'Updates a folder permission (requires Mail.ReadWrite.Shared scope)',
    {
      folderId: z.string().describe('Folder ID'),
      permissionId: z.string().describe('Permission ID (from list-folder-permissions)'),
      role: z.enum(['none', 'owner', 'publishingEditor', 'editor', 'publishingAuthor', 'author', 'noneditingAuthor', 'reviewer', 'contributor']).describe('New permission role'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateFolderPermission
  );

  server.tool(
    'remove-folder-permission',
    'Removes a folder permission (requires Mail.ReadWrite.Shared scope)',
    {
      folderId: z.string().describe('Folder ID'),
      permissionId: z.string().describe('Permission ID (from list-folder-permissions)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleRemoveFolderPermission
  );

  // Folder Statistics & Search
  server.tool(
    'get-folder-statistics',
    'Gets detailed statistics for a folder (counts, sizes, categories, date range)',
    {
      folderId: z.string().describe('Folder ID (from list-folders)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetFolderStatistics
  );

  server.tool(
    'get-all-folders-statistics',
    'Gets statistics summary for all folders',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetAllFoldersStatistics
  );

  server.tool(
    'search-in-folder',
    'Searches emails within a specific folder',
    {
      folderId: z.string().describe('Folder ID (from list-folders)'),
      query: z.string().describe('Search query text'),
      count: z.number().int().positive().max(500).optional().describe('Number of results (default: 50, max: 500)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleSearchInFolder
  );

  // Folder Hierarchy Management
  server.tool(
    'get-folder-tree',
    'Gets the complete folder hierarchy as a tree structure',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetFolderTree
  );

  server.tool(
    'move-folder',
    'Moves a folder to a new parent folder',
    {
      folderId: z.string().describe('Folder ID to move (from list-folders)'),
      newParentFolderId: z.string().describe('New parent folder ID (from list-folders)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleMoveFolder
  );

  server.tool(
    'create-folder-path',
    'Creates a nested folder path (creates intermediate folders as needed)',
    {
      path: z.string().describe('Folder path (e.g., "Projects/ClientA/2024")'),
      parentFolder: z.string().optional().describe('Optional parent folder ID (default: root)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateFolderPath
  );

  server.tool(
    'empty-folder',
    'Empties a folder by moving all emails to Deleted Items (or permanently deleting)',
    {
      folderId: z.string().describe('Folder ID (from list-folders)'),
      permanent: z.boolean().optional().describe('If true, permanently delete emails (default: false, moves to Deleted Items)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleEmptyFolder
  );

  server.tool(
    'get-folder-path',
    'Gets the full path of a folder from root',
    {
      folderId: z.string().describe('Folder ID (from list-folders)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetFolderPath
  );
}

module.exports = {
  registerFolderTools,
  registerFolderResources,
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails,
  handleRenameFolder,
  handleDeleteFolder,
  handleCopyEmails,
  handleBulkDeleteEmails,
  handleBulkMarkRead,
  handleBulkFlagEmails,
  handleBulkCategorizeEmails,
  handleListFolderPermissions,
  handleAddFolderPermission,
  handleUpdateFolderPermission,
  handleRemoveFolderPermission,
  handleGetFolderStatistics,
  handleGetAllFoldersStatistics,
  handleSearchInFolder,
  handleGetFolderTree,
  handleMoveFolder,
  handleCreateFolderPath,
  handleEmptyFolder,
  handleGetFolderPath,
};
