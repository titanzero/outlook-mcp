/**
 * Folder management module for Outlook MCP server
 */
const handleListFolders = require('./list');
const handleCreateFolder = require('./create');
const handleMoveEmails = require('./move');
const handleRenameFolder = require('./rename');
const { z } = require('zod');

/**
 * Register folder tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerFolderTools(server) {
  server.tool(
    "list-folders",
    "Lists mail folders in your Outlook account",
    {
      includeItemCounts: z.boolean().optional().describe("Include counts of total and unread items"),
      includeChildren: z.boolean().optional().describe("Include child folders in hierarchy")
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleListFolders
  );

  server.tool(
    "create-folder",
    "Creates a new mail folder",
    {
      name: z.string().describe("Name of the folder to create"),
      parentFolder: z.string().optional().describe("Optional parent folder name (default is root)")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false
      }
    },
    handleCreateFolder
  );

  server.tool(
    "move-emails",
    "Moves emails from one folder to another",
    {
      emailIds: z.string().describe("Comma-separated list of email IDs to move"),
      targetFolder: z.string().describe("Name of the folder to move emails to"),
      sourceFolder: z.string().optional().describe("Optional name of the source folder (default is inbox)")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false
      }
    },
    handleMoveEmails
  );

  server.tool(
    "rename-folder",
    "Renames an existing mail folder",
    {
      name: z.string().describe("Current name of the folder to rename"),
      newName: z.string().describe("New name for the folder")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false
      }
    },
    handleRenameFolder
  );
}

module.exports = {
  registerFolderTools,
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails,
  handleRenameFolder
};