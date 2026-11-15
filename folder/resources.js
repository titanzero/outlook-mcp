/**
 * Folder resources for MCP
 * Exposes mail folders as addressable resources
 */
const { getGraphClient } = require('../utils/graph-client');
const { getAllFolders } = require('../email/folder-utils');
const { isAuthError } = require('../utils/response-helpers');
const { ResourceTemplate } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { getFolderCompletions } = require('../completion');

/**
 * Register folder resources with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerFolderResources(server) {
  // Resource template for individual folders by name/path with completion
  const folderTemplate = new ResourceTemplate('outlook://folders/{folderName}', {
    complete: {
      folderName: getFolderCompletions,
    },
  });

  server.resource(
    'outlook-folder',
    folderTemplate,
    {
      title: 'Mail Folder',
      description: "Access a mail folder by name or path (e.g., 'Inbox', 'Inbox/Projects/ClientA')",
    },
    async (uri, { folderName }) => {
      try {
        const client = await getGraphClient();
        const folders = await getAllFolders();

        // Find folder by name or path
        const folder = folders.find(f => f.displayName === folderName || f.path === folderName);

        if (!folder) {
          throw new Error(`Folder "${folderName}" not found`);
        }

        const structured = {
          id: folder.id,
          name: folder.displayName,
          path: folder.path,
          parentFolderId: folder.parentFolderId,
          totalItems: folder.totalItemCount || 0,
          unreadItems: folder.unreadItemCount || 0,
          childFolderCount: folder.childFolderCount || 0,
          wellKnownName: folder.wellKnownName || null,
        };

        const textFallback = [
          `Folder: ${folder.displayName}`,
          `Path: ${folder.path || folder.displayName}`,
          `ID: ${folder.id}`,
          `Items: ${folder.totalItemCount || 0} (${folder.unreadItemCount || 0} unread)`,
          `Subfolders: ${folder.childFolderCount || 0}`,
        ].join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading folder: ${error.message}`);
      }
    }
  );

  // Resource for folder hierarchy tree
  server.resource(
    'outlook-folder-tree',
    'outlook://folders',
    {
      title: 'Mail Folder Hierarchy',
      description: 'Complete folder hierarchy as a tree structure',
    },
    async uri => {
      try {
        const folders = await getAllFolders();

        // Build hierarchy
        const folderMap = new Map();
        const rootFolders = [];

        folders.forEach(folder => {
          folderMap.set(folder.id, { ...folder, children: [] });
          if (!folder.path || !folder.path.includes('/')) {
            rootFolders.push(folder.id);
          }
        });

        folders.forEach(folder => {
          if (folder.path && folder.path.includes('/') && folder.parentFolderId) {
            const parent = folderMap.get(folder.parentFolderId);
            if (parent) {
              parent.children.push(folder.id);
            } else {
              rootFolders.push(folder.id);
            }
          }
        });

        function buildTree(folderId, level = 0) {
          const folder = folderMap.get(folderId);
          if (!folder) {
            return null;
          }

          const node = {
            id: folder.id,
            name: folder.displayName,
            path: folder.path,
            totalItems: folder.totalItemCount || 0,
            unreadItems: folder.unreadItemCount || 0,
            level,
          };

          if (folder.children.length > 0) {
            node.children = folder.children.map(id => buildTree(id, level + 1)).filter(Boolean);
          }

          return node;
        }

        const tree = rootFolders.map(id => buildTree(id)).filter(Boolean);

        const structured = { folders: tree };

        function formatTree(folders, indent = 0) {
          return folders
            .map(f => {
              const prefix = '  '.repeat(indent);
              const unread = f.unreadItems > 0 ? ` (${f.unreadItems} unread)` : '';
              let line = `${prefix}📁 ${f.name} [${f.totalItems} items${unread}]`;
              if (f.children) {
                line += '\n' + formatTree(f.children, indent + 1);
              }
              return line;
            })
            .join('\n');
        }

        const textFallback = `Mail Folder Hierarchy:\n\n${formatTree(tree)}`;

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading folder tree: ${error.message}`);
      }
    }
  );
}

module.exports = { registerFolderResources };
