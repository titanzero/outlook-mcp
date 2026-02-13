/**
 * Email folder utilities
 * Uses the Microsoft Graph JS SDK for API calls.
 *
 * Includes a TTL-based in-memory cache so repeated folder look-ups
 * (e.g. during list + search in the same conversation turn) don't
 * trigger extra Graph API round-trips.
 */
const { getGraphClient } = require('../utils/graph-client');
const config = require('../config');

// ---------------------------------------------------------------------------
// Folder cache (TTL-based, in-memory)
// ---------------------------------------------------------------------------
let _folderCache = { folders: null, timestamp: 0 };

/**
 * Returns true when the cached folder list is still valid.
 * @returns {boolean}
 */
function isCacheValid() {
  return _folderCache.folders !== null &&
    (Date.now() - _folderCache.timestamp < config.FOLDER_CACHE_TTL_MS);
}

/**
 * Invalidate the folder cache.
 * Call this after any folder-mutating operation (create, move, delete).
 */
function invalidateFolderCache() {
  _folderCache = { folders: null, timestamp: 0 };
}

/**
 * Well-known folder names and their endpoints
 */
const WELL_KNOWN_FOLDERS = {
  'inbox': 'me/mailFolders/inbox/messages',
  'drafts': 'me/mailFolders/drafts/messages',
  'sent': 'me/mailFolders/sentItems/messages',
  'deleted': 'me/mailFolders/deletedItems/messages',
  'junk': 'me/mailFolders/junkemail/messages',
  'archive': 'me/mailFolders/archive/messages'
};

/**
 * Resolve a folder name to its endpoint path
 * @param {string} folderName - Folder name to resolve
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(folderName) {
  // Default to inbox if no folder specified
  if (!folderName) {
    return WELL_KNOWN_FOLDERS['inbox'];
  }

  // Check if it's a well-known folder (case-insensitive)
  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  // Try to find the folder by name
  const folderId = await getFolderIdByName(folderName);
  if (folderId) {
    const path = `me/mailFolders/${folderId}/messages`;
    console.error(`Resolved folder "${folderName}" to path: ${path}`);
    return path;
  }

  // Folder not found — throw a descriptive error
  if (folderName.includes('/')) {
    throw new Error(`Folder path "${folderName}" not found. Use 'list-folders' to see available folders and their hierarchy.`);
  } else {
    throw new Error(`Folder "${folderName}" not found. Use 'list-folders' to see available folders, or specify a path like 'ParentFolder/SubFolder' for nested folders.`);
  }
}

/**
 * Get the ID of a mail folder by its name
 * @param {string} folderName - Name of the folder to find
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(folderName) {
  // If the folder name contains '/', delegate to resolveNestedFolderPath
  if (folderName && folderName.includes('/')) {
    console.error(`Folder name "${folderName}" contains '/', resolving as nested path`);
    return resolveNestedFolderPath(folderName);
  }

  try {
    // Try the cache first – avoids any API call when the cache is warm
    if (isCacheValid()) {
      const lowerFolderName = folderName.toLowerCase();
      const cached = _folderCache.folders.find(
        f => f.displayName.toLowerCase() === lowerFolderName
      );
      if (cached) {
        console.error(`Folder "${folderName}" resolved from cache (ID: ${cached.id})`);
        return cached.id;
      }
      // Not in cache – may be a brand-new folder; fall through to API
      console.error(`Folder "${folderName}" not in cache, querying API`);
    }

    const client = await getGraphClient();

    // First try with exact match filter
    console.error(`Looking for folder with name "${folderName}"`);
    const response = await client.api('me/mailFolders')
      .filter(`displayName eq '${folderName}'`)
      .get();
    
    if (response.value && response.value.length > 0) {
      console.error(`Found folder "${folderName}" with ID: ${response.value[0].id}`);
      return response.value[0].id;
    }
    
    // If exact match fails, try to get all folders and do a case-insensitive comparison
    console.error(`No exact match found for "${folderName}", trying case-insensitive search`);
    const allFoldersResponse = await client.api('me/mailFolders')
      .top(config.FOLDER_PAGE_SIZE)
      .get();
    
    if (allFoldersResponse.value) {
      const lowerFolderName = folderName.toLowerCase();
      const matchingFolder = allFoldersResponse.value.find(
        folder => folder.displayName.toLowerCase() === lowerFolderName
      );
      
      if (matchingFolder) {
        console.error(`Found case-insensitive match for "${folderName}" with ID: ${matchingFolder.id}`);
        return matchingFolder.id;
      }
    }
    
    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Find a folder by name within a specific API endpoint.
 * Tries exact filter first, then falls back to case-insensitive comparison.
 * @param {object} client - Microsoft Graph client instance
 * @param {string} endpoint - API endpoint (e.g. 'me/mailFolders' or 'me/mailFolders/{id}/childFolders')
 * @param {string} name - Folder display name to find
 * @returns {Promise<{id: string, displayName: string}|null>} - Folder object or null if not found
 */
async function findFolderInCollection(client, endpoint, name) {
  // Try exact match filter first
  try {
    const response = await client.api(endpoint)
      .filter(`displayName eq '${name}'`)
      .get();

    if (response.value && response.value.length > 0) {
      return response.value[0];
    }
  } catch (error) {
    console.error(`Exact filter failed for "${name}" at ${endpoint}: ${error.message}`);
  }

  // Fallback: fetch all and do case-insensitive comparison
  try {
    const allResponse = await client.api(endpoint)
      .top(config.FOLDER_PAGE_SIZE)
      .get();

    if (allResponse.value) {
      const lowerName = name.toLowerCase();
      const match = allResponse.value.find(
        folder => folder.displayName.toLowerCase() === lowerName
      );
      if (match) {
        return match;
      }
    }
  } catch (error) {
    console.error(`Case-insensitive search failed for "${name}" at ${endpoint}: ${error.message}`);
  }

  return null;
}

/**
 * Resolve a nested folder path like "Finance/Cards/Amex" to the deepest folder's ID.
 * Each segment is resolved in sequence: the first under top-level mailFolders,
 * subsequent segments under the previous folder's childFolders.
 * @param {string} path - Slash-separated folder path (e.g. "Finance/Cards/Amex")
 * @returns {Promise<string>} - The ID of the deepest (last) folder in the path
 * @throws {Error} If any segment cannot be found or maxDepth is exceeded
 */
async function resolveNestedFolderPath(path) {
  const segments = path.split('/').filter(s => s.length > 0);

  if (segments.length === 0) {
    throw new Error(`Invalid folder path: "${path}"`);
  }

  if (segments.length > config.FOLDER_PATH_MAX_DEPTH) {
    throw new Error(`Folder path "${path}" exceeds maximum depth of ${config.FOLDER_PATH_MAX_DEPTH}`);
  }

  console.error(`Resolving nested folder path: "${path}" (${segments.length} segments)`);

  const client = await getGraphClient();
  let parentId = null;
  let resolvedSegments = [];

  for (const segment of segments) {
    const endpoint = parentId
      ? `me/mailFolders/${parentId}/childFolders`
      : 'me/mailFolders';

    console.error(`Looking for segment "${segment}" at endpoint: ${endpoint}`);
    const folder = await findFolderInCollection(client, endpoint, segment);

    if (!folder) {
      const parentDesc = resolvedSegments.length > 0
        ? `"${resolvedSegments.join('/')}"`
        : 'top-level folders';
      throw new Error(
        `Folder segment "${segment}" not found under ${parentDesc}. Full path: "${path}"`
      );
    }

    console.error(`Resolved segment "${segment}" to ID: ${folder.id}`);
    resolvedSegments.push(folder.displayName);
    parentId = folder.id;
  }

  console.error(`Fully resolved path "${path}" to folder ID: ${parentId}`);
  return parentId;
}

/**
 * Get all mail folders recursively
 * @param {number} maxDepth - Maximum recursion depth (default: 5)
 * @returns {Promise<Array>} - Flat array of folder objects, each with a 'path' property
 */
async function getAllFolders(maxDepth = config.FOLDER_TRAVERSAL_MAX_DEPTH) {
  // Return cached result if still valid
  if (isCacheValid()) {
    console.error(`Returning ${_folderCache.folders.length} folders from cache`);
    return _folderCache.folders;
  }

  const selectFields = config.FOLDER_SELECT_FIELDS;

  /**
   * Recursively fetch folders from an endpoint
   * @param {object} client - Graph client
   * @param {string} endpoint - API endpoint
   * @param {string} parentPath - Path of the parent folder
   * @param {number} depth - Current recursion depth
   * @returns {Promise<Array>} - Flat array of folders
   */
  async function fetchFoldersRecursive(client, endpoint, parentPath, depth) {
    try {
      const response = await client.api(endpoint)
        .top(config.FOLDER_PAGE_SIZE)
        .select(selectFields)
        .get();

      if (!response.value || response.value.length === 0) {
        return [];
      }

      const allFolders = [];

      for (const folder of response.value) {
        // Build the full path for this folder
        folder.path = parentPath ? `${parentPath}/${folder.displayName}` : folder.displayName;
        allFolders.push(folder);

        // Recurse into child folders if they exist and we haven't hit max depth
        if (folder.childFolderCount > 0 && depth < maxDepth) {
          try {
            const childEndpoint = `me/mailFolders/${folder.id}/childFolders`;
            const children = await fetchFoldersRecursive(client, childEndpoint, folder.path, depth + 1);
            allFolders.push(...children);
          } catch (error) {
            console.error(`Error getting child folders for "${folder.path}": ${error.message}`);
          }
        }
      }

      return allFolders;
    } catch (error) {
      console.error(`Error fetching folders from ${endpoint}: ${error.message}`);
      return [];
    }
  }

  try {
    const client = await getGraphClient();
    const folders = await fetchFoldersRecursive(client, 'me/mailFolders', '', 0);

    // Populate the cache
    _folderCache = { folders, timestamp: Date.now() };
    console.error(`Cached ${folders.length} folders (TTL ${config.FOLDER_CACHE_TTL_MS / 1000}s)`);

    return folders;
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  resolveNestedFolderPath,
  getAllFolders,
  invalidateFolderCache,
};
