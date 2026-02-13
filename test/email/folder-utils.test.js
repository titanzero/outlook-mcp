const {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  resolveNestedFolderPath,
  getAllFolders,
  invalidateFolderCache
} = require('../../email/folder-utils');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');
jest.mock('../../auth', () => ({
  ensureAuthenticated: jest.fn().mockResolvedValue(true)
}));

/**
 * Helper to create a mock Graph client with chainable .api().filter().select().top().get()
 * Responses are consumed sequentially as each .api() call triggers .get().
 */
function createMockClient(responses) {
  let callIndex = 0;
  const mockClient = {
    api: jest.fn(() => {
      const currentResponse = responses[callIndex++] || { value: [] };
      const chain = {
        filter: jest.fn(() => chain),
        select: jest.fn(() => chain),
        top: jest.fn(() => chain),
        get: jest.fn().mockResolvedValue(currentResponse),
      };
      return chain;
    })
  };
  return mockClient;
}

describe('resolveFolderPath', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    invalidateFolderCache();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('well-known folders', () => {
    test('should return inbox endpoint when no folder name is provided', async () => {
      const result = await resolveFolderPath(null);
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return inbox endpoint when undefined folder name is provided', async () => {
      const result = await resolveFolderPath(undefined);
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return inbox endpoint when empty string is provided', async () => {
      const result = await resolveFolderPath('');
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return correct endpoint for well-known folders', async () => {
      const result = await resolveFolderPath('drafts');
      expect(result).toBe(WELL_KNOWN_FOLDERS['drafts']);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should handle case-insensitive well-known folder names', async () => {
      const result1 = await resolveFolderPath('INBOX');
      const result2 = await resolveFolderPath('Drafts');
      const result3 = await resolveFolderPath('SENT');

      expect(result1).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(result2).toBe(WELL_KNOWN_FOLDERS['drafts']);
      expect(result3).toBe(WELL_KNOWN_FOLDERS['sent']);
      expect(getGraphClient).not.toHaveBeenCalled();
    });
  });

  describe('custom folders', () => {
    test('should resolve custom folder by ID when found', async () => {
      const customFolderId = 'custom-folder-id-123';
      const customFolderName = 'MyCustomFolder';

      const mockClient = createMockClient([
        { value: [{ id: customFolderId, displayName: customFolderName }] }
      ]);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await resolveFolderPath(customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders');
    });

    test('should try case-insensitive search when exact match fails', async () => {
      const customFolderId = 'custom-folder-id-456';
      const customFolderName = 'ProjectAlpha';

      const mockClient = createMockClient([
        { value: [] }, // exact match fails
        { value: [
          { id: 'other-id', displayName: 'OtherFolder' },
          { id: customFolderId, displayName: 'projectalpha' }
        ]}
      ]);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await resolveFolderPath(customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(mockClient.api).toHaveBeenCalledTimes(2);
    });

    test('should throw error when simple folder not found', async () => {
      const mockClient = createMockClient([
        { value: [] }, // exact match fails
        { value: [
          { id: 'id1', displayName: 'Folder1' },
          { id: 'id2', displayName: 'Folder2' }
        ]} // case-insensitive search also fails
      ]);
      getGraphClient.mockResolvedValue(mockClient);

      await expect(resolveFolderPath('NonExistent')).rejects.toThrow(
        'Folder "NonExistent" not found'
      );
    });

    test('should throw error when nested path not found', async () => {
      // First segment "Finance" found, second segment "NonExistent" not found
      const mockClient = createMockClient([
        // findFolderInCollection for "Finance" at me/mailFolders — exact filter finds it
        { value: [{ id: 'finance-id', displayName: 'Finance' }] },
        // findFolderInCollection for "NonExistent" at childFolders — exact filter empty
        { value: [] },
        // findFolderInCollection for "NonExistent" — top(100) fallback also empty
        { value: [] }
      ]);
      getGraphClient.mockResolvedValue(mockClient);

      await expect(resolveFolderPath('Finance/NonExistent')).rejects.toThrow();
    });

    test('should resolve nested folder path', async () => {
      const mockClient = createMockClient([
        // findFolderInCollection for "Finance" at me/mailFolders — exact filter finds it
        { value: [{ id: 'finance-id', displayName: 'Finance' }] },
        // findFolderInCollection for "Cards" at me/mailFolders/finance-id/childFolders — exact filter finds it
        { value: [{ id: 'cards-id', displayName: 'Cards' }] }
      ]);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await resolveFolderPath('Finance/Cards');

      expect(result).toBe('me/mailFolders/cards-id/messages');
    });
  });
});

describe('getFolderIdByName', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    invalidateFolderCache();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return folder ID when exact match is found', async () => {
    const folderId = 'folder-id-123';
    const folderName = 'TestFolder';

    const mockClient = createMockClient([
      { value: [{ id: folderId, displayName: folderName }] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getFolderIdByName(folderName);

    expect(result).toBe(folderId);
    expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders');
  });

  test('should return folder ID when case-insensitive match is found', async () => {
    const folderId = 'folder-id-456';
    const folderName = 'TestFolder';

    const mockClient = createMockClient([
      { value: [] }, // exact match fails
      { value: [{ id: folderId, displayName: 'testfolder' }] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getFolderIdByName(folderName);

    expect(result).toBe(folderId);
    expect(mockClient.api).toHaveBeenCalledTimes(2);
  });

  test('should return null when folder is not found', async () => {
    const folderName = 'NonExistentFolder';

    const mockClient = createMockClient([
      { value: [] },
      { value: [{ id: 'id1', displayName: 'OtherFolder' }] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getFolderIdByName(folderName);

    expect(result).toBeNull();
    expect(mockClient.api).toHaveBeenCalledTimes(2);
  });

  test('should return null when API call fails', async () => {
    const folderName = 'TestFolder';

    const mockClient = {
      api: jest.fn(() => ({
        filter: jest.fn(() => ({
          get: jest.fn().mockRejectedValue(new Error('API Error'))
        }))
      }))
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getFolderIdByName(folderName);

    expect(result).toBeNull();
  });

  test('should delegate to resolveNestedFolderPath when name contains "/"', async () => {
    const mockClient = createMockClient([
      // findFolderInCollection for "Finance" at me/mailFolders — exact filter finds it
      { value: [{ id: 'finance-id', displayName: 'Finance' }] },
      // findFolderInCollection for "Cards" at me/mailFolders/finance-id/childFolders — exact filter finds it
      { value: [{ id: 'cards-id', displayName: 'Cards' }] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getFolderIdByName('Finance/Cards');

    expect(result).toBe('cards-id');
  });
});

describe('resolveNestedFolderPath', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    invalidateFolderCache();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should resolve a 2-segment path', async () => {
    const mockClient = createMockClient([
      // Segment "Finance" — exact filter at me/mailFolders
      { value: [{ id: 'finance-id', displayName: 'Finance' }] },
      // Segment "Cards" — exact filter at me/mailFolders/finance-id/childFolders
      { value: [{ id: 'cards-id', displayName: 'Cards' }] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await resolveNestedFolderPath('Finance/Cards');

    expect(result).toBe('cards-id');
    expect(mockClient.api).toHaveBeenCalledTimes(2);
  });

  test('should resolve a 3-segment path', async () => {
    const mockClient = createMockClient([
      // Segment "Finance" — exact filter at me/mailFolders
      { value: [{ id: 'finance-id', displayName: 'Finance' }] },
      // Segment "Cards" — exact filter at me/mailFolders/finance-id/childFolders
      { value: [{ id: 'cards-id', displayName: 'Cards' }] },
      // Segment "Amex" — exact filter at me/mailFolders/cards-id/childFolders
      { value: [{ id: 'amex-id', displayName: 'Amex' }] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await resolveNestedFolderPath('Finance/Cards/Amex');

    expect(result).toBe('amex-id');
    expect(mockClient.api).toHaveBeenCalledTimes(3);
  });

  test('should throw when first segment not found', async () => {
    const mockClient = createMockClient([
      // Segment "Unknown" — exact filter at me/mailFolders returns empty
      { value: [] },
      // Fallback top(100) also returns empty
      { value: [] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    await expect(resolveNestedFolderPath('Unknown/Child')).rejects.toThrow(
      'top-level folders'
    );
  });

  test('should throw when intermediate segment not found', async () => {
    const mockClient = createMockClient([
      // Segment "Finance" — exact filter finds it
      { value: [{ id: 'finance-id', displayName: 'Finance' }] },
      // Segment "Missing" — exact filter returns empty
      { value: [] },
      // Fallback top(100) also returns empty
      { value: [] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    await expect(resolveNestedFolderPath('Finance/Missing/Deep')).rejects.toThrow(
      '"Finance"'
    );
  });

  test('should throw for empty path', async () => {
    await expect(resolveNestedFolderPath('')).rejects.toThrow('Invalid folder path');
  });

  test('should throw when exceeding max depth', async () => {
    const segments = Array.from({ length: 11 }, (_, i) => `Folder${i}`);
    const deepPath = segments.join('/');

    await expect(resolveNestedFolderPath(deepPath)).rejects.toThrow(
      'exceeds maximum depth'
    );
  });

  test('should use case-insensitive fallback for segments', async () => {
    const mockClient = createMockClient([
      // Segment "finance" — exact filter returns empty
      { value: [] },
      // Fallback top(100) returns case-insensitive match
      { value: [
        { id: 'other-id', displayName: 'Other' },
        { id: 'finance-id', displayName: 'Finance' }
      ]},
      // Segment "cards" — exact filter returns empty
      { value: [] },
      // Fallback top(100) returns case-insensitive match
      { value: [{ id: 'cards-id', displayName: 'Cards' }] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await resolveNestedFolderPath('finance/cards');

    expect(result).toBe('cards-id');
  });
});

describe('getAllFolders', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    invalidateFolderCache();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return empty array when no folders', async () => {
    const mockClient = createMockClient([
      { value: [] }
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getAllFolders();

    expect(result).toEqual([]);
  });

  test('should return top-level folders with path property', async () => {
    const mockClient = createMockClient([
      { value: [
        { id: 'id1', displayName: 'Inbox', childFolderCount: 0, totalItemCount: 10, unreadItemCount: 2 },
        { id: 'id2', displayName: 'Sent', childFolderCount: 0, totalItemCount: 5, unreadItemCount: 0 }
      ]}
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getAllFolders();

    expect(result).toHaveLength(2);
    expect(result[0].path).toBe('Inbox');
    expect(result[1].path).toBe('Sent');
    expect(result[0].id).toBe('id1');
    expect(result[1].id).toBe('id2');
  });

  test('should recursively fetch child folders', async () => {
    const mockClient = createMockClient([
      // Top-level folders: Finance has 1 child
      { value: [
        { id: 'finance-id', displayName: 'Finance', childFolderCount: 1, totalItemCount: 3, unreadItemCount: 0 }
      ]},
      // Child folders of Finance
      { value: [
        { id: 'cards-id', displayName: 'Cards', childFolderCount: 0, totalItemCount: 7, unreadItemCount: 1 }
      ]}
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getAllFolders();

    expect(result).toHaveLength(2);
    expect(result[0].path).toBe('Finance');
    expect(result[0].id).toBe('finance-id');
    expect(result[1].path).toBe('Finance/Cards');
    expect(result[1].id).toBe('cards-id');
  });

  test('should respect maxDepth', async () => {
    const mockClient = createMockClient([
      // Top-level (depth 0)
      { value: [
        { id: 'a-id', displayName: 'A', childFolderCount: 1, totalItemCount: 0, unreadItemCount: 0 }
      ]},
      // Children of A (depth 1)
      { value: [
        { id: 'b-id', displayName: 'B', childFolderCount: 1, totalItemCount: 0, unreadItemCount: 0 }
      ]}
      // B's children (depth 2) should NOT be fetched because maxDepth=1
    ]);
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getAllFolders(1);

    expect(result).toHaveLength(2);
    expect(result[0].path).toBe('A');
    expect(result[1].path).toBe('A/B');
    // Should only have made 2 API calls (top-level + A's children), not 3
    expect(mockClient.api).toHaveBeenCalledTimes(2);
  });

  test('should handle API errors gracefully', async () => {
    let callIndex = 0;
    const mockClient = {
      api: jest.fn(() => {
        callIndex++;
        if (callIndex === 2) {
          // Child folders call throws
          const chain = {
            filter: jest.fn(() => chain),
            select: jest.fn(() => chain),
            top: jest.fn(() => chain),
            get: jest.fn().mockRejectedValue(new Error('Network Error')),
          };
          return chain;
        }
        // Top-level call succeeds
        const chain = {
          filter: jest.fn(() => chain),
          select: jest.fn(() => chain),
          top: jest.fn(() => chain),
          get: jest.fn().mockResolvedValue({
            value: [
              { id: 'parent-id', displayName: 'Parent', childFolderCount: 2, totalItemCount: 0, unreadItemCount: 0 }
            ]
          }),
        };
        return chain;
      })
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await getAllFolders();

    // Should return the parent folder, child folders silently skipped
    expect(result).toHaveLength(1);
    expect(result[0].path).toBe('Parent');
    expect(result[0].id).toBe('parent-id');
  });
});

