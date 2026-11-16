/**
 * Tests for folder rename functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleRenameFolder = require('../../folder/rename');
const { getGraphClient } = require('../../utils/graph-client');
const { getFolderIdByName, invalidateFolderCache } = require('../../email/folder-utils');

jest.mock('../../utils/graph-client');
jest.mock('../../email/folder-utils', () => ({
  getFolderIdByName: jest.fn(),
  invalidateFolderCache: jest.fn(),
}));

describe('handleRenameFolder', () => {
  let mockClient;
  let mockApi;
  let mockPatch;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPatch = jest.fn().mockResolvedValue({ id: 'folder-id', displayName: 'New Name' });
    mockApi = { patch: mockPatch };
    mockClient = { api: jest.fn().mockReturnValue(mockApi) };
    getGraphClient.mockResolvedValue(mockClient);
    invalidateFolderCache.mockResolvedValue();
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful folder renaming', () => {
    test('should rename folder', async () => {
      getFolderIdByName
        .mockResolvedValueOnce('folder-id')    // old folder exists
        .mockResolvedValueOnce(null);          // new name doesn't exist

      const result = await handleRenameFolder({
        name: 'Old Name',
        newName: 'New Name',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(getFolderIdByName).toHaveBeenCalledWith('Old Name');
      expect(getFolderIdByName).toHaveBeenCalledWith('New Name');
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/folder-id');
      expect(mockApi.patch).toHaveBeenCalledWith({ displayName: 'New Name' });
      expect(invalidateFolderCache).toHaveBeenCalledTimes(1);

      expect(result.content[0].text).toBe('Successfully renamed folder "Old Name" to "New Name".');
    });
  });

  describe('validation', () => {
    test('should return error when name is missing', async () => {
      const result = await handleRenameFolder({ newName: 'New Name' });
      expect(result.content[0].text).toBe('Folder name is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when newName is missing', async () => {
      const result = await handleRenameFolder({ name: 'Old Name' });
      expect(result.content[0].text).toBe('New folder name is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when new name is same as old name', async () => {
      const result = await handleRenameFolder({ name: 'Same Name', newName: 'Same Name' });
      expect(result.content[0].text).toBe('New name must be different from the current name.');
      expect(result.isError).toBe(true);
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleRenameFolder({ name: 'Old', newName: 'New' });
      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should return error when folder not found', async () => {
      getFolderIdByName.mockResolvedValue(null);

      const result = await handleRenameFolder({ name: 'Non-existent', newName: 'New' });
      expect(result.content[0].text).toContain('Folder "Non-existent" not found');
    });

    test('should return error when new name already exists', async () => {
      getFolderIdByName
        .mockResolvedValueOnce('folder-id')
        .mockResolvedValueOnce('existing-folder-id');

      const result = await handleRenameFolder({ name: 'Old', newName: 'Existing' });
      expect(result.content[0].text).toBe('A folder named "Existing" already exists.');
    });

    test('should handle Graph API error during rename', async () => {
      getFolderIdByName
        .mockResolvedValueOnce('folder-id')
        .mockResolvedValueOnce(null);
      mockApi.patch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleRenameFolder({ name: 'Old', newName: 'New' });
      expect(result.content[0].text).toBe('Error renaming folder: Graph API Error');
    });

    test('should handle error when server does not return folder ID', async () => {
      getFolderIdByName
        .mockResolvedValueOnce('folder-id')
        .mockResolvedValueOnce(null);
      mockApi.patch.mockResolvedValue({});

      const result = await handleRenameFolder({ name: 'Old', newName: 'New' });
      expect(result.content[0].text).toContain("The server didn't return a folder ID");
    });
  });
});
