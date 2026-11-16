/**
 * Tests for folder create functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleCreateFolder = require('../../folder/create');
const { getGraphClient } = require('../../utils/graph-client');
const { getFolderIdByName, invalidateFolderCache } = require('../../email/folder-utils');

jest.mock('../../utils/graph-client');
jest.mock('../../email/folder-utils', () => ({
  getFolderIdByName: jest.fn(),
  invalidateFolderCache: jest.fn(),
}));

describe('handleCreateFolder', () => {
  let mockClient;
  let mockPost;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPost = jest.fn().mockResolvedValue({ id: 'new-folder-id' });
    mockClient = {
      api: jest.fn(() => ({
        post: mockPost,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
    getFolderIdByName.mockResolvedValue(null);
    invalidateFolderCache.mockResolvedValue();
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful folder creation', () => {
    test('should create folder at root level', async () => {
      const result = await handleCreateFolder({ name: 'New Folder' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(getFolderIdByName).toHaveBeenCalledWith('New Folder');
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders');
      expect(mockPost).toHaveBeenCalledWith({ displayName: 'New Folder' });
      expect(invalidateFolderCache).toHaveBeenCalledTimes(1);

      expect(result.content[0].text).toContain('Successfully created folder "New Folder" at the root level.');
    });

    test('should create folder inside parent folder', async () => {
      getFolderIdByName
        .mockResolvedValueOnce(null) // Check if folder exists
        .mockResolvedValueOnce('parent-folder-id'); // Get parent folder ID

      const result = await handleCreateFolder({
        name: 'Sub Folder',
        parentFolder: 'Parent Folder',
      });

      expect(getFolderIdByName).toHaveBeenNthCalledWith(2, 'Parent Folder');
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/parent-folder-id/childFolders');
      expect(result.content[0].text).toContain('Successfully created folder "Sub Folder" inside "Parent Folder".');
    });

    test('should return error when folder already exists', async () => {
      getFolderIdByName.mockResolvedValue('existing-folder-id');

      const result = await handleCreateFolder({ name: 'Existing Folder' });

      expect(result.content[0].text).toBe('A folder named "Existing Folder" already exists.');
      expect(mockPost).not.toHaveBeenCalled();
    });

    test('should return error when parent folder not found', async () => {
      getFolderIdByName
        .mockResolvedValueOnce(null) // Check if folder exists
        .mockResolvedValueOnce(null); // Parent folder not found

      const result = await handleCreateFolder({
        name: 'Sub Folder',
        parentFolder: 'Non-existent Parent',
      });

      expect(result.content[0].text).toContain('Parent folder "Non-existent Parent" not found');
    });
  });

  describe('validation', () => {
    test('should return error when name is missing', async () => {
      const result = await handleCreateFolder({});

      expect(result.content[0].text).toBe('Folder name is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when name is empty string', async () => {
      const result = await handleCreateFolder({ name: '' });

      expect(result.content[0].text).toBe('Folder name is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleCreateFolder({ name: 'Test Folder' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error during creation', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleCreateFolder({ name: 'Test Folder' });

      expect(result.content[0].text).toBe('Error creating folder: Graph API Error');
    });

    test('should handle error when server does not return folder ID', async () => {
      mockPost.mockResolvedValue({});

      const result = await handleCreateFolder({ name: 'Test Folder' });

      expect(result.content[0].text).toContain("The server didn't return a folder ID");
    });
  });
});
