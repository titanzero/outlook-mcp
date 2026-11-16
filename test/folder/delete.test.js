/**
 * Tests for folder delete functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleDeleteFolder = require('../../folder/delete');
const { getGraphClient } = require('../../utils/graph-client');
const { getFolderIdByName, invalidateFolderCache } = require('../../email/folder-utils');

jest.mock('../../utils/graph-client');
jest.mock('../../email/folder-utils', () => ({
  getFolderIdByName: jest.fn(),
  invalidateFolderCache: jest.fn(),
}));

describe('handleDeleteFolder', () => {
  let mockClient;
  let mockDelete;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockDelete = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn(() => ({
        delete: mockDelete,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
    getFolderIdByName.mockResolvedValue('folder-id');
    invalidateFolderCache.mockResolvedValue();
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful folder deletion', () => {
    test('should delete folder', async () => {
      const result = await handleDeleteFolder({ name: 'Test Folder' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(getFolderIdByName).toHaveBeenCalledWith('Test Folder');
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/folder-id');
      expect(mockDelete).toHaveBeenCalledTimes(1);
      expect(invalidateFolderCache).toHaveBeenCalledTimes(1);

      expect(result.content[0].text).toBe('Successfully deleted folder "Test Folder".');
    });
  });

  describe('validation', () => {
    test('should return error when name is missing', async () => {
      const result = await handleDeleteFolder({});

      expect(result.content[0].text).toBe('Folder name is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when name is empty string', async () => {
      const result = await handleDeleteFolder({ name: '' });

      expect(result.content[0].text).toBe('Folder name is required.');
    });
  });

  describe('well-known folder protection', () => {
    const wellKnownFolders = [
      'Inbox',
      'Drafts',
      'Sent Items',
      'Deleted Items',
      'Junk Email',
      'Archive',
      'Outbox',
      'Conversation History',
    ];

    test.each(wellKnownFolders)('should not delete well-known folder "%s"', async (folderName) => {
      getFolderIdByName.mockResolvedValue('some-id');

      const result = await handleDeleteFolder({ name: folderName });

      expect(result.content[0].text).toContain(`Cannot delete well-known folder "${folderName}"`);
      expect(result.content[0].text).toContain('These folders are system-managed');
      expect(mockDelete).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleDeleteFolder({ name: 'Test Folder' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should return error when folder not found', async () => {
      getFolderIdByName.mockResolvedValue(null);

      const result = await handleDeleteFolder({ name: 'Non-existent' });

      expect(result.content[0].text).toContain('Folder "Non-existent" not found');
    });

    test('should handle non-empty folder error', async () => {
      const folderError = new Error('FolderNotEmpty');
      mockDelete.mockRejectedValue(folderError);

      const result = await handleDeleteFolder({ name: 'Test Folder' });

      expect(result.content[0].text).toContain('is not empty');
      expect(result.content[0].text).toContain('Move or delete its contents first');
    });

    test('should handle generic Graph API error', async () => {
      mockDelete.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleDeleteFolder({ name: 'Test Folder' });

      expect(result.content[0].text).toBe('Error deleting folder: Graph API Error');
    });
  });
});
