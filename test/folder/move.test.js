/**
 * Tests for folder move emails functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleMoveEmails = require('../../folder/move');
const { getGraphClient } = require('../../utils/graph-client');
const { getFolderIdByName, invalidateFolderCache } = require('../../email/folder-utils');

jest.mock('../../utils/graph-client');
jest.mock('../../email/folder-utils', () => ({
  getFolderIdByName: jest.fn(),
  invalidateFolderCache: jest.fn(),
}));

describe('handleMoveEmails', () => {
  let mockClient;
  let mockPost;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPost = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn(() => ({
        post: mockPost,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
    getFolderIdByName.mockResolvedValue('target-folder-id');
    invalidateFolderCache.mockResolvedValue();
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful email moving', () => {
    test('should move single email to target folder', async () => {
      const result = await handleMoveEmails({
        emailIds: 'email-123',
        targetFolder: 'Archive',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(getFolderIdByName).toHaveBeenCalledWith('Archive');
      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-123/move');
      expect(mockPost).toHaveBeenCalledWith({ destinationId: 'target-folder-id' });
      expect(invalidateFolderCache).toHaveBeenCalledTimes(1);

      expect(result.content[0].text).toContain('Successfully moved 1 email(s) to "Archive".');
    });

    test('should move multiple emails to target folder', async () => {
      const result = await handleMoveEmails({
        emailIds: 'email-1, email-2, email-3',
        targetFolder: 'Archive',
      });

      expect(mockClient.api).toHaveBeenCalledTimes(3);
      expect(mockPost).toHaveBeenCalledTimes(3);
      expect(result.content[0].text).toContain('Successfully moved 3 email(s) to "Archive".');
    });

    test('should handle partial failures', async () => {
      mockPost
        .mockResolvedValueOnce({}) // email-1 success
        .mockRejectedValueOnce(new Error('Not found')) // email-2 fail
        .mockResolvedValueOnce({}); // email-3 success

      const result = await handleMoveEmails({
        emailIds: 'email-1, email-2, email-3',
        targetFolder: 'Archive',
      });

      expect(result.content[0].text).toContain('Successfully moved 2 email(s) to "Archive".');
      expect(result.content[0].text).toContain('Failed to move 1 email(s)');
      expect(result.content[0].text).toContain('Not found');
    });

    test('should trim whitespace from email IDs', async () => {
      await handleMoveEmails({
        emailIds: '  email-1 , email-2  ',
        targetFolder: 'Archive',
      });

      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-1/move');
      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-2/move');
    });

    test('should ignore empty email IDs', async () => {
      await handleMoveEmails({
        emailIds: 'email-1, , email-2, ',
        targetFolder: 'Archive',
      });

      expect(mockClient.api).toHaveBeenCalledTimes(2);
    });
  });

  describe('validation', () => {
    test('should return error when emailIds is missing', async () => {
      const result = await handleMoveEmails({ targetFolder: 'Archive' });

      expect(result.content[0].text).toContain('Email IDs are required');
      expect(result.isError).toBe(true);
    });

    test('should return error when targetFolder is missing', async () => {
      const result = await handleMoveEmails({ emailIds: 'email-1' });

      expect(result.content[0].text).toBe('Target folder name is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when no valid email IDs provided', async () => {
      const result = await handleMoveEmails({ emailIds: ', , ', targetFolder: 'Archive' });

      expect(result.content[0].text).toBe('No valid email IDs provided.');
      expect(result.isError).toBe(true);
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleMoveEmails({
        emailIds: 'email-1',
        targetFolder: 'Archive',
      });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle target folder not found', async () => {
      getFolderIdByName.mockResolvedValue(null);

      const result = await handleMoveEmails({
        emailIds: 'email-1',
        targetFolder: 'Non-existent',
      });

      expect(result.content[0].text).toContain('Target folder "Non-existent" not found');
    });

    test('should handle Graph API error for all emails', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleMoveEmails({
        emailIds: 'email-1, email-2',
        targetFolder: 'Archive',
      });

      expect(result.content[0].text).toContain('Failed to move 2 email(s)');
      expect(result.content[0].text).toContain('Graph API Error');
    });
  });
});
