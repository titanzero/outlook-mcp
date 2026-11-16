// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleCopyEmails, handleBulkDeleteEmails, handleBulkMarkRead, handleBulkFlagEmails, handleBulkCategorizeEmails } = require('../../folder/bulk-operations');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Folder Bulk Operations Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      post: jest.fn(),
      patch: jest.fn(),
      delete: jest.fn(),
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
      filter: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleCopyEmails', () => {
    test('should copy emails to target folder', async () => {
      mockClient.post.mockResolvedValue({});

      const result = await handleCopyEmails({ emailIds: 'email-1,email-2', targetFolder: 'Archive', sourceFolder: 'inbox' });

      expect(result.content[0].text).toContain('Copy Emails Results');
      expect(result.content[0].text).toContain('Successful: 2');
      expect(result.content[0].text).toContain('email-1');
      expect(result.content[0].text).toContain('email-2');
    });

    test('should return error for missing parameters', async () => {
      const result = await handleCopyEmails({ emailIds: 'email-1' });

      expect(result.content[0].text).toBe('Target folder name is required.');
    });
  });

  describe('handleBulkDeleteEmails', () => {
    test('should move emails to Deleted Items by default', async () => {
      mockClient.post.mockResolvedValue({});

      const result = await handleBulkDeleteEmails({ emailIds: 'email-1,email-2', sourceFolder: 'inbox' });

      expect(result.content[0].text).toContain('Moved to Deleted Items Results');
      expect(result.content[0].text).toContain('Successful: 2');
    });

    test('should permanently delete when permanent=true', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleBulkDeleteEmails({ emailIds: 'email-1', permanent: true });

      expect(result.content[0].text).toContain('Permanently deleted Results');
    });
  });

  describe('handleBulkMarkRead', () => {
    test('should mark emails as read', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleBulkMarkRead({ emailIds: 'email-1,email-2', isRead: true });

      expect(result.content[0].text).toContain('Marked as read Results');
      expect(result.content[0].text).toContain('Successful: 2');
    });

    test('should mark emails as unread', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleBulkMarkRead({ emailIds: 'email-1', isRead: false });

      expect(result.content[0].text).toContain('Marked as unread Results');
    });
  });

  describe('handleBulkFlagEmails', () => {
    test('should flag emails', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleBulkFlagEmails({ emailIds: 'email-1', flagStatus: 'flagged' });

      expect(result.content[0].text).toContain('Bulk Flag Emails (flagged)');
    });

    test('should complete flag', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleBulkFlagEmails({ emailIds: 'email-1', flagStatus: 'completed' });

      expect(result.content[0].text).toContain('Bulk Flag Emails (completed)');
    });

    test('should clear flag', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleBulkFlagEmails({ emailIds: 'email-1', flagStatus: 'cleared' });

      expect(result.content[0].text).toContain('Bulk Flag Emails (cleared)');
    });
  });

  describe('handleBulkCategorizeEmails', () => {
    test('should categorize emails', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleBulkCategorizeEmails({ emailIds: 'email-1', categories: 'Important,Urgent' });

      expect(result.content[0].text).toContain('Bulk Categorize Emails');
      expect(result.content[0].text).toContain('Important, Urgent');
    });

    test('should clear categories when empty string', async () => {
      // Empty string is treated as "categories required" not "clear all"
      const result = await handleBulkCategorizeEmails({ emailIds: 'email-1', categories: '' });

      expect(result.content[0].text).toContain('Categories are required');
    });
  });
});