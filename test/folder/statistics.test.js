// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleGetFolderStatistics, handleGetAllFoldersStatistics, handleSearchInFolder } = require('../../folder/statistics');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Folder Statistics & Search Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
      filter: jest.fn(() => mockClient),
      orderby: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleGetFolderStatistics', () => {
    test('should get folder statistics', async () => {
      const mockFolder = { id: 'folder-1', displayName: 'Inbox', totalItemCount: 100, unreadItemCount: 5, childFolderCount: 2 };
      const mockMessages = [
        { size: 1024, hasAttachments: true, importance: 'high', isRead: false, categories: ['Work'], flag: { flagStatus: 'flagged' }, receivedDateTime: '2024-01-15T10:00:00Z' },
        { size: 2048, hasAttachments: false, importance: 'normal', isRead: true, categories: [], flag: null, receivedDateTime: '2024-01-14T10:00:00Z' },
      ];

      mockClient.get.mockResolvedValue(mockFolder);
      graphGetPaginated.mockResolvedValue({ value: mockMessages });

      const result = await handleGetFolderStatistics({ folderId: 'folder-1' });

      expect(result.content[0].text).toContain('Folder Statistics: Inbox');
      expect(result.content[0].text).toContain('Total Items: 100');
      expect(result.content[0].text).toContain('Total Size:');
      expect(result.content[0].text).toContain('With Attachments: 1');
      expect(result.content[0].text).toContain('High Importance: 1');
      expect(result.content[0].text).toContain('Flagged: 1');
    });
  });

  describe('handleGetAllFoldersStatistics', () => {
    test('should get all folders statistics', async () => {
      const mockFolders = [
        { id: 'folder-1', displayName: 'Inbox', totalItemCount: 100, unreadItemCount: 5, childFolderCount: 2, wellKnownName: 'inbox' },
        { id: 'folder-2', displayName: 'Sent', totalItemCount: 50, unreadItemCount: 0, childFolderCount: 0, wellKnownName: 'sentitems' },
      ];

      mockClient.get.mockResolvedValue({ value: mockFolders });

      const result = await handleGetAllFoldersStatistics({});

      expect(result.content[0].text).toContain('All Folders Statistics Summary');
      expect(result.content[0].text).toContain('Total Folders: 2');
      expect(result.content[0].text).toContain('Total Items: 150');
      expect(result.content[0].text).toContain('Inbox [inbox]: 100 items (5 unread)');
    });
  });

  describe('handleSearchInFolder', () => {
    test('should search in folder', async () => {
      const mockEmails = [
        { id: 'email-1', subject: 'Test Email', from: { emailAddress: { name: 'John', address: 'john@example.com' } }, toRecipients: [], ccRecipients: [], receivedDateTime: '2024-01-15T10:30:00Z', bodyPreview: 'Test content', hasAttachments: false, importance: 'normal', isRead: false, categories: ['Work'] },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      const result = await handleSearchInFolder({ folderId: 'folder-1', query: 'test', count: 50 });

      expect(result.content[0].text).toContain('Search in Folder');
      expect(result.content[0].text).toContain('Query: test');
      expect(result.content[0].text).toContain('Test Email');
    });

    test('should return error for missing parameters', async () => {
      const result = await handleSearchInFolder({ folderId: 'folder-1' });

      expect(result.content[0].text).toBe('Search query is required.');
    });
  });
});