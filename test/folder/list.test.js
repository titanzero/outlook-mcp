/**
 * Tests for folder list functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleListFolders = require('../../folder/list');
const { getAllFolders } = require('../../email/folder-utils');

jest.mock('../../email/folder-utils');
jest.mock('../../utils/response-formatter', () => ({
  formatResponse: jest.fn((structured, text) => text),
}));

const { formatResponse } = require('../../utils/response-formatter');

describe('handleListFolders', () => {
  const mockFolders = [
    {
      id: 'f1',
      displayName: 'Inbox',
      path: 'Inbox',
      totalItemCount: 10,
      unreadItemCount: 3,
      parentFolderId: null,
    },
    {
      id: 'f2',
      displayName: 'Drafts',
      path: 'Drafts',
      totalItemCount: 2,
      unreadItemCount: 0,
      parentFolderId: null,
    },
    {
      id: 'f3',
      displayName: 'Project Alpha',
      path: 'Inbox/Project Alpha',
      totalItemCount: 25,
      unreadItemCount: 5,
      parentFolderId: 'f1',
    },
    {
      id: 'f4',
      displayName: 'Archive',
      path: 'Archive',
      totalItemCount: 100,
      unreadItemCount: 0,
      parentFolderId: null,
    },
  ];

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    getAllFolders.mockResolvedValue(mockFolders);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful folder listing', () => {
    test('should list folders without item counts by default', async () => {
      const result = await handleListFolders({});

      expect(getAllFolders).toHaveBeenCalledTimes(1);
      expect(formatResponse).toHaveBeenCalled();
      expect(result.content[0].text).toContain('Found 4 folders:');
      expect(result.content[0].text).toContain('Inbox');
      expect(result.content[0].text).toContain('Drafts');
      expect(result.content[0].text).toContain('Project Alpha');
      expect(result.content[0].text).toContain('Archive');
    });

    test('should list folders with item counts when requested', async () => {
      const result = await handleListFolders({ includeItemCounts: true });

      expect(result.content[0].text).toContain('10 items');
      expect(result.content[0].text).toContain('(3 unread)');
      expect(result.content[0].text).toContain('2 items');
      expect(result.content[0].text).toContain('25 items');
      expect(result.content[0].text).toContain('(5 unread)');
      expect(result.content[0].text).toContain('100 items');
    });

    test('should format as hierarchy when includeChildren is true', async () => {
      const result = await handleListFolders({ includeChildren: true });

      expect(result.content[0].text).toContain('Folder Hierarchy:');
    });

    test('should format as flat list when includeChildren is false', async () => {
      const result = await handleListFolders({ includeChildren: false });

      expect(result.content[0].text).toContain('Found 4 folders:');
    });

    test('should sort well-known folders first', async () => {
      const result = await handleListFolders({});

      // Inbox should appear before custom folders
      const inboxIndex = result.content[0].text.indexOf('Inbox');
      const projectIndex = result.content[0].text.indexOf('Project Alpha');
      expect(inboxIndex).toBeLessThan(projectIndex);
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no folders found', async () => {
      getAllFolders.mockResolvedValue([]);

      const result = await handleListFolders({});

      expect(result.content[0].text).toBe('No folders found.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getAllFolders.mockRejectedValue(authError);

      const result = await handleListFolders({});

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      getAllFolders.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleListFolders({});

      expect(result.content[0].text).toBe('Error listing folders: Graph API Error');
    });
  });
});
