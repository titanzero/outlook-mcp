// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleGetFolderTree, handleMoveFolder, handleCreateFolderPath, handleEmptyFolder, handleGetFolderPath } = require('../../folder/hierarchy');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Folder Hierarchy Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      patch: jest.fn(),
      post: jest.fn(),
      delete: jest.fn(),
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleGetFolderTree', () => {
    test('should get folder hierarchy tree', async () => {
      const mockFolders = [
        { id: 'root-1', displayName: 'Inbox', parentFolderId: null, childFolderCount: 1, totalItemCount: 100, unreadItemCount: 5, wellKnownName: 'inbox' },
        { id: 'child-1', displayName: 'Projects', parentFolderId: 'root-1', childFolderCount: 0, totalItemCount: 20, unreadItemCount: 2, wellKnownName: null },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockFolders });

      const result = await handleGetFolderTree({});

      expect(result.content[0].text).toContain('Mail Folder Hierarchy');
      expect(result.content[0].text).toContain('Inbox [inbox] [100 items (5 unread)]');
      expect(result.content[0].text).toContain('Projects [20 items (2 unread)]');
    });
  });

  describe('handleMoveFolder', () => {
    test('should move folder to new parent', async () => {
      const mockFolders = [
        { id: 'folder-1', parentFolderId: 'root-1' },
        { id: 'new-parent', parentFolderId: 'root-1' },
        { id: 'root-1', parentFolderId: null },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockFolders });
      mockClient.patch.mockResolvedValue({});

      const result = await handleMoveFolder({ folderId: 'folder-1', newParentFolderId: 'new-parent' });

      expect(result.content[0].text).toBe('Folder moved successfully.');
    });

    test('should prevent moving folder into itself', async () => {
      const result = await handleMoveFolder({ folderId: 'folder-1', newParentFolderId: 'folder-1' });

      expect(result.content[0].text).toBe('Cannot move folder into itself.');
    });

    test('should prevent cycle creation', async () => {
      const mockFolders = [
        { id: 'folder-1', parentFolderId: 'root-1' },
        { id: 'child-1', parentFolderId: 'folder-1' },
        { id: 'root-1', parentFolderId: null },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockFolders });

      const result = await handleMoveFolder({ folderId: 'root-1', newParentFolderId: 'child-1' });

      expect(result.content[0].text).toContain('Cannot move folder into its own descendant');
    });
  });

  describe('handleCreateFolderPath', () => {
    test('should create nested folder path', async () => {
      mockClient.get
        .mockResolvedValueOnce({ value: [] }) // root level - no Projects
        .mockResolvedValueOnce({ value: [] }) // Projects level - no ClientA
        .mockResolvedValueOnce({ value: [] }); // ClientA level - no 2024

      mockClient.post
        .mockResolvedValueOnce({ id: 'projects-id', displayName: 'Projects' })
        .mockResolvedValueOnce({ id: 'client-id', displayName: 'ClientA' })
        .mockResolvedValueOnce({ id: 'year-id', displayName: '2024' });

      const result = await handleCreateFolderPath({ path: 'Projects/ClientA/2024' });

      expect(result.content[0].text).toContain('Folder Path: Projects/ClientA/2024');
      expect(result.content[0].text).toContain('Created Folders (3)');
      expect(result.content[0].text).toContain('Projects');
    });
  });

  describe('handleEmptyFolder', () => {
    test('should empty folder', async () => {
      const mockEmails = [
        { id: 'email-1' },
        { id: 'email-2' },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockEmails });
      mockClient.post.mockResolvedValue({});

      const result = await handleEmptyFolder({ folderId: 'folder-1', permanent: false });

      expect(result.content[0].text).toContain('Empty Folder: Moved to Deleted Items');
      expect(result.content[0].text).toContain('Total Emails: 2');
      expect(result.content[0].text).toContain('Successful: 2');
    });
  });

  describe('handleGetFolderPath', () => {
    test('should get folder path from root', async () => {
      const mockFolders = [
        { id: 'root', displayName: 'Inbox', parentFolderId: null },
        { id: 'child', displayName: 'Projects', parentFolderId: 'root' },
        { id: 'grandchild', displayName: '2024', parentFolderId: 'child' },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockFolders });

      const result = await handleGetFolderPath({ folderId: 'grandchild' });

      expect(result.content[0].text).toContain('Folder Path: Inbox / Projects / 2024');
      expect(result.content[0].text).toContain('Depth: 2');
    });
  });
});