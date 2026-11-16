// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListFolderPermissions, handleAddFolderPermission, handleUpdateFolderPermission, handleRemoveFolderPermission } = require('../../folder/permissions');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Folder Permissions Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      patch: jest.fn(),
      delete: jest.fn(),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListFolderPermissions', () => {
    test('should list folder permissions', async () => {
      const mockPermissions = [
        { id: 'perm-1', emailAddress: { address: 'user@company.com', name: 'User One' }, role: 'editor', isInsideOrganization: true, isRemovable: true },
        { id: 'perm-2', emailAddress: { address: 'external@partner.com', name: 'External User' }, role: 'reviewer', isInsideOrganization: false, isRemovable: true },
      ];

      mockClient.get.mockResolvedValue({ value: mockPermissions });

      const result = await handleListFolderPermissions({ folderId: 'folder-1' });

      expect(result.content[0].text).toContain('Folder Permissions (2)');
      expect(result.content[0].text).toContain('User One <user@company.com>: Editor (internal) [removable]');
      expect(result.content[0].text).toContain('External User <external@partner.com>: Reviewer (external) [removable]');
    });
  });

  describe('handleAddFolderPermission', () => {
    test('should add folder permission', async () => {
      mockClient.post.mockResolvedValue({});

      const result = await handleAddFolderPermission({ folderId: 'folder-1', emailAddress: 'new@company.com', role: 'author' });

      expect(result.content[0].text).toBe('Folder permission added for new@company.com as "author".');
    });

    test('should validate role', async () => {
      const result = await handleAddFolderPermission({ folderId: 'folder-1', emailAddress: 'user@company.com', role: 'invalid' });

      expect(result.content[0].text).toContain('Invalid role');
    });
  });

  describe('handleUpdateFolderPermission', () => {
    test('should update folder permission', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleUpdateFolderPermission({ folderId: 'folder-1', permissionId: 'perm-1', role: 'editor' });

      expect(result.content[0].text).toBe('Folder permission updated to "editor" successfully.');
    });
  });

  describe('handleRemoveFolderPermission', () => {
    test('should remove folder permission', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleRemoveFolderPermission({ folderId: 'folder-1', permissionId: 'perm-1' });

      expect(result.content[0].text).toBe('Folder permission removed successfully.');
    });
  });
});