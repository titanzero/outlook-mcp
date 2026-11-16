// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListContactFolders, handleGetContactFolder, handleCreateContactFolder, handleUpdateContactFolder, handleDeleteContactFolder, handleMoveContactToFolder } = require('../../contacts/folders');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Contact Folders Tools', () => {
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
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListContactFolders', () => {
    test('should list contact folders', async () => {
      const mockFolders = [
        { id: 'folder-1', displayName: 'Personal', parentFolderId: null, contactCount: 50 },
        { id: 'folder-2', displayName: 'Work', parentFolderId: 'folder-1', contactCount: 100 },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockFolders });

      const result = await handleListContactFolders({});

      expect(result.content[0].text).toContain('Contact Folders (2)');
      expect(result.content[0].text).toContain('Personal (50 contacts) [root]');
      expect(result.content[0].text).toContain('Work (100 contacts) [parent: folder-1]');
    });
  });

  describe('handleGetContactFolder', () => {
    test('should get contact folder with contacts', async () => {
      const mockFolder = { id: 'folder-1', displayName: 'Work', parentFolderId: null, contactCount: 100 };
      const mockContacts = [
        { id: 'contact-1', displayName: 'John Doe', emailAddresses: [{ address: 'john@company.com', name: 'John' }], mobilePhone: '555-1234', businessPhones: ['555-5678'], homePhones: [], jobTitle: 'Manager', companyName: 'Acme Corp' },
      ];

      mockClient.get.mockResolvedValue(mockFolder);
      graphGetPaginated.mockResolvedValue({ value: mockContacts });

      const result = await handleGetContactFolder({ folderId: 'folder-1', count: 50 });

      expect(result.content[0].text).toContain('Contact Folder: Work');
      expect(result.content[0].text).toContain('John Doe (john@company.com)');
    });
  });

  describe('handleCreateContactFolder', () => {
    test('should create contact folder', async () => {
      mockClient.post.mockResolvedValue({ id: 'new-folder-id', displayName: 'New Folder' });

      const result = await handleCreateContactFolder({ name: 'New Folder', parentFolderId: 'folder-1' });

      expect(result.content[0].text).toContain('Contact folder "New Folder" created successfully');
      expect(result.content[0].text).toContain('ID: new-folder-id');
    });

    test('should return error when name is missing', async () => {
      const result = await handleCreateContactFolder({});

      expect(result.content[0].text).toBe('Folder name is required.');
    });
  });

  describe('handleUpdateContactFolder', () => {
    test('should update contact folder', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleUpdateContactFolder({ folderId: 'folder-1', name: 'Updated Name' });

      expect(result.content[0].text).toBe('Contact folder updated successfully.');
    });
  });

  describe('handleDeleteContactFolder', () => {
    test('should delete contact folder', async () => {
      mockClient.get.mockResolvedValue({ contactCount: 0 });
      mockClient.delete.mockResolvedValue({});

      const result = await handleDeleteContactFolder({ folderId: 'folder-1' });

      expect(result.content[0].text).toBe('Contact folder deleted successfully.');
    });

    test('should return error when folder has contacts', async () => {
      mockClient.get.mockResolvedValue({ contactCount: 5 });

      const result = await handleDeleteContactFolder({ folderId: 'folder-1' });

      expect(result.content[0].text).toContain('Cannot delete folder with 5 contacts');
    });
  });

  describe('handleMoveContactToFolder', () => {
    test('should move contact to folder', async () => {
      mockClient.post.mockResolvedValue({});

      const result = await handleMoveContactToFolder({ contactId: 'contact-1', folderId: 'folder-2' });

      expect(result.content[0].text).toBe('Contact moved to folder successfully.');
    });
  });
});