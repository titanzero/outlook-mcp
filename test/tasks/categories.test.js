// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListCategories, handleCreateCategory, handleUpdateCategory, handleDeleteCategory, handleGetTaskCategories, handleSetTaskCategories } = require('../../tasks/categories');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Task Categories Tools', () => {
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

  describe('handleListCategories', () => {
    test('should list master categories', async () => {
      const mockCategories = [
        { name: 'Red Category', color: 'preset0', predefined: true },
        { name: 'Custom Category', color: 'preset1', predefined: false },
      ];

      mockClient.get.mockResolvedValue({ value: mockCategories });

      const result = await handleListCategories({});

      expect(result.content[0].text).toContain('Master Categories (2)');
      expect(result.content[0].text).toContain('Red Category [Red] (built-in)');
      expect(result.content[0].text).toContain('Custom Category [Orange]');
    });
  });

  describe('handleCreateCategory', () => {
    test('should create category', async () => {
      mockClient.post.mockResolvedValue({ name: 'New Category', color: 'preset2', predefined: false });

      const result = await handleCreateCategory({ name: 'New Category', color: 'preset2' });

      expect(result.content[0].text).toContain('Category "New Category" created successfully with color preset2');
    });

    test('should return error when name is missing', async () => {
      const result = await handleCreateCategory({});

      expect(result.content[0].text).toBe('Category name is required.');
    });

    test('should validate color', async () => {
      const result = await handleCreateCategory({ name: 'Test', color: 'invalid' });

      expect(result.content[0].text).toContain('Invalid color');
    });
  });

  describe('handleUpdateCategory', () => {
    test('should update category', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleUpdateCategory({ name: 'Old Name', newName: 'New Name', color: 'preset3' });

      expect(result.content[0].text).toBe('Category updated successfully.');
    });
  });

  describe('handleDeleteCategory', () => {
    test('should delete category', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleDeleteCategory({ name: 'Category to Delete' });

      expect(result.content[0].text).toBe('Category "Category to Delete" deleted successfully.');
    });
  });

  describe('handleGetTaskCategories', () => {
    test('should get task categories', async () => {
      const mockTask = { id: 'task-1', title: 'Task with Categories', categories: ['Red Category', 'Custom Category'] };

      const mockApi = { select: jest.fn().mockReturnThis(), get: jest.fn().mockResolvedValue(mockTask) };
      mockClient.api.mockReturnValue(mockApi);

      const result = await handleGetTaskCategories({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toContain('Task Categories');
      expect(result.content[0].text).toContain('Red Category');
    });
  });

  describe('handleSetTaskCategories', () => {
    test('should set task categories', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleSetTaskCategories({ listId: 'list-1', taskId: 'task-1', categories: ['Category 1', 'Category 2'] });

      expect(result.content[0].text).toBe('Task categories set to: Category 1, Category 2');
    });

    test('should clear categories when empty array', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleSetTaskCategories({ listId: 'list-1', taskId: 'task-1', categories: [] });

      expect(result.content[0].text).toBe('Task categories cleared.');
    });
  });
});