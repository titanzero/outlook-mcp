/**
 * Tests for tasks delete functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleDeleteTask } = require('../../tasks/index');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleDeleteTask', () => {
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
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful task deletion', () => {
    test('should delete task', async () => {
      const result = await handleDeleteTask({ listId: 'list-1', taskId: 'task-123' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/todo/lists/list-1/tasks/task-123');
      expect(mockDelete).toHaveBeenCalledTimes(1);

      expect(result.content[0].text).toBe('Task deleted successfully.');
    });
  });

  describe('validation', () => {
    test('should return error when listId is missing', async () => {
      const result = await handleDeleteTask({ taskId: 'task-123' });

      expect(result.content[0].text).toBe('listId is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when taskId is missing', async () => {
      const result = await handleDeleteTask({ listId: 'list-1' });

      expect(result.content[0].text).toBe('taskId is required.');
      expect(result.isError).toBe(true);
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleDeleteTask({ listId: 'list-1', taskId: 'task-123' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockDelete.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleDeleteTask({ listId: 'list-1', taskId: 'task-123' });

      expect(result.content[0].text).toBe('Error deleting task: Graph API Error');
    });
  });
});
