/**
 * Tests for tasks complete functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleCompleteTask } = require('../../tasks/index');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleCompleteTask', () => {
  let mockClient;
  let mockPatch;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPatch = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn(() => ({
        patch: mockPatch,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful task completion', () => {
    test('should mark task as completed', async () => {
      const result = await handleCompleteTask({ listId: 'list-1', taskId: 'task-123' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/todo/lists/list-1/tasks/task-123');
      expect(mockPatch).toHaveBeenCalledWith({
        status: 'completed',
        completedDateTime: expect.objectContaining({
          dateTime: expect.any(String),
          timeZone: 'UTC',
        }),
      });
      expect(result.content[0].text).toBe('Task marked as completed.');
    });

    test('should include current date/time in completedDateTime', async () => {
      const beforeTime = new Date().toISOString();
      await handleCompleteTask({ listId: 'list-1', taskId: 'task-123' });
      const afterTime = new Date().toISOString();

      const patch = mockPatch.mock.calls[0][0];
      const completedTime = new Date(patch.completedDateTime.dateTime).toISOString();
      expect(completedTime >= beforeTime && completedTime <= afterTime).toBe(true);
    });
  });

  describe('validation', () => {
    test('should return error when listId is missing', async () => {
      const result = await handleCompleteTask({ taskId: 'task-123' });

      expect(result.content[0].text).toBe('listId is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when taskId is missing', async () => {
      const result = await handleCompleteTask({ listId: 'list-1' });

      expect(result.content[0].text).toBe('taskId is required.');
      expect(result.isError).toBe(true);
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleCompleteTask({ listId: 'list-1', taskId: 'task-123' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPatch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleCompleteTask({ listId: 'list-1', taskId: 'task-123' });

      expect(result.content[0].text).toBe('Error completing task: Graph API Error');
    });
  });
});
