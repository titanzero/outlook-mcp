/**
 * Tests for tasks update functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleUpdateTask = require('../../tasks/update');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleUpdateTask', () => {
  let mockClient;
  let mockPatch;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPatch = jest.fn().mockResolvedValue({ id: 'task-123', title: 'Updated Task', status: 'notStarted' });
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

  describe('successful task updates', () => {
    test('should update task title', async () => {
      const result = await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        title: 'Updated Title',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/todo/lists/list-1/tasks/task-123');
      expect(mockPatch).toHaveBeenCalledWith({ title: 'Updated Title' });
      expect(result.content[0].text).toContain('Task updated successfully');
      expect(result.content[0].text).toContain('Updated Task');
    });

    test('should update task importance', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        importance: 'high',
      });

      expect(mockPatch).toHaveBeenCalledWith({ importance: 'high' });
    });

    test('should update task dueDateTime', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        dueDateTime: '2024-01-25T18:00:00',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        dueDateTime: { dateTime: '2024-01-25T18:00:00', timeZone: 'UTC' },
      });
    });

    test('should clear dueDateTime when empty string provided', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        dueDateTime: '',
      });

      expect(mockPatch).toHaveBeenCalledWith({ dueDateTime: null });
    });

    test('should update task note', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        note: 'New note content',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        body: { contentType: 'text', content: 'New note content' },
      });
    });

    test('should clear note when empty string provided', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        note: '',
      });

      expect(mockPatch).toHaveBeenCalledWith({ body: null });
    });

    test('should update task status', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        status: 'completed',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        status: 'completed',
        completedDateTime: expect.objectContaining({
          dateTime: expect.any(String),
          timeZone: 'UTC',
        }),
      });
    });

    test('should update task status without completedDateTime when not completed', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        status: 'inProgress',
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.status).toBe('inProgress');
      expect(patch.completedDateTime).toBeUndefined();
    });

    test('should update multiple fields at once', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        title: 'New Title',
        importance: 'high',
        dueDateTime: '2024-01-25T18:00:00',
        note: 'New note',
        status: 'inProgress',
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.title).toBe('New Title');
      expect(patch.importance).toBe('high');
      expect(patch.dueDateTime).toEqual({ dateTime: '2024-01-25T18:00:00', timeZone: 'UTC' });
      expect(patch.body).toEqual({ contentType: 'text', content: 'New note' });
      expect(patch.status).toBe('inProgress');
      expect(patch.completedDateTime).toBeUndefined();
    });

    test('should not include undefined fields in patch', async () => {
      await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
        title: 'Only Title',
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.title).toBe('Only Title');
      expect(patch.importance).toBeUndefined();
      expect(patch.dueDateTime).toBeUndefined();
      expect(patch.body).toBeUndefined();
      expect(patch.status).toBeUndefined();
    });

    test('should return early if no fields to update', async () => {
      const result = await handleUpdateTask({
        listId: 'list-1',
        taskId: 'task-123',
      });

      expect(mockPatch).not.toHaveBeenCalled();
      expect(result.content[0].text).toBe('No changes provided.');
    });
  });

  describe('validation', () => {
    test('should return error when listId is missing', async () => {
      const result = await handleUpdateTask({ taskId: 'task-123', title: 'Test' });

      expect(result.content[0].text).toBe('listId is required (from list-task-lists).');
      expect(result.isError).toBe(true);
    });

    test('should return error when taskId is missing', async () => {
      const result = await handleUpdateTask({ listId: 'list-1', title: 'Test' });

      expect(result.content[0].text).toBe('taskId is required (from list-tasks).');
      expect(result.isError).toBe(true);
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleUpdateTask({ listId: 'list-1', taskId: 'task-123', title: 'Test' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPatch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleUpdateTask({ listId: 'list-1', taskId: 'task-123', title: 'Test' });

      expect(result.content[0].text).toBe('Error updating task: Graph API Error');
    });
  });
});
