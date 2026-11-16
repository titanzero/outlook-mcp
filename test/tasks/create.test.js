/**
 * Tests for tasks create functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleCreateTask } = require('../../tasks/index');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleCreateTask', () => {
  let mockClient;
  let mockPost;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPost = jest.fn().mockResolvedValue({ id: 'task-123', title: 'New Task' });
    mockClient = {
      api: jest.fn(() => ({
        post: mockPost,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful task creation', () => {
    test('should create task with required listId and title', async () => {
      const result = await handleCreateTask({ listId: 'list-1', title: 'New Task' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/todo/lists/list-1/tasks');
      expect(mockPost).toHaveBeenCalledWith({ title: 'New Task', importance: 'normal' });

      expect(result.content[0].text).toContain('Task created successfully');
      expect(result.content[0].text).toContain('task-123');
      expect(result.content[0].text).toContain('New Task');
    });

    test('should create task with dueDateTime', async () => {
      await handleCreateTask({
        listId: 'list-1',
        title: 'Task with due date',
        dueDateTime: '2024-01-20T18:00:00',
      });

      expect(mockPost).toHaveBeenCalledWith({
        title: 'Task with due date',
        importance: 'normal',
        dueDateTime: { dateTime: '2024-01-20T18:00:00', timeZone: 'UTC' },
      });
    });

    test('should create task with importance', async () => {
      await handleCreateTask({
        listId: 'list-1',
        title: 'High importance task',
        importance: 'high',
      });

      expect(mockPost).toHaveBeenCalledWith({
        title: 'High importance task',
        importance: 'high',
      });
    });

    test('should create task with note', async () => {
      await handleCreateTask({
        listId: 'list-1',
        title: 'Task with note',
        note: 'This is a note',
      });

      expect(mockPost).toHaveBeenCalledWith({
        title: 'Task with note',
        importance: 'normal',
        body: { contentType: 'text', content: 'This is a note' },
      });
    });

    test('should create task with all fields', async () => {
      await handleCreateTask({
        listId: 'list-1',
        title: 'Complete task',
        dueDateTime: '2024-01-20T18:00:00',
        importance: 'high',
        note: 'Details here',
      });

      const task = mockPost.mock.calls[0][0];
      expect(task.title).toBe('Complete task');
      expect(task.importance).toBe('high');
      expect(task.dueDateTime).toEqual({ dateTime: '2024-01-20T18:00:00', timeZone: 'UTC' });
      expect(task.body).toEqual({ contentType: 'text', content: 'Details here' });
    });
  });

  describe('validation', () => {
    test('should return error when listId is missing', async () => {
      const result = await handleCreateTask({ title: 'Test' });

      expect(result.content[0].text).toBe('listId is required (from list-task-lists).');
      expect(result.isError).toBe(true);
    });

    test('should return error when title is missing', async () => {
      const result = await handleCreateTask({ listId: 'list-1' });

      expect(result.content[0].text).toBe('title is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when listId is empty string', async () => {
      const result = await handleCreateTask({ listId: '', title: 'Test' });

      expect(result.content[0].text).toBe('listId is required (from list-task-lists).');
    });

    test('should return error when title is empty string', async () => {
      const result = await handleCreateTask({ listId: 'list-1', title: '' });

      expect(result.content[0].text).toBe('title is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleCreateTask({ listId: 'list-1', title: 'Test' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleCreateTask({ listId: 'list-1', title: 'Test' });

      expect(result.content[0].text).toBe('Error creating task: Graph API Error');
    });
  });
});
