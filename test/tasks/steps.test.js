// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListTaskSteps, handleCreateTaskStep, handleUpdateTaskStep, handleDeleteTaskStep, handleCompleteTaskStep } = require('../../tasks/steps');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Task Steps / Checklist Tools', () => {
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
      orderby: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListTaskSteps', () => {
    test('should list task steps', async () => {
      const mockSteps = [
        { id: 'step-1', title: 'Step 1', status: 'notStarted', createdDateTime: '2024-01-15T10:00:00Z', lastModifiedDateTime: '2024-01-15T10:00:00Z' },
        { id: 'step-2', title: 'Step 2', status: 'completed', createdDateTime: '2024-01-15T10:05:00Z', lastModifiedDateTime: '2024-01-15T10:10:00Z' },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockSteps });

      const result = await handleListTaskSteps({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toContain('Task Steps (2)');
      expect(result.content[0].text).toContain('☐ Step 1');
      expect(result.content[0].text).toContain('✓ Step 2');
    });

    test('should return error when listId is missing', async () => {
      const result = await handleListTaskSteps({ taskId: 'task-1' });

      expect(result.content[0].text).toBe('listId is required (from list-task-lists).');
    });
  });

  describe('handleCreateTaskStep', () => {
    test('should create task step', async () => {
      mockClient.post.mockResolvedValue({ id: 'step-3', title: 'Step 3' });

      const result = await handleCreateTaskStep({ listId: 'list-1', taskId: 'task-1', title: 'Step 3' });

      expect(result.content[0].text).toContain('Task step created successfully');
      expect(result.content[0].text).toContain('Step 3');
    });
  });

  describe('handleUpdateTaskStep', () => {
    test('should update task step', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleUpdateTaskStep({ listId: 'list-1', taskId: 'task-1', stepId: 'step-1', status: 'completed' });

      expect(result.content[0].text).toBe('Task step updated successfully.');
    });
  });

  describe('handleDeleteTaskStep', () => {
    test('should delete task step', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleDeleteTaskStep({ listId: 'list-1', taskId: 'task-1', stepId: 'step-1' });

      expect(result.content[0].text).toBe('Task step deleted successfully.');
    });
  });

  describe('handleCompleteTaskStep', () => {
    test('should complete task step', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleCompleteTaskStep({ listId: 'list-1', taskId: 'task-1', stepId: 'step-1' });

      expect(result.content[0].text).toBe('Task step marked as completed.');
    });
  });
});