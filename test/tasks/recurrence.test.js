// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleSetTaskRecurrence, handleGetTaskRecurrence, handleRemoveTaskRecurrence, handleRecurrenceExamples } = require('../../tasks/recurrence');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Task Recurrence Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      patch: jest.fn(),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleSetTaskRecurrence', () => {
    test('should set task recurrence', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleSetTaskRecurrence({
        listId: 'list-1',
        taskId: 'task-1',
        pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday', 'wednesday', 'friday'] },
        range: { type: 'noEnd', startDate: '2024-01-01' },
      });

      expect(result.content[0].text).toBe('Task recurrence set successfully.');
    });

    test('should return error for missing parameters', async () => {
      const result = await handleSetTaskRecurrence({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toContain('required');
    });
  });

  describe('handleGetTaskRecurrence', () => {
    test('should get task recurrence with human-readable description', async () => {
      const mockTask = {
        id: 'task-1',
        title: 'Weekly Report',
        recurrence: {
          pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday', 'wednesday', 'friday'] },
          range: { type: 'noEnd', startDate: '2024-01-01' },
        },
      };

      const mockApi = { select: jest.fn().mockReturnThis(), get: jest.fn().mockResolvedValue(mockTask) };
      mockClient.api.mockReturnValue(mockApi);

      const result = await handleGetTaskRecurrence({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toContain('Task Recurrence');
      expect(result.content[0].text).toContain('Weekly Report');
    });

    test('should return message when no recurrence', async () => {
      const mockTask = { id: 'task-1', title: 'No Recurrence', recurrence: null };

      const mockApi = { select: jest.fn().mockReturnThis(), get: jest.fn().mockResolvedValue(mockTask) };
      mockClient.api.mockReturnValue(mockApi);

      const result = await handleGetTaskRecurrence({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toContain('no recurrence pattern');
    });
  });

  describe('handleRemoveTaskRecurrence', () => {
    test('should remove task recurrence', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleRemoveTaskRecurrence({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toBe('Task recurrence removed successfully.');
    });
  });

  describe('handleRecurrenceExamples', () => {
    test('should return recurrence examples', async () => {
      const result = await handleRecurrenceExamples({});

      expect(result.content[0].text).toContain('TASK RECURRENCE PATTERN EXAMPLES');
      expect(result.content[0].text).toContain('DAILY');
      expect(result.content[0].text).toContain('WEEKLY');
      expect(result.content[0].text).toContain('MONTHLY');
      expect(result.content[0].text).toContain('YEARLY');
      expect(result.content[0].text).toContain('RANGE TYPES');
      expect(result.content[0].text).toContain('noEnd');
      expect(result.content[0].text).toContain('endDate');
      expect(result.content[0].text).toContain('numbered');
    });
  });
});