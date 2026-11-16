/**
 * Tests for tasks list functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListTaskLists, handleListTasks } = require('../../tasks/index');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

const mockAllTasks = [
  {
    id: 'task-1',
    title: 'Buy groceries',
    status: 'notStarted',
    importance: 'normal',
    dueDateTime: { dateTime: '2024-01-20T18:00:00' },
    body: { content: 'Milk, eggs, bread' },
  },
  {
    id: 'task-2',
    title: 'Finish report',
    status: 'completed',
    importance: 'high',
    dueDateTime: { dateTime: '2024-01-15T17:00:00' },
    body: { content: 'Q4 report' },
  },
  {
    id: 'task-3',
    title: 'Call mom',
    status: 'inProgress',
    importance: 'low',
    dueDateTime: null,
    body: { content: null },
  },
];

const mockLists = [
  { id: 'list-1', displayName: 'Personal', isOwner: true, isShared: false },
  { id: 'list-2', displayName: 'Work', isOwner: true, isShared: true },
];

function createTaskMockClient(includeCompleted = false) {
  let filterCalled = false;
  let filterValue = null;

  const mockChain = {
    select: jest.fn().mockReturnThis(),
    filter: jest.fn((f) => {
      filterCalled = true;
      filterValue = f;
      return mockChain;
    }),
    get: jest.fn().mockImplementation(() => {
      const tasks = includeCompleted
        ? mockAllTasks
        : mockAllTasks.filter(t => t.status !== 'completed');
      return Promise.resolve({ value: tasks });
    }),
  };
  return {
    api: jest.fn().mockReturnValue(mockChain),
    _filterCalled: () => filterCalled,
    _filterValue: () => filterValue,
    _includeCompleted: includeCompleted,
  };
}

function createListMockClient(lists = mockLists) {
  const mockChain = {
    select: jest.fn().mockReturnThis(),
    get: jest.fn().mockResolvedValue({ value: lists }),
  };
  return {
    api: jest.fn().mockReturnValue(mockChain),
  };
}

describe('handleListTaskLists', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = createListMockClient();
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful task list listing', () => {
    test('should list all task lists', async () => {
      const result = await handleListTaskLists({});

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/todo/lists');
      expect(mockClient.api().select).toHaveBeenCalledWith('id,displayName,isOwner,isShared');

      expect(result.content[0].text).toContain('Task lists (2):');
      expect(result.content[0].text).toContain('Personal (id: list-1)');
      expect(result.content[0].text).toContain('Work (id: list-2)');
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no task lists found', async () => {
      mockClient = createListMockClient([]);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListTaskLists({});

      expect(result.content[0].text).toBe('No task lists found.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleListTaskLists({});

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      const mockChain = {
        select: jest.fn().mockReturnThis(),
        get: jest.fn().mockRejectedValue(new Error('Graph API Error')),
      };
      mockClient = { api: jest.fn().mockReturnValue(mockChain) };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListTaskLists({});

      expect(result.content[0].text).toBe('Error listing task lists: Graph API Error');
    });
  });
});

describe('handleListTasks', () => {
  let mockClient;

  function createTaskMockClient(includeCompleted = false) {
    let filterCalled = false;
    let filterValue = null;

    const mockChain = {
      select: jest.fn().mockReturnThis(),
      filter: jest.fn((f) => {
        filterCalled = true;
        filterValue = f;
        return mockChain;
      }),
      get: jest.fn().mockImplementation(() => {
        const tasks = includeCompleted
          ? mockAllTasks
          : mockAllTasks.filter(t => t.status !== 'completed');
        return Promise.resolve({ value: tasks });
      }),
    };
    return {
      api: jest.fn().mockReturnValue(mockChain),
      _filterCalled: () => filterCalled,
      _filterValue: () => filterValue,
      _includeCompleted: includeCompleted,
    };
  }

  function createListMockClient(lists = mockLists) {
    const mockChain = {
      select: jest.fn().mockReturnThis(),
      get: jest.fn().mockResolvedValue({ value: lists }),
    };
    return {
      api: jest.fn().mockReturnValue(mockChain),
    };
  }

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = createTaskMockClient(false);
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful task listing', () => {
    test('should list tasks excluding completed by default', async () => {
      const result = await handleListTasks({ listId: 'list-1' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/todo/lists/list-1/tasks');
      expect(mockClient.api().select).toHaveBeenCalledWith('id,title,status,importance,dueDateTime,body');
      expect(mockClient._filterCalled()).toBe(true);
      expect(mockClient._filterValue()).toBe("status ne 'completed'");

      // The filter should exclude completed tasks
      expect(result.content[0].text).toContain('Tasks (2):');
      expect(result.content[0].text).toContain('Buy groceries');
      expect(result.content[0].text).toContain('Call mom');
      expect(result.content[0].text).not.toContain('Finish report');
    });

    test('should include completed tasks when requested', async () => {
      mockClient = createTaskMockClient(true);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListTasks({ listId: 'list-1', includeCompleted: true });

      // Filter should NOT be called when includeCompleted is true
      expect(mockClient._filterCalled()).toBe(false);
      expect(result.content[0].text).toContain('Tasks (3):');
      // Check for completed task with checkmark
      expect(result.content[0].text).toContain('Finish report');
      expect(result.content[0].text).toContain('✓');
    });

    test('should format tasks with importance, due date, and completion status', async () => {
      const result = await handleListTasks({ listId: 'list-1' });

      expect(result.content[0].text).toContain('[normal]');
      expect(result.content[0].text).toContain('[low]');
      expect(result.content[0].text).toContain('(due: 2024-01-20T18:00:00)');
      // Note: body content is not in text fallback, only in structured data
    });

    test('should handle tasks without due date', async () => {
      const result = await handleListTasks({ listId: 'list-1' });

      expect(result.content[0].text).toContain('Call mom');
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no tasks found', async () => {
      mockClient = createTaskMockClient(true);
      // Override get to return empty array
      mockClient.api().get.mockResolvedValue({ value: [] });
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListTasks({ listId: 'list-1' });

      expect(result.content[0].text).toBe('No tasks found.');
    });
  });

  describe('validation', () => {
    test('should return error when listId is missing', async () => {
      const result = await handleListTasks({});

      expect(result.content[0].text).toBe('listId is required (from list-task-lists).');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when listId is empty string', async () => {
      const result = await handleListTasks({ listId: '' });

      expect(result.content[0].text).toBe('listId is required (from list-task-lists).');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleListTasks({ listId: 'list-1' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      const mockChain = {
        select: jest.fn().mockReturnThis(),
        filter: jest.fn().mockReturnThis(),
        get: jest.fn().mockRejectedValue(new Error('Graph API Error')),
      };
      mockClient = { api: jest.fn().mockReturnValue(mockChain) };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListTasks({ listId: 'list-1' });

      expect(result.content[0].text).toBe('Error listing tasks: Graph API Error');
    });
  });
});
