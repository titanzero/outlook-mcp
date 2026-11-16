/**
 * Tests for completion handlers
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const {
  getFolderCompletions,
  getCalendarCompletions,
  getContactCompletions,
  getTaskListCompletions,
  getCategoryCompletions,
  getRuleCompletions,
  getImportanceCompletions,
  getEmailFolderCompletions,
  getEventCompletions,
  registerCompletionCallbacks,
} = require('../../completion/index');
const { getGraphClient } = require('../../utils/graph-client');
const { getAllFolders } = require('../../email/folder-utils');

jest.mock('../../utils/graph-client');
jest.mock('../../email/folder-utils');
jest.mock('../../utils/response-helpers', () => ({
  isAuthError: jest.fn(),
}));

describe('Completion handlers', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('getFolderCompletions', () => {
    const mockFolders = [
      { displayName: 'Inbox', path: 'Inbox' },
      { displayName: 'Sent Items', path: 'Sent Items' },
      { displayName: 'Drafts', path: 'Drafts' },
      { displayName: 'Projects/Alpha', path: 'Projects/Alpha' },
      { displayName: 'Projects/Beta', path: 'Projects/Beta' },
    ];

    beforeEach(() => {
      getAllFolders.mockResolvedValue(mockFolders);
    });

    test('should return matching folders', async () => {
      const results = await getFolderCompletions('pro');
      expect(results).toHaveLength(2);
      expect(results.map(r => r.value)).toContain('Projects/Alpha');
      expect(results.map(r => r.value)).toContain('Projects/Beta');
    });

    test('should be case insensitive', async () => {
      const results = await getFolderCompletions('INBOX');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('Inbox');
    });

    test('should return empty array on auth error', async () => {
      const { isAuthError } = require('../../utils/response-helpers');
      isAuthError.mockReturnValue(true);
      getAllFolders.mockRejectedValue(new Error('Auth failed'));

      const results = await getFolderCompletions('test');
      expect(results).toEqual([]);
    });

    test('should return empty array on error', async () => {
      getAllFolders.mockRejectedValue(new Error('Network error'));

      const results = await getFolderCompletions('test');
      expect(results).toEqual([]);
    });

    test('should limit to 20 results', async () => {
      const manyFolders = Array.from({ length: 25 }, (_, i) => ({ displayName: `Folder${i}`, path: `Folder${i}` }));
      getAllFolders.mockResolvedValue(manyFolders);

      const results = await getFolderCompletions('Folder');
      expect(results).toHaveLength(20);
    });
  });

  describe('getCalendarCompletions', () => {
    const mockCalendars = [
      { name: 'Personal' },
      { name: 'Work' },
      { name: 'Shared Calendar' },
    ];

    beforeEach(() => {
      const mockClient = {
        api: jest.fn(() => ({
          select: jest.fn().mockReturnThis(),
          get: jest.fn().mockResolvedValue({ value: mockCalendars }),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);
    });

    test('should return matching calendars', async () => {
      const results = await getCalendarCompletions('per');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('Personal');
    });

    test('should return empty array on auth error', async () => {
      const { isAuthError } = require('../../utils/response-helpers');
      isAuthError.mockReturnValue(true);
      getGraphClient.mockRejectedValue(new Error('Auth failed'));

      const results = await getCalendarCompletions('test');
      expect(results).toEqual([]);
    });
  });

  describe('getContactCompletions', () => {
    const mockContacts = [
      { displayName: 'John Doe', emailAddresses: [{ address: 'john@example.com' }] },
      { displayName: 'Jane Smith', emailAddresses: [{ address: 'jane@work.com' }] },
    ];

    beforeEach(() => {
      const mockClient = {
        api: jest.fn(() => ({
          select: jest.fn().mockReturnThis(),
          top: jest.fn().mockReturnThis(),
          get: jest.fn().mockResolvedValue({ value: mockContacts }),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);
    });

    test('should return matching contacts by email', async () => {
      const results = await getContactCompletions('john');
      expect(results.length).toBeGreaterThanOrEqual(1);
      expect(results.some(r => r.value === 'john@example.com')).toBe(true);
    });

    test('should return matching contacts by name', async () => {
      const results = await getContactCompletions('jane');
      expect(results.length).toBeGreaterThanOrEqual(1);
      expect(results.some(r => r.value === 'Jane Smith')).toBe(true);
    });

    test('should return empty array on auth error', async () => {
      const { isAuthError } = require('../../utils/response-helpers');
      isAuthError.mockReturnValue(true);
      getGraphClient.mockRejectedValue(new Error('Auth failed'));

      const results = await getContactCompletions('test');
      expect(results).toEqual([]);
    });
  });

  describe('getTaskListCompletions', () => {
    const mockLists = [
      { displayName: 'Personal' },
      { displayName: 'Work' },
      { displayName: 'Shopping' },
    ];

    beforeEach(() => {
      const mockClient = {
        api: jest.fn(() => ({
          select: jest.fn().mockReturnThis(),
          get: jest.fn().mockResolvedValue({ value: mockLists }),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);
    });

    test('should return matching task lists', async () => {
      const results = await getTaskListCompletions('per');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('Personal');
    });
  });

  describe('getCategoryCompletions', () => {
    test('should return matching categories', async () => {
      const results = await getCategoryCompletions('red');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('Red Category');
    });

    test('should return matching categories case insensitive', async () => {
      const results = await getCategoryCompletions('IMPORTANT');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('Important');
    });

    test('should return matching custom categories', async () => {
      const results = await getCategoryCompletions('work');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('Work');
    });
  });

  describe('getRuleCompletions', () => {
    const mockRules = [
      { displayName: 'Move GitHub emails' },
      { displayName: 'Archive newsletters' },
    ];

    beforeEach(() => {
      const mockClient = {
        api: jest.fn(() => ({
          select: jest.fn().mockReturnThis(),
          get: jest.fn().mockResolvedValue({ value: mockRules }),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);
    });

    test('should return matching rules', async () => {
      const results = await getRuleCompletions('github');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('Move GitHub emails');
    });

    test('should return empty array on auth error', async () => {
      const { isAuthError } = require('../../utils/response-helpers');
      isAuthError.mockReturnValue(true);
      getGraphClient.mockRejectedValue(new Error('Auth failed'));

      const results = await getRuleCompletions('test');
      expect(results).toEqual([]);
    });
  });

  describe('getImportanceCompletions', () => {
    test('should return matching importance levels', async () => {
      const results = await getImportanceCompletions('hi');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('high');
    });

    test('should return all when empty', async () => {
      const results = await getImportanceCompletions('');
      expect(results).toHaveLength(3);
      expect(results.map(r => r.value)).toEqual(['low', 'normal', 'high']);
    });
  });

  describe('getEmailFolderCompletions', () => {
    test('should return matching well-known folders', async () => {
      const results = await getEmailFolderCompletions('in');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('inbox');
    });

    test('should return matching well-known folders for sent', async () => {
      const results = await getEmailFolderCompletions('sent');
      expect(results).toHaveLength(1);
      expect(results[0].value).toBe('sent');
    });
  });

  describe('getEventCompletions', () => {
    const mockEvents = [
      { id: 'event-1', subject: 'Team Meeting' },
      { id: 'event-2', subject: 'Project Review' },
    ];

    beforeEach(() => {
      const mockClient = {
        api: jest.fn(() => ({
          select: jest.fn().mockReturnThis(),
          filter: jest.fn().mockReturnThis(),
          top: jest.fn().mockReturnThis(),
          get: jest.fn().mockResolvedValue({ value: mockEvents }),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);
    });

    test('should return matching events', async () => {
      const results = await getEventCompletions('team');
      expect(results.length).toBeGreaterThanOrEqual(1);
      expect(results.some(r => r.value === 'event-1')).toBe(true);
    });

    test('should return empty array on auth error', async () => {
      const { isAuthError } = require('../../utils/response-helpers');
      isAuthError.mockReturnValue(true);
      getGraphClient.mockRejectedValue(new Error('Auth failed'));

      const results = await getEventCompletions('test');
      expect(results).toEqual([]);
    });
  });

  describe('registerCompletionCallbacks', () => {
    test('should log completion availability', () => {
      const mockServer = {};
      registerCompletionCallbacks(mockServer);
      expect(console.error).toHaveBeenCalledWith('Completion callbacks available for resource templates');
    });
  });
});
