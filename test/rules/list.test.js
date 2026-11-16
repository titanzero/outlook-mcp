/**
 * Tests for rules list functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListRules, getInboxRules } = require('../../rules/list');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleListRules', () => {
  let mockClient;

  const mockRules = [
    {
      id: 'rule-1',
      displayName: 'Move GitHub notifications',
      isEnabled: true,
      sequence: 1,
      conditions: {
        fromAddresses: [{ emailAddress: { address: 'notifications@github.com' } }],
        subjectContains: ['GitHub'],
        hasAttachment: false,
        importance: 'normal',
      },
      actions: {
        moveToFolder: 'GitHub',
        markAsRead: true,
      },
    },
    {
      id: 'rule-2',
      displayName: 'Archive newsletters',
      isEnabled: false,
      sequence: 2,
      conditions: {
        fromAddresses: [],
        subjectContains: ['newsletter', 'weekly'],
        hasAttachment: false,
      },
      actions: {
        moveToFolder: 'Newsletters',
      },
    },
  ];

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: mockRules }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful rules listing', () => {
    test('should list rules with default simple format', async () => {
      const result = await handleListRules({});

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/inbox/messageRules');

      expect(result.content[0].text).toContain('Found 2 inbox rules (sorted by execution order):');
      expect(result.content[0].text).toContain('Move GitHub notifications - Sequence: 1');
      expect(result.content[0].text).toContain('Archive newsletters (Disabled) - Sequence: 2');
    });

    test('should list rules with detailed format when includeDetails is true', async () => {
      const result = await handleListRules({ includeDetails: true });

      expect(result.content[0].text).toContain('Move GitHub notifications - Sequence: 1');
      expect(result.content[0].text).toContain('Conditions: From: notifications@github.com; Subject contains: "GitHub"');
      expect(result.content[0].text).toContain('Actions: Move to folder: GitHub; Mark as read');
      expect(result.content[0].text).toContain('Archive newsletters (Disabled) - Sequence: 2');
      expect(result.content[0].text).toContain('Subject contains: "newsletter, weekly"');
      expect(result.content[0].text).toContain('Move to folder: Newsletters');
    });

    test('should sort rules by sequence', async () => {
      const unsortedRules = [
        { ...mockRules[0], sequence: 2, displayName: 'Rule B' },
        { ...mockRules[1], sequence: 1, displayName: 'Rule A' },
      ];
      mockClient = {
        api: jest.fn(() => ({
          get: jest.fn().mockResolvedValue({ value: unsortedRules }),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListRules({});

      const ruleAIndex = result.content[0].text.indexOf('Rule A');
      const ruleBIndex = result.content[0].text.indexOf('Rule B');
      expect(ruleAIndex).toBeLessThan(ruleBIndex);
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no rules found', async () => {
      mockClient = {
        api: jest.fn(() => ({
          get: jest.fn().mockResolvedValue({ value: [] }),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListRules({});

      expect(result.content[0].text).toContain('No inbox rules found');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleListRules({});

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockClient = {
        api: jest.fn(() => ({
          get: jest.fn().mockRejectedValue(new Error('Graph API Error')),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListRules({});

      expect(result.content[0].text).toBe('Error listing rules: Graph API Error');
    });
  });
});

describe('formatRuleConditions (via handleListRules)', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should format fromAddresses condition', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {
        fromAddresses: [
          { emailAddress: { address: 'test1@example.com' } },
          { emailAddress: { address: 'test2@example.com' } },
        ],
      },
      actions: {},
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('From: test1@example.com, test2@example.com');
  });

  test('should format subjectContains condition', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {
        subjectContains: ['urgent', 'important'],
      },
      actions: {},
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Subject contains: "urgent, important"');
  });

  test('should format bodyContains condition', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {
        bodyContains: ['meeting', 'deadline'],
      },
      actions: {},
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Body contains: "meeting, deadline"');
  });

  test('should format hasAttachment condition', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {
        hasAttachment: true,
      },
      actions: {},
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Has attachment');
  });

  test('should format importance condition', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {
        importance: 'high',
      },
      actions: {},
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Importance: high');
  });

  test('should combine multiple conditions', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {
        fromAddresses: [{ emailAddress: { address: 'test@example.com' } }],
        subjectContains: ['urgent'],
        hasAttachment: true,
        importance: 'high',
      },
      actions: {},
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('From: test@example.com');
    expect(result.content[0].text).toContain('Subject contains: "urgent"');
    expect(result.content[0].text).toContain('Has attachment');
    expect(result.content[0].text).toContain('Importance: high');
  });
});

describe('formatRuleActions (via handleListRules)', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should format moveToFolder action', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {},
      actions: {
        moveToFolder: 'Archive',
      },
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Move to folder: Archive');
  });

  test('should format copyToFolder action', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {},
      actions: {
        copyToFolder: 'Backup',
      },
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Copy to folder: Backup');
  });

  test('should format markAsRead action', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {},
      actions: {
        markAsRead: true,
      },
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Mark as read');
  });

  test('should format markImportance action', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {},
      actions: {
        markImportance: 'high',
      },
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Mark importance: high');
  });

  test('should format forwardTo action', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {},
      actions: {
        forwardTo: [
          { emailAddress: { address: 'user1@example.com' } },
          { emailAddress: { address: 'user2@example.com' } },
        ],
      },
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Forward to: user1@example.com, user2@example.com');
  });

  test('should format delete action', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {},
      actions: {
        delete: true,
      },
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Delete');
  });

  test('should combine multiple actions', async () => {
    const rule = {
      id: 'rule-1',
      displayName: 'Test Rule',
      isEnabled: true,
      sequence: 1,
      conditions: {},
      actions: {
        moveToFolder: 'Archive',
        markAsRead: true,
        markImportance: 'high',
      },
    };

    const mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [rule] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const result = await handleListRules({ includeDetails: true });

    expect(result.content[0].text).toContain('Move to folder: Archive; Mark as read; Mark importance: high');
  });
});

describe('getInboxRules', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: [{ id: 'rule-1' }] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should get inbox rules from Graph API', async () => {
    const rules = await getInboxRules(mockClient);

    expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/inbox/messageRules');
    expect(rules).toHaveLength(1);
    expect(rules[0].id).toBe('rule-1');
  });

  test('should return empty array when response has no value', async () => {
    mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({}),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const rules = await getInboxRules(mockClient);
    expect(rules).toEqual([]);
  });

  test('should return empty array when value is not an array', async () => {
    mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({ value: 'not an array' }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    const rules = await getInboxRules(mockClient);
    expect(rules).toEqual([]);
  });

  test('should throw error on Graph API failure', async () => {
    mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockRejectedValue(new Error('Graph API Error')),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    await expect(getInboxRules(mockClient)).rejects.toThrow('Graph API Error');
  });
});
