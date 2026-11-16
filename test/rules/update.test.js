/**
 * Tests for rules update functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleUpdateRule } = require('../../rules/index');
const { getGraphClient } = require('../../utils/graph-client');
const { getInboxRules } = require('../../rules/list');

jest.mock('../../utils/graph-client');
jest.mock('../../rules/list', () => ({
  getInboxRules: jest.fn(),
}));

describe('handleUpdateRule', () => {
  let mockClient;
  let mockPatch;
  let mockGetInboxRules;

  const existingRule = {
    id: 'rule-123',
    displayName: 'Test Rule',
    isEnabled: true,
    sequence: 5,
    conditions: {
      fromAddresses: [{ emailAddress: { address: 'old@example.com' } }],
      subjectContains: ['old subject'],
    },
    actions: {
      moveToFolder: 'OldFolder',
      markAsRead: false,
    },
  };

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
    mockGetInboxRules = getInboxRules;
    mockGetInboxRules.mockResolvedValue([existingRule]);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful rule updates', () => {
    test('should update rule isEnabled', async () => {
      const result = await handleUpdateRule({
        ruleName: 'Test Rule',
        isEnabled: false,
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockGetInboxRules).toHaveBeenCalledWith(mockClient);
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/inbox/messageRules/rule-123');
      expect(mockPatch).toHaveBeenCalledWith({ isEnabled: false });
      expect(result.content[0].text).toBe('Rule "Test Rule" updated successfully.');
    });

    test('should update rule sequence', async () => {
      await handleUpdateRule({
        ruleName: 'Test Rule',
        sequence: 10,
      });

      expect(mockPatch).toHaveBeenCalledWith({ sequence: 10 });
    });

    test('should update rule fromAddresses', async () => {
      await handleUpdateRule({
        ruleName: 'Test Rule',
        fromAddresses: 'new1@example.com, new2@example.com',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        conditions: {
          fromAddresses: [
            { emailAddress: { address: 'new1@example.com', name: 'new1@example.com' } },
            { emailAddress: { address: 'new2@example.com', name: 'new2@example.com' } },
          ],
          subjectContains: ['old subject'],
        },
      });
    });

    test('should update rule containsSubject', async () => {
      await handleUpdateRule({
        ruleName: 'Test Rule',
        containsSubject: 'new subject filter',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        conditions: {
          fromAddresses: [{ emailAddress: { address: 'old@example.com' } }],
          subjectContains: ['new subject filter'],
        },
      });
    });

    test('should update rule moveToFolder', async () => {
      await handleUpdateRule({
        ruleName: 'Test Rule',
        moveToFolder: 'NewFolder',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        actions: {
          moveToFolder: 'NewFolder',
          markAsRead: false,
        },
      });
    });

    test('should update rule markAsRead', async () => {
      await handleUpdateRule({
        ruleName: 'Test Rule',
        markAsRead: true,
      });

      expect(mockPatch).toHaveBeenCalledWith({
        actions: {
          moveToFolder: 'OldFolder',
          markAsRead: true,
        },
      });
    });

    test('should update multiple fields at once', async () => {
      await handleUpdateRule({
        ruleName: 'Test Rule',
        isEnabled: false,
        sequence: 15,
        fromAddresses: 'updated@example.com',
        containsSubject: 'updated subject',
        moveToFolder: 'UpdatedFolder',
        markAsRead: true,
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.isEnabled).toBe(false);
      expect(patch.sequence).toBe(15);
      expect(patch.conditions.fromAddresses).toHaveLength(1);
      expect(patch.conditions.subjectContains).toEqual(['updated subject']);
      expect(patch.actions.moveToFolder).toBe('UpdatedFolder');
      expect(patch.actions.markAsRead).toBe(true);
    });
  });

  describe('validation', () => {
    test('should return error when ruleName is missing', async () => {
      const result = await handleUpdateRule({ isEnabled: false });

      expect(result.content[0].text).toBe('Rule name is required.');
      expect(result.isError).toBe(true);
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleUpdateRule({ ruleName: 'Test Rule', isEnabled: false });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should return error when rule not found', async () => {
      mockGetInboxRules.mockResolvedValue([]);

      const result = await handleUpdateRule({ ruleName: 'Non-existent', isEnabled: false });

      expect(result.content[0].text).toBe('Rule "Non-existent" not found.');
    });

    test('should handle Graph API error', async () => {
      mockPatch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleUpdateRule({ ruleName: 'Test Rule', isEnabled: false });

      expect(result.content[0].text).toBe('Error updating rule: Graph API Error');
    });
  });
});
