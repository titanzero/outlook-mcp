/**
 * Tests for rules edit-sequence functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleEditRuleSequence } = require('../../rules/index');
const { getGraphClient } = require('../../utils/graph-client');
const { getInboxRules } = require('../../rules/list');

jest.mock('../../utils/graph-client');
jest.mock('../../rules/list', () => ({
  getInboxRules: jest.fn(),
}));

describe('handleEditRuleSequence', () => {
  let mockClient;
  let mockPatch;
  let mockGetInboxRules;

  const existingRule = {
    id: 'rule-123',
    displayName: 'Test Rule',
    isEnabled: true,
    sequence: 5,
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

  describe('successful sequence update', () => {
    test('should update rule sequence', async () => {
      const result = await handleEditRuleSequence({
        ruleName: 'Test Rule',
        sequence: 10,
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockGetInboxRules).toHaveBeenCalledWith(mockClient);
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/inbox/messageRules/rule-123');
      expect(mockPatch).toHaveBeenCalledWith({ sequence: 10 });
      expect(result.content[0].text).toBe('Successfully updated the sequence of rule "Test Rule" to 10.');
    });

    test('should update rule to sequence 1 (highest priority)', async () => {
      await handleEditRuleSequence({
        ruleName: 'Test Rule',
        sequence: 1,
      });

      expect(mockPatch).toHaveBeenCalledWith({ sequence: 1 });
    });

    test('should update rule to higher sequence (lower priority)', async () => {
      await handleEditRuleSequence({
        ruleName: 'Test Rule',
        sequence: 100,
      });

      expect(mockPatch).toHaveBeenCalledWith({ sequence: 100 });
    });
  });

  describe('validation', () => {
    test('should return error when ruleName is missing', async () => {
      const result = await handleEditRuleSequence({ sequence: 10 });

      expect(result.content[0].text).toBe('Rule name is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when sequence is missing', async () => {
      const result = await handleEditRuleSequence({ ruleName: 'Test Rule' });

      expect(result.content[0].text).toContain('A positive sequence number is required');
      expect(result.isError).toBe(true);
    });

    test('should return error when sequence is not a number', async () => {
      const result = await handleEditRuleSequence({ ruleName: 'Test Rule', sequence: 'abc' });

      expect(result.content[0].text).toContain('A positive sequence number is required');
    });

    test('should return error when sequence is zero', async () => {
      const result = await handleEditRuleSequence({ ruleName: 'Test Rule', sequence: 0 });

      expect(result.content[0].text).toContain('A positive sequence number is required');
    });

    test('should return error when sequence is negative', async () => {
      const result = await handleEditRuleSequence({ ruleName: 'Test Rule', sequence: -5 });

      expect(result.content[0].text).toContain('A positive sequence number is required');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleEditRuleSequence({ ruleName: 'Test Rule', sequence: 10 });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should return error when rule not found', async () => {
      mockGetInboxRules.mockResolvedValue([]);

      const result = await handleEditRuleSequence({ ruleName: 'Non-existent', sequence: 10 });

      expect(result.content[0].text).toBe('Rule "Non-existent" not found.');
    });

    test('should handle Graph API error', async () => {
      mockPatch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleEditRuleSequence({ ruleName: 'Test Rule', sequence: 10 });

      expect(result.content[0].text).toBe('Error updating rule sequence: Graph API Error');
    });
  });
});
