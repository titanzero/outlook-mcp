/**
 * Tests for rules delete functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleDeleteRule } = require('../../rules/index');
const { getGraphClient } = require('../../utils/graph-client');
const { getInboxRules } = require('../../rules/list');

jest.mock('../../utils/graph-client');
jest.mock('../../rules/list', () => ({
  getInboxRules: jest.fn(),
}));

describe('handleDeleteRule', () => {
  let mockClient;
  let mockDelete;
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
    mockDelete = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn(() => ({
        delete: mockDelete,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
    mockGetInboxRules = getInboxRules;
    mockGetInboxRules.mockResolvedValue([existingRule]);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful rule deletion', () => {
    test('should delete rule', async () => {
      const result = await handleDeleteRule({ ruleName: 'Test Rule' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockGetInboxRules).toHaveBeenCalledWith(mockClient);
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/inbox/messageRules/rule-123');
      expect(mockDelete).toHaveBeenCalledTimes(1);
      expect(result.content[0].text).toBe('Rule "Test Rule" deleted successfully.');
    });
  });

  describe('validation', () => {
    test('should return error when ruleName is missing', async () => {
      const result = await handleDeleteRule({});

      expect(result.content[0].text).toBe('Rule name is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when ruleName is empty string', async () => {
      const result = await handleDeleteRule({ ruleName: '' });

      expect(result.content[0].text).toBe('Rule name is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleDeleteRule({ ruleName: 'Test Rule' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should return error when rule not found', async () => {
      mockGetInboxRules.mockResolvedValue([]);

      const result = await handleDeleteRule({ ruleName: 'Non-existent' });

      expect(result.content[0].text).toBe('Rule "Non-existent" not found.');
    });

    test('should handle Graph API error', async () => {
      mockDelete.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleDeleteRule({ ruleName: 'Test Rule' });

      expect(result.content[0].text).toBe('Error deleting rule: Graph API Error');
    });
  });
});
