// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleCreateAdvancedRule, handleUpdateAdvancedRule, handleGetRule, handleTestRule, handleBulkEnableRules } = require('../../rules/advanced');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Advanced Rules Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      patch: jest.fn(),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleCreateAdvancedRule', () => {
    test('should create advanced rule with full conditions and actions', async () => {
      mockClient.post.mockResolvedValue({ id: 'rule-new', displayName: 'Advanced Rule', sequence: 10 });

      const result = await handleCreateAdvancedRule({
        name: 'Advanced Rule',
        fromAddresses: 'sender@example.com',
        containsSubject: 'Important',
        importance: 'high',
        hasAttachments: true,
        moveToFolder: 'Important',
        markAsRead: true,
        assignCategories: 'Priority,Urgent',
      });

      expect(result.content[0].text).toContain('Advanced rule "Advanced Rule" created successfully');
      expect(result.content[0].text).toContain('ID: rule-new');
    });

    test('should return error when name is missing', async () => {
      const result = await handleCreateAdvancedRule({});

      expect(result.content[0].text).toBe('Rule name is required.');
    });
  });

  describe('handleUpdateAdvancedRule', () => {
    test('should update advanced rule', async () => {
      const mockRule = { id: 'rule-1', displayName: 'Test Rule', conditions: {}, actions: {} };

      mockClient.get.mockResolvedValue({ value: [mockRule] });
      mockClient.patch.mockResolvedValue({});

      const result = await handleUpdateAdvancedRule({
        ruleName: 'Test Rule',
        isEnabled: false,
        moveToFolder: 'Archive',
        clearConditions: true,
      });

      expect(result.content[0].text).toBe('Advanced rule "Test Rule" updated successfully.');
    });
  });

  describe('handleGetRule', () => {
    test('should get rule details', async () => {
      const mockRule = {
        id: 'rule-1',
        displayName: 'Test Rule',
        sequence: 5,
        isEnabled: true,
        conditions: { fromAddresses: [{ emailAddress: { address: 'sender@example.com' } }], subjectContains: ['Important'] },
        actions: { moveToFolder: 'Important', markAsRead: true },
      };

      mockClient.get.mockResolvedValue({ value: [mockRule] });

      const result = await handleGetRule({ ruleName: 'Test Rule' });

      expect(result.content[0].text).toContain('Rule: Test Rule');
      expect(result.content[0].text).toContain('Sequence: 5');
      expect(result.content[0].text).toContain('Enabled: Yes');
    });
  });

  describe('handleTestRule', () => {
    test('should test rule against email', async () => {
      const mockRule = {
        id: 'rule-1',
        displayName: 'Test Rule',
        conditions: { fromAddresses: [{ emailAddress: { address: 'sender@example.com' } }], subjectContains: ['Report'] },
        actions: { moveToFolder: 'Reports' },
      };
      const mockEmail = {
        id: 'email-1',
        subject: 'Monthly Report',
        from: { emailAddress: { address: 'sender@example.com', name: 'Sender' } },
        toRecipients: [{ emailAddress: { address: 'me@example.com' } }],
        ccRecipients: [],
        receivedDateTime: '2024-01-15T10:00:00Z',
      };

      const mockApi = { select: jest.fn().mockReturnThis(), get: jest.fn() };
      mockApi.get
        .mockResolvedValueOnce({ value: [mockRule] })
        .mockResolvedValueOnce(mockEmail);
      mockClient.api.mockReturnValue(mockApi);

      const result = await handleTestRule({ ruleName: 'Test Rule', emailId: 'email-1' });

      expect(result.content[0].text).toContain('Rule Test: "Test Rule"');
      expect(result.content[0].text).toContain('Would Match: YES');
      expect(result.content[0].text).toContain('fromAddresses');
      expect(result.content[0].text).toContain('subjectContains');
    });
  });

  describe('handleBulkEnableRules', () => {
    test('should bulk enable rules', async () => {
      const mockRules = [
        { id: 'rule-1', displayName: 'Rule 1' },
        { id: 'rule-2', displayName: 'Rule 2' },
      ];

      mockClient.get.mockResolvedValue({ value: mockRules });
      mockClient.patch.mockResolvedValue({});

      const result = await handleBulkEnableRules({ ruleNames: ['Rule 1', 'Rule 2'], isEnabled: false });

      expect(result.content[0].text).toContain('Bulk Disabled Rules');
      expect(result.content[0].text).toContain('Successful: 2');
    });
  });
});