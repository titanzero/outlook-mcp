// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListRuleTemplates, handleCreateRuleFromTemplate, handleExportRules, handleImportRules } = require('../../rules/templates');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Rule Templates & Import/Export Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      patch: jest.fn(),
      select: jest.fn(() => mockClient),
      orderby: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListRuleTemplates', () => {
    test('should list rule templates', async () => {
      const result = await handleListRuleTemplates({});

      expect(result.content[0].text).toContain('Rule Templates (10)');
      expect(result.content[0].text).toContain('Move emails from sender to folder');
      expect(result.content[0].text).toContain('Move emails with subject to folder');
      expect(result.content[0].text).toContain('Auto-reply to specific sender');
      expect(result.content[0].text).toContain('Flag high importance emails');
      expect(result.content[0].text).toContain('Categorize by sender domain');
      expect(result.content[0].text).toContain('Delete obvious spam');
      expect(result.content[0].text).toContain('Forward specific emails to team');
      expect(result.content[0].text).toContain('Mark newsletters as read');
      expect(result.content[0].text).toContain('Archive large/old emails');
      expect(result.content[0].text).toContain("Move CC'd emails to folder");
    });
  });

  describe('handleCreateRuleFromTemplate', () => {
    test('should create rule from move-sender-to-folder template', async () => {
      mockClient.post.mockResolvedValue({ id: 'rule-new', displayName: 'Move from sender@example.com to Folder', sequence: 10 });

      const result = await handleCreateRuleFromTemplate({
        templateId: 'move-sender-to-folder',
        parameters: { fromAddresses: 'sender@example.com', folderName: 'Folder' },
      });

      expect(result.content[0].text).toContain('Rule "Move from sender@example.com to Folder" created from template');
    });

    test('should return error for missing templateId', async () => {
      const result = await handleCreateRuleFromTemplate({ parameters: {} });

      expect(result.content[0].text).toBe('Template ID is required (from list-rule-templates).');
    });

    test('should return error for missing required parameters', async () => {
      const result = await handleCreateRuleFromTemplate({
        templateId: 'move-sender-to-folder',
        parameters: { folderName: 'Folder' },
      });

      expect(result.content[0].text).toContain('fromAddresses and folderName are required');
    });

    test('should return error for unknown template', async () => {
      const result = await handleCreateRuleFromTemplate({ templateId: 'unknown', parameters: {} });

      expect(result.content[0].text).toContain('Unknown template: unknown');
    });
  });

  describe('handleExportRules', () => {
    test('should export rules as JSON', async () => {
      const mockRules = [
        { id: 'rule-1', displayName: 'Rule 1', sequence: 1, isEnabled: true, conditions: {}, actions: {} },
      ];

      mockClient.get.mockResolvedValue({ value: mockRules });

      const result = await handleExportRules({});

      expect(result.content[0].text).toContain('Rules Export');
      expect(result.content[0].text).toContain('Total Rules: 1');
      expect(result.content[0].text).toContain('Rule 1 [seq: 1] ENABLED');
    });
  });

  describe('handleImportRules', () => {
    test('should import rules', async () => {
      const existingRules = { value: [{ id: 'existing-1', displayName: 'Existing Rule', sequence: 1 }] };
      const importRules = [
        { displayName: 'New Rule', sequence: 10, isEnabled: true, conditions: {}, actions: {} },
        { displayName: 'Existing Rule', sequence: 1, isEnabled: true, conditions: {}, actions: {} },
      ];

      mockClient.get.mockResolvedValue(existingRules);
      mockClient.post.mockResolvedValue({ id: 'new-rule-id' });

      const result = await handleImportRules({ rules: importRules, skipExisting: true });

      expect(result.content[0].text).toContain('Rules Import Results');
      expect(result.content[0].text).toContain('Created: 1');
      expect(result.content[0].text).toContain('Skipped: 1');
    });
  });
});