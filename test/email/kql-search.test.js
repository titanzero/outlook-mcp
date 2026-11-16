// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleKqlSearch, handleKqlHelp } = require('../../email/kql-search');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('KQL Search Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleKqlSearch', () => {
    test('should execute KQL search and return results', async () => {
      const mockEmails = [
        { id: 'email-1', subject: 'Quarterly Report', from: { emailAddress: { name: 'John', address: 'john@example.com' } }, toRecipients: [], ccRecipients: [], receivedDateTime: '2024-01-15T10:30:00Z', bodyPreview: 'Q4 report attached', hasAttachments: true, importance: 'high', isRead: false, categories: ['Reports'], flag: { flagStatus: 'flagged', dueDateTime: { dateTime: '2024-01-20T00:00:00Z' } } },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      const result = await handleKqlSearch({ query: 'from:john subject:report hasattachment:true', folder: 'inbox', count: 50 });

      expect(result.content[0].text).toContain('KQL Search Results');
      expect(result.content[0].text).toContain('Query: from:john subject:report hasattachment:true');
      expect(result.content[0].text).toContain('Quarterly Report');
      expect(result.content[0].text).toContain('[Reports]');
      expect(result.content[0].text).toContain('[flagged]');
    });

    test('should return error when query is missing', async () => {
      const result = await handleKqlSearch({});

      expect(result.content[0].text).toContain('KQL query is required');
    });

    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleKqlSearch({ query: 'test' });

      expect(result.content[0].text).toBe('Authentication required');
    });
  });

  describe('handleKqlHelp', () => {
    test('should return KQL syntax reference', async () => {
      const result = await handleKqlHelp({});

      expect(result.content[0].text).toContain('KQL (Keyword Query Language) Syntax Reference');
      expect(result.content[0].text).toContain('PROPERTY RESTRICTIONS');
      expect(result.content[0].text).toContain('from:');
      expect(result.content[0].text).toContain('subject:');
      expect(result.content[0].text).toContain('hasattachment:');
      expect(result.content[0].text).toContain('importance:');
      expect(result.content[0].text).toContain('LOGICAL OPERATORS');
      expect(result.content[0].text).toContain('EXAMPLES');
      expect(result.content[0].text).toContain('DATE FORMATS');
    });
  });
});