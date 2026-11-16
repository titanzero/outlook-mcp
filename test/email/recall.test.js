// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleRecallMessage, handleGetRecallStatus } = require('../../email/recall');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Message Recall Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      filter: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
      orderby: jest.fn(() => mockClient),
      top: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleRecallMessage', () => {
    test('should recall a sent message', async () => {
      const mockMessage = { id: 'msg-1', subject: 'Test Email', sentDateTime: '2024-01-15T10:30:00Z', isDraft: false, toRecipients: [] };

      mockClient.get.mockResolvedValue(mockMessage);
      mockClient.post.mockResolvedValue({});

      const result = await handleRecallMessage({ messageId: 'msg-1' });

      expect(result.content[0].text).toContain('Message Recall Initiated');
      expect(result.content[0].text).toContain('Test Email');
      expect(result.content[0].text).toContain('Recall request has been sent');
    });

    test('should return error when messageId is missing', async () => {
      const result = await handleRecallMessage({});

      expect(result.content[0].text).toBe('Message ID is required (from sent items).');
    });

    test('should return error for draft message', async () => {
      const mockMessage = { id: 'msg-1', subject: 'Draft', isDraft: true };

      mockClient.get.mockResolvedValue(mockMessage);

      const result = await handleRecallMessage({ messageId: 'msg-1' });

      expect(result.content[0].text).toBe('Cannot recall a draft message. Only sent messages can be recalled.');
    });

    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleRecallMessage({ messageId: 'msg-1' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error for unsupported recall', async () => {
      const mockMessage = { id: 'msg-1', subject: 'Test', sentDateTime: '2024-01-15T10:30:00Z', isDraft: false };
      mockClient.get.mockResolvedValue(mockMessage);
      const error = new Error('Bad Request');
      error.statusCode = 400;
      mockClient.post.mockRejectedValue(error);

      const result = await handleRecallMessage({ messageId: 'msg-1' });

      expect(result.content[0].text).toContain('Cannot recall this message');
    });
  });

  describe('handleGetRecallStatus', () => {
    test('should get recall status notifications', async () => {
      const mockNotifications = [
        { id: 'notif-1', subject: 'Message Recall: msg-1', receivedDateTime: '2024-01-15T10:35:00Z', bodyPreview: 'Recall successful', isRead: false },
      ];

      mockClient.get.mockResolvedValue({ value: mockNotifications });

      const result = await handleGetRecallStatus({ messageId: 'msg-1' });

      expect(result.content[0].text).toContain('Recall Status for Message: msg-1');
      expect(result.content[0].text).toContain('Message Recall: msg-1');
    });

    test('should return error when messageId is missing', async () => {
      const result = await handleGetRecallStatus({});

      expect(result.content[0].text).toBe('Message ID is required.');
    });
  });
});