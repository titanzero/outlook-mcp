// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListConversations, handleGetConversation } = require('../../email/conversations');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Conversations/Thread View Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
      orderby: jest.fn(() => mockClient),
      filter: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListConversations', () => {
    test('should group messages by conversationId', async () => {
      const mockMessages = [
        { id: 'msg-1', conversationId: 'conv-1', subject: 'Thread 1', from: { emailAddress: { name: 'John', address: 'john@example.com' } }, toRecipients: [{ emailAddress: { address: 'me@example.com' } }], receivedDateTime: '2024-01-15T10:30:00Z', isRead: false, hasAttachments: true, importance: 'high', bodyPreview: 'First message' },
        { id: 'msg-2', conversationId: 'conv-1', subject: 'Thread 1', from: { emailAddress: { name: 'Jane', address: 'jane@example.com' } }, toRecipients: [{ emailAddress: { address: 'me@example.com' } }], receivedDateTime: '2024-01-15T11:00:00Z', isRead: true, hasAttachments: false, importance: 'normal', bodyPreview: 'Reply' },
        { id: 'msg-3', conversationId: 'conv-2', subject: 'Thread 2', from: { emailAddress: { name: 'Bob', address: 'bob@example.com' } }, toRecipients: [{ emailAddress: { address: 'me@example.com' } }], receivedDateTime: '2024-01-14T09:00:00Z', isRead: false, hasAttachments: false, importance: 'normal', bodyPreview: 'Another thread' },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockMessages });

      const result = await handleListConversations({ folder: 'inbox', count: 20 });

      expect(result.content[0].text).toContain('Conversations in inbox (2)');
      expect(result.content[0].text).toContain('Thread 1');
      expect(result.content[0].text).toContain('Thread 2');
      expect(result.content[0].text).toContain('[UNREAD]');
      expect(result.content[0].text).toContain('📎');
    });

    test('should handle empty results', async () => {
      graphGetPaginated.mockResolvedValue({ value: [] });

      const result = await handleListConversations({});

      expect(result.content[0].text).toContain('No conversations found.');
    });
  });

  describe('handleGetConversation', () => {
    test('should get full conversation thread', async () => {
      const mockMessages = [
        { id: 'msg-1', conversationId: 'conv-1', subject: 'Thread 1', from: { emailAddress: { name: 'John', address: 'john@example.com' } }, toRecipients: [{ emailAddress: { address: 'me@example.com' } }], ccRecipients: [], bccRecipients: [], receivedDateTime: '2024-01-15T10:30:00Z', sentDateTime: null, bodyPreview: 'First message', body: { content: 'Full body', contentType: 'text' }, hasAttachments: false, importance: 'normal', isRead: false, isDraft: false, categories: [], flag: null, to: ['me@example.com'], cc: [] },
        { id: 'msg-2', conversationId: 'conv-1', subject: 'Thread 1', from: { emailAddress: { name: 'Me', address: 'me@example.com' } }, toRecipients: [{ emailAddress: { address: 'john@example.com' } }], ccRecipients: [], bccRecipients: [], receivedDateTime: null, sentDateTime: '2024-01-15T11:00:00Z', bodyPreview: 'Reply', body: { content: 'Reply body', contentType: 'text' }, hasAttachments: false, importance: 'normal', isRead: true, isDraft: false, categories: [], flag: null, to: ['john@example.com'], cc: [] },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockMessages });

      const result = await handleGetConversation({ conversationId: 'conv-1', count: 50 });

      expect(result.content[0].text).toContain('Conversation Thread');
      expect(result.content[0].text).toContain('Thread 1');
    });

    test('should return error when conversationId is missing', async () => {
      const result = await handleGetConversation({});

      expect(result.content[0].text).toBe('Conversation ID is required (from list-conversations).');
    });
  });
});