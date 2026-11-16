/**
 * Tests for email resources
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { registerEmailResources } = require('../../email/resources');
const { getGraphClient } = require('../../utils/graph-client');
const { resolveFolderPath } = require('../../email/folder-utils');

jest.mock('../../utils/graph-client');
jest.mock('../../email/folder-utils');
jest.mock('@modelcontextprotocol/sdk/server/mcp.js', () => ({
  ResourceTemplate: jest.fn().mockImplementation((uri, options) => ({ uri, options })),
}));

describe('registerEmailResources', () => {
  let mockServer;
  let mockClient;

  const mockEmails = [
    {
      id: 'email-1',
      subject: 'Test Email 1',
      from: { emailAddress: { name: 'John', address: 'john@example.com' } },
      toRecipients: [{ emailAddress: { address: 'me@example.com' } }],
      ccRecipients: [],
      receivedDateTime: '2024-01-15T10:30:00Z',
      bodyPreview: 'Preview 1',
      hasAttachments: false,
      importance: 'normal',
      isRead: true,
    },
    {
      id: 'email-2',
      subject: 'Test Email 2',
      from: { emailAddress: { name: 'Jane', address: 'jane@example.com' } },
      toRecipients: [{ emailAddress: { address: 'me@example.com' } }],
      ccRecipients: [],
      receivedDateTime: '2024-01-14T15:20:00Z',
      bodyPreview: 'Preview 2',
      hasAttachments: true,
      importance: 'high',
      isRead: false,
    },
  ];

  const mockEmailDetail = {
    id: 'email-1',
    subject: 'Test Email 1',
    from: { emailAddress: { name: 'John', address: 'john@example.com' } },
    toRecipients: [{ emailAddress: { address: 'me@example.com' } }],
    ccRecipients: [],
    bccRecipients: [],
    replyTo: [],
    receivedDateTime: '2024-01-15T10:30:00Z',
    sentDateTime: '2024-01-15T10:25:00Z',
    bodyPreview: 'Preview content',
    body: { content: 'Full body content', contentType: 'text' },
    hasAttachments: false,
    importance: 'normal',
    isRead: true,
    isDraft: false,
    conversationId: 'conv-1',
    conversationIndex: 'index-1',
    internetMessageId: 'msg-1',
    categories: ['Work'],
    flag: { flagStatus: 'flagged', startDateTime: { dateTime: '2024-01-15T10:30:00Z' } },
  };

  const mockDrafts = [
    {
      id: 'draft-1',
      subject: 'Draft Subject',
      toRecipients: [{ emailAddress: { address: 'to@example.com' } }],
      ccRecipients: [],
      bccRecipients: [],
      bodyPreview: 'Draft preview',
      body: { content: 'Draft body', contentType: 'text' },
      hasAttachments: false,
      importance: 'normal',
      createdDateTime: '2024-01-15T10:00:00Z',
      lastModifiedDateTime: '2024-01-15T10:30:00Z',
    },
  ];

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});

    mockClient = {
      api: jest.fn(() => ({
        top: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        orderby: jest.fn().mockReturnThis(),
        filter: jest.fn().mockReturnThis(),
        get: jest.fn().mockResolvedValue({ value: mockEmails }),
      })),
    };

    getGraphClient.mockResolvedValue(mockClient);
    resolveFolderPath.mockResolvedValue('me/mailFolders/inbox/messages');

    mockServer = {
      resource: jest.fn(),
    };
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('registerEmailResources', () => {
    test('should register 4 email resources', () => {
      registerEmailResources(mockServer);
      expect(mockServer.resource).toHaveBeenCalledTimes(4);
    });

    test('should register outlook-emails resource', () => {
      registerEmailResources(mockServer);
      const calls = mockServer.resource.mock.calls;
      const emailsCall = calls.find(c => c[0] === 'outlook-emails');
      expect(emailsCall).toBeDefined();
      expect(emailsCall[1]).toBe('outlook://emails');
    });

    test('should register outlook-email resource with template', () => {
      registerEmailResources(mockServer);
      const calls = mockServer.resource.mock.calls;
      const emailCall = calls.find(c => c[0] === 'outlook-email');
      expect(emailCall).toBeDefined();
      expect(emailCall[1]).toBeDefined(); // ResourceTemplate
    });

    test('should register outlook-drafts resource', () => {
      registerEmailResources(mockServer);
      const calls = mockServer.resource.mock.calls;
      const draftsCall = calls.find(c => c[0] === 'outlook-drafts');
      expect(draftsCall).toBeDefined();
      expect(draftsCall[1]).toBe('outlook://drafts');
    });

    test('should register outlook-draft resource', () => {
      registerEmailResources(mockServer);
      const calls = mockServer.resource.mock.calls;
      const draftCall = calls.find(c => c[0] === 'outlook-draft');
      expect(draftCall).toBeDefined();
      expect(draftCall[1]).toBe('outlook://drafts/{draftId}');
    });
  });
});
