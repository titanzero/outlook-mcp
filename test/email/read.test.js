/**
 * Tests for email read functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleReadEmail = require('../../email/read');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');
jest.mock('../../config', () => ({
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,importance,bodyPreview,hasAttachments',
  EMAIL_FULL_BODY_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,importance,bodyPreview,hasAttachments,body',
  DEFAULT_TIMEZONE: 'Central European Standard Time',
}));

describe('handleReadEmail', () => {
  let mockClient;
  const mockEmail = {
    id: 'email-123',
    subject: 'Test Subject',
    from: {
      emailAddress: { name: 'John Doe', address: 'john@example.com' },
    },
    toRecipients: [
      { emailAddress: { name: 'Jane Smith', address: 'jane@example.com' } },
    ],
    ccRecipients: [
      { emailAddress: { name: 'Bob Wilson', address: 'bob@example.com' } },
    ],
    bccRecipients: [],
    receivedDateTime: '2024-01-15T10:30:00Z',
    importance: 'high',
    bodyPreview: 'This is a preview of the email body.',
    hasAttachments: true,
    body: {
      contentType: 'html',
      content: '<html><body><p>This is the full HTML body content.</p></body></html>',
    },
  };

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => ({
        select: jest.fn(() => ({
          get: jest.fn().mockResolvedValue(mockEmail),
        })),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful email reading', () => {
    test('should read email with preview by default', async () => {
      const result = await handleReadEmail({ id: 'email-123' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-123');
      expect(result.content[0].text).toContain('Test Subject');
      expect(result.content[0].text).toContain('John Doe (john@example.com)');
      expect(result.content[0].text).toContain('Jane Smith (jane@example.com)');
      expect(result.content[0].text).toContain('Bob Wilson (bob@example.com)');
      expect(result.content[0].text).toContain('This is a preview of the email body.');
      expect(result.content[0].text).toContain('(Preview — use fullBody=true for complete content)');
    });

    test('should read email with full body when requested', async () => {
      const result = await handleReadEmail({ id: 'email-123', fullBody: true });

      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-123');
      expect(result.content[0].text).toContain('This is the full HTML body content.');
      expect(result.content[0].text).not.toContain('Preview — use fullBody=true');
    });

    test('should handle email without sender', async () => {
      const emailNoSender = { ...mockEmail, from: null };
      mockClient.api.mockReturnValueOnce({
        select: jest.fn(() => ({
          get: jest.fn().mockResolvedValue(emailNoSender),
        })),
      });

      const result = await handleReadEmail({ id: 'email-123' });

      expect(result.content[0].text).toContain('Unknown');
    });

    test('should handle email without CC/BCC recipients', async () => {
      const emailNoCcBcc = { ...mockEmail, ccRecipients: [], bccRecipients: [] };
      mockClient.api.mockReturnValueOnce({
        select: jest.fn(() => ({
          get: jest.fn().mockResolvedValue(emailNoCcBcc),
        })),
      });

      const result = await handleReadEmail({ id: 'email-123' });

      expect(result.content[0].text).not.toContain('CC:');
      expect(result.content[0].text).not.toContain('BCC:');
    });

    test('should handle plain text body when fullBody requested', async () => {
      const emailPlainText = {
        ...mockEmail,
        body: {
          contentType: 'text',
          content: 'Plain text body content.',
        },
      };
      mockClient.api.mockReturnValueOnce({
        select: jest.fn(() => ({
          get: jest.fn().mockResolvedValue(emailPlainText),
        })),
      });

      const result = await handleReadEmail({ id: 'email-123', fullBody: true });

      expect(result.content[0].text).toContain('Plain text body content.');
    });

    test('should handle email without bodyPreview', async () => {
      const emailNoPreview = { ...mockEmail, bodyPreview: null };
      mockClient.api.mockReturnValueOnce({
        select: jest.fn(() => ({
          get: jest.fn().mockResolvedValue(emailNoPreview),
        })),
      });

      const result = await handleReadEmail({ id: 'email-123' });

      expect(result.content[0].text).toContain('No content');
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleReadEmail({});

      expect(result.content[0].text).toBe('Email ID is required.');
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when id is empty string', async () => {
      const result = await handleReadEmail({ id: '' });

      expect(result.content[0].text).toBe('Email ID is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleReadEmail({ id: 'email-123' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle email not found error', async () => {
      const notFoundError = new Error("doesn't belong to the targeted mailbox");
      mockClient.api.mockReturnValueOnce({
        select: jest.fn(() => ({
          get: jest.fn().mockRejectedValue(notFoundError),
        })),
      });

      const result = await handleReadEmail({ id: 'invalid-id' });

      expect(result.content[0].text).toContain("doesn't belong to your mailbox");
    });

    test('should handle generic Graph API error', async () => {
      const apiError = new Error('Graph API Error');
      mockClient.api.mockReturnValueOnce({
        select: jest.fn(() => ({
          get: jest.fn().mockRejectedValue(apiError),
        })),
      });

      const result = await handleReadEmail({ id: 'email-123' });

      expect(result.content[0].text).toContain('Failed to read email: Graph API Error');
    });

    test('should handle getGraphClient error', async () => {
      const clientError = new Error('Client creation failed');
      getGraphClient.mockRejectedValue(clientError);

      const result = await handleReadEmail({ id: 'email-123' });

      expect(result.content[0].text).toContain('Error accessing email: Client creation failed');
    });
  });
});
