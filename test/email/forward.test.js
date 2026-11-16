/**
 * Tests for email forward functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleForwardEmail = require('../../email/forward');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleForwardEmail', () => {
  let mockClient;
  let mockPost;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPost = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn(() => ({
        post: mockPost,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful forwards', () => {
    test('should forward email to single recipient', async () => {
      const result = await handleForwardEmail({
        id: 'email-123',
        to: 'recipient@example.com',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-123/forward');
      expect(mockPost).toHaveBeenCalledWith({
        toRecipients: [{ emailAddress: { address: 'recipient@example.com' } }],
        comment: '',
      });

      expect(result.content[0].text).toBe('Email forwarded successfully to recipient@example.com.');
    });

    test('should forward email to multiple recipients', async () => {
      const result = await handleForwardEmail({
        id: 'email-123',
        to: 'recipient1@example.com, recipient2@example.com',
        comment: 'Please review',
      });

      expect(mockPost).toHaveBeenCalledWith({
        toRecipients: [
          { emailAddress: { address: 'recipient1@example.com' } },
          { emailAddress: { address: 'recipient2@example.com' } },
        ],
        comment: 'Please review',
      });

      expect(result.content[0].text).toBe(
        'Email forwarded successfully to recipient1@example.com, recipient2@example.com.'
      );
    });

    test('should forward email without comment', async () => {
      const result = await handleForwardEmail({
        id: 'email-123',
        to: 'recipient@example.com',
      });

      expect(mockPost).toHaveBeenCalledWith({
        toRecipients: [{ emailAddress: { address: 'recipient@example.com' } }],
        comment: '',
      });
      expect(result.content[0].text).toContain('Email forwarded successfully');
    });

    test('should trim whitespace from recipient emails', async () => {
      await handleForwardEmail({
        id: 'email-123',
        to: '  test1@example.com , test2@example.com  ',
      });

      expect(mockPost).toHaveBeenCalledWith({
        toRecipients: [
          { emailAddress: { address: 'test1@example.com' } },
          { emailAddress: { address: 'test2@example.com' } },
        ],
        comment: '',
      });
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleForwardEmail({});

      expect(result.content[0].text).toBe('Email ID is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when to is missing', async () => {
      const result = await handleForwardEmail({ id: 'email-123' });

      expect(result.content[0].text).toBe('At least one recipient (to) is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when to is empty string', async () => {
      const result = await handleForwardEmail({ id: 'email-123', to: '' });

      expect(result.content[0].text).toBe('At least one recipient (to) is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleForwardEmail({
        id: 'email-123',
        to: 'test@example.com',
      });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleForwardEmail({
        id: 'email-123',
        to: 'test@example.com',
      });

      expect(result.content[0].text).toBe('Error forwarding email: Graph API Error');
    });
  });
});
