/**
 * Tests for email reply functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleReplyEmail = require('../../email/reply');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleReplyEmail', () => {
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

  describe('successful replies', () => {
    test('should reply to email without comment', async () => {
      const result = await handleReplyEmail({ id: 'email-123' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-123/reply');
      expect(mockPost).toHaveBeenCalledWith({ comment: '' });

      expect(result.content[0].text).toBe('Email replied successfully.');
    });

    test('should reply to email with comment', async () => {
      const result = await handleReplyEmail({
        id: 'email-123',
        comment: 'Thanks for the email!',
      });

      expect(mockPost).toHaveBeenCalledWith({ comment: 'Thanks for the email!' });
      expect(result.content[0].text).toBe('Email replied successfully.');
    });

    test('should reply all when replyAll is true', async () => {
      const result = await handleReplyEmail({
        id: 'email-123',
        replyAll: true,
        comment: 'Reply to all',
      });

      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-123/replyAll');
      expect(result.content[0].text).toBe('Email replied to all successfully.');
    });

    test('should use default replyAll false', async () => {
      await handleReplyEmail({ id: 'email-123', comment: 'Test' });

      expect(mockClient.api).toHaveBeenCalledWith('me/messages/email-123/reply');
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleReplyEmail({});

      expect(result.content[0].text).toBe('Email ID is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when id is empty string', async () => {
      const result = await handleReplyEmail({ id: '' });

      expect(result.content[0].text).toBe('Email ID is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleReplyEmail({ id: 'email-123' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleReplyEmail({ id: 'email-123' });

      expect(result.content[0].text).toBe('Error replying to email: Graph API Error');
    });
  });
});
