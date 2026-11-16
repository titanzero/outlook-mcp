/**
 * Tests for email send functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleSendEmail = require('../../email/send');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleSendEmail', () => {
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

  describe('successful email sending', () => {
    test('should send email with required parameters only', async () => {
      const result = await handleSendEmail({
        to: 'recipient@example.com',
        subject: 'Test Subject',
        body: 'Test body content',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/sendMail');
      expect(mockPost).toHaveBeenCalledTimes(1);

      const emailObject = mockPost.mock.calls[0][0];
      expect(emailObject.message.subject).toBe('Test Subject');
      expect(emailObject.message.body.content).toBe('Test body content');
      expect(emailObject.message.body.contentType).toBe('text');
      expect(emailObject.message.toRecipients).toHaveLength(1);
      expect(emailObject.message.toRecipients[0].emailAddress.address).toBe('recipient@example.com');
      expect(emailObject.message.importance).toBe('normal');
      expect(emailObject.saveToSentItems).toBe(true);

      expect(result.content[0].text).toContain('Email sent successfully!');
      expect(result.content[0].text).toContain('Test Subject');
      expect(result.content[0].text).toContain('Recipients: 1');
    });

    test('should send email with CC and BCC', async () => {
      const result = await handleSendEmail({
        to: 'to1@example.com, to2@example.com',
        cc: 'cc1@example.com',
        bcc: 'bcc1@example.com, bcc2@example.com',
        subject: 'Test with CC/BCC',
        body: 'Body content',
      });

      const emailObject = mockPost.mock.calls[0][0];
      expect(emailObject.message.toRecipients).toHaveLength(2);
      expect(emailObject.message.ccRecipients).toHaveLength(1);
      expect(emailObject.message.bccRecipients).toHaveLength(2);

      expect(result.content[0].text).toContain('Recipients: 2 + 1 CC + 2 BCC');
    });

    test('should detect HTML content and set contentType to html', async () => {
      await handleSendEmail({
        to: 'test@example.com',
        subject: 'HTML Email',
        body: '<html><body><p>HTML content</p></body></html>',
      });

      const emailObject = mockPost.mock.calls[0][0];
      expect(emailObject.message.body.contentType).toBe('html');
    });

    test('should use custom importance', async () => {
      await handleSendEmail({
        to: 'test@example.com',
        subject: 'High Importance',
        body: 'Body',
        importance: 'high',
      });

      const emailObject = mockPost.mock.calls[0][0];
      expect(emailObject.message.importance).toBe('high');
    });

    test('should respect saveToSentItems false', async () => {
      await handleSendEmail({
        to: 'test@example.com',
        subject: 'No Save',
        body: 'Body',
        saveToSentItems: false,
      });

      const emailObject = mockPost.mock.calls[0][0];
      expect(emailObject.saveToSentItems).toBe(false);
    });

    test('should trim whitespace from recipient emails', async () => {
      await handleSendEmail({
        to: '  test1@example.com , test2@example.com  ',
        subject: 'Test',
        body: 'Body',
      });

      const emailObject = mockPost.mock.calls[0][0];
      expect(emailObject.message.toRecipients[0].emailAddress.address).toBe('test1@example.com');
      expect(emailObject.message.toRecipients[1].emailAddress.address).toBe('test2@example.com');
    });
  });

  describe('validation', () => {
    test('should return error when to is missing', async () => {
      const result = await handleSendEmail({
        subject: 'Test',
        body: 'Body',
      });

      expect(result.content[0].text).toBe('Recipient (to) is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when subject is missing', async () => {
      const result = await handleSendEmail({
        to: 'test@example.com',
        body: 'Body',
      });

      expect(result.content[0].text).toBe('Subject is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when body is missing', async () => {
      const result = await handleSendEmail({
        to: 'test@example.com',
        subject: 'Test',
      });

      expect(result.content[0].text).toBe('Body content is required.');
      expect(result.isError).toBe(true);
    });

    test('should return error when to is empty string', async () => {
      const result = await handleSendEmail({
        to: '',
        subject: 'Test',
        body: 'Body',
      });

      expect(result.content[0].text).toBe('Recipient (to) is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleSendEmail({
        to: 'test@example.com',
        subject: 'Test',
        body: 'Body',
      });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleSendEmail({
        to: 'test@example.com',
        subject: 'Test',
        body: 'Body',
      });

      expect(result.content[0].text).toBe('Error sending email: Graph API Error');
    });
  });
});
