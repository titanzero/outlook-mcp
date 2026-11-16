/**
 * Tests for email draft functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleCreateDraft, handleUpdateDraft, handleSendDraft } = require('../../email/draft');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('email draft', () => {
  let mockClient;
  let mockPost;
  let mockPatch;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPost = jest.fn().mockResolvedValue({ id: 'draft-123', subject: 'Draft Subject' });
    mockPatch = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn((path) => {
        if (path.includes('/send')) {
          return { post: mockPost };
        }
        return {
          post: mockPost,
          patch: mockPatch,
        };
      }),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleCreateDraft', () => {
    test('should create draft with required parameters', async () => {
      const result = await handleCreateDraft({
        to: 'recipient@example.com',
        subject: 'Draft Subject',
        body: 'Draft body content',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockPost).toHaveBeenCalledTimes(1);

      const message = mockPost.mock.calls[0][0];
      expect(message.subject).toBe('Draft Subject');
      expect(message.body.content).toBe('Draft body content');
      expect(message.body.contentType).toBe('text');
      expect(message.toRecipients).toHaveLength(1);
      expect(message.toRecipients[0].emailAddress.address).toBe('recipient@example.com');
      expect(message.importance).toBe('normal');

      expect(result.content[0].text).toContain('Draft created successfully');
      expect(result.content[0].text).toContain('draft-123');
      expect(result.content[0].text).toContain('Draft Subject');
    });

    test('should create draft with CC and BCC', async () => {
      await handleCreateDraft({
        to: 'to@example.com',
        cc: 'cc@example.com',
        bcc: 'bcc@example.com',
        subject: 'Draft with CC/BCC',
        body: 'Body',
      });

      const message = mockPost.mock.calls[0][0];
      expect(message.ccRecipients).toHaveLength(1);
      expect(message.bccRecipients).toHaveLength(1);
    });

    test('should detect HTML content', async () => {
      await handleCreateDraft({
        to: 'test@example.com',
        subject: 'HTML Draft',
        body: '<html><body>HTML</body></html>',
      });

      const message = mockPost.mock.calls[0][0];
      expect(message.body.contentType).toBe('html');
    });

    test('should use custom importance', async () => {
      await handleCreateDraft({
        to: 'test@example.com',
        subject: 'High Importance',
        body: 'Body',
        importance: 'high',
      });

      const message = mockPost.mock.calls[0][0];
      expect(message.importance).toBe('high');
    });

    test('should create draft with minimal parameters (only body)', async () => {
      await handleCreateDraft({
        body: 'Just body',
      });

      const message = mockPost.mock.calls[0][0];
      expect(message.subject).toBe('');
      expect(message.body.content).toBe('Just body');
      expect(message.toRecipients).toBeUndefined();
    });

    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleCreateDraft({
        to: 'test@example.com',
        subject: 'Test',
        body: 'Body',
      });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleCreateDraft({
        to: 'test@example.com',
        subject: 'Test',
        body: 'Body',
      });

      expect(result.content[0].text).toBe('Error creating draft: Graph API Error');
    });
  });

  describe('handleUpdateDraft', () => {
    test('should update draft with all parameters', async () => {
      const result = await handleUpdateDraft({
        id: 'draft-123',
        to: 'new@example.com',
        cc: 'newcc@example.com',
        subject: 'Updated Subject',
        body: 'Updated body',
        importance: 'high',
      });

      expect(mockClient.api).toHaveBeenCalledWith('me/messages/draft-123');
      expect(mockPatch).toHaveBeenCalledTimes(1);

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.subject).toBe('Updated Subject');
      expect(patch.body.content).toBe('Updated body');
      expect(patch.body.contentType).toBe('text');
      expect(patch.importance).toBe('high');
      expect(patch.toRecipients).toHaveLength(1);
      expect(patch.ccRecipients).toHaveLength(1);

      expect(result.content[0].text).toBe('Draft updated successfully.');
    });

    test('should update only provided fields', async () => {
      await handleUpdateDraft({
        id: 'draft-123',
        subject: 'Only Subject Updated',
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.subject).toBe('Only Subject Updated');
      expect(patch.body).toBeUndefined();
      expect(patch.toRecipients).toBeUndefined();
      expect(patch.importance).toBeUndefined();
    });

    test('should detect HTML content in update', async () => {
      await handleUpdateDraft({
        id: 'draft-123',
        body: '<html><body>HTML Update</body></html>',
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.body.contentType).toBe('html');
    });

    test('should return error when id is missing', async () => {
      const result = await handleUpdateDraft({
        subject: 'Test',
        body: 'Body',
      });

      expect(result.content[0].text).toBe('Draft ID is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleUpdateDraft({
        id: 'draft-123',
        subject: 'Test',
      });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPatch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleUpdateDraft({
        id: 'draft-123',
        subject: 'Test',
      });

      expect(result.content[0].text).toBe('Error updating draft: Graph API Error');
    });
  });

  describe('handleSendDraft', () => {
    test('should send draft', async () => {
      const result = await handleSendDraft({ id: 'draft-123' });

      expect(mockClient.api).toHaveBeenCalledWith('me/messages/draft-123/send');
      expect(mockPost).toHaveBeenCalledWith({});
      expect(result.content[0].text).toBe('Draft sent successfully.');
    });

    test('should return error when id is missing', async () => {
      const result = await handleSendDraft({});

      expect(result.content[0].text).toBe('Draft ID is required.');
      expect(result.isError).toBe(true);
    });

    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleSendDraft({ id: 'draft-123' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleSendDraft({ id: 'draft-123' });

      expect(result.content[0].text).toBe('Error sending draft: Graph API Error');
    });
  });
});
