// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListTaskAttachments, handleUploadTaskAttachment, handleDownloadTaskAttachment, handleDeleteTaskAttachment } = require('../../tasks/attachments');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Task Attachments Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      delete: jest.fn(),
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListTaskAttachments', () => {
    test('should list task attachments', async () => {
      const mockAttachments = [
        { id: 'att-1', name: 'document.pdf', contentType: 'application/pdf', size: 102400, lastModifiedDateTime: '2024-01-15T10:00:00Z', isInline: false },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockAttachments });

      const result = await handleListTaskAttachments({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toContain('Task Attachments (1)');
      expect(result.content[0].text).toContain('document.pdf (application/pdf, 102400 bytes)');
    });
  });

  describe('handleUploadTaskAttachment', () => {
    test('should upload task attachment', async () => {
      mockClient.post.mockResolvedValue({ id: 'att-new', name: 'test.pdf', size: 1024 });

      const base64Content = Buffer.from('test content').toString('base64');
      const result = await handleUploadTaskAttachment({
        listId: 'list-1',
        taskId: 'task-1',
        name: 'test.pdf',
        contentBytes: base64Content,
      });

      expect(result.content[0].text).toContain('Task attachment uploaded successfully');
      expect(result.content[0].text).toContain('test.pdf');
    });

    test('should return error for invalid base64', async () => {
      const result = await handleUploadTaskAttachment({ listId: 'list-1', taskId: 'task-1', name: 'test.pdf', contentBytes: 'invalid!' });

      expect(result.content[0].text).toContain('Error uploading task attachment');
    });

    test('should return error for oversized attachment', async () => {
      const largeBuffer = Buffer.alloc(26 * 1024 * 1024); // 26MB
      const largeBase64 = largeBuffer.toString('base64');

      const result = await handleUploadTaskAttachment({ listId: 'list-1', taskId: 'task-1', name: 'large.pdf', contentBytes: largeBase64 });

      expect(result.content[0].text).toContain('Attachment too large');
    });
  });

  describe('handleDownloadTaskAttachment', () => {
    test('should download task attachment', async () => {
      const mockAttachment = { id: 'att-1', name: 'test.pdf', contentType: 'application/pdf', size: 1024, lastModifiedDateTime: '2024-01-15T10:00:00Z', isInline: false };
      const mockContent = Buffer.from('file content');

      mockClient.get.mockResolvedValue(mockAttachment);
      mockClient.api.mockReturnValue(mockClient);
      mockClient.get.mockResolvedValueOnce(mockAttachment).mockResolvedValueOnce(mockContent);

      const result = await handleDownloadTaskAttachment({ listId: 'list-1', taskId: 'task-1', attachmentId: 'att-1' });

      expect(result.content[0].text).toContain('Task Attachment: test.pdf');
      expect(result.content[0].text).toContain('Base64 Content:');
    });
  });

  describe('handleDeleteTaskAttachment', () => {
    test('should delete task attachment', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleDeleteTaskAttachment({ listId: 'list-1', taskId: 'task-1', attachmentId: 'att-1' });

      expect(result.content[0].text).toBe('Task attachment deleted successfully.');
    });
  });
});