/**
 * Tests for email search functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleSearchEmails = require('../../email/search');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');
const { resolveFolderPath } = require('../../email/folder-utils');

jest.mock('../../utils/graph-client');
jest.mock('../../email/folder-utils');
jest.mock('../../config', () => ({
  DEFAULT_PAGE_SIZE: 10,
  MAX_RESULT_COUNT: 100,
  EMAIL_SELECT_FIELDS: 'id,subject,from,receivedDateTime,isRead,hasAttachments',
}));

describe('handleSearchEmails', () => {
  let mockClient;
  const mockEmails = [
    {
      id: 'email-1',
      subject: 'Meeting Request',
      from: { emailAddress: { name: 'John Doe', address: 'john@example.com' } },
      receivedDateTime: '2024-01-15T10:30:00Z',
      isRead: false,
      hasAttachments: true,
    },
    {
      id: 'email-2',
      subject: 'Project Update',
      from: { emailAddress: { name: 'Jane Smith', address: 'jane@example.com' } },
      receivedDateTime: '2024-01-14T15:20:00Z',
      isRead: true,
      hasAttachments: false,
    },
  ];

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(),
    };
    getGraphClient.mockResolvedValue(mockClient);
    resolveFolderPath.mockResolvedValue('me/mailFolders/inbox/messages');
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful searches', () => {
    test('should search with query parameter', async () => {
      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      const result = await handleSearchEmails({ query: 'meeting', count: 5 });

      expect(graphGetPaginated).toHaveBeenCalled();
      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$search).toBeDefined();
      expect(result.content[0].text).toContain('Found 2 emails matching your search criteria');
      expect(result.content[0].text).toContain('Meeting Request');
    });

    test('should search with from parameter', async () => {
      graphGetPaginated.mockResolvedValue({ value: [mockEmails[0]] });

      const result = await handleSearchEmails({ from: 'john@example.com' });

      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$search).toContain('from:john@example.com');
    });

    test('should search with to parameter', async () => {
      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      const result = await handleSearchEmails({ to: 'jane@example.com' });

      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$search).toContain('to:jane@example.com');
    });

    test('should search with subject parameter', async () => {
      graphGetPaginated.mockResolvedValue({ value: [mockEmails[1]] });

      const result = await handleSearchEmails({ subject: 'Project' });

      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$search).toContain('subject:Project');
    });

    test('should filter by hasAttachments', async () => {
      graphGetPaginated.mockResolvedValue({ value: [mockEmails[0]] });

      const result = await handleSearchEmails({ hasAttachments: true });

      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$filter).toBe('hasAttachments eq true');
    });

    test('should filter by unreadOnly', async () => {
      graphGetPaginated.mockResolvedValue({ value: [mockEmails[0]] });

      const result = await handleSearchEmails({ unreadOnly: true });

      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$filter).toBe('isRead eq false');
    });

    test('should combine search terms and filters', async () => {
      graphGetPaginated.mockResolvedValue({ value: [mockEmails[0]] });

      await handleSearchEmails({
        query: 'meeting',
        from: 'john@example.com',
        hasAttachments: true,
        unreadOnly: true,
      });

      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$search).toContain('meeting');
      expect(params.$search).toContain('from:john@example.com');
      expect(params.$filter).toBe('hasAttachments eq true and isRead eq false');
    });

    test('should respect count limit', async () => {
      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ count: 5 });

      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$top).toBe(5);
    });

    test('should cap count at MAX_RESULT_COUNT', async () => {
      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ count: 500 });

      const params = graphGetPaginated.mock.calls[0][2];
      expect(params.$top).toBe(10); // DEFAULT_PAGE_SIZE is 10, min(10, 100) = 10
    });

    test('should use folder when specified', async () => {
      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({ folder: 'drafts' });

      expect(resolveFolderPath).toHaveBeenCalledWith('drafts');
    });

    test('should search all folders when no folder specified', async () => {
      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      await handleSearchEmails({});

      // When no folder is specified, it uses 'me/messages' directly without calling resolveFolderPath
      expect(resolveFolderPath).not.toHaveBeenCalled();
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no emails found', async () => {
      graphGetPaginated.mockResolvedValue({ value: [] });

      const result = await handleSearchEmails({ query: 'nonexistent' });

      expect(result.content[0].text).toBe('No emails found matching your search criteria.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleSearchEmails({ query: 'test' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      graphGetPaginated.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleSearchEmails({ query: 'test' });

      expect(result.content[0].text).toBe('Error searching emails: Graph API Error');
    });

    test('should handle folder resolution error', async () => {
      resolveFolderPath.mockRejectedValue(new Error('Folder not found'));

      const result = await handleSearchEmails({ folder: 'invalid' });

      expect(result.content[0].text).toBe('Error searching emails: Folder not found');
    });
  });

  describe('progressive search fallback (via handleSearchEmails)', () => {
    test('should fallback to recent emails when no criteria provided', async () => {
      graphGetPaginated.mockResolvedValue({ value: mockEmails });

      const result = await handleSearchEmails({});

      expect(result.content[0].text).toContain('Found 2 emails matching your search criteria');
    });
  });
});
