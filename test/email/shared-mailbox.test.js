// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListSharedMailboxes, handleAccessSharedMailbox, handleListSharedMailboxFolders } = require('../../email/shared-mailbox');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Shared Mailbox Tools', () => {
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
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListSharedMailboxes', () => {
    test('should list own folders and shared mailboxes', async () => {
      const mockFolders = [
        { id: '1', displayName: 'Inbox', totalItemCount: 10, unreadItemCount: 2, childFolderCount: 0, wellKnownName: 'inbox' },
        { id: '2', displayName: 'Sent', totalItemCount: 5, unreadItemCount: 0, childFolderCount: 0, wellKnownName: 'sentitems' },
      ];
      const mockDelegates = [
        { emailAddress: { address: 'shared@company.com' }, displayName: 'Shared Mailbox', permissions: ['read', 'write'] },
      ];

      mockClient.get
        .mockResolvedValueOnce({ value: mockFolders })
        .mockResolvedValueOnce({ value: mockDelegates });

      const result = await handleListSharedMailboxes({});

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders');
      expect(result.content[0].text).toContain('Inbox (10 items, 2 unread)');
      expect(result.content[0].text).toContain('Shared Mailbox <shared@company.com>');
    });

    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleListSharedMailboxes({});

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockClient.get.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleListSharedMailboxes({});

      expect(result.content[0].text).toBe('Error listing shared mailboxes: Graph API Error');
    });
  });

  describe('handleAccessSharedMailbox', () => {
    test('should access shared mailbox and list emails', async () => {
      const mockEmails = [
        { id: 'email-1', subject: 'Shared Email', from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } }, toRecipients: [], ccRecipients: [], receivedDateTime: '2024-01-15T10:30:00Z', bodyPreview: 'Preview', hasAttachments: false, importance: 'normal', isRead: true },
      ];

      mockClient.get.mockResolvedValue({ value: mockEmails });
      mockClient.api.mockReturnValue(mockClient);

      const result = await handleAccessSharedMailbox({ emailAddress: 'shared@company.com', folder: 'inbox', count: 20 });

      expect(result.content[0].text).toContain('Shared Mailbox: shared@company.com');
      expect(result.content[0].text).toContain('Shared Email');
    });

    test('should return error when email address is missing', async () => {
      const result = await handleAccessSharedMailbox({ folder: 'inbox' });

      expect(result.content[0].text).toBe('Email address of the shared mailbox is required.');
    });

    test('should handle permission denied error', async () => {
      const error = new Error('Forbidden');
      error.statusCode = 403;
      mockClient.get.mockRejectedValue(error);

      const result = await handleAccessSharedMailbox({ emailAddress: 'shared@company.com' });

      expect(result.content[0].text).toBe('Forbidden');
    });
  });

  describe('handleListSharedMailboxFolders', () => {
    test('should list folders in shared mailbox', async () => {
      const mockFolders = [
        { id: '1', displayName: 'Inbox', totalItemCount: 10, unreadItemCount: 2, childFolderCount: 0, wellKnownName: 'inbox' },
        { id: '2', displayName: 'Projects', totalItemCount: 5, unreadItemCount: 1, childFolderCount: 2, wellKnownName: null },
      ];

      mockClient.get.mockResolvedValue({ value: mockFolders });

      const result = await handleListSharedMailboxFolders({ emailAddress: 'shared@company.com' });

      expect(result.content[0].text).toContain('Folders in shared@company.com');
      expect(result.content[0].text).toContain('Inbox (10 items, 2 unread)');
      expect(result.content[0].text).toContain('Projects (5 items, 1 unread)');
    });

    test('should return error when email address is missing', async () => {
      const result = await handleListSharedMailboxFolders({});

      expect(result.content[0].text).toBe('Email address of the shared mailbox is required.');
    });
  });
});