// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListDelegates, handleAddDelegate, handleUpdateDelegate, handleRemoveDelegate } = require('../../mailbox/delegates');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Delegates Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      patch: jest.fn(),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListDelegates', () => {
    test('should list delegates with permissions', async () => {
      const mockDelegates = [
        {
          emailAddress: { address: 'delegate@company.com', name: 'Delegate User' },
          permissions: { calendar: { role: 'editor' }, inbox: { role: 'reviewer' }, contacts: { role: 'author' } },
          canViewPrivateItems: true,
          receiveCopiesOfMeetingMessages: false,
        },
      ];

      mockClient.get.mockResolvedValue({ value: mockDelegates });

      const result = await handleListDelegates({});

      expect(result.content[0].text).toContain('Delegates (1)');
      expect(result.content[0].text).toContain('Delegate User <delegate@company.com>');
      expect(result.content[0].text).toContain('Calendar: Editor (read/write)');
      expect(result.content[0].text).toContain('Inbox: Reviewer (read only)');
      expect(result.content[0].text).toContain('Private items: Yes');
    });

    test('should handle API not available', async () => {
      const error = new Error('Not Found');
      error.statusCode = 404;
      mockClient.get.mockRejectedValue(error);

      const result = await handleListDelegates({});

      expect(result.content[0].text).toContain('Delegates API not available');
    });
  });

  describe('handleAddDelegate', () => {
    test('should add delegate', async () => {
      mockClient.get.mockResolvedValue({ value: [] });
      mockClient.patch.mockResolvedValue({});

      const result = await handleAddDelegate({
        emailAddress: 'new@company.com',
        calendarRole: 'editor',
        inboxRole: 'reviewer',
        contactsRole: 'author',
        canViewPrivateItems: true,
      });

      expect(result.content[0].text).toBe('Delegate new@company.com added successfully.');
    });

    test('should validate roles', async () => {
      const result = await handleAddDelegate({ emailAddress: 'user@company.com', calendarRole: 'invalid' });

      expect(result.content[0].text).toContain('Invalid role');
    });
  });

  describe('handleUpdateDelegate', () => {
    test('should update delegate', async () => {
      const mockDelegates = {
        value: [
          { emailAddress: { address: 'delegate@company.com' }, permissions: {}, canViewPrivateItems: false },
        ],
      };

      mockClient.get.mockResolvedValue(mockDelegates);
      mockClient.patch.mockResolvedValue({});

      const result = await handleUpdateDelegate({ emailAddress: 'delegate@company.com', calendarRole: 'author' });

      expect(result.content[0].text).toBe('Delegate delegate@company.com updated successfully.');
    });
  });

  describe('handleRemoveDelegate', () => {
    test('should remove delegate', async () => {
      const mockDelegates = {
        value: [{ emailAddress: { address: 'delegate@company.com' } }],
      };

      mockClient.get.mockResolvedValue(mockDelegates);
      mockClient.patch.mockResolvedValue({});

      const result = await handleRemoveDelegate({ emailAddress: 'delegate@company.com' });

      expect(result.content[0].text).toBe('Delegate delegate@company.com removed successfully.');
    });
  });
});