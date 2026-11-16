// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListContactLists, handleCreateContactList, handleUpdateContactList, handleDeleteContactList, handleSendEmailToContactList } = require('../../contacts/lists');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Contact Lists / Distribution Lists Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      patch: jest.fn(),
      delete: jest.fn(),
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
      filter: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListContactLists', () => {
    test('should list contact lists with members', async () => {
      const mockLists = [
        {
          id: 'list-1',
          displayName: 'Team Distribution',
          emailAddresses: [{ address: 'team@company.com', name: 'Team' }],
          companyName: 'Acme Corp',
          members: [
            { id: 'contact-1', displayName: 'John Doe', emailAddresses: [{ address: 'john@company.com' }] },
            { id: 'contact-2', displayName: 'Jane Smith', emailAddresses: [{ address: 'jane@company.com' }] },
          ],
        },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockLists });

      const result = await handleListContactLists({ count: 50 });

      expect(result.content[0].text).toContain('Contact Lists (1)');
      expect(result.content[0].text).toContain('Team Distribution');
      expect(result.content[0].text).toContain('undefined members');
      expect(result.content[0].text).toContain('John Doe');
    });
  });

  describe('handleCreateContactList', () => {
    test('should create contact list with members', async () => {
      const mockMembers = [
        { id: 'contact-1', displayName: 'John', emailAddresses: [{ address: 'john@company.com' }] },
        { id: 'contact-2', displayName: 'Jane', emailAddresses: [{ address: 'jane@company.com' }] },
      ];

      mockClient.get
        .mockResolvedValueOnce({ id: 'contact-1', displayName: 'John', emailAddresses: [{ address: 'john@company.com' }] })
        .mockResolvedValueOnce({ id: 'contact-2', displayName: 'Jane', emailAddresses: [{ address: 'jane@company.com' }] });

      mockClient.post.mockResolvedValue({ id: 'new-list-id', displayName: 'New List' });

      const result = await handleCreateContactList({
        name: 'New List',
        memberIds: ['contact-1', 'contact-2'],
        email: 'newlist@company.com',
      });

      expect(result.content[0].text).toContain('Contact list "New List" created successfully with 2 members');
    });

    test('should return error when name is missing', async () => {
      const result = await handleCreateContactList({ memberIds: ['contact-1'] });

      expect(result.content[0].text).toBe('Contact list name is required.');
    });

    test('should return error when memberIds is missing', async () => {
      const result = await handleCreateContactList({ name: 'Test' });

      expect(result.content[0].text).toBe('At least one member ID is required (from list-contacts).');
    });

    test('should return error for non-existent member', async () => {
      const error = new Error('Not Found');
      error.statusCode = 404;
      mockClient.get.mockRejectedValue(error);

      const result = await handleCreateContactList({ name: 'Test', memberIds: ['invalid'] });

      expect(result.content[0].text).toContain('Member contact invalid not found');
    });
  });

  describe('handleUpdateContactList', () => {
    test('should update contact list', async () => {
      const mockList = {
        id: 'list-1',
        displayName: 'Old Name',
        members: [
          { id: 'contact-1', displayName: 'John', emailAddresses: [{ address: 'john@company.com' }] },
        ],
      };

      mockClient.get.mockResolvedValue(mockList);
      mockClient.patch.mockResolvedValue({});

      const result = await handleUpdateContactList({
        listId: 'list-1',
        name: 'New Name',
        addMemberIds: ['contact-2'],
        removeMemberIds: ['contact-1'],
      });

      expect(result.content[0].text).toContain('Contact list updated successfully');
    });
  });

  describe('handleDeleteContactList', () => {
    test('should delete contact list', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleDeleteContactList({ listId: 'list-1' });

      expect(result.content[0].text).toBe('Contact list deleted successfully.');
    });
  });

  describe('handleSendEmailToContactList', () => {
    test('should send email to contact list', async () => {
      const mockList = {
        id: 'list-1',
        displayName: 'Team',
        emailAddresses: [{ address: 'team@company.com', name: 'Team' }],
        members: [
          { id: 'contact-1', displayName: 'John', emailAddresses: [{ address: 'john@company.com' }] },
        ],
      };

      mockClient.get.mockResolvedValue(mockList);
      mockClient.api.mockReturnValue(mockClient);
      mockClient.post.mockResolvedValue({});

      const result = await handleSendEmailToContactList({
        listId: 'list-1',
        subject: 'Test',
        body: '<p>Test</p>',
      });

      expect(result.content[0].text).toContain('Email sent to contact list "Team" (2 recipients)');
    });

    test('should return error when list has no emails', async () => {
      const mockList = { id: 'list-1', displayName: 'Empty', emailAddresses: [], members: [] };

      mockClient.get.mockResolvedValue(mockList);

      const result = await handleSendEmailToContactList({ listId: 'list-1', subject: 'Test', body: 'Test' });

      expect(result.content[0].text).toContain('has no email addresses to send to');
    });
  });
});