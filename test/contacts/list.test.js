/**
 * Tests for contacts list functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleListContacts = require('../../contacts/list');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');
jest.mock('../../config', () => ({
  MAX_RESULT_COUNT: 100,
}));

function createMockClient(contacts) {
  const mockChain = {
    select: jest.fn().mockReturnThis(),
    top: jest.fn().mockReturnThis(),
    filter: jest.fn().mockReturnThis(),
    get: jest.fn().mockResolvedValue({ value: contacts }),
  };
  return {
    api: jest.fn().mockReturnValue(mockChain),
  };
}

describe('handleListContacts', () => {
  let mockClient;

  const mockContacts = [
    {
      id: 'contact-1',
      displayName: 'John Doe',
      emailAddresses: [{ address: 'john@example.com', name: 'John Doe' }],
      mobilePhone: '+1-555-1234',
      businessPhones: ['+1-555-5678'],
      jobTitle: 'Software Engineer',
      companyName: 'Acme Corp',
    },
    {
      id: 'contact-2',
      displayName: 'Jane Smith',
      emailAddresses: [{ address: 'jane@example.com', name: 'Jane Smith' }],
      mobilePhone: null,
      businessPhones: [],
      jobTitle: 'Product Manager',
      companyName: 'Tech Inc',
    },
  ];

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = createMockClient(mockContacts);
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful contact listing', () => {
    test('should list contacts with default count', async () => {
      const result = await handleListContacts({});

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/contacts');
      expect(result.content[0].text).toContain('Contacts (2):');
      expect(result.content[0].text).toContain('John Doe <john@example.com>');
      expect(result.content[0].text).toContain('Software Engineer');
      expect(result.content[0].text).toContain('Jane Smith <jane@example.com>');
      expect(result.content[0].text).toContain('Product Manager');
    });

    test('should respect custom count parameter', async () => {
      await handleListContacts({ count: 10 });

      expect(mockClient.api().select().top).toHaveBeenCalledWith(10);
    });

    test('should cap count at MAX_RESULT_COUNT', async () => {
      await handleListContacts({ count: 500 });

      expect(mockClient.api().select().top).toHaveBeenCalledWith(100);
    });

    test('should filter by query parameter', async () => {
      await handleListContacts({ query: 'john' });

      expect(mockClient.api().select().top().filter).toHaveBeenCalled();
      const filterCall = mockClient.api().select().top().filter.mock.calls[0][0];
      expect(filterCall).toContain("startsWith(displayName,'john')");
      expect(filterCall).toContain("startsWith(e/address,'john')");
    });

    test('should return structured data with emails, phone, title, company', async () => {
      const result = await handleListContacts({});

      // The response should include structured data for TOON
      expect(result.content[0].text).toBeDefined();
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no contacts found', async () => {
      mockClient = createMockClient([]);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListContacts({});

      expect(result.content[0].text).toBe('No contacts found.');
    });

    test('should return no contacts when query matches nothing', async () => {
      mockClient = createMockClient([]);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListContacts({ query: 'nonexistent' });

      expect(result.content[0].text).toBe('No contacts found.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleListContacts({});

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      const mockChain = {
        select: jest.fn().mockReturnThis(),
        top: jest.fn().mockReturnThis(),
        filter: jest.fn().mockReturnThis(),
        get: jest.fn().mockRejectedValue(new Error('Graph API Error')),
      };
      mockClient = { api: jest.fn().mockReturnValue(mockChain) };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListContacts({});

      expect(result.content[0].text).toBe('Error listing contacts: Graph API Error');
    });
  });
});
