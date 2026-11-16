/**
 * Tests for contacts create functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleCreateContact = require('../../contacts/create');
const handleCreateContactFull = require('../../contacts/create-full');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleCreateContact', () => {
  let mockClient;
  let mockPost;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPost = jest.fn().mockResolvedValue({ id: 'contact-123', displayName: 'John Doe' });
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

  describe('successful contact creation', () => {
    test('should create contact with required displayName', async () => {
      const result = await handleCreateContact({ displayName: 'John Doe' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/contacts');
      expect(mockPost).toHaveBeenCalledWith({ displayName: 'John Doe' });

      expect(result.content[0].text).toContain('Contact created successfully');
      expect(result.content[0].text).toContain('contact-123');
      expect(result.content[0].text).toContain('John Doe');
    });

    test('should create contact with email', async () => {
      await handleCreateContact({ displayName: 'John Doe', email: 'john@example.com' });

      expect(mockPost).toHaveBeenCalledWith({
        displayName: 'John Doe',
        emailAddresses: [{ address: 'john@example.com', name: 'John Doe' }],
      });
    });

    test('should create contact with phone', async () => {
      await handleCreateContact({ displayName: 'John Doe', phone: '+1-555-1234' });

      expect(mockPost).toHaveBeenCalledWith({
        displayName: 'John Doe',
        mobilePhone: '+1-555-1234',
      });
    });

    test('should create contact with jobTitle', async () => {
      await handleCreateContact({ displayName: 'John Doe', jobTitle: 'Engineer' });

      expect(mockPost).toHaveBeenCalledWith({
        displayName: 'John Doe',
        jobTitle: 'Engineer',
      });
    });

    test('should create contact with companyName', async () => {
      await handleCreateContact({ displayName: 'John Doe', companyName: 'Acme Corp' });

      expect(mockPost).toHaveBeenCalledWith({
        displayName: 'John Doe',
        companyName: 'Acme Corp',
      });
    });

    test('should create contact with all fields', async () => {
      await handleCreateContact({
        displayName: 'John Doe',
        email: 'john@example.com',
        phone: '+1-555-1234',
        jobTitle: 'Engineer',
        companyName: 'Acme Corp',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.displayName).toBe('John Doe');
      expect(contact.emailAddresses).toEqual([{ address: 'john@example.com', name: 'John Doe' }]);
      expect(contact.mobilePhone).toBe('+1-555-1234');
      expect(contact.jobTitle).toBe('Engineer');
      expect(contact.companyName).toBe('Acme Corp');
    });
  });

  describe('validation', () => {
    test('should return error when displayName is missing', async () => {
      const result = await handleCreateContact({});

      expect(result.content[0].text).toBe('displayName is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when displayName is empty string', async () => {
      const result = await handleCreateContact({ displayName: '' });

      expect(result.content[0].text).toBe('displayName is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleCreateContact({ displayName: 'Test' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleCreateContact({ displayName: 'Test' });

      expect(result.content[0].text).toBe('Error creating contact: Graph API Error');
    });
  });
});

describe('handleCreateContactFull', () => {
  let mockClient;
  let mockPost;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPost = jest.fn().mockResolvedValue({ id: 'contact-full-123', displayName: 'John Doe' });
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

  describe('successful full contact creation', () => {
    test('should create contact with all basic fields', async () => {
      const result = await handleCreateContactFull({
        displayName: 'John Doe',
        givenName: 'John',
        surname: 'Doe',
        email: 'john@example.com',
        phone: '+1-555-1234',
        jobTitle: 'Engineer',
        companyName: 'Acme Corp',
        department: 'Engineering',
        officeLocation: 'Building A',
        profession: 'Software Developer',
        spouseName: 'Jane Doe',
        personalNotes: 'Notes here',
        manager: 'Bob Manager',
        assistantName: 'Alice Assistant',
        fileAs: 'Doe, John',
        initials: 'JD',
        generation: 'Jr.',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.displayName).toBe('John Doe');
      expect(contact.givenName).toBe('John');
      expect(contact.surname).toBe('Doe');
      expect(contact.emailAddresses).toEqual([{ address: 'john@example.com', name: 'John Doe' }]);
      expect(contact.mobilePhone).toBe('+1-555-1234');
      expect(contact.jobTitle).toBe('Engineer');
      expect(contact.companyName).toBe('Acme Corp');
      expect(contact.department).toBe('Engineering');
      expect(contact.officeLocation).toBe('Building A');
      expect(contact.profession).toBe('Software Developer');
      expect(contact.spouseName).toBe('Jane Doe');
      expect(contact.personalNotes).toBe('Notes here');
      expect(contact.manager).toBe('Bob Manager');
      expect(contact.assistantName).toBe('Alice Assistant');
      expect(contact.fileAs).toBe('Doe, John');
      expect(contact.initials).toBe('JD');
      expect(contact.generation).toBe('Jr.');

      expect(result.content[0].text).toContain('Contact created successfully');
    });

    test('should create contact with multiple email addresses', async () => {
      await handleCreateContactFull({
        displayName: 'John Doe',
        emailAddresses: 'john@example.com|John Work, john.personal@example.com|John Personal',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.emailAddresses).toHaveLength(2);
      expect(contact.emailAddresses[0]).toEqual({ address: 'john@example.com', name: 'John Work' });
      expect(contact.emailAddresses[1]).toEqual({ address: 'john.personal@example.com', name: 'John Personal' });
    });

    test('should create contact with multiple business phones', async () => {
      await handleCreateContactFull({
        displayName: 'John Doe',
        businessPhones: '+1-555-1234, +1-555-5678',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.businessPhones).toEqual(['+1-555-1234', '+1-555-5678']);
    });

    test('should create contact with home phones', async () => {
      await handleCreateContactFull({
        displayName: 'John Doe',
        homePhones: '+1-555-9999',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.homePhones).toEqual(['+1-555-9999']);
    });

    test('should create contact with addresses', async () => {
      await handleCreateContactFull({
        displayName: 'John Doe',
        businessAddress: '123 Main St;Seattle;WA;98101;USA',
        homeAddress: '456 Home Ave;Redmond;WA;98052;USA',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.businessAddress).toEqual({
        street: '123 Main St',
        city: 'Seattle',
        state: 'WA',
        postalCode: '98101',
        countryOrRegion: 'USA',
      });
      expect(contact.homeAddress).toEqual({
        street: '456 Home Ave',
        city: 'Redmond',
        state: 'WA',
        postalCode: '98052',
        countryOrRegion: 'USA',
      });
    });

    test('should create contact with birthday', async () => {
      await handleCreateContactFull({
        displayName: 'John Doe',
        birthday: '1990-01-15',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.birthday).toBe('1990-01-15');
    });

    test('should create contact with IM addresses', async () => {
      await handleCreateContactFull({
        displayName: 'John Doe',
        imAddresses: 'john@im.com, john2@im.com',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.imAddresses).toEqual(['john@im.com', 'john2@im.com']);
    });

    test('should create contact with titles array', async () => {
      await handleCreateContactFull({
        displayName: 'John Doe',
        titles: 'Senior Engineer, Tech Lead',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.titles).toEqual(['Senior Engineer', 'Tech Lead']);
    });

    test('should prefer mobilePhone over phone parameter', async () => {
      await handleCreateContactFull({
        displayName: 'John Doe',
        phone: '+1-555-1234',
        mobilePhone: '+1-555-9999',
      });

      const contact = mockPost.mock.calls[0][0];
      expect(contact.mobilePhone).toBe('+1-555-9999');
    });
  });

  describe('validation', () => {
    test('should return error when displayName is missing', async () => {
      const result = await handleCreateContactFull({});

      expect(result.content[0].text).toBe('displayName is required.');
      expect(result.isError).toBe(true);
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleCreateContactFull({ displayName: 'Test' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPost.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleCreateContactFull({ displayName: 'Test' });

      expect(result.content[0].text).toBe('Error creating contact: Graph API Error');
    });
  });
});
