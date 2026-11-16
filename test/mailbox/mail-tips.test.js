// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleGetMailTips, handleGetMailboxUsage, handleGetUserProfile } = require('../../mailbox/mail-tips');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Mail Tips & Mailbox Info Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      select: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleGetMailTips', () => {
    test('should get mail tips for recipients', async () => {
      const mockTips = [
        {
          emailAddress: { address: 'user@company.com' },
          automaticReplies: { message: 'I am out of office', scheduledStartTime: { dateTime: '2024-01-15T00:00:00Z' }, scheduledEndTime: { dateTime: '2024-01-20T00:00:00Z' } },
          mailboxFull: false,
          customMailTip: 'Custom tip',
          maxMessageSize: 10485760,
        },
      ];

      mockClient.post.mockResolvedValue({ value: mockTips });

      const result = await handleGetMailTips({ emailAddresses: ['user@company.com'] });

      expect(result.content[0].text).toContain('Mail Tips (1 recipients)');
      expect(result.content[0].text).toContain('user@company.com');
      expect(result.content[0].text).toContain('Auto-reply: I am out of office');
      expect(result.content[0].text).toContain('Max message size: 10240 KB');
    });

    test('should return error for empty email array', async () => {
      const result = await handleGetMailTips({ emailAddresses: [] });

      expect(result.content[0].text).toBe('emailAddresses array is required (at least one recipient).');
    });

    test('should limit to 20 recipients', async () => {
      const result = await handleGetMailTips({ emailAddresses: Array(21).fill('user@company.com') });

      expect(result.content[0].text).toContain('Maximum 20 email addresses allowed');
    });
  });

  describe('handleGetMailboxUsage', () => {
    test('should get mailbox usage', async () => {
      const mockSettings = {
        displayName: 'Test User',
        userPrincipalName: 'test@company.com',
        mail: 'test@company.com',
        timeZone: 'Pacific Standard Time',
        language: { displayName: 'English (United States)' },
        automaticRepliesSetting: { status: 'disabled' },
        workingHours: { daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'], startTime: '09:00:00.0000000', endTime: '17:00:00.0000000', timeZone: { name: 'Pacific Standard Time' } },
      };

      const mockFolders = { value: [{ totalItemCount: 100 }, { totalItemCount: 50 }] };

      mockClient.get.mockResolvedValue(mockSettings);
      mockClient.api.mockReturnValue(mockClient);
      mockClient.get.mockResolvedValueOnce(mockSettings).mockResolvedValueOnce(mockFolders);

      const result = await handleGetMailboxUsage({});

      expect(result.content[0].text).toContain('Mailbox Usage');
      expect(result.content[0].text).toContain('Name: Test User');
      expect(result.content[0].text).toContain('Email: test@company.com');
      expect(result.content[0].text).toContain('Total Messages (estimate): 150');
    });
  });

  describe('handleGetUserProfile', () => {
    test('should get user profile', async () => {
      const mockUser = {
        id: 'user-1',
        displayName: 'Test User',
        givenName: 'Test',
        surname: 'User',
        userPrincipalName: 'test@company.com',
        mail: 'test@company.com',
        mailNickname: 'testuser',
        jobTitle: 'Developer',
        officeLocation: 'Building A',
        department: 'Engineering',
        companyName: 'Acme Corp',
        businessPhones: ['555-1234'],
        mobilePhone: '555-5678',
        preferredLanguage: 'en-US',
        employeeId: 'EMP001',
        employeeHireDate: '2020-01-15T00:00:00Z',
        createdDateTime: '2020-01-01T00:00:00Z',
      };

      mockClient.get.mockResolvedValue(mockUser);

      const result = await handleGetUserProfile({});

      expect(result.content[0].text).toContain('User Profile');
      expect(result.content[0].text).toContain('Name: Test User');
      expect(result.content[0].text).toContain('Job Title: Developer');
      expect(result.content[0].text).toContain('Employee ID: EMP001');
    });
  });
});