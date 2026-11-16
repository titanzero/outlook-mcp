// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListGroups, handleGetGroup, handleListGroupMembers, handleSendEmailToGroup } = require('../../email/groups');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Groups / Distribution Lists Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListGroups', () => {
    test('should list groups user is member of', async () => {
      const mockGroups = [
        { id: 'group-1', displayName: 'Marketing Team', description: 'Marketing department', mail: 'marketing@company.com', mailEnabled: true, securityEnabled: false, groupTypes: ['Unified'], visibility: 'Private', createdDateTime: '2024-01-01T00:00:00Z' },
        { id: 'group-2', displayName: 'All Employees', description: 'Company-wide distribution list', mail: 'all@company.com', mailEnabled: true, securityEnabled: false, groupTypes: [], visibility: 'Public', createdDateTime: '2023-01-01T00:00:00Z' },
      ];

      mockClient.get.mockResolvedValue({ value: mockGroups });

      const result = await handleListGroups({ count: 50 });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/memberOf/microsoft.graph.group');
      expect(result.content[0].text).toContain('Groups (2)');
      expect(result.content[0].text).toContain('Marketing Team (Microsoft 365 Group)');
      expect(result.content[0].text).toContain('All Employees (Distribution List)');
    });

    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleListGroups({});

      expect(result.content[0].text).toBe('Authentication required');
    });
  });

  describe('handleGetGroup', () => {
    test('should get group details with members and owners', async () => {
      const mockGroup = {
        id: 'group-1',
        displayName: 'Test Group',
        description: 'Test description',
        mail: 'test@company.com',
        mailEnabled: true,
        securityEnabled: false,
        groupTypes: ['Unified'],
        visibility: 'Private',
        createdDateTime: '2024-01-01T00:00:00Z',
      };
      const mockMembers = [
        { id: 'user-1', displayName: 'John Doe', userPrincipalName: 'john@company.com', mail: 'john@company.com', jobTitle: 'Manager', department: 'IT' },
      ];
      const mockOwners = [
        { id: 'owner-1', displayName: 'Jane Smith', userPrincipalName: 'jane@company.com', mail: 'jane@company.com' },
      ];

      mockClient.get
        .mockResolvedValueOnce(mockGroup)
        .mockResolvedValueOnce({ value: mockMembers })
        .mockResolvedValueOnce({ value: mockOwners });

      graphGetPaginated
        .mockResolvedValueOnce({ value: mockMembers })
        .mockResolvedValueOnce({ value: mockOwners });

      const result = await handleGetGroup({ groupId: 'group-1' });

      expect(result.content[0].text).toContain('Group: Test Group (Microsoft 365 Group)');
      expect(result.content[0].text).toContain('Members (1)');
      expect(result.content[0].text).toContain('Owners (1)');
    });

    test('should return error when groupId is missing', async () => {
      const result = await handleGetGroup({});

      expect(result.content[0].text).toBe('Group ID is required (from list-groups).');
    });
  });

  describe('handleListGroupMembers', () => {
    test('should list group members', async () => {
      const mockMembers = [
        { id: 'user-1', displayName: 'John Doe', userPrincipalName: 'john@company.com', mail: 'john@company.com', jobTitle: 'Developer', department: 'Engineering', officeLocation: 'Building A' },
        { id: 'user-2', displayName: 'Jane Smith', userPrincipalName: 'jane@company.com', mail: 'jane@company.com', jobTitle: 'Designer', department: 'Design', officeLocation: 'Building B' },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockMembers });

      const result = await handleListGroupMembers({ groupId: 'group-1', count: 100 });

      expect(result.content[0].text).toContain('Group Members (2)');
      expect(result.content[0].text).toContain('John Doe <john@company.com>');
    });

    test('should return error when groupId is missing', async () => {
      const result = await handleListGroupMembers({});

      expect(result.content[0].text).toBe('Group ID is required (from list-groups).');
    });
  });

  describe('handleSendEmailToGroup', () => {
    test('should send email to group', async () => {
      const mockGroup = { id: 'group-1', displayName: 'Test Group', mail: 'test@company.com' };

      mockClient.get.mockResolvedValue(mockGroup);
      mockClient.api.mockReturnValue(mockClient);
      mockClient.post.mockResolvedValue({});

      const result = await handleSendEmailToGroup({
        groupId: 'group-1',
        subject: 'Test Subject',
        body: '<p>Test body</p>',
        importance: 'high',
      });

      expect(result.content[0].text).toContain('Email sent successfully to group "Test Group"');
    });

    test('should return error when group has no email', async () => {
      const mockGroup = { id: 'group-1', displayName: 'Test Group', mail: null };

      mockClient.get.mockResolvedValue(mockGroup);

      const result = await handleSendEmailToGroup({
        groupId: 'group-1',
        subject: 'Test',
        body: 'Test',
      });

      expect(result.content[0].text).toContain('does not have an email address');
    });

    test('should return error when required fields missing', async () => {
      const result = await handleSendEmailToGroup({ groupId: 'group-1', subject: 'Test' });

      expect(result.content[0].text).toBe('Body is required.');
    });
  });
});