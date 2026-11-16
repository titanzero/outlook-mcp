// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleShareCalendar, handleListCalendarPermissions, handleUpdateCalendarPermission, handleRemoveCalendarPermission, handleGetSharedCalendars } = require('../../calendar/sharing');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Calendar Sharing Tools', () => {
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

  describe('handleShareCalendar', () => {
    test('should share calendar with user', async () => {
      mockClient.post.mockResolvedValue({});

      const result = await handleShareCalendar({
        calendarId: 'cal-1',
        emailAddress: 'user@company.com',
        role: 'read',
        message: 'Please review my calendar',
      });

      expect(result.content[0].text).toContain('Calendar shared successfully with user@company.com as "Full details (read)"');
    });

    test('should return error when calendarId is missing', async () => {
      const result = await handleShareCalendar({ emailAddress: 'user@company.com' });

      expect(result.content[0].text).toBe('Calendar ID is required (from list-calendars).');
    });

    test('should return error when emailAddress is missing', async () => {
      const result = await handleShareCalendar({ calendarId: 'cal-1' });

      expect(result.content[0].text).toBe('Email address of the person to share with is required.');
    });

    test('should validate role', async () => {
      const result = await handleShareCalendar({
        calendarId: 'cal-1',
        emailAddress: 'user@company.com',
        role: 'invalid',
      });

      expect(result.content[0].text).toContain('Invalid role');
    });

    test('should handle permission denied', async () => {
      const error = new Error('Forbidden');
      error.statusCode = 403;
      mockClient.post.mockRejectedValue(error);

      const result = await handleShareCalendar({
        calendarId: 'cal-1',
        emailAddress: 'user@company.com',
      });

      expect(result.content[0].text).toBe('Forbidden');
    });
  });

  describe('handleListCalendarPermissions', () => {
    test('should list calendar permissions', async () => {
      const mockPermissions = [
        { id: 'perm-1', emailAddress: { address: 'user@company.com', name: 'User One' }, role: 'read', isInsideOrganization: true, isRemovable: true },
        { id: 'perm-2', emailAddress: { address: 'external@partner.com', name: 'External User' }, role: 'freeBusyRead', isInsideOrganization: false, isRemovable: true },
      ];

      mockClient.get.mockResolvedValue({ value: mockPermissions });

      const result = await handleListCalendarPermissions({ calendarId: 'cal-1' });

      expect(result.content[0].text).toContain('Calendar Permissions (2)');
      expect(result.content[0].text).toContain('User One <user@company.com>: Full details (read) (internal) [removable]');
      expect(result.content[0].text).toContain('External User <external@partner.com>: Availability only (external) [removable]');
    });
  });

  describe('handleUpdateCalendarPermission', () => {
    test('should update calendar permission', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleUpdateCalendarPermission({
        calendarId: 'cal-1',
        permissionId: 'perm-1',
        role: 'write',
      });

      expect(result.content[0].text).toBe('Calendar permission updated to "write" successfully.');
    });

    test('should return error for missing parameters', async () => {
      const result = await handleUpdateCalendarPermission({ calendarId: 'cal-1' });

      expect(result.content[0].text).toContain('required');
    });
  });

  describe('handleRemoveCalendarPermission', () => {
    test('should remove calendar permission', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleRemoveCalendarPermission({
        calendarId: 'cal-1',
        permissionId: 'perm-1',
      });

      expect(result.content[0].text).toBe('Calendar permission removed successfully.');
    });
  });

  describe('handleGetSharedCalendars', () => {
    test('should get shared and owned calendars', async () => {
      const mockOwned = [
        { id: 'cal-1', name: 'Personal', color: 'auto', isDefaultCalendar: true, canEdit: true, canShare: true, canViewPrivateItems: true, owner: 'me@company.com' },
      ];
      const mockShared = [
        { id: 'cal-2', name: 'Team Calendar', color: 'lightBlue', owner: 'team@company.com' },
      ];

      mockClient.get
        .mockResolvedValueOnce({ value: mockOwned })
        .mockResolvedValueOnce({ value: [] }); // calendarGroups empty

      graphGetPaginated.mockResolvedValue({ value: mockShared });

      const result = await handleGetSharedCalendars({});

      expect(result.content[0].text).toContain('Owned Calendars (1)');
      expect(result.content[0].text).toContain('Personal (DEFAULT)');
      expect(result.content[0].text).toContain('Shared Calendars (0)');
      expect(result.content[0].text).toContain('None found');
    });
  });
});