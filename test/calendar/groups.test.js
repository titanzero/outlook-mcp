// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListCalendarGroups, handleCreateCalendarGroup, handleDeleteCalendarGroup, handleAddCalendarToGroup, handleRemoveCalendarFromGroup } = require('../../calendar/groups');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Calendar Groups Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      delete: jest.fn(),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListCalendarGroups', () => {
    test('should list calendar groups with calendars', async () => {
      const mockGroups = [
        { id: 'group-1', name: 'Work', classId: 'class-1' },
        { id: 'group-2', name: 'Personal', classId: 'class-2' },
      ];
      const mockCalendars = [
        { id: 'cal-1', name: 'Work Calendar', color: 'lightBlue', isDefaultCalendar: true },
      ];

      mockClient.get
        .mockResolvedValueOnce({ value: mockGroups })
        .mockResolvedValueOnce({ value: mockCalendars })
        .mockResolvedValueOnce({ value: [] });

      const result = await handleListCalendarGroups({});

      expect(result.content[0].text).toContain('Calendar Groups (2)');
      expect(result.content[0].text).toContain('Work');
      expect(result.content[0].text).toContain('Work Calendar [lightBlue] (DEFAULT)');
    });
  });

  describe('handleCreateCalendarGroup', () => {
    test('should create calendar group', async () => {
      mockClient.post.mockResolvedValue({ id: 'new-group-id', name: 'New Group' });

      const result = await handleCreateCalendarGroup({ name: 'New Group' });

      expect(result.content[0].text).toContain('Calendar group "New Group" created successfully');
      expect(result.content[0].text).toContain('ID: new-group-id');
    });

    test('should return error when name is missing', async () => {
      const result = await handleCreateCalendarGroup({});

      expect(result.content[0].text).toBe('Calendar group name is required.');
    });
  });

  describe('handleDeleteCalendarGroup', () => {
    test('should delete calendar group', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleDeleteCalendarGroup({ groupId: 'group-1' });

      expect(result.content[0].text).toBe('Calendar group deleted successfully. Note: Calendars in the group are not deleted, only the group itself.');
    });
  });

  describe('handleAddCalendarToGroup', () => {
    test('should add calendar to group', async () => {
      const mockApi = { 
        select: jest.fn().mockReturnThis(), 
        get: jest.fn().mockResolvedValue({ id: 'cal-1' }),
        post: jest.fn().mockResolvedValue({})
      };
      mockClient.api.mockReturnValue(mockApi);

      const result = await handleAddCalendarToGroup({ groupId: 'group-1', calendarId: 'cal-1' });

      expect(result.content[0].text).toContain('Calendar added to group successfully');
    });
  });

  describe('handleRemoveCalendarFromGroup', () => {
    test('should remove calendar from group', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleRemoveCalendarFromGroup({ groupId: 'group-1', calendarId: 'cal-1' });

      expect(result.content[0].text).toBe('Calendar removed from group successfully.');
    });
  });
});