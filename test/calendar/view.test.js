// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleGetCalendarView, handleGetMergedCalendarView } = require('../../calendar/view');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Calendar View Tools', () => {
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
      filter: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleGetCalendarView', () => {
    test('should get calendar view for date range', async () => {
      const mockEvents = [
        { id: 'event-1', subject: 'Team Meeting', bodyPreview: 'Weekly sync', start: { dateTime: '2024-01-15T10:00:00Z', timeZone: 'UTC' }, end: { dateTime: '2024-01-15T11:00:00Z', timeZone: 'UTC' }, location: { displayName: 'Room A' }, isAllDay: false, isCancelled: false, importance: 'normal', categories: ['Work'], showAs: 'busy', sensitivity: 'normal', organizer: { emailAddress: { address: 'organizer@company.com' } }, attendees: [{ emailAddress: { address: 'attendee@company.com', name: 'Attendee' }, status: { response: 'accepted' }, type: 'required' }], recurrence: null, onlineMeeting: null },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockEvents });

      const result = await handleGetCalendarView({
        startTime: '2024-01-15T00:00:00Z',
        endTime: '2024-01-16T00:00:00Z',
        calendarId: 'cal-1',
        count: 100,
      });

      expect(result.content[0].text).toContain('Calendar View (1 events)');
      expect(result.content[0].text).toContain('Team Meeting');
      expect(result.content[0].text).toContain('Room A');
    });

    test('should return error when startTime/endTime missing', async () => {
      const result = await handleGetCalendarView({ startTime: '2024-01-15T00:00:00Z' });

      expect(result.content[0].text).toBe('startTime and endTime are required (ISO 8601 format).');
    });
  });

  describe('handleGetMergedCalendarView', () => {
    test('should get merged view across multiple calendars', async () => {
      const mockCalendars = [
        { id: 'cal-1', name: 'Work', color: 'lightBlue' },
        { id: 'cal-2', name: 'Personal', color: 'lightGreen' },
      ];
      const mockEventsCal1 = [
        { id: 'event-1', subject: 'Work Meeting', bodyPreview: 'Work', start: { dateTime: '2024-01-15T10:00:00Z', timeZone: 'UTC' }, end: { dateTime: '2024-01-15T11:00:00Z', timeZone: 'UTC' }, location: { displayName: 'Office' }, isAllDay: false, isCancelled: false, importance: 'normal', categories: [], recurrence: null, onlineMeeting: null },
      ];
      const mockEventsCal2 = [
        { id: 'event-2', subject: 'Personal Appointment', bodyPreview: 'Personal', start: { dateTime: '2024-01-15T14:00:00Z', timeZone: 'UTC' }, end: { dateTime: '2024-01-15T15:00:00Z', timeZone: 'UTC' }, location: { displayName: 'Home' }, isAllDay: false, isCancelled: false, importance: 'normal', categories: [], recurrence: null, onlineMeeting: null },
      ];

      mockClient.get
        .mockResolvedValueOnce({ value: mockCalendars }) // list calendars
        .mockResolvedValueOnce({ id: 'cal-1', name: 'Work', color: 'lightBlue' }) // cal1 info
        .mockResolvedValueOnce({ id: 'cal-2', name: 'Personal', color: 'lightGreen' }); // cal2 info

      graphGetPaginated
        .mockResolvedValueOnce({ value: mockEventsCal1 })
        .mockResolvedValueOnce({ value: mockEventsCal2 });

      const result = await handleGetMergedCalendarView({
        startTime: '2024-01-15T00:00:00Z',
        endTime: '2024-01-16T00:00:00Z',
      });

      expect(result.content[0].text).toContain('Merged Calendar View (2 events from 2 calendars)');
      expect(result.content[0].text).toContain('Work Meeting');
      expect(result.content[0].text).toContain('Personal Appointment');
      expect(result.content[0].text).toContain('(Work)');
      expect(result.content[0].text).toContain('(Personal)');
    });
  });
});