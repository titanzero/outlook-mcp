/**
 * Tests for calendar update functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleUpdateEvent = require('../../calendar/update');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');
jest.mock('../../config', () => ({
  DEFAULT_TIMEZONE: 'Central European Standard Time',
}));

describe('handleUpdateEvent', () => {
  let mockClient;
  let mockPatch;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPatch = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn(() => ({
        patch: mockPatch,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful updates', () => {
    test('should update event subject', async () => {
      const result = await handleUpdateEvent({
        eventId: 'event-123',
        subject: 'Updated Subject',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/events/event-123');
      expect(mockPatch).toHaveBeenCalledWith({ subject: 'Updated Subject' });
      expect(result.content[0].text).toBe('Event updated successfully.');
    });

    test('should update event location', async () => {
      await handleUpdateEvent({
        eventId: 'event-123',
        location: 'New Conference Room',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        location: { displayName: 'New Conference Room' },
      });
    });

    test('should update event body', async () => {
      await handleUpdateEvent({
        eventId: 'event-123',
        body: 'Updated body content',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        body: { contentType: 'text', content: 'Updated body content' },
      });
    });

    test('should update isAllDay flag', async () => {
      await handleUpdateEvent({
        eventId: 'event-123',
        isAllDay: true,
      });

      expect(mockPatch).toHaveBeenCalledWith({ isAllDay: true });
    });

    test('should update start time', async () => {
      await handleUpdateEvent({
        eventId: 'event-123',
        start: '2024-01-20T14:00:00',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        start: { dateTime: '2024-01-20T14:00:00', timeZone: 'Central European Standard Time' },
      });
    });

    test('should update end time', async () => {
      await handleUpdateEvent({
        eventId: 'event-123',
        end: '2024-01-20T15:00:00',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        end: { dateTime: '2024-01-20T15:00:00', timeZone: 'Central European Standard Time' },
      });
    });

    test('should update attendees', async () => {
      await handleUpdateEvent({
        eventId: 'event-123',
        attendees: ['alice@example.com', 'bob@example.com'],
      });

      expect(mockPatch).toHaveBeenCalledWith({
        attendees: [
          { emailAddress: { address: 'alice@example.com' }, type: 'required' },
          { emailAddress: { address: 'bob@example.com' }, type: 'required' },
        ],
      });
    });

    test('should update multiple fields at once', async () => {
      await handleUpdateEvent({
        eventId: 'event-123',
        subject: 'New Subject',
        location: 'New Location',
        body: 'New Body',
        isAllDay: false,
        start: '2024-01-20T10:00:00',
        end: '2024-01-20T11:00:00',
        attendees: ['test@example.com'],
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.subject).toBe('New Subject');
      expect(patch.location).toEqual({ displayName: 'New Location' });
      expect(patch.body).toEqual({ contentType: 'text', content: 'New Body' });
      expect(patch.isAllDay).toBe(false);
      expect(patch.start).toEqual({
        dateTime: '2024-01-20T10:00:00',
        timeZone: 'Central European Standard Time',
      });
      expect(patch.end).toEqual({
        dateTime: '2024-01-20T11:00:00',
        timeZone: 'Central European Standard Time',
      });
      expect(patch.attendees).toHaveLength(1);
    });

    test('should not include undefined fields in patch', async () => {
      await handleUpdateEvent({
        eventId: 'event-123',
        subject: 'Only Subject',
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.subject).toBe('Only Subject');
      expect(patch.location).toBeUndefined();
      expect(patch.body).toBeUndefined();
      expect(patch.start).toBeUndefined();
      expect(patch.end).toBeUndefined();
      expect(patch.attendees).toBeUndefined();
    });
  });

  describe('validation', () => {
    test('should return error when eventId is missing', async () => {
      const result = await handleUpdateEvent({
        subject: 'Test',
      });

      expect(result.content[0].text).toBe('Event ID is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when eventId is empty string', async () => {
      const result = await handleUpdateEvent({ eventId: '' });

      expect(result.content[0].text).toBe('Event ID is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleUpdateEvent({ eventId: 'event-123' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPatch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleUpdateEvent({ eventId: 'event-123' });

      expect(result.content[0].text).toBe('Error updating event: Graph API Error');
    });
  });
});
