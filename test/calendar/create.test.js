const handleCreateEvent = require('../../calendar/create');
const { DEFAULT_TIMEZONE } = require('../../config');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleCreateEvent', () => {
  let mockPost;
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    mockPost = jest.fn().mockResolvedValue({ id: 'test_event_id' });
    mockClient = {
      api: jest.fn(() => ({
        post: mockPost
      }))
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  test('should use default timezone when no timezone is provided', async () => {
    const args = {
      subject: 'Test Event',
      start: '2024-03-10T10:00:00',
      end: '2024-03-10T11:00:00',
    };

    await handleCreateEvent(args);

    expect(getGraphClient).toHaveBeenCalledTimes(1);
    expect(mockClient.api).toHaveBeenCalledWith('me/events');
    expect(mockPost).toHaveBeenCalledTimes(1);
    const bodyContent = mockPost.mock.calls[0][0];
    expect(bodyContent.start.timeZone).toBe(DEFAULT_TIMEZONE);
    expect(bodyContent.end.timeZone).toBe(DEFAULT_TIMEZONE);
  });

  test('should use specified timezone when provided', async () => {
    const specifiedTimeZone = 'Pacific Standard Time';
    const args = {
      subject: 'Test Event with Specific Timezone',
      start: { dateTime: '2024-03-10T10:00:00', timeZone: specifiedTimeZone },
      end: { dateTime: '2024-03-10T11:00:00', timeZone: specifiedTimeZone },
    };

    await handleCreateEvent(args);

    expect(mockPost).toHaveBeenCalledTimes(1);
    const bodyContent = mockPost.mock.calls[0][0];
    expect(bodyContent.start.timeZone).toBe(specifiedTimeZone);
    expect(bodyContent.end.timeZone).toBe(specifiedTimeZone);
  });

  test('should use default timezone if only start timezone is provided', async () => {
    const specifiedTimeZone = 'Pacific Standard Time';
    const args = {
      subject: 'Test Event with Specific Start Timezone',
      start: { dateTime: '2024-03-10T10:00:00', timeZone: specifiedTimeZone },
      end: { dateTime: '2024-03-10T11:00:00' },
    };

    await handleCreateEvent(args);

    const bodyContent = mockPost.mock.calls[0][0];
    expect(bodyContent.start.timeZone).toBe(specifiedTimeZone);
    expect(bodyContent.end.timeZone).toBe(DEFAULT_TIMEZONE);
  });

  test('should use default timezone if only end timezone is provided', async () => {
    const specifiedTimeZone = 'Pacific Standard Time';
    const args = {
      subject: 'Test Event with Specific End Timezone',
      start: { dateTime: '2024-03-10T10:00:00' },
      end: { dateTime: '2024-03-10T11:00:00', timeZone: specifiedTimeZone },
    };

    await handleCreateEvent(args);

    const bodyContent = mockPost.mock.calls[0][0];
    expect(bodyContent.start.timeZone).toBe(DEFAULT_TIMEZONE);
    expect(bodyContent.end.timeZone).toBe(specifiedTimeZone);
  });

  test('should return error if subject is missing', async () => {
    const args = {
      start: '2024-03-10T10:00:00',
      end: '2024-03-10T11:00:00',
    };

    const result = await handleCreateEvent(args);
    expect(result.content[0].text).toBe("Subject, start, and end times are required to create an event.");
    expect(getGraphClient).not.toHaveBeenCalled();
    expect(mockClient.api).not.toHaveBeenCalled();
  });

  test('should return error if start is missing', async () => {
    const args = {
      subject: 'Test Event',
      end: '2024-03-10T11:00:00',
    };

    const result = await handleCreateEvent(args);
    expect(result.content[0].text).toBe("Subject, start, and end times are required to create an event.");
    expect(getGraphClient).not.toHaveBeenCalled();
  });

  test('should return error if end is missing', async () => {
    const args = {
      subject: 'Test Event',
      start: '2024-03-10T10:00:00',
    };

    const result = await handleCreateEvent(args);
    expect(result.content[0].text).toBe("Subject, start, and end times are required to create an event.");
    expect(getGraphClient).not.toHaveBeenCalled();
  });

  test('should handle authentication error', async () => {
    const authError = new Error('Authentication required');
    authError.isAuthError = true;
    getGraphClient.mockRejectedValue(authError);

    const args = {
      subject: 'Test Event',
      start: '2024-03-10T10:00:00',
      end: '2024-03-10T11:00:00',
    };

    const result = await handleCreateEvent(args);
    expect(result.content[0].text).toBe('Authentication required');
    expect(mockClient.api).not.toHaveBeenCalled();
  });

  test('should handle Graph API call error', async () => {
    mockPost.mockRejectedValue(new Error('Graph API Error'));
    const args = {
      subject: 'Test Event',
      start: '2024-03-10T10:00:00',
      end: '2024-03-10T11:00:00',
    };

    const result = await handleCreateEvent(args);
    expect(result.content[0].text).toBe("Error creating event: Graph API Error");
  });
});
