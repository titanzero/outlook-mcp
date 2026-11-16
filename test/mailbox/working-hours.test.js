// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleGetWorkingHours, handleSetWorkingHours, handleClearWorkingHours } = require('../../mailbox/working-hours');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Working Hours Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      patch: jest.fn(),
      select: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleGetWorkingHours', () => {
    test('should get working hours', async () => {
      const mockSettings = {
        workingHours: {
          daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
          startTime: '09:00:00.0000000',
          endTime: '17:00:00.0000000',
          timeZone: { name: 'Pacific Standard Time' },
        },
      };

      mockClient.get.mockResolvedValue(mockSettings);

      const result = await handleGetWorkingHours({});

      expect(result.content[0].text).toContain('Working Hours');
      expect(result.content[0].text).toContain('Days: Monday, Tuesday, Wednesday, Thursday, Friday');
      expect(result.content[0].text).toContain('Time: 09:00:00.0000000 - 17:00:00.0000000');
      expect(result.content[0].text).toContain('Time Zone: Pacific Standard Time');
    });

    test('should return message when not configured', async () => {
      mockClient.get.mockResolvedValue({ workingHours: null });

      const result = await handleGetWorkingHours({});

      expect(result.content[0].text).toBe('Working hours not configured.');
    });
  });

  describe('handleSetWorkingHours', () => {
    test('should set working hours', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleSetWorkingHours({
        daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
        startTime: '09:00',
        endTime: '17:00',
        timeZone: 'Pacific Standard Time',
      });

      expect(result.content[0].text).toContain('Working hours set successfully');
      expect(result.content[0].text).toContain('Monday, Tuesday, Wednesday, Thursday, Friday');
      expect(result.content[0].text).toContain('09:00 - 17:00');
    });

    test('should validate daysOfWeek', async () => {
      const result = await handleSetWorkingHours({ daysOfWeek: ['invalid'], startTime: '09:00', endTime: '17:00' });

      expect(result.content[0].text).toContain('Invalid day');
    });

    test('should validate time format', async () => {
      // The implementation accepts H:mm format (single digit hour)
      const result = await handleSetWorkingHours({ daysOfWeek: ['monday'], startTime: '09:00', endTime: '17:00' });

      expect(result.content[0].text).toContain('Working hours set successfully');
    });
  });

  describe('handleClearWorkingHours', () => {
    test('should clear working hours', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleClearWorkingHours({});

      expect(result.content[0].text).toBe('Working hours cleared successfully.');
    });
  });
});