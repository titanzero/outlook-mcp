// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleGetLanguageRegion, handleSetLanguage, handleSetDateFormat, handleSetTimeFormat, handleListLocales, handleListDateFormats, handleListTimeFormats } = require('../../mailbox/language-region');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Language & Region Tools', () => {
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

  describe('handleGetLanguageRegion', () => {
    test('should get language and region settings', async () => {
      const mockSettings = {
        language: { displayName: 'English (United States)', locale: 'en-US' },
        dateFormat: 'MM/dd/yyyy',
        timeFormat: 'h:mm tt',
      };

      mockClient.get.mockResolvedValue(mockSettings);

      const result = await handleGetLanguageRegion({});

      expect(result.content[0].text).toContain('Language & Region Settings');
      expect(result.content[0].text).toContain('Language: English (United States) (en-US)');
      expect(result.content[0].text).toContain('Date Format: MM/dd/yyyy');
      expect(result.content[0].text).toContain('Time Format: h:mm tt');
    });
  });

  describe('handleSetLanguage', () => {
    test('should set language', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleSetLanguage({ locale: 'fr-FR' });

      expect(result.content[0].text).toBe('Language set to fr-FR successfully.');
    });

    test('should return error when locale is missing', async () => {
      const result = await handleSetLanguage({});

      expect(result.content[0].text).toBe('Locale is required (e.g., "en-US", "fr-FR", "de-DE")');
    });
  });

  describe('handleSetDateFormat', () => {
    test('should set date format', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleSetDateFormat({ format: 'dd/MM/yyyy' });

      expect(result.content[0].text).toBe('Date format set to "dd/MM/yyyy" successfully.');
    });
  });

  describe('handleSetTimeFormat', () => {
    test('should set time format', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleSetTimeFormat({ format: 'HH:mm' });

      expect(result.content[0].text).toBe('Time format set to "HH:mm" successfully.');
    });
  });

  describe('handleListLocales', () => {
    test('should list common locales', async () => {
      const result = await handleListLocales({});

      expect(result.content[0].text).toContain('Common Locales (28)');
      expect(result.content[0].text).toContain('en-US: English (United States)');
      expect(result.content[0].text).toContain('fr-FR: French (France)');
    });
  });

  describe('handleListDateFormats', () => {
    test('should list date formats', async () => {
      const result = await handleListDateFormats({});

      expect(result.content[0].text).toContain('Common Date Formats');
      expect(result.content[0].text).toContain('MM/dd/yyyy (US (Month/Day/Year)): 12/31/2024');
    });
  });

  describe('handleListTimeFormats', () => {
    test('should list time formats', async () => {
      const result = await handleListTimeFormats({});

      expect(result.content[0].text).toContain('Common Time Formats');
      expect(result.content[0].text).toContain('H:mm (24-hour (HH:MM)): 14:30');
    });
  });
});