/**
 * Tests for contacts delete functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleDeleteContact = require('../../contacts/delete');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleDeleteContact', () => {
  let mockClient;
  let mockDelete;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockDelete = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn(() => ({
        delete: mockDelete,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful contact deletion', () => {
    test('should delete contact', async () => {
      const result = await handleDeleteContact({ id: 'contact-123' });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/contacts/contact-123');
      expect(mockDelete).toHaveBeenCalledTimes(1);

      expect(result.content[0].text).toBe('Contact deleted successfully.');
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleDeleteContact({});

      expect(result.content[0].text).toBe('Contact ID is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when id is empty string', async () => {
      const result = await handleDeleteContact({ id: '' });

      expect(result.content[0].text).toBe('Contact ID is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleDeleteContact({ id: 'contact-123' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockDelete.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleDeleteContact({ id: 'contact-123' });

      expect(result.content[0].text).toBe('Error deleting contact: Graph API Error');
    });
  });
});
