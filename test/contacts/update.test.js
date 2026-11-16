/**
 * Tests for contacts update functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleUpdateContact = require('../../contacts/update');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleUpdateContact', () => {
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

  describe('successful contact updates', () => {
    test('should update contact displayName', async () => {
      const result = await handleUpdateContact({
        id: 'contact-123',
        displayName: 'Jane Doe',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/contacts/contact-123');
      expect(mockPatch).toHaveBeenCalledWith({ displayName: 'Jane Doe' });
      expect(result.content[0].text).toBe('Contact updated successfully.');
    });

    test('should update contact email', async () => {
      await handleUpdateContact({
        id: 'contact-123',
        email: 'jane@example.com',
        displayName: 'Jane Doe',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        displayName: 'Jane Doe',
        emailAddresses: [{ address: 'jane@example.com', name: 'Jane Doe' }],
      });
    });

    test('should update contact phone', async () => {
      await handleUpdateContact({
        id: 'contact-123',
        phone: '+1-555-9999',
      });

      expect(mockPatch).toHaveBeenCalledWith({ mobilePhone: '+1-555-9999' });
    });

    test('should update contact jobTitle', async () => {
      await handleUpdateContact({
        id: 'contact-123',
        jobTitle: 'Senior Engineer',
      });

      expect(mockPatch).toHaveBeenCalledWith({ jobTitle: 'Senior Engineer' });
    });

    test('should update contact companyName', async () => {
      await handleUpdateContact({
        id: 'contact-123',
        companyName: 'New Corp',
      });

      expect(mockPatch).toHaveBeenCalledWith({ companyName: 'New Corp' });
    });

    test('should update multiple fields at once', async () => {
      await handleUpdateContact({
        id: 'contact-123',
        displayName: 'Jane Smith',
        email: 'jane@example.com',
        phone: '+1-555-0000',
        jobTitle: 'Manager',
        companyName: 'New Company',
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.displayName).toBe('Jane Smith');
      expect(patch.emailAddresses).toEqual([{ address: 'jane@example.com', name: 'Jane Smith' }]);
      expect(patch.mobilePhone).toBe('+1-555-0000');
      expect(patch.jobTitle).toBe('Manager');
      expect(patch.companyName).toBe('New Company');
    });

    test('should only include provided fields in patch', async () => {
      await handleUpdateContact({
        id: 'contact-123',
        displayName: 'Only Name',
      });

      const patch = mockPatch.mock.calls[0][0];
      expect(patch.displayName).toBe('Only Name');
      expect(patch.emailAddresses).toBeUndefined();
      expect(patch.mobilePhone).toBeUndefined();
      expect(patch.jobTitle).toBeUndefined();
      expect(patch.companyName).toBeUndefined();
    });
  });

  describe('validation', () => {
    test('should return error when id is missing', async () => {
      const result = await handleUpdateContact({ displayName: 'Test' });

      expect(result.content[0].text).toBe('Contact ID is required.');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when id is empty string', async () => {
      const result = await handleUpdateContact({ id: '', displayName: 'Test' });

      expect(result.content[0].text).toBe('Contact ID is required.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleUpdateContact({ id: 'contact-123', displayName: 'Test' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPatch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleUpdateContact({ id: 'contact-123', displayName: 'Test' });

      expect(result.content[0].text).toBe('Error updating contact: Graph API Error');
    });
  });
});
