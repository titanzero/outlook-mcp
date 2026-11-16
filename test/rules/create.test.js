/**
 * Tests for rules create functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

// Use isolateModules to avoid mock interference
jest.isolateModules(() => {
  // Mock dependencies
  jest.mock('../../utils/graph-client');
  jest.mock('../../email/folder-utils', () => ({
    getFolderIdByName: jest.fn(),
  }));
  jest.mock('../../rules/list', () => ({
    getInboxRules: jest.fn(),
  }));
  jest.mock('../../config', () => ({
    DEFAULT_RULE_SEQUENCE: 100,
  }));

  const handleCreateRule = require('../../rules/create');
  const { getGraphClient } = require('../../utils/graph-client');
  const { getFolderIdByName } = require('../../email/folder-utils');
  const { getInboxRules } = require('../../rules/list');

  describe('handleCreateRule', () => {
    let mockClient;
    let mockPost;

    beforeEach(() => {
      jest.clearAllMocks();
      jest.spyOn(console, 'error').mockImplementation(() => {});
      mockPost = jest.fn().mockResolvedValue({ id: 'new-rule-id' });
      mockClient = {
        api: jest.fn(() => ({
          post: mockPost,
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);
      getFolderIdByName.mockResolvedValue('target-folder-id');
      getInboxRules.mockResolvedValue([]);
    });

    afterEach(() => {
      console.error.mockRestore();
    });

    describe('successful rule creation', () => {
      test('should create rule with fromAddresses and moveToFolder', async () => {
        const result = await handleCreateRule({
          name: 'Move GitHub emails',
          fromAddresses: 'notifications@github.com',
          moveToFolder: 'GitHub',
        });

        expect(getGraphClient).toHaveBeenCalledTimes(1);
        expect(getFolderIdByName).toHaveBeenCalledWith('GitHub');
        expect(mockClient.api).toHaveBeenCalledWith('me/mailFolders/inbox/messageRules');
        expect(mockPost).toHaveBeenCalledWith(
          expect.objectContaining({
            displayName: 'Move GitHub emails',
            isEnabled: true,
            sequence: 100,
            conditions: {
              fromAddresses: [{ emailAddress: { address: 'notifications@github.com' } }],
            },
            actions: { moveToFolder: 'target-folder-id' },
          })
        );

        expect(result.content[0].text).toContain('Successfully created rule "Move GitHub emails" with sequence 100.');
      });

      test('should create rule with containsSubject and markAsRead', async () => {
        const result = await handleCreateRule({
          name: 'Mark newsletters as read',
          containsSubject: 'newsletter',
          markAsRead: true,
        });

        expect(mockPost).toHaveBeenCalledWith(
          expect.objectContaining({
            displayName: 'Mark newsletters as read',
            conditions: { subjectContains: ['newsletter'] },
            actions: { markAsRead: true },
          })
        );
        expect(result.content[0].text).toContain('Successfully created rule');
      });

      test('should create rule with hasAttachments', async () => {
        await handleCreateRule({
          name: 'Move attachments',
          hasAttachments: true,
          moveToFolder: 'Attachments',
        });

        expect(mockPost).toHaveBeenCalledWith(
          expect.objectContaining({
            conditions: { hasAttachment: true },
          })
        );
      });

      test('should create rule with multiple fromAddresses', async () => {
        await handleCreateRule({
          name: 'Multiple senders',
          fromAddresses: 'sender1@example.com, sender2@example.com, sender3@example.com',
          moveToFolder: 'Filtered',
        });

        expect(mockPost).toHaveBeenCalledWith(
          expect.objectContaining({
            conditions: {
              fromAddresses: [
                { emailAddress: { address: 'sender1@example.com' } },
                { emailAddress: { address: 'sender2@example.com' } },
                { emailAddress: { address: 'sender3@example.com' } },
              ],
            },
          })
        );
      });

      test('should create rule with custom sequence', async () => {
        const result = await handleCreateRule({
          name: 'Custom sequence rule',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
          sequence: 5,
        });

        expect(mockPost).toHaveBeenCalledWith(
          expect.objectContaining({ sequence: 5 })
        );
        expect(result.content[0].text).toContain('with sequence 5');
      });

      test('should create rule with isEnabled false', async () => {
        await handleCreateRule({
          name: 'Disabled rule',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
          isEnabled: false,
        });

        expect(mockPost).toHaveBeenCalledWith(
          expect.objectContaining({ isEnabled: false })
        );
      });

      test('should use DEFAULT_RULE_SEQUENCE when no existing rules', async () => {
        getInboxRules.mockImplementation(() => Promise.resolve([]));

        await handleCreateRule({
          name: 'First rule',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
        });

        expect(getInboxRules).toHaveBeenCalledWith(mockClient);
        expect(mockPost).toHaveBeenCalledWith(
          expect.objectContaining({ sequence: 100 })
        );
      });
    });

    describe('validation', () => {
      test('should return error when name is missing', async () => {
        const result = await handleCreateRule({ fromAddresses: 'test@example.com' });
        expect(result.content[0].text).toBe('Rule name is required.');
        expect(result.isError).toBe(true);
      });

      test('should return error when no condition specified', async () => {
        const result = await handleCreateRule({
          name: 'No condition',
          moveToFolder: 'Test',
        });
        expect(result.content[0].text).toContain('At least one condition is required');
        expect(result.isError).toBe(true);
      });

      test('should return error when no action specified', async () => {
        const result = await handleCreateRule({
          name: 'No action',
          fromAddresses: 'test@example.com',
        });
        expect(result.content[0].text).toContain('At least one action is required');
        expect(result.isError).toBe(true);
      });

      test('should return error when sequence is invalid', async () => {
        const result = await handleCreateRule({
          name: 'Invalid sequence',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
          sequence: 0,
        });
        expect(result.content[0].text).toBe('Sequence must be a positive number greater than zero.');
        expect(result.isError).toBe(true);
      });

      test('should return error when sequence is negative', async () => {
        const result = await handleCreateRule({
          name: 'Negative sequence',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
          sequence: -5,
        });
        expect(result.content[0].text).toBe('Sequence must be a positive number greater than zero.');
      });
    });

    describe('error handling', () => {
      test('should handle authentication error', async () => {
        const authError = new Error('Authentication required');
        authError.isAuthError = true;
        getGraphClient.mockRejectedValue(authError);

        const result = await handleCreateRule({
          name: 'Test',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
        });
        expect(result.content[0].text).toBe('Authentication required');
      });

      test('should return error when target folder not found', async () => {
        getFolderIdByName.mockResolvedValue(null);

        const result = await handleCreateRule({
          name: 'Test',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Non-existent',
        });
        expect(result.content[0].text).toContain('Target folder "Non-existent" not found');
      });

      test('should return error when folder resolution fails', async () => {
        getFolderIdByName.mockRejectedValue(new Error('Folder resolution failed'));

        const result = await handleCreateRule({
          name: 'Test',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
        });
        expect(result.content[0].text).toContain('Error resolving folder');
      });

      test('should handle Graph API error during rule creation', async () => {
        mockPost.mockRejectedValue(new Error('Graph API Error'));

        const result = await handleCreateRule({
          name: 'Test',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
        });
        expect(result.content[0].text).toBe('Error creating rule: Graph API Error');
      });

      test('should handle error when server does not return rule ID', async () => {
        mockPost.mockResolvedValue({});

        const result = await handleCreateRule({
          name: 'Test',
          fromAddresses: 'test@example.com',
          moveToFolder: 'Test',
        });
        expect(result.content[0].text).toContain("The server didn't return a rule ID");
      });
    });
  });
});
