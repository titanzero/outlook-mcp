// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListTaskLinkedResources, handleAddTaskLinkedResource, handleRemoveTaskLinkedResource, handleLinkEmailToTask, handleLinkEventToTask } = require('../../tasks/linked-resources');
const { getGraphClient, graphGetPaginated } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Task Linked Resources Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
      delete: jest.fn(),
      top: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListTaskLinkedResources', () => {
    test('should list linked resources', async () => {
      const mockResources = [
        { id: 'res-1', webUrl: 'https://outlook.office.com/mail/id/123', applicationName: 'Outlook', displayName: 'Important Email', externalId: 'message:123' },
        { id: 'res-2', webUrl: 'https://teams.microsoft.com/l/meeting/456', applicationName: 'Teams', displayName: 'Team Meeting', externalId: 'event:456' },
      ];

      graphGetPaginated.mockResolvedValue({ value: mockResources });

      const result = await handleListTaskLinkedResources({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toContain('Task Linked Resources (2)');
      expect(result.content[0].text).toContain('Important Email (Outlook)');
      expect(result.content[0].text).toContain('Team Meeting (Teams)');
    });
  });

  describe('handleAddTaskLinkedResource', () => {
    test('should add linked resource', async () => {
      mockClient.post.mockResolvedValue({ id: 'res-new', displayName: 'New Link' });

      const result = await handleAddTaskLinkedResource({
        listId: 'list-1',
        taskId: 'task-1',
        webUrl: 'https://example.com',
        applicationName: 'Example',
        displayName: 'Example Link',
        externalId: 'ext-123',
      });

      expect(result.content[0].text).toContain('Linked resource added successfully');
      expect(result.content[0].text).toContain('New Link');
    });

    test('should return error for missing fields', async () => {
      const result = await handleAddTaskLinkedResource({ listId: 'list-1', taskId: 'task-1' });

      expect(result.content[0].text).toContain('required');
    });
  });

  describe('handleRemoveTaskLinkedResource', () => {
    test('should remove linked resource', async () => {
      mockClient.delete.mockResolvedValue({});

      const result = await handleRemoveTaskLinkedResource({ listId: 'list-1', taskId: 'task-1', linkedResourceId: 'res-1' });

      expect(result.content[0].text).toBe('Linked resource removed successfully.');
    });
  });

  describe('handleLinkEmailToTask', () => {
    test('should link email to task', async () => {
      const mockEmail = { id: 'email-1', subject: 'Important Email', webLink: 'https://outlook.office.com/mail/id/email-1' };

      mockClient.get.mockResolvedValue(mockEmail);
      mockClient.post.mockResolvedValue({ id: 'res-new', displayName: 'Important Email' });

      const result = await handleLinkEmailToTask({ listId: 'list-1', taskId: 'task-1', emailId: 'email-1' });

      expect(result.content[0].text).toContain('Email linked to task successfully');
      expect(result.content[0].text).toContain('Important Email');
    });
  });

  describe('handleLinkEventToTask', () => {
    test('should link event to task', async () => {
      const mockEvent = { id: 'event-1', subject: 'Team Meeting', webLink: 'https://outlook.office.com/calendar/id/event-1' };

      mockClient.get.mockResolvedValue(mockEvent);
      mockClient.post.mockResolvedValue({ id: 'res-new', displayName: 'Team Meeting' });

      const result = await handleLinkEventToTask({ listId: 'list-1', taskId: 'task-1', eventId: 'event-1' });

      expect(result.content[0].text).toContain('Event linked to task successfully');
      expect(result.content[0].text).toContain('Team Meeting');
    });
  });
});