/**
 * Tests for prompts
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { registerPrompts } = require('../../prompts/index');
const { z } = require('zod');

jest.mock('@modelcontextprotocol/sdk/server/mcp.js');

describe('registerPrompts', () => {
  let mockServer;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});

    mockServer = {
      prompt: jest.fn(),
    };
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should register 7 prompts', () => {
    registerPrompts(mockServer);
    expect(mockServer.prompt).toHaveBeenCalledTimes(7);
  });

  const expectedPrompts = [
    'summarize-inbox',
    'schedule-meeting',
    'triage-email',
    'daily-briefing',
    'weekly-review',
    'follow-up-draft',
    'enrich-contact',
  ];

  test.each(expectedPrompts)('should register %s prompt', async (promptName) => {
    registerPrompts(mockServer);
    const call = mockServer.prompt.mock.calls.find(c => c[0] === promptName);
    expect(call).toBeDefined();
    expect(call[1]).toBeDefined(); // description
    expect(call[2]).toBeDefined(); // arguments schema
    expect(typeof call[3]).toBe('function'); // handler
  });

  test('summarize-inbox should have correct arguments', () => {
    registerPrompts(mockServer);
    const call = mockServer.prompt.mock.calls.find(c => c[0] === 'summarize-inbox');
    const argsSchema = call[2];
    // Zod object schema - check keys
    expect(argsSchema.shape || argsSchema).toBeDefined();
    const schemaObj = argsSchema.shape || argsSchema;
    expect(schemaObj.folder).toBeDefined();
    expect(schemaObj.count).toBeDefined();
    expect(schemaObj.timeRange).toBeDefined();
  });

  test('schedule-meeting should have correct arguments', () => {
    registerPrompts(mockServer);
    const call = mockServer.prompt.mock.calls.find(c => c[0] === 'schedule-meeting');
    const argsSchema = call[2];
    const schemaObj = argsSchema.shape || argsSchema;
    expect(schemaObj.attendees).toBeDefined();
    expect(schemaObj.durationMinutes).toBeDefined();
    expect(schemaObj.preferredTimes).toBeDefined();
    expect(schemaObj.subject).toBeDefined();
    expect(schemaObj.description).toBeDefined();
  });

  test('triage-email should have emailId argument', () => {
    registerPrompts(mockServer);
    const call = mockServer.prompt.mock.calls.find(c => c[0] === 'triage-email');
    const argsSchema = call[2];
    const schemaObj = argsSchema.shape || argsSchema;
    expect(schemaObj.emailId).toBeDefined();
  });

  test('daily-briefing should have date, includeTasks, includeEmails arguments', () => {
    registerPrompts(mockServer);
    const call = mockServer.prompt.mock.calls.find(c => c[0] === 'daily-briefing');
    const argsSchema = call[2];
    const schemaObj = argsSchema.shape || argsSchema;
    expect(schemaObj.date).toBeDefined();
    expect(schemaObj.includeTasks).toBeDefined();
    expect(schemaObj.includeEmails).toBeDefined();
  });

  test('weekly-review should have weekStart argument', () => {
    registerPrompts(mockServer);
    const call = mockServer.prompt.mock.calls.find(c => c[0] === 'weekly-review');
    const argsSchema = call[2];
    const schemaObj = argsSchema.shape || argsSchema;
    expect(schemaObj.weekStart).toBeDefined();
  });

  test('follow-up-draft should have emailId, tone, context arguments', () => {
    registerPrompts(mockServer);
    const call = mockServer.prompt.mock.calls.find(c => c[0] === 'follow-up-draft');
    const argsSchema = call[2];
    const schemaObj = argsSchema.shape || argsSchema;
    expect(schemaObj.emailId).toBeDefined();
    expect(schemaObj.tone).toBeDefined();
    expect(schemaObj.context).toBeDefined();
  });

  test('enrich-contact should have contactId argument', () => {
    registerPrompts(mockServer);
    const call = mockServer.prompt.mock.calls.find(c => c[0] === 'enrich-contact');
    const argsSchema = call[2];
    const schemaObj = argsSchema.shape || argsSchema;
    expect(schemaObj.contactId).toBeDefined();
  });
});
