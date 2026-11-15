const handleListContacts = require('./list');
const handleCreateContact = require('./create');
const handleCreateContactFull = require('./create-full');
const handleUpdateContact = require('./update');
const handleDeleteContact = require('./delete');
const { registerContactsResources } = require('./resources');
const { z } = require('zod');

/**
 * Register contacts tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerContactsTools(server) {
  server.tool(
    'list-contacts',
    'Lists contacts from your Outlook address book. Optionally filter by name or email.',
    {
      query: z.string().optional().describe('Filter by name or email prefix'),
      count: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('Number of contacts to retrieve (default: 25)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListContacts
  );

  server.tool(
    'create-contact',
    'Creates a new contact in your Outlook address book (basic fields)',
    {
      displayName: z.string().describe('Full name of the contact'),
      email: z.string().optional().describe('Email address'),
      phone: z.string().optional().describe('Mobile phone number'),
      jobTitle: z.string().optional().describe('Job title'),
      companyName: z.string().optional().describe('Company name'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateContact
  );

  server.tool(
    'create-contact-full',
    'Creates a new contact with all available fields including multiple emails, addresses, phones, birthday, etc.',
    {
      displayName: z.string().describe('Full name of the contact'),
      // Basic
      givenName: z.string().optional().describe('First name'),
      surname: z.string().optional().describe('Last name'),
      email: z.string().optional().describe('Primary email address (simple)'),
      emailAddresses: z.string().optional().describe("Multiple emails: 'addr1|name1,addr2|name2'"),
      phone: z.string().optional().describe('Mobile phone (simple)'),
      mobilePhone: z.string().optional().describe('Mobile phone'),
      businessPhones: z.string().optional().describe('Comma-separated business phones'),
      homePhones: z.string().optional().describe('Comma-separated home phones'),
      jobTitle: z.string().optional().describe('Job title'),
      companyName: z.string().optional().describe('Company name'),
      department: z.string().optional().describe('Department'),
      officeLocation: z.string().optional().describe('Office location'),
      profession: z.string().optional().describe('Profession'),
      // Addresses (format: 'street;city;state;postalCode;country')
      businessAddress: z
        .string()
        .optional()
        .describe("Business address: 'street;city;state;postalCode;country'"),
      homeAddress: z
        .string()
        .optional()
        .describe("Home address: 'street;city;state;postalCode;country'"),
      otherAddress: z
        .string()
        .optional()
        .describe("Other address: 'street;city;state;postalCode;country'"),
      // Personal
      birthday: z.string().optional().describe('Birthday in ISO 8601 date format (YYYY-MM-DD)'),
      spouseName: z.string().optional().describe('Spouse/partner name'),
      personalNotes: z.string().optional().describe('Personal notes'),
      manager: z.string().optional().describe('Manager name'),
      assistantName: z.string().optional().describe('Assistant name'),
      fileAs: z.string().optional().describe('File-as name for sorting'),
      initials: z.string().optional().describe('Initials'),
      generation: z.string().optional().describe('Generation suffix (e.g., Jr., Sr.)'),
      titles: z.string().optional().describe('Comma-separated titles'),
      imAddresses: z.string().optional().describe('Comma-separated IM addresses'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateContactFull
  );

  server.tool(
    'update-contact',
    'Updates an existing contact in your Outlook address book',
    {
      id: z.string().describe('ID of the contact to update (from list-contacts)'),
      displayName: z.string().optional().describe('New full name'),
      email: z.string().optional().describe('New email address'),
      phone: z.string().optional().describe('New mobile phone number'),
      jobTitle: z.string().optional().describe('New job title'),
      companyName: z.string().optional().describe('New company name'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateContact
  );

  server.tool(
    'delete-contact',
    'Deletes a contact from your Outlook address book',
    {
      id: z.string().describe('ID of the contact to delete (from list-contacts)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteContact
  );
}

module.exports = {
  registerContactsTools,
  registerContactsResources,
  handleListContacts,
  handleCreateContact,
  handleCreateContactFull,
  handleUpdateContact,
  handleDeleteContact,
};
