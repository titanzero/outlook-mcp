const handleListContacts = require('./list');
const handleCreateContact = require('./create');
const handleUpdateContact = require('./update');
const handleDeleteContact = require('./delete');
const { z } = require('zod');

/**
 * Register contacts tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerContactsTools(server) {
  server.tool(
    "list-contacts",
    "Lists contacts from your Outlook address book. Optionally filter by name or email.",
    {
      query: z.string().optional().describe("Filter by name or email prefix"),
      count: z.number().int().positive().optional().describe("Number of contacts to retrieve (default: 25)")
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false
      }
    },
    handleListContacts
  );

  server.tool(
    "create-contact",
    "Creates a new contact in your Outlook address book",
    {
      displayName: z.string().describe("Full name of the contact"),
      email: z.string().optional().describe("Email address"),
      phone: z.string().optional().describe("Mobile phone number"),
      jobTitle: z.string().optional().describe("Job title"),
      companyName: z.string().optional().describe("Company name")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false
      }
    },
    handleCreateContact
  );

  server.tool(
    "update-contact",
    "Updates an existing contact in your Outlook address book",
    {
      id: z.string().describe("ID of the contact to update (from list-contacts)"),
      displayName: z.string().optional().describe("New full name"),
      email: z.string().optional().describe("New email address"),
      phone: z.string().optional().describe("New mobile phone number"),
      jobTitle: z.string().optional().describe("New job title"),
      companyName: z.string().optional().describe("New company name")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false
      }
    },
    handleUpdateContact
  );

  server.tool(
    "delete-contact",
    "Deletes a contact from your Outlook address book",
    {
      id: z.string().describe("ID of the contact to delete (from list-contacts)")
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false
      }
    },
    handleDeleteContact
  );
}

module.exports = {
  registerContactsTools,
  handleListContacts,
  handleCreateContact,
  handleUpdateContact,
  handleDeleteContact
};