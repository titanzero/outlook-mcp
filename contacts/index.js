const handleListContacts = require('./list');
const handleCreateContact = require('./create');
const handleUpdateContact = require('./update');
const handleDeleteContact = require('./delete');

const contactsTools = [
  {
    name: "list-contacts",
    description: "Lists contacts from your Outlook address book. Optionally filter by name or email.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Filter by name or email prefix" },
        count: { type: "number", description: "Number of contacts to retrieve (default: 25)" }
      },
      required: []
    },
    handler: handleListContacts
  },
  {
    name: "create-contact",
    description: "Creates a new contact in your Outlook address book",
    inputSchema: {
      type: "object",
      properties: {
        displayName: { type: "string", description: "Full name of the contact" },
        email: { type: "string", description: "Email address" },
        phone: { type: "string", description: "Mobile phone number" },
        jobTitle: { type: "string", description: "Job title" },
        companyName: { type: "string", description: "Company name" }
      },
      required: ["displayName"]
    },
    handler: handleCreateContact
  },
  {
    name: "update-contact",
    description: "Updates an existing contact in your Outlook address book",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the contact to update (from list-contacts)" },
        displayName: { type: "string", description: "New full name" },
        email: { type: "string", description: "New email address" },
        phone: { type: "string", description: "New mobile phone number" },
        jobTitle: { type: "string", description: "New job title" },
        companyName: { type: "string", description: "New company name" }
      },
      required: ["id"]
    },
    handler: handleUpdateContact
  },
  {
    name: "delete-contact",
    description: "Deletes a contact from your Outlook address book",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "ID of the contact to delete (from list-contacts)" }
      },
      required: ["id"]
    },
    handler: handleDeleteContact
  }
];

module.exports = { contactsTools };
