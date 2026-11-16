const handleListContacts = require('./list');
const handleCreateContact = require('./create');
const handleCreateContactFull = require('./create-full');
const handleUpdateContact = require('./update');
const handleDeleteContact = require('./delete');
const { registerContactsResources } = require('./resources');
const { handleListContactFolders, handleGetContactFolder, handleCreateContactFolder, handleUpdateContactFolder, handleDeleteContactFolder, handleMoveContactToFolder } = require('./folders');
const { handleGetContactPhoto, handleSetContactPhoto, handleDeleteContactPhoto } = require('./photos');
const { handleListContactLists, handleCreateContactList, handleUpdateContactList, handleDeleteContactList, handleSendEmailToContactList } = require('./lists');
const { handleSearchContacts, handleGetContactStats } = require('./search');
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

  // Contact Folders tools (requires Contacts.ReadWrite)
  server.tool(
    'list-contact-folders',
    'Lists all contact folders in your address book',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListContactFolders
  );

  server.tool(
    'get-contact-folder',
    'Gets a contact folder with its contacts',
    {
      folderId: z.string().describe('Folder ID (from list-contact-folders)'),
      count: z.number().int().positive().max(1000).optional().describe('Number of contacts to retrieve (default: 50, max: 1000)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetContactFolder
  );

  server.tool(
    'create-contact-folder',
    'Creates a new contact folder',
    {
      name: z.string().describe('Name of the folder to create'),
      parentFolderId: z.string().optional().describe('Optional parent folder ID (default: root)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateContactFolder
  );

  server.tool(
    'update-contact-folder',
    'Updates a contact folder (rename or move)',
    {
      folderId: z.string().describe('Folder ID (from list-contact-folders)'),
      name: z.string().optional().describe('New folder name'),
      parentFolderId: z.string().optional().describe('New parent folder ID (to move)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateContactFolder
  );

  server.tool(
    'delete-contact-folder',
    'Deletes a contact folder (must be empty)',
    {
      folderId: z.string().describe('Folder ID (from list-contact-folders)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteContactFolder
  );

  server.tool(
    'move-contact-to-folder',
    'Moves a contact to a different folder',
    {
      contactId: z.string().describe('Contact ID (from list-contacts)'),
      folderId: z.string().describe('Destination folder ID (from list-contact-folders)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleMoveContactToFolder
  );

  // Contact Photos tools (requires Contacts.ReadWrite)
  server.tool(
    'get-contact-photo',
    'Gets a contact\'s photo as base64-encoded image',
    {
      contactId: z.string().describe('Contact ID (from list-contacts)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetContactPhoto
  );

  server.tool(
    'set-contact-photo',
    'Sets or updates a contact\'s photo',
    {
      contactId: z.string().describe('Contact ID (from list-contacts)'),
      imageBase64: z.string().describe('Base64-encoded image data (max 4MB, JPEG/PNG/GIF/BMP)'),
      contentType: z.enum(['image/jpeg', 'image/png', 'image/gif', 'image/bmp']).optional().describe('MIME type (default: image/jpeg)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleSetContactPhoto
  );

  server.tool(
    'delete-contact-photo',
    'Deletes a contact\'s photo',
    {
      contactId: z.string().describe('Contact ID (from list-contacts)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteContactPhoto
  );

  // Contact Lists / Distribution Lists tools (requires Contacts.ReadWrite)
  server.tool(
    'list-contact-lists',
    'Lists contact lists (distribution lists) in your contacts',
    {
      folderId: z.string().optional().describe('Optional folder ID to search in (from list-contact-folders)'),
      count: z.number().int().positive().max(100).optional().describe('Number of lists to retrieve (default: 50)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListContactLists
  );

  server.tool(
    'create-contact-list',
    'Creates a new contact list (distribution list) with members',
    {
      name: z.string().describe('Name of the contact list'),
      memberIds: z.array(z.string()).describe('Array of contact IDs to add as members (from list-contacts)'),
      email: z.string().optional().describe('Optional email address for the list'),
      folderId: z.string().optional().describe('Optional folder ID to create in (from list-contact-folders)'),
      companyName: z.string().optional().describe('Optional company name'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateContactList
  );

  server.tool(
    'update-contact-list',
    'Updates a contact list (add/remove members, rename)',
    {
      listId: z.string().describe('Contact list ID (from list-contact-lists)'),
      name: z.string().optional().describe('New name for the list'),
      addMemberIds: z.array(z.string()).optional().describe('Contact IDs to add as members'),
      removeMemberIds: z.array(z.string()).optional().describe('Contact IDs to remove from members'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateContactList
  );

  server.tool(
    'delete-contact-list',
    'Deletes a contact list',
    {
      listId: z.string().describe('Contact list ID (from list-contact-lists)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteContactList
  );

  server.tool(
    'send-email-to-contact-list',
    'Sends an email to all members of a contact list',
    {
      listId: z.string().describe('Contact list ID (from list-contact-lists)'),
      subject: z.string().describe('Email subject'),
      body: z.string().describe('Email body content (HTML)'),
      importance: z.enum(['normal', 'high', 'low']).optional().describe('Email importance (default: normal)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: true,
      },
    },
    handleSendEmailToContactList
  );

  // Advanced Contact Search tools
  server.tool(
    'search-contacts',
    'Searches contacts with advanced filters (company, job title, email domain, etc.)',
    {
      query: z.string().optional().describe('Search text in name, email'),
      folderId: z.string().optional().describe('Optional folder ID to search in (from list-contact-folders)'),
      companyName: z.string().optional().describe('Filter by company name'),
      jobTitle: z.string().optional().describe('Filter by job title'),
      emailDomain: z.string().optional().describe('Filter by email domain (e.g., "microsoft.com")'),
      count: z.number().int().positive().max(1000).optional().describe('Number of results (default: 50, max: 1000)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleSearchContacts
  );

  server.tool(
    'get-contact-stats',
    'Gets statistics about your contacts (counts, top companies, email domains, folder breakdown)',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetContactStats
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
  handleListContactFolders,
  handleGetContactFolder,
  handleCreateContactFolder,
  handleUpdateContactFolder,
  handleDeleteContactFolder,
  handleMoveContactToFolder,
  handleGetContactPhoto,
  handleSetContactPhoto,
  handleDeleteContactPhoto,
  handleListContactLists,
  handleCreateContactList,
  handleUpdateContactList,
  handleDeleteContactList,
  handleSendEmailToContactList,
  handleSearchContacts,
  handleGetContactStats,
};
