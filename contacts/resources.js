/**
 * Contacts resources for MCP
 * Exposes contacts and contact folders as addressable resources
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError } = require('../utils/response-helpers');
const { ResourceTemplate } = require("@modelcontextprotocol/sdk/server/mcp.js");
const { getContactCompletions } = require('../completion');

/**
 * Register contacts resources with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerContactsResources(server) {
  // Resource template for individual contact by ID with completion
  const contactTemplate = new ResourceTemplate(
    "outlook://contacts/{contactId}",
    {
      complete: {
        contactId: getContactCompletions
      }
    }
  );

  server.resource(
    "outlook-contact",
    contactTemplate,
    {
      title: "Contact",
      description: "Access a contact by ID"
    },
    async (uri, { contactId }) => {
      try {
        const client = await getGraphClient();
        const contact = await client.api(`me/contacts/${contactId}`)
          .select('id,displayName,givenName,surname,emailAddresses,businessPhones,homePhones,mobilePhone,jobTitle,companyName,department,officeLocation,businessAddress,homeAddress,otherAddress,birthday,spouseName,personalNotes,manager,assistantName,fileAs,initials,generation,titles,imAddresses')
          .get();
        
        const structured = {
          id: contact.id,
          displayName: contact.displayName,
          givenName: contact.givenName,
          surname: contact.surname,
          emails: contact.emailAddresses?.map(e => ({ address: e.address, name: e.name })) || [],
          businessPhones: contact.businessPhones || [],
          homePhones: contact.homePhones || [],
          mobilePhone: contact.mobilePhone,
          jobTitle: contact.jobTitle,
          companyName: contact.companyName,
          department: contact.department,
          officeLocation: contact.officeLocation,
          businessAddress: contact.businessAddress,
          homeAddress: contact.homeAddress,
          otherAddress: contact.otherAddress,
          birthday: contact.birthday,
          spouseName: contact.spouseName,
          personalNotes: contact.personalNotes,
          manager: contact.manager,
          assistantName: contact.assistantName,
          fileAs: contact.fileAs,
          initials: contact.initials,
          generation: contact.generation,
          titles: contact.titles || [],
          imAddresses: contact.imAddresses || []
        };
        
        const textFallback = [
          `Contact: ${contact.displayName}`,
          `ID: ${contact.id}`,
          contact.emailAddresses?.map(e => `Email: ${e.address} (${e.name || 'no name'})`).join('\n') || 'Email: (none)',
          contact.mobilePhone ? `Mobile: ${contact.mobilePhone}` : '',
          contact.businessPhones?.length ? `Business: ${contact.businessPhones.join(', ')}` : '',
          contact.jobTitle ? `Title: ${contact.jobTitle}` : '',
          contact.companyName ? `Company: ${contact.companyName}` : '',
          contact.department ? `Dept: ${contact.department}` : '',
          contact.birthday ? `Birthday: ${contact.birthday}` : ''
        ].filter(Boolean).join('\n');
        
        return {
          contents: [{
            uri: uri.toString(),
            mimeType: "application/json",
            text: JSON.stringify(structured, null, 2)
          }]
        };
      } catch (error) {
        if (isAuthError(error)) throw error;
        throw new Error(`Error reading contact: ${error.message}`);
      }
    }
  );

  // Resource for contact folders
  server.resource(
    "outlook-contact-folders",
    "outlook://contact-folders",
    {
      title: "Contact Folders",
      description: "List all contact folders"
    },
    async (uri) => {
      try {
        const client = await getGraphClient();
        const response = await client.api('me/contactFolders')
          .select('id,displayName,parentFolderId,contactCount')
          .get();
        
        const folders = response.value || [];
        
        const structured = folders.map(f => ({
          id: f.id,
          name: f.displayName,
          parentFolderId: f.parentFolderId,
          contactCount: f.contactCount || 0
        }));
        
        const textFallback = folders.length === 0 
          ? 'No contact folders found.'
          : `Contact Folders (${folders.length}):\n\n` + 
            folders.map(f => `- ${f.displayName} (${f.contactCount} contacts) [id: ${f.id}]`).join('\n');
        
        return {
          contents: [{
            uri: uri.toString(),
            mimeType: "application/json",
            text: JSON.stringify({ folders: structured }, null, 2)
          }]
        };
      } catch (error) {
        if (isAuthError(error)) throw error;
        throw new Error(`Error reading contact folders: ${error.message}`);
      }
    }
  );

  // Resource for individual contact folder
  server.resource(
    "outlook-contact-folder",
    "outlook://contact-folders/{folderId}",
    {
      title: "Contact Folder",
      description: "Access a contact folder by ID with its contacts"
    },
    async (uri, { folderId }) => {
      try {
        const client = await getGraphClient();
        
        // Get folder info
        const folder = await client.api(`me/contactFolders/${folderId}`).get();
        
        // Get contacts in this folder
        const response = await graphGetPaginated(client, `me/contactFolders/${folderId}/contacts`, {
          $select: 'id,displayName,emailAddresses,mobilePhone,jobTitle,companyName'
        });
        
        const contacts = response.value || [];
        
        const structured = {
          id: folder.id,
          name: folder.displayName,
          parentFolderId: folder.parentFolderId,
          contactCount: folder.contactCount || 0,
          contacts: contacts.map(c => ({
            id: c.id,
            displayName: c.displayName,
            emails: c.emailAddresses?.map(e => ({ address: e.address, name: e.name })) || [],
            mobilePhone: c.mobilePhone,
            jobTitle: c.jobTitle,
            companyName: c.companyName
          }))
        };
        
        const textFallback = [
          `Contact Folder: ${folder.displayName}`,
          `ID: ${folder.id}`,
          `Contacts: ${contacts.length}`,
          '',
          contacts.map(c => `- ${c.displayName} (${c.emailAddresses?.[0]?.address || 'no email'})`).join('\n')
        ].join('\n');
        
        return {
          contents: [{
            uri: uri.toString(),
            mimeType: "application/json",
            text: JSON.stringify(structured, null, 2)
          }]
        };
      } catch (error) {
        if (isAuthError(error)) throw error;
        throw new Error(`Error reading contact folder: ${error.message}`);
      }
    }
  );
}

module.exports = { registerContactsResources };