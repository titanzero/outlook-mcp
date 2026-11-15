/**
 * Create contact with full fields functionality
 * Uses the Microsoft Graph JS SDK.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

function parseAddress(csv) {
  if (!csv) {
    return undefined;
  }
  return csv.split('|').map(part => {
    const [street, city, state, postalCode, country] = part.split(';').map(s => s.trim());
    return { street, city, state, postalCode, countryOrRegion: country };
  });
}

function parsePhones(csv) {
  if (!csv) {
    return undefined;
  }
  const phones = {};
  csv.split(',').forEach(p => {
    const [type, number] = p.split(':').map(s => s.trim());
    if (type && number) {
      phones[type] = number;
    }
  });
  return Object.keys(phones).length ? phones : undefined;
}

/**
 * Create contact with full fields handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateContactFull(args) {
  const {
    displayName,
    email,
    phone,
    jobTitle,
    companyName,
    // Additional fields
    givenName,
    surname,
    emailAddresses,
    businessPhones,
    homePhones,
    mobilePhone,
    businessAddress,
    homeAddress,
    otherAddress,
    birthday,
    spouseName,
    personalNotes,
    manager,
    assistantName,
    department,
    officeLocation,
    profession,
    titles,
    fileAs,
    initials,
    generation,
    imAddresses,
  } = args;

  if (!displayName) {
    return makeErrorResponse('displayName is required.');
  }

  try {
    const client = await getGraphClient();

    const contact = { displayName };

    // Basic fields
    if (givenName) {
      contact.givenName = givenName;
    }
    if (surname) {
      contact.surname = surname;
    }
    if (jobTitle) {
      contact.jobTitle = jobTitle;
    }
    if (companyName) {
      contact.companyName = companyName;
    }
    if (department) {
      contact.department = department;
    }
    if (officeLocation) {
      contact.officeLocation = officeLocation;
    }
    if (profession) {
      contact.profession = profession;
    }
    if (spouseName) {
      contact.spouseName = spouseName;
    }
    if (personalNotes) {
      contact.personalNotes = personalNotes;
    }
    if (manager) {
      contact.manager = manager;
    }
    if (assistantName) {
      contact.assistantName = assistantName;
    }
    if (fileAs) {
      contact.fileAs = fileAs;
    }
    if (initials) {
      contact.initials = initials;
    }
    if (generation) {
      contact.generation = generation;
    }

    // Email addresses - support multiple
    if (emailAddresses) {
      contact.emailAddresses = emailAddresses.split(',').map(e => {
        const [addr, name] = e.split('|').map(s => s.trim());
        return { address: addr, name: name || displayName };
      });
    } else if (email) {
      contact.emailAddresses = [{ address: email, name: displayName }];
    }

    // Phone numbers
    if (businessPhones) {
      contact.businessPhones = businessPhones.split(',').map(p => p.trim());
    }
    if (homePhones) {
      contact.homePhones = homePhones.split(',').map(p => p.trim());
    }
    if (mobilePhone) {
      contact.mobilePhone = mobilePhone;
    } else if (phone) {
      contact.mobilePhone = phone;
    } // Backward compat

    // Addresses
    if (businessAddress) {
      const addr = parseAddress(businessAddress);
      if (addr) {
        contact.businessAddress = addr[0];
      }
    }
    if (homeAddress) {
      const addr = parseAddress(homeAddress);
      if (addr) {
        contact.homeAddress = addr[0];
      }
    }
    if (otherAddress) {
      const addr = parseAddress(otherAddress);
      if (addr) {
        contact.otherAddress = addr[0];
      }
    }

    // Birthday (ISO 8601 date string)
    if (birthday) {
      contact.birthday = birthday;
    }

    // IM addresses
    if (imAddresses) {
      contact.imAddresses = imAddresses.split(',').map(a => a.trim());
    }

    // Titles array
    if (titles) {
      contact.titles = titles.split(',').map(t => t.trim());
    }

    const created = await client.api('me/contacts').post(contact);

    return makeResponse(
      `Contact created successfully.\nID: ${created.id}\nName: ${created.displayName}`
    );
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating contact: ${error.message}`);
  }
}

module.exports = handleCreateContactFull;
