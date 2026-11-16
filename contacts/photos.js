/**
 * Contact photos tools for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Get contact photo
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetContactPhoto(args) {
  const { contactId } = args;
  
  if (!contactId) {
    return makeErrorResponse('Contact ID is required (from list-contacts).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get photo metadata first
    let photoMetadata;
    try {
      photoMetadata = await client.api(`me/contacts/${contactId}/photo`).get();
    } catch (e) {
      return makeErrorResponse('Contact has no photo.');
    }
    
    // Get photo content
    const photoResponse = await client.api(`me/contacts/${contactId}/photo/$value`).get();
    
    // Convert to base64
    const base64Photo = Buffer.from(photoResponse).toString('base64');
    const mimeType = photoMetadata['@odata.mediaContentType'] || 'image/jpeg';
    
    const structured = {
      contactId,
      photo: {
        id: photoMetadata.id,
        width: photoMetadata.width,
        height: photoMetadata.height,
        mimeType,
        size: photoResponse.length,
        base64: base64Photo,
      },
    };
    
    const textFallback = [
      `Contact Photo: ${contactId}`,
      `Dimensions: ${photoMetadata.width}x${photoMetadata.height}`,
      `Type: ${mimeType}`,
      `Size: ${photoResponse.length} bytes`,
      '',
      `Base64 data: ${base64Photo.substring(0, 100)}... (truncated)`,
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Contact not found or has no photo.');
    }
    return makeErrorResponse(`Error getting contact photo: ${error.message}`);
  }
}

/**
 * Set/Update contact photo
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSetContactPhoto(args) {
  const { contactId, imageBase64, contentType = 'image/jpeg' } = args;
  
  if (!contactId) {
    return makeErrorResponse('Contact ID is required.');
  }
  if (!imageBase64) {
    return makeErrorResponse('Base64-encoded image data is required.');
  }
  
  // Validate base64
  let imageBuffer;
  try {
    imageBuffer = Buffer.from(imageBase64, 'base64');
  } catch (e) {
    return makeErrorResponse('Invalid base64 image data.');
  }
  
  // Check size (max 4MB for contact photos)
  if (imageBuffer.length > 4 * 1024 * 1024) {
    return makeErrorResponse('Image too large. Maximum size is 4MB.');
  }
  
  // Validate content type
  const validTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/bmp'];
  if (!validTypes.includes(contentType)) {
    return makeErrorResponse(`Invalid content type. Must be one of: ${validTypes.join(', ')}`);
  }
  
  try {
    const client = await getGraphClient();
    
    // Upload photo
    await client.api(`me/contacts/${contactId}/photo/$value`)
      .header('Content-Type', contentType)
      .put(imageBuffer);
    
    return makeResponse(`Contact photo updated successfully (${imageBuffer.length} bytes, ${contentType}).`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Contact not found.');
    }
    return makeErrorResponse(`Error setting contact photo: ${error.message}`);
  }
}

/**
 * Delete contact photo
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteContactPhoto(args) {
  const { contactId } = args;
  
  if (!contactId) {
    return makeErrorResponse('Contact ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/contacts/${contactId}/photo`).delete();
    
    return makeResponse('Contact photo deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Contact not found or has no photo.');
    }
    return makeErrorResponse(`Error deleting contact photo: ${error.message}`);
  }
}

module.exports = {
  handleGetContactPhoto,
  handleSetContactPhoto,
  handleDeleteContactPhoto,
};