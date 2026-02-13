/**
 * Create rule functionality
 * Uses the Microsoft Graph JS SDK.
 */
const config = require('../config');
const { getGraphClient } = require('../utils/graph-client');
const { getFolderIdByName } = require('../email/folder-utils');
const { getInboxRules } = require('./list');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Create rule handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateRule(args) {
  const {
    name,
    fromAddresses,
    containsSubject,
    hasAttachments,
    moveToFolder,
    markAsRead,
    isEnabled = true,
    sequence
  } = args;
  
  // Add validation for sequence parameter
  if (sequence !== undefined && (isNaN(sequence) || sequence < 1)) {
    return makeErrorResponse('Sequence must be a positive number greater than zero.');
  }
  
  if (!name) {
    return makeErrorResponse('Rule name is required.');
  }
  
  // Validate that at least one condition or action is specified
  const hasCondition = fromAddresses || containsSubject || hasAttachments === true;
  const hasAction = moveToFolder || markAsRead === true;
  
  if (!hasCondition) {
    return makeErrorResponse('At least one condition is required. Specify fromAddresses, containsSubject, or hasAttachments.');
  }
  
  if (!hasAction) {
    return makeErrorResponse('At least one action is required. Specify moveToFolder or markAsRead.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Create rule
    const result = await createInboxRule(client, {
      name,
      fromAddresses,
      containsSubject,
      hasAttachments,
      moveToFolder,
      markAsRead,
      isEnabled,
      sequence
    });
    
    let responseText = result.message;
    
    // Add a tip about sequence if it wasn't provided
    if (sequence === undefined && !result.error) {
      responseText += "\n\nTip: You can specify a 'sequence' parameter when creating rules to control their execution order. Lower sequence numbers run first.";
    }
    
    return makeResponse(responseText);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error creating rule: ${error.message}`);
  }
}

/**
 * Create a new inbox rule
 * @param {object} client - Microsoft Graph SDK client
 * @param {object} ruleOptions - Rule creation options
 * @returns {Promise<object>} - Result object with status and message
 */
async function createInboxRule(client, ruleOptions) {
  try {
    const {
      name,
      fromAddresses,
      containsSubject,
      hasAttachments,
      moveToFolder,
      markAsRead,
      isEnabled,
      sequence
    } = ruleOptions;
    
    // Get existing rules to determine sequence if not provided
    let ruleSequence = sequence;
    if (!ruleSequence) {
      try {
        // Default to 100 if we can't get existing rules
        ruleSequence = config.DEFAULT_RULE_SEQUENCE;
        
        // Get existing rules to find highest sequence
        const existingRules = await getInboxRules(client);
        if (existingRules && existingRules.length > 0) {
          // Find the highest sequence
          const highestSequence = Math.max(...existingRules.map(r => r.sequence || 0));
          // Set new rule sequence to be higher
          ruleSequence = Math.max(highestSequence + 1, config.DEFAULT_RULE_SEQUENCE);
          console.error(`Auto-generated sequence: ${ruleSequence} (based on highest existing: ${highestSequence})`);
        }
      } catch (sequenceError) {
        console.error(`Error determining rule sequence: ${sequenceError.message}`);
        // Fall back to default value
        ruleSequence = config.DEFAULT_RULE_SEQUENCE;
      }
    }
    
    console.error(`Using rule sequence: ${ruleSequence}`);
    
    // Make sure sequence is a positive integer
    ruleSequence = Math.max(1, Math.floor(ruleSequence));
    
    // Build rule object with sequence
    const rule = {
      displayName: name,
      isEnabled: isEnabled === true,
      sequence: ruleSequence,
      conditions: {},
      actions: {}
    };
    
    // Add conditions
    if (fromAddresses) {
      // Parse email addresses
      const emailAddresses = fromAddresses.split(',')
        .map(email => email.trim())
        .filter(email => email)
        .map(email => ({
          emailAddress: {
            address: email
          }
        }));
      
      if (emailAddresses.length > 0) {
        rule.conditions.fromAddresses = emailAddresses;
      }
    }
    
    if (containsSubject) {
      rule.conditions.subjectContains = [containsSubject];
    }
    
    if (hasAttachments === true) {
      rule.conditions.hasAttachment = true;
    }
    
    // Add actions
    if (moveToFolder) {
      // Get folder ID
      try {
        const folderId = await getFolderIdByName(moveToFolder);
        if (!folderId) {
          return {
            success: false,
            message: `Target folder "${moveToFolder}" not found. Please specify a valid folder name.`
          };
        }
        
        rule.actions.moveToFolder = folderId;
      } catch (folderError) {
        console.error(`Error resolving folder "${moveToFolder}": ${folderError.message}`);
        return {
          success: false,
          message: `Error resolving folder "${moveToFolder}": ${folderError.message}`
        };
      }
    }
    
    if (markAsRead === true) {
      rule.actions.markAsRead = true;
    }
    
    // Create the rule
    const response = await client.api('me/mailFolders/inbox/messageRules').post(rule);
    
    if (response && response.id) {
      return {
        success: true,
        message: `Successfully created rule "${name}" with sequence ${ruleSequence}.`,
        ruleId: response.id
      };
    } else {
      return {
        success: false,
        message: "Failed to create rule. The server didn't return a rule ID."
      };
    }
  } catch (error) {
    console.error(`Error creating rule: ${error.message}`);
    throw error;
  }
}

module.exports = handleCreateRule;
