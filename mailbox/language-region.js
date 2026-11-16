/**
 * Language and region settings tools for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Get language and region settings
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetLanguageRegion(_args) {
  try {
    const client = await getGraphClient();
    
    const settings = await client.api('me/mailboxSettings').select('language,dateFormat,timeFormat').get();
    
    const structured = {
      language: settings.language?.displayName || null,
      languageLocale: settings.language?.locale || null,
      dateFormat: settings.dateFormat || null,
      timeFormat: settings.timeFormat || null,
    };
    
    const textFallback = [
      `Language & Region Settings`,
      `Language: ${structured.language || 'Not set'} (${structured.languageLocale || 'Not set'})`,
      `Date Format: ${structured.dateFormat || 'Not set'}`,
      `Time Format: ${structured.timeFormat || 'Not set'}`,
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting language/region: ${error.message}`);
  }
}

/**
 * Set language
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSetLanguage(args) {
  const { locale } = args;
  
  if (!locale) {
    return makeErrorResponse('Locale is required (e.g., "en-US", "fr-FR", "de-DE")');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api('me/mailboxSettings').patch({
      language: { locale, displayName: locale },
    });
    
    return makeResponse(`Language set to ${locale} successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error setting language: ${error.message}`);
  }
}

/**
 * Set date format
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSetDateFormat(args) {
  const { format } = args;
  
  if (!format) {
    return makeErrorResponse('Date format is required (e.g., "MM/dd/yyyy", "dd/MM/yyyy", "yyyy-MM-dd")');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api('me/mailboxSettings').patch({ dateFormat: format });
    
    return makeResponse(`Date format set to "${format}" successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error setting date format: ${error.message}`);
  }
}

/**
 * Set time format
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSetTimeFormat(args) {
  const { format } = args;
  
  if (!format) {
    return makeErrorResponse('Time format is required (e.g., "H:mm", "h:mm tt", "HH:mm")');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api('me/mailboxSettings').patch({ timeFormat: format });
    
    return makeResponse(`Time format set to "${format}" successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error setting time format: ${error.message}`);
  }
}

/**
 * Get available locales (common ones)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListLocales(_args) {
  const commonLocales = [
    { locale: 'en-US', name: 'English (United States)', region: 'US' },
    { locale: 'en-GB', name: 'English (United Kingdom)', region: 'GB' },
    { locale: 'en-CA', name: 'English (Canada)', region: 'CA' },
    { locale: 'en-AU', name: 'English (Australia)', region: 'AU' },
    { locale: 'fr-FR', name: 'French (France)', region: 'FR' },
    { locale: 'fr-CA', name: 'French (Canada)', region: 'CA' },
    { locale: 'de-DE', name: 'German (Germany)', region: 'DE' },
    { locale: 'es-ES', name: 'Spanish (Spain)', region: 'ES' },
    { locale: 'es-MX', name: 'Spanish (Mexico)', region: 'MX' },
    { locale: 'it-IT', name: 'Italian (Italy)', region: 'IT' },
    { locale: 'pt-BR', name: 'Portuguese (Brazil)', region: 'BR' },
    { locale: 'pt-PT', name: 'Portuguese (Portugal)', region: 'PT' },
    { locale: 'ja-JP', name: 'Japanese (Japan)', region: 'JP' },
    { locale: 'ko-KR', name: 'Korean (South Korea)', region: 'KR' },
    { locale: 'zh-CN', name: 'Chinese (Simplified, China)', region: 'CN' },
    { locale: 'zh-TW', name: 'Chinese (Traditional, Taiwan)', region: 'TW' },
    { locale: 'ru-RU', name: 'Russian (Russia)', region: 'RU' },
    { locale: 'nl-NL', name: 'Dutch (Netherlands)', region: 'NL' },
    { locale: 'sv-SE', name: 'Swedish (Sweden)', region: 'SE' },
    { locale: 'da-DK', name: 'Danish (Denmark)', region: 'DK' },
    { locale: 'fi-FI', name: 'Finnish (Finland)', region: 'FI' },
    { locale: 'no-NO', name: 'Norwegian (Norway)', region: 'NO' },
    { locale: 'pl-PL', name: 'Polish (Poland)', region: 'PL' },
    { locale: 'cs-CZ', name: 'Czech (Czech Republic)', region: 'CZ' },
    { locale: 'hu-HU', name: 'Hungarian (Hungary)', region: 'HU' },
    { locale: 'tr-TR', name: 'Turkish (Turkey)', region: 'TR' },
    { locale: 'he-IL', name: 'Hebrew (Israel)', region: 'IL' },
    { locale: 'ar-SA', name: 'Arabic (Saudi Arabia)', region: 'SA' },
  ];
  
  const structured = { locales: commonLocales };
  
  const textFallback = [
    `Common Locales (${commonLocales.length}):`,
    '',
    commonLocales.map(l => `${l.locale}: ${l.name}`).join('\n'),
    '',
    'Use set-language with locale code (e.g., "en-US").',
  ].join('\n');
  
  return makeResponse(formatResponse(structured, textFallback));
}

/**
 * List common date formats
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListDateFormats(_args) {
  const formats = [
    { format: 'MM/dd/yyyy', example: '12/31/2024', name: 'US (Month/Day/Year)' },
    { format: 'dd/MM/yyyy', example: '31/12/2024', name: 'UK/EU (Day/Month/Year)' },
    { format: 'yyyy-MM-dd', example: '2024-12-31', name: 'ISO (Year-Month-Day)' },
    { format: 'dd.MM.yyyy', example: '31.12.2024', name: 'German/European' },
    { format: 'yyyy/MM/dd', example: '2024/12/31', name: 'Japanese/Chinese' },
    { format: 'MMM d, yyyy', example: 'Dec 31, 2024', name: 'US Long' },
    { format: 'd MMM yyyy', example: '31 Dec 2024', name: 'UK Long' },
  ];
  
  const structured = { formats };
  
  const textFallback = [
    `Common Date Formats:`,
    '',
    formats.map(f => `${f.format} (${f.name}): ${f.example}`).join('\n'),
    '',
    'Use set-date-format with format string.',
  ].join('\n');
  
  return makeResponse(formatResponse(structured, textFallback));
}

/**
 * List common time formats
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListTimeFormats(_args) {
  const formats = [
    { format: 'H:mm', example: '14:30', name: '24-hour (HH:MM)' },
    { format: 'HH:mm', example: '14:30', name: '24-hour zero-padded' },
    { format: 'h:mm tt', example: '2:30 PM', name: '12-hour with AM/PM' },
    { format: 'hh:mm tt', example: '02:30 PM', name: '12-hour zero-padded' },
    { format: 'H:mm:ss', example: '14:30:00', name: '24-hour with seconds' },
    { format: 'h:mm:ss tt', example: '2:30:00 PM', name: '12-hour with seconds' },
  ];
  
  const structured = { formats };
  
  const textFallback = [
    `Common Time Formats:`,
    '',
    formats.map(f => `${f.format} (${f.name}): ${f.example}`).join('\n'),
    '',
    'Use set-time-format with format string.',
  ].join('\n');
  
  return makeResponse(formatResponse(structured, textFallback));
}

module.exports = {
  handleGetLanguageRegion,
  handleSetLanguage,
  handleSetDateFormat,
  handleSetTimeFormat,
  handleListLocales,
  handleListDateFormats,
  handleListTimeFormats,
};