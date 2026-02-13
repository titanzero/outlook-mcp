/**
 * Response formatter utility.
 *
 * Provides a TOON / plain-text toggle so every handler can
 * emit structured data and have it automatically encoded in
 * the format chosen by config.RESPONSE_FORMAT.
 *
 * - "toon"  → Token-Oriented Object Notation (compact, LLM-friendly)
 * - "text"  → human-readable plain text (legacy default)
 */
const { encode } = require('@toon-format/toon');
const config = require('../config');

/**
 * Whether the TOON response format is currently active.
 * Reads the env var directly so tests can override at runtime.
 * @returns {boolean}
 */
function isToonEnabled() {
  return (process.env.OUTLOOK_RESPONSE_FORMAT || config.RESPONSE_FORMAT) === 'toon';
}

/**
 * Encode a plain JS object / array as a TOON string.
 * @param {*} data - JSON-serialisable value
 * @returns {string}
 */
function formatAsToon(data) {
  return encode(data);
}

/**
 * Return whichever representation the current config demands.
 *
 * @param {*}      structuredData  - data shaped for TOON encoding
 * @param {string} textFallback    - pre-formatted plain-text string
 * @returns {string}
 */
function formatResponse(structuredData, textFallback) {
  if (isToonEnabled()) {
    return formatAsToon(structuredData);
  }
  return textFallback;
}

module.exports = {
  isToonEnabled,
  formatAsToon,
  formatResponse,
};
