/**
 * Jest CJS mock for @toon-format/toon (ESM-only package).
 * Provides a minimal encode() that serialises data as JSON
 * so tests can exercise the formatter path without ESM issues.
 */
module.exports = {
  encode: function encode(data) {
    return JSON.stringify(data);
  },
};
