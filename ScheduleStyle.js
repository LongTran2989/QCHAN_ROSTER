/**
 * Pure color/style helpers for the A/C Schedule. No SpreadsheetApp dependency, so they can be
 * exercised outside Apps Script (see test/scheduleStyle.test.js).
 */

var STYLE_BLACK = "#000000";
var STYLE_WHITE = "#FFFFFF";

/**
 * WCAG 2.1 relative luminance of a 6-digit hex color.
 * @param {string} hex e.g. "#B71C1C"
 * @returns {number}
 */
function relativeLuminance(hex) {
  var channels = [1, 3, 5].map(function (i) {
    var c = parseInt(hex.substr(i, 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}

/**
 * WCAG contrast ratio between two hex colors (1..21).
 * @param {string} hexA
 * @param {string} hexB
 * @returns {number}
 */
function contrastRatio(hexA, hexB) {
  var a = relativeLuminance(hexA);
  var b = relativeLuminance(hexB);
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}

/**
 * Black unless white has strictly higher contrast against the fill.
 * @param {string} bgHex
 * @returns {string} "#000000" or "#FFFFFF"
 */
function pickFontColor(bgHex) {
  return contrastRatio(bgHex, STYLE_WHITE) > contrastRatio(bgHex, STYLE_BLACK) ? STYLE_WHITE : STYLE_BLACK;
}

// Allow `node test/scheduleStyle.test.js` to require() this file, while GAS still sees plain globals.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    relativeLuminance: relativeLuminance,
    contrastRatio: contrastRatio,
    pickFontColor: pickFontColor
  };
}
