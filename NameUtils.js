/**
 * Compacts a full Vietnamese name (family + middle + given) into initials-plus-given-name
 * for on-grid display, e.g. "Tran Thanh Long" -> "TTLong". Pure, no SpreadsheetApp
 * dependency, so it's exercised outside Apps Script (see test/nameUtils.test.js).
 */

/**
 * @param {string} fullName - e.g. "Trần Thanh Long"
 * @returns {string} e.g. "TTLong". A single-word name is returned unchanged.
 */
function toInitialsWithFirstName(fullName) {
  var words = (fullName || "").trim().split(/\s+/).filter(function (w) { return w.length > 0; });
  if (words.length === 0) return "";
  if (words.length === 1) return words[0];

  var givenName = words[words.length - 1];
  var initials = words.slice(0, -1).map(function (w) {
    return w.charAt(0).toUpperCase();
  }).join("");

  return initials + givenName;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    toInitialsWithFirstName: toInitialsWithFirstName
  };
}
