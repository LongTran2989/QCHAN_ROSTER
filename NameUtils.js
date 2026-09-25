/**
 * Compacts a full Vietnamese name (family + middle + given) into initials-plus-given-name
 * for on-grid display, e.g. "Tran Thanh Long" -> "TTLong". Pure, no SpreadsheetApp
 * dependency, so it's exercised outside Apps Script (see test/nameUtils.test.js).
 */

var ASSIGNED_PEOPLE_SEPARATOR = ", ";

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

/**
 * Splits the snapshot's stored "AssignedPerson" cell -- one or more full names joined by
 * ASSIGNED_PEOPLE_SEPARATOR -- back into an array. Tolerant of stray whitespace/empties so a
 * hand-edited cell or a single legacy name doesn't break.
 * @param {string} value
 * @returns {string[]}
 */
function splitAssignedPeople(value) {
  return (value || "")
    .split(",")
    .map(function (name) { return name.trim(); })
    .filter(function (name) { return name !== ""; });
}

/**
 * Inverse of splitAssignedPeople(): the canonical stored form for a set of assignees.
 * @param {string[]} names
 * @returns {string}
 */
function joinAssignedPeople(names) {
  return (names || [])
    .map(function (name) { return (name || "").trim(); })
    .filter(function (name) { return name !== ""; })
    .join(ASSIGNED_PEOPLE_SEPARATOR);
}

/**
 * Compacts a list of full names into their on-grid short form, e.g.
 * ["Tran Thanh Long", "Vu Hong Hai"] -> "TTLong, VHHải".
 * @param {string[]} names
 * @returns {string}
 */
function formatAssignedShortNames(names) {
  return (names || []).map(toInitialsWithFirstName).join(ASSIGNED_PEOPLE_SEPARATOR);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    toInitialsWithFirstName: toInitialsWithFirstName,
    splitAssignedPeople: splitAssignedPeople,
    joinAssignedPeople: joinAssignedPeople,
    formatAssignedShortNames: formatAssignedShortNames
  };
}
