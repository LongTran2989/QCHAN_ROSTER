/**
 * Pure layout math for the A/C Schedule block (no SpreadsheetApp dependency, so it can be
 * exercised outside Apps Script; see test/scheduleLayout.test.js).
 *
 * Section order: PHASE CHECKS -> NORMAL CHECKS -> OTHER SHORT CHECKS -> END OF LIST -> STO.
 * One row cursor drives all positions, so no block can land inside another.
 */

// Minimum number of rows reserved for the Phase block (matches the historical layout).
var PHASE_MIN_ROWS = 9;

/**
 * Stacks Phase WPs that start on the same day column: the i-th entry gets slot k when k earlier
 * entries share its column. Phase WP with slot k is drawn at phaseStartRow + 2*k (label) with its
 * reg row directly below.
 * @param {number[]} cols day column of each Phase WP, in render order
 * @returns {{slots: number[], depth: number}} depth = number of stack levels used (0 when empty)
 */
function assignPhaseSlots(cols) {
  var countByCol = {};
  var slots = [];
  var maxSlot = -1;
  for (var i = 0; i < cols.length; i++) {
    var used = countByCol[cols[i]] || 0;
    slots.push(used);
    countByCol[cols[i]] = used + 1;
    if (used > maxSlot) maxSlot = used;
  }
  return { slots: slots, depth: maxSlot + 1 };
}

/**
 * Computes every section's row positions from the section sizes.
 * @param {{phaseHeaderRow: number, phaseDepth: number, normalCount: number, otherCount: number, stoCount: number}} opts
 * @returns {{phaseHeaderRow: number, phaseStartRow: number, phaseRowCount: number,
 *            normalHeaderRow: number, normalStartRow: number,
 *            otherHeaderRow: (number|null), otherStartRow: (number|null),
 *            endOfListRow: number, stoStartRow: number}}
 */
function computeScheduleLayout(opts) {
  var phaseStartRow = opts.phaseHeaderRow + 1;
  var phaseRowCount = Math.max(PHASE_MIN_ROWS, 2 * opts.phaseDepth + 1);
  var normalHeaderRow = phaseStartRow + phaseRowCount;
  var normalStartRow = normalHeaderRow + 1;

  var otherHeaderRow = null;
  var otherStartRow = null;
  var endOfListRow;
  if (opts.otherCount > 0) {
    otherHeaderRow = normalStartRow + opts.normalCount + 1;
    otherStartRow = otherHeaderRow + 1;
    endOfListRow = otherStartRow + opts.otherCount + 2;
  } else {
    endOfListRow = normalStartRow + opts.normalCount + 2;
  }

  return {
    phaseHeaderRow: opts.phaseHeaderRow,
    phaseStartRow: phaseStartRow,
    phaseRowCount: phaseRowCount,
    normalHeaderRow: normalHeaderRow,
    normalStartRow: normalStartRow,
    otherHeaderRow: otherHeaderRow,
    otherStartRow: otherStartRow,
    endOfListRow: endOfListRow,
    stoStartRow: endOfListRow + 1
  };
}

// Allow `node test/scheduleLayout.test.js` to require() this file, while GAS still sees plain globals.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    PHASE_MIN_ROWS: PHASE_MIN_ROWS,
    assignPhaseSlots: assignPhaseSlots,
    computeScheduleLayout: computeScheduleLayout
  };
}
