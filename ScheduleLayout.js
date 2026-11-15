/**
 * Pure layout math for the A/C Schedule block (no SpreadsheetApp dependency, so it can be
 * exercised outside Apps Script; see test/scheduleLayout.test.js).
 *
 * Section order: PHASE CHECKS -> NORMAL CHECKS -> OTHER CHECKS -> END OF LIST -> STO.
 * One row cursor drives all positions, so no block can land inside another.
 */

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
 *            endOfListRow: number, stoStartRow: number, lastRow: number}}
 */
function computeScheduleLayout(opts) {
  var phaseStartRow = opts.phaseHeaderRow + 1;
  // Exactly the rows the stacked Phase entries use (label + reg per level), so NORMAL CHECKS
  // starts on the row right after the last Phase line. 0 when there are no Phase WPs.
  var phaseRowCount = 2 * opts.phaseDepth;
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
    stoStartRow: endOfListRow + 1,
    // Last row any section writes: callers use it to size the cleared area / check the scan limit.
    lastRow: Math.max(endOfListRow, endOfListRow + opts.stoCount)
  };
}

/**
 * Row order for a check block: one continuous staircase. Start day, then end day
 * (earlier-ending first), then reg. Rows are schedule arrays whose FROM/TO (idx 3/4)
 * are already day-of-month numbers; reg is idx 1.
 */
function compareScheduleRows(a, b) {
  if (a[3] !== b[3]) return a[3] - b[3];
  if (a[4] !== b[4]) return a[4] - b[4];
  var ra = String(a[1]);
  var rb = String(b[1]);
  return ra < rb ? -1 : (ra > rb ? 1 : 0);
}

/** A/C type ascending (A321 before A350), then the staircase order within each type. */
function compareByTypeThenStair(a, b) {
  var ta = String(a[0]);
  var tb = String(b[0]);
  if (ta !== tb) return ta < tb ? -1 : 1;
  return compareScheduleRows(a, b);
}

// Allow `node test/scheduleLayout.test.js` to require() this file, while GAS still sees plain globals.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    compareScheduleRows: compareScheduleRows,
    compareByTypeThenStair: compareByTypeThenStair,
    assignPhaseSlots: assignPhaseSlots,
    computeScheduleLayout: computeScheduleLayout
  };
}
