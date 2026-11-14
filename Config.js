/**
 * Global Configuration for the SQD Roster Application.
 * Centralizes all constants, sheet IDs, layout bounds, and recipients.
 */
const CONFIG = {
  // Spreadsheet IDs and Names
  SHEET_IDS: {
    PUBLIC_ROSTER: "1cC8OJAlAp5TIcXMUj6Pot2zVA_I8voS7qcCI0F7vtZw",
    SCHEDULE: 2119712554,
    ROSTER_TEMPLATE: "ROSTER_TEMPLATE",
    EMAILS_SHEET: "Email_Config",
    CHANGE_LOG_SHEET: "Change Log",
    ASSIGNMENTS_DB: "Assignments_DB"
  },

  // Schedule source: SharePoint Excel link (anonymous share). Downloaded with cookie-aware
  // redirect handling (see ExternalSource.js). ENABLED = false falls back to the SCHEDULE tab.
  EXTERNAL_SOURCE: {
    URL: "https://vaecovna-my.sharepoint.com/:x:/g/personal/nguyendinhkhanh_vaeco_com_vn/IQDQlCIeF2Y6R7vOHjuFSWo1ASGlooDnTJBXBv6y5Y75axI?e=FBI5ph",
    ENABLED: true, // true = SharePoint raw AMOS export (tab 'RAW', converted by convertRawToSchedule); false = in-spreadsheet SCHEDULE tab
    SHEET_NAME: "RAW" // tab to read; empty = first tab
  },

  SNAPSHOT_SHEET_PREFIX: "_Snapshot_",

  // Layout constraints for the Roster table
  ROSTER: {
    UPPER_ROW: 2,
    LOWER_ROW: 73,
    LEFT_COL: 3,
    RIGHT_COL: 33,
    UPDATE_INFO_CELL: "B75"
  },

  // Personnel info indexes assuming column mappings (legacy, from live Apps Script project)
  PERSONNEL: {
    TT: 0,
    VAECO_ID: 1,
    NAME: 2,
    TITLE: 3,
    PJID: 37
  },

  // Source of names for the Assign Personnel sidebar
  PERSONNEL_SHEET: {
    NAME: "Personel info",
    NAME_COL: 2, // column B
    START_ROW: 2
  },

  // Color mapping. All hex so ScheduleStyle.js can compute contrast (WCAG >= 4.5:1, enforced by
  // test/scheduleStyle.test.js). Pastel fills take black text; NIGHT_SHIFT_FLAG takes white text.
  COLORS: {
    A320: "#FFF59D",
    A321: "#FFF59D",
    A350: "#81D4FA",
    B787: "#FFCC80",
    DEFAULT: "#CE93D8",
    EA_LAN: "#80DEEA", // L-HAN
    EA_HAN: "#A5D6A7", // HAN
    BG_SAT_SUN: "gray",
    BG_NULL: "black",
    WHITE: "white",
    NIGHT_SHIFT_FLAG: "#B71C1C", // start/end day of a WP needing night-shift coverage (white text)
    CHANGE_HIGHLIGHT: "#6A1B9A", // border around a WP row that was added/changed this run
    UNASSIGNED_BORDER: "#424242" // dashed border around a WP row with no assignee
  },

  // Fixed row height (px) for schedule rows so long labels never make short checks tall
  LAYOUT: {
    ROW_HEIGHT: 21
  }
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = { CONFIG: CONFIG };
}
