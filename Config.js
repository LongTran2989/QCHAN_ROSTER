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

  // Schedule source: public view-only SharePoint Excel link (replaced the old Google Sheet)
  EXTERNAL_SOURCE: {
    URL: "https://vaecovna-my.sharepoint.com/:x:/g/personal/nguyendinhkhanh_vaeco_com_vn/IQDQlCIeF2Y6R7vOHjuFSWo1ASGlooDnTJBXBv6y5Y75axI?e=FBI5ph",
    ENABLED: true, // false = fall back to the old in-spreadsheet SCHEDULE tab
    SHEET_NAME: "" // tab to read; empty = first tab
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

  // Color mapping by Aircraft Type
  COLORS: {
    A320: "yellow",
    A321: "yellow",
    A350: "#56a9cb",
    B787: "orange",
    DEFAULT: "#fca8a8",
    EA_LAN: "#00ffff", // L-HAN
    EA_HAN: "#00ff00", // HAN
    BG_SAT_SUN: "gray",
    BG_NULL: "black",
    WHITE: "white",
    NIGHT_SHIFT_FLAG: "#ff4d4d", // marks the start/end day of a WP needing night-shift coverage
    CHANGE_HIGHLIGHT: "orange" // border around a WP row that was added/changed this run
  }
};
