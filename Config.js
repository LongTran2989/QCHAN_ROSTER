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
    EMAILS_SHEET: "Email_Config"
  },
  
  // Layout constraints for the Roster table
  ROSTER: {
    UPPER_ROW: 2,
    LOWER_ROW: 73,
    LEFT_COL: 3,
    RIGHT_COL: 33,
    UPDATE_INFO_CELL: "B75"
  },

  // Personnel info indexes assuming column mappings
  PERSONNEL: {
    TT: 0,
    VAECO_ID: 1,
    NAME: 2,
    TITLE: 3,
    PJID: 37 // Using fixed index for compatibility, though we recommend dynamic header search
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
    WHITE: "white"
  }
};
