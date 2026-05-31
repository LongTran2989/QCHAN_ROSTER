# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script (GAS) project that manages aircraft crew rosters and schedules for SQD. The application runs entirely within Google Sheets, providing UI menus to automate roster generation, schedule updates, public roster syncing, and timesheet exports.

## Key Technologies

- **Google Apps Script** (V8 runtime) - The primary platform
- **Google Sheets API** - For spreadsheet manipulation
- **Google MailApp** - For email notifications
- **Clasp** (Google Apps Script CLI v3.3.0) - For deployment and version control

## Project Setup

### Dependencies
```bash
npm install
```

The only npm dependency is `@google/clasp` (v3.3.0), used for pushing code to Google Apps Script.

### Deployment
```bash
npx clasp push
```

This syncs all `.js` files from the local directory to the Google Apps Script project (ID: `1JwHDL-OLbrBIP807d6pX9zpi9fF3qQYEoqxcYGyrffKjIdFOjq5eUWN1`).

**Note:** The `.clasp.json` file contains the script ID and configuration. The `rootDir` is empty, meaning all files at the repository root are pushed.

## Architecture Overview

### Core Modules

The codebase follows a modular architecture with clear separation of concerns:

1. **Main.js** - Entry point and UI trigger
   - `onOpen()`: Builds the "QC HAN" custom menu when spreadsheet opens
   - `publishAsNewRevision()`: Increments revision number, renames sheet, sends email notifications
   - `showDialog()`: Displays link to public roster

2. **Config.js** - Global configuration singleton
   - Spreadsheet IDs (public roster, schedule data source, template)
   - Roster layout bounds (row/column indices)
   - Color mappings by aircraft type (A320, A321, A350, B787)
   - Personnel column indices
   - Timezone: Asia/Bangkok

3. **ScheduleManager.js** - Aircraft schedule synchronization (largest/most complex module)
   - `updateACSchedules()`: Main function triggered by UI menu
   - Fetches aircraft schedule data from external schedule sheet
   - Filters schedules by month and station (HAN/L-HAN only)
   - Sorts by aircraft type and departure date
   - Splits into three categories: Phase Checks (EA), Normal Checks, and STO (Storage) Checks
   - Renders schedules to roster grid with color-coding and formatting
   - Uses in-memory batch processing for performance
   - Contains `SCHEDULE_INDEX` constants mapping array positions

4. **RosterManager.js** - Roster sheet creation and formatting
   - `createNewRoster()`: UI menu handler prompting for month/year
   - `fillDateToSheet()`: Duplicates template sheet, generates date/day headers for the month
   - `setBackgroundColor()`: Colors weekends (gray) and empty cells (black) in batches
   - `updatePublicRoster()`: Copies active roster to public spreadsheet with security filtering
     - Applies date window filter (-3 to +3 days from current date)
     - Only displays BA1-C and BA2 assignments in the public view
     - Clears data validations and extraneous content

5. **TimesheetManager.js** - Timesheet export (Vietnamese: "Bảng Chấm Công")
   - `ccExport()`: Exports roster data to timesheet format
   - `createFormCC()`: Creates formatted timesheet sheet with dynamic column sizing
   - Includes day-of-week and date rows
   - Colors weekends gray
   - Validates sheet selection (rejects guide and personnel sheets)

6. **Utils.js** - Common utility functions
   - Date/array manipulation: `arraymove()`, `daysInMonth()`, `textDay()`, `textMonth()`
   - Grid creation: `createGrid(rows, cols, defaultVal)`
   - Safe spreadsheet access: `openSpreadsheetSafe(id)`
   - Email management: `getEmailRecipients()` - reads from "Email_Config" sheet
   - Metadata tracking: `recordUpdateMetadata(sheet)` - logs update timestamp and user

7. **TEST.js** - Development test functions (not part of production flow)
   - Various debugging and exploration functions
   - Examples: `test()`, `getFileName()`, `test1()`, `test2()`, `test3()`

8. **macros.js** - Legacy Google Sheets macro
   - `sortHAN()` - Old macro for sorting schedules (superseded by in-memory sorting in ScheduleManager)

## Build and Test Commands

### Testing
```bash
npm test
```

Currently returns an error (no test framework configured). The codebase has a `TEST.js` file with manual test functions that can be run in Google Apps Script's editor.

### Linting

No built-in linting is configured. For GAS projects, consider using:
- ESLint with Google Apps Script dialect
- Manual code review

### Code Quality Notes

- **In-Memory Batch Processing**: ScheduleManager and TimesheetManager use in-memory grids (`createGrid()`) to build entire data structures before batching API calls. This significantly improves performance by reducing network round-trips.
- **Error Handling**: All major functions wrap operations in try-catch blocks with `console.error()` and user-facing `SpreadsheetApp.getUi().alert()` notifications.
- **Configuration Centralization**: All constants (sheet IDs, layout bounds, colors) are defined in Config.js to simplify maintenance.

## Spreadsheet Configuration

### Required Sheets

The application expects these sheets to exist in the deployed Google Sheets:

1. **ROSTER_TEMPLATE** - Template sheet duplicated to create monthly rosters
2. **AC CHECKS** (external sheet) - Schedule data source (accessed via sheet ID 2119712554)
3. **Email_Config** - Hidden sheet containing email recipients for publish notifications (created automatically if missing)
4. **CC_TEMP** - Hidden template for timesheet generation
5. **Public Roster** (external spreadsheet) - Synced public view (ID in CONFIG)

### Typical Workflow

1. User opens the main roster spreadsheet
2. "QC HAN" menu appears with options:
   - "Create new Roster" → prompts for month/year → duplicates template → fills dates
   - "Update A/C Schedules" → fetches and renders aircraft schedules
   - "Update Public Roster" → copies filtered view to public spreadsheet
   - "Xuất bảng chấm công" → exports to timesheet format
   - "Publish as new revision" → increments version, renames sheet, sends email

## Important Behavioral Details

### Schedule Filtering
- Only includes flights departing from HAN or L-HAN stations
- Filters to schedules within the current month (with cross-month boundary handling)
- Splits Phase Checks (TAT 0.5-1 day) from normal checks for separate rendering

### Color Coding
- A320/A321 → Yellow
- A350 → Light blue (#56a9cb)
- B787 → Orange
- Default → Pink (#fca8a8)
- L-HAN Phase Checks → Cyan (#00ffff)
- HAN Phase Checks → Green (#00ff00)
- Weekends → Gray
- Empty cells → Black

### Public Roster Security
- Displays schedules only within ±3 days of current date
- Only shows BA1-C and BA2 assignments
- Removes data validations
- Clears extraneous rows/columns

### Assignment Persistence
- User assignments in Column B are backed up before redrawing schedules
- Previous assignments are restored to preserve manual edits during schedule updates

## Recent Changes (Git History)

- **98389dc**: Fixed String casted to Date - Date handling improvements
- **4f357e2**: Phase 1: Implement In-Memory Processing - Major performance optimization
- **409eae8**: Move update info to B75 - Changed update metadata cell location
- **c27de92**: Added update tracking to cell B75 - Added timestamp and user logging
- **1acb123**: Initial refactor: modular architecture and dynamic emails - Refactored from monolithic to modular design

## Notes for Future Development

- **No automated tests**: The TEST.js file contains manual test functions that must be executed in the GAS editor.
- **Google Apps Script limitations**: Bound scripts have a 6-minute execution limit; complex operations are optimized with in-memory batch processing.
- **Clasp push only**: There is no automated CI/CD pipeline; code is manually pushed via `npx clasp push`.
- **Time zone**: All date/time operations use Asia/Bangkok timezone as configured in appsscript.json.
