# A/C Schedule Colors, Labels, Unassigned Flag & Legend — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make every A/C Schedule cell legible (WCAG-checked colors, font color derived from fill), shorten/clarify check labels without tall rows, flag unassigned WPs, and add an on-sheet legend.

**Architecture:** A new pure, node-testable `ScheduleStyle.js` holds contrast math, font-color selection, per-cell style resolution and legend entries. `Config.js` gets a hex-only, contrast-validated palette plus layout constants. `ScheduleManager.js` consumes these when building render payloads; label text comes from `ShiftUtils.js`. Label cells stop wrapping (overflow/clip) and row heights are fixed, so short-TAT checks never grow tall.

**Tech Stack:** Google Apps Script (V8, shared global scope), plain-Node tests (`node test/*.test.js`), clasp.

**Spec:** None written; implements items 1–4 from the conversation proposal (no red text; font color from fill; readable night/CHK markers; unassigned flag; legend; new change-highlight color), plus user notes: long labels must not make short-TAT rows tall; use standardized colors with verified contrast.

## Global Constraints

- All palette colors are 6-digit hex (no named colors) so `ScheduleStyle.js` can compute luminance.
- WCAG 2.1: every (fill, font) pair used on the schedule has contrast ratio **≥ 4.5:1** (tests enforce it); pastel fills with black text target **≥ 7:1**.
- Red is never a font color. Red (`NIGHT_SHIFT_FLAG`) means only "night-shift coverage required".
- Normal/STO/Phase rows use fixed height `CONFIG.LAYOUT.ROW_HEIGHT = 21` px; no `autoResizeRows`, no wrap on label cells.
- New `.js` files must work both as Apps Script globals and `require()`able (same guarded `module.exports` pattern as `ShiftUtils.js`).
- Keep `npm test` green; add new test files to the `test` script in `package.json`.

## Review Focus

- CHK check that also needs night shift (the original red-on-red): font must be white on `#B71C1C`, bold, `★` prefix.
- 1-day check with a long WP description + assignee: stays one line at `ROW_HEIGHT`, text overflows right, never wraps.
- Unknown `AC_TYPE` falls back to the default fill with passing contrast.
- WP both unassigned **and** changed this run: solid thick purple highlight must win over the dashed unassigned border.
- `shiftInfo` missing (null/undefined): no suffix, no crash, style still resolves.
- Month with zero checks: legend still draws, no error.

---

## File Structure

- Create `ScheduleStyle.js` — pure: `relativeLuminance`, `contrastRatio`, `pickFontColor`, `resolveCheckStyle`, `buildLegendEntries`.
- Create `test/scheduleStyle.test.js` — palette contrast + style resolution + legend tests.
- Modify `Config.js` — new palette, `LAYOUT`, guarded `module.exports`.
- Modify `ShiftUtils.js` + `test/shiftUtils.test.js` — new label suffix, `buildCheckLabel`.
- Modify `ScheduleManager.js` — payload styling, unassigned border, wrap/row-height, legend row.
- Modify `package.json` — test script.

## Palette (decided here; tests verify the contrast)

| Key | Hex | Font |
|---|---|---|
| `A320` / `A321` | `#FFF59D` | black |
| `A350` | `#81D4FA` | black |
| `B787` | `#FFCC80` | black |
| `DEFAULT` | `#CE93D8` | black |
| `EA_LAN` (L-HAN) | `#80DEEA` | black |
| `EA_HAN` (HAN) | `#A5D6A7` | black |
| `NIGHT_SHIFT_FLAG` | `#B71C1C` | white, bold |
| `CHANGE_HIGHLIGHT` (border) | `#6A1B9A` | n/a |
| `UNASSIGNED_BORDER` (dashed) | `#424242` | n/a |

Black = `#000000`, white = `#FFFFFF`. Weekend/null backgrounds (`BG_SAT_SUN`, `BG_NULL`) are not touched by the schedule renderer and stay as-is.

---

### Task 1: Palette + contrast utilities

**Files:**
- Create: `ScheduleStyle.js`, `test/scheduleStyle.test.js`
- Modify: `Config.js:51-65`, `Config.js` (append guarded export), `package.json`

**Interfaces:**
- Produces: `contrastRatio(hexA: string, hexB: string): number`; `pickFontColor(bgHex: string): "#000000" | "#FFFFFF"` (whichever has higher contrast); `CONFIG.COLORS` per the palette table; `CONFIG.LAYOUT = { ROW_HEIGHT: 21 }`; `CONFIG.COLORS.UNASSIGNED_BORDER`.

- [ ] **Step 1: Write failing tests** in `test/scheduleStyle.test.js` (same `test()/passed/failed` harness as `test/shiftUtils.test.js`):
  - `contrastRatio("#000000","#FFFFFF")` ≈ 21 (±0.01); `contrastRatio(x,x)` = 1.
  - `pickFontColor("#B71C1C") === "#FFFFFF"`; `pickFontColor("#FFF59D") === "#000000"`.
  - For every fill key in `{A320,A350,B787,DEFAULT,EA_LAN,EA_HAN,NIGHT_SHIFT_FLAG}` from `require("../Config.js").CONFIG.COLORS`: `contrastRatio(fill, pickFontColor(fill)) >= 4.5`; for the six pastel fills `>= 7`.
  - Every `COLORS` value used as a fill/border matches `/^#[0-9A-Fa-f]{6}$/`.
- [ ] **Step 2: Run** `node test/scheduleStyle.test.js` → FAIL (module not found).
- [ ] **Step 3: Implement** `ScheduleStyle.js` using the WCAG relative-luminance formula (sRGB linearization, `0.2126/0.7152/0.0722`); `pickFontColor` returns black unless white has strictly higher contrast. Replace `CONFIG.COLORS` with the palette table, add `CONFIG.LAYOUT`, and add `if (typeof module !== "undefined" && module.exports) module.exports = { CONFIG: CONFIG };` at the end of `Config.js`. Keep existing keys `WHITE`, `BG_SAT_SUN`, `BG_NULL` untouched.
- [ ] **Step 4: Add** `&& node test/scheduleStyle.test.js` to the `test` script; run `npm test` → all PASS.
- [ ] **Step 5: Commit** `feat: contrast-validated schedule palette and ScheduleStyle utils`.

---

### Task 2: Readable labels (`★`, `🌙`, no wrap)

**Files:**
- Modify: `ShiftUtils.js:98-103` and exports; `test/shiftUtils.test.js`

**Interfaces:**
- Consumes: `classifyWP` result (`shiftInfo`, may be null/undefined).
- Produces: `buildShiftLabelSuffix(shiftInfo): string` → `" 🌙"` when `nightShiftRequired`, else `""` (null-safe → `""`); `buildCheckLabel(opts: { acCheck: string, isChk: boolean, shiftInfo: object|null, assigneeShort?: string }): string` → `[★ ]<acCheck><suffix>[ · <assigneeShort>]`, **single line (no `\n`)**.

- [ ] **Step 1: Update/add tests** in `test/shiftUtils.test.js`: replace old `(E)`/`(E N)` suffix assertions with: night → `" 🌙"`, non-night E → `""`, M → `""`, null → `""`; `buildCheckLabel({acCheck:"4C",isChk:true,shiftInfo:{nightShiftRequired:true},assigneeShort:"Nguyen T."})` === `"★ 4C 🌙 · Nguyen T."`; without assignee → `"4C"` (non-chk, null shiftInfo); result never contains `"\n"`. `buildShiftNote` keeps shift start/end lines (shift letters now live only in the note) — assert unchanged.
- [ ] **Step 2: Run** `node test/shiftUtils.test.js` → FAIL.
- [ ] **Step 3: Implement** both functions in `ShiftUtils.js`; export `buildCheckLabel`.
- [ ] **Step 4: Run** `npm test` → PASS.
- [ ] **Step 5: Commit** `feat: single-line check labels with chk star and night moon`.

---

### Task 3: Per-cell style resolution

**Files:**
- Modify: `ScheduleStyle.js`, `test/scheduleStyle.test.js`

**Interfaces:**
- Consumes: `pickFontColor`, `CONFIG.COLORS` (passed in as `colors` param so the file needs no `require`).
- Produces: `resolveCheckStyle(opts: { acType: string, isChk: boolean, nightShiftRequired: boolean, assigned: boolean, changed: boolean }, colors: object): { barColor: string, labelBg: string, labelFont: string, bold: boolean, border: null | { color: string, style: "SOLID_THICK" | "DASHED" } }`.
  Rules: `barColor` by type (A320/A321→`A320`, A350, B787, else `DEFAULT`); `labelBg` = `NIGHT_SHIFT_FLAG` if night else `barColor`; `labelFont = pickFontColor(labelBg)`; `bold = isChk || night`; `border` = changed → `{CHANGE_HIGHLIGHT, "SOLID_THICK"}`, else unassigned → `{UNASSIGNED_BORDER, "DASHED"}`, else `null`.

- [ ] **Step 1: Write failing tests**: CHK + night → `labelFont "#FFFFFF"`, `labelBg "#B71C1C"`, `bold true`; plain A350 → `barColor "#81D4FA"`, font black, `bold false`; `acType "XYZ"` → `DEFAULT`; changed + unassigned → `border.style === "SOLID_THICK"`; unassigned only → `"DASHED"`; assigned & unchanged → `null`; `labelFont` has ≥4.5 contrast on `labelBg` for all 4 type × night/no-night combos.
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement** `resolveCheckStyle` per the rules; export it.
- [ ] **Step 4: Run** `npm test` → PASS.
- [ ] **Step 5: Commit** `feat: resolveCheckStyle with contrast-safe font and border rules`.

---

### Task 4: Wire styles + compact layout into `ScheduleManager.js`

**Files:**
- Modify: `ScheduleManager.js` (`drawChecksBlock` 286-357; Phase Checks 195-229; apply loop 256-263)

**Interfaces:**
- Consumes: `resolveCheckStyle`, `buildCheckLabel`, `CONFIG.LAYOUT.ROW_HEIGHT`.

- [ ] **Step 1: `drawChecksBlock`** — replace the inline `fCol`, `barColor` switch, night override and label building with `resolveCheckStyle` / `buildCheckLabel` (`assigned = !!assignedPerson`, `changed = highlightedPjids.has(pjid)`). Label payload: `color: style.labelFont`, `bold: style.bold`, `bg: style.labelBg`. Bar cells use `style.barColor`; start/end cells use `NIGHT_SHIFT_FLAG` when night (keep current behavior) — their font is set by the label payload only on the start cell; the end cell is empty. Border: if `style.border`, apply `setBorder` on the AC Reg→`toCol` range with `SOLID_THICK` / `DASHED` accordingly (changed is a single call, so it already wins over unassigned).
- [ ] **Step 2: Layout** — replace `setWrap(true)` on the label range (line 347) with `setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)`; delete the per-row `setWrap(false)` TAT>2 loop (now redundant); replace `autoResizeRows(...)` with `currentSheet.setRowHeights(startRow, dataBlock.length, CONFIG.LAYOUT.ROW_HEIGHT)`.
- [ ] **Step 3: Phase Checks** — use `resolveCheckStyle` for font/bg (`acType` from the row, `assigned: true`, `changed: false` — no border for phase checks), `buildCheckLabel` without assignee; change `setWrap(true)` at line 229 to `setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)` and set row heights to `ROW_HEIGHT` for the block (full text stays in the cell note). Station fill still `EA_LAN`/`EA_HAN`, overridden by night red as today.
- [ ] **Step 4: Apply loop** — no change needed beyond ensuring `payload.bg`/`payload.color` are written as before; verify no payload still carries `"red"` (`Grep "red"` in `ScheduleManager.js` returns nothing).
- [ ] **Step 5: Verify** `npm test` → PASS. Manual in the sheet (after `clasp push`, run Update A/C Schedules): CHK+night cell shows white bold on dark red; 1-day check with long description is one line at default height; unassigned bars have a dashed dark border; changed WP has thick purple border on B787 rows too.
- [ ] **Step 6: Commit** `feat: apply contrast-safe styles, unassigned border, fixed row heights`.

---

### Task 5: Legend row

**Files:**
- Modify: `ScheduleStyle.js`, `test/scheduleStyle.test.js`, `ScheduleManager.js`

**Interfaces:**
- Produces: `buildLegendEntries(colors: object): Array<{ label: string, bg: string|null, font: string, border?: string }>` in this order: `A320/A321`, `A350`, `B787`, `Other type`, `Night shift 🌙` (night red/white), `★ CHK`, `Unassigned` (white bg, border `UNASSIGNED_BORDER`), `Changed` (white bg, border `CHANGE_HIGHLIGHT`). Phase-check station colors are 2 extra entries (`L-HAN`, `HAN`) appended last.

- [ ] **Step 1: Write failing tests**: 10 entries; first four `bg`s equal the type palette; night entry `font "#FFFFFF"`; every entry with a `bg` has `contrastRatio(bg,font) >= 4.5`; `★ CHK` entry has black font.
- [ ] **Step 2: Run** → FAIL. **Step 3: Implement** `buildLegendEntries`; export.
- [ ] **Step 4: Draw it** in `ScheduleManager.js`: new `drawLegend(sheet, row, startCol)` called from `runUpdateACSchedules` after the clear (line ~186) on row `CONFIG.ROSTER.LOWER_ROW + 3` (same row as the "PHASE CHECKS" title) starting at column `CONFIG.ROSTER.LEFT_COL`; each entry occupies 3 cells (label in the first, `bg` on all three, `font` color, thin border where `border` is set), 3 columns per entry → 30 columns ≤ 31 day columns. Inside the existing clear range, so it redraws each run.
- [ ] **Step 5: Verify** `npm test` → PASS; manually confirm legend renders, does not overlap the title or the first Phase Check row, and still draws when the month has zero checks.
- [ ] **Step 6: Commit** `feat: on-sheet color legend`.

---

## Self-Review Notes

- Coverage: item 1 (no red text, `★`) → Tasks 3–4; item 2 (readable label, no tall rows) → Tasks 2, 4; item 3 (unassigned flag) → Tasks 3–4; item 4 (font-from-fill + palette + change color) → Tasks 1, 3; legend → Task 5.
- Signature consistency: `resolveCheckStyle(opts, colors)` used identically in Tasks 3–4; `buildCheckLabel` in Tasks 2 and 4.
- Open assumptions to confirm during Task 4/5: the legend row (76) is free of other content, and column widths are wide enough for 3-cell legend entries.
