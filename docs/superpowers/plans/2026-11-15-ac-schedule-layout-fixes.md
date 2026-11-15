# A/C Schedule Layout Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Phase block shows only CHK WPs (no star/moon, night = bold red reg), non-CHK short WPs move to an OTHER SHORT CHECKS group under NORMAL CHECKS, Normal/Other/STO labels carry the reg, HAN/L-HAN get new distinguishable colors, notes drop the timezone wording, and no block can overwrite another.

**Architecture:** Replace the fragile `eaCheckBlockLength` counter in `ScheduleManager.js` with a pure, node-tested layout module (`ScheduleLayout.js`: slot packing + one row cursor). New style/label helpers go in the existing pure modules (`ScheduleStyle.js`, `ShiftUtils.js`); `ScheduleManager.js` only wires them to the sheet.

**Tech Stack:** Google Apps Script (V8, shared global scope, no imports), plain-Node tests (`npm test`), clasp.

**Spec:** `docs/superpowers/specs/2026-11-15-ac-schedule-layout-fixes-design.md` (read it first).

## Global Constraints

- New `.js` files use the guarded `if (typeof module !== "undefined" && module.exports)` export pattern (see `ScheduleStyle.js`) so they work as GAS globals and under `require()`. Top level contains only `var`/`function` declarations.
- Palette: `EA_HAN = "#66BB6A"`, `EA_LAN = "#A1887F"`, new `NIGHT_REG_FONT = "#C62828"`. All hex. Red is never a font color on a colored fill; red reg font sits only on the unfilled (white) reg cell.
- Every (fill, font) pair keeps WCAG >= 4.5:1 (existing tests enforce it).
- Section markers in column B must stay exactly `NORMAL CHECKS` and `END OF LIST` (the `Utils.js` export parses them). The new header text is `OTHER SHORT CHECKS`; it must not contain either marker string.
- Layout order: PHASE CHECKS -> NORMAL CHECKS -> OTHER SHORT CHECKS -> END OF LIST -> STO.
- Notes show times only: `Start: 05/01 22:00 (Evening shift)`; no "Hanoi"/"Bangkok" wording.
- Keep `npm test` green; add new test files to the `test` script in `package.json`. Do not run `clasp push`.

## Review Focus

- Many Phase WPs starting the same day, then a shallow column: NORMAL header must never land inside the Phase block (Task 3 property test).
- Zero CHK phase WPs / zero Other-short / zero Normal / zero STO: END OF LIST still exists between Normal(+Other) and STO, and no "OTHER SHORT CHECKS" header when empty (Task 3).
- Night CHK Phase WP at HAN and at L-HAN: fill stays the station color with black text; only the reg font turns red; non-night reg font is untouched (Task 2).
- Assign-by-double-click on a new-format label (`★ A350-1 4C · X`): `setLabelAssignee` keeps star+reg (Task 1).
- Palette: new HAN/L-HAN stay clearly apart from every A/C-type fill and the night red (Task 2 distance test).

---

### Task 1: Labels with reg, timezone-free notes

**Files:**
- Modify: `ShiftUtils.js` (`buildShiftNote`, `buildCheckLabel`, header comments)
- Test: `test/shiftUtils.test.js`

**Interfaces:**
- Produces: `buildCheckLabel(opts: {acCheck: string, isChk: boolean, shiftInfo: object|null, assigneeShort?: string, acReg?: string}) -> string` = `[★ ][<acReg> ]<acCheck>[ 🌙][ · <assignee>]`, newlines flattened. `acReg` omitted/empty adds nothing, so callers that don't pass it behave as before.
- Produces: `buildShiftNote` lines are `Start: dd/mm HH:mm (<Morning|Evening> shift)` / `End: ...`, then the unchanged night warning line.

- [ ] **Step 1: Write failing tests** in `test/shiftUtils.test.js`:
  - `buildCheckLabel({acReg:"A350-1", acCheck:"4C", isChk:true, shiftInfo:{nightShiftRequired:true}, assigneeShort:"Nguyen T."})` === `"★ A350-1 4C 🌙 · Nguyen T."`
  - `buildCheckLabel({acReg:"A350-1", acCheck:"4C", isChk:false, shiftInfo:null})` === `"A350-1 4C"`
  - `buildCheckLabel({acCheck:"4C", isChk:false, shiftInfo:null})` still `"4C"` (existing test stays).
  - `acReg` containing a newline never yields `\n` in the label.
  - `setLabelAssignee("★ A350-1 4C 🌙", "Tran B.")` === `"★ A350-1 4C 🌙 · Tran B."`; `setLabelAssignee("A350-1 4C · Nguyen T.", "")` === `"A350-1 4C"`.
  - Update the existing `buildShiftNote` test to expect `"Start: 05/01 22:00 (Evening shift)"` and `"End: 06/01 03:00 (Morning shift)"`, and assert the note contains neither `"Hanoi"` nor `"Bangkok"`.
- [ ] **Step 2: Run** `node test/shiftUtils.test.js` — Expected: new/updated tests FAIL.
- [ ] **Step 3: Implement** the signature above in `ShiftUtils.js`; remove `, Hanoi time` from `buildShiftNote` and drop "Hanoi local time" wording from its doc comment (keep the UTC+7 explanation out of user-visible text only).
- [ ] **Step 4: Run** `node test/shiftUtils.test.js` — Expected: all PASS.
- [ ] **Step 5: Commit** `git add ShiftUtils.js test/shiftUtils.test.js && git commit -m "feat: reg in check labels, timezone-free shift notes"`

### Task 2: New station colors and Phase/Other style helpers

**Files:**
- Modify: `Config.js` (COLORS), `ScheduleStyle.js`
- Test: `test/scheduleStyle.test.js`

**Interfaces:**
- Produces in `ScheduleStyle.js`:
  - `colorDistance(hexA: string, hexB: string) -> number` — CIELAB delta E76 (D65, sRGB).
  - `stationFill(station: string, colors) -> string` — `"L-HAN"` -> `colors.EA_LAN`, anything else -> `colors.EA_HAN`.
  - `resolvePhaseStyle(opts: {station: string, nightShiftRequired: boolean}, colors) -> {labelBg: string, labelFont: string, bold: true, regFont: string|null, regBold: boolean}` — `labelBg = stationFill(...)` (never the night red), `labelFont = pickFontColor(labelBg)`, `regFont = colors.NIGHT_REG_FONT` and `regBold = true` only when night, else `null`/`false`.
  - `resolveCheckStyle` gains optional `opts.fill: string`; when given, `barColor = opts.fill` (night still overrides `labelBg`). Existing behavior unchanged when absent.
- Config: set `EA_HAN`, `EA_LAN`, add `NIGHT_REG_FONT` (values in Global Constraints); update the comments (L-HAN taupe, HAN green, reg red font).

- [ ] **Step 1: Write failing tests** in `test/scheduleStyle.test.js`:
  - `COLORS.EA_HAN === "#66BB6A"`, `COLORS.EA_LAN === "#A1887F"`, `COLORS.NIGHT_REG_FONT === "#C62828"` (add `NIGHT_REG_FONT` to the 6-digit hex test).
  - `contrastRatio(COLORS.NIGHT_REG_FONT, "#FFFFFF") >= 4.5`.
  - `colorDistance("#000000","#000000") === 0`; for each of EA_HAN and EA_LAN, `colorDistance` to each of A320, A350, B787, DEFAULT, NIGHT_SHIFT_FLAG and to the other EA color is `>= 25`.
  - `stationFill("L-HAN", COLORS)` is EA_LAN; `stationFill("HAN", COLORS)` is EA_HAN.
  - `resolvePhaseStyle({station:"HAN", nightShiftRequired:true})` -> `labelBg` EA_HAN, `labelFont "#000000"`, `regFont "#C62828"`, `regBold true`; same with `"L-HAN"` -> EA_LAN; with `nightShiftRequired:false` -> `regFont null`, `regBold false`. Label font has >= 4.5:1 contrast on both station fills.
  - `resolveCheckStyle({..., fill: COLORS.EA_LAN})` -> `barColor` EA_LAN, `labelBg` EA_LAN (not night), and with `nightShiftRequired:true` `labelBg` is `#EF9A9A`.
  - In the palette loops: EA_LAN leaves `PASTEL_KEYS` (taupe is 6.3:1 on black, below the 7:1 pastel bar) and joins `FILL_KEYS` only (>= 4.5:1); EA_HAN stays in both. Fix the duplicated `NIGHT_SHIFT_FLAG` entry while there.
- [ ] **Step 2: Run** `node test/scheduleStyle.test.js` — Expected: FAIL.
- [ ] **Step 3: Implement** the Config values and the three functions/option in `ScheduleStyle.js`, add them to its `module.exports`.
- [ ] **Step 4: Run** `node test/scheduleStyle.test.js` — Expected: all PASS.
- [ ] **Step 5: Commit** `git add Config.js ScheduleStyle.js test/scheduleStyle.test.js && git commit -m "feat: distinct HAN/L-HAN colors, phase style helper"`

### Task 3: Pure layout module

**Files:**
- Create: `ScheduleLayout.js`, `test/scheduleLayout.test.js`
- Modify: `package.json` (append `&& node test/scheduleLayout.test.js` to the `test` script)

**Interfaces:**
- Produces:
  - `PHASE_MIN_ROWS` = `9` (global `var`).
  - `assignPhaseSlots(cols: number[]) -> {slots: number[], depth: number}` — `slots[i]` = count of earlier entries in `cols` equal to `cols[i]` (0-based stack index per day column); `depth = max(slots)+1`, `0` for an empty input.
  - `computeScheduleLayout(opts: {phaseHeaderRow: number, phaseDepth: number, normalCount: number, otherCount: number, stoCount: number}) -> {phaseHeaderRow, phaseStartRow, phaseRowCount, normalHeaderRow, normalStartRow, otherHeaderRow: number|null, otherStartRow: number|null, endOfListRow, stoStartRow}` with:
    - `phaseStartRow = phaseHeaderRow + 1`; `phaseRowCount = max(PHASE_MIN_ROWS, 2*phaseDepth + 1)` (Phase WP k in a column sits at `phaseStartRow + 2*slot`, its reg row one below);
    - `normalHeaderRow = phaseStartRow + phaseRowCount`; `normalStartRow = normalHeaderRow + 1`;
    - if `otherCount > 0`: `otherHeaderRow = normalStartRow + normalCount + 1`, `otherStartRow = otherHeaderRow + 1`, `endOfListRow = otherStartRow + otherCount + 2`; else both `null` and `endOfListRow = normalStartRow + normalCount + 2`;
    - `stoStartRow = endOfListRow + 1`.
  - With `phaseHeaderRow=76, phaseDepth=1, otherCount=0` this reproduces today's rows (normal header 86, start 87, END OF LIST `87+n+2`).

- [ ] **Step 1: Write failing tests** (`test/scheduleLayout.test.js`, same harness style as the other test files):
  - `assignPhaseSlots([5,5,5,7,5,7])` -> slots `[0,1,2,0,3,1]`, depth `4`; `assignPhaseSlots([])` -> `{slots:[], depth:0}`.
  - Property: for random `cols` arrays, no two entries share `(col, slot)`.
  - `computeScheduleLayout({phaseHeaderRow:76, phaseDepth:1, normalCount:10, otherCount:0, stoCount:3})` -> `normalHeaderRow 86`, `normalStartRow 87`, `otherHeaderRow null`, `endOfListRow 99`, `stoStartRow 100`.
  - Same with `otherCount:4` -> `otherHeaderRow 98`, `otherStartRow 99`, `endOfListRow 105`, `stoStartRow 106`.
  - `phaseDepth:0` and `phaseDepth:1` both give `phaseRowCount 9`; `phaseDepth:6` gives `13` and `normalHeaderRow = 77 + 13`.
  - Property over `phaseDepth` 0..12, counts 0..40: the last Phase row used (`phaseStartRow + 2*(depth-1) + 1`) is `< normalHeaderRow`; Normal rows end before `otherHeaderRow`/`endOfListRow`; Other rows end before `endOfListRow`; `stoStartRow === endOfListRow + 1`; rows strictly increase section to section; `endOfListRow` is never inside any data range.
  - Zero counts everywhere still returns a valid `endOfListRow > normalStartRow`.
- [ ] **Step 2: Run** `node test/scheduleLayout.test.js` — Expected: FAIL (module missing).
- [ ] **Step 3: Implement** `ScheduleLayout.js` per the signatures above.
- [ ] **Step 4: Run** `npm test` — Expected: every file reports 0 failed.
- [ ] **Step 5: Commit** `git add ScheduleLayout.js test/scheduleLayout.test.js package.json && git commit -m "feat: pure schedule layout (slot packing + row cursor)"`

### Task 4: Wire it into the sheet

**Files:**
- Modify: `ScheduleManager.js` (`runUpdateACSchedules`, `drawChecksBlock`), `AssignmentManager.js` (label-format comment at ~line 120)

**Interfaces:**
- Consumes: everything from Tasks 1-3.
- `drawChecksBlock(dataBlock, startRow, assignedPersonByPjid, highlightedPjids, renderQueue, shiftInfoByRow, currentSheet, borderOps, useStationFill)` — new trailing boolean; when true, each row's style is `resolveCheckStyle({..., fill: stationFill(row[SCHEDULE_INDEX.STATION], CONFIG.COLORS)})`. Its `buildCheckLabel` call now passes `acReg: acReg`.

- [ ] **Step 1: Split the short-TAT list** after `filteredData_EA` is sorted (currently ~line 192): `phaseChk` = rows whose PJID contains `"CHK"`, `otherShort` = the rest (both keep the `comparator` order). Rows stay in the snapshot/diff exactly as today; only rendering changes. A non-CHK short WP whose PJID contains `STO` goes to `otherShort`, not the STO block (spec decision).
- [ ] **Step 2: Phase block** (replaces the `eaCheckBlockLength`/`paintRow` collision code, ~lines 195-234): compute `assignPhaseSlots` over each `phaseChk` row's day column, then `computeScheduleLayout({phaseHeaderRow: CONFIG.ROSTER.LOWER_ROW + 3, phaseDepth: depth, normalCount: filteredDataNormal.length, otherCount: otherShort.length, stoCount: filteredDataSTO.length})` once (split Normal/STO before this call). For each Phase WP: label cell at `(phaseStartRow + 2*slot, col)` with value `buildCheckLabel({acCheck, isChk:false, shiftInfo:null})` (no star, no moon), fill/font/bold from `resolvePhaseStyle({station, nightShiftRequired})`, note as today; reg cell directly below with today's `nsLabel` text, `regFont`/`regBold` applied (payloads need a `color`+`bold` pair; use `regFont`/`regBold`). Apply `CLIP` + `setRowHeights(phaseStartRow, phaseRowCount, ROW_HEIGHT)` over `phaseRowCount` rows.
- [ ] **Step 3: Section positions**: write headers/rows from the layout object: `"NORMAL CHECKS"` at `normalHeaderRow`, `drawChecksBlock(normal, normalStartRow, ...)`; if `otherShort.length > 0`, `"OTHER SHORT CHECKS"` at `otherHeaderRow` (column `LEFT_COL - 1`) and `drawChecksBlock(otherShort, otherStartRow, ..., true)`; `"END OF LIST"` at `endOfListRow` column 2; STO at `stoStartRow`. Delete the old `normalCheckStartRow`/`listEndRow`/`stoStartRow` arithmetic.
- [ ] **Step 4: Update the assignment comment** in `AssignmentManager.js` to `[star ][REG ]AC_CHECK[ moon][ middot ShortName]`; no logic change.
- [ ] **Step 5: Verify**: `npm test` -> all PASS; `node --check ScheduleManager.js` and `node --check AssignmentManager.js` -> no output; `grep -n "eaCheckBlockLength\|listEndRow\|normalCheckStartRow" ScheduleManager.js` -> no matches.
- [ ] **Step 6: Commit** `git add ScheduleManager.js AssignmentManager.js && git commit -m "feat: phase CHK-only block, OTHER SHORT CHECKS group, row-cursor layout"`
- [ ] **Step 7: Manual sheet check** (needs the live sheet; report as "not run" if unavailable): after `clasp push` by the user, run Update A/C Schedules on a month with deep same-day Phase stacks and confirm: Phase has only CHK WPs with no star/moon and red bold reg under night ones; HAN green vs L-HAN taupe; OTHER SHORT CHECKS sits directly under Normal; END OF LIST between it and STO; nothing overwritten; Normal/Other/STO labels read `★ REG Desc`; double-click assign on an Other-short row works; notes read `Start: ... (Morning shift)`; the Utils.js sheet export still finds both markers.
