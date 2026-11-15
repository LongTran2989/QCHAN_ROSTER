# A/C Schedule layout fixes (follow-up to colors/labels PR)

## Intent
Phase checks are the WPs where staff go in person and supervise. The sheet should surface those
first and legibly, and the layout must never overwrite cells.

## Goals
- Phase block shows only CHK WPs (TAT 0.5 / 1 day). Label has no star and no moon.
- Night coverage in Phase = the reg cell (row below the painted label) in bold red text on the unfilled cell.
- HAN / L-HAN fills are distinguishable from each other and from every A/C-type color and the night-red.
- Normal / Other short / STO labels become `[star ]REG Desc[ moon][ · Assignee]` (reg also stays in left column).
- Non-CHK short-TAT WPs move to an "OTHER SHORT CHECKS" group directly below the last Normal row,
  drawn like Normal rows (bar, borders, assignee, PJID in col 38) with HAN/L-HAN fills.
- Layout: PHASE CHECKS -> NORMAL CHECKS -> OTHER SHORT CHECKS -> END OF LIST -> STO. END OF LIST stays
  between Other short and STO (Utils.js export uses NORMAL CHECKS..END OF LIST, so Other short is exported with Normal).
- No block overlaps another; no cell written twice.
- Shift notes: times only ("Start: dd/mm HH:mm (Morning shift)"), no "Hanoi/Bangkok time".

## Non-goals
Assigning personnel to Phase WPs; snapshot / change-log / contrast rules; changing Normal-row A/C-type colors;
moon and start/end night tint on Normal and STO rows (kept as is).

## Approaches considered
- Patch the existing counter (`eaCheckBlockLength`): rejected, it is the source of the overlap bug.
- **Chosen:** one row cursor + first-free-row packing for Phase WPs, extracted into a pure tested function
  (like ScheduleStyle.js), with a "no cell written twice" test.

## Decisions
- Colors: HAN `#66BB6A` (green), L-HAN `#A1887F` (taupe). Min CIELAB distance to existing fills: 44 / 28
  (old L-HAN cyan was 15 from A350 blue). Black text passes WCAG 4.5:1 on both.
- Reg red for night: a dark red that passes 4.5:1 on white (e.g. `#C62828`), bold.
- Label separator is a plain space (assignee parsing splits on the first " · ").

## Left to the builder
Exact packing function signature, test names, where the pure layout module lives.
