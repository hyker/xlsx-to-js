# Production readiness — Spec

## Pitch
Implement the reviewed release requirements for bounded spreadsheet previews, correct hidden-cell merges, explicit resource budgets, browser validation, responsive loading, and documented fidelity limits.

## Assumptions
- The user's instruction to implement the reviewed list authorizes implementation without additional planning gates.
- Keep existing full-HTML export APIs. Add a bounded page API and use it for the interactive demo.
- Use row and column pagination, with at most 5,000 positions per preview page.
- Preserve original coordinates and clip merges and drawings to each page.
- Parsing runs in a terminable browser worker using a strict XML parser; callbacks remain on the caller side.

## Stories
1. S1: Preview large sheets with a bounded DOM and navigate rows, columns, and sheets.
2. S2: Preserve layout for hidden rows/columns and merged regions, including page boundaries.
3. S3: Reject oversized inputs and report resource-limit failures with consistent parse/render budgets.
4. S4: Cancel parsing, report progress, and prevent stale loads from replacing newer results.
5. S5: Validate security/correctness with regression tests and the complete UI in browser engines.
6. S6: Reduce duplicate XML work, date-formatting overhead, grid allocations, and repeated cell CSS.
7. S7: Document supported inputs, budgets, package worker integration, and fidelity limitations.

## Domain model
A Workbook contains WorkSheets. A WorkSheet has positioned cells, styles, merges, and drawings. A SheetPage selects visible row and column coordinates and returns HTML plus navigation metadata. A worker parse request owns a worker lifetime, cancellation signal, and progress callbacks.

## Non-goals
- Excel formula evaluation, complete format fidelity, or unlimited workbooks.
- Changing full HTML export to silently truncate its output.
