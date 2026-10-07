# Production readiness — Level 1 Plan

## Status
Implemented. Both XML implementations pass all 64 regression/adversarial tests. All 30 browser checks pass across Chromium, Firefox, and WebKit, including the 100,000-cell styled-date fixture, cancellation, upload races, embedded images, and full-export merge alignment. Library/example type checks, package entry smoke checks, and Storybook production build pass.

The large fixture mounts 5,000 cells (5,377 demo DOM nodes); the measured first display was about 1.1 seconds on Chromium with 4× CPU throttling in this environment. Timings and available heap measurements are recorded by the browser suite and are not device guarantees.

## Plan
1. **Parser corrections and allocation improvements** — Infer references in one pass, expose drawing relationships, reuse date formatters and resolved styles, fill dense blanks in place. `serves: S2, S6` [not expanded]
2. **Bounded preview renderer** — Add `toHTMLSheetPage(workbook, sheetIndex, options): XlsxSheetPage`, visible-axis pagination and merge clipping; share CSS rules and precompute column widths. `serves: S1, S2, S6` [not expanded]
3. **Worker loading lifecycle** — Add a worker entry and `XlsxWorkerParser.readFile`, progress, cancellation, timeout, and cleanup. `serves: S3, S4` [not expanded]
4. **Demo integration** — Use explicit preview budgets, preflight uploads, pagination, cancellation, stale-load protection, and lifecycle cleanup. `serves: S1, S3, S4` [not expanded]
5. **Regression and browser checks** — Run security cases with both XML implementations; add merged/hidden/page tests and browser fixtures with performance measurements and responsive/cancellation checks. `serves: S2, S5, S6` [not expanded]
6. **Documentation and CI** — Explain production integration, inputs and fidelity; automate unit, browser and demo build checks. `serves: S5, S7` [not expanded]

## Open questions
None required to implement the approved scope. Performance timings are measured and recorded rather than presented as universal device guarantees.
