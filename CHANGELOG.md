# Changelog

## v1.1.0 — 2026-08-18

Word Office Add-in support and internal architecture cleanup.

- Added a Word Task Pane Add-in with a Review-tab ribbon command
- Added direct checking of the currently open Word document on Windows / Mac Word
- Added jump-to-finding actions in Word
- Added safe individual and same-rule bulk fixes, with review-only fixes kept visually distinct
- Added finding anchors and fix metadata without coupling Office behavior into the rule engine
- Added `core.js` as the shared public facade for Standalone and Add-in UIs
- Split the Standalone UI into `src/standalone/` while retaining `src/app.js` as a compatibility entry point
- Added `parseOfficeBuffer()` so File API and Office document snapshots share the same OOXML parser
- Fixed the Standalone initial empty-results placeholder behavior
- Added Word Add-in manifest, task pane UI, icons and sideload/deployment documentation
- Extended CI to cover nested JavaScript, finding metadata regression tests, Add-in CSP, manifest structure and allowed Office.js CDN usage
- Preserved the existing 68-rule deterministic QC behavior

Word on the web can load the task pane, but v1.1.0 does not run the full compressed-OOXML check there.

## v1.0.0 — 2026-08-18

Initial stable release of Document QC.

- Fully static, client-side `.docx` / `.xlsx` quality checking
- 68 generic and public-information-based rules
- Word text, style-hygiene, figure/table reference and structural checks
- Excel error, reference, hidden-element, external-link and formula-pattern checks
- Context-aware typo checks designed to reduce false positives
- Public-information master checks for selected official document names and dates
- Rule enable/disable controls, severity filtering and result search
- Local CSV / XLSX result export
- GitHub Pages deployment support
- CSP `connect-src 'none'`, runtime network guards and no browser persistence of document contents
- CI checks for JavaScript syntax, rule regressions, hidden-state behavior, no-network and no-persistence invariants

v1 intentionally does not include cloud AI or automatic semantic rewriting.
