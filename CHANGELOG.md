# Changelog

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
