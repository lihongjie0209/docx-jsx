# Changelog

## 0.3.1 - 2026-08-15

- Harden domestic-mirror downloads in CI and cargo-dist release builds with retries,
  longer timeouts, and tolerance for transient low-speed connections.
- Republish the 0.3 feature set after the v0.3.0 Apple ARM64 artifact build was
  interrupted by a transient rsproxy.cn timeout.

## 0.3.0 - 2026-08-14

- Add typed paragraph, character, numbering, and table style references with
  inheritance-cycle, `next`, and reciprocal linked-style validation.
- Preserve external DOCX style definitions and inheritance metadata, including
  `basedOn`, `next`, `link`, quick-format, visibility, and UI-priority values.
- Preserve exact paragraph line spacing, line rules, zero spacing, and
  first-line/hanging indentation during external DOCX reversal.
- Preserve external paragraph, run, and table style references, tracked
  revisions, whitespace, and inline bookmark anchors.
- Recover footer content from external floating text boxes and normalize
  section-property ordering for schema-valid header/footer references.
- Reject table-style properties that docx-rs 0.4.22 serializes as invalid
  OOXML, with targeted repair suggestions.
- Expand style inheritance, external reversal, CLI diagnostics, round-trip,
  and Microsoft Open XML SDK integration coverage.

## 0.2.0 - 2026-08-14

- Add advanced paragraph formatting, complete line spacing, paragraph borders,
  identifiers, and tracked property revisions.
- Add positioned and cleared table, cell, paragraph, and style borders.
- Add Office task panes, multiple Web Extensions, and multiple Custom XML
  data-store items with corrected OPC content types.
- Add advanced raster image rotation and floating-anchor positioning, relative
  origins, distances, overlap, and layer height.
- Add Microsoft Open XML SDK 3.5.1 integration validation with structured,
  repair-oriented diagnostics.
- Normalize generated WordprocessingML property, style, and settings element
  order so representative output passes the Microsoft 365 Open XML validator.
- Expand the Agent specification, reverse/recompile coverage, README feature
  matrix, local/GPU test workflow, and GitHub CI.

## 0.1.0 - 2026-08-14

- Initial public release of the Rust JSX/TSX-to-DOCX CLI with incremental
  caching, DOCX-to-JSX reversal, DSL validation, and Agent-readable specs.
