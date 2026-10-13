# Changelog

## 0.8.9 - 2026-10-13

- Extend the Agent JSON Schema to constrain raster `Image` source/dimensions,
  accessibility metadata, decorative-image exclusivity, floating-anchor
  properties, and signed DrawingML simple-position coordinates.

## 0.8.8 - 2026-10-13

- Extend the Agent JSON Schema to constrain Document defaults, metadata,
  settings, styles, web extensions, custom XML items, and editing protection,
  including font-slot and zoom/revision bounds.

## 0.8.7 - 2026-10-13

- Extend the Agent JSON Schema to constrain Section page size, orientation,
  margins, text direction, section breaks, document grid, page/line numbering,
  and equal-width or individually sized column layouts. Reject custom page
  sizes with missing, unknown, or non-positive dimensions during DSL validation.

## 0.8.6 - 2026-10-13

- Extend the Agent JSON Schema to constrain common `Run` formatting, including
  underline/theme enums, color formats, font slots, positive sizes, enabled-only
  hidden flags, and incompatible strike combinations.

## 0.8.5 - 2026-10-13

- Extend the Agent JSON Schema to constrain `Paragraph` and `Heading` layout,
  paragraph-mark revision, font-default, and formatting properties, including
  the required heading level and mutually exclusive indentation/revision forms.

## 0.8.4 - 2026-10-13

- Make the Agent JSON Schema constrain `TableRow` and `TableCell` properties,
  including types, bounds, required margins, enums, and mutually exclusive
  values; document that other component properties are still covered by the
  Markdown specification and semantic validator pending future batches.
- Correct `TableCell.widthPercent` documentation to match validation: greater
  than 0 and at most 100.

## 0.8.3 - 2026-10-13

- Add `TableCell.noWrap` to control cell text wrapping and preserve explicit
  enabled/disabled values on external DOCX reverse without the IR manifest.
- Add schema-valid OOXML, invalid-value, CLI round-trip, and Microsoft Open XML
  SDK coverage, including no-wrap values on cells with other table properties.

## 0.8.2 - 2026-10-13

- Add `TableRow.header` to repeat table header rows across page breaks and
  preserve enabled/disabled states in external DOCX reverse without the IR
  manifest.
- Normalize generated table-row property order to valid OOXML when row headers,
  sizing, omitted grid cells, and revisions are combined.
- Add CLI round-trip, invalid-property, compiler XML, and Microsoft Open XML SDK
  coverage for repeating table headers.

## 0.8.1 - 2026-10-13

- Fix Rust 1.99 Clippy's single-element-loop lint in image accessibility
  reverse validation so the v0.8.0 feature set passes the hosted CI toolchain.

## 0.8.0 - 2026-10-13

- Add `Image.decorative` to explicitly mark an image as decorative, compiling
  an empty `wp:docPr/@descr` and restoring the property on external reverse.
- Preserve per-drawing accessibility metadata when identical image media is
  reused, and assign unique drawing property IDs required by Open XML.
- Validate decorative-image property combinations and add same-media,
  manifest-free round-trip and Microsoft Open XML SDK coverage.

## 0.7.0 - 2026-10-13

- Add `Image.description` and `Image.title` for accessible image metadata,
  preserving `wp:docPr/@descr` and `@title` when compiling and reversing
  external DOCX packages without the embedded IR manifest.
- Validate non-empty image metadata and relationship IDs, cover escaped XML
  values and image metadata in simple-position anchors, and run the fixture
  through Microsoft Open XML SDK validation.

## 0.6.0 - 2026-10-13

- Preserve paragraph `mirrorIndents` and `suppressOverlap` on paragraphs,
  headings, list items, and paragraph styles when compiling and reversing
  external DOCX files.
- Keep explicit on/off values in OOXML and validate the paragraph flags as
  booleans; extend the Agent specification, feature matrix, and Open XML
  integration fixture.

## 0.5.0 - 2026-08-15

- Add optional PDF export through a local Microsoft Word, WPS Writer, or
  LibreOffice install. `docx-jsx report.tsx -o report.pdf` infers PDF from
  the output suffix; `docx-jsx pdf` converts an existing DOCX.
- Auto-detect Word, then WPS, then LibreOffice on Windows; Word then
  LibreOffice on macOS; LibreOffice on Linux. Pin with `--engine`,
  `--soffice`, `DOCX_JSX_WORD`, `DOCX_JSX_WPS`, or `DOCX_JSX_SOFFICE`.
- Expose `List.levels` for schema-valid numbering writer APIs and rebuild
  contiguous `w:numPr` paragraphs as `List`/`ListItem` on external reverse.
- Recover header/footer raster images from their own relationship parts,
  restore document settings and `customXmlItems` including `schemaRefs`,
  and keep positional tabs, signed indentation, and start/end paragraph
  alignment from real Word packages.

## 0.4.0 - 2026-08-15

- Reconstruct hyperlink composite children (`ContentControl`, `Inserted`,
  `Deleted`, `InlineBookmark`, `Comment`) on compile and external reverse.
- Compile document-level block `ContentControl` through
  `Docx::add_structured_data_tag`, and reverse body/cell structured tags
  without the IR manifest.
- Preserve distinct section page sizes, and reverse TOC/TOF/TOE fields as
  dedicated components instead of empty content controls.
- Reverse raster drawings as `Image`, extract package media beside the JSX,
  and restore section `margins` / `documentGrid`.
- Restore paragraph and style `keepNext`, `keepLines`, and `outlineLevel`,
  plus run `themeColor` / `themeShade` / `themeTint` from OOXML.
- Omit unused empty `comments`, `footnotes`, and `numbering` parts the
  backend would otherwise inject, and skip a stock `Normal` style so
  recompilation does not duplicate it.
- Add `reverse_package` so `docx-jsx reverse` writes extracted image assets
  next to the JSX module.

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
