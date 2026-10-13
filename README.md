# docx-jsx

Compile executable JSX/TSX to DOCX with an embedded V8 runtime and `docx-rs`.

```tsx
import { Document, Section, Paragraph, Run } from "docx-jsx";

export default ({ name = "World" } = {}) => (
  <Document defaultFont="Arial">
    <Section pageSize="A4">
      <Paragraph align="center">
        <Run bold size={18}>Hello {name}</Run>
      </Paragraph>
    </Section>
  </Document>
);
```

```sh
cargo install --path .
docx-jsx report.tsx --data data.json -o report.docx
docx-jsx report.tsx -o report.pdf
```

Validate the complete executable DSL without creating a DOCX file:

```sh
docx-jsx validate report.tsx --data data.json
```

Reverse a generated document back to deterministic JSX:

```sh
docx-jsx reverse report.docx -o report.jsx
```

Compile can write PDF when `-o` ends with `.pdf`. `docx-jsx pdf` still converts
an existing DOCX. Both use a local office app (not bundled):

```sh
docx-jsx report.tsx -o report.pdf --engine word
docx-jsx pdf report.docx -o report.pdf
docx-jsx pdf report.docx --engine wps --force
```

`auto` uses Microsoft Word, then WPS, then LibreOffice when those are
installed. If none is found, the command prints install steps and how to set
`--engine`, `--soffice`, `DOCX_JSX_SOFFICE`, `DOCX_JSX_WORD`, or `DOCX_JSX_WPS`.

Define inherited Word styles once and reference them to avoid repeating
formatting on every component:

```tsx
const styles = [
  { id: "BodyBase", name: "Body Base", type: "paragraph", paragraph: { spacingAfter: 6 } },
  { id: "Body", name: "Body", type: "paragraph", basedOn: "BodyBase", run: { font: "Arial" } },
  { id: "Emphasis", name: "Emphasis", type: "character", run: { italic: true } },
];

export default (
  <Document styles={styles}>
    <Section><Paragraph style="Body"><Run style="Emphasis">Reusable formatting</Run></Paragraph></Section>
  </Document>
);
```

Print the complete Agent-readable contract or its machine-readable IR schema:

```sh
docx-jsx spec
docx-jsx spec --format json-schema
```

Release archives and installers are published on GitHub for Linux, macOS, and
Windows.

## Feature matrix

| Area | Status | Highlights |
| --- | --- | --- |
| Executable JSX/TSX | Supported | Embedded V8, local ESM imports, async components, JSON data injection |
| Styles and inheritance | Supported | Paragraph/character/numbering/table styles, native `basedOn`, paragraph `next`, linked styles, typed references and cycle detection; paragraph styles preserve page-break-before, widow control, bidi, shading and all six character-unit indentation values; run styles preserve independent Latin/complex-script size, bold and italic, explicit on/off emphasis, caps, strike/double-strike, vertical alignment, character spacing, fit width, stretch, special hidden text and patterned shading in both directions |
| Document defaults and metadata | Supported | Independent ASCII/High ANSI/East Asian/complex-script fonts, font themes/hint, spacing, view zoom, global revision tracking/move/formatting settings, full core properties, settings, document editing restrictions with legacy and Office 2010+ precomputed OOXML hash/salt, custom properties, variables, Web Extensions and Custom XML |
| Sections | Supported | Page/column/continuous/even/odd section breaks, line numbering and restart/distance controls, page size/orientation/margins, equal-width and individually sized columns with per-column spacing/separators, headers/footers, title pages, text direction, document grid, page-number format and chapter separators |
| Paragraphs, headings, and list items | Supported | Full line spacing, point and character-unit indentation (`indentLeftChars`, `indentStartChars`, `indentRightChars`, `indentEndChars`, `firstLineChars`, `hangingChars`), contextual spacing, mirrored inside/outside indents, frame-overlap suppression, nine-slot run-default fonts, frames, borders, IDs, tabs, bidi, outline, paragraph-mark insertion/deletion revisions, and tracked format changes |
| Runs and semantic text | Supported | Independent physical/theme font slots and hint; independent Latin/complex-script size, bold and italic; caps on/off, all Word underline styles (boolean shorthand retained), hidden/special-hidden text, baseline/superscript/subscript alignment, effects, borders/shading, symbols, breaks, tracked text revisions, deleted field instructions (`w:delInstrText`), and semantic wrappers; schema-invalid run-property revision setters are intentionally not exposed |
| Tables | Supported | Grid/width/layout, positioning, margins, point/percent cell widths, no-wrap cells, all border positions/styles, row height/rule, repeatable table header rows, cant-split, gridBefore/gridAfter and point-based widthBefore/widthAfter, row revisions and cell layout; external reverse preserves table, row, and cell properties and errors on unrepresentable values |
| Fields and indexes | Supported | Generic and typed fields; TOC/figures/entries with level filters, page-number ranges, style mappings, separators, bookmark scoping, refresh flags, reversible tracked deletion and field-paragraph formatting (including line-number suppression, character-unit indentation, paragraph fonts, size, emphasis, color, and character spacing) for dynamic indexes, childless dynamic indexes without an outer SDT, before/after paragraphs and tables, static entries with PAGEREF targets/page placeholders, validated dotted-tab positions and OOXML field switches; captions, index fields and cross-references |
| Lists and numbering | Supported | Bullet/ordered shorthand, `styleLink`/`numStyleLink`, OOXML multi-level type, per-level format/text/suffix/restart/legal/indent/run formatting, and per-instance start/full-level overrides; `ListItem` preserves character-unit indentation; external reverse rebuilds `w:numPr` runs as `List` |
| Annotations and controls | Supported | Bookmarks, comments with stable IDs, threaded replies, and resolved state (including comment ranges over deleted text), footnotes, inline and document-level bound content controls (including block controls in headers/footers), and hyperlinks with explicit false/true history plus ContentControl/revision/bookmark/comment children |
| Images | Supported | Raster embedding, dimensions, rotation, accessibility description/title and decorative-image semantics, floating anchors, alignment/offsets, relative origins, distances, overlap, in-cell layout, layer height and signed-EMU `simplePosition` anchors. External reverse extracts body, header, and footer media and preserves accessibility metadata and simple-position coordinates |
| DOCX to JSX | Supported | Component-level 1:1 reversal for generated DOCX; external paragraph and table styles/inheritance including table/cell properties and paragraph contextual spacing, mirrored indents, frame-overlap suppression, page-break/widow/bidi/shading properties, run-style emphasis/caps/vertical alignment/character spacing, all paragraph/style character-unit indentation names, section margins/grids, point-based paragraph indentation and keep/outline, paragraph and inherited style line-number suppression, theme colors, row height/rule/cant-split and revisions including distinct deleted field instructions, body/header/footer raster images with accessible descriptions, titles, and decorative-image semantics, document settings/metadata, custom XML data-store items, task-pane web extensions, nested bookmark/comment ranges with stable threaded-comment IDs and resolved state, and hyperlink composite children. Text boxes are rejected with repair guidance because docx-rs 0.4.22 has no text-box writer and flattening them would lose geometry. Reverse→recompile omits empty comments/footnotes/numbering parts the backend would inject |
| Validation and Agent spec | Supported | DSL validation with repair-oriented diagnostics, Agent contract, .NET Open XML conformance tests; JSON Schema constrains the IR envelope/tree, Document style metadata and run/paragraph/table/cell formatting, and Table, TableRow, TableCell, Paragraph, Heading, Run, Section, Header, Footer, Image, Hyperlink, Bookmark, InlineBookmark, List, ListItem, Caption, TableOfContents, TableOfFigures, TableOfEntries, ContentControl, and native field-component props; remaining component props are documented and semantically validated, with JSON Schema coverage planned in priority batches |
| PDF export | Optional | Compile `-o file.pdf` or `docx-jsx pdf` auto-detects Microsoft Word, WPS, or LibreOffice; `--engine` pins one. Install/config diagnostics when none is present. Not shipped in the release binary |
| Backend-limited APIs | Not emitted | Theme parts and embedded fonts are reader-only in docx-rs 0.4.22; see the coverage audit for missing setters and schema-invalid writer APIs |

The detailed writer-API audit is maintained in
[docs/docx-rs-coverage.md](docs/docx-rs-coverage.md).

See [the v1 specification](docs/spec.md) for the complete component contract.

## Open XML conformance test

The integration gate compiles a representative JSX fixture and validates the
resulting package with Microsoft `DocumentFormat.OpenXml` 3.5.1. Diagnostics
are JSON Lines containing the package part, XPath, explanation, and a repair
suggestion. It requires .NET SDK 8 and restores packages through the configured
Huawei Cloud NuGet mirror.

```sh
./scripts/openxml-test.sh
```

This complements Rust DSL validation: it checks OPC/OOXML schema and semantic
constraints in the generated DOCX, while the Rust tests continue to verify
component mappings and DOCX-to-JSX round trips.
