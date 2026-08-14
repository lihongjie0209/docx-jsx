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
```

Validate the complete executable DSL without creating a DOCX file:

```sh
docx-jsx validate report.tsx --data data.json
```

Reverse a generated document back to deterministic JSX:

```sh
docx-jsx reverse report.docx -o report.jsx
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
| Document defaults and metadata | Supported | Fonts, size, character/line spacing, settings, core/custom properties, variables, custom styles |
| Sections | Supported | Page size/orientation/margins, headers/footers, title pages, text direction, document grid, page numbering |
| Paragraphs and headings | Supported | Full line spacing, indentation, frames, borders, IDs, tabs, bidi, outline and tracked format changes |
| Runs and semantic text | Supported | Fonts, themes, emphasis, effects, borders/shading, symbols, breaks, revisions and semantic wrappers |
| Tables | Supported | Grid/width/layout, positioning, margins, all border positions/styles, row revisions and cell layout |
| Fields and indexes | Supported | Generic and typed fields, TOC/figures/entries, captions, index fields and cross-references |
| Annotations and controls | Supported | Bookmarks, comments, footnotes, hyperlinks and bound content controls |
| Images | Supported | Raster image embedding with point-based dimensions |
| DOCX to JSX | Supported | Component-level 1:1 reversal for generated DOCX; strict structural fallback for external DOCX |
| Validation and Agent spec | Supported | `validate`, repair-oriented diagnostics, Markdown contract and JSON Schema |
| Additional package parts | Planned | docx-rs task panes/web extensions and custom XML items |
| Backend-limited APIs | Not emitted | Character-unit paragraph indentation and row before/after grids are dropped by docx-rs 0.4.22 |

The detailed writer-API audit is maintained in
[docs/docx-rs-coverage.md](docs/docx-rs-coverage.md).

See [the v1 specification](docs/spec.md) for the complete component contract.
