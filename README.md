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
| Styles and inheritance | Supported | Paragraph/character/numbering/table styles, native `basedOn`, paragraph `next`, linked styles, typed references and cycle detection |
| Document defaults and metadata | Supported | Independent ASCII/High ANSI/East Asian/complex-script fonts, font themes/hint, spacing, settings, properties, variables, Web Extensions and Custom XML |
| Sections | Supported | Page size/orientation/margins, headers/footers, title pages, text direction, document grid, page numbering |
| Paragraphs and headings | Supported | Full line spacing, indentation, nine-slot run-default fonts, frames, borders, IDs, tabs, bidi, outline and tracked format changes |
| Runs and semantic text | Supported | Independent physical/theme font slots and hint, emphasis, effects, borders/shading, symbols, breaks, revisions and semantic wrappers |
| Tables | Supported | Grid/width/layout, positioning, margins, all border positions/styles, row revisions and cell layout |
| Fields and indexes | Supported | Generic and typed fields, TOC/figures/entries, captions, index fields and cross-references |
| Lists and numbering | Supported | Bullet/ordered shorthand plus per-level format, text, suffix, restart, legal numbering, indent and run formatting; external reverse rebuilds `w:numPr` runs as `List` |
| Annotations and controls | Supported | Bookmarks, comments, footnotes, inline and document-level bound content controls, and hyperlinks with ContentControl/revision/bookmark/comment children |
| Images | Supported | Raster embedding, dimensions, rotation, floating anchors, alignment/offsets, relative origins, distances, overlap and layer height. External reverse extracts body, header, and footer package media beside the JSX |
| DOCX to JSX | Supported | Component-level 1:1 reversal for generated DOCX; external styles/inheritance, section margins/grids, paragraph keep/outline, theme colors, body/header/footer raster images, document settings/metadata, custom XML data-store items, revisions, nested bookmark/comment ranges, hyperlink composite children and footer text-box fallback. Reverse→recompile omits empty comments/footnotes/numbering parts the backend would inject |
| Validation and Agent spec | Supported | DSL validation, repair-oriented diagnostics, Agent contract, JSON Schema, and .NET Open XML conformance tests |
| PDF export | Optional | Compile `-o file.pdf` or `docx-jsx pdf` auto-detects Microsoft Word, WPS, or LibreOffice; `--engine` pins one. Install/config diagnostics when none is present. Not shipped in the release binary |
| Backend-limited APIs | Not emitted | Character-unit paragraph indentation, row before/after grids, image simple-position, theme parts and embedded fonts are dropped or hardcoded by docx-rs 0.4.22 |

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
