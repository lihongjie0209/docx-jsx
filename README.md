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

See [the v1 specification](docs/spec.md) for the complete component contract.
