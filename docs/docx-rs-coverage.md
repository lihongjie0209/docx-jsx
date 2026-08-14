# docx-rs 0.4.22 coverage

This inventory tracks public document-building APIs that have a stable JSX
representation. Every covered feature must have IR validation, DOCX output
tests, CLI round-trip coverage, and specification text.

| Area | Covered | Next gaps |
| --- | --- | --- |
| Run content and formatting | text, breaks, images, symbols, fields, tabs, fonts, size, RGB and theme color/tint/shade, highlight, explicit bold/italic/strike/double-strike on/off, underline, caps, hidden, spacing, scaling, fit, border, shading, revisions, footnotes | remaining public run-property combinations and reader-only run children |
| Paragraph | alignment, style, spacing, indentation, keep flags, page break, tabs, outline through Heading, snap-to-grid, widow control, paragraph run defaults | character-unit indentation (upstream serializer currently drops it), frame properties, property revisions |
| Section | page size, orientation, margins, headers, footers | section type, columns, page numbering, line numbering, title page |
| Table | width, alignment, layout, grid, borders | style, indentation, margins, floating position |
| TableRow | height/rule, cant-split | before/after grid widths, revisions |
| TableCell | width, span, vertical alignment, shading, borders | vertical merge, text direction, structured data tags, cell indexes |
| Document | defaults, sections, numbering, comments, footnotes, indexes | styles, settings, themes, custom properties, document protection |

Reader-only docx-rs APIs and constructors marked private by docx-rs are not
counted as build coverage until the backend exposes a public writer API.
