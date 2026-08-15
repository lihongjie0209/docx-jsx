# docx-rs 0.4.22 coverage

This inventory tracks public document-building APIs that have a stable JSX
representation. Every covered feature must have IR validation, DOCX output
tests, CLI round-trip coverage, and specification text.

| Area | Covered | Next gaps |
| --- | --- | --- |
| Run content and formatting | text, breaks, raster images with relationship IDs, dimensions, rotation, floating positioning, overlap, offsets/alignments, relative origins, distances and layer height, symbols, fields, tabs, fonts, size, RGB and theme color/tint/shade, highlight, explicit bold/italic/strike/double-strike on/off, underline, caps, hidden, spacing, scaling, fit, border, shading, revisions, footnotes | image simple-position flag is hardcoded by the upstream serializer; remaining public run-property combinations and reader-only run children |
| Paragraph | alignment, style, all line-spacing fields/rules, indentation, keep flags, page break, tabs, outline, snap-to-grid, widow control, bidi, text alignment, right-indent adjustment, shading, frame properties, paragraph borders, explicit paragraph IDs, paragraph run defaults, formatting insert/delete and property revisions | character-unit indentation (upstream serializer currently drops it) |
| Section | page size, orientation, margins, headers, footers, title page, text direction, document grid, page numbering | section type and columns (no public `Section` setters), line numbering (no public writer API) |
| Table | width, alignment, layout, grid, uniform/positioned/cleared borders, style, indentation, cell margins, floating position | none in the public writer API |
| TableRow | height/rule, cant-split, inserted/deleted revisions | before/after grid widths (upstream serializer currently drops them) |
| TableCell | width, span, vertical alignment/merge, text direction, margins, shading, uniform/positioned/diagonal/cleared borders, structured data tags, table of contents | cell indexes are reader-only |
| Document | all nine `RunFonts` default slots (four physical fonts, four theme fonts and hint), size/character-spacing and complete paragraph-spacing defaults, sections, numbering, comments, footnotes, indexes, created/modified metadata, string custom properties, document ID/variables, default tab stop, odd/even headers, table line-height compatibility, character-spacing control, task panes/web extensions, multiple custom XML data-store items, custom paragraph/character/numbering/table style definitions with metadata and valid run/paragraph/frame/table/cell properties and text borders, including positioned/cleared style borders | theme *parts* are reader-only; table-style width/layout/nested-style setters emit schema-invalid `StyleTableProperties` in docx-rs 0.4.22 and are rejected; remaining core properties have no public setters; document protection is not modeled by the writer |

Reader-only docx-rs APIs and constructors marked private by docx-rs are not
counted as build coverage until the backend exposes a public writer API.
