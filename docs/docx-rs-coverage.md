# docx-rs 0.4.22 coverage

This inventory tracks public document-building APIs that have a stable JSX
representation. Every covered feature must have IR validation, DOCX output
tests, CLI round-trip coverage, and specification text.

Classification used in **Next gaps**:

- **implementable-next** — public writer API, schema-valid output, not yet
  wrapped as JSX/IR
- **upstream-writer-missing** — no public writer setter
- **schema-invalid** — public setter exists but 0.4.22 emits illegal OOXML or
  drops the value
- **reader-only** — present on the reader/model, not a writer API

| Area | Covered | Next gaps |
| --- | --- | --- |
| Run content and formatting | text, breaks, raster images with relationship IDs, dimensions, rotation, floating positioning, overlap, offsets/alignments, relative origins, distances and layer height, symbols, fields, tabs, all nine `RunFonts` slots, size, RGB and theme color/tint/shade, highlight, explicit bold/italic/strike/double-strike on/off, underline, caps, hidden, spacing, scaling, fit, border, shading, revisions, footnotes. External reverse recovers raster `Pic` drawings from `Docx.images` / `word/media` and theme color modifiers from OOXML | **schema-invalid**: image simple-position is hardcoded by the upstream serializer. **reader-only**: `Shape` children (`todo!` in the writer), leftover reader-only run children. Vector drawings (EMF/WMF/PICT) error on reverse |
| Paragraph | alignment, style, all line-spacing fields/rules, indentation, keep flags, page break, tabs, outline, snap-to-grid, widow control, bidi, text alignment, right-indent adjustment, shading, frame properties, paragraph borders, explicit paragraph IDs, paragraph run defaults including all nine font slots, formatting insert/delete and property revisions, nested `InlineBookmark`/`Comment`/`ContentControl` reconstructed by matching range IDs | **schema-invalid**: character-unit indentation (`hanging_chars` / `first_line_chars`) is dropped by the serializer |
| Hyperlink | `href`/`anchor`/`history`, run children, and every public `Hyperlink::add_*` child: `ContentControl` (`add_structured_data_tag`), `Inserted`, `Deleted`, `InlineBookmark` (`add_bookmark_start`/`end`), `Comment` (`add_comment_start`/`end`). External reverse reconstructs those children without the IR manifest, including `w:sdt` that the 0.4.22 hyperlink reader flattens | **upstream-writer-missing**: no public `history` setter (the field is written directly). Nested hyperlinks are rejected |
| Section | page size, orientation, margins, headers, footers, title page, text direction, document grid, page numbering. External reverse emits `margins` and `documentGrid` from `w:sectPr` | **upstream-writer-missing**: section type, columns, and line numbering have no public `Section` setters |
| Table | width, alignment, layout, grid, uniform/positioned/cleared borders, style, indentation, cell margins, floating position | none in the public writer API |
| TableRow | height/rule, cant-split, inserted/deleted revisions | **schema-invalid**: `grid_before` / `grid_after` / `width_before` / `width_after` are dropped by the serializer |
| TableCell | width, span, vertical alignment/merge, text direction, margins, shading, uniform/positioned/diagonal/cleared borders, structured data tags, table of contents | **reader-only**: cell indexes |
| Document | all nine `RunFonts` default slots (four physical fonts, four theme fonts and hint), size/character-spacing and complete paragraph-spacing defaults, sections, numbering, comments, footnotes, indexes, created/modified metadata, string custom properties, document ID/variables, default tab stop, odd/even headers, table line-height compatibility, character-spacing control, task panes/web extensions, multiple custom XML data-store items, custom paragraph/character/numbering/table style definitions with metadata and valid run/paragraph/frame/table/cell properties and text borders, including positioned/cleared style borders and style `keepNext`/`keepLines`/`outlineLevel`; document-level `Bookmark` ranges reverse as nested JSX; document-level `add_structured_data_tag` as a section `ContentControl` with block children. A user `Normal` style replaces the backend default instead of duplicating it. Unused empty `comments` / `commentsExtended` / `footnotes` / `numbering` parts the writer injects are stripped when the document has no matching components | **reader-only**: theme *parts*. **schema-invalid**: table-style width/layout/nested-style setters emit illegal `StyleTableProperties` and are rejected. **upstream-writer-missing**: remaining core properties and document protection. Header/footer `StructuredDataTag` has no public writer setter and reverse errors instead of flattening |

Reader-only docx-rs APIs and constructors marked private by docx-rs are not
counted as build coverage until the backend exposes a public writer API.
