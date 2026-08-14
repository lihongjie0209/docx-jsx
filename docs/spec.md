# docx-jsx v1 specification

`docx-jsx` compiles an executable JSX or TSX module to a DOCX archive. The
compiler is a single Rust executable. It embeds V8; Node and Deno are not
runtime dependencies.

## Command line

```text
docx-jsx <INPUT.jsx|tsx> [-o <OUTPUT.docx>] [--data <DATA.json>] [--force]
docx-jsx validate <INPUT.jsx|tsx> [--data <DATA.json>]
docx-jsx reverse <INPUT.docx> [-o <OUTPUT.jsx>] [--force]
docx-jsx spec [--format markdown|json-schema]
```

When `-o` is omitted, the input extension is replaced with `.docx`. Existing
outputs are rejected unless `--force` is present. Compilation completes in
memory and output is written atomically, so a failed compilation never leaves
a partial DOCX.

`validate` runs module resolution, JSX/TSX transpilation, optional JSON data
injection, module evaluation, IR decoding, and the complete component semantic
validation pass without invoking the DOCX backend or writing an output file.
It prints `valid: INPUT` and exits with status 0 on success. Invalid DSL uses
the standard detailed diagnostic and exits with status 1.

The entry module must default-export either a document JSX value or a function
receiving the JSON value supplied by `--data`. The function and nested function
components may return promises. With no `--data`, the root function receives
`undefined`.

## JavaScript environment

The compiler accepts `.js`, `.jsx`, `.ts`, `.tsx`, and `.json`. Relative and
absolute local ESM imports support explicit filenames, extension lookup in the
order `.ts`, `.tsx`, `.js`, `.jsx`, `.json`, and directory `index.*` lookup.
Bare imports are limited to `docx-jsx` and `docx-jsx/jsx-runtime`. npm, remote
URLs, Node built-ins, and Deno APIs are not supported.

JSX uses the automatic runtime. `Fragment`, arrays, and nested arrays are
flattened. `null`, `undefined`, and booleans render nothing. String and numeric
children are text. Pure whitespace between structural elements is ignored.

## Component tree

| Parent | Allowed children |
| --- | --- |
| `Document` | one or more `Section` |
| `Section` | `Header`, `Footer`, `Heading`, `Paragraph`, `Caption`, `Index`, `TableOfContents`, `TableOfFigures`, `TableOfEntries`, `Bookmark`, `Table`, `List` |
| `Header`, `Footer` | `Paragraph`, `Caption`, `Table`, `List` |
| `Paragraph` | `Run`, `Text`, `Break`, `CarriageReturn`, `NonBreakingSpace`, `SoftHyphen`, `NonBreakingHyphen`, `Image`, `Hyperlink`, `ContentControl`, `Field`, `DateField`, `TimeField`, `FileNameField`, `AuthorField`, `TitleField`, `SubjectField`, `SequenceField`, `ReferenceField`, `MergeField`, `DocumentPropertyField`, `FormulaField`, `IndexEntry`, `Comment`, `Inserted`, `Deleted`, `MovedFrom`, `MovedTo`, `Footnote`, `Tab`, `TabStop`, `PositionalTab`, `Symbol`, `Bold`, `Italic`, `Underline`, `StrikeThrough`, `Superscript`, `Subscript`, `AllCaps`, `HiddenText`, `SpecialHiddenText`, `DoubleStrike`, `SpacedText`, `ScaledText`, `FitText`, `BorderedText`, `ShadedText`, `PageNumber`, `TotalPages`, `PageReference`, `TocEntry`, string, number |
| `Heading` | same inline children as `Paragraph` |
| `Caption` | same inline children as `Paragraph` |
| `Run` | `Text`, `Break`, `CarriageReturn`, `NonBreakingSpace`, `SoftHyphen`, `NonBreakingHyphen`, `Image`, `Footnote`, `Tab`, `Symbol`, string, number |
| `Comment` | inline children selected by the comment |
| `Footnote` | `Run`, `Text`, `Break`, `Image`, `Tab`, `Symbol`, string, number |
| `Inserted` | run-level inline content |
| `Deleted` | `Run`, `Text`, string, number |
| `MovedFrom`, `MovedTo` | run-level inline content |
| `Hyperlink` | `Run`, `Text`, `Break`, `Image`, string, number |
| `ContentControl` | `Run`, `Text`, `Break`, `CarriageReturn`, `Image`, `Footnote`, `Tab`, `PositionalTab`, `Symbol`, `PageReference`, string, number |
| `Field` | `Run`, `Text`, `Break`, `CarriageReturn`, `Image`, `Tab`, `PositionalTab`, `Symbol`, string, number |
| `DateField`, `TimeField`, `FileNameField`, `AuthorField`, `TitleField`, `SubjectField`, `SequenceField`, `ReferenceField`, `MergeField`, `DocumentPropertyField`, `FormulaField` | same result children as `Field` |
| `Text` | string, number |
| `Bold`, `Italic`, `Underline`, `StrikeThrough`, `Superscript`, `Subscript`, `AllCaps`, `HiddenText`, `SpecialHiddenText`, `DoubleStrike`, `SpacedText`, `ScaledText`, `FitText`, `BorderedText`, `ShadedText` | `Text`, string, number |
| `Table` | `TableRow` |
| `TableRow` | `TableCell` |
| `TableCell` | `Paragraph`, `Caption`, `Table`, `List`, `TableOfContents`, `ContentControl` |
| `List` | one or more `ListItem` |
| `ListItem` | `Run`, `Text`, `Break`, `Image`, string, number |
| `Bookmark` | `Heading`, `Paragraph`, `Caption`, `Table`, `List` |
| `TableOfContents` | none |
| `TableOfFigures`, `TableOfEntries` | none |
| `Index`, `IndexEntry` | none |

Text, breaks, and images placed directly in a paragraph are wrapped in an
implicit run. Unknown components, unknown properties, and invalid nesting are
errors and include a component path in diagnostics.

## Properties

All numeric dimensions use points (`pt`). Font sizes convert to half-points,
layout dimensions to twentieths of a point, and image dimensions to EMU.
Colors are uppercase or lowercase six-digit RGB strings without `#`.

- `Table`: optional non-empty `style`, signed `indent` in points, and `margins`
  object with required non-negative `top`, `right`, `bottom`, and `left` point
  values. Optional `position` supports non-negative `leftFromText` and
  `rightFromText`, `verticalAnchor`/`horizontalAnchor: "margin" | "page" |
  "text"`, `xAlign: "center" | "inside" | "left" | "outside" | "right"`,
  `yAlign: "bottom" | "center" | "inline" | "inside" | "outside" | "top"`,
  and signed point coordinates `x`/`y`. `x` conflicts with `xAlign`; `y`
  conflicts with `yAlign`.
- `TableRow`: optional `inserted` or `deleted` revision metadata object with
  string `author?` and `date?`; the two revision states are mutually exclusive.
- `TableCell`: optional `verticalMerge: "restart" | "continue"`,
  `textDirection: "lr" | "lrV" | "rl" | "rlV" | "tb" | "tbV" | "tbRlV" |
  "tbRl" | "btLr" | "lrTbV"`, and a `margins` object with the same four
  required point values as `Table.margins`.

## DOCX reverse conversion

`docx-jsx reverse INPUT.docx [-o OUTPUT.jsx] [--force]` reads a DOCX with
`docx-rs` and emits deterministic JSX that can be compiled by this tool again.
The default output is `INPUT.jsx`. The command refuses to replace an existing
file unless `--force` is supplied.

For DOCX files produced by this compiler, reverse conversion is component-level
1:1: every v1 component, property, child, and scalar is restored from a
normalized IR manifest embedded in the package. External DOCX files fall back
to structural OOXML conversion; unsupported structures are reported as errors
instead of being silently discarded. The generated module imports only the
components it uses and default-exports one `Document`.

- `Document`: `defaultFont?: string`, positive `defaultSize?: number`, signed
  `defaultCharacterSpacing?: number` (points), `createdAt?: string`,
  `updatedAt?: string`, `customProperties?: Record<string,string>`,
  `documentId?: string`, positive `defaultTabStop?: number` (points),
  `documentVariables?: Record<string,string>`, `evenAndOddHeaders?: boolean`,
  `adjustLineHeightInTable?: boolean`, and
  `characterSpacingControl?: "doNotCompress" | "compressPunctuation" | "compressPunctuationAndJapaneseKana"`.
  Metadata strings, IDs, property names, and variable names must be non-empty;
  maps must not be empty. These values are emitted to the standard core,
  custom-property, styles, and settings package parts.
  `defaultLineSpacing?: {before?,after?,line?,beforeLines?,afterLines?,lineRule?}`
  configures default paragraph properties. `before`, `after`, and `line` are
  finite point values (`before` and `after` are non-negative);
  `beforeLines` and `afterLines` are non-negative integers in OOXML
  hundredths-of-a-line units; `lineRule` is `auto`, `atLeast`, or `exact`.
  The object must configure at least one field.
- `Document.styles?: StyleDefinition[]` defines reusable Word styles. Every
  definition requires unique non-empty `id`, non-empty `name`, and `type` of
  `paragraph`, `character`, `numbering`, or `table`. Optional metadata is
  `basedOn`, `next`, `link`, `quickFormat`, `uiPriority`, `semiHidden`, and
  `unhideWhenUsed`. `run?: {font,size,color,themeColor,themeShade,themeTint,highlight,bold,italic,underline,hidden,textBorder}`
  uses the same units and color rules as `Run`.
  `paragraph?: {align,textAlign,snapToGrid,spacingBefore,spacingAfter,lineSpacing,spacingBeforeLines,spacingAfterLines,lineRule,indentLeft,indentRight,firstLine,hanging,hangingChars,firstLineChars,outlineLevel,frame}`
  uses point dimensions; `firstLine` and `hanging` are mutually exclusive.
  `textAlign` is `auto`, `baseline`, `bottom`, `center`, or `top`; line-spacing
  fields use the same units and `lineRule` values as `Paragraph`.
  `frame` accepts `wrap`, `verticalAnchor`, `horizontalAnchor`, `heightRule`,
  `xAlign`, `yAlign`, `horizontalSpace`, `verticalSpace`, `x`, `y`, `width`,
  and `height`; its dimensions are points and coordinate/alignment pairs are
  mutually exclusive.
  Table styles additionally accept
  `table?: {style,indent,width,widthPercent,align,layout,margins,border}` and
  `cell?: {width,colSpan,verticalAlign,verticalMerge,textDirection,shading,margins,border}`.
  These reuse the corresponding `Table` and `TableCell` property formats;
  `width`/`widthPercent` are mutually exclusive. `run.textBorder` reuses the
  `BorderedText` object shape. `table` and `cell` are rejected on non-table
  styles because docx-rs would otherwise omit them during serialization.
- `Section`: `pageSize?: "A4" | "Letter" | {width,height}`,
  `orientation?: "portrait" | "landscape"`, `margins?: {top,right,bottom,left,header?,footer?,gutter?}`,
  `titlePage?: boolean`, `textDirection?: "lrTb" | "tbRl" | "btLr" | "lrTbV" | "tbRlV"`,
  `documentGrid?: {type,linePitch?,charSpace?}`, and
  `pageNumbering?: {start?,chapterStyle?}`. Document-grid `type` is `default`,
  `lines`, `linesAndChars`, or `snapToChars`; `linePitch` is a positive point
  measurement, while `charSpace` is a signed integer in OOXML 1/4096-em units.
  `pageNumbering.start` is a non-negative integer and `chapterStyle` is a
  non-empty Word chapter-style value. Both configuration objects must be
  non-empty.
- `Paragraph`: `align?: "left" | "center" | "right" | "both"`,
  `spacingBefore?`, `spacingAfter?`, `lineSpacing?`, `indentLeft?`,
  `indentRight?`, `firstLine?`, `hanging?`, `keepNext?`, `keepLines?`,
  `pageBreakBefore?`. `firstLine` and `hanging` are mutually exclusive.
  Optional `spacingBeforeLines?`, `spacingAfterLines?`, and
  `lineRule?: "auto" | "atLeast" | "exact"` expose the remaining docx-rs
  line-spacing controls. The `*Lines` values are non-negative integers in
  OOXML hundredths-of-a-line units and may be combined with point spacing.
- `Paragraph`: also accepts `snapToGrid?`, `widowControl?`, `bidi?`,
  `textAlign?: "auto" | "baseline" | "bottom" | "center" | "top"`, signed
  integer `adjustRightIndent?`, six-digit RGB `shading?`, and integer
  `outlineLevel?: 0..9`. `frame?` has the same shape, units, enumerations, and
  coordinate/alignment conflicts as `Document.styles[].paragraph.frame`.
  Optional `inserted?` and `deleted?` objects mark the paragraph's default run
  formatting as revised and are mutually exclusive; both accept non-empty
  `author?` and `date?` strings. `propertyChange?` records previous paragraph
  formatting as `{author?, date?, previous}`. `previous` is non-empty and
  accepts the same paragraph formatting properties except `inserted`,
  `deleted`, and `propertyChange`.
- `Paragraph`: optional `paragraphId` is exactly eight hexadecimal characters
  and is emitted as `w14:paraId`. `border` accepts either a uniform
  `{style?,size?,color?,space?}` object or positioned entries `top`, `right`,
  `bottom`, `left`, `between`, and `bar`. Each positioned entry is a border
  object or `false` for an explicit `nil` border; `{clearAll:true}` clears all
  six positions. Uniform and positioned forms cannot be mixed. Border sizes
  use points and `space` is a non-negative integer in OOXML point units. All
  border-bearing components accept every `BorderType` token exposed by the
  pinned docx-rs backend, from `nil`, `none`, `single`, `thick`, `double`,
  `dotted`, and `dashed` through compound, wave, 3D, and art values such as
  `thinThickThinLargeGap`, `doubleWave`, `threeDEmboss`, and `babyRattle`.
  The complete style set is `nil`, `none`, `single`, `thick`, `double`,
  `dotted`, `dashed`, `dotDash`, `dotDotDash`, `triple`,
  `thinThickSmallGap`, `thickThinSmallGap`, `thinThickThinSmallGap`,
  `thinThickMediumGap`, `thickThinMediumGap`, `thinThickThinMediumGap`,
  `thinThickLargeGap`, `thickThinLargeGap`, `thinThickThinLargeGap`, `wave`,
  `doubleWave`, `dashSmallGap`, `dashDotStroked`, `threeDEmboss`,
  `threeDEngrave`, `outset`, `inset`, `apples`, `archedScallops`,
  `babyPacifier`, and `babyRattle`.
- `Paragraph`: also accepts paragraph-default
  `font?`, positive `size?`, `bold?`, `italic?`, six-digit RGB `color?`, and
  signed `characterSpacing?` in points.
- `Heading`: required `level: 1..9`, optional `style`, plus all `Paragraph`
  layout properties except `outlineLevel`. Its `level` emits both a heading
  style and the corresponding Word outline level.
- `Run`: `themeColor?` accepts Word theme tokens `dark1`, `light1`, `dark2`,
  `light2`, `accent1` through `accent6`, `hyperlink`, `followedHyperlink`,
  `none`, `background1`, `text1`, `background2`, or `text2`. `themeShade?` and
  `themeTint?` are two-digit hex modifiers. Boolean `bold`, `italic`, `strike`,
  and `doubleStrike` emit explicit on/off OOXML; enabled strike modes are
  mutually exclusive.
- `Caption`: required Word identifier `label`; optional Word identifier
  `identifier?` defaults to `label`. `format?`, `restart?`, `placeholder?`, and
  `dirty?` have the same meaning as on `SequenceField`; `style?` defaults to
  `Caption`, `numberSeparator?` defaults to one space, and `textSeparator?`
  defaults to `: `. It emits a styled paragraph containing the visible label,
  a native `SEQ` field, the separator, and its inline children.
- `Run`: `font?`, `size?`, `bold?`, `italic?`, `strike?`, `underline?`,
  `color?`, `highlight?`, `style?`.
- `Paragraph`: also accepts `style?: string` for custom or built-in styles.
- `Text`: optional `value`; it may alternatively contain string/number children.
- `Break`: `type?: "line" | "page" | "column"` (default `line`).
- `CarriageReturn`: no properties or children; emits a run-level `w:cr`.
- `NonBreakingSpace`, `SoftHyphen`, `NonBreakingHyphen`: no properties or
  children. They emit Unicode U+00A0, U+00AD, and U+2011 respectively, keeping
  adjacent words together, exposing an optional hyphenation point, or keeping
  a hyphenated term together. They are valid wherever run-level text is valid.
- `Image`: required `src`, `width`, and `height`. Images are inline. Relative
  sources resolve from the entry module directory.
- `Table`: `width?`, `widthPercent?`, `align?`, `layout?: "auto" | "fixed"`,
  `columnWidths?: number[]`, `border?: {style?,size?,color?}`. Point and percent
  widths are mutually exclusive.
- `TableRow`: `height?`, `heightRule?: "auto" | "atLeast" | "exact"`,
  `cantSplit?`.
- `TableCell`: `width?`, `colSpan?`,
  `verticalAlign?: "top" | "center" | "bottom"`, `shading?`, `border?`.
- Table, TableCell, and their custom-style equivalents also accept an advanced
  border object. Table positions are `top`, `right`, `bottom`, `left`,
  `insideHorizontal`, and `insideVertical`; cells additionally accept
  `topLeftToBottomRight` and `topRightToBottomLeft`. Each position is either a
  `{style?,size?,color?}` border or `false` to emit an explicit `nil` border.
  `clearAll: true` clears every supported position and cannot be combined with
  position entries. Uniform `{style?,size?,color?}` remains supported and
  cannot be mixed with position entries.
- `Header`, `Footer`: `type?: "default" | "first" | "even"`. A `first`
  header or footer enables the section's different-first-page setting.
- `Hyperlink`: exactly one of `href?: string` (external relationship) or
  `anchor?: string` (document bookmark), plus `history?: boolean`.
- `PageNumber`, `TotalPages`: no properties. They emit dynamic Word fields.
- `Comment`: required non-empty `text` plus selected inline children;
  `author?: string`, `date?: string`. It emits a comment range, reference, and
  an entry in `comments.xml`.
- `Footnote`: requires inline content and emits a native footnote reference and
  entry in `footnotes.xml`.
- `Tab`: no properties or children; emits a run-level tab.
- `TabStop`: required non-negative `position` in points,
  `align?: "left" | "center" | "right" | "decimal" | "bar" | "clear"`, and
  `leader?: "none" | "dot" | "heavy" | "hyphen" | "middleDot" | "underscore"`.
  It defines a paragraph tab stop (`w:tabs/w:tab`); pair it with `Tab` to move
  content to that stop.
- `Symbol`: required `font` and four-hex-digit `char` (for example
  `<Symbol font="Wingdings" char="F0A7" />`).
- `Superscript`, `Subscript`, `AllCaps`, `HiddenText`: no properties and
  require text content. Each emits an independent Run with native
  `w:vertAlign`, `w:caps`, or `w:vanish` formatting, so surrounding text is
  unaffected.
- `Bold`, `Italic`, `StrikeThrough`: no properties and require text content;
  they emit independent runs with native `w:b`, `w:i`, or `w:strike` formatting.
  `Underline` requires text content and accepts `type?: "single" | "double" |
  "dotted" | "dash" | "wave"` (default `single`), emitting native `w:u`.
- `DoubleStrike`: no properties; emits native `w:dstrike` formatting.
- `SpacedText`: required finite `amount` in points. Positive values expand and
  negative values condense character spacing; the value is converted to
  signed twips.
- `ScaledText`: required integer `percent` from 1 to 600; emits native
  horizontal character scaling (`w:w`).
- `FitText`: required positive `width` in points and optional unsigned integer
  `id`. Word compresses or expands the contained text to the requested width
  using native `w:fitText` formatting.
- `BorderedText`: optional `style?: "single" | "double" | "dotted" |
  "dashed"`, positive `size?` in points, six-digit RGB `color?`, and
  non-negative integer `space?` in points. It emits a native `w:bdr` run
  border around the contained text.
- `ShadedText`: required six-digit RGB `fill`, optional six-digit RGB `color`,
  and optional OOXML `pattern` (default `clear`). Supported patterns are the
  native clear/solid, stripe/cross, thin-stripe/cross, and `pct5` through
  `pct95` values exposed by docx-rs. It emits character-level `w:shd`.
- `SpecialHiddenText`: no properties. It emits `w:specVanish`, which Word uses
  for text that should remain hidden even when ordinary hidden text is shown.
- `Inserted`, `Deleted`: `author?: string`, `date?: string`; both require
  content. They emit native tracked revisions (`w:ins`/`w:del`). Deleted text
  is serialized as `w:delText`, preserving optional `Run` formatting.
- `MovedFrom`, `MovedTo`: `author?: string`, `date?: string`; both require
  content and emit native move revision containers (`w:moveFrom`/`w:moveTo`).
- `TocEntry`: required non-empty `text`, optional `level?: 1..9`,
  `omitPageNumber?: boolean`, and non-empty `identifier?: string`. It emits a
  native hidden `TC` field so arbitrary paragraph content can appear in a
  table of contents without being styled as a heading.
- `PageReference`: required `bookmark`, optional `hyperlink?`,
  `relativePosition?`, `placeholder?`, `dirty?`. It emits an updateable
  `PAGEREF` field and requires a matching `Bookmark` in the document.
- `PositionalTab`: `align?: "left" | "center" | "right"`,
  `relativeTo?: "margin" | "indent"`, and
  `leader?: "none" | "dot" | "heavy" | "hyphen" | "middleDot" | "underscore"`.
- `ContentControl`: optional `alias`, `xpath`, `prefixMappings`, and
  `storeItemId` strings. It requires inline content and emits a native Word
  structured document tag (`w:sdt`). When any data-binding property is used,
  `xpath` is required and the values are emitted as `w:dataBinding` attributes.
- `Field`: required non-empty `instruction`, optional `dirty?: boolean`
  (default `true`), and optional inline result content. It emits a complex Word
  field (`begin`, `instrText`, `separate`, result, `end`). This supports native
  field codes such as `DATE`, `DOCPROPERTY`, `FORMULA`, and `MERGEFIELD` without
  adding a separate JSX component for every instruction.
- `DateField`, `TimeField`: optional non-empty `format` without quote/control
  characters and optional `dirty?: boolean`; they emit native `DATE` and
  `TIME` fields. `FileNameField` accepts `fullPath?: boolean` and `dirty?` and
  emits `FILENAME` with the `\\p` switch when requested. `AuthorField`,
  `TitleField`, and `SubjectField` accept `dirty?` and emit the corresponding
  built-in document-property fields. All six accept optional inline result
  content and default `dirty` to `true`.
- `SequenceField`: required identifier matching `[A-Za-z_][A-Za-z0-9_]*`,
  optional `format?: "arabic" | "roman" | "Roman" | "alphabetic" |
  "Alphabetic"`, non-negative integer `restart?`, string `placeholder?`, and
  `dirty?`. It emits a native `SEQ` field for captions and independent number
  series. `ReferenceField`: required `bookmark`, optional `hyperlink?`,
  `relativePosition?`, `placeholder?`, and `dirty?`; it emits a native `REF`
  field and requires a matching `Bookmark` in the document.
- `IndexEntry`: required non-empty `text`; optional non-empty `subentry` and
  `identifier`, `boldPageNumber?`, `italicPageNumber?`, `pageRangeBookmark?`,
  and `crossReference?`. The page-range and cross-reference properties are
  mutually exclusive. It emits a hidden native `XE` field.
- `Index`: optional non-empty `identifier` filters matching entries,
  `columns?: 1..4`, `runIn?: boolean`, `placeholder?: string`, `dirty?: boolean`,
  and non-empty `style?: string`. It emits a native `INDEX` complex field in a
  paragraph and defaults its placeholder to `Update index`.
- `MergeField`: required Word identifier `name`, optional
  `preserveFormatting?` (default `true`), `placeholder?`, and `dirty?`.
  `DocumentPropertyField`: required non-empty `name` without quote/control
  characters plus `placeholder?` and `dirty?`. `FormulaField`: required
  non-empty `expression`, optional `numberFormat` without quote/control
  characters, `placeholder?`, and `dirty?`. They emit native `MERGEFIELD`,
  `DOCPROPERTY`, and formula (`=`) complex fields.
- `List`: `type?: "bullet" | "ordered"` (default `bullet`) and optional
  positive integer `start`.
- `ListItem`: `level?: 0..8` (default `0`). Each list compiles to native Word
  numbering definitions and numbered paragraphs.
- `Bookmark`: required non-empty `name`. Its structural children are enclosed
  by a unique Word bookmark pair and may be targeted by `Hyperlink anchor`.
- `TableOfContents`: `startLevel?: 1..9`, `endLevel?: 1..9`,
  `hyperlinks?: boolean`, `dirty?: boolean`, `alias?: string`. Defaults to
  levels 1–3, hyperlinks enabled, and automatic refresh requested. A section
  accepts at most one TOC, placed before its body content (headers and footers
  may precede it).
- `TableOfFigures`: required non-empty `label` (for example `Figure`),
  `includeLabelAndNumber?: boolean`, `separator?: string`,
  `hyperlinks?: boolean`, `dirty?: boolean`, and `alias?: string`. It emits a
  native TOC field filtered by caption/SEQ label.
- `TableOfEntries`: required non-empty `identifier`, plus `hyperlinks?`,
  `dirty?`, and `alias?`. It emits a native custom TOC containing `TocEntry`
  fields with the same `identifier`.

Each section accepts at most one of each index component. All index components
must precede body content; headers and footers may appear before them.

Widths and font/image sizes must be positive. Spacing, margins, indents, border
sizes, and row heights must be non-negative. Percent width is in `(0, 100]`.

## IR and diagnostics

The JavaScript runtime emits the versioned JSON envelope described by
`spec/ir-v1.schema.json`. Rust deserializes the envelope into typed structures
and independently validates structural and semantic constraints.

Every compiler diagnostic includes the concrete error, a phase-specific
explanation, an actionable repair suggestion, and the `docx-jsx spec` discovery
command. Failures exit with a non-zero status.
