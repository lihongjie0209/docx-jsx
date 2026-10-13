# docx-jsx v1 specification

`docx-jsx` compiles an executable JSX or TSX module to a DOCX archive. The
compiler is a single Rust executable. It embeds V8; Node and Deno are not
runtime dependencies.

## Command line

```text
docx-jsx <INPUT.jsx|tsx> [-o <OUTPUT.docx|pdf>] [--data <DATA.json>] [--engine auto|libreoffice|word|wps] [--soffice <PATH>] [--force]
docx-jsx validate <INPUT.jsx|tsx> [--data <DATA.json>]
docx-jsx reverse <INPUT.docx> [-o <OUTPUT.jsx>] [--force]
docx-jsx pdf <INPUT.docx> [-o <OUTPUT.pdf>] [--engine auto|libreoffice|word|wps] [--soffice <PATH>] [--force]
docx-jsx spec [--format markdown|json-schema]
```

`spec --format json-schema` describes the IR envelope and recursive tree. In
this schema revision, `Document`, `TableRow`, `TableCell`, `Paragraph`,
`Heading`, `Run`, `Section`, `Image`, `Hyperlink`, `List`, `ListItem`, `Table`,
`Caption`, `Header`, `Footer`, `Bookmark`, `InlineBookmark`, and
`TableOfContents`, `TableOfFigures`, `TableOfEntries`, `ContentControl`, the
native field components, `TocItem`, `PageReference`, `TabStop`, `PositionalTab`,
`Symbol`, and tracked revision properties are additionally typed and constrained;
other component `props` remain generic
objects pending later coverage batches. The
schema constrains style definition metadata and all four nested style property
groups; style identity uniqueness, reference resolution/type compatibility,
and inheritance cycles remain the responsibility of `validate`. List override
level uniqueness and legal component placement are likewise checked by
`validate`.
`TableCell.widthPercent` must be greater than 0 and at most 100.
The Table schema also constrains width exclusivity, percentage widths, layout,
grid widths, margins, borders, and floating-position coordinate/alignment
exclusivity. Caption schema validates the required Word label identifier,
optional sequence identifier/format/restart/placeholder/dirty values, and style
and separator strings.
`StyleDefinition.run`, `.paragraph`, `.table`, and `.cell` receive typed schemas
for formatting, effects, layout, borders, and shading; cross-definition style
identity/reference semantics remain authoritative in `validate`.
Header and Footer schema properties constrain their optional story type to
`default`, `first`, or `even`.
Bookmark and InlineBookmark schema properties require non-empty anchor names.
TableOfContents schema covers level filters, field instruction options, style
level mappings, deletion metadata, and generated field-paragraph formatting;
cross-property and child-dependent rules remain checked by `validate`.
TableOfFigures and TableOfEntries schemas require their respective non-empty
`label` and `identifier`, constrain field switches and shared index formatting,
and reject unknown properties; index placement and child-dependent rules
remain checked by `validate`.
ContentControl schema constrains its alias and data-binding strings, and requires
both a non-empty `xpath` and `storeItemId` whenever any binding property is used.
Child-kind and placement rules remain checked by `validate`.
Field component schemas constrain required instructions/identifiers, switch
types, sequence formats, Word identifiers, and quoted field arguments; bookmark
resolution and XML control-character rules remain checked by `validate`.
Static `TocItem` rows require text and a bookmark, constrain level and tab
position, and `PageReference` shares the strict bookmark-target schema with
`ReferenceField`.
Tab-stop and positional-tab schemas constrain alignment and leader enums, while
`Symbol` requires a font and four hexadecimal character code.
Tracked revision schemas allow optional author/date metadata and require at least
one child for `Inserted`, `Deleted`, `MovedFrom`, and `MovedTo`.

When `-o` is omitted, the input extension is replaced with `.docx`. The output
format is inferred from the `-o` suffix: `.docx` writes a DOCX archive; `.pdf`
compiles the document in memory and converts that archive with the same local
office detection as `docx-jsx pdf`. The intermediate DOCX is not kept beside
the PDF. Other suffixes are rejected. `--engine` and `--soffice` on the compile
command apply only to `.pdf` output and are rejected when writing `.docx`.
Existing outputs are rejected unless `--force` is present. Compilation completes
in memory and output is written atomically, so a failed compilation never leaves
a partial DOCX. The compiler then drops backend-injected annotation and
numbering parts that the document does not use: `word/comments.xml`,
`word/commentsExtended.xml`, `word/footnotes.xml`, and `word/numbering.xml`,
together with their document relationships and content-type overrides. A
document with no `Comment`, `Footnote`, or `List` therefore does not grow
those parts on reverse→recompile. The parts are written only when the
corresponding components are present.

`validate` runs module resolution, JSX/TSX transpilation, optional JSON data
injection, module evaluation, IR decoding, and the complete component semantic
validation pass without invoking the DOCX backend or writing an output file.
It prints `valid: INPUT` and exits with status 0 on success. Invalid DSL uses
the standard detailed diagnostic and exits with status 1.

`pdf` converts an existing `.docx` file to PDF by calling a local office
application. The compiler does not bundle a layout engine. `--engine` is
`auto` (default), `libreoffice`, `word`, or `wps`.

- `libreoffice` uses `soffice --headless --convert-to pdf`. Layout follows
  LibreOffice Writer.
- `word` uses Microsoft Word (`Word.Application` on Windows, AppleScript on
  macOS). Layout follows Word. Linux has no Word automation path.
- `wps` uses WPS Writer (`KWPS.Application` / `wps.Application` on Windows).
  WPS on Linux/macOS can be detected but has no supported headless PDF export.

`auto` prefers Word, then WPS, then LibreOffice on Windows and macOS, and
LibreOffice then WPS on Linux. Explicit `--soffice` or `DOCX_JSX_SOFFICE`
selects LibreOffice. `DOCX_JSX_WORD` and `DOCX_JSX_WPS` pin those binaries
when present. `DOCX_JSX_PDF_ENGINE` sets the default engine. When `-o` is
omitted the input extension is replaced with `.pdf`. Existing outputs are
rejected unless `--force` is present. If no usable converter is found the
command fails with install and configuration instructions instead of writing
a PDF. A successful run prints the output path.

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
| `Section` | `Header`, `Footer`, `Heading`, `Paragraph`, `Caption`, `Index`, `TableOfContents`, `TableOfFigures`, `TableOfEntries`, `Bookmark`, `Table`, `List`, `ContentControl` |
| `Header`, `Footer` | `Paragraph`, `Caption`, `Table`, `List`, block `ContentControl` |
| `Paragraph` | `Run`, `Text`, `Break`, `CarriageReturn`, `NonBreakingSpace`, `SoftHyphen`, `NonBreakingHyphen`, `Image`, `Hyperlink`, `ContentControl`, `Field`, `DateField`, `TimeField`, `FileNameField`, `AuthorField`, `TitleField`, `SubjectField`, `SequenceField`, `ReferenceField`, `MergeField`, `DocumentPropertyField`, `FormulaField`, `IndexEntry`, `Comment`, `InlineBookmark`, `Inserted`, `Deleted`, `MovedFrom`, `MovedTo`, `Footnote`, `Tab`, `TabStop`, `PositionalTab`, `Symbol`, `Bold`, `Italic`, `Underline`, `StrikeThrough`, `Superscript`, `Subscript`, `AllCaps`, `HiddenText`, `SpecialHiddenText`, `DoubleStrike`, `SpacedText`, `ScaledText`, `FitText`, `BorderedText`, `ShadedText`, `PageNumber`, `TotalPages`, `PageReference`, `TocEntry`, string, number |
| `Heading` | same inline children as `Paragraph` |
| `Caption` | same inline children as `Paragraph` |
| `Run` | `Text`, `Break`, `CarriageReturn`, `NonBreakingSpace`, `SoftHyphen`, `NonBreakingHyphen`, `Image`, `Footnote`, `Tab`, `Symbol`, `DeletedInstruction` (only inside `Deleted`), string, number |
| `Comment` | inline children selected by the comment |
| `InlineBookmark` | same inline children as `Paragraph` |
| `Footnote` | `Run`, `Text`, `Break`, `Image`, `Tab`, `Symbol`, string, number |
| `Inserted` | run-level inline content |
| `Deleted` | `Run`, `Text`, `Comment`, string, number |
| `DeletedInstruction` | none |
| `MovedFrom`, `MovedTo` | run-level inline content |
| `Hyperlink` | `Run`, `Text`, `Break`, `Image`, `ContentControl`, `Inserted`, `Deleted`, `InlineBookmark`, `Comment`, string, number |
| `ContentControl` | inline: `Run`, `Text`, `Break`, `CarriageReturn`, `Image`, `Footnote`, `Tab`, `PositionalTab`, `Symbol`, `PageReference`, string, number. block (`Section`, `Header`, or `Footer` child): `Paragraph`, `Heading`, `Caption`, `Table`, `List` |
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
| `TableOfContents`, `TableOfFigures`, `TableOfEntries` | `TocItem`, `TocBefore`, `TocAfter` |
| `TocItem` | none |
| `TocBefore`, `TocAfter` | `Paragraph`, `Table` |
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
  External reverse restores `height`, `heightRule`, `cantSplit`, and row
  revision metadata without the IR manifest. `gridBefore` and `gridAfter` are
  non-negative integer grid-unit counts. `widthBefore` and `widthAfter` are
  non-negative point values, serialized as `dxa` twips. External reverse
  restores these properties from `w:trPr`; non-`dxa` widths are rejected with
  a conversion suggestion.
- Paragraphs, headings, list items, and paragraph styles may use integer character-unit
  indentation properties: `indentLeftChars`, `indentStartChars`,
  `indentRightChars`, `indentEndChars`, `firstLineChars`, and `hangingChars`.
  Values map directly to the corresponding OOXML `w:*Chars` values without
  point conversion. The `leftChars`/`startChars`, `rightChars`/`endChars`,
  `firstLine`/`firstLineChars`, and `hanging`/`hangingChars` pairs are
  mutually exclusive; first-line and hanging modes are also mutually
  exclusive. External reverse preserves the exact OOXML attribute name and
  integer value, and reports malformed or conflicting values with a repair
  suggestion.
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
instead of being silently discarded. Bookmark, comment, and hyperlink range
markers are reconstructed into nested `Bookmark` / `InlineBookmark` /
`Comment` JSX by matching IDs. Overlapping or unmatched range markers return
a reason and a repair suggestion instead of empty adjacent markers.
Wordprocessing text boxes (`w:drawing` with `wps:wsp`) are rejected with an
explanation and repair suggestion: docx-rs 0.4.22 has reader models for them
but its `Drawing` writer panics for text boxes, so flattening their paragraphs
would lose geometry and prevent a 1:1 reverse/recompile round trip.
Body-level `w:sdt` that is a table of contents, figures, or entries field
reverses as `TableOfContents` / `TableOfFigures` / `TableOfEntries`. Other
body-level and table-cell structured tags reverse as `ContentControl`.
Header/footer block structured tags reverse as `ContentControl` in the same
container, preserving supported aliases and data bindings. The generated
module imports only the components it uses and default-exports
one `Document`. External reverse also restores section `margins` and
`documentGrid` from `w:sectPr`, paragraph `keepNext` / `keepLines` /
`outlineLevel` from `w:pPr`, and run `themeColor` / `themeShade` /
`themeTint` from `w:color` in `document.xml` and `styles.xml` because
docx-rs 0.4.22 serializes `Color` as a bare hex string. Raster
`w:drawing` pictures reverse as `Image`; media bytes are written beside
the JSX as `media/<filename>` and referenced by `src`. Header and footer
drawings resolve media through `word/_rels/headerN.xml.rels` and
`word/_rels/footerN.xml.rels` independently of `word/_rels/document.xml.rels`,
so a header `rId` is not merged with a body relationship that happens to
share the same id. Vector drawings
such as EMF, WMF, and PICT have no schema-valid writer path and error
instead of being dropped. A stock unused `Normal` style is omitted so
recompilation does not emit a duplicate `w:style` that the backend always
prepends. Reverse of a package that lacks comments, footnotes, or numbering
does not invent those components; recompilation then omits the empty parts
the 0.4.22 writer would otherwise inject. External reverse restores already
writable `Document` settings from `word/settings.xml`, `docProps/core.xml`,
and `docProps/custom.xml` when they differ from backend defaults, and
restores `customXmlItems` from `customXml/itemN.xml` plus the matching
`itemPropsN.xml`, including `ds:schemaRef` URIs as `schemaRefs`. A custom
XML part without an item id or whose payload is not well-formed XML is an
error instead of a silent drop.

- `Document`: `defaultFont?: string` is a shorthand that assigns the same font
  to all four physical font slots. `defaultFonts?: {ascii?,hiAnsi?,eastAsia?,cs?,asciiTheme?,hiAnsiTheme?,eastAsiaTheme?,csTheme?,hint?}`
  exposes every field of `docx_rs::RunFonts`; it must contain at least one
  non-empty string. `defaultFont` and `defaultFonts` are mutually exclusive.
  Reverse conversion collapses four equal physical slots to `defaultFont` only
  when no theme or hint is present; otherwise it emits `defaultFonts` without
  merging slots. Positive `defaultSize?: number`, signed
  `defaultCharacterSpacing?: number` (points), `createdAt?: string`,
  `updatedAt?: string`, `customProperties?: Record<string,string>`,
  `creator?: string`, `lastModifiedBy?: string`, `description?: string`,
  `language?: string`, `revision?: integer`, `subject?: string`,
  `title?: string`,
  `documentId?: string`, positive `defaultTabStop?: number` (points),
  `zoom?: integer` sets the Word document view zoom as `w:zoom/@w:percent`;
  it must be between 10 and 500. Reverse conversion preserves explicit zoom
  percentages without the IR manifest and omits the backend default of 100.
  Non-percent zoom modes (`fullPage`, `bestFit`, and `textFit`) cannot be
  represented and are rejected with repair guidance.
  `documentProtection?: {edit?,formatting?,enforcement?,cryptProviderType?,cryptAlgorithmClass?,cryptAlgorithmType?,cryptAlgorithmSid?,cryptSpinCount?,cryptProvider?,algIdExt?,algIdExtSource?,cryptProviderTypeExt?,cryptProviderTypeExtSource?,hash?,salt?,algorithmName?,hashValue?,saltValue?,spinCount?}`.
  `edit` is `none | readOnly | comments | trackedChanges | forms`; the two
  protection flags are booleans; cryptographic enums follow the OOXML values
  (`rsaFull | rsaAES`, `hash | encryption | custom`, `typeAny | custom`).
  Numeric cryptographic fields are non-negative 32-bit integers; string fields are
  non-empty; `algIdExt` and `cryptProviderTypeExt` are even-length hexadecimal
  data. `hash` and `salt` must be supplied together and are precomputed
  OOXML base64 values—plaintext passwords are not accepted or hashed by this
  tool. The Office 2010+ fields `algorithmName`, `hashValue`, `saltValue`, and
  `spinCount` are also preserved. `hashValue` and `saltValue` must be supplied
  together. Restrictions are enforced only when `enforcement` is true.
  This is an editing restriction, not encryption or a security boundary.
  Reverse preserves the supported `w:documentProtection` attributes without
  the IR manifest and rejects unsupported attributes rather than dropping them.
  `documentVariables?: Record<string,string>`, `evenAndOddHeaders?: boolean`,
  `trackRevisions?: boolean` enables/disables Word's document-wide tracking of
  future edits. The setting is an on/off `w:trackRevisions` element and is
  preserved in both directions, including explicit false values.
  `doNotTrackMoves?: boolean` and `doNotTrackFormatting?: boolean` configure
  Word's revision behavior for moves and formatting changes; both are on/off
  settings preserved from external DOCX, including explicit false values.
  `adjustLineHeightInTable?: boolean`, and
  `characterSpacingControl?: "doNotCompress" | "compressPunctuation" | "compressPunctuationAndJapaneseKana"`.
  Core-property strings, metadata strings, IDs, property names, and variable
  names must be non-empty; `revision` must be a positive integer. Core
  properties map to `docProps/core.xml` and reverse from external DOCX without
  the IR manifest. Empty or default values injected by docx-rs (`unknown` for
  creator/modifier and `1` for revision) are omitted on reverse when no
  meaningful distinction can be recovered.
  maps must not be empty. These values are emitted to the standard core,
  custom-property, styles, and settings package parts.
  `defaultLineSpacing?: {before?,after?,line?,beforeLines?,afterLines?,lineRule?}`
  configures default paragraph properties. `before`, `after`, and `line` are
  finite point values (`before` and `after` are non-negative);
  `beforeLines` and `afterLines` are non-negative integers in OOXML
  hundredths-of-a-line units; `lineRule` is `auto`, `atLeast`, or `exact`.
  The object must configure at least one field.
  `webExtensions?: WebExtension[]` adds Office task-pane web extensions. Each
  item requires non-empty `id`, `referenceId`, `version`, `store`, and
  `storeType`; optional `properties` is a string map. Extension IDs must be
  unique. Supplying the array creates the task-pane part and one relationship
  and `webextensionN.xml` part per item. External reverse reads the task-pane
  catalog and its relationships without the IR manifest, retaining extension
  IDs, metadata, and string properties. Missing or ambiguous relationships,
  malformed entries, and non-string properties are errors rather than silently
  dropping extensions.
  `customXmlItems?: {id:string,xml:string}[]` embeds parsed XML data-store
  items. IDs must be non-empty and unique, XML must be well-formed, and both
  arrays must be non-empty when present. The compiler emits item, properties,
  relationship, document relationship, and content-type parts for every item.
  Optional `schemaRefs?: string[]` records `ds:schemaRef/@ds:uri` values.
  External reverse reads those parts without the IR manifest. The `id` comes
  from `ds:itemID` with surrounding braces stripped. Missing `itemProps` for
  an item part errors with a repair suggestion. Settings reverse omits the
  backend default tab stop (`840` twips) and the epoch core timestamps the
  writer injects when `createdAt` / `updatedAt` are absent.
- `Document.styles?: StyleDefinition[]` defines reusable Word styles. Every
  definition requires unique non-empty `id`, non-empty `name`, and `type` of
  `paragraph`, `character`, `numbering`, or `table`. Optional metadata is
  `basedOn`, `next`, `link`, `quickFormat`, `uiPriority`, `semiHidden`, and
  `unhideWhenUsed`. `run?: {font,fonts,size,sizeCs,color,themeColor,themeShade,themeTint,highlight,bold,boldCs,italic,italicCs,strike,doubleStrike,caps,underline,hidden,verticalAlign,characterSpacing,fitText,stretch,specVanish,shading,textBorder}`
  uses the same units and color rules as `Run`. `strike` and
  `doubleStrike` cannot both be true; `verticalAlign` is `baseline`,
  `superscript`, or `subscript`; `characterSpacing` is a signed point value.
  False `bold`, `italic`, `strike`, `doubleStrike`, and `caps` values are
  retained as explicit off properties, rather than omitted from the style.
  `fitText:{width,id?}` uses a positive point width and optional unsigned
  32-bit integer id; `stretch` is an integer percentage from 1 through 600;
  `specVanish` only accepts `true` because docx-rs has no explicit-off writer
  value; `shading:{fill,color?,pattern?}` requires six-digit RGB `fill`, accepts
  optional six-digit RGB or `auto` for `color`, and uses the same OOXML pattern
  names as `ShadedText`. Reverse conversion reads these values from the style's
  direct `w:rPr` XML; unsupported foreground/fill forms and disabled
  `w:specVanish` fail with a repair suggestion.
  `paragraph?: {align,textAlign,snapToGrid,keepNext,keepLines,pageBreakBefore,widowControl,bidi,shading,suppressLineNumbers,spacingBefore,spacingAfter,lineSpacing,spacingBeforeLines,spacingAfterLines,lineRule,indentLeft,indentRight,firstLine,hanging,hangingChars,firstLineChars,outlineLevel,frame}`
  uses point dimensions; `firstLine` and `hanging` are mutually exclusive.
  `keepNext`, `keepLines`, `pageBreakBefore`, `widowControl`, `bidi`, and
  `suppressLineNumbers` are booleans and map to the style's corresponding
  paragraph properties. `shading` is a solid six-digit RGB fill. These values
  are restored from external `styles.xml`; unsupported patterned
  shading returns an error instead of being discarded. `adjustRightIndent`
  accepts only `0` or `1`, the schema-valid on/off values emitted by the pinned
  writer for `w:adjustRightInd`.
  `textAlign` is `auto`, `baseline`, `bottom`, `center`, or `top`; line-spacing
  fields use the same units and `lineRule` values as `Paragraph`.
  `frame` accepts `wrap`, `verticalAnchor`, `horizontalAnchor`, `heightRule`,
  `xAlign`, `yAlign`, `horizontalSpace`, `verticalSpace`, `x`, `y`, `width`,
  and `height`; its dimensions are points and coordinate/alignment pairs are
  mutually exclusive.
  Table styles additionally accept
  `table?: {indent,align,margins,border}` and
  `cell?: {width,colSpan,verticalAlign,verticalMerge,textDirection,shading,margins,border}`.
  These reuse the corresponding `Table` and `TableCell` property formats, and
  external reverse restores table-style table/cell properties without the IR
  manifest. Values not representable in these objects cause a repair-oriented
  reverse error rather than being omitted.
  Width, layout, and nested style references belong on `Table`; they are
  rejected inside a style because OOXML `StyleTableProperties` does not permit
  them even though docx-rs exposes setters that serialize invalid markup.
  `run.textBorder` reuses the
  `BorderedText` object shape. `table` and `cell` are rejected on non-table
  styles because docx-rs would otherwise omit them during serialization.
  `basedOn` is emitted as native Word style inheritance. A locally declared
  base must have the same type as the derived style, and inheritance cycles
  are rejected. `Paragraph`, `Heading`, `Caption`, and `Index` may reference a
  paragraph style; `Run` may reference a character style; `Table` may reference
  a table style. References to locally declared styles are type-checked.
  References not declared in `Document.styles` remain legal so built-in styles
  from Word templates can be used. Reverse conversion preserves style ids and
  definitions rather than expanding inherited properties onto every node.
- `Section`: `type?: "nextPage" | "nextColumn" | "continuous" | "evenPage" | "oddPage"`
  (`w:type`; omitted means Word's default `nextPage`), `pageSize?: "A4" | "Letter" | {width,height}`,
  where custom `width` and `height` are positive point values,
  `orientation?: "portrait" | "landscape"`, `margins?: {top,right,bottom,left,header?,footer?,gutter?}`,
  `titlePage?: boolean`, `textDirection?: "lrTb" | "tbRl" | "btLr" | "lrTbV" | "tbRlV"`,
  `documentGrid?: {type,linePitch?,charSpace?}`, and
  `pageNumbering?: {start?,chapterStyle?,format?,chapterSeparator?}`, `lineNumbering?: {countBy?,start?,distance?,restart?}`,
  and `columns?: 1..45 | {count, spacing?, separator?} | {widths, separator?}`.
  The numeric shorthand retains equal-width columns with the backend's default
  425-twip spacing. In `{count, spacing?, separator?}`, `count` is 1..45,
  `spacing` is a non-negative point measure (default 0), and `separator` is a
  boolean. In `{widths, separator?}`, `widths` contains 1..45 column objects;
  each accepts optional non-negative `width` and `spacing` point measures.
  `count` and `widths` are mutually exclusive. Reverse preserves equal and
  individual widths, spacing, and separator settings without the IR manifest.
  Document-grid `type` is `default`,
  `lines`, `linesAndChars`, or `snapToChars`; `linePitch` is a positive point
  measurement, while `charSpace` is a signed integer in OOXML 1/4096-em units.
  `pageNumbering.start` is a non-negative integer, `chapterStyle` is a
  non-empty Word chapter-style value, `format` maps to `w:pgNumType/@w:fmt`, and
  `chapterSeparator` maps to `w:pgNumType/@w:chapSep`. Supported formats are
  `decimal`, `upperRoman`, `lowerRoman`, `upperLetter`,
  `lowerLetter`, `ordinal`, `cardinalText`, `ordinalText`, `hex`, `chicago`,
  `ideographDigital`, `japaneseCounting`, `aiueo`, `iroha`, `decimalFullWidth`,
  `decimalHalfWidth`, `japaneseLegal`, `japaneseDigitalTenThousand`,
  `decimalEnclosedCircle`, `decimalFullWidth2`, `aiueoFullWidth`,
  `irohaFullWidth`, `decimalZero`, `bullet`, `ganada`, `chosung`,
  `decimalEnclosedFullstop`, `decimalEnclosedParen`, `decimalEnclosedCircleChinese`,
  `ideographEnclosedCircle`, `ideographTraditional`, `ideographZodiac`,
  `ideographZodiacTraditional`, `taiwaneseCounting`, `ideographLegalTraditional`,
  `taiwaneseCountingThousand`, `taiwaneseDigital`, `chineseCounting`,
  `chineseLegalSimplified`, `chineseCountingThousand`, `koreanDigital`,
  `koreanCounting`, `koreanLegal`, `koreanDigital2`, `vietnameseCounting`,
  `russianLower`, `russianUpper`, `numberInDash`, `hebrew1`, `hebrew2`,
  `arabicAlpha`, `arabicAbjad`, `hindiVowels`, `hindiConsonants`,
  `hindiNumbers`, `hindiCounting`, `thaiLetters`, `thaiNumbers`, or
  `thaiCounting`, `none`, `bahtText`, `dollarText`, and `custom`. Chapter
  separators are `hyphen`, `period`, `colon`, `emDash`, or `enDash`. The pinned
  `docx-rs` writer has no `fmt` or `chapSep` setter; the compiler emits these
  schema-valid attributes directly in `w:pgNumType`, and external reverse reads
  them from package XML. `documentGrid` and `pageNumbering`
  objects must be non-empty. `lineNumbering` enables section line numbering; its `countBy?`
  and `start?` are positive integers (defaults 1), `distance?` is a
  non-negative point measurement converted to OOXML twips (20 twips per point;
  omitted for Word's automatic placement),
  and `restart?` is `newPage` (default), `newSection`, or `continuous`.
  An empty object enables default line numbering. External reverse preserves
  the presence of `w:lnNumType` and every explicitly present attribute.
- `Paragraph`: `align?: "left" | "center" | "right" | "both" | "distribute" | "start" | "end" | "justified"`,
  `spacingBefore?`, `spacingAfter?`, `lineSpacing?`, signed
  `indentLeft?` / `indentRight?`, `firstLine?`, `hanging?`, `keepNext?`, `keepLines?`,
  `pageBreakBefore?`. `firstLine` and `hanging` are mutually exclusive.
  Optional `spacingBeforeLines?`, `spacingAfterLines?`, and
  `lineRule?: "auto" | "atLeast" | "exact"` expose the remaining docx-rs
  line-spacing controls. The `*Lines` values are non-negative integers in
  OOXML hundredths-of-a-line units and may be combined with point spacing.
- `Paragraph`: also accepts `snapToGrid?`, `widowControl?`, `bidi?`,
  `contextualSpacing?: boolean` preserves `w:contextualSpacing` (including an
  explicit false value) on `Paragraph`, `Heading`, and `ListItem`; the inherited
  form is `Document.styles[].paragraph.contextualSpacing?: boolean`.
  `mirrorIndents?: boolean` writes `w:mirrorIndents` to interpret left/right
  indents as inside/outside indents on mirrored pages. `suppressOverlap?: boolean`
  writes `w:suppressOverlap` for framed paragraphs to prevent overlapping text
  frames. Both preserve explicit false and are supported on `Paragraph`,
  `Heading`, `ListItem`, and `Document.styles[].paragraph`.
  `suppressLineNumbers?: boolean` suppresses active section line numbers for this
  paragraph; it is independent of section `lineNumbering` and preserves explicit
  `false` when reversing external DOCX. It is supported on `Paragraph`, `Heading`,
  and `ListItem`. `Document.styles[].paragraph.suppressLineNumbers?: boolean`
  provides the inherited style form.
  `textAlign?: "auto" | "baseline" | "bottom" | "center" | "top"`,
  `adjustRightIndent?: 0 | 1`, six-digit RGB `shading?`, and integer
  `outlineLevel?: 0..9`. `frame?` has the same shape, units, enumerations, and
  coordinate/alignment conflicts as `Document.styles[].paragraph.frame`.
  Optional `inserted?` and `deleted?` objects mark the paragraph mark as
  inserted or deleted (OOXML `w:pPr/w:rPr/w:ins` or `w:del`) and are mutually
  exclusive; both accept non-empty `author?` and `date?` strings.
  `propertyChange?` records previous paragraph
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
  `font?` or `fonts?`, positive `size?`, `bold?`, `italic?`, six-digit RGB `color?`, and
  signed `characterSpacing?` in points.
- `Heading`: required `level: 1..9`, optional `style`, plus all `Paragraph`
  layout properties except `outlineLevel`. Its `level` emits both a heading
  style and the corresponding Word outline level.
- `Run`: `themeColor?` accepts Word theme tokens `dark1`, `light1`, `dark2`,
  `light2`, `accent1` through `accent6`, `hyperlink`, `followedHyperlink`,
  `none`, `background1`, `text1`, `background2`, or `text2`. `themeShade?` and
  `themeTint?` are two-digit hex modifiers. Boolean `bold`, `italic`, `strike`,
  `doubleStrike`, and `caps` emit explicit on/off OOXML; enabled strike modes are
  mutually exclusive. `hidden?` and `specVanish?` accept only `true` and emit
  `w:vanish` / `w:specVanish`; docx-rs has no explicit-off writer setter.
  External reverse rejects explicit off values because docx-rs's reader treats
  either element's presence as on. `verticalAlign?` accepts `baseline`, `superscript`, or
  `subscript` and maps to `w:vertAlign`; the `Superscript` and `Subscript`
  components remain composable shorthand. `RunProperty::insert` and `RunProperty::delete` are not
  exposed as JSX properties: docx-rs 0.4.22 writes `w:ins` / `w:del` inside
  `w:rPr`, which the Open XML SDK identifies as invalid there. Use the
  `Inserted`/`Deleted` components for tracked text revisions.
- `Caption`: required Word identifier `label`; optional Word identifier
  `identifier?` defaults to `label`. `format?`, `restart?`, `placeholder?`, and
  `dirty?` have the same meaning as on `SequenceField`; `style?` defaults to
  `Caption`, `numberSeparator?` defaults to one space, and `textSeparator?`
  defaults to `: `. It emits a styled paragraph containing the visible label,
  a native `SEQ` field, the separator, and its inline children.
- `Run`: `font?`, `fonts?`, `size?`, `bold?`, `italic?`, `strike?`,
  `doubleStrike?`, `caps?`, `hidden?: true`, `specVanish?: true`,
  `underline?: boolean | UnderlineStyle`,
  `sizeCs?`, `boldCs?`, `italicCs?`,
  `color?`, `highlight?`, `style?`, and `verticalAlign?: "baseline" |
  "superscript" | "subscript"`.
  `sizeCs?` is the complex-script font size in points, converted to half-points
  as `w:szCs`; it is independent from `size?`. `boldCs?` and `italicCs?` are
  explicit on/off values for complex-script characters only (`w:bCs` / `w:iCs`)
  and are independent from `bold?` / `italic?`. External reverse preserves
  these values independently rather than merging them into the non-complex
  script properties. `PositionalTab` is represented by a run child component,
  not by `Run` or `StyleDefinition.run` properties; external `w:ptab` inside
  `w:rPr` is schema-invalid and reverse reports it instead of dropping it.
  `true` is the compatibility shorthand for `"single"`; `false` means
  `"none"`. String values preserve the exact `w:u/@w:val` token. Supported
  `UnderlineStyle` values are `single`, `words`, `double`, `thick`, `dotted`,
  `dottedHeavy`, `dash`, `dashedHeavy`, `dashLong`, `dashLongHeavy`,
  `dotDash`, `dashDotHeavy`, `dotDotDash`, `dashDotDotHeavy`, `wave`,
  `wavyHeavy`, `wavyDouble`, and `none`. External reverse always emits the
  string form so underline style is not collapsed to a boolean.
- `fonts` on `Run`, `Paragraph`, `Heading`, and `StyleDefinition.run` has the
  same nine optional fields as `Document.defaultFonts`. It is mutually
  exclusive with `font`, must not be empty, and preserves distinct physical,
  theme, and hint values during external-DOCX reverse conversion.
- `Paragraph`: also accepts `style?: string` for custom or built-in styles.
- `Text`: optional `value`; it may alternatively contain string/number children.
- `Break`: `type?: "line" | "page" | "column"` (default `line`).
- `CarriageReturn`: no properties or children; emits a run-level `w:cr`.
- `NonBreakingSpace`, `SoftHyphen`, `NonBreakingHyphen`: no properties or
  children. They emit Unicode U+00A0, U+00AD, and U+2011 respectively, keeping
  adjacent words together, exposing an optional hyphenation point, or keeping
  a hyphenated term together. They are valid wherever run-level text is valid.
- `Image`: required `src`, `width`, and `height` in points. Optional
  `description` supplies accessible alternative text and maps to
  `wp:docPr/@descr`; `title` maps to `wp:docPr/@title`. Both must be non-empty
  strings when supplied and are preserved by external DOCX reverse without
  relying on the embedded IR manifest. `decorative?: boolean` marks an image
  as decorative by writing an explicitly empty `wp:docPr/@descr`; it cannot be
  combined with `description` or `title`. External reverse maps an empty
  `wp:docPr/@descr` to `decorative`. `relationshipId` overrides the generated
  image relationship identifier and must be non-empty. `rotate` is an integer
  angle from 0 through 65535 degrees.
  `floating` switches from inline layout to an anchored drawing. Anchored
  images support `allowOverlap`, `positionH` (`left`, `center`, `right`, or a
  point offset), `positionV` (`top`, `center`, `bottom`, or a point offset),
  `relativeFromH` (`character`, `column`, `insideMargin`, `leftMargin`,
  `margin`, `outsideMargin`, `page`, or `rightMargin`), `relativeFromV`
  (`bottomMargin`, `insideMargin`, `line`, `margin`, `outsideMargin`, `page`,
  `paragraph`, or `topMargin`), point-valued `distanceTop`, `distanceBottom`,
  `distanceLeft`, and `distanceRight`, integer `relativeHeight`, and
  `layoutInCell?: boolean` to control whether a floating drawing moves with its
  containing table cell. `simplePosition?: { x: integer, y: integer }` selects
  the DrawingML simple-position anchor mode and supplies signed EMU coordinates
  in the i32 range [-2,147,483,648, 2,147,483,647];
  it is mutually independent from the positional offsets/alignments (which are
  preserved even though simple positioning takes precedence in Word). All
  anchor-only properties, including `simplePosition`, require `floating={true}`.
  Omitting `layoutInCell` keeps the backend default;
  external reverse emits the reader's effective boolean so recompilation keeps
  the same layout behavior.
  Image sources resolve from the entry module directory. External reverse
  extracts raster media from the package, restores `wp:docPr/@descr` and
  `wp:docPr/@title`, and points `src` at
  `media/<filename>` next to the emitted JSX, including pictures whose
  relationships live on a header or footer part. `docx-jsx reverse` writes
  those sidecar files beside the JSX output. The pinned writer hardcodes
  `simplePos="0"` and writes its `simple_pos` field into `allowOverlap`; the
  compiler repairs both anchor attributes in the package. External reverse
  preserves the exact `simplePos` flag and signed coordinates without relying
  on the embedded IR manifest.
- `Table`: `width?`, `widthPercent?`, `align?`, `layout?: "auto" | "fixed"`,
  `columnWidths?: number[]`, `border?: {style?,size?,color?}`. Point and percent
  widths are mutually exclusive. `indent?`, four-edge `margins?`, `style?`,
  and floating `position?: {leftFromText,rightFromText,verticalAnchor,
  horizontalAnchor,xAlign,yAlign,x,y}` are also supported. Position distances
  use points; anchors/alignment use OOXML string values. External reverse
  restores table properties and grid column widths without the IR manifest;
  unsupported width units, invalid grid values, and unrepresentable values fail
  with repair guidance.
- `TableRow`: optional `height` in points (`1 pt = 20 twips`),
  `heightRule?: "auto" | "atLeast" | "exact"`, `cantSplit?: boolean`, and
  `header?: boolean`. `header={true}` emits `w:tblHeader` so Word repeats that
  row as a table header across page breaks; mark the leading contiguous header
  rows for predictable Word behavior. `header={false}` is preserved as an
  explicit disabled value in OOXML. External reverse restores this property
  without relying on the IR manifest. The default height rule is supplied by
  Word when omitted. Height is converted to/from whole twips; an external row
  height that is not a whole twip cannot be represented and causes reverse to
  fail with a repair suggestion.
- `TableCell`: mutually exclusive `width?` in points or `widthPercent?` greater
  than 0 and at most 100, `colSpan?`,
  `verticalAlign?: "top" | "center" | "bottom"`, `verticalMerge?`,
  `textDirection?`, `noWrap?: boolean`, `margins?`, `shading?`, and `border?`.
  `noWrap={true}` prevents text wrapping in the cell; explicit false is retained
  as an off value. This controls cell text wrapping only and is distinct from
  paragraph `fitText`. External reverse
  restores these properties from DOCX cells without the IR manifest. A cell
  width that uses an unsupported unit, margins with missing edges or non-`dxa`
  units, nonzero border spacing, and patterned shading are reported instead of
  silently dropped. The compiler emits cell properties in OOXML `CT_TcPr` order,
  including when `verticalAlign` and `shading` are both present.
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
  `anchor?: string` (document bookmark), plus `history?: boolean`. Reverse
  preserves `w:history="0"` as `history={false}` and `w:history="1"` as the
  true shorthand; other values are rejected with a repair suggestion. In addition
  to ordinary run-level children it accepts `ContentControl`, `Inserted`,
  `Deleted`, `InlineBookmark`, and `Comment`, mapping to the corresponding
  public `docx_rs::Hyperlink::add_*` APIs. Hyperlinks cannot be nested.
  External reverse restores those children from OOXML (including `w:sdt` that
  the 0.4.22 hyperlink reader would otherwise flatten) and resolves `href`
  from the document relationship, not from the writer-only `path` field.
- `InlineBookmark`: required non-empty `name`. Its inline children are
  enclosed by a matching `w:bookmarkStart`/`w:bookmarkEnd` pair inside a
  paragraph or hyperlink. Reverse reconstructs the element by matching IDs
  rather than emitting an empty marker beside the marked runs.
- `PageNumber`, `TotalPages`: no properties. They emit dynamic Word fields.
- `Comment`: required non-empty `text` plus selected inline children;
  `author?: string`, `date?: string`, `id?: non-negative integer`, and
  `parentCommentId?: non-negative integer`, `done?: boolean`. Omitted IDs are
  assigned the next unused positive comment ID; explicit IDs are preserved.
  IDs are unique document-wide. A reply's `parentCommentId` must reference
  another comment with an explicit `id`, and reply cycles are invalid. It emits
  a comment range/reference in `comments.xml` and the parent relationship in
  `commentsExtended.xml`; `done` maps to the comment's resolved state in
  `commentsExtended.xml`. Reverse preserves IDs, reply relationships, and
  resolved state from external DOCX without relying on the IR manifest.
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
  `Underline` requires text content and accepts `type?: UnderlineStyle`
  (default `single`), emitting native `w:u`.
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
  is serialized as `w:delText`, preserving optional `Run` formatting. A
  `DeletedInstruction` is only accepted inside a `Run` that is nested in
  `Deleted`. It serializes as `w:delInstrText` and preserves the exact field
  instruction. `Comment` inside `Deleted` creates a matching comment range inside `w:del`,
  wraps the selected deleted text, and preserves its annotation in `comments.xml`.
- `DeletedInstruction`: required `instruction: string`, legal only inside a
  `Run` nested within `Deleted`. It emits `w:delInstrText` inside the tracked deletion. External
  reverse preserves this element separately from ordinary `w:delText`; the
  pinned docx-rs reader does not parse `w:delInstrText`, so reverse normalizes it
  only in a temporary in-memory package and retains its exact instruction text.
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
  `storeItemId` strings. It emits a native Word structured document tag
  (`w:sdt`). When any data-binding property is used, non-empty `xpath` and
  `storeItemId` are required; the latter is the custom XML data-store item GUID
  required by OOXML. The values are emitted as `w:dataBinding` attributes. As a paragraph,
  hyperlink, or table-cell child it requires inline content. As a `Section`,
  `Header`, or `Footer` child it maps to docx-rs `StructuredDataTag` block
  children (`Paragraph`, `Heading`, `Caption`, `Table`, `List`); inline and
  block children cannot be mixed. External reverse restores a body-level
  `w:sdt` as a `ContentControl` wrapping those block children in the original
  `Section`, `Header`, or `Footer` container.
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
- `List`: `type?: "bullet" | "ordered"` (default `bullet`), optional
  positive integer `start`, and optional `multiLevelType?: "singleLevel" |
  "multilevel" | "hybridMultilevel"` (`w:multiLevelType`; omitted by default).
  Optional non-empty `styleLink?` (`w:styleLink`) declares that this abstract
  numbering is the base definition for the named numbering style;
  `numStyleLink?` (`w:numStyleLink`) references a numbering style supplying the
  actual numbering properties. When either link is present and `levels` is
  omitted, no default `w:lvl` children are generated. `numStyleLink` is
  mutually exclusive with `type`, `start`, and explicit `levels`, because its
  abstract numbering definition must not carry the referenced style's actual
  level properties. `multiLevelType` and per-instance `overrides` may still be
  used. External reverse preserves both link values and their level structure.
  External reverse also preserves an explicit `multiLevelType`. Optional
  `levels` is a non-empty array of level
  objects, one per `w:lvl` in document order (`ilvl` is the array index,
  `0..8`). When `levels` is omitted and neither style link is present, the compiler emits nine default levels
  (`decimal` + `%N.` for `ordered`, `bullet` + `•` otherwise) with
  `(level+1)*36` pt left indent and `18` pt hanging indent. `levels` and the
  `type`/`start` shorthand may be combined: omitted level fields inherit the
  shorthand defaults for that list. Each level accepts `format?` (non-empty
  Word `ST_NumberFormat` token such as `decimal`, `bullet`, `lowerLetter`,
  `upperRoman`), `text?` (non-empty `w:lvlText`), positive `start?`,
  `align?: "left" | "center" | "right"` (default `left`),
  `suffix?: "tab" | "space" | "nothing"` (default `tab`, omitted from OOXML),
  non-empty `paragraphStyle?`, `restart?` (`w:lvlRestart`, non-negative
  integer), `legal?` (`w:isLgl`), point `indentLeft?` / `indentRight?` /
  `hanging?` / `firstLine?` (`firstLine` and `hanging` are mutually
  exclusive), and the same run formatting as `Run` (`font`/`fonts`, `size`,
  `color`, `highlight`, `bold`, `italic`, `strike`, `doubleStrike`,
  `underline` as a Word underline style string, `boldCs`, `italicCs`, `sizeCs`, `hidden`,
  `characterSpacing`). Boolean run flags emit explicit on/off OOXML.
  `overrides?` is an array of
  per-instance `w:lvlOverride` objects: each requires a unique `level: 0..8`,
  may contain positive `start?` (`w:startOverride`) and/or `definition?` (a
  list-level object emitted as replacement `w:lvl`; its `ilvl` is the enclosing
  override's `level`, and omitted fields use the same defaults as `levels`). An
  override must set `start` or `definition`;
  duplicate levels are invalid. External reverse preserves start overrides
  and full level replacements as `List.overrides`, independently of the
  shared abstract list definition. Contiguous paragraphs that share a `w:numId`
  reverse as one `List`; a missing numbering definition is an error rather
  than a plain `Paragraph`.
- `ListItem`: `level?: 0..8` (default `0`) and optional `style` for the
  numbered paragraph. Each list compiles to native Word numbering
  definitions and numbered paragraphs.
- `Bookmark`: required non-empty `name`. Its structural children are enclosed
  by a unique Word bookmark pair and may be targeted by `Hyperlink anchor`.
  Reverse reconstructs the element by matching start/end IDs around the
  enclosed paragraphs and tables; orphan or overlapping markers are errors.
- `TableOfContents`: `startLevel?: 1..9`, `endLevel?: 1..9`,
  `tcFieldLevelRange?: {start:1..9,end:1..9}`,
  `omitPageNumbersLevelRange?: {start:1..9,end:1..9}`,
  `entryBookmarkName?: string`, `entryAndPageNumberSeparator?: string`,
  `sequenceAndPageNumbersSeparator?: string`,
  `seqFieldIdentifierForPrefix?: string`,
  `stylesWithLevels?: [{style:string,level:1..9}]`,
  `preserveTab?: boolean`, `preserveNewLine?: boolean`,
  `useAppliedParagraphLineLevel?: boolean`,
  `hideTabAndPageNumbersInWebview?: boolean`, `hyperlinks?: boolean`,
  `dirty?: boolean`, `alias?: string`, `withoutSdt?: boolean`, and
  `deletion?: {author:string,date:string}`, and optional `fieldParagraph`
  formatting. Defaults to
  levels 1–3, hyperlinks enabled, and automatic refresh requested. A section
  accepts at most one TOC, placed before its body content (headers and footers
  may precede it. Optional `TocItem` children provide a static entry row with
  required non-empty `text` and `bookmark`, optional `level?: 1..9` (default
  `1`), optional `pageRef` display text (default `"1"`), and optional positive
  `tabPosition` in points (default `420`, maximum `1584`, corresponding to the
  OOXML maximum of `31680` twips). The compiler emits
  a native `w:hyperlink` entry and a `PAGEREF` field targeting `bookmark`;
  reverse restores the item from the external DOCX without the IR manifest.
  Optional `TocBefore` and `TocAfter` wrappers map to docx-rs
  `add_before_paragraph` / `add_before_table` and their `add_after_*` variants;
  each wrapper accepts one or more `Paragraph` or `Table` children and can
  appear at most once. They preserve ordered content around the generated
  field/static entries and are restored from external DOCX without the IR
  manifest. The blank paragraph docx-rs inserts before a first after-table is
  treated as a writer artifact. `deletion` maps to docx-rs `delete(author,
  date)` and marks the generated field instruction as a tracked deletion. It
  is valid only without index children and is mutually exclusive with
  `withoutSdt`; static entries are rejected because docx-rs silently ignores
  `delete` in its static-items serialization branch. Reverse restores the
  deletion author and date from the external DOCX. Unlike `withoutSdt`,
  deletion is not restricted to direct section children where the index
  component is otherwise allowed. `fieldParagraph?: object` applies
  formatting to the generated automatic-field paragraph. It accepts the
  existing `Paragraph` formatting names and units: `style`, `align`,
  `spacingBefore`, `spacingAfter`, `lineSpacing`, `spacingBeforeLines`,
  `spacingAfterLines`, `lineRule`, `indentLeft`, `indentRight`, `firstLine`,
  `hanging`, `keepNext`, `keepLines`, `pageBreakBefore`, `snapToGrid`,
  `widowControl`, `bidi`, `textAlign`, `shading`, `outlineLevel`, and
  `suppressLineNumbers`, `indentLeftChars`, `indentStartChars`,
  `indentRightChars`, `indentEndChars`, `firstLineChars`, and
  `hangingChars`, plus paragraph run defaults `font`, `fonts`, `size`, `bold`,
  `italic`, `color`, and `characterSpacing`. Point-valued spacing, indentation,
  font size, and character spacing use points. `font` sets all four physical
  font slots; `fonts` can configure the nine physical/theme/hint slots
  independently. `bold` and `italic` preserve both explicit true and false
  values. It is valid only without static
  `TocItem` rows because docx-rs ignores paragraph properties in its static-
  items branch. Character-unit indentation values are integers mapped
  directly to the corresponding OOXML `w:*Chars` attributes without point
  conversion; aliases and first-line/hanging modes follow the paragraph
  indentation exclusivity rules. `false` for `keepNext`,
  `keepLines`, or `pageBreakBefore`, and properties not listed here are
  rejected; reverse restores the listed values (including unwrapped indexes)
  from external DOCX and errors on unsupported properties rather than
  silently dropping them. The pinned `docx-rs` version has no paragraph writer
  setter for `suppressLineNumbers`; the compiler emits the equivalent
  schema-valid `w:suppressLineNumbers` element and reverse reads it from the
  field paragraph XML.
  `withoutSdt?: boolean` omits the outer structured document
  tag and leaves only the native field paragraphs. It is valid only when the
  index is a direct `Section` child and has no `TocItem`, `TocBefore`, or
  `TocAfter` children; reverse recognizes the adjacent field-start and
  field-end paragraphs without relying on a manifest.
- `TableOfFigures`: required non-empty `label` (for example `Figure`),
  `includeLabelAndNumber?: boolean`, `separator?: string`,
  `entryBookmarkName?: string`, `entryAndPageNumberSeparator?: string`,
  `sequenceAndPageNumbersSeparator?: string`,
  `seqFieldIdentifierForPrefix?: string`,
  `stylesWithLevels?: [{style:string,level:1..9}]`,
  `preserveTab?: boolean`, `preserveNewLine?: boolean`,
  `useAppliedParagraphLineLevel?: boolean`,
  `hideTabAndPageNumbersInWebview?: boolean`, `hyperlinks?: boolean`,
  `dirty?: boolean`, `alias?: string`, `withoutSdt?: boolean`, and
  `deletion?: {author:string,date:string}`, and optional `fieldParagraph`
  formatting with the same contract as `TableOfContents`. It emits a native
  TOC field filtered
  by caption/SEQ label. It also accepts `TocItem` children with the same
  contract as `TableOfContents`. `withoutSdt` has the same direct `Section`
  child and no-child restrictions as `TableOfContents`; `deletion` has the
  same no-child and mutual-exclusion restrictions and is allowed in any
  otherwise-valid index placement.
- `TableOfEntries`: required non-empty `identifier`, plus `hyperlinks?`,
  `tcFieldLevelRange?: {start:1..9,end:1..9}`,
  `omitPageNumbersLevelRange?: {start:1..9,end:1..9}`,
  `entryBookmarkName?: string`, `entryAndPageNumberSeparator?: string`,
  `sequenceAndPageNumbersSeparator?: string`,
  `seqFieldIdentifierForPrefix?: string`,
  `stylesWithLevels?: [{style:string,level:1..9}]`,
  `preserveTab?: boolean`, `preserveNewLine?: boolean`,
  `useAppliedParagraphLineLevel?: boolean`,
  `hideTabAndPageNumbersInWebview?: boolean`, `dirty?`, `alias?`, and
  `withoutSdt?: boolean`, `deletion?: {author:string,date:string}`, and
  optional `fieldParagraph` formatting with the same contract as
  `TableOfContents`. It emits
  a native custom TOC containing `TocEntry` fields with the same `identifier`.
  It also accepts `TocItem` children with the same contract as
  `TableOfContents`. `withoutSdt` uses the same direct `Section` child and
  no-child restrictions as `TableOfContents`; `deletion` uses the same
  no-child and mutual-exclusion restrictions and is allowed in any
  otherwise-valid index placement.
  Range starts must be no greater than ends; levels are integers from 1 to 9.
  Style names are non-empty and cannot contain commas or quotes, matching the
  OOXML `\\t` field-argument syntax. The Boolean switches default to false at
  the field level except component defaults already stated above. Reverse
  preserves supported `TOC` instruction switches without the IR manifest and
  errors on malformed or unsupported switch arguments. Text passed as an OOXML
  quoted field argument cannot contain a quote; validation rejects such values
  rather than emitting an ambiguous field instruction.

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
