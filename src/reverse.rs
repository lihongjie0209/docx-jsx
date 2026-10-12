//! DOCX to deterministic, recompilable JSX conversion.

use std::collections::{BTreeMap, BTreeSet, HashMap, VecDeque};
use std::fmt::Write as _;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use docx_rs::{
    AbstractNumbering, Comment, CommentChild, Delete, DeleteChild, DocumentChild, DrawingData,
    DrawingPosition, DrawingPositionType, FooterChild, HeaderChild, Hyperlink, HyperlinkData,
    Insert, InsertChild, Level, MoveFrom, MoveFromChild, MoveTo, MoveToChild, Numbering,
    ParagraphChild, Pic, PicAlign, RelativeFromHType, RelativeFromVType, RunChild,
    StructuredDataTag, StructuredDataTagChild, TableCellContent, TableChild, TableRowChild,
};
use num_traits::ToPrimitive;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use serde::Serialize;
use serde_json::{Map, Number, Value};

use crate::ir::{
    Child, IrEnvelope, Node, PAGE_NUMBER_CHAPTER_SEPARATORS, PAGE_NUMBER_FORMATS, SHADING_PATTERNS,
};
use crate::{Error, Result};

/// JSX module plus sidecar files produced by external DOCX reverse conversion.
#[derive(Debug, Default)]
pub struct ReversedDocument {
    /// Deterministic JSX module text.
    pub jsx: String,
    /// Relative paths (POSIX) to extracted package assets, keyed from the JSX
    /// file directory. Raster images are stored under `media/`.
    pub assets: BTreeMap<String, Vec<u8>>,
}

/// Converts DOCX package bytes into a deterministic JSX module.
///
/// # Errors
///
/// Returns an error when the package is invalid or contains a structure that
/// the JSX component model cannot currently represent without data loss.
pub fn reverse_document(bytes: &[u8]) -> Result<String> {
    Ok(reverse_package(bytes)?.jsx)
}

/// Converts DOCX package bytes into JSX and any extracted sidecar assets.
///
/// # Errors
///
/// Returns an error when the package is invalid or contains a structure that
/// the JSX component model cannot currently represent without data loss.
pub fn reverse_package(bytes: &[u8]) -> Result<ReversedDocument> {
    if let Some(ir) = read_ir_manifest(bytes)? {
        ir.validate()?;
        return Ok(ReversedDocument {
            jsx: ir_to_jsx(&ir)?,
            assets: BTreeMap::new(),
        });
    }
    validate_external_xml_features(bytes)?;
    let style_metadata = read_style_metadata(bytes)?;
    let theme_colors = parse_theme_colors(bytes)?;
    let section_columns = collect_section_columns(bytes)?;
    let section_types = collect_section_types(bytes)?;
    let section_line_numbering = collect_section_line_numbering(bytes)?;
    let section_page_numbering = collect_section_page_numbering(bytes)?;
    let numbering_multi_level_types = collect_numbering_multi_level_types(bytes)?;
    let table_row_xml = collect_table_row_xml_metadata(bytes)?;
    let table_xml = collect_table_xml_metadata(bytes)?;
    let paragraph_indent_xml = collect_paragraph_indent_metadata(bytes)?;
    let style_character_indents = collect_style_character_indent_metadata(bytes)?;
    let style_suppress_line_numbers = collect_style_line_number_suppression_metadata(bytes)?;
    let style_paragraph_properties = collect_style_paragraph_xml_metadata(bytes)?;
    let style_run_properties = collect_style_run_xml_metadata(bytes)?;
    let index_field_paragraph_xml = collect_index_field_paragraph_metadata(bytes)?;
    let hyperlink_histories = collect_hyperlink_histories(bytes)?;
    let toc_reader_bytes = normalize_nested_toc_item_sdts(bytes)?;
    let (reader_bytes, deleted_instructions) =
        normalize_deleted_instruction_parts(&toc_reader_bytes)?;
    let docx =
        docx_rs::read_docx(&reader_bytes).map_err(|error| Error::Reverse(error.to_string()))?;
    let mut images = collect_package_images(bytes)?;
    for (id, asset) in collect_images(&docx, &images.used_names) {
        images.by_id.entry(id).or_insert(asset);
    }
    let mut writer = Writer {
        images,
        image_scope: Some("word/document.xml".to_owned()),
        header_footer_parts: collect_header_footer_parts(bytes)?,
        comment_done: docx
            .comments_extended
            .children
            .iter()
            .map(|extended| (extended.paragraph_id.clone(), extended.done))
            .collect(),
        style_theme_colors: theme_colors.styles,
        run_theme_colors: theme_colors.runs,
        abstract_nums: docx.numberings.abstract_nums.clone(),
        numbering_instances: docx.numberings.numberings.clone(),
        numbering_multi_level_types,
        section_columns,
        section_types,
        section_line_numbering,
        section_page_numbering,
        table_row_xml,
        table_xml,
        paragraph_indent_xml,
        style_character_indents,
        style_suppress_line_numbers,
        style_paragraph_properties,
        style_run_properties,
        deleted_instructions,
        index_field_paragraph_xml,
        hyperlink_histories,
        ..Writer::default()
    };
    writer.components.insert("Document");
    writer.components.insert("Section");
    writer.comments = docx
        .comments
        .inner()
        .iter()
        .map(|comment| (comment.id(), comment.clone()))
        .collect();
    writer.hyperlink_targets = docx
        .hyperlinks
        .iter()
        .map(|(id, path, _)| (id.clone(), path.clone()))
        .collect();
    writer.sdt_plan = parse_sdt_plan(bytes)?;
    writer.header_footer_sdt_props =
        parse_header_footer_sdt_props(bytes, &writer.header_footer_parts)?;
    let document_attrs = reverse_document_attributes(&docx, &style_metadata, &writer, bytes)?;
    writer.line(0, &format!("<Document{document_attrs}>"));
    writer.emit_sections(&docx)?;
    writer.validate_index_field_paragraph_xml_consumed()?;
    writer.validate_table_row_xml_consumed()?;
    writer.validate_table_xml_consumed()?;
    writer.line(0, "</Document>");

    let imports = writer
        .components
        .iter()
        .copied()
        .collect::<Vec<_>>()
        .join(", ");
    Ok(ReversedDocument {
        jsx: format!(
            "import {{ {imports} }} from \"docx-jsx\";\n\nexport default (\n{});\n",
            writer.output
        ),
        assets: writer.assets,
    })
}

fn normalize_nested_toc_item_sdts(bytes: &[u8]) -> Result<Vec<u8>> {
    let xml = read_zip_text(bytes, "word/document.xml")?;
    let normalized = unwrap_nested_toc_item_sdts(&xml)?;
    if normalized == xml {
        return Ok(bytes.to_vec());
    }

    let mut source = zip::ZipArchive::new(Cursor::new(bytes.to_vec()))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let file = source
            .by_index(index)
            .map_err(|error| Error::Reverse(format!("cannot read DOCX entry: {error}")))?;
        let name = file.name().to_owned();
        if name == "word/document.xml" {
            output
                .start_file(name, file.options())
                .map_err(|error| Error::Reverse(format!("cannot rewrite DOCX entry: {error}")))?;
            output
                .write_all(normalized.as_bytes())
                .map_err(|error| Error::Reverse(format!("cannot rewrite document XML: {error}")))?;
        } else {
            output
                .raw_copy_file(file)
                .map_err(|error| Error::Reverse(format!("cannot copy DOCX entry: {error}")))?;
        }
    }
    output
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Reverse(format!("cannot finish normalized DOCX: {error}")))
}

fn unwrap_nested_toc_item_sdts(xml: &str) -> Result<String> {
    #[derive(Default)]
    struct SdtSpan {
        start: usize,
        content_start: Option<usize>,
        content_end: Option<usize>,
    }

    let mut reader = Reader::from_str(xml);
    let mut stack = Vec::<SdtSpan>::new();
    let mut replacements = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Reverse("document XML offset is out of range".to_owned()))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Reverse(format!("malformed document XML: {error}")))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::Reverse("document XML offset is out of range".to_owned()))?;
        match event {
            Event::Start(element) if element.local_name().as_ref() == b"sdt" => {
                stack.push(SdtSpan {
                    start,
                    ..SdtSpan::default()
                });
            }
            Event::Start(element) if element.local_name().as_ref() == b"sdtContent" => {
                if let Some(span) = stack.last_mut() {
                    span.content_start = Some(end);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"sdtContent" => {
                if let Some(span) = stack.last_mut() {
                    span.content_end = Some(start);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"sdt" => {
                let Some(span) = stack.pop() else {
                    return Err(Error::Reverse(
                        "document XML contains an unmatched `w:sdt` end tag".to_owned(),
                    ));
                };
                if stack.is_empty() {
                    continue;
                }
                if let (Some(content_start), Some(content_end)) =
                    (span.content_start, span.content_end)
                {
                    let block = &xml[span.start..end];
                    if block.contains("PAGEREF")
                        && block.contains("ToC")
                        && sdt_properties_are_empty(block)
                    {
                        replacements.push((
                            span.start,
                            end,
                            xml[content_start..content_end].to_owned(),
                        ));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err(Error::Reverse(
            "document XML contains an unclosed `w:sdt` element".to_owned(),
        ));
    }
    replacements.sort_unstable_by_key(|replacement| std::cmp::Reverse(replacement.0));
    let mut normalized = xml.to_owned();
    for (start, end, content) in replacements {
        // Later offsets are replaced first so earlier XML positions remain stable.
        normalized.replace_range(start..end, &content);
    }
    Ok(normalized)
}

fn sdt_properties_are_empty(block: &str) -> bool {
    let Some(property_start) = block.find("sdtPr") else {
        return false;
    };
    let Some(open_end) = block[property_start..].find('>') else {
        return false;
    };
    let open_end = property_start + open_end;
    let Some(tag_start) = block[..property_start].rfind('<') else {
        return false;
    };
    let property_name_end = property_start + "sdtPr".len();
    let attributes = block[property_name_end..open_end]
        .trim()
        .trim_end_matches('/');
    if !attributes.is_empty() {
        return false;
    }
    if block[tag_start..=open_end].trim_end().ends_with("/>") {
        return true;
    }
    let Some(close_offset) = block[open_end + 1..].find("sdtPr>") else {
        return false;
    };
    block[open_end + 1..open_end + 1 + close_offset]
        .trim()
        .is_empty()
}

fn normalize_deleted_instruction_parts(bytes: &[u8]) -> Result<(Vec<u8>, HashMap<String, String>)> {
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    let mut entries = Vec::with_capacity(source.len());
    let mut instructions = HashMap::new();
    let mut sequence = 0_u64;
    for index in 0..source.len() {
        let mut file = source
            .by_index(index)
            .map_err(|error| Error::Reverse(format!("cannot read DOCX part: {error}")))?;
        let name = file.name().to_owned();
        if name.starts_with("word/")
            && Path::new(&name)
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        {
            let mut data = Vec::new();
            file.read_to_end(&mut data)
                .map_err(|error| Error::Reverse(format!("cannot read `{name}`: {error}")))?;
            if data
                .windows(b"delInstrText".len())
                .any(|window| window == b"delInstrText")
            {
                let xml = String::from_utf8(data).map_err(|error| {
                    Error::Reverse(format!("`{name}` is not UTF-8 XML: {error}"))
                })?;
                let (normalized, found) =
                    normalize_deleted_instruction_xml(&xml, &name, &mut sequence)?;
                if found.is_empty() {
                    entries.push((name, xml.into_bytes()));
                } else {
                    instructions.extend(found);
                    entries.push((name, normalized.into_bytes()));
                }
            } else {
                entries.push((name, data));
            }
        } else {
            let mut data = Vec::new();
            file.read_to_end(&mut data)
                .map_err(|error| Error::Reverse(format!("cannot read `{name}`: {error}")))?;
            entries.push((name, data));
        }
    }
    if instructions.is_empty() {
        return Ok((bytes.to_vec(), instructions));
    }
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in entries {
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        if name.ends_with('/') {
            output.add_directory(name, options).map_err(|error| {
                Error::Reverse(format!("cannot rewrite DOCX directory: {error}"))
            })?;
        } else {
            output
                .start_file(name, options)
                .map_err(|error| Error::Reverse(format!("cannot rewrite DOCX part: {error}")))?;
            output.write_all(&data).map_err(|error| {
                Error::Reverse(format!("cannot write normalized DOCX part: {error}"))
            })?;
        }
    }
    let normalized = output
        .finish()
        .map_err(|error| Error::Reverse(format!("cannot finish normalized DOCX: {error}")))?
        .into_inner();
    Ok((normalized, instructions))
}

fn normalize_deleted_instruction_xml(
    xml: &str,
    part: &str,
    sequence: &mut u64,
) -> Result<(String, HashMap<String, String>)> {
    let mut output = String::with_capacity(xml.len());
    let mut instructions = HashMap::new();
    let mut scan_cursor = 0;
    let mut output_cursor = 0;
    while let Some(relative_start) = xml[scan_cursor..].find('<') {
        let start = scan_cursor + relative_start;
        let name_start = start + 1;
        if xml.as_bytes().get(name_start) == Some(&b'/') {
            scan_cursor = name_start + 1;
            continue;
        }
        let name_end = xml[name_start..]
            .find(|character: char| {
                character.is_whitespace() || character == '>' || character == '/'
            })
            .map_or(xml.len(), |offset| name_start + offset);
        let qualified_name = &xml[name_start..name_end];
        if qualified_name.rsplit(':').next() != Some("delInstrText") {
            scan_cursor = name_end;
            continue;
        }
        let open_end = xml[name_end..]
            .find('>')
            .map(|offset| name_end + offset + 1)
            .ok_or_else(|| Error::Reverse(format!("`{part}` contains an unterminated w:delInstrText start tag; repair the XML element")))?;
        let closing = format!("</{qualified_name}>");
        let close_start = xml[open_end..]
            .find(&closing)
            .map(|offset| open_end + offset)
            .ok_or_else(|| Error::Reverse(format!("`{part}` contains w:delInstrText without a matching end tag; repair the revision markup")))?;
        let instruction = quick_xml::escape::unescape(&xml[open_end..close_start])
            .map_err(|error| {
                Error::Reverse(format!(
                    "`{part}` has invalid w:delInstrText content: {error}"
                ))
            })?
            .into_owned();
        if instruction.is_empty() {
            return Err(Error::Reverse(format!(
                "`{part}` contains empty w:delInstrText; provide a non-empty instruction or remove the empty element"
            )));
        }
        let marker = loop {
            let candidate = format!("Z{sequence}Q");
            *sequence += 1;
            if !xml.contains(&candidate) {
                break candidate;
            }
        };
        let prefix = qualified_name
            .rsplit_once(':')
            .map_or("", |(prefix, _)| prefix);
        let normalized_name = if prefix.is_empty() {
            "delText".to_owned()
        } else {
            format!("{prefix}:delText")
        };
        output.push_str(&xml[output_cursor..name_start]);
        output.push_str(&normalized_name);
        output.push_str(&xml[name_end..open_end]);
        output.push_str(&marker);
        output.push_str("</");
        output.push_str(&normalized_name);
        output.push('>');
        instructions.insert(marker, instruction);
        output_cursor = close_start + closing.len();
        scan_cursor = output_cursor;
    }
    output.push_str(&xml[output_cursor..]);
    Ok((output, instructions))
}

fn read_ir_manifest(bytes: &[u8]) -> Result<Option<IrEnvelope>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    let Ok(mut file) = archive.by_name("docx-jsx/ir-v1.json") else {
        return Ok(None);
    };
    let mut manifest = String::new();
    file.read_to_string(&mut manifest)
        .map_err(|error| Error::Reverse(format!("cannot read embedded IR: {error}")))?;
    serde_json::from_str(&manifest)
        .map(Some)
        .map_err(|error| Error::Reverse(format!("invalid embedded IR: {error}")))
}

fn ir_to_jsx(ir: &IrEnvelope) -> Result<String> {
    let mut components = BTreeSet::new();
    let mut body = String::new();
    write_ir_node(&ir.document, 0, &mut body, &mut components)?;
    let imports = components.into_iter().collect::<Vec<_>>().join(", ");
    Ok(format!(
        "import {{ {imports} }} from \"docx-jsx\";\n\nexport default (\n{body});\n"
    ))
}

fn write_ir_node(
    node: &Node,
    depth: usize,
    output: &mut String,
    components: &mut BTreeSet<&'static str>,
) -> Result<()> {
    let name = node.kind.name();
    components.insert(name);
    output.push_str(&"  ".repeat(depth));
    output.push('<');
    output.push_str(name);
    let mut props = node.props.iter().collect::<Vec<_>>();
    props.sort_unstable_by_key(|(key, _)| *key);
    for (key, value) in props {
        output.push(' ');
        output.push_str(key);
        match value {
            Value::Bool(true) => {}
            Value::String(value) => {
                output.push_str("=\"");
                output.push_str(&escape_attr(value));
                output.push('"');
            }
            value => {
                output.push_str("={");
                output.push_str(
                    &serde_json::to_string(value)
                        .map_err(|error| Error::Reverse(error.to_string()))?,
                );
                output.push('}');
            }
        }
    }
    if node.children.is_empty() {
        output.push_str(" />\n");
        return Ok(());
    }
    output.push_str(">\n");
    for child in &node.children {
        match child {
            Child::Node(child) => write_ir_node(child, depth + 1, output, components)?,
            Child::String(value) => {
                output.push_str(&"  ".repeat(depth + 1));
                output.push('{');
                output.push_str(
                    &serde_json::to_string(value)
                        .map_err(|error| Error::Reverse(error.to_string()))?,
                );
                output.push('}');
                output.push('\n');
            }
            Child::Number(value) => {
                output.push_str(&"  ".repeat(depth + 1));
                output.push('{');
                output.push_str(&value.to_string());
                output.push_str("}\n");
            }
        }
    }
    output.push_str(&"  ".repeat(depth));
    output.push_str("</");
    output.push_str(name);
    output.push_str(">\n");
    Ok(())
}

#[derive(Default)]
struct Writer {
    output: String,
    components: BTreeSet<&'static str>,
    comments: HashMap<usize, Comment>,
    comment_done: HashMap<String, bool>,
    hyperlink_targets: HashMap<String, String>,
    hyperlink_histories: HashMap<String, VecDeque<Option<usize>>>,
    sdt_plan: SdtPlan,
    paragraph_index: usize,
    hyperlink_index: usize,
    body_sdt_index: usize,
    cell_sdt_index: usize,
    apply_sdt: bool,
    images: ImageCatalog,
    image_scope: Option<String>,
    header_footer_parts: HashMap<String, String>,
    header_footer_sdt_props: HashMap<String, Vec<SdtProps>>,
    header_footer_sdt_indices: HashMap<String, usize>,
    assets: BTreeMap<String, Vec<u8>>,
    style_theme_colors: HashMap<String, ThemeColorInfo>,
    run_theme_colors: Vec<ThemeColorInfo>,
    run_theme_index: usize,
    abstract_nums: Vec<AbstractNumbering>,
    numbering_instances: Vec<Numbering>,
    numbering_multi_level_types: HashMap<usize, String>,
    section_columns: Vec<SectionColumnsMetadata>,
    section_types: Vec<Option<String>>,
    section_line_numbering: Vec<Option<Map<String, Value>>>,
    section_page_numbering: Vec<Option<Map<String, Value>>>,
    table_row_xml: HashMap<String, Vec<TableRowXmlMetadata>>,
    table_row_indices: HashMap<String, usize>,
    table_xml: HashMap<String, Vec<TableXmlMetadata>>,
    table_indices: HashMap<String, usize>,
    paragraph_indent_xml: HashMap<String, Vec<ParagraphIndentMetadata>>,
    paragraph_indent_indices: HashMap<String, usize>,
    style_character_indents: HashMap<String, BTreeMap<String, i64>>,
    style_suppress_line_numbers: HashMap<String, bool>,
    style_paragraph_properties: HashMap<String, Map<String, Value>>,
    style_run_properties: HashMap<String, Map<String, Value>>,
    deleted_instructions: HashMap<String, String>,
    index_field_paragraph_xml: HashMap<String, Vec<Map<String, Value>>>,
    index_field_paragraph_indices: HashMap<String, usize>,
    in_deleted_revision: bool,
}

#[derive(Clone, Debug, Default)]
struct ParagraphIndentMetadata {
    values: BTreeMap<String, i64>,
    suppress_line_numbers: Option<bool>,
}

type RawIndexParagraphProperties = Vec<(String, BTreeMap<String, String>)>;
type RawIndexFieldParagraph = (String, RawIndexParagraphProperties);

#[derive(Clone, Debug, Default)]
struct TableRowXmlMetadata {
    cant_split: Option<bool>,
    width_before_type: Option<String>,
    width_after_type: Option<String>,
}

#[derive(Clone, Debug, Default)]
struct TableXmlMetadata {
    layout: Option<String>,
}

#[derive(Clone, Debug)]
struct SectionColumnsMetadata {
    count: usize,
    spacing_twips: Option<u64>,
    separator: bool,
    widths: Option<Vec<SectionColumnWidthMetadata>>,
}

impl Default for SectionColumnsMetadata {
    fn default() -> Self {
        Self {
            count: 1,
            spacing_twips: None,
            separator: false,
            widths: None,
        }
    }
}

#[derive(Clone, Debug, Default)]
struct SectionColumnWidthMetadata {
    width_twips: Option<u64>,
    spacing_twips: Option<u64>,
}

#[derive(Clone, Debug)]
struct ImageAsset {
    src: String,
    bytes: Vec<u8>,
}

#[derive(Clone, Debug, Default)]
struct ImageCatalog {
    scoped: HashMap<String, HashMap<String, ImageAsset>>,
    by_id: HashMap<String, ImageAsset>,
    used_names: BTreeSet<String>,
}

impl ImageCatalog {
    fn insert(&mut self, owner: &str, id: String, asset: ImageAsset) {
        self.by_id.entry(id.clone()).or_insert(asset.clone());
        self.scoped
            .entry(owner.to_owned())
            .or_default()
            .insert(id, asset);
    }

    fn get(&self, owner: Option<&str>, id: &str) -> Option<&ImageAsset> {
        owner
            .and_then(|owner| self.scoped.get(owner)?.get(id))
            .or_else(|| self.by_id.get(id))
    }
}

#[derive(Clone, Debug, Default)]
struct ThemeColorInfo {
    val: Option<String>,
    theme_color: Option<String>,
    theme_shade: Option<String>,
    theme_tint: Option<String>,
}

impl Writer {
    fn line(&mut self, depth: usize, value: &str) {
        self.output.push_str(&"  ".repeat(depth));
        self.output.push_str(value);
        self.output.push('\n');
    }

    fn emit_sections(&mut self, docx: &docx_rs::Docx) -> Result<()> {
        let nested = nest_block_children(&docx.document.children)?;
        let mut groups = split_section_groups(&nested);
        if groups.is_empty() {
            groups.push((nested.as_slice(), None));
        }
        if !self.section_columns.is_empty() && self.section_columns.len() != groups.len() {
            return Err(Error::Reverse(format!(
                "word/document.xml contains {} section properties for {} reconstructed sections; repair the section-break structure so each section has one `w:sectPr`",
                self.section_columns.len(),
                groups.len()
            )));
        }
        if !self.section_types.is_empty() && self.section_types.len() != groups.len() {
            return Err(Error::Reverse(format!(
                "word/document.xml contains {} section types for {} reconstructed sections; repair the section-break structure so each section is unambiguous",
                self.section_types.len(),
                groups.len()
            )));
        }
        if !self.section_line_numbering.is_empty()
            && self.section_line_numbering.len() != groups.len()
        {
            return Err(Error::Reverse(format!(
                "word/document.xml contains {} line-numbering entries for {} reconstructed sections; repair the section-break structure so each section is unambiguous",
                self.section_line_numbering.len(),
                groups.len()
            )));
        }
        if !self.section_page_numbering.is_empty()
            && self.section_page_numbering.len() != groups.len()
        {
            return Err(Error::Reverse(format!(
                "word/document.xml contains {} page-numbering entries for {} reconstructed sections; repair the section-break structure so each section has one `w:sectPr`",
                self.section_page_numbering.len(),
                groups.len()
            )));
        }
        for (index, (children, break_property)) in groups.into_iter().enumerate() {
            let property = break_property.unwrap_or(&docx.document.section_property);
            let columns = self.section_columns.get(index).cloned().unwrap_or_default();
            let section_type = self.section_types.get(index).and_then(Option::as_deref);
            let line_numbering = self
                .section_line_numbering
                .get(index)
                .and_then(Option::as_ref);
            let page_numbering = self
                .section_page_numbering
                .get(index)
                .and_then(Option::as_ref);
            self.line(
                1,
                &format!(
                    "<Section{}>",
                    section_jsx_attrs(
                        property,
                        &columns,
                        section_type,
                        line_numbering,
                        page_numbering,
                    )?
                ),
            );
            self.apply_sdt = false;
            self.headers_and_footers(property, 2)?;
            self.apply_sdt = true;
            self.emit_nested_blocks(children, 2)?;
            self.line(1, "</Section>");
        }
        Ok(())
    }

    fn emit_nested_blocks(&mut self, children: &[NestedBlock<'_>], depth: usize) -> Result<()> {
        let mut index = 0;
        while index < children.len() {
            if let Some((mut field, consumed_paragraphs)) = unwrapped_index_at(children, index)? {
                let Some(NestedBlock::Child(DocumentChild::Paragraph(begin))) = children.get(index)
                else {
                    return Err(Error::Reverse(
                        "unwrapped index start paragraph disappeared during reconstruction"
                            .to_owned(),
                    ));
                };
                let instruction = toc_instruction_from_paragraph(begin).ok_or_else(|| {
                    Error::Reverse("unwrapped index field instruction is malformed".to_owned())
                })?;
                self.append_index_field_paragraph_properties(&mut field, &instruction)?;
                self.index_component(&field, depth)?;
                self.paragraph_index += consumed_paragraphs;
                index += consumed_paragraphs;
                continue;
            }
            if section_break_of(&children[index]).is_some()
                && is_section_break_marker(&children[index])
            {
                index += 1;
                continue;
            }
            if let Some(num_id) = nested_block_num_id(&children[index]) {
                let start = index;
                index += 1;
                while index < children.len()
                    && nested_block_num_id(&children[index]) == Some(num_id)
                {
                    index += 1;
                }
                let paragraphs = children[start..index]
                    .iter()
                    .map(nested_block_paragraph)
                    .collect::<Result<Vec<_>>>()?;
                self.emit_list_from_paragraphs(num_id, paragraphs, depth)?;
            } else {
                self.emit_nested_block(&children[index], depth)?;
                index += 1;
            }
        }
        Ok(())
    }

    fn emit_nested_block(&mut self, child: &NestedBlock<'_>, depth: usize) -> Result<()> {
        match child {
            NestedBlock::Child(DocumentChild::Paragraph(value)) => self.paragraph(value, depth),
            NestedBlock::Child(DocumentChild::Table(value)) => self.table(value, depth),
            NestedBlock::Child(value) => Err(unsupported("document", value)),
            NestedBlock::Bookmark { name, children } => {
                self.components.insert("Bookmark");
                if children.is_empty() {
                    self.line(depth, &format!("<Bookmark{} />", attr("name", name)));
                    return Ok(());
                }
                self.line(depth, &format!("<Bookmark{}>", attr("name", name)));
                self.emit_nested_blocks(children, depth + 1)?;
                self.line(depth, "</Bookmark>");
                Ok(())
            }
            NestedBlock::ContentControl { tag } => self.body_structured_tag(tag, depth),
        }
    }

    fn body_structured_tag(&mut self, tag: &StructuredDataTag, depth: usize) -> Result<()> {
        if let Some(mut index) = index_from_sdt(tag, &self.deleted_instructions)? {
            let deletion = deleted_index_field_from_sdt(tag, &self.deleted_instructions)?;
            let instruction = collect_sdt_instruction(tag, deletion.as_ref());
            self.append_index_field_paragraph_properties(&mut index, &instruction)?;
            if self.apply_sdt {
                if let Some(props) = self.sdt_plan.body_tags.get(self.body_sdt_index) {
                    add_index_sdt_properties(&mut index, props);
                }
                self.body_sdt_index += 1;
            }
            self.advance_paragraphs_in_sdt(tag);
            self.index_component(&index, depth)?;
            return Ok(());
        }
        self.block_content_control(tag, depth)
    }

    fn append_index_field_paragraph_properties(
        &mut self,
        index: &mut IndexField,
        instruction: &str,
    ) -> Result<()> {
        let key = instruction.trim();
        let index_in_bucket = self
            .index_field_paragraph_indices
            .entry(key.to_owned())
            .or_default();
        let Some(properties) = self
            .index_field_paragraph_xml
            .get(key)
            .and_then(|values| values.get(*index_in_bucket))
            .cloned()
        else {
            return Err(Error::Reverse(format!(
                "index field `{key}` has no matching raw paragraph properties in word/document.xml; retain the source package or repair the field structure"
            )));
        };
        *index_in_bucket += 1;
        append_index_field_paragraph_properties(index, properties)
    }

    fn validate_index_field_paragraph_xml_consumed(&self) -> Result<()> {
        for (instruction, values) in &self.index_field_paragraph_xml {
            let consumed = self
                .index_field_paragraph_indices
                .get(instruction)
                .copied()
                .unwrap_or_default();
            if consumed != values.len() {
                return Err(Error::Reverse(format!(
                    "word/document.xml has {} TOC field paragraph(s) for `{instruction}`, but only {consumed} were reconstructed as index components; repair the field structure so each TOC field has a matching index control or adjacent begin/end paragraphs",
                    values.len()
                )));
            }
        }
        Ok(())
    }

    fn advance_paragraphs_in_sdt(&mut self, tag: &StructuredDataTag) {
        for child in &tag.children {
            match child {
                StructuredDataTagChild::Paragraph(_) => self.paragraph_index += 1,
                StructuredDataTagChild::Table(table) => {
                    self.advance_paragraphs_in_table(table);
                }
                StructuredDataTagChild::StructuredDataTag(inner) => {
                    self.advance_paragraphs_in_sdt(inner);
                }
                _ => {}
            }
        }
    }

    fn advance_paragraphs_in_table(&mut self, table: &docx_rs::Table) {
        for TableChild::TableRow(row) in &table.rows {
            for TableRowChild::TableCell(cell) in &row.cells {
                for child in &cell.children {
                    match child {
                        TableCellContent::Paragraph(_) => self.paragraph_index += 1,
                        TableCellContent::Table(inner) => self.advance_paragraphs_in_table(inner),
                        TableCellContent::StructuredDataTag(tag) => {
                            self.advance_paragraphs_in_sdt(tag);
                        }
                        TableCellContent::TableOfContents(_) => {}
                    }
                }
            }
        }
    }

    fn index_component(&mut self, index: &IndexField, depth: usize) -> Result<()> {
        match index {
            IndexField::Contents {
                start,
                end,
                attributes,
                items,
                before,
                after,
            } => {
                self.index_with_items(
                    "TableOfContents",
                    depth,
                    &format!(" startLevel={{{start}}} endLevel={{{end}}}{attributes}"),
                    items,
                    before,
                    after,
                )?;
            }
            IndexField::Figures {
                label,
                include_label_and_number,
                separator,
                attributes,
                items,
                before,
                after,
            } => {
                let mut properties = attr("label", label);
                if !include_label_and_number {
                    properties.push_str(" includeLabelAndNumber={false}");
                }
                if let Some(separator) = separator {
                    properties.push_str(&attr("separator", separator));
                }
                properties.push_str(attributes);
                self.index_with_items("TableOfFigures", depth, &properties, items, before, after)?;
            }
            IndexField::Entries {
                identifier,
                attributes,
                items,
                before,
                after,
            } => {
                let mut properties = attr("identifier", identifier);
                properties.push_str(attributes);
                self.index_with_items("TableOfEntries", depth, &properties, items, before, after)?;
            }
        }
        Ok(())
    }

    fn index_with_items(
        &mut self,
        component: &'static str,
        depth: usize,
        attributes: &str,
        items: &[TocItemField],
        before: &[IndexContentBlock],
        after: &[IndexContentBlock],
    ) -> Result<()> {
        if items.is_empty() && before.is_empty() && after.is_empty() {
            self.empty(component, depth, attributes);
            return Ok(());
        }
        self.components.insert(component);
        self.line(depth, &format!("<{component}{attributes}>"));
        if !before.is_empty() {
            self.components.insert("TocBefore");
            self.line(depth + 1, "<TocBefore>");
            for content in before {
                match content {
                    IndexContentBlock::Paragraph(paragraph) => {
                        self.paragraph(paragraph, depth + 2)?;
                    }
                    IndexContentBlock::Table(table) => self.table(table, depth + 2)?,
                }
            }
            self.line(depth + 1, "</TocBefore>");
        }
        for item in items {
            self.components.insert("TocItem");
            let mut attributes = attr("text", &item.text);
            attributes.push_str(&attr("bookmark", &item.bookmark));
            if item.level != 1 {
                let _ = write!(attributes, " level={{{}}}", item.level);
            }
            if item.page_ref != "1" {
                attributes.push_str(&attr("pageRef", &item.page_ref));
            }
            if item.tab_position != 8_400 {
                let _ = write!(
                    attributes,
                    " tabPosition={{{}}}",
                    format_twips_as_points(i64::from(item.tab_position))
                );
            }
            self.line(depth + 1, &format!("<TocItem{attributes} />"));
        }
        if !after.is_empty() {
            self.components.insert("TocAfter");
            self.line(depth + 1, "<TocAfter>");
            for content in after {
                match content {
                    IndexContentBlock::Paragraph(paragraph) => {
                        self.paragraph(paragraph, depth + 2)?;
                    }
                    IndexContentBlock::Table(table) => self.table(table, depth + 2)?,
                }
            }
            self.line(depth + 1, "</TocAfter>");
        }
        self.line(depth, &format!("</{component}>"));
        Ok(())
    }

    fn block_content_control(&mut self, tag: &StructuredDataTag, depth: usize) -> Result<()> {
        self.components.insert("ContentControl");
        let mut props = sdt_props_from_tag(tag);
        if let Some(part) = self
            .image_scope
            .as_deref()
            .filter(|part| self.header_footer_sdt_props.contains_key(*part))
        {
            let index = self
                .header_footer_sdt_indices
                .entry(part.to_owned())
                .or_default();
            if let Some(xml_props) = self
                .header_footer_sdt_props
                .get(part)
                .and_then(|values| values.get(*index))
            {
                merge_sdt_props(&mut props, xml_props);
            }
            *index += 1;
        } else if self.apply_sdt {
            if let Some(xml_props) = self.sdt_plan.body_tags.get(self.body_sdt_index) {
                merge_sdt_props(&mut props, xml_props);
            }
            self.body_sdt_index += 1;
        }
        if tag.children.is_empty() {
            return Err(Error::Reverse(
                "block-level structured document tag has no representable content; add paragraphs or tables inside `w:sdtContent`"
                    .to_owned(),
            ));
        }
        self.line(depth, &format!("<ContentControl{}>", sdt_attrs(&props)));
        self.emit_sdt_children(&tag.children, depth + 1)?;
        self.line(depth, "</ContentControl>");
        Ok(())
    }

    fn headers_and_footers(
        &mut self,
        property: &docx_rs::SectionProperty,
        depth: usize,
    ) -> Result<()> {
        for (kind, value) in [
            ("default", property.header.as_ref()),
            ("first", property.first_header.as_ref()),
            ("even", property.even_header.as_ref()),
        ] {
            if let Some((rid, header)) = value {
                self.components.insert("Header");
                self.line(depth, &format!(r#"<Header type="{kind}">"#));
                self.with_image_scope(rid, |writer| {
                    writer.emit_header_children(&header.children, depth + 1)
                })?;
                self.line(depth, "</Header>");
            }
        }
        for (kind, value) in [
            ("default", property.footer.as_ref()),
            ("first", property.first_footer.as_ref()),
            ("even", property.even_footer.as_ref()),
        ] {
            if let Some((rid, footer)) = value {
                self.components.insert("Footer");
                self.line(depth, &format!(r#"<Footer type="{kind}">"#));
                self.with_image_scope(rid, |writer| {
                    writer.emit_footer_children(&footer.children, depth + 1)
                })?;
                self.line(depth, "</Footer>");
            }
        }
        Ok(())
    }

    fn with_image_scope(
        &mut self,
        relationship_id: &str,
        emit: impl FnOnce(&mut Self) -> Result<()>,
    ) -> Result<()> {
        let previous = self.image_scope.take();
        self.image_scope = self
            .header_footer_parts
            .get(relationship_id)
            .cloned()
            .or_else(|| previous.clone());
        let result = emit(self);
        self.image_scope = previous;
        result
    }

    fn footer_paragraph(&mut self, paragraph: &docx_rs::Paragraph, depth: usize) -> Result<()> {
        for child in &paragraph.children {
            let ParagraphChild::Run(run) = child else {
                continue;
            };
            for child in &run.children {
                let RunChild::Drawing(drawing) = child else {
                    continue;
                };
                let Some(DrawingData::TextBox(_)) = &drawing.data else {
                    continue;
                };
                return Err(text_box_reverse_error("footer"));
            }
        }
        self.paragraph(paragraph, depth)
    }

    #[expect(
        clippy::too_many_lines,
        reason = "paragraph properties are emitted together to keep raw OOXML fallbacks aligned"
    )]
    fn paragraph(&mut self, paragraph: &docx_rs::Paragraph, depth: usize) -> Result<()> {
        self.components.insert("Paragraph");
        let raw_indent = self.next_paragraph_indent_metadata()?;
        let property = json(&paragraph.property)?;
        let mut attrs = Vec::new();
        if let Some(value) = raw_indent.suppress_line_numbers {
            attrs.push(jsx_prop("suppressLineNumbers", &Value::Bool(value))?);
        }
        if let Some(value) = property.get("alignment").and_then(scalar_string) {
            attrs.push(attr("align", value));
        }
        if let Some(value) = property
            .get("style")
            .and_then(scalar_string)
            .or_else(|| nested_string(&property, &["style", "val"]))
            .or_else(|| nested_string(&property, &["style", "styleId"]))
        {
            attrs.push(attr("style", value));
        }
        if let Some(spacing) = property.get("lineSpacing") {
            for (source, target) in [
                ("before", "spacingBefore"),
                ("after", "spacingAfter"),
                ("line", "lineSpacing"),
            ] {
                if let Some(value) = spacing.get(source).and_then(Value::as_i64) {
                    attrs.push(point_attr(target, value));
                }
            }
            for (source, target) in [
                ("beforeLines", "spacingBeforeLines"),
                ("afterLines", "spacingAfterLines"),
            ] {
                if let Some(value) = spacing.get(source).and_then(Value::as_u64) {
                    attrs.push(format!(" {target}={{{value}}}"));
                }
            }
            if let Some(value) = spacing.get("lineRule").and_then(Value::as_str) {
                attrs.push(attr("lineRule", value));
            }
        }
        if let Some(indent) = property.get("indent") {
            for (source, target) in [("start", "indentLeft"), ("end", "indentRight")] {
                if let Some(value) = indent.get(source).and_then(Value::as_i64) {
                    attrs.push(point_attr(target, value));
                }
            }
            if let Some(special) = indent.get("specialIndent")
                && let (Some(kind), Some(value)) = (
                    special.get("type").and_then(Value::as_str),
                    special.get("val").and_then(Value::as_i64),
                )
            {
                let target = if kind == "hanging" {
                    "hanging"
                } else {
                    "firstLine"
                };
                attrs.push(point_attr(target, value));
            }
        }
        for (name, value) in &raw_indent.values {
            attrs.push(jsx_prop(name, &Value::from(*value))?);
        }
        if let Some(run_properties) =
            reverse_run_properties(property.get("runProperty").unwrap_or(&Value::Null))?
        {
            for name in [
                "font",
                "fonts",
                "size",
                "bold",
                "italic",
                "color",
                "characterSpacing",
            ] {
                if let Some(value) = run_properties.get(name) {
                    attrs.push(jsx_prop(name, value)?);
                }
            }
        }
        append_paragraph_mark_revision_attributes(&mut attrs, &property)?;
        if property.get("keepNext").and_then(Value::as_bool) == Some(true) {
            attrs.push(" keepNext".to_owned());
        }
        if property.get("keepLines").and_then(Value::as_bool) == Some(true) {
            attrs.push(" keepLines".to_owned());
        }
        if let Some(value) = property
            .get("outlineLvl")
            .and_then(Value::as_u64)
            .or_else(|| property.get("outlineLvl").and_then(scalar_u64))
        {
            attrs.push(format!(" outlineLevel={{{value}}}"));
        }
        self.line(depth, &format!("<Paragraph{}>", attrs.concat()));
        let events = if self.apply_sdt {
            let index = self.paragraph_index;
            self.paragraph_index += 1;
            self.sdt_plan
                .paragraphs
                .get(index)
                .cloned()
                .unwrap_or_default()
        } else {
            Vec::new()
        };
        let nested = nest_inline_children(&paragraph.children, &self.comments, &events)?;
        for child in nested {
            self.emit_nested_inline(&child, depth + 1)?;
        }
        self.line(depth, "</Paragraph>");
        Ok(())
    }

    fn emit_nested_inline(&mut self, child: &NestedInline<'_>, depth: usize) -> Result<()> {
        match child {
            NestedInline::Child(value) => self.paragraph_child(value, depth),
            NestedInline::InlineBookmark { name, children } => {
                self.components.insert("InlineBookmark");
                if children.is_empty() {
                    self.line(depth, &format!("<InlineBookmark{} />", attr("name", name)));
                    return Ok(());
                }
                self.line(depth, &format!("<InlineBookmark{}>", attr("name", name)));
                for child in children {
                    self.emit_nested_inline(child, depth + 1)?;
                }
                self.line(depth, "</InlineBookmark>");
                Ok(())
            }
            NestedInline::Comment { comment, children } => self.comment(comment, children, depth),
            NestedInline::ContentControl { props, children } => {
                self.content_control(props, children, depth)
            }
        }
    }

    fn comment(
        &mut self,
        comment: &Comment,
        children: &[NestedInline<'_>],
        depth: usize,
    ) -> Result<()> {
        self.components.insert("Comment");
        let text = comment_text(comment)?;
        self.line(
            depth,
            &format!(
                "<Comment{}{}{}{}>",
                attr("text", &text),
                attr("author", &comment.author),
                attr("date", &comment.date),
                comment_identity_attrs(comment, comment_done(comment, &self.comment_done))
            ),
        );
        for child in children {
            self.emit_nested_inline(child, depth + 1)?;
        }
        self.line(depth, "</Comment>");
        Ok(())
    }

    fn content_control(
        &mut self,
        props: &SdtProps,
        children: &[NestedInline<'_>],
        depth: usize,
    ) -> Result<()> {
        self.components.insert("ContentControl");
        if children.is_empty() {
            return Err(Error::Reverse(
                "structured document tag has no representable content; add inline runs inside `w:sdtContent` or remove the empty control"
                    .to_owned(),
            ));
        }
        self.line(depth, &format!("<ContentControl{}>", sdt_attrs(props)));
        for child in children {
            self.emit_nested_inline(child, depth + 1)?;
        }
        self.line(depth, "</ContentControl>");
        Ok(())
    }

    fn hyperlink(&mut self, value: &Hyperlink, depth: usize) -> Result<()> {
        self.components.insert("Hyperlink");
        let attrs = self.hyperlink_attrs(value)?;
        self.line(depth, &format!("<Hyperlink{attrs}>"));
        let events = if self.apply_sdt {
            let index = self.hyperlink_index;
            self.hyperlink_index += 1;
            self.sdt_plan
                .hyperlinks
                .get(index)
                .cloned()
                .unwrap_or_default()
        } else {
            Vec::new()
        };
        let nested = nest_inline_children(&value.children, &self.comments, &events)?;
        for child in nested {
            self.emit_nested_inline(&child, depth + 1)?;
        }
        self.line(depth, "</Hyperlink>");
        Ok(())
    }

    fn hyperlink_attrs(&mut self, value: &Hyperlink) -> Result<String> {
        let data = json(value)?;
        let (link, key) = match &value.link {
            HyperlinkData::Anchor { anchor } => (attr("anchor", anchor), format!("a:{anchor}")),
            HyperlinkData::External { rid, path } => {
                let href = if path.is_empty() {
                    self.hyperlink_targets.get(rid).ok_or_else(|| {
                        Error::Reverse(format!(
                            "hyperlink relationship `{rid}` is missing a target; restore the external relationship in word/_rels/document.xml.rels"
                        ))
                    })?
                } else {
                    path
                };
                (attr("href", href), format!("r:{rid}"))
            }
        };
        let raw_history = self
            .hyperlink_histories
            .get_mut(&key)
            .and_then(VecDeque::pop_front);
        let history_value = raw_history.unwrap_or_else(|| {
            data.get("history")
                .and_then(Value::as_u64)
                .map(|value| usize::try_from(value).unwrap_or(usize::MAX))
        });
        let history = match history_value {
            Some(0) => " history={false}",
            Some(1) => " history",
            Some(value) => {
                return Err(Error::Reverse(format!(
                    "hyperlink `w:history` has unsupported value `{value}`; use 0 or 1"
                )));
            }
            None => "",
        };
        Ok(format!("{link}{history}"))
    }

    fn paragraph_child(&mut self, child: &ParagraphChild, depth: usize) -> Result<()> {
        match child {
            ParagraphChild::Run(value) if run_is_ignorable(value) => Ok(()),
            ParagraphChild::Run(value) => self.run(value, depth),
            ParagraphChild::Insert(value) => self.inserted(value, depth),
            ParagraphChild::Delete(value) => self.deleted(value, depth),
            ParagraphChild::MoveFrom(value) => self.moved_from(value, depth),
            ParagraphChild::MoveTo(value) => self.moved_to(value, depth),
            ParagraphChild::Hyperlink(value) => self.hyperlink(value, depth),
            ParagraphChild::StructuredDataTag(value) => self.native_content_control(value, depth),
            ParagraphChild::PageNum(_) => {
                self.empty("PageNumber", depth, "");
                Ok(())
            }
            ParagraphChild::NumPages(_) => {
                self.empty("TotalPages", depth, "");
                Ok(())
            }
            ParagraphChild::BookmarkStart(_)
            | ParagraphChild::BookmarkEnd(_)
            | ParagraphChild::CommentStart(_)
            | ParagraphChild::CommentEnd(_) => Err(Error::Reverse(
                "unpaired bookmark or comment marker remained after range reconstruction; pair matching start/end IDs so the range can nest"
                    .to_owned(),
            )),
        }
    }

    fn cell_content_control(&mut self, tag: &StructuredDataTag, depth: usize) -> Result<()> {
        let mut props = sdt_props_from_tag(tag);
        if self.apply_sdt {
            if let Some(xml_props) = self.sdt_plan.cell_tags.get(self.cell_sdt_index) {
                merge_sdt_props(&mut props, xml_props);
            }
            self.cell_sdt_index += 1;
        }
        self.emit_inline_content_control(tag, &props, depth)
    }

    fn native_content_control(&mut self, tag: &StructuredDataTag, depth: usize) -> Result<()> {
        let props = sdt_props_from_tag(tag);
        self.emit_inline_content_control(tag, &props, depth)
    }

    fn emit_inline_content_control(
        &mut self,
        tag: &StructuredDataTag,
        props: &SdtProps,
        depth: usize,
    ) -> Result<()> {
        let runs = tag
            .children
            .iter()
            .map(|child| match child {
                StructuredDataTagChild::Run(run) => Ok(Some(run.as_ref())),
                value => Err(unsupported("content control", value)),
            })
            .collect::<Result<Vec<_>>>()?
            .into_iter()
            .flatten()
            .filter(|run| !run_is_ignorable(run))
            .collect::<Vec<_>>();
        if runs.is_empty() {
            return Err(Error::Reverse(
                "structured document tag has no representable content; add inline runs inside `w:sdtContent` or remove the empty control"
                    .to_owned(),
            ));
        }
        self.components.insert("ContentControl");
        self.line(depth, &format!("<ContentControl{}>", sdt_attrs(props)));
        for run in runs {
            self.run(run, depth + 1)?;
        }
        self.line(depth, "</ContentControl>");
        Ok(())
    }

    fn inserted(&mut self, value: &Insert, depth: usize) -> Result<()> {
        self.components.insert("Inserted");
        self.line(
            depth,
            &format!(
                "<Inserted{}{}>",
                attr("author", &value.author),
                attr("date", &value.date)
            ),
        );
        for child in &value.children {
            match child {
                InsertChild::Run(run) => self.run(run, depth + 1)?,
                InsertChild::Delete(deleted) => self.deleted(deleted, depth + 1)?,
                child => return Err(unsupported("inserted revision", child)),
            }
        }
        self.line(depth, "</Inserted>");
        Ok(())
    }

    fn deleted(&mut self, value: &Delete, depth: usize) -> Result<()> {
        self.components.insert("Deleted");
        self.line(
            depth,
            &format!(
                "<Deleted{}{}>",
                attr("author", &value.author),
                attr("date", &value.date)
            ),
        );
        let was_in_deleted_revision = std::mem::replace(&mut self.in_deleted_revision, true);
        let result = (|| {
            for child in nest_deleted_ranges(&value.children, &self.comments)? {
                self.emit_deleted_inline(&child, depth + 1)?;
            }
            Ok(())
        })();
        self.in_deleted_revision = was_in_deleted_revision;
        result?;
        self.line(depth, "</Deleted>");
        Ok(())
    }

    fn emit_deleted_inline(&mut self, child: &DeletedInline<'_>, depth: usize) -> Result<()> {
        match child {
            DeletedInline::Run(run) => self.run(run, depth),
            DeletedInline::Comment { comment, children } => {
                self.components.insert("Comment");
                let text = comment_text(comment)?;
                self.line(
                    depth,
                    &format!(
                        "<Comment{}{}{}{}>",
                        attr("text", &text),
                        attr("author", &comment.author),
                        attr("date", &comment.date),
                        comment_identity_attrs(comment, comment_done(comment, &self.comment_done))
                    ),
                );
                for child in children {
                    self.emit_deleted_inline(child, depth + 1)?;
                }
                self.line(depth, "</Comment>");
                Ok(())
            }
        }
    }

    fn moved_from(&mut self, value: &MoveFrom, depth: usize) -> Result<()> {
        self.components.insert("MovedFrom");
        self.line(
            depth,
            &format!(
                "<MovedFrom{}{}>",
                attr("author", &value.author),
                attr("date", &value.date)
            ),
        );
        for child in &value.children {
            match child {
                MoveFromChild::Run(run) => self.run(run, depth + 1)?,
                child => return Err(unsupported("moved-from revision", child)),
            }
        }
        self.line(depth, "</MovedFrom>");
        Ok(())
    }

    fn moved_to(&mut self, value: &MoveTo, depth: usize) -> Result<()> {
        self.components.insert("MovedTo");
        self.line(
            depth,
            &format!(
                "<MovedTo{}{}>",
                attr("author", &value.author),
                attr("date", &value.date)
            ),
        );
        for child in &value.children {
            match child {
                MoveToChild::Run(run) => self.run(run, depth + 1)?,
                MoveToChild::Delete(deleted) => self.deleted(deleted, depth + 1)?,
                child => return Err(unsupported("moved-to revision", child)),
            }
        }
        self.line(depth, "</MovedTo>");
        Ok(())
    }

    fn run_attributes(&mut self, property: &Value) -> Result<Vec<String>> {
        let mut attrs = Vec::new();
        for (key, name) in [("bold", "bold"), ("italic", "italic"), ("strike", "strike")] {
            if property_enabled(property, key) {
                attrs.push(format!(" {name}"));
            }
        }
        if property.get("caps").is_some() {
            attrs.push(jsx_prop(
                "caps",
                &Value::Bool(property_enabled(property, "caps")),
            )?);
        }
        if property.get("vanish").is_some() {
            attrs.push(jsx_prop("hidden", &Value::Bool(true))?);
        }
        if property.get("specVanish").is_some() {
            attrs.push(jsx_prop("specVanish", &Value::Bool(true))?);
        }
        if property.get("underline").is_some() {
            attrs.push(" underline".to_owned());
        }
        if let Some(value) = property
            .get("color")
            .and_then(scalar_string)
            .filter(|value| value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
        {
            attrs.push(attr("color", value));
            attrs.extend(self.next_run_theme_attrs(value));
        }
        if let Some(value) = property.get("sz").and_then(scalar_u64) {
            let size = if value % 2 == 0 {
                (value / 2).to_string()
            } else {
                format!("{}.5", value / 2)
            };
            attrs.push(format!(" size={{{size}}}"));
        }
        if let Some((name, value)) = reverse_fonts_property(property, "font", "fonts") {
            attrs.push(jsx_prop(&name, &value)?);
        }
        if let Some(value) = property
            .get("style")
            .and_then(scalar_string)
            .or_else(|| nested_string(property, &["style", "val"]))
        {
            attrs.push(attr("style", value));
        }
        if let Some(value) = property.get("vertAlign").and_then(scalar_string) {
            if !["baseline", "superscript", "subscript"].contains(&value) {
                return Err(Error::Reverse(format!(
                    "run has unsupported vertical alignment `{value}`; use baseline, superscript, or subscript"
                )));
            }
            attrs.push(attr("verticalAlign", value));
        }
        Ok(attrs)
    }

    fn run(&mut self, run: &docx_rs::Run, depth: usize) -> Result<()> {
        self.components.insert("Run");
        let property = json(&run.run_property)?;
        if property.get("ins").is_some_and(|value| !value.is_null())
            || property.get("del").is_some_and(|value| !value.is_null())
        {
            return Err(Error::Reverse(
                "run properties contain `w:rPr/w:ins` or `w:rPr/w:del`, which is schema-invalid in this location and cannot be represented safely; use `<Inserted>` or `<Deleted>` for tracked text, or preserve the original DOCX until run-format revision support is available".to_owned(),
            ));
        }
        let attrs = self.run_attributes(&property)?;
        self.line(depth, &format!("<Run{}>", attrs.concat()));
        for child in &run.children {
            match child {
                RunChild::Text(value) => self.line(depth + 1, &jsx_string(&value.text)?),
                RunChild::DeleteText(value) => {
                    let data = json(value)?;
                    let text = data.get("text").and_then(Value::as_str).unwrap_or_default();
                    if let Some(instruction) = self.deleted_instructions.remove(text) {
                        if !self.in_deleted_revision {
                            return Err(Error::Reverse(
                                "w:delInstrText was found outside a tracked deletion; move it inside w:del and retry".to_owned(),
                            ));
                        }
                        self.components.insert("DeletedInstruction");
                        self.empty(
                            "DeletedInstruction",
                            depth + 1,
                            &attr("instruction", &instruction),
                        );
                    } else {
                        self.line(depth + 1, &jsx_string(text)?);
                    }
                }
                RunChild::Tab(_) => self.empty("Tab", depth + 1, ""),
                RunChild::PTab(tab) => self.positional_tab(tab, depth + 1)?,
                RunChild::Break(value) => {
                    let data = json(value)?;
                    let kind = data
                        .get("breakType")
                        .or_else(|| data.get("type"))
                        .and_then(Value::as_str)
                        .unwrap_or("textWrapping");
                    let kind = match kind {
                        "page" => "page",
                        "column" => "column",
                        _ => "line",
                    };
                    self.empty("Break", depth + 1, &attr("type", kind));
                }
                RunChild::CarriageReturn(_) => self.empty("CarriageReturn", depth + 1, ""),
                RunChild::FieldChar(_) | RunChild::InstrTextString(_) => {}
                RunChild::Drawing(drawing) => self.drawing(drawing, depth + 1)?,
                value => return Err(unsupported("run", value)),
            }
        }
        self.line(depth, "</Run>");
        Ok(())
    }

    fn positional_tab(&mut self, tab: &docx_rs::PositionalTab, depth: usize) -> Result<()> {
        self.components.insert("PositionalTab");
        let data = json(tab)?;
        let mut attrs = String::new();
        if let Some(value) = data
            .get("alignment")
            .and_then(Value::as_str)
            .filter(|value| *value != "left")
        {
            attrs.push_str(&attr("align", value));
        }
        if let Some(value) = data
            .get("relativeTo")
            .and_then(Value::as_str)
            .filter(|value| *value != "margin")
        {
            attrs.push_str(&attr("relativeTo", value));
        }
        if let Some(value) = data
            .get("leader")
            .and_then(Value::as_str)
            .filter(|value| *value != "none")
        {
            attrs.push_str(&attr("leader", value));
        }
        self.empty("PositionalTab", depth, &attrs);
        Ok(())
    }

    fn drawing(&mut self, drawing: &docx_rs::Drawing, depth: usize) -> Result<()> {
        let Some(data) = &drawing.data else {
            return Err(Error::Reverse(
                "drawing has no representable picture; restore the `w:drawing` content or replace it with a raster image"
                    .to_owned(),
            ));
        };
        match data {
            DrawingData::Pic(picture) => self.image(picture, depth),
            DrawingData::TextBox(_) => Err(text_box_reverse_error("run")),
        }
    }

    fn image(&mut self, picture: &Pic, depth: usize) -> Result<()> {
        let asset = self.lookup_image(&picture.id).ok_or_else(|| {
            Error::Reverse(format!(
                "image relationship `{}` is missing media bytes; restore the image part in word/media and the relationship in word/_rels/document.xml.rels or the matching header/footer rels",
                picture.id
            ))
        })?;
        if image::load_from_memory(&asset.bytes).is_err() {
            return Err(Error::Reverse(format!(
                "unsupported image format at `{}`; convert EMF/WMF/PICT drawings to PNG or JPEG before reversing",
                asset.src
            )));
        }
        self.components.insert("Image");
        self.assets.insert(asset.src.clone(), asset.bytes);
        let mut attrs = attr("src", &asset.src);
        attrs.push_str(&emu_point_attr("width", i64::from(picture.size.0)));
        attrs.push_str(&emu_point_attr("height", i64::from(picture.size.1)));
        if !picture.id.is_empty() {
            attrs.push_str(&attr("relationshipId", &picture.id));
        }
        if picture.rot != 0 {
            let _ = write!(attrs, " rotate={{{}}}", picture.rot);
        }
        if picture.position_type == DrawingPositionType::Anchor {
            attrs.push_str(" floating");
            attrs.push_str(&reverse_image_anchor_attrs(picture));
        }
        self.empty("Image", depth, &attrs);
        Ok(())
    }

    fn lookup_image(&self, id: &str) -> Option<ImageAsset> {
        self.images.get(self.image_scope.as_deref(), id).cloned()
    }

    fn next_run_theme_attrs(&mut self, hex: &str) -> Vec<String> {
        let expected = hex.to_ascii_uppercase();
        let Some(index) = self.run_theme_colors[self.run_theme_index..]
            .iter()
            .position(|color| {
                color
                    .val
                    .as_deref()
                    .is_some_and(|value| value.eq_ignore_ascii_case(&expected))
            })
        else {
            return Vec::new();
        };
        self.run_theme_index += index + 1;
        theme_color_attrs(&self.run_theme_colors[self.run_theme_index - 1])
    }

    fn table(&mut self, table: &docx_rs::Table, depth: usize) -> Result<()> {
        self.components.insert("Table");
        let property = json(&table.property)?;
        let metadata = self.next_table_xml_metadata()?;
        let attrs = table_jsx_attrs(table, &property, &metadata)?;
        self.line(depth, &format!("<Table{attrs}>"));
        for row in &table.rows {
            let TableChild::TableRow(row) = row;
            self.components.insert("TableRow");
            let row_xml = self.next_table_row_xml_metadata()?;
            self.line(
                depth + 1,
                &format!("<TableRow{}>", table_row_jsx_attrs(row, &row_xml)?),
            );
            for cell in &row.cells {
                let TableRowChild::TableCell(cell) = cell;
                self.components.insert("TableCell");
                self.line(
                    depth + 2,
                    &format!("<TableCell{}>", table_cell_jsx_attrs(cell)?),
                );
                self.emit_cell_children(&cell.children, depth + 3)?;
                self.line(depth + 2, "</TableCell>");
            }
            self.line(depth + 1, "</TableRow>");
        }
        self.line(depth, "</Table>");
        Ok(())
    }

    fn next_table_row_xml_metadata(&mut self) -> Result<TableRowXmlMetadata> {
        let part = self
            .image_scope
            .as_deref()
            .unwrap_or("word/document.xml")
            .to_owned();
        let index = self.table_row_indices.entry(part.clone()).or_default();
        let Some(metadata) = self
            .table_row_xml
            .get(&part)
            .and_then(|rows| rows.get(*index))
            .cloned()
        else {
            return Err(Error::Reverse(format!(
                "docx-rs exposed a table row that is missing from `{part}`; repair the table structure or retain the original DOCX"
            )));
        };
        *index += 1;
        Ok(metadata)
    }

    fn next_paragraph_indent_metadata(&mut self) -> Result<ParagraphIndentMetadata> {
        let part = self
            .image_scope
            .as_deref()
            .unwrap_or("word/document.xml")
            .to_owned();
        let index = self
            .paragraph_indent_indices
            .entry(part.clone())
            .or_default();
        let metadata = self
            .paragraph_indent_xml
            .get(&part)
            .and_then(|paragraphs| paragraphs.get(*index))
            .cloned()
            .ok_or_else(|| Error::Reverse(format!(
                "docx-rs exposed a paragraph missing from `{part}` XML metadata; retain the original DOCX or repair the paragraph tree"
            )))?;
        *index += 1;
        Ok(metadata)
    }

    fn next_table_xml_metadata(&mut self) -> Result<TableXmlMetadata> {
        let part = self
            .image_scope
            .as_deref()
            .unwrap_or("word/document.xml")
            .to_owned();
        let index = self.table_indices.entry(part.clone()).or_default();
        let Some(metadata) = self
            .table_xml
            .get(&part)
            .and_then(|tables| tables.get(*index))
            .cloned()
        else {
            return Err(Error::Reverse(format!(
                "docx-rs exposed a table that is missing from `{part}`; repair the table structure or retain the original DOCX"
            )));
        };
        *index += 1;
        Ok(metadata)
    }

    fn validate_table_row_xml_consumed(&self) -> Result<()> {
        for (part, rows) in &self.table_row_xml {
            let consumed = self
                .table_row_indices
                .get(part)
                .copied()
                .unwrap_or_default();
            if consumed != rows.len() {
                return Err(Error::Reverse(format!(
                    "`{part}` contains {} table rows but docx-rs exposed {consumed}; reverse would lose row structure, so repair the table XML or keep the original DOCX",
                    rows.len()
                )));
            }
        }
        Ok(())
    }

    fn validate_table_xml_consumed(&self) -> Result<()> {
        for (part, tables) in &self.table_xml {
            let consumed = self.table_indices.get(part).copied().unwrap_or_default();
            if consumed != tables.len() {
                return Err(Error::Reverse(format!(
                    "`{part}` contains {} tables but docx-rs exposed {consumed}; reverse would lose table structure, so repair the table XML or keep the original DOCX",
                    tables.len()
                )));
            }
        }
        Ok(())
    }

    fn emit_header_children(&mut self, children: &[HeaderChild], depth: usize) -> Result<()> {
        let mut index = 0;
        while index < children.len() {
            match &children[index] {
                HeaderChild::Paragraph(paragraph) if paragraph_num_id(paragraph).is_some() => {
                    let num_id = paragraph_num_id(paragraph).expect("checked");
                    let mut group = vec![paragraph];
                    index += 1;
                    while index < children.len() {
                        if let HeaderChild::Paragraph(next) = &children[index]
                            && paragraph_num_id(next) == Some(num_id)
                        {
                            group.push(next);
                            index += 1;
                        } else {
                            break;
                        }
                    }
                    self.emit_list_from_paragraphs(num_id, group.iter().map(AsRef::as_ref), depth)?;
                }
                HeaderChild::Paragraph(paragraph) => {
                    self.paragraph(paragraph, depth)?;
                    index += 1;
                }
                HeaderChild::Table(table) => {
                    self.table(table, depth)?;
                    index += 1;
                }
                HeaderChild::StructuredDataTag(tag) => {
                    self.block_content_control(tag, depth)?;
                    index += 1;
                }
            }
        }
        Ok(())
    }

    fn emit_footer_children(&mut self, children: &[FooterChild], depth: usize) -> Result<()> {
        let mut index = 0;
        while index < children.len() {
            match &children[index] {
                FooterChild::Paragraph(paragraph) if paragraph_num_id(paragraph).is_some() => {
                    let num_id = paragraph_num_id(paragraph).expect("checked");
                    let mut group = vec![paragraph];
                    index += 1;
                    while index < children.len() {
                        if let FooterChild::Paragraph(next) = &children[index]
                            && paragraph_num_id(next) == Some(num_id)
                        {
                            group.push(next);
                            index += 1;
                        } else {
                            break;
                        }
                    }
                    self.emit_list_from_paragraphs(num_id, group.iter().map(AsRef::as_ref), depth)?;
                }
                FooterChild::Paragraph(paragraph) => {
                    self.footer_paragraph(paragraph, depth)?;
                    index += 1;
                }
                FooterChild::Table(table) => {
                    self.table(table, depth)?;
                    index += 1;
                }
                FooterChild::StructuredDataTag(tag) => {
                    self.block_content_control(tag, depth)?;
                    index += 1;
                }
            }
        }
        Ok(())
    }

    fn emit_sdt_children(
        &mut self,
        children: &[StructuredDataTagChild],
        depth: usize,
    ) -> Result<()> {
        let mut index = 0;
        while index < children.len() {
            match &children[index] {
                StructuredDataTagChild::Paragraph(paragraph)
                    if paragraph_num_id(paragraph).is_some() =>
                {
                    let num_id = paragraph_num_id(paragraph).expect("checked");
                    let mut group = vec![paragraph];
                    index += 1;
                    while index < children.len() {
                        if let StructuredDataTagChild::Paragraph(next) = &children[index]
                            && paragraph_num_id(next) == Some(num_id)
                        {
                            group.push(next);
                            index += 1;
                        } else {
                            break;
                        }
                    }
                    self.emit_list_from_paragraphs(num_id, group.iter().map(AsRef::as_ref), depth)?;
                }
                StructuredDataTagChild::Paragraph(paragraph) => {
                    self.paragraph(paragraph, depth)?;
                    index += 1;
                }
                StructuredDataTagChild::Table(table) => {
                    self.table(table, depth)?;
                    index += 1;
                }
                StructuredDataTagChild::StructuredDataTag(inner) => {
                    self.block_content_control(inner, depth)?;
                    index += 1;
                }
                StructuredDataTagChild::Run(_) => {
                    return Err(Error::Reverse(
                        "block-level structured document tag contains inline runs; wrap them in a paragraph before reversing"
                            .to_owned(),
                    ));
                }
                value => return Err(unsupported("block content control", value)),
            }
        }
        Ok(())
    }

    fn emit_cell_children(&mut self, children: &[TableCellContent], depth: usize) -> Result<()> {
        let mut index = 0;
        while index < children.len() {
            match &children[index] {
                TableCellContent::Paragraph(paragraph) if paragraph_num_id(paragraph).is_some() => {
                    let num_id = paragraph_num_id(paragraph).expect("checked");
                    let mut group = vec![paragraph];
                    index += 1;
                    while index < children.len() {
                        if let TableCellContent::Paragraph(next) = &children[index]
                            && paragraph_num_id(next) == Some(num_id)
                        {
                            group.push(next);
                            index += 1;
                        } else {
                            break;
                        }
                    }
                    self.emit_list_from_paragraphs(num_id, group.iter().map(AsRef::as_ref), depth)?;
                }
                TableCellContent::Paragraph(paragraph) => {
                    self.paragraph(paragraph, depth)?;
                    index += 1;
                }
                TableCellContent::Table(table) => {
                    self.table(table, depth)?;
                    index += 1;
                }
                TableCellContent::StructuredDataTag(tag) => {
                    if let Some(mut kind) = index_from_sdt(tag, &self.deleted_instructions)? {
                        let deletion =
                            deleted_index_field_from_sdt(tag, &self.deleted_instructions)?;
                        let instruction = collect_sdt_instruction(tag, deletion.as_ref());
                        self.append_index_field_paragraph_properties(&mut kind, &instruction)?;
                        if self.apply_sdt {
                            if let Some(props) = self.sdt_plan.cell_tags.get(self.cell_sdt_index) {
                                add_index_sdt_properties(&mut kind, props);
                            }
                            self.cell_sdt_index += 1;
                        }
                        self.index_component(&kind, depth)?;
                    } else {
                        self.cell_content_control(tag, depth)?;
                    }
                    index += 1;
                }
                TableCellContent::TableOfContents(_) => {
                    self.empty("TableOfContents", depth, "");
                    index += 1;
                }
            }
        }
        Ok(())
    }

    fn emit_list_from_paragraphs<'a, I>(
        &mut self,
        num_id: usize,
        paragraphs: I,
        depth: usize,
    ) -> Result<()>
    where
        I: IntoIterator<Item = &'a docx_rs::Paragraph>,
    {
        self.components.insert("List");
        self.components.insert("ListItem");
        let attrs = self.list_jsx_attrs(num_id)?;
        self.line(depth, &format!("<List{attrs}>"));
        for paragraph in paragraphs {
            self.list_item(paragraph, depth + 1)?;
        }
        self.line(depth, "</List>");
        Ok(())
    }

    fn list_jsx_attrs(&self, num_id: usize) -> Result<String> {
        let instance = self
            .numbering_instances
            .iter()
            .find(|numbering| numbering.id == num_id)
            .ok_or_else(|| {
                Error::Reverse(format!(
                    "paragraph references numbering `{num_id}` which is missing from word/numbering.xml; restore the numbering definition or remove `w:numPr`"
                ))
            })?;
        let definition = self
            .abstract_nums
            .iter()
            .find(|value| value.id == instance.abstract_num_id)
            .ok_or_else(|| {
                Error::Reverse(format!(
                    "numbering `{num_id}` points at missing abstract numbering {}; restore word/numbering.xml",
                    instance.abstract_num_id
                ))
            })?;
        if definition.levels.is_empty()
            && definition.style_link.is_none()
            && definition.num_style_link.is_none()
        {
            return Err(Error::Reverse(
                "numbering has no `w:lvl`, `w:styleLink`, or `w:numStyleLink` definition; restore a numbering definition in word/numbering.xml"
                    .to_owned(),
            ));
        }
        let mut attrs = reverse_list_definition(
            definition,
            self.numbering_multi_level_types
                .get(&definition.id)
                .map(String::as_str),
        )?;
        if !instance.level_overrides.is_empty() {
            let mut overrides = Vec::with_capacity(instance.level_overrides.len());
            for override_level in &instance.level_overrides {
                let mut value = Map::new();
                value.insert("level".to_owned(), Value::from(override_level.level));
                if let Some(start) = override_level.override_start {
                    value.insert("start".to_owned(), Value::from(start));
                }
                if let Some(definition) = &override_level.override_level {
                    value.insert(
                        "definition".to_owned(),
                        reverse_level_definition(definition)?,
                    );
                }
                overrides.push(Value::Object(value));
            }
            let encoded = serde_json::to_string(&overrides).map_err(|error| {
                Error::Reverse(format!("cannot serialize list overrides: {error}"))
            })?;
            let _ = write!(attrs, " overrides={{{encoded}}}");
        }
        Ok(attrs)
    }

    fn list_item(&mut self, paragraph: &docx_rs::Paragraph, depth: usize) -> Result<()> {
        let raw_indent = self.next_paragraph_indent_metadata()?;
        let property = json(&paragraph.property)?;
        let mut attrs = String::new();
        if let Some(value) = raw_indent.suppress_line_numbers {
            attrs.push_str(&jsx_prop("suppressLineNumbers", &Value::Bool(value))?);
        }
        if let Some(level) =
            nested(&property, &["numberingProperty", "level"]).and_then(Value::as_u64)
            && level > 0
        {
            let _ = write!(attrs, " level={{{level}}}");
        }
        if let Some(value) = property
            .get("style")
            .and_then(scalar_string)
            .or_else(|| nested_string(&property, &["style", "val"]))
            .or_else(|| nested_string(&property, &["style", "styleId"]))
        {
            attrs.push_str(&attr("style", value));
        }
        for (name, value) in &raw_indent.values {
            attrs.push_str(&jsx_prop(name, &Value::from(*value))?);
        }
        self.line(depth, &format!("<ListItem{attrs}>"));
        if self.apply_sdt {
            self.paragraph_index += 1;
        }
        let events = if self.apply_sdt {
            self.sdt_plan
                .paragraphs
                .get(self.paragraph_index - 1)
                .cloned()
                .unwrap_or_default()
        } else {
            Vec::new()
        };
        let nested = nest_inline_children(&paragraph.children, &self.comments, &events)?;
        for child in nested {
            self.emit_nested_inline(&child, depth + 1)?;
        }
        self.line(depth, "</ListItem>");
        Ok(())
    }

    fn empty(&mut self, component: &'static str, depth: usize, attrs: &str) {
        self.components.insert(component);
        self.line(depth, &format!("<{component}{attrs} />"));
    }
}

#[derive(Clone, Debug, Default)]
struct SdtProps {
    alias: Option<String>,
    xpath: Option<String>,
    prefix_mappings: Option<String>,
    store_item_id: Option<String>,
}

#[derive(Clone, Debug)]
enum SdtEvent {
    Start(SdtProps),
    Leaf,
    End,
}

#[derive(Clone, Debug, Default)]
struct SdtPlan {
    paragraphs: Vec<Vec<SdtEvent>>,
    hyperlinks: Vec<Vec<SdtEvent>>,
    body_tags: Vec<SdtProps>,
    cell_tags: Vec<SdtProps>,
}

enum NestedBlock<'a> {
    Child(&'a DocumentChild),
    Bookmark {
        name: &'a str,
        children: Vec<NestedBlock<'a>>,
    },
    ContentControl {
        tag: &'a StructuredDataTag,
    },
}

enum NestedInline<'a> {
    Child(&'a ParagraphChild),
    InlineBookmark {
        name: &'a str,
        children: Vec<NestedInline<'a>>,
    },
    Comment {
        comment: Comment,
        children: Vec<NestedInline<'a>>,
    },
    ContentControl {
        props: SdtProps,
        children: Vec<NestedInline<'a>>,
    },
}

enum DeletedInline<'a> {
    Run(&'a docx_rs::Run),
    Comment {
        comment: Comment,
        children: Vec<DeletedInline<'a>>,
    },
}

enum RangeFrame<'a> {
    Bookmark {
        id: usize,
        name: &'a str,
        children: Vec<NestedInline<'a>>,
    },
    Comment {
        id: usize,
        comment: Comment,
        children: Vec<NestedInline<'a>>,
    },
}

fn nest_block_children(children: &[DocumentChild]) -> Result<Vec<NestedBlock<'_>>> {
    let mut root = Vec::new();
    let mut stack: Vec<(usize, &str, Vec<NestedBlock<'_>>)> = Vec::new();
    let mut ignored = BTreeSet::new();
    for child in children {
        match child {
            DocumentChild::BookmarkStart(value) if value.name == "_GoBack" => {
                ignored.insert(value.id);
            }
            DocumentChild::BookmarkStart(value) => {
                stack.push((value.id, value.name.as_str(), Vec::new()));
            }
            DocumentChild::BookmarkEnd(value) if ignored.remove(&value.id) => {}
            DocumentChild::BookmarkEnd(value) => {
                let Some((start_id, name, children)) = stack.pop() else {
                    return Err(range_error(
                        "unmatched bookmark end",
                        value.id,
                        "pair each `w:bookmarkStart` with a matching `w:bookmarkEnd` that nests, not overlaps",
                    ));
                };
                if start_id != value.id {
                    return Err(range_error(
                        "overlapping bookmark ranges",
                        value.id,
                        "close the inner bookmark before the outer one so ranges nest as JSX",
                    ));
                }
                push_block(
                    &mut stack,
                    &mut root,
                    NestedBlock::Bookmark { name, children },
                );
            }
            DocumentChild::CommentStart(_) | DocumentChild::CommentEnd(_) => {
                return Err(Error::Reverse(
                    "document-level comment ranges cannot be represented as JSX Comment; move the commented content into a Paragraph"
                        .to_owned(),
                ));
            }
            DocumentChild::StructuredDataTag(tag) => {
                push_block(&mut stack, &mut root, NestedBlock::ContentControl { tag });
            }
            value => push_block(&mut stack, &mut root, NestedBlock::Child(value)),
        }
    }
    if let Some((id, _, _)) = stack.last() {
        return Err(range_error(
            "unmatched bookmark start",
            *id,
            "add the matching `w:bookmarkEnd` or remove the orphan start marker",
        ));
    }
    Ok(root)
}

fn push_block<'a>(
    stack: &mut Vec<(usize, &str, Vec<NestedBlock<'a>>)>,
    root: &mut Vec<NestedBlock<'a>>,
    child: NestedBlock<'a>,
) {
    if let Some((_, _, children)) = stack.last_mut() {
        children.push(child);
    } else {
        root.push(child);
    }
}

fn nest_inline_children<'a>(
    children: &'a [ParagraphChild],
    comments: &HashMap<usize, Comment>,
    events: &[SdtEvent],
) -> Result<Vec<NestedInline<'a>>> {
    let has_native_sdt = children
        .iter()
        .any(|child| matches!(child, ParagraphChild::StructuredDataTag(_)));
    let leaves = children.iter().map(NestedInline::Child).collect();
    let wrapped = if has_native_sdt || events.is_empty() {
        leaves
    } else {
        apply_sdt_events(events, leaves)?
    };
    nest_inline_ranges(wrapped, comments)
}

fn apply_sdt_events<'a>(
    events: &[SdtEvent],
    children: Vec<NestedInline<'a>>,
) -> Result<Vec<NestedInline<'a>>> {
    let mut remaining = children.into_iter();
    let mut root = Vec::new();
    let mut stack: Vec<(SdtProps, Vec<NestedInline<'a>>)> = Vec::new();
    for event in events {
        match event {
            SdtEvent::Start(props) => stack.push((props.clone(), Vec::new())),
            SdtEvent::End => {
                let Some((props, children)) = stack.pop() else {
                    return Err(Error::Reverse(
                        "unmatched structured document tag end; keep each `w:sdt` well-formed"
                            .to_owned(),
                    ));
                };
                push_inline(
                    &mut stack,
                    &mut root,
                    NestedInline::ContentControl { props, children },
                );
            }
            SdtEvent::Leaf => {
                let Some(child) = remaining.next() else {
                    return Err(Error::Reverse(
                        "structured document tag markers do not match readable paragraph children; simplify nested `w:sdt` content"
                            .to_owned(),
                    ));
                };
                push_inline_frame(&mut stack, &mut root, child);
            }
        }
    }
    if stack.last().is_some() {
        return Err(Error::Reverse(
            "unmatched structured document tag start; close each `w:sdt` before reversing"
                .to_owned(),
        ));
    }
    root.extend(remaining);
    Ok(root)
}

fn push_inline<'a>(
    stack: &mut Vec<(SdtProps, Vec<NestedInline<'a>>)>,
    root: &mut Vec<NestedInline<'a>>,
    child: NestedInline<'a>,
) {
    push_inline_frame(stack, root, child);
}

fn push_inline_frame<'a>(
    stack: &mut Vec<(SdtProps, Vec<NestedInline<'a>>)>,
    root: &mut Vec<NestedInline<'a>>,
    child: NestedInline<'a>,
) {
    if let Some((_, children)) = stack.last_mut() {
        children.push(child);
    } else {
        root.push(child);
    }
}

fn nest_inline_ranges<'a>(
    children: Vec<NestedInline<'a>>,
    comments: &HashMap<usize, Comment>,
) -> Result<Vec<NestedInline<'a>>> {
    let mut root = Vec::new();
    let mut stack: Vec<RangeFrame<'a>> = Vec::new();
    let mut ignored = BTreeSet::new();
    for child in children {
        apply_inline_range(child, comments, &mut stack, &mut root, &mut ignored)?;
    }
    if let Some(frame) = stack.last() {
        let id = match frame {
            RangeFrame::Bookmark { id, .. } | RangeFrame::Comment { id, .. } => *id,
        };
        return Err(range_error(
            "unmatched range start",
            id,
            "add the matching end marker or remove the orphan start so the range can nest",
        ));
    }
    Ok(root)
}

fn apply_inline_range<'a>(
    child: NestedInline<'a>,
    comments: &HashMap<usize, Comment>,
    stack: &mut Vec<RangeFrame<'a>>,
    root: &mut Vec<NestedInline<'a>>,
    ignored: &mut BTreeSet<usize>,
) -> Result<()> {
    let NestedInline::Child(value) = child else {
        push_range_child(stack, root, child);
        return Ok(());
    };
    match value {
        ParagraphChild::BookmarkStart(start) if start.name == "_GoBack" => {
            ignored.insert(start.id);
        }
        ParagraphChild::BookmarkStart(start) => {
            stack.push(RangeFrame::Bookmark {
                id: start.id,
                name: start.name.as_str(),
                children: Vec::new(),
            });
        }
        ParagraphChild::BookmarkEnd(end) if ignored.remove(&end.id) => {}
        ParagraphChild::BookmarkEnd(end) => close_bookmark_range(stack, root, end.id)?,
        ParagraphChild::CommentStart(start) => {
            stack.push(RangeFrame::Comment {
                id: start.id,
                comment: resolve_comment(start.id, &start.comment, comments)?,
                children: Vec::new(),
            });
        }
        ParagraphChild::CommentEnd(end) => {
            close_comment_range(stack, root, comment_end_id(end)?)?;
        }
        _ => push_range_child(stack, root, NestedInline::Child(value)),
    }
    Ok(())
}

fn push_range_child<'a>(
    stack: &mut Vec<RangeFrame<'a>>,
    root: &mut Vec<NestedInline<'a>>,
    child: NestedInline<'a>,
) {
    match stack.last_mut() {
        Some(RangeFrame::Bookmark { children, .. } | RangeFrame::Comment { children, .. }) => {
            children.push(child);
        }
        None => root.push(child),
    }
}

fn close_bookmark_range<'a>(
    stack: &mut Vec<RangeFrame<'a>>,
    root: &mut Vec<NestedInline<'a>>,
    end_id: usize,
) -> Result<()> {
    match stack.pop() {
        Some(RangeFrame::Bookmark { id, name, children }) if id == end_id => {
            push_range_child(stack, root, NestedInline::InlineBookmark { name, children });
            Ok(())
        }
        Some(RangeFrame::Bookmark { .. } | RangeFrame::Comment { .. }) => Err(range_error(
            "overlapping bookmark or comment ranges",
            end_id,
            "close the inner range before the outer one so bookmarks and comments nest as JSX",
        )),
        None => Err(range_error(
            "unmatched bookmark end",
            end_id,
            "pair each `w:bookmarkStart` with a matching `w:bookmarkEnd` that nests, not overlaps",
        )),
    }
}

fn close_comment_range<'a>(
    stack: &mut Vec<RangeFrame<'a>>,
    root: &mut Vec<NestedInline<'a>>,
    end_id: usize,
) -> Result<()> {
    match stack.pop() {
        Some(RangeFrame::Comment {
            id,
            comment,
            children,
        }) if id == end_id => {
            push_range_child(stack, root, NestedInline::Comment { comment, children });
            Ok(())
        }
        Some(RangeFrame::Bookmark { .. } | RangeFrame::Comment { .. }) => Err(range_error(
            "overlapping bookmark or comment ranges",
            end_id,
            "close the inner range before the outer one so bookmarks and comments nest as JSX",
        )),
        None => Err(range_error(
            "unmatched comment end",
            end_id,
            "pair each `w:commentRangeStart` with a matching `w:commentRangeEnd`",
        )),
    }
}

fn resolve_comment(
    id: usize,
    attached: &Comment,
    comments: &HashMap<usize, Comment>,
) -> Result<Comment> {
    if !attached.children.is_empty() {
        return Ok(attached.clone());
    }
    comments.get(&id).cloned().ok_or_else(|| {
        Error::Reverse(format!(
            "comment {id} is missing from comments.xml; add the comment part or remove the orphan comment range"
        ))
    })
}

fn comment_end_id(value: &docx_rs::CommentRangeEnd) -> Result<usize> {
    json(value)?
        .get("id")
        .and_then(Value::as_u64)
        .and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| Error::Reverse("comment range end is missing an id".to_owned()))
}

fn nest_deleted_ranges<'a>(
    children: &'a [DeleteChild],
    comments: &HashMap<usize, Comment>,
) -> Result<Vec<DeletedInline<'a>>> {
    let mut root = Vec::new();
    let mut stack: Vec<(usize, Comment, Vec<DeletedInline<'a>>)> = Vec::new();
    for child in children {
        match child {
            DeleteChild::Run(run) => {
                push_deleted_child(&mut stack, &mut root, DeletedInline::Run(run.as_ref()));
            }
            DeleteChild::CommentStart(start) => stack.push((
                start.id,
                resolve_comment(start.id, &start.comment, comments)?,
                Vec::new(),
            )),
            DeleteChild::CommentEnd(end) => {
                let end_id = comment_end_id(end)?;
                match stack.pop() {
                    Some((id, comment, nested)) if id == end_id => push_deleted_child(
                        &mut stack,
                        &mut root,
                        DeletedInline::Comment {
                            comment,
                            children: nested,
                        },
                    ),
                    Some((id, _, _)) => {
                        return Err(range_error(
                            "overlapping comment ranges in deleted text",
                            end_id.max(id),
                            "close the inner comment range before the outer one so deleted comments nest as JSX",
                        ));
                    }
                    None => {
                        return Err(range_error(
                            "unmatched comment end in deleted text",
                            end_id,
                            "pair each `w:commentRangeStart` with a matching `w:commentRangeEnd` inside the deleted revision",
                        ));
                    }
                }
            }
        }
    }
    if let Some((id, _, _)) = stack.last() {
        return Err(range_error(
            "unmatched comment start in deleted text",
            *id,
            "add the matching `w:commentRangeEnd` inside the deleted revision or remove the orphan marker",
        ));
    }
    Ok(root)
}

fn push_deleted_child<'a>(
    stack: &mut Vec<(usize, Comment, Vec<DeletedInline<'a>>)>,
    root: &mut Vec<DeletedInline<'a>>,
    child: DeletedInline<'a>,
) {
    if let Some((_, _, children)) = stack.last_mut() {
        children.push(child);
    } else {
        root.push(child);
    }
}

fn comment_text(comment: &Comment) -> Result<String> {
    let mut text = String::new();
    for child in &comment.children {
        match child {
            CommentChild::Paragraph(paragraph) => {
                for child in &paragraph.children {
                    let ParagraphChild::Run(run) = child else {
                        return Err(Error::Reverse(
                            "comment body contains non-run content that JSX Comment `text` cannot represent; keep comment text as a single paragraph"
                                .to_owned(),
                        ));
                    };
                    for child in &run.children {
                        if let RunChild::Text(value) = child {
                            text.push_str(&value.text);
                        }
                    }
                }
            }
            CommentChild::Table(_) => {
                return Err(Error::Reverse(
                    "comment body tables are not representable as JSX Comment `text`; flatten the comment to a paragraph"
                        .to_owned(),
                ));
            }
        }
    }
    if text.is_empty() {
        return Err(Error::Reverse(format!(
            "comment {} has empty text; store the annotation in comments.xml or remove the range",
            comment.id()
        )));
    }
    Ok(text)
}

fn comment_identity_attrs(comment: &Comment, done: Option<bool>) -> String {
    let mut attributes = format!(" id={{{}}}", comment.id);
    if let Some(parent) = comment.parent_comment_id {
        let _ = write!(attributes, " parentCommentId={{{parent}}}");
    }
    if let Some(done) = done {
        let _ = write!(attributes, " done={{{done}}}");
    }
    attributes
}

fn comment_done(comment: &Comment, states: &HashMap<String, bool>) -> Option<bool> {
    comment.children.iter().find_map(|child| match child {
        CommentChild::Paragraph(paragraph) => states.get(&paragraph.id).copied(),
        CommentChild::Table(_) => None,
    })
}

fn range_error(kind: &str, id: usize, suggestion: &str) -> Error {
    Error::Reverse(format!("{kind} id={id}: {suggestion}"))
}

fn run_is_ignorable(run: &docx_rs::Run) -> bool {
    run.children
        .iter()
        .all(|child| matches!(child, RunChild::FieldChar(_) | RunChild::InstrTextString(_)))
}

fn sdt_props_from_tag(tag: &StructuredDataTag) -> SdtProps {
    SdtProps {
        alias: tag.property.alias.clone(),
        xpath: tag
            .property
            .data_binding
            .as_ref()
            .and_then(|binding| binding.xpath.clone()),
        prefix_mappings: tag
            .property
            .data_binding
            .as_ref()
            .and_then(|binding| binding.prefix_mappings.clone()),
        store_item_id: tag
            .property
            .data_binding
            .as_ref()
            .and_then(|binding| binding.store_item_id.clone()),
    }
}

fn merge_sdt_props(target: &mut SdtProps, source: &SdtProps) {
    if target.alias.is_none() {
        target.alias.clone_from(&source.alias);
    }
    if target.xpath.is_none() {
        target.xpath.clone_from(&source.xpath);
    }
    if target.prefix_mappings.is_none() {
        target.prefix_mappings.clone_from(&source.prefix_mappings);
    }
    if target.store_item_id.is_none() {
        target.store_item_id.clone_from(&source.store_item_id);
    }
}

fn sdt_attrs(props: &SdtProps) -> String {
    let mut attrs = String::new();
    if let Some(value) = &props.alias {
        attrs.push_str(&attr("alias", value));
    }
    if let Some(value) = &props.xpath {
        attrs.push_str(&attr("xpath", value));
    }
    if let Some(value) = &props.prefix_mappings {
        attrs.push_str(&attr("prefixMappings", value));
    }
    if let Some(value) = &props.store_item_id {
        attrs.push_str(&attr("storeItemId", value));
    }
    attrs
}

enum IndexField {
    Contents {
        start: u64,
        end: u64,
        attributes: String,
        items: Vec<TocItemField>,
        before: Vec<IndexContentBlock>,
        after: Vec<IndexContentBlock>,
    },
    Figures {
        label: String,
        include_label_and_number: bool,
        separator: Option<String>,
        attributes: String,
        items: Vec<TocItemField>,
        before: Vec<IndexContentBlock>,
        after: Vec<IndexContentBlock>,
    },
    Entries {
        identifier: String,
        attributes: String,
        items: Vec<TocItemField>,
        before: Vec<IndexContentBlock>,
        after: Vec<IndexContentBlock>,
    },
}

enum IndexContentBlock {
    Paragraph(Box<docx_rs::Paragraph>),
    Table(Box<docx_rs::Table>),
}

#[derive(Debug)]
struct TocItemField {
    text: String,
    bookmark: String,
    level: u64,
    page_ref: String,
    tab_position: u32,
}

#[derive(Debug)]
struct DeletedIndexField {
    paragraph_index: usize,
    author: String,
    date: String,
    instruction: String,
}

fn index_from_sdt(
    tag: &StructuredDataTag,
    deleted_instructions: &HashMap<String, String>,
) -> Result<Option<IndexField>> {
    let deletion = deleted_index_field_from_sdt(tag, deleted_instructions)?;
    let instruction = collect_sdt_instruction(tag, deletion.as_ref());
    let Some(mut index) = parse_toc_instruction(&instruction)? else {
        return Ok(None);
    };
    let (items, before, after) = toc_content_from_sdt(tag, deletion.as_ref())?;
    if deletion.is_some() && !items.is_empty() {
        return Err(Error::Reverse(
            "deleted index field also contains static PAGEREF rows, but docx-rs drops the deletion when compiling rows; remove the rows or keep the field automatic before reversing".to_owned(),
        ));
    }
    match &mut index {
        IndexField::Contents { items: target, .. }
        | IndexField::Figures { items: target, .. }
        | IndexField::Entries { items: target, .. } => *target = items,
    }
    match &mut index {
        IndexField::Contents {
            before: target_before,
            after: target_after,
            ..
        }
        | IndexField::Figures {
            before: target_before,
            after: target_after,
            ..
        }
        | IndexField::Entries {
            before: target_before,
            after: target_after,
            ..
        } => {
            *target_before = before;
            *target_after = after;
        }
    }
    let attributes = match &mut index {
        IndexField::Contents { attributes, .. }
        | IndexField::Figures { attributes, .. }
        | IndexField::Entries { attributes, .. } => attributes,
    };
    if let Some(alias) = &tag.property.alias {
        attributes.push_str(&attr("alias", alias));
    }
    if let Some(dirty) = sdt_field_dirty(tag) {
        if dirty {
            attributes.push_str(" dirty");
        } else {
            attributes.push_str(" dirty={false}");
        }
    }
    if let Some(deletion) = &deletion {
        let value = serde_json::to_string(&serde_json::json!({
            "author": deletion.author,
            "date": deletion.date,
        }))
        .map_err(|error| {
            Error::Reverse(format!("cannot encode index deletion metadata: {error}"))
        })?;
        let _ = write!(attributes, " deletion={{{value}}}");
    }
    Ok(Some(index))
}

fn append_index_field_paragraph_properties(
    index: &mut IndexField,
    mut properties: Map<String, Value>,
) -> Result<()> {
    if index_has_static_items(index) {
        let static_row_only_properties = properties
            .keys()
            .all(|key| key == "style" || key == "__unsupported")
            && properties
                .get("__unsupported")
                .and_then(Value::as_array)
                .is_none_or(|values| {
                    !values.is_empty()
                        && values
                            .iter()
                            .all(|value| matches!(value.as_str(), Some("tab" | "tabs")))
                });
        if properties.is_empty() || static_row_only_properties {
            return Ok(());
        }
        return Err(Error::Reverse(
            "index field paragraph formatting is present together with static PAGEREF rows, but docx-rs ignores paragraph_property in its static-items branch; remove the rows or keep the automatic field before reversing".to_owned(),
        ));
    }
    if let Some(unsupported) = properties.remove("__unsupported") {
        let property = unsupported
            .as_array()
            .and_then(|values| values.first())
            .and_then(Value::as_str)
            .unwrap_or("unknown OOXML paragraph formatting");
        return Err(unsupported_index_field_paragraph_property(property));
    }
    for key in ["keepNext", "keepLines", "pageBreakBefore"] {
        if properties.get(key) == Some(&Value::Bool(false)) {
            return Err(unsupported_index_field_paragraph_property(key));
        }
    }
    if properties.is_empty() {
        return Ok(());
    }
    let value = serde_json::to_string(&properties).map_err(|error| {
        Error::Reverse(format!(
            "cannot encode index field paragraph properties: {error}"
        ))
    })?;
    let attributes = match index {
        IndexField::Contents { attributes, .. }
        | IndexField::Figures { attributes, .. }
        | IndexField::Entries { attributes, .. } => attributes,
    };
    let _ = write!(attributes, " fieldParagraph={{{value}}}");
    Ok(())
}

fn index_has_static_items(index: &IndexField) -> bool {
    match index {
        IndexField::Contents { items, .. }
        | IndexField::Figures { items, .. }
        | IndexField::Entries { items, .. } => !items.is_empty(),
    }
}

fn unsupported_index_field_paragraph_property(property: &str) -> Error {
    Error::Reverse(format!(
        "index field paragraph property `{property}` is not represented by `fieldParagraph`; remove it or use only the documented `fieldParagraph` properties before reversing"
    ))
}

fn deleted_index_field_from_sdt(
    tag: &StructuredDataTag,
    deleted_instructions: &HashMap<String, String>,
) -> Result<Option<DeletedIndexField>> {
    let mut result = None;
    for (paragraph_index, child) in tag.children.iter().enumerate() {
        let StructuredDataTagChild::Paragraph(paragraph) = child else {
            continue;
        };
        for child in &paragraph.children {
            let ParagraphChild::Delete(deletion) = child else {
                continue;
            };
            let mut instruction_marker = None;
            for child in &deletion.children {
                let DeleteChild::Run(run) = child else {
                    continue;
                };
                for child in &run.children {
                    let RunChild::DeleteText(text) = child else {
                        continue;
                    };
                    let data = json(text)?;
                    let Some(marker) = data.get("text").and_then(Value::as_str) else {
                        continue;
                    };
                    if deleted_instructions
                        .get(marker)
                        .is_some_and(|value| value.trim_start().starts_with("TOC"))
                    {
                        instruction_marker = Some((marker.to_owned(), data));
                    }
                }
            }
            let Some((marker, _)) = instruction_marker else {
                continue;
            };
            let Some(instruction) = deleted_instructions.get(&marker).cloned() else {
                continue;
            };
            let exact_field_shape = deletion.children.len() == 1
                && matches!(
                    deletion.children.first(),
                    Some(DeleteChild::Run(run))
                        if run.children.len() == 3
                            && matches!(run.children.first(), Some(RunChild::FieldChar(field)) if field.field_char_type == docx_rs::FieldCharType::Begin)
                            && matches!(run.children.get(1), Some(RunChild::DeleteText(text)) if json(text).ok().and_then(|value| value.get("text").and_then(Value::as_str).map(str::to_owned)).as_deref() == Some(marker.as_str()))
                            && matches!(run.children.get(2), Some(RunChild::FieldChar(field)) if field.field_char_type == docx_rs::FieldCharType::Separate)
                );
            if !exact_field_shape {
                return Err(Error::Reverse(
                    "TOC instruction inside `w:del` has additional revision content that the index DSL cannot represent; keep only the field begin, instruction, and separator in the deletion".to_owned(),
                ));
            }
            if result.is_some() {
                return Err(Error::Reverse(
                    "index control contains multiple deleted TOC field instructions; keep one deleted field per index control".to_owned(),
                ));
            }
            result = Some(DeletedIndexField {
                paragraph_index,
                author: deletion.author.clone(),
                date: deletion.date.clone(),
                instruction,
            });
        }
    }
    Ok(result)
}

fn unwrapped_index_at(
    children: &[NestedBlock<'_>],
    index: usize,
) -> Result<Option<(IndexField, usize)>> {
    let Some(NestedBlock::Child(DocumentChild::Paragraph(begin))) = children.get(index) else {
        return Ok(None);
    };
    let Some(NestedBlock::Child(DocumentChild::Paragraph(end))) = children.get(index + 1) else {
        return Ok(None);
    };
    if !paragraph_has_only_field_end(end) {
        return Ok(None);
    }
    let begin_flags = toc_field_paragraph_flags(begin);
    let end_flags = toc_field_paragraph_flags(end);
    if !begin_flags.0 || begin_flags.1 || begin_flags.2 || !end_flags.1 || end_flags.0 {
        return Ok(None);
    }
    let Some(instruction) = toc_instruction_from_paragraph(begin) else {
        return Ok(None);
    };
    let Some(mut field) = parse_toc_instruction(&instruction)? else {
        return Ok(None);
    };
    let attributes = match &mut field {
        IndexField::Contents { attributes, .. }
        | IndexField::Figures { attributes, .. }
        | IndexField::Entries { attributes, .. } => attributes,
    };
    attributes.push_str(" withoutSdt");
    if let Some(dirty) = paragraph_field_dirty(begin) {
        if dirty {
            attributes.push_str(" dirty");
        } else {
            attributes.push_str(" dirty={false}");
        }
    }
    Ok(Some((field, 2)))
}

fn toc_instruction_from_paragraph(paragraph: &docx_rs::Paragraph) -> Option<String> {
    let mut started = false;
    let mut in_instruction = false;
    let mut separated = false;
    let mut instruction = String::new();
    for child in &paragraph.children {
        let ParagraphChild::Run(run) = child else {
            return None;
        };
        for child in &run.children {
            match child {
                RunChild::FieldChar(field)
                    if field.field_char_type == docx_rs::FieldCharType::Begin
                        && !started
                        && !separated =>
                {
                    started = true;
                    in_instruction = true;
                }
                RunChild::InstrTextString(value) if in_instruction => {
                    instruction.push_str(value);
                }
                RunChild::FieldChar(field)
                    if field.field_char_type == docx_rs::FieldCharType::Separate
                        && in_instruction
                        && !separated =>
                {
                    in_instruction = false;
                    separated = true;
                }
                _ => return None,
            }
        }
    }
    (started && separated && !instruction.is_empty()).then_some(instruction)
}

fn paragraph_has_only_field_end(paragraph: &docx_rs::Paragraph) -> bool {
    let mut end_count = 0;
    for child in &paragraph.children {
        let ParagraphChild::Run(run) = child else {
            return false;
        };
        for child in &run.children {
            if matches!(child, RunChild::FieldChar(field) if field.field_char_type == docx_rs::FieldCharType::End)
            {
                end_count += 1;
            } else {
                return false;
            }
        }
    }
    end_count == 1
}

fn paragraph_field_dirty(paragraph: &docx_rs::Paragraph) -> Option<bool> {
    paragraph.children.iter().find_map(|child| {
        let ParagraphChild::Run(run) = child else {
            return None;
        };
        run.children.iter().find_map(|child| {
            let RunChild::FieldChar(field) = child else {
                return None;
            };
            (field.field_char_type == docx_rs::FieldCharType::Begin).then_some(field.dirty)
        })
    })
}

fn toc_content_from_sdt(
    tag: &StructuredDataTag,
    deletion: Option<&DeletedIndexField>,
) -> Result<(
    Vec<TocItemField>,
    Vec<IndexContentBlock>,
    Vec<IndexContentBlock>,
)> {
    let mut items = Vec::new();
    let mut before = Vec::new();
    let mut after = Vec::new();
    let mut field_started = false;
    let mut field_ended = false;
    for (index, child) in tag.children.iter().enumerate() {
        match child {
            StructuredDataTagChild::Paragraph(paragraph) => {
                if deletion.is_some_and(|deletion| deletion.paragraph_index == index) {
                    field_started = true;
                    continue;
                }
                let flags = toc_field_paragraph_flags(paragraph);
                if flags.2 && (flags.0 || (field_started && !field_ended)) {
                    let item = toc_item_from_paragraph(paragraph)?.ok_or_else(|| {
                        Error::Reverse(
                            "PAGEREF field could not be restored as a static index row".to_owned(),
                        )
                    })?;
                    items.push(item);
                    field_started |= flags.0;
                    continue;
                }
                if flags.0 {
                    field_started = true;
                    continue;
                }
                if flags.1 {
                    field_ended = true;
                    if let Some(paragraph) = paragraph_without_field_end(paragraph) {
                        after.push(IndexContentBlock::Paragraph(Box::new(paragraph)));
                    }
                    continue;
                }
                if field_started && !field_ended {
                    return Err(Error::Reverse(
                        "index field contains a cached result paragraph that is not a representable PAGEREF row; add one PAGEREF target per row or remove the cached result before reversing"
                            .to_owned(),
                    ));
                } else if field_ended {
                    let next_is_table = matches!(
                        tag.children.get(index + 1),
                        Some(StructuredDataTagChild::Table(_))
                    );
                    if !(after.is_empty() && next_is_table && paragraph_is_empty(paragraph)) {
                        after.push(IndexContentBlock::Paragraph(Box::new(
                            (**paragraph).clone(),
                        )));
                    }
                } else {
                    before.push(IndexContentBlock::Paragraph(Box::new(
                        (**paragraph).clone(),
                    )));
                }
            }
            StructuredDataTagChild::Table(table) => {
                if field_started && !field_ended {
                    return Err(Error::Reverse(
                        "table appears inside the cached result of an index field; move it before or after the field, or use a PAGEREF row instead"
                            .to_owned(),
                    ));
                }
                if field_ended {
                    after.push(IndexContentBlock::Table(Box::new((**table).clone())));
                } else {
                    before.push(IndexContentBlock::Table(Box::new((**table).clone())));
                }
            }
            StructuredDataTagChild::StructuredDataTag(_)
            | StructuredDataTagChild::Run(_)
            | StructuredDataTagChild::BookmarkStart(_)
            | StructuredDataTagChild::BookmarkEnd(_)
            | StructuredDataTagChild::CommentStart(_)
            | StructuredDataTagChild::CommentEnd(_) => {
                return Err(Error::Reverse(
                    "non-paragraph content inside an index control cannot be represented; keep static entries as `ToCItem` rows and remove unsupported controls or markers"
                        .to_owned(),
                ));
            }
        }
    }
    Ok((items, before, after))
}

fn paragraph_is_empty(paragraph: &docx_rs::Paragraph) -> bool {
    let mut runs = Vec::new();
    let mut anchors = Vec::new();
    collect_toc_item_runs(&paragraph.children, &mut runs, &mut anchors).is_ok()
        && runs.iter().all(|run| {
            run.children
                .iter()
                .all(|child| matches!(child, RunChild::Text(text) if text.text.is_empty()))
        })
}

fn toc_field_paragraph_flags(paragraph: &docx_rs::Paragraph) -> (bool, bool, bool) {
    let mut runs = Vec::new();
    let mut anchors = Vec::new();
    if collect_toc_item_runs(&paragraph.children, &mut runs, &mut anchors).is_err() {
        return (false, false, false);
    }
    let mut has_toc_instruction = false;
    let mut has_field_end = false;
    let mut has_page_reference = false;
    for run in runs {
        for child in &run.children {
            match child {
                RunChild::InstrTextString(instruction)
                    if instruction.trim_start().starts_with("TOC") =>
                {
                    has_toc_instruction = true;
                }
                RunChild::InstrTextString(instruction)
                    if instruction.trim_start().starts_with("PAGEREF") =>
                {
                    has_page_reference = true;
                }
                RunChild::FieldChar(field)
                    if field.field_char_type == docx_rs::FieldCharType::End =>
                {
                    has_field_end = true;
                }
                _ => {}
            }
        }
    }
    (has_toc_instruction, has_field_end, has_page_reference)
}

fn paragraph_without_field_end(paragraph: &docx_rs::Paragraph) -> Option<docx_rs::Paragraph> {
    let mut paragraph = paragraph.clone();
    for child in &mut paragraph.children {
        let ParagraphChild::Run(run) = child else {
            continue;
        };
        run.children.retain(|child| {
            !matches!(child, RunChild::FieldChar(field) if field.field_char_type == docx_rs::FieldCharType::End)
        });
    }
    paragraph
        .children
        .retain(|child| !matches!(child, ParagraphChild::Run(run) if run.children.is_empty()));
    (!paragraph.children.is_empty()).then_some(paragraph)
}

fn toc_item_from_paragraph(paragraph: &docx_rs::Paragraph) -> Result<Option<TocItemField>> {
    let mut runs = Vec::new();
    let mut anchors = Vec::new();
    collect_toc_item_runs(&paragraph.children, &mut runs, &mut anchors)?;
    let mut text = String::new();
    let mut bookmark = None;
    let mut page_ref = String::new();
    let mut in_page_ref = false;
    let mut saw_page_ref_text = false;
    for run in runs {
        for child in &run.children {
            match child {
                RunChild::InstrTextString(instruction)
                    if instruction.trim_start().starts_with("PAGEREF") =>
                {
                    if bookmark.is_some() {
                        return Err(Error::Reverse(
                            "index row contains multiple PAGEREF fields; split it into one entry per paragraph"
                                .to_owned(),
                        ));
                    }
                    bookmark = Some(parse_toc_item_pageref(instruction)?);
                    in_page_ref = true;
                }
                RunChild::InstrTextString(instruction)
                    if instruction.trim_start().starts_with("TOC") => {}
                RunChild::InstrTextString(instruction) => {
                    return Err(Error::Reverse(format!(
                        "static index row contains unsupported field instruction `{instruction}`; use the TOC field and one PAGEREF per entry"
                    )));
                }
                RunChild::FieldChar(field)
                    if field.field_char_type == docx_rs::FieldCharType::End && in_page_ref =>
                {
                    in_page_ref = false;
                }
                RunChild::Text(value) if in_page_ref => {
                    page_ref.push_str(&value.text);
                    saw_page_ref_text = true;
                }
                RunChild::Text(value) => text.push_str(&value.text),
                RunChild::Tab(_) | RunChild::FieldChar(_) => {}
                child => {
                    return Err(Error::Reverse(format!(
                        "static index row contains unsupported run content `{child:?}`; use text, tab, and a PAGEREF field only"
                    )));
                }
            }
        }
    }
    if bookmark.is_none() {
        if text.is_empty() {
            return Ok(None);
        }
        return Err(Error::Reverse(
            "index control contains text without a PAGEREF target; use representable static `TocItem` rows or remove the unsupported field result"
                .to_owned(),
        ));
    }
    if text.is_empty() {
        return Err(Error::Reverse(
            "static index row has a PAGEREF field but no entry text; add visible text before the tab"
                .to_owned(),
        ));
    }
    if !saw_page_ref_text {
        return Err(Error::Reverse(
            "static index row has no cached page-reference text; add a visible PAGEREF result or update the field in Word before reversing"
                .to_owned(),
        ));
    }
    let bookmark = bookmark.unwrap_or_default();
    if anchors.iter().any(|anchor| anchor != &bookmark) {
        return Err(Error::Reverse(
            "index entry hyperlink target does not match its PAGEREF bookmark; make both targets identical before reversing"
                .to_owned(),
        ));
    }
    let style = paragraph
        .property
        .style
        .as_ref()
        .map(|style| style.val.as_str())
        .and_then(|style| style.strip_prefix("ToC"))
        .and_then(|level| level.parse::<u64>().ok())
        .filter(|level| (1..=9).contains(level))
        .ok_or_else(|| {
            Error::Reverse(
                "static index row has no supported `ToC1`–`ToC9` paragraph style; apply a TOC level style before reversing"
                    .to_owned(),
            )
        })?;
    let tab_position = toc_item_tab_position(paragraph)?;
    Ok(Some(TocItemField {
        text,
        bookmark,
        level: style,
        page_ref,
        tab_position,
    }))
}

fn toc_item_tab_position(paragraph: &docx_rs::Paragraph) -> Result<u32> {
    let matching_tabs = paragraph
        .property
        .tabs
        .iter()
        .filter(|tab| {
            tab.val == Some(docx_rs::TabValueType::Right)
                && tab.leader == Some(docx_rs::TabLeaderType::Dot)
        })
        .collect::<Vec<_>>();
    if matching_tabs.len() != 1 {
        return Err(Error::Reverse(
            "static index row must contain exactly one right-aligned dotted page-number tab; add that tab or reduce the row to a representable TOC entry"
                .to_owned(),
        ));
    }
    let tab_position = matching_tabs[0].pos.ok_or_else(|| {
        Error::Reverse(
            "static index row page-number tab has no explicit position; set a positive `tabPosition` before reversing"
                .to_owned(),
        )
    })?;
    let tab_position = u32::try_from(tab_position)
        .map_err(|_| Error::Reverse("static index row tab position is out of range".to_owned()))?;
    if tab_position == 0 || tab_position > 31_680 {
        return Err(Error::Reverse(format!(
            "static index row tab position {tab_position} twips is outside OOXML's legal range 1..=31680; set a legal `tabPosition` no greater than 1584 points"
        )));
    }
    Ok(tab_position)
}

fn format_twips_as_points(twips: i64) -> String {
    let whole = twips / 20;
    let remainder = twips % 20;
    if remainder == 0 {
        whole.to_string()
    } else {
        format!("{whole}.{fraction:02}", fraction = remainder * 5)
    }
}

fn collect_toc_item_runs<'a>(
    children: &'a [ParagraphChild],
    runs: &mut Vec<&'a docx_rs::Run>,
    anchors: &mut Vec<String>,
) -> Result<()> {
    for child in children {
        match child {
            ParagraphChild::Run(run) => runs.push(run),
            ParagraphChild::Hyperlink(hyperlink) => {
                match &hyperlink.link {
                    docx_rs::HyperlinkData::Anchor { anchor } => anchors.push(anchor.clone()),
                    docx_rs::HyperlinkData::External { .. } => {
                        return Err(Error::Reverse(
                            "static index entry uses an external hyperlink; use a bookmark link matching its PAGEREF target"
                                .to_owned(),
                        ));
                    }
                }
                collect_toc_item_runs(&hyperlink.children, runs, anchors)?;
            }
            child => {
                return Err(Error::Reverse(format!(
                    "static index row contains unsupported paragraph content `{child:?}`; use runs or a bookmark hyperlink only"
                )));
            }
        }
    }
    Ok(())
}

fn parse_toc_item_pageref(instruction: &str) -> Result<String> {
    let mut words = instruction.split_whitespace();
    if words.next() != Some("PAGEREF") {
        return Err(Error::Reverse(
            "index row has an unsupported field instruction; use `PAGEREF bookmark`".to_owned(),
        ));
    }
    let bookmark = words
        .next()
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            Error::Reverse(
                "PAGEREF in index row has no bookmark target; add a target bookmark name"
                    .to_owned(),
            )
        })?;
    if words.any(|switch| switch != "\\h") {
        return Err(Error::Reverse(
            "PAGEREF in index row contains unsupported switches; use only the optional `\\h` switch"
                .to_owned(),
        ));
    }
    Ok(bookmark.to_owned())
}

fn add_index_sdt_properties(index: &mut IndexField, props: &SdtProps) {
    let (IndexField::Contents { attributes, .. }
    | IndexField::Figures { attributes, .. }
    | IndexField::Entries { attributes, .. }) = index;
    if let Some(alias) = &props.alias
        && !attributes.contains(" alias=")
    {
        attributes.push_str(&attr("alias", alias));
    }
}

fn sdt_field_dirty(tag: &StructuredDataTag) -> Option<bool> {
    tag.children.iter().find_map(|child| {
        let StructuredDataTagChild::Paragraph(paragraph) = child else {
            return None;
        };
        paragraph.children.iter().find_map(|child| {
            let ParagraphChild::Run(run) = child else {
                return None;
            };
            run.children.iter().find_map(|child| {
                let RunChild::FieldChar(field) = child else {
                    return None;
                };
                (field.field_char_type == docx_rs::FieldCharType::Begin).then_some(field.dirty)
            })
        })
    })
}

fn collect_sdt_instruction(
    tag: &StructuredDataTag,
    deletion: Option<&DeletedIndexField>,
) -> String {
    if let Some(deletion) = deletion {
        return deletion.instruction.clone();
    }
    let mut instruction = String::new();
    let mut in_first_field = false;
    let mut first_field_started = false;
    for child in &tag.children {
        let StructuredDataTagChild::Paragraph(paragraph) = child else {
            continue;
        };
        for child in &paragraph.children {
            let ParagraphChild::Run(run) = child else {
                continue;
            };
            for child in &run.children {
                match child {
                    RunChild::FieldChar(field)
                        if field.field_char_type == docx_rs::FieldCharType::Begin
                            && !first_field_started =>
                    {
                        first_field_started = true;
                        in_first_field = true;
                    }
                    RunChild::InstrTextString(value) if in_first_field => {
                        instruction.push_str(value);
                    }
                    RunChild::FieldChar(field)
                        if field.field_char_type == docx_rs::FieldCharType::Separate
                            && in_first_field =>
                    {
                        if instruction.trim_start().starts_with("TOC") {
                            return instruction;
                        }
                        in_first_field = false;
                    }
                    _ => {}
                }
            }
        }
    }
    instruction
}

#[expect(
    clippy::too_many_lines,
    reason = "TOC instruction decoding must validate switch compatibility before constructing one lossless JSX index component"
)]
fn parse_toc_instruction(instruction: &str) -> Result<Option<IndexField>> {
    let trimmed = instruction.trim();
    if trimmed != "TOC" && !trimmed.starts_with("TOC ") {
        return Ok(None);
    }
    let switches = parse_toc_switches(trimmed)?;
    for flag in switches.keys() {
        if !matches!(
            flag,
            'a' | 'b'
                | 'c'
                | 'd'
                | 'f'
                | 'h'
                | 'l'
                | 'n'
                | 'o'
                | 'p'
                | 's'
                | 't'
                | 'u'
                | 'w'
                | 'x'
                | 'z'
        ) {
            return Err(Error::Reverse(format!(
                "TOC field contains unsupported `\\{flag}` switch; extend the index DSL or remove that switch"
            )));
        }
    }
    for flag in ['h', 'u', 'w', 'x', 'z'] {
        if switches.get(&flag).is_some_and(Option::is_some) {
            return Err(Error::Reverse(format!(
                "TOC field `\\{flag}` does not take an argument; remove the following value"
            )));
        }
    }
    for flag in ['a', 'c'] {
        if switches.contains_key(&flag) && switches.get(&flag).is_some_and(Option::is_none) {
            return Err(Error::Reverse(format!(
                "TOC caption switch `\\{flag}` has no label; add a non-empty quoted caption name"
            )));
        }
    }
    let caption_number_label = switches.get(&'c').and_then(Option::as_ref);
    let caption_label = switches.get(&'a').and_then(Option::as_ref);
    if caption_number_label.is_some() && caption_label.is_some() {
        return Err(Error::Reverse(
            "TOC field cannot contain both caption switches `\\a` and `\\c`; retain only one"
                .to_owned(),
        ));
    }
    let is_figure_index = caption_number_label.is_some() || caption_label.is_some();
    let attributes = index_instruction_attributes(&switches, is_figure_index)?;
    if let Some(label) = caption_number_label.or(caption_label) {
        if label.is_empty() {
            return Err(Error::Reverse(
                "TOC caption switch has an empty label; provide a non-empty caption label"
                    .to_owned(),
            ));
        }
        if switches.contains_key(&'f') {
            return Err(Error::Reverse(
                "TOC field mixes caption and custom-entry switches; keep either the figure field or the custom-entry field".to_owned(),
            ));
        }
        return Ok(Some(IndexField::Figures {
            label: label.clone(),
            include_label_and_number: caption_number_label.is_some(),
            separator: switches.get(&'d').and_then(Clone::clone),
            attributes,
            items: Vec::new(),
            before: Vec::new(),
            after: Vec::new(),
        }));
    }
    if let Some(identifier) = switches.get(&'f') {
        let Some(identifier) = identifier else {
            return Err(Error::Reverse(
                "TOC field uses `\\f` without an identifier; provide a non-empty `TableOfEntries identifier`".to_owned(),
            ));
        };
        if identifier.is_empty() {
            return Err(Error::Reverse(
                "TOC field `\\f` has an empty identifier; provide a non-empty `TableOfEntries identifier`".to_owned(),
            ));
        }
        return Ok(Some(IndexField::Entries {
            identifier: identifier.clone(),
            attributes,
            items: Vec::new(),
            before: Vec::new(),
            after: Vec::new(),
        }));
    }
    let (start, end) = match switches.get(&'o') {
        Some(Some(value)) => parse_toc_range(value, "\\o")?,
        Some(None) => {
            return Err(Error::Reverse(
                "TOC field uses `\\o` without a level range; provide a range such as 1-3"
                    .to_owned(),
            ));
        }
        None => (1, 3),
    };
    Ok(Some(IndexField::Contents {
        start,
        end,
        attributes,
        items: Vec::new(),
        before: Vec::new(),
        after: Vec::new(),
    }))
}

fn parse_toc_switches(instruction: &str) -> Result<HashMap<char, Option<String>>> {
    let chars = instruction.chars().collect::<Vec<_>>();
    let mut cursor = 3;
    let mut switches = HashMap::new();
    while cursor < chars.len() {
        while chars.get(cursor).is_some_and(|value| value.is_whitespace()) {
            cursor += 1;
        }
        if cursor == chars.len() {
            break;
        }
        if chars[cursor] != '\\' {
            return Err(Error::Reverse(format!(
                "TOC field contains unexpected text after `TOC`; repair `{instruction}` so every option uses a supported switch"
            )));
        }
        cursor += 1;
        let Some(flag) = chars.get(cursor).copied() else {
            return Err(Error::Reverse(
                "TOC field ends with an incomplete switch; repair the field instruction".to_owned(),
            ));
        };
        cursor += 1;
        if switches.contains_key(&flag) {
            return Err(Error::Reverse(format!(
                "TOC field repeats `\\{flag}`; remove the duplicate switch"
            )));
        }
        while chars.get(cursor).is_some_and(|value| value.is_whitespace()) {
            cursor += 1;
        }
        let argument = match chars.get(cursor).copied() {
            None | Some('\\') => None,
            Some('"') => {
                cursor += 1;
                let mut value = String::new();
                let mut closed = false;
                while let Some(character) = chars.get(cursor).copied() {
                    cursor += 1;
                    if character == '"' {
                        closed = true;
                        break;
                    }
                    value.push(character);
                }
                if !closed {
                    return Err(Error::Reverse(format!(
                        "TOC field `\\{flag}` has an unterminated quoted argument; close the quote"
                    )));
                }
                Some(value)
            }
            Some(_) => {
                let mut value = String::new();
                while let Some(character) = chars.get(cursor).copied() {
                    if character.is_whitespace() || character == '\\' {
                        break;
                    }
                    value.push(character);
                    cursor += 1;
                }
                Some(value)
            }
        };
        switches.insert(flag, argument);
    }
    Ok(switches)
}

fn parse_toc_range(value: &str, switch: &str) -> Result<(u64, u64)> {
    let Some((start, end)) = value.split_once('-') else {
        return Err(Error::Reverse(format!(
            "TOC field `{switch}` range `{value}` is malformed; use a range such as 1-3"
        )));
    };
    let (Ok(start), Ok(end)) = (start.parse::<u64>(), end.parse::<u64>()) else {
        return Err(Error::Reverse(format!(
            "TOC field `{switch}` range `{value}` is malformed; use integer levels from 1 to 9"
        )));
    };
    if !(1..=9).contains(&start) || !(start..=9).contains(&end) {
        return Err(Error::Reverse(format!(
            "TOC field `{switch}` range `{value}` is invalid; use levels with 1 <= start <= end <= 9"
        )));
    }
    Ok((start, end))
}

#[expect(
    clippy::too_many_lines,
    reason = "the supported OOXML field switches share one deterministic attribute encoder and validation path"
)]
fn index_instruction_attributes(
    switches: &HashMap<char, Option<String>>,
    is_figure_index: bool,
) -> Result<String> {
    let mut attributes = String::new();
    for (flag, property) in [
        ('l', "tcFieldLevelRange"),
        ('n', "omitPageNumbersLevelRange"),
    ] {
        if let Some(value) = switches.get(&flag) {
            let Some(value) = value else {
                return Err(Error::Reverse(format!(
                    "TOC field `\\{flag}` has no level range; provide a value such as 1-3"
                )));
            };
            let (start, end) = parse_toc_range(value, &format!("\\{flag}"))?;
            write!(
                attributes,
                " {property}={{{{\"start\":{start},\"end\":{end}}}}}"
            )
            .map_err(|error| Error::Reverse(error.to_string()))?;
        }
    }
    for (flag, property) in [
        ('b', "entryBookmarkName"),
        ('p', "entryAndPageNumberSeparator"),
        ('s', "seqFieldIdentifierForPrefix"),
        (
            'd',
            if is_figure_index {
                "separator"
            } else {
                "sequenceAndPageNumbersSeparator"
            },
        ),
    ] {
        if flag == 'd' && is_figure_index {
            continue;
        }
        if let Some(value) = switches.get(&flag) {
            let Some(value) = value else {
                return Err(Error::Reverse(format!(
                    "TOC field `\\{flag}` has no text argument; add a quoted non-empty value"
                )));
            };
            if matches!(flag, 'b' | 's') && value.is_empty() {
                return Err(Error::Reverse(format!(
                    "TOC field `\\{flag}` has an empty value; provide a non-empty bookmark or sequence identifier"
                )));
            }
            attributes.push(' ');
            attributes.push_str(property);
            attributes.push('=');
            attributes.push_str(&serde_json::to_string(value).map_err(|error| {
                Error::Reverse(format!("cannot serialize TOC switch value: {error}"))
            })?);
        }
    }
    if let Some(value) = switches.get(&'t') {
        let Some(value) = value else {
            return Err(Error::Reverse(
                "TOC field `\\t` has no style mapping; provide quoted style,level pairs".to_owned(),
            ));
        };
        let entries = value.split(',').collect::<Vec<_>>();
        if entries.is_empty() || !entries.len().is_multiple_of(2) {
            return Err(Error::Reverse(format!(
                "TOC field `\\t` mapping `{value}` is malformed; use comma-separated style,level pairs"
            )));
        }
        let mut mappings = Vec::new();
        for pair in entries.chunks_exact(2) {
            let level = pair[1]
                .parse::<u64>()
                .ok()
                .filter(|level| (1..=9).contains(level));
            let Some(level) = level else {
                return Err(Error::Reverse(format!(
                    "TOC field `\\t` mapping `{value}` has invalid level `{}`; use 1 to 9",
                    pair[1]
                )));
            };
            if pair[0].is_empty() {
                return Err(Error::Reverse(
                    "TOC field `\\t` contains an empty style name; provide a non-empty style name"
                        .to_owned(),
                ));
            }
            mappings.push(serde_json::json!({"style": pair[0], "level": level}));
        }
        write!(
            attributes,
            " stylesWithLevels={{{}}}",
            serde_json::to_string(&mappings)
                .map_err(|error| Error::Reverse(format!("cannot serialize TOC styles: {error}")))?
        )
        .map_err(|error| Error::Reverse(error.to_string()))?;
    }
    for (flag, property) in [
        ('u', "useAppliedParagraphLineLevel"),
        ('w', "preserveTab"),
        ('x', "preserveNewLine"),
        ('z', "hideTabAndPageNumbersInWebview"),
    ] {
        if switches.contains_key(&flag) {
            write!(attributes, " {property}").map_err(|error| Error::Reverse(error.to_string()))?;
        }
    }
    if !switches.contains_key(&'h') {
        attributes.push_str(" hyperlinks={false}");
    }
    Ok(attributes)
}

fn is_body_level_sdt(stack: &[Vec<u8>]) -> bool {
    stack.iter().any(|name| name.as_slice() == b"body")
        && !stack.iter().any(|name| {
            matches!(
                name.as_slice(),
                b"p" | b"tbl" | b"tr" | b"tc" | b"hyperlink" | b"sdt"
            )
        })
}

fn is_cell_level_sdt(stack: &[Vec<u8>]) -> bool {
    stack.iter().any(|name| name.as_slice() == b"tc")
        && !stack.iter().any(|name| name.as_slice() == b"p")
}

fn parse_sdt_plan(bytes: &[u8]) -> Result<SdtPlan> {
    let Ok(xml) = read_zip_text(bytes, "word/document.xml") else {
        return Ok(SdtPlan::default());
    };
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);
    let mut plan = SdtPlan::default();
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"p" => {
                plan.paragraphs.push(Vec::new());
            }
            Ok(Event::Start(element)) => {
                let name = element.local_name().as_ref().to_vec();
                if name == b"p" {
                    let events = parse_container_events(&mut reader, b"p", &mut plan)?;
                    plan.paragraphs.push(events);
                } else if name == b"hyperlink" {
                    let events = parse_container_events(&mut reader, b"hyperlink", &mut plan)?;
                    plan.hyperlinks.push(events);
                } else if name == b"sdt" {
                    if is_body_level_sdt(&stack) {
                        plan.body_tags
                            .push(parse_sdt_properties_until_content(&mut reader)?);
                    } else if is_cell_level_sdt(&stack) {
                        plan.cell_tags
                            .push(parse_sdt_properties_until_content(&mut reader)?);
                        skip_to_end(&mut reader, b"sdtContent")?;
                    } else {
                        skip_to_end(&mut reader, b"sdt")?;
                    }
                } else {
                    stack.push(name);
                }
            }
            Ok(Event::End(element)) => {
                let name = element.local_name();
                if stack
                    .last()
                    .is_some_and(|value| value.as_slice() == name.as_ref())
                {
                    stack.pop();
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid word/document.xml while reading structured tags: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(plan)
}

fn parse_header_footer_sdt_props(
    bytes: &[u8],
    parts: &HashMap<String, String>,
) -> Result<HashMap<String, Vec<SdtProps>>> {
    let mut output = HashMap::new();
    for part in parts.values() {
        let Ok(xml) = read_zip_text(bytes, part) else {
            continue;
        };
        let mut reader = Reader::from_str(&xml);
        reader.config_mut().trim_text(true);
        let mut props = Vec::new();
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(element)) if element.local_name().as_ref() == b"sdt" => {
                    props.push(parse_sdt_properties_until_content(&mut reader)?);
                }
                Ok(Event::Eof) => break,
                Err(error) => {
                    return Err(Error::Reverse(format!(
                        "invalid {part} while reading structured tags: {error}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }
        output.insert(part.clone(), props);
    }
    Ok(output)
}

fn parse_container_events(
    reader: &mut Reader<&[u8]>,
    end_name: &[u8],
    plan: &mut SdtPlan,
) -> Result<Vec<SdtEvent>> {
    let mut events = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(element)) => {
                let name = element.local_name().as_ref().to_vec();
                if is_leaf_child(&name) {
                    events.push(SdtEvent::Leaf);
                }
            }
            Ok(Event::Start(element)) => {
                let name = element.local_name().as_ref().to_vec();
                if name == b"sdt" {
                    let (props, inner) = parse_sdt_wrapper(reader, plan)?;
                    events.push(SdtEvent::Start(props));
                    events.extend(inner);
                    events.push(SdtEvent::End);
                } else if name == b"hyperlink" && end_name != b"hyperlink" {
                    events.push(SdtEvent::Leaf);
                    let inner = parse_container_events(reader, b"hyperlink", plan)?;
                    plan.hyperlinks.push(inner);
                } else if is_leaf_child(&name) {
                    events.push(SdtEvent::Leaf);
                    skip_to_end(reader, &name)?;
                } else if name == b"p" && end_name != b"p" {
                    return Err(Error::Reverse(
                        "block-level structured document tags are not representable as ContentControl; unwrap them or convert to inline `w:sdt` children"
                            .to_owned(),
                    ));
                } else {
                    skip_to_end(reader, &name)?;
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == end_name => {
                return Ok(events);
            }
            Ok(Event::Eof) => {
                return Err(Error::Reverse(format!(
                    "unterminated <w:{}> while reading structured tags",
                    String::from_utf8_lossy(end_name)
                )));
            }
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid word/document.xml while reading structured tags: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
}

fn parse_sdt_properties_until_content(reader: &mut Reader<&[u8]>) -> Result<SdtProps> {
    let mut props = SdtProps::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => {
                let name = element.local_name().as_ref().to_vec();
                if name == b"sdtPr" {
                    props = parse_sdt_properties(reader)?;
                } else if name == b"sdtContent" {
                    return Ok(props);
                } else {
                    skip_to_end(reader, &name)?;
                }
            }
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"sdtContent" => {
                return Ok(props);
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sdt" => {
                return Ok(props);
            }
            Ok(Event::Eof) => {
                return Err(Error::Reverse(
                    "unterminated `w:sdt` while reading structured tags".to_owned(),
                ));
            }
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid word/document.xml while reading structured tags: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
}

fn parse_sdt_wrapper(
    reader: &mut Reader<&[u8]>,
    plan: &mut SdtPlan,
) -> Result<(SdtProps, Vec<SdtEvent>)> {
    let mut props = SdtProps::default();
    let mut inner = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => {
                let name = element.local_name().as_ref().to_vec();
                if name == b"sdtPr" {
                    props = parse_sdt_properties(reader)?;
                } else if name == b"sdtContent" {
                    inner = parse_container_events(reader, b"sdtContent", plan)?;
                } else {
                    skip_to_end(reader, &name)?;
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sdt" => {
                return Ok((props, inner));
            }
            Ok(Event::Eof) => {
                return Err(Error::Reverse(
                    "unterminated `w:sdt` while reading structured tags".to_owned(),
                ));
            }
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid word/document.xml while reading structured tags: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
}

fn parse_sdt_properties(reader: &mut Reader<&[u8]>) -> Result<SdtProps> {
    let mut props = SdtProps::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(element)) => {
                apply_sdt_property(&mut props, reader, &element)?;
            }
            Ok(Event::Start(element)) => {
                apply_sdt_property(&mut props, reader, &element)?;
                let name = element.local_name().as_ref().to_vec();
                skip_to_end(reader, &name)?;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sdtPr" => {
                return Ok(props);
            }
            Ok(Event::Eof) => {
                return Err(Error::Reverse(
                    "unterminated `w:sdtPr` while reading structured tags".to_owned(),
                ));
            }
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid word/document.xml while reading structured tags: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
}

fn apply_sdt_property(
    props: &mut SdtProps,
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<()> {
    let name = element.local_name();
    if name.as_ref() == b"alias" {
        props.alias = xml_attribute(reader, element, b"val")?;
    } else if name.as_ref() == b"dataBinding" {
        props.xpath = xml_attribute(reader, element, b"xpath")?;
        props.prefix_mappings = xml_attribute(reader, element, b"prefixMappings")?;
        props.store_item_id = xml_attribute(reader, element, b"storeItemID")?;
    }
    Ok(())
}

fn is_leaf_child(name: &[u8]) -> bool {
    matches!(
        name,
        b"r" | b"ins"
            | b"del"
            | b"moveFrom"
            | b"moveTo"
            | b"bookmarkStart"
            | b"bookmarkEnd"
            | b"commentRangeStart"
            | b"commentRangeEnd"
            | b"fldSimple"
            | b"hyperlink"
    )
}

fn skip_to_end(reader: &mut Reader<&[u8]>, name: &[u8]) -> Result<()> {
    let mut depth = 1_usize;
    let mut buf = Vec::new();
    while depth > 0 {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == name => {
                depth += 1;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == name => {
                depth -= 1;
            }
            Ok(Event::Eof) => {
                return Err(Error::Reverse(format!(
                    "unterminated <w:{}> while reading structured tags",
                    String::from_utf8_lossy(name)
                )));
            }
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid word/document.xml while reading structured tags: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn read_zip_text(bytes: &[u8], name: &str) -> Result<String> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    let mut file = archive
        .by_name(name)
        .map_err(|error| Error::Reverse(format!("missing {name}: {error}")))?;
    let mut text = String::new();
    file.read_to_string(&mut text)
        .map_err(|error| Error::Reverse(format!("cannot read {name}: {error}")))?;
    Ok(text)
}

fn collect_hyperlink_histories(bytes: &[u8]) -> Result<HashMap<String, VecDeque<Option<usize>>>> {
    let xml = read_zip_text(bytes, "word/document.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut histories: HashMap<String, VecDeque<Option<usize>>> = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"hyperlink" =>
            {
                let mut key = None;
                let mut history = None;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Reverse(format!(
                            "invalid hyperlink attribute while reading `w:history`: {error}"
                        ))
                    })?;
                    let name = attribute.key.local_name();
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::Reverse(format!(
                                "cannot decode hyperlink attribute while reading `w:history`: {error}"
                            ))
                        })?;
                    match name.as_ref() {
                        b"id" => key = Some(format!("r:{value}")),
                        b"anchor" => key = Some(format!("a:{value}")),
                        b"history" => {
                            history = Some(match value.as_ref() {
                                "1" | "true" | "on" => 1,
                                "0" | "false" | "off" => 0,
                                _ => {
                                    return Err(Error::Reverse(format!(
                                        "hyperlink `w:history` has invalid value `{value}`; use true, false, 1, 0, on, or off"
                                    )));
                                }
                            });
                        }
                        _ => {}
                    }
                }
                if let Some(key) = key {
                    histories.entry(key).or_default().push_back(history);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "invalid `word/document.xml` while reading hyperlink history: {error}"
                )));
            }
        }
    }
    Ok(histories)
}

fn split_section_groups<'a>(
    children: &'a [NestedBlock<'a>],
) -> Vec<(&'a [NestedBlock<'a>], Option<&'a docx_rs::SectionProperty>)> {
    let mut groups = Vec::new();
    let mut start = 0;
    for (index, child) in children.iter().enumerate() {
        if let Some(property) = section_break_of(child) {
            groups.push((&children[start..=index], Some(property)));
            start = index + 1;
        }
    }
    groups.push((&children[start..], None));
    groups
        .into_iter()
        .filter(|(group, property)| property.is_some() || !group.is_empty())
        .collect()
}

fn section_break_of<'a>(child: &'a NestedBlock<'a>) -> Option<&'a docx_rs::SectionProperty> {
    match child {
        NestedBlock::Child(DocumentChild::Paragraph(paragraph)) => {
            paragraph.property.section_property.as_ref()
        }
        _ => None,
    }
}

fn is_section_break_marker(child: &NestedBlock<'_>) -> bool {
    let NestedBlock::Child(DocumentChild::Paragraph(paragraph)) = child else {
        return false;
    };
    paragraph.property.section_property.is_some()
        && paragraph.children.iter().all(|child| match child {
            ParagraphChild::Run(run) => run_is_ignorable(run),
            _ => false,
        })
}

fn table_jsx_attrs(
    table: &docx_rs::Table,
    source: &Value,
    metadata: &TableXmlMetadata,
) -> Result<String> {
    let mut attributes = String::new();
    attributes.push_str(&table_width_jsx_attrs(source)?);
    if let Some(value) = source
        .get("justification")
        .and_then(scalar_string)
        .filter(|value| *value != "left")
    {
        attributes.push_str(&attr("align", value));
    }
    if let Some(value) = metadata
        .layout
        .as_deref()
        .or_else(|| source.get("layout").and_then(Value::as_str))
    {
        let layout = match value {
            "fixed" => "fixed",
            "autofit" => "auto",
            _ => {
                return Err(Error::Reverse(format!(
                    "Table layout `{value}` is unsupported; use fixed or autofit in the DOCX"
                )));
            }
        };
        attributes.push_str(&attr("layout", layout));
    }
    if !table.grid.is_empty() {
        let widths = table
            .grid
            .iter()
            .map(|width| {
                let width = i64::try_from(*width).map_err(|_| {
                    Error::Reverse("Table grid column width exceeds supported range".to_owned())
                })?;
                if width <= 0 {
                    return Err(Error::Reverse(
                        "Table grid contains a non-positive column width; set every `columnWidths` entry to a positive point value or repair `w:tblGrid`".to_owned(),
                    ));
                }
                points_value(width)
            })
            .collect::<Result<Vec<_>>>()?;
        attributes.push_str(&jsx_prop("columnWidths", &Value::Array(widths))?);
    }
    if let Some(value) = source.get("style").and_then(scalar_string) {
        attributes.push_str(&attr("style", value));
    }
    if let Some(value) = source.get("indent").filter(|value| !value.is_null()) {
        let width_type = value
            .get("widthType")
            .and_then(Value::as_str)
            .unwrap_or_default();
        if width_type != "dxa" {
            return Err(Error::Reverse(format!(
                "Table indent uses `{width_type}`, which JSX cannot represent; convert it to dxa or remove the indent"
            )));
        }
        let twips = value.get("width").and_then(Value::as_i64).ok_or_else(|| {
            Error::Reverse("Table indent is malformed; repair `w:tblInd`".to_owned())
        })?;
        attributes.push_str(&jsx_prop("indent", &points_value(twips)?)?);
    }
    if let Some(value) = source.get("margins").filter(|value| !value.is_null()) {
        attributes.push_str(&jsx_prop("margins", &table_cell_margins_jsx(value)?)?);
    }
    if let Some(value) = source.get("borders").filter(|value| !value.is_null()) {
        attributes.push_str(&jsx_prop("border", &table_cell_border_jsx(value)?)?);
    }
    if let Some(value) = source.get("position").filter(|value| !value.is_null()) {
        attributes.push_str(&jsx_prop("position", &table_position_jsx(value)?)?);
    }
    Ok(attributes)
}

fn table_width_jsx_attrs(source: &Value) -> Result<String> {
    let mut attributes = String::new();
    if let Some(value) = source.get("width").filter(|value| !value.is_null()) {
        let width = value.get("width").and_then(Value::as_u64).ok_or_else(|| {
            Error::Reverse("Table width is malformed; repair `w:tblW`".to_owned())
        })?;
        let width_type = value
            .get("widthType")
            .and_then(Value::as_str)
            .ok_or_else(|| Error::Reverse("Table width unit is missing".to_owned()))?;
        match width_type {
            "dxa" => {
                let twips = i64::try_from(width).map_err(|_| {
                    Error::Reverse("Table width exceeds supported range".to_owned())
                })?;
                attributes.push_str(&jsx_prop("width", &points_value(twips)?)?);
            }
            "pct" if width <= 5_000 => {
                attributes.push_str(&jsx_prop(
                    "widthPercent",
                    &Value::Number(percent_value(width)?),
                )?);
            }
            "pct" => {
                return Err(Error::Reverse(format!(
                    "Table percentage width `{width}` exceeds JSX's 100 percent limit; reduce the width or retain the DOCX"
                )));
            }
            "auto" => {}
            _ => {
                return Err(Error::Reverse(format!(
                    "Table width uses `{width_type}`, which JSX cannot represent; convert it to dxa, pct, or auto"
                )));
            }
        }
    }
    Ok(attributes)
}

fn percent_value(fiftieths: u64) -> Result<Number> {
    if fiftieths.is_multiple_of(50) {
        return Ok(Number::from(fiftieths / 50));
    }
    Number::from_f64(
        f64::from(
            u32::try_from(fiftieths)
                .map_err(|_| Error::Reverse("percentage value is out of range".to_owned()))?,
        ) / 50.0,
    )
    .ok_or_else(|| Error::Reverse("percentage value is out of range".to_owned()))
}

fn table_position_jsx(source: &Value) -> Result<Value> {
    let mut output = Map::new();
    for (source_key, target_key) in [
        ("leftFromText", "leftFromText"),
        ("rightFromText", "rightFromText"),
        ("positionX", "x"),
        ("positionY", "y"),
    ] {
        if let Some(value) = source.get(source_key).and_then(Value::as_i64) {
            output.insert(target_key.to_owned(), points_value(value)?);
        }
    }
    for (source_key, target_key) in [
        ("verticalAnchor", "verticalAnchor"),
        ("horizontalAnchor", "horizontalAnchor"),
        ("positionXAlignment", "xAlign"),
        ("positionYAlignment", "yAlign"),
    ] {
        if let Some(value) = source.get(source_key).and_then(Value::as_str) {
            output.insert(target_key.to_owned(), Value::String(value.to_owned()));
        }
    }
    Ok(Value::Object(output))
}

#[allow(clippy::too_many_lines)]
fn table_cell_jsx_attrs(cell: &docx_rs::TableCell) -> Result<String> {
    let source = json(&cell.property)?;
    let properties = table_cell_jsx_props(&source)?;
    let mut attributes = String::new();
    for (name, value) in properties {
        attributes.push_str(&jsx_prop(&name, &value)?);
    }
    Ok(attributes)
}

#[allow(clippy::too_many_lines)]
fn table_cell_jsx_props(source: &Value) -> Result<Map<String, Value>> {
    let mut properties = Map::new();
    if let Some(width) = source.get("width").filter(|value| !value.is_null()) {
        let value = width
            .get("width")
            .and_then(Value::as_u64)
            .ok_or_else(|| Error::Reverse("TableCell width is malformed".to_owned()))?;
        let width_type = width
            .get("widthType")
            .and_then(Value::as_str)
            .ok_or_else(|| Error::Reverse("TableCell width unit is missing".to_owned()))?;
        match width_type {
            "dxa" => {
                let twips = i64::try_from(value).map_err(|_| {
                    Error::Reverse("TableCell width exceeds the supported point range".to_owned())
                })?;
                properties.insert("width".to_owned(), points_value(twips)?);
            }
            "pct" if value <= 5_000 => {
                let percent = if value % 50 == 0 {
                    Number::from(value / 50)
                } else {
                    Number::from_f64(
                        f64::from(u32::try_from(value).map_err(|_| {
                            Error::Reverse("TableCell percentage width is out of range".to_owned())
                        })?) / 50.0,
                    )
                    .ok_or_else(|| {
                        Error::Reverse("TableCell percentage width is out of range".to_owned())
                    })?
                };
                properties.insert("widthPercent".to_owned(), Value::Number(percent));
            }
            "pct" => {
                return Err(Error::Reverse(format!(
                    "TableCell percentage width `{value}` exceeds JSX's 100 percent limit; reduce the width or retain the DOCX"
                )));
            }
            _ => {
                return Err(Error::Reverse(format!(
                    "TableCell width uses `{width_type}`, which JSX cannot represent; convert the width to `dxa` or `pct` or retain the DOCX"
                )));
            }
        }
    }
    if let Some(value) = source.get("gridSpan").filter(|value| !value.is_null()) {
        let span = scalar_u64(value).ok_or_else(|| {
            Error::Reverse("TableCell grid span is malformed; repair `w:gridSpan`".to_owned())
        })?;
        if span == 0 {
            return Err(Error::Reverse(
                "TableCell grid span is zero and is not valid; set `colSpan` to a positive integer or repair `w:gridSpan`".to_owned(),
            ));
        }
        if span > 1 {
            properties.insert("colSpan".to_owned(), Value::Number(span.into()));
        }
    }
    if let Some(value) = source.get("verticalAlign").filter(|value| !value.is_null()) {
        let alignment = scalar_string(value).ok_or_else(|| {
            Error::Reverse("TableCell vertical alignment is malformed".to_owned())
        })?;
        if !["top", "center", "bottom"].contains(&alignment) {
            return Err(Error::Reverse(format!(
                "TableCell vertical alignment `{alignment}` is unsupported; use top, center, or bottom in the DOCX"
            )));
        }
        properties.insert(
            "verticalAlign".to_owned(),
            Value::String(alignment.to_owned()),
        );
    }
    if let Some(value) = source.get("verticalMerge").filter(|value| !value.is_null()) {
        let merge = scalar_string(value)
            .ok_or_else(|| Error::Reverse("TableCell vertical merge is malformed".to_owned()))?;
        if !["restart", "continue"].contains(&merge) {
            return Err(Error::Reverse(format!(
                "TableCell vertical merge `{merge}` is unsupported; repair `w:vMerge`"
            )));
        }
        properties.insert("verticalMerge".to_owned(), Value::String(merge.to_owned()));
    }
    if let Some(value) = source.get("textDirection").filter(|value| !value.is_null()) {
        let direction = scalar_string(value)
            .ok_or_else(|| Error::Reverse("TableCell text direction is malformed".to_owned()))?;
        if ![
            "lr", "lrV", "rl", "rlV", "tb", "tbV", "tbRlV", "tbRl", "btLr", "lrTbV",
        ]
        .contains(&direction)
        {
            return Err(Error::Reverse(format!(
                "TableCell text direction `{direction}` is unsupported; repair `w:textDirection`"
            )));
        }
        properties.insert(
            "textDirection".to_owned(),
            Value::String(direction.to_owned()),
        );
    }
    if let Some(value) = source.get("margins").filter(|value| !value.is_null()) {
        let margins = table_cell_margins_jsx(value)?;
        properties.insert("margins".to_owned(), margins);
    }
    if let Some(value) = source.get("shading").filter(|value| !value.is_null()) {
        let fill = value.get("fill").and_then(Value::as_str).ok_or_else(|| {
            Error::Reverse("TableCell shading has no representable fill color".to_owned())
        })?;
        if fill.len() != 6 || !fill.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(Error::Reverse(
                "TableCell shading fill is not six-digit RGB; repair `w:shd/@w:fill` or retain the DOCX".to_owned(),
            ));
        }
        if value.get("shdType").and_then(Value::as_str) != Some("clear")
            || value.get("color").and_then(Value::as_str) != Some("auto")
        {
            return Err(Error::Reverse(
                "TableCell shading uses a pattern or foreground color not represented by JSX; use a solid fill or retain the DOCX".to_owned(),
            ));
        }
        if fill != "FFFFFF" {
            properties.insert("shading".to_owned(), Value::String(fill.to_owned()));
        }
    }
    if let Some(value) = source.get("borders").filter(|value| !value.is_null()) {
        properties.insert("border".to_owned(), table_cell_border_jsx(value)?);
    }
    Ok(properties)
}

fn table_cell_margins_jsx(source: &Value) -> Result<Value> {
    let mut margins = Map::new();
    for edge in ["top", "right", "bottom", "left"] {
        let value = source.get(edge).filter(|value| !value.is_null()).ok_or_else(|| {
            Error::Reverse(format!(
                "TableCell margins specify only some edges; JSX requires all four, so add explicit {edge} spacing or retain the DOCX"
            ))
        })?;
        let width_type = value
            .get("widthType")
            .and_then(Value::as_str)
            .ok_or_else(|| Error::Reverse(format!("TableCell `{edge}` margin unit is missing")))?;
        if width_type != "dxa" {
            return Err(Error::Reverse(format!(
                "TableCell `{edge}` margin uses `{width_type}`, which JSX cannot represent; convert it to `dxa` or retain the DOCX"
            )));
        }
        let twips = value
            .get("val")
            .and_then(Value::as_u64)
            .and_then(|value| i64::try_from(value).ok())
            .ok_or_else(|| Error::Reverse(format!("TableCell `{edge}` margin is malformed")))?;
        margins.insert(edge.to_owned(), points_value(twips)?);
    }
    Ok(Value::Object(margins))
}

fn table_cell_border_jsx(source: &Value) -> Result<Value> {
    let borders = source
        .as_object()
        .ok_or_else(|| Error::Reverse("TableCell borders are malformed".to_owned()))?;
    let mut result = Map::new();
    for (source_edge, target_edge) in [
        ("top", "top"),
        ("right", "right"),
        ("bottom", "bottom"),
        ("left", "left"),
        ("insideH", "insideHorizontal"),
        ("insideV", "insideVertical"),
        ("tl2br", "topLeftToBottomRight"),
        ("tr2bl", "topRightToBottomLeft"),
    ] {
        let Some(border) = borders.get(source_edge).filter(|value| !value.is_null()) else {
            continue;
        };
        let style = border
            .get("borderType")
            .and_then(Value::as_str)
            .ok_or_else(|| {
                Error::Reverse(format!("TableCell `{source_edge}` border style is missing"))
            })?;
        if style == "nil" || style == "none" {
            result.insert(target_edge.to_owned(), Value::Bool(false));
            continue;
        }
        if border
            .get("space")
            .and_then(Value::as_u64)
            .unwrap_or_default()
            != 0
        {
            return Err(Error::Reverse(format!(
                "TableCell `{source_edge}` border has nonzero spacing, which JSX cannot represent; remove the border spacing or retain the DOCX"
            )));
        }
        let size = border.get("size").and_then(Value::as_u64).ok_or_else(|| {
            Error::Reverse(format!(
                "TableCell `{source_edge}` border size is malformed"
            ))
        })?;
        let color = border
            .get("color")
            .and_then(Value::as_str)
            .filter(|color| color.len() == 6 && color.bytes().all(|byte| byte.is_ascii_hexdigit()))
            .ok_or_else(|| {
                Error::Reverse(format!(
                    "TableCell `{source_edge}` border color is not six-digit RGB; use an explicit RGB color or retain the DOCX"
                ))
            })?;
        let border_size = if size % 8 == 0 {
            Number::from(size / 8)
        } else {
            Number::from_f64(
                f64::from(u32::try_from(size).map_err(|_| {
                    Error::Reverse(format!(
                        "TableCell `{source_edge}` border size is out of range"
                    ))
                })?) / 8.0,
            )
            .ok_or_else(|| {
                Error::Reverse(format!(
                    "TableCell `{source_edge}` border size is out of range"
                ))
            })?
        };
        let mut spec = Map::new();
        spec.insert("style".to_owned(), Value::String(style.to_owned()));
        spec.insert("size".to_owned(), Value::Number(border_size));
        spec.insert("color".to_owned(), Value::String(color.to_owned()));
        result.insert(target_edge.to_owned(), Value::Object(spec));
    }
    Ok(Value::Object(result))
}

fn table_row_jsx_attrs(row: &docx_rs::TableRow, row_xml: &TableRowXmlMetadata) -> Result<String> {
    let source = json(&row.property)?;
    let mut attributes = String::new();
    for key in ["gridBefore", "gridAfter"] {
        if let Some(value) = source.get(key).filter(|value| !value.is_null()) {
            let count = value.as_u64().ok_or_else(|| {
                Error::Reverse(format!(
                    "TableRow `{key}` is not a non-negative integer; repair this row grid value"
                ))
            })?;
            attributes.push_str(&jsx_prop(key, &Value::from(count))?);
        }
    }
    for (key, xml_type) in [
        ("widthBefore", row_xml.width_before_type.as_deref()),
        ("widthAfter", row_xml.width_after_type.as_deref()),
    ] {
        if let Some(value) = source.get(key).filter(|value| !value.is_null()) {
            if xml_type.is_some_and(|kind| kind != "dxa") {
                return Err(Error::Reverse(format!(
                    "TableRow `{key}` uses width type `{}` which cannot be represented in JSX; convert it to dxa (twips) in the DOCX",
                    xml_type.unwrap_or_default()
                )));
            }
            let twips = value.as_f64().filter(|value| value.is_finite() && *value >= 0.0 && value.fract() == 0.0).and_then(|value| value.to_i64()).ok_or_else(|| Error::Reverse(format!("TableRow `{key}` is not a non-negative whole number of twips; repair this width")))?;
            attributes.push_str(&jsx_prop(key, &points_value(twips)?)?);
        }
    }
    if let Some(value) = source.get("rowHeight").filter(|value| !value.is_null()) {
        let height = value.as_f64().ok_or_else(|| {
            Error::Reverse(
                "TableRow height is not numeric; repair the row height in the DOCX".to_owned(),
            )
        })?;
        let twips = height.to_i64().filter(|twips| {
            height.is_finite() && height.fract() == 0.0 && *twips >= 0
        }).ok_or_else(|| {
            Error::Reverse(format!(
                "TableRow height `{height}` cannot be represented as a non-negative whole number of twips; repair the row height"
            ))
        })?;
        attributes.push_str(&jsx_prop("height", &points_value(twips)?)?);
    }
    if let Some(value) = source.get("heightRule").filter(|value| !value.is_null()) {
        let rule = scalar_string(value).ok_or_else(|| {
            Error::Reverse(
                "TableRow height rule is invalid; use `auto`, `atLeast`, or `exact` in the DOCX"
                    .to_owned(),
            )
        })?;
        if !["auto", "atLeast", "exact"].contains(&rule) {
            return Err(Error::Reverse(format!(
                "TableRow height rule `{rule}` is unsupported; use `auto`, `atLeast`, or `exact`"
            )));
        }
        attributes.push_str(&attr("heightRule", rule));
    }
    if row_xml.cant_split.unwrap_or_else(|| {
        source
            .get("cantSplit")
            .is_some_and(|value| value.as_bool().unwrap_or(false))
    }) {
        attributes.push_str(" cantSplit");
    }
    for (source_key, target_key) in [("ins", "inserted"), ("del", "deleted")] {
        let Some(value) = source.get(source_key).filter(|value| !value.is_null()) else {
            continue;
        };
        let revision = value.as_object().ok_or_else(|| {
            Error::Reverse(format!(
                "TableRow `{source_key}` revision metadata is malformed; repair its author/date or remove the revision"
            ))
        })?;
        if revision
            .get("children")
            .and_then(Value::as_array)
            .is_some_and(|children| !children.is_empty())
        {
            return Err(Error::Reverse(format!(
                "TableRow `{source_key}` contains child revision data not modeled by JSX; preserve the original DOCX or remove the nested row revision"
            )));
        }
        let mut metadata = Map::new();
        for key in ["author", "date"] {
            if let Some(value) = revision.get(key).and_then(Value::as_str) {
                metadata.insert(key.to_owned(), Value::String(value.to_owned()));
            }
        }
        attributes.push_str(&jsx_prop(target_key, &Value::Object(metadata))?);
    }
    Ok(attributes)
}

fn collect_table_row_xml_metadata(
    bytes: &[u8],
) -> Result<HashMap<String, Vec<TableRowXmlMetadata>>> {
    let mut result = HashMap::new();
    for part in zip_entry_names(bytes)?.into_iter().filter(|name| {
        name == "word/document.xml"
            || (name.starts_with("word/header") || name.starts_with("word/footer"))
                && Path::new(name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
                && !name.contains("/_rels/")
    }) {
        let xml = read_zip_text(bytes, &part)?;
        let mut reader = Reader::from_str(&xml);
        let mut buffer = Vec::new();
        let mut rows = Vec::<TableRowXmlMetadata>::new();
        let mut active_rows = Vec::<usize>::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element)) if element.local_name().as_ref() == b"tr" => {
                    rows.push(TableRowXmlMetadata::default());
                    active_rows.push(rows.len() - 1);
                }
                Ok(Event::Empty(element)) if element.local_name().as_ref() == b"tr" => {
                    rows.push(TableRowXmlMetadata::default());
                }
                Ok(Event::Start(element) | Event::Empty(element))
                    if element.local_name().as_ref() == b"cantSplit" =>
                {
                    if let Some(index) = active_rows.last().copied() {
                        rows[index].cant_split = Some(xml_on_off(&reader, &element)?);
                    }
                }
                Ok(Event::Start(element) | Event::Empty(element))
                    if element.local_name().as_ref() == b"wBefore"
                        || element.local_name().as_ref() == b"wAfter" =>
                {
                    if let Some(index) = active_rows.last().copied() {
                        let kind = xml_attribute(&reader, &element, b"type")?;
                        if element.local_name().as_ref() == b"wBefore" {
                            rows[index].width_before_type = kind;
                        } else {
                            rows[index].width_after_type = kind;
                        }
                    }
                }
                Ok(Event::End(element)) if element.local_name().as_ref() == b"tr" => {
                    active_rows.pop();
                }
                Ok(Event::Eof) => break,
                Err(error) => {
                    return Err(Error::Reverse(format!(
                        "cannot parse {part} table row properties at byte {}: {error}",
                        reader.error_position()
                    )));
                }
                _ => {}
            }
            buffer.clear();
        }
        result.insert(part, rows);
    }
    Ok(result)
}

fn collect_table_xml_metadata(bytes: &[u8]) -> Result<HashMap<String, Vec<TableXmlMetadata>>> {
    let mut result = HashMap::new();
    for part in zip_entry_names(bytes)?.into_iter().filter(|name| {
        name == "word/document.xml"
            || (name.starts_with("word/header") || name.starts_with("word/footer"))
                && Path::new(name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
                && !name.contains("/_rels/")
    }) {
        let xml = read_zip_text(bytes, &part)?;
        let mut reader = Reader::from_str(&xml);
        let mut buffer = Vec::new();
        let mut tables = Vec::<TableXmlMetadata>::new();
        let mut active_tables = Vec::<usize>::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element)) if element.local_name().as_ref() == b"tbl" => {
                    tables.push(TableXmlMetadata::default());
                    active_tables.push(tables.len() - 1);
                }
                Ok(Event::Empty(element)) if element.local_name().as_ref() == b"tbl" => {
                    tables.push(TableXmlMetadata::default());
                }
                Ok(Event::Start(element) | Event::Empty(element))
                    if element.local_name().as_ref() == b"tblLayout" =>
                {
                    if let Some(index) = active_tables.last().copied() {
                        tables[index].layout = xml_attribute(&reader, &element, b"type")?;
                    }
                }
                Ok(Event::End(element)) if element.local_name().as_ref() == b"tbl" => {
                    active_tables.pop();
                }
                Ok(Event::Eof) => break,
                Err(error) => {
                    return Err(Error::Reverse(format!(
                        "cannot parse {part} table properties at byte {}: {error}",
                        reader.error_position()
                    )));
                }
                _ => {}
            }
            buffer.clear();
        }
        result.insert(part, tables);
    }
    Ok(result)
}

fn validate_character_indent_xml(bytes: &[u8]) -> Result<()> {
    const UNSUPPORTED: &[&[u8]] = &[
        b"leftChars",
        b"startChars",
        b"rightChars",
        b"endChars",
        b"hangingChars",
        b"firstLineChars",
    ];
    for part in zip_entry_names(bytes)?.into_iter().filter(|name| {
        name.starts_with("word/")
            && Path::new(name)
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
            && !name.contains("/_rels/")
    }) {
        let Ok(xml) = read_zip_text(bytes, &part) else {
            continue;
        };
        let mut reader = Reader::from_str(&xml);
        let mut buffer = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element) | Event::Empty(element))
                    if element.local_name().as_ref() == b"ind" =>
                {
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| {
                            Error::Reverse(format!(
                                "invalid indentation attribute in {part}: {error}"
                            ))
                        })?;
                        let local_name = attribute.key.local_name();
                        if UNSUPPORTED.contains(&local_name.as_ref()) {
                            if part == "word/styles.xml"
                                || part == "word/document.xml"
                                || part.starts_with("word/header")
                                || part.starts_with("word/footer")
                            {
                                continue;
                            }
                            let raw_value = attribute.value.as_ref();
                            if raw_value == b"0" || raw_value == b"+0" {
                                continue;
                            }
                            let name = String::from_utf8_lossy(local_name.as_ref());
                            let value = String::from_utf8_lossy(raw_value);
                            return Err(Error::Reverse(format!(
                                "{part} uses `w:{name}={value}` character-unit indentation, which JSX cannot represent in points; replace it with a point-based indentation (`left`/`start`/`right`/`end`/`firstLine`/`hanging`) or keep the original DOCX"
                            )));
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(error) => {
                    return Err(Error::Reverse(format!(
                        "cannot parse {part} indentation at byte {}: {error}",
                        reader.error_position()
                    )));
                }
                _ => {}
            }
            buffer.clear();
        }
    }
    Ok(())
}

#[expect(
    clippy::too_many_lines,
    reason = "raw paragraph properties are parsed in one ordered XML pass"
)]
fn collect_paragraph_indent_metadata(
    bytes: &[u8],
) -> Result<HashMap<String, Vec<ParagraphIndentMetadata>>> {
    const ATTRIBUTES: &[(&[u8], &str)] = &[
        (b"leftChars", "indentLeftChars"),
        (b"startChars", "indentStartChars"),
        (b"rightChars", "indentRightChars"),
        (b"endChars", "indentEndChars"),
        (b"firstLineChars", "firstLineChars"),
        (b"hangingChars", "hangingChars"),
    ];
    let parts = zip_entry_names(bytes)?.into_iter().filter(|name| {
        name == "word/document.xml"
            || (name.starts_with("word/header") || name.starts_with("word/footer"))
                && Path::new(name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
                && !name.contains("/_rels/")
    });
    let mut result = HashMap::new();
    for part in parts {
        let xml = read_zip_text(bytes, &part)?;
        let mut reader = Reader::from_str(&xml);
        let mut buffer = Vec::new();
        let mut paragraphs = Vec::new();
        let mut active_paragraphs = Vec::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element)) if element.local_name().as_ref() == b"p" => {
                    paragraphs.push(ParagraphIndentMetadata::default());
                    active_paragraphs.push(paragraphs.len() - 1);
                }
                Ok(Event::Empty(element)) if element.local_name().as_ref() == b"p" => {
                    paragraphs.push(ParagraphIndentMetadata::default());
                }
                Ok(Event::Start(element) | Event::Empty(element))
                    if element.local_name().as_ref() == b"ind" =>
                {
                    if let Some(index) = active_paragraphs.last().copied() {
                        for attribute in element.attributes() {
                            let attribute = attribute.map_err(|error| {
                                Error::Reverse(format!(
                                    "invalid indentation attribute in {part}: {error}"
                                ))
                            })?;
                            let local_name = attribute.key.local_name();
                            let Some((_, target)) = ATTRIBUTES
                                .iter()
                                .find(|(name, _)| *name == local_name.as_ref())
                            else {
                                continue;
                            };
                            let value = attribute
                                .decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| {
                                    Error::Reverse(format!(
                                        "invalid w:{} value in {part}: {error}",
                                        String::from_utf8_lossy(local_name.as_ref())
                                    ))
                                })?;
                            let value = value.parse::<i64>().map_err(|_| Error::Reverse(format!("{part} has non-integer w:{} value `{value}`; repair the character-unit indentation", String::from_utf8_lossy(local_name.as_ref()))))?;
                            if paragraphs[index]
                                .values
                                .insert((*target).to_owned(), value)
                                .is_some()
                            {
                                return Err(Error::Reverse(format!(
                                    "{part} contains duplicate w:{} attributes in one paragraph",
                                    String::from_utf8_lossy(local_name.as_ref())
                                )));
                            }
                        }
                    }
                }
                Ok(Event::Start(element) | Event::Empty(element))
                    if element.local_name().as_ref() == b"suppressLineNumbers" =>
                {
                    if let Some(index) = active_paragraphs.last().copied() {
                        for attribute in element.attributes() {
                            let attribute = attribute.map_err(|error| {
                                Error::Reverse(format!(
                                    "invalid w:suppressLineNumbers attribute in {part}: {error}"
                                ))
                            })?;
                            if attribute.key.local_name().as_ref() != b"val" {
                                return Err(Error::Reverse(format!(
                                    "{part} has unsupported attribute `{}` on w:suppressLineNumbers; remove it or extend the line-number suppression spec",
                                    String::from_utf8_lossy(attribute.key.local_name().as_ref())
                                )));
                            }
                        }
                        let value = xml_attribute(&reader, &element, b"val")?
                            .as_deref()
                            .map(parse_on_off)
                            .transpose()?
                            .unwrap_or(true);
                        if paragraphs[index]
                            .suppress_line_numbers
                            .replace(value)
                            .is_some()
                        {
                            return Err(Error::Reverse(format!(
                                "{part} contains duplicate w:suppressLineNumbers elements in one paragraph"
                            )));
                        }
                    }
                }
                Ok(Event::End(element)) if element.local_name().as_ref() == b"p" => {
                    if let Some(index) = active_paragraphs.pop() {
                        validate_raw_character_indent(&paragraphs[index].values, "paragraph")?;
                    }
                }
                Ok(Event::Eof) => break,
                Err(error) => {
                    return Err(Error::Reverse(format!(
                        "cannot parse {part} paragraph indentation at byte {}: {error}",
                        reader.error_position()
                    )));
                }
                _ => {}
            }
            buffer.clear();
        }
        result.insert(part, paragraphs);
    }
    Ok(result)
}

#[expect(
    clippy::too_many_lines,
    reason = "This parser keeps paragraph, pPr, and field-instruction state synchronized in one streaming pass."
)]
fn collect_index_field_paragraph_metadata(
    bytes: &[u8],
) -> Result<HashMap<String, Vec<Map<String, Value>>>> {
    let xml = read_zip_text(bytes, "word/document.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut result: HashMap<String, Vec<Map<String, Value>>> = HashMap::new();
    let mut paragraph: Option<RawIndexFieldParagraph> = None;
    let mut paragraph_properties_depth = 0_usize;
    let mut in_paragraph_run_properties = false;
    let mut in_instruction = false;
    let mut instruction_seen = false;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"p" => {
                paragraph = Some((String::new(), Vec::new()));
                paragraph_properties_depth = 0;
                in_paragraph_run_properties = false;
                in_instruction = false;
                instruction_seen = false;
            }
            Ok(Event::Start(element)) if paragraph.is_some() => {
                let name = element.local_name();
                if name.as_ref() == b"pPr" {
                    paragraph_properties_depth = 1;
                } else if paragraph_properties_depth > 0 && name.as_ref() == b"rPr" {
                    let attributes = raw_xml_attributes(&reader, &element)?;
                    if let Some((_, entries)) = paragraph.as_mut() {
                        entries.push(("rPr".to_owned(), attributes));
                    }
                    in_paragraph_run_properties = true;
                    paragraph_properties_depth += 1;
                } else if !instruction_seen
                    && (name.as_ref() == b"instrText" || name.as_ref() == b"delInstrText")
                {
                    in_instruction = true;
                } else if paragraph_properties_depth > 0 {
                    let attributes = raw_xml_attributes(&reader, &element)?;
                    if let Some((_, entries)) = paragraph.as_mut() {
                        let name = String::from_utf8_lossy(name.as_ref()).into_owned();
                        entries.push((
                            if in_paragraph_run_properties {
                                format!("rPr/{name}")
                            } else {
                                name
                            },
                            attributes,
                        ));
                    }
                    paragraph_properties_depth += 1;
                }
            }
            Ok(Event::Empty(element)) if paragraph.is_some() => {
                let name = element.local_name();
                if name.as_ref() == b"pPr" {
                    paragraph_properties_depth = 0;
                } else if paragraph_properties_depth > 0 && name.as_ref() == b"rPr" {
                    let attributes = raw_xml_attributes(&reader, &element)?;
                    if let Some((_, entries)) = paragraph.as_mut() {
                        entries.push(("rPr".to_owned(), attributes));
                    }
                    paragraph_properties_depth += 1;
                } else if name.as_ref() == b"instrText" || name.as_ref() == b"delInstrText" {
                    in_instruction = false;
                    instruction_seen = true;
                } else if paragraph_properties_depth > 0 {
                    let attributes = raw_xml_attributes(&reader, &element)?;
                    if let Some((_, entries)) = paragraph.as_mut() {
                        let name = String::from_utf8_lossy(name.as_ref()).into_owned();
                        entries.push((
                            if in_paragraph_run_properties {
                                format!("rPr/{name}")
                            } else {
                                name
                            },
                            attributes,
                        ));
                    }
                }
            }
            Ok(Event::Text(text)) if paragraph.is_some() && in_instruction => {
                let decoded = text.decode().map_err(|error| {
                    Error::Reverse(format!("cannot decode index field instruction: {error}"))
                })?;
                let value = quick_xml::escape::unescape(&decoded).map_err(|error| {
                    Error::Reverse(format!("cannot unescape index field instruction: {error}"))
                })?;
                if let Some((instruction, _)) = paragraph.as_mut() {
                    instruction.push_str(&value);
                }
            }
            Ok(Event::GeneralRef(reference)) if paragraph.is_some() && in_instruction => {
                let entity = reference.decode().map_err(|error| {
                    Error::Reverse(format!("cannot decode index field entity: {error}"))
                })?;
                if let Some(value) = quick_xml::escape::resolve_predefined_entity(&entity) {
                    if let Some((instruction, _)) = paragraph.as_mut() {
                        instruction.push_str(value);
                    }
                } else if let Some(value) = reference.resolve_char_ref().map_err(|error| {
                    Error::Reverse(format!(
                        "cannot decode index field character reference: {error}"
                    ))
                })? {
                    if let Some((instruction, _)) = paragraph.as_mut() {
                        instruction.push(value);
                    }
                } else {
                    return Err(Error::Reverse(format!(
                        "index field instruction contains unsupported entity `&{entity};`"
                    )));
                }
            }
            Ok(Event::End(element)) if paragraph.is_some() => {
                let name = element.local_name();
                if name.as_ref() == b"p" {
                    if let Some((instruction, entries)) = paragraph.take()
                        && instruction.trim_start().starts_with("TOC")
                    {
                        let properties = raw_index_field_paragraph_properties(&entries)?;
                        result
                            .entry(instruction.trim().to_owned())
                            .or_default()
                            .push(properties);
                    }
                    paragraph_properties_depth = 0;
                    in_instruction = false;
                    instruction_seen = false;
                } else if name.as_ref() == b"instrText" || name.as_ref() == b"delInstrText" {
                    in_instruction = false;
                    instruction_seen = true;
                } else if name.as_ref() == b"pPr" {
                    paragraph_properties_depth = 0;
                    in_paragraph_run_properties = false;
                } else if name.as_ref() == b"rPr" {
                    in_paragraph_run_properties = false;
                    paragraph_properties_depth = paragraph_properties_depth.saturating_sub(1);
                } else if paragraph_properties_depth > 1 {
                    paragraph_properties_depth -= 1;
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot read index field paragraph properties from word/document.xml: {error}"
                )));
            }
        }
        buffer.clear();
    }
    Ok(result)
}

fn raw_xml_attributes(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<BTreeMap<String, String>> {
    let mut attributes = BTreeMap::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::Reverse(format!(
                "malformed attribute on index field property: {error}"
            ))
        })?;
        let key = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::Reverse(format!("invalid value on index field property: {error}"))
            })?
            .into_owned();
        attributes.insert(key, value);
    }
    Ok(attributes)
}

#[expect(
    clippy::too_many_lines,
    reason = "The whitelist is intentionally explicit so unsupported OOXML is never silently discarded."
)]
fn raw_index_field_paragraph_properties(
    entries: &RawIndexParagraphProperties,
) -> Result<Map<String, Value>> {
    let mut properties = Map::new();
    let mut special_indent = None;
    let mut character_indents = BTreeMap::new();
    for (element, attributes) in entries {
        if let Some(run_element) = element.strip_prefix("rPr/") {
            raw_index_run_property(run_element, attributes, &mut properties)?;
            continue;
        }
        let mut recognized = true;
        match element.as_str() {
            "rPr" | "tabs" if attributes.is_empty() => continue,
            "pStyle" => {
                raw_string_property(attributes, "val", "style", &mut properties);
            }
            "jc" => raw_string_property(attributes, "val", "align", &mut properties),
            "spacing" => {
                for (source, target) in [
                    ("before", "spacingBefore"),
                    ("after", "spacingAfter"),
                    ("line", "lineSpacing"),
                ] {
                    if let Some(value) = attributes.get(source) {
                        let twips = parse_raw_i64(value, source)?;
                        properties.insert(target.to_owned(), points_value(twips)?);
                    }
                }
                for (source, target) in [
                    ("beforeLines", "spacingBeforeLines"),
                    ("afterLines", "spacingAfterLines"),
                ] {
                    if let Some(value) = attributes.get(source) {
                        properties.insert(
                            target.to_owned(),
                            Value::from(parse_raw_i64(value, source)?),
                        );
                    }
                }
                if let Some(value) = attributes.get("lineRule") {
                    properties.insert("lineRule".to_owned(), Value::String(value.clone()));
                }
                if attributes.keys().any(|key| {
                    ![
                        "before",
                        "after",
                        "line",
                        "beforeLines",
                        "afterLines",
                        "lineRule",
                    ]
                    .contains(&key.as_str())
                }) {
                    raw_mark_unsupported(&mut properties, "spacing attributes");
                }
            }
            "ind" => {
                for (source, target) in [("left", "indentLeft"), ("right", "indentRight")] {
                    if let Some(value) = attributes.get(source) {
                        properties.insert(
                            target.to_owned(),
                            points_value(parse_raw_i64(value, source)?)?,
                        );
                    }
                }
                for (source, target) in [("firstLine", "firstLine"), ("hanging", "hanging")] {
                    if let Some(value) = attributes.get(source) {
                        if special_indent.is_some() {
                            raw_mark_unsupported(&mut properties, "conflicting paragraph indents");
                        }
                        special_indent = Some((target, parse_raw_i64(value, source)?));
                    }
                }
                for (source, target) in [
                    ("leftChars", "indentLeftChars"),
                    ("startChars", "indentStartChars"),
                    ("rightChars", "indentRightChars"),
                    ("endChars", "indentEndChars"),
                    ("firstLineChars", "firstLineChars"),
                    ("hangingChars", "hangingChars"),
                ] {
                    if let Some(value) = attributes.get(source) {
                        let value = parse_raw_i64(value, source)?;
                        properties.insert(target.to_owned(), Value::from(value));
                        character_indents.insert(target.to_owned(), value);
                    }
                }
                if attributes.keys().any(|key| {
                    ![
                        "left",
                        "right",
                        "firstLine",
                        "hanging",
                        "leftChars",
                        "startChars",
                        "rightChars",
                        "endChars",
                        "firstLineChars",
                        "hangingChars",
                    ]
                    .contains(&key.as_str())
                }) {
                    raw_mark_unsupported(&mut properties, "indentation attributes");
                }
            }
            "keepNext"
            | "keepLines"
            | "pageBreakBefore"
            | "snapToGrid"
            | "widowControl"
            | "bidi"
            | "suppressLineNumbers" => {
                let value = attributes
                    .get("val")
                    .map(|value| parse_on_off(value))
                    .transpose()?
                    .unwrap_or(true);
                if matches!(
                    element.as_str(),
                    "keepNext" | "keepLines" | "pageBreakBefore"
                ) && !value
                {
                    raw_mark_unsupported(&mut properties, element);
                } else {
                    properties.insert(element.clone(), Value::Bool(value));
                }
                if attributes.keys().any(|key| key != "val") {
                    raw_mark_unsupported(&mut properties, "paragraph flag attributes");
                }
            }
            "outlineLvl" => {
                raw_number_property(attributes, "val", "outlineLevel", &mut properties)?;
            }
            "textAlignment" => {
                raw_string_property(attributes, "val", "textAlign", &mut properties);
            }
            "shd" => {
                let supported = attributes.get("val").is_none_or(|value| value == "clear")
                    && attributes.get("color").is_none_or(|value| value == "auto")
                    && attributes.contains_key("fill")
                    && attributes
                        .keys()
                        .all(|key| ["val", "color", "fill"].contains(&key.as_str()));
                if supported {
                    if let Some(fill) = attributes.get("fill") {
                        properties.insert("shading".to_owned(), Value::String(fill.clone()));
                    }
                } else {
                    raw_mark_unsupported(&mut properties, "shading");
                }
            }
            _ => recognized = false,
        }
        if !recognized {
            raw_mark_unsupported(&mut properties, element);
        }
    }
    if let Some((target, value)) = special_indent {
        properties.insert(target.to_owned(), points_value(value)?);
    }
    validate_raw_character_indent(&character_indents, "index field paragraph")?;
    if (properties.contains_key("firstLine") || properties.contains_key("firstLineChars"))
        && (properties.contains_key("hanging") || properties.contains_key("hangingChars"))
    {
        return Err(Error::Reverse(
            "index field paragraph has conflicting first-line and hanging indents; retain only one indentation mode".to_owned(),
        ));
    }
    if let Some((name, value)) =
        reverse_fonts_property(&Value::Object(properties.clone()), "font", "fonts")
    {
        properties.remove("fonts");
        properties.insert(name, value);
    }
    Ok(properties)
}

#[expect(
    clippy::too_many_lines,
    reason = "The explicit run-property whitelist keeps unsupported OOXML from being silently discarded."
)]
fn raw_index_run_property(
    element: &str,
    attributes: &BTreeMap<String, String>,
    properties: &mut Map<String, Value>,
) -> Result<()> {
    match element {
        "b" | "bCs" | "i" | "iCs" => {
            let key = if element.starts_with('b') {
                "bold"
            } else {
                "italic"
            };
            let value = attributes
                .get("val")
                .map(|value| parse_on_off(value))
                .transpose()?
                .unwrap_or(true);
            if attributes.keys().any(|attribute| attribute != "val") {
                raw_mark_unsupported(properties, &format!("{element} attributes"));
            }
            if let Some(previous) = properties.get(key).and_then(Value::as_bool)
                && previous != value
            {
                raw_mark_unsupported(properties, &format!("conflicting {key} run defaults"));
            } else {
                properties.insert(key.to_owned(), Value::Bool(value));
            }
        }
        "sz" | "szCs" => {
            let Some(value) = attributes.get("val") else {
                raw_mark_unsupported(properties, "run font size without w:val");
                return Ok(());
            };
            if attributes.keys().any(|attribute| attribute != "val") {
                raw_mark_unsupported(properties, "run font size attributes");
            }
            let half_points = parse_raw_i64(value, "run font size")?;
            if half_points < 0 {
                return Err(Error::Reverse(format!(
                    "index field paragraph run font size `{value}` is negative; use a non-negative w:sz value"
                )));
            }
            let size = scaled_decimal_value(half_points, 2)?;
            if let Some(previous) = properties.get("size")
                && previous != &size
            {
                raw_mark_unsupported(properties, "conflicting sz and szCs run defaults");
            } else {
                properties.insert("size".to_owned(), size);
            }
        }
        "color" => {
            if attributes.keys().any(|attribute| attribute != "val") {
                raw_mark_unsupported(properties, "run color attributes");
            }
            if let Some(value) = attributes.get("val") {
                if value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                    properties.insert(
                        "color".to_owned(),
                        Value::String(value.to_ascii_uppercase()),
                    );
                } else {
                    raw_mark_unsupported(properties, "non-RGB run color");
                }
            }
        }
        "spacing" => {
            if attributes.keys().any(|attribute| attribute != "val") {
                raw_mark_unsupported(properties, "run character-spacing attributes");
            }
            if let Some(value) = attributes.get("val") {
                properties.insert(
                    "characterSpacing".to_owned(),
                    points_value(parse_raw_i64(value, "character spacing")?)?,
                );
            }
        }
        "rFonts" => {
            let names = [
                ("ascii", "ascii"),
                ("hAnsi", "hiAnsi"),
                ("eastAsia", "eastAsia"),
                ("cs", "cs"),
                ("asciiTheme", "asciiTheme"),
                ("hAnsiTheme", "hiAnsiTheme"),
                ("eastAsiaTheme", "eastAsiaTheme"),
                ("cstheme", "csTheme"),
                ("csTheme", "csTheme"),
                ("hint", "hint"),
            ];
            let mut fonts = Map::new();
            for &(source, target) in &names {
                if let Some(value) = attributes.get(source) {
                    if fonts.get(target).is_some_and(|previous| previous != value) {
                        raw_mark_unsupported(properties, "conflicting complex-script theme fonts");
                    }
                    fonts.insert(target.to_owned(), Value::String(value.clone()));
                }
            }
            if attributes
                .keys()
                .any(|attribute| !names.iter().any(|(source, _)| source == attribute))
            {
                raw_mark_unsupported(properties, "run font attributes");
            }
            if !fonts.is_empty() {
                properties.insert("fonts".to_owned(), Value::Object(fonts));
            }
        }
        _ => raw_mark_unsupported(properties, &format!("run property {element}")),
    }
    Ok(())
}

fn raw_string_property(
    attributes: &BTreeMap<String, String>,
    source: &str,
    target: &str,
    properties: &mut Map<String, Value>,
) {
    if let Some(value) = attributes.get(source) {
        properties.insert(target.to_owned(), Value::String(value.clone()));
    }
    if attributes.keys().any(|key| key != source) {
        raw_mark_unsupported(properties, target);
    }
}

fn raw_number_property(
    attributes: &BTreeMap<String, String>,
    source: &str,
    target: &str,
    properties: &mut Map<String, Value>,
) -> Result<()> {
    if let Some(value) = attributes.get(source) {
        properties.insert(
            target.to_owned(),
            Value::from(parse_raw_i64(value, source)?),
        );
    }
    if attributes.keys().any(|key| key != source) {
        raw_mark_unsupported(properties, target);
    }
    Ok(())
}

fn parse_raw_i64(value: &str, property: &str) -> Result<i64> {
    value.parse::<i64>().map_err(|_| {
        Error::Reverse(format!(
            "index field paragraph `{property}` value `{value}` is not an integer; repair the OOXML value"
        ))
    })
}

fn raw_mark_unsupported(properties: &mut Map<String, Value>, property: &str) {
    let unsupported = properties
        .entry("__unsupported".to_owned())
        .or_insert_with(|| Value::Array(Vec::new()));
    if let Some(unsupported) = unsupported.as_array_mut() {
        unsupported.push(Value::String(property.to_owned()));
    }
}

fn collect_style_character_indent_metadata(
    bytes: &[u8],
) -> Result<HashMap<String, BTreeMap<String, i64>>> {
    const ATTRIBUTES: &[(&[u8], &str)] = &[
        (b"leftChars", "indentLeftChars"),
        (b"startChars", "indentStartChars"),
        (b"rightChars", "indentRightChars"),
        (b"endChars", "indentEndChars"),
        (b"firstLineChars", "firstLineChars"),
        (b"hangingChars", "hangingChars"),
    ];
    let xml = read_zip_text(bytes, "word/styles.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut current_style: Option<String> = None;
    let mut result = HashMap::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"style" => {
                current_style = xml_attribute(&reader, &element, b"styleId")?;
                if let Some(style_id) = &current_style {
                    result.entry(style_id.clone()).or_insert_with(BTreeMap::new);
                }
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"ind" =>
            {
                let Some(style_id) = current_style.as_deref() else {
                    buffer.clear();
                    continue;
                };
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Reverse(format!("invalid style indentation attribute: {error}"))
                    })?;
                    let local_name = attribute.key.local_name();
                    let Some((_, target)) = ATTRIBUTES
                        .iter()
                        .find(|(name, _)| *name == local_name.as_ref())
                    else {
                        continue;
                    };
                    let raw_value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::Reverse(format!(
                                "invalid w:{} in style `{style_id}`: {error}",
                                String::from_utf8_lossy(local_name.as_ref())
                            ))
                        })?;
                    let value = raw_value.parse::<i64>().map_err(|_| {
                        Error::Reverse(format!(
                            "style `{style_id}` has non-integer w:{} value `{raw_value}`; repair its character-unit indentation",
                            String::from_utf8_lossy(local_name.as_ref())
                        ))
                    })?;
                    let properties = result
                        .entry(style_id.to_owned())
                        .or_insert_with(BTreeMap::new);
                    if properties.insert((*target).to_owned(), value).is_some() {
                        return Err(Error::Reverse(format!(
                            "style `{style_id}` contains duplicate w:{} indentation attributes",
                            String::from_utf8_lossy(local_name.as_ref())
                        )));
                    }
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"style" => {
                if let Some(style_id) = current_style.as_deref()
                    && let Some(properties) = result.get(style_id)
                {
                    validate_raw_character_indent(properties, &format!("style `{style_id}`"))?;
                }
                current_style = None;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/styles.xml character indentation at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(result)
}

fn collect_style_line_number_suppression_metadata(bytes: &[u8]) -> Result<HashMap<String, bool>> {
    let xml = read_zip_text(bytes, "word/styles.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut current_style: Option<String> = None;
    let mut result = HashMap::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"style" => {
                current_style = xml_attribute(&reader, &element, b"styleId")?;
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"suppressLineNumbers" =>
            {
                let Some(style_id) = current_style.as_deref() else {
                    buffer.clear();
                    continue;
                };
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        Error::Reverse(format!(
                            "invalid w:suppressLineNumbers attribute in style `{style_id}`: {error}"
                        ))
                    })?;
                    if attribute.key.local_name().as_ref() != b"val" {
                        return Err(Error::Reverse(format!(
                            "style `{style_id}` has unsupported attribute `{}` on w:suppressLineNumbers; remove it or extend the style line-number suppression spec",
                            String::from_utf8_lossy(attribute.key.local_name().as_ref())
                        )));
                    }
                }
                let value = xml_attribute(&reader, &element, b"val")?
                    .as_deref()
                    .map(parse_on_off)
                    .transpose()?
                    .unwrap_or(true);
                if result.insert(style_id.to_owned(), value).is_some() {
                    return Err(Error::Reverse(format!(
                        "style `{style_id}` contains duplicate w:suppressLineNumbers elements"
                    )));
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"style" => {
                current_style = None;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/styles.xml line-number suppression at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(result)
}

fn collect_style_paragraph_xml_metadata(
    bytes: &[u8],
) -> Result<HashMap<String, Map<String, Value>>> {
    let xml = read_zip_text(bytes, "word/styles.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut current_style: Option<String> = None;
    let mut in_paragraph_properties = false;
    let mut result = HashMap::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"style" => {
                current_style = xml_attribute(&reader, &element, b"styleId")?;
            }
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"pPr" => {
                in_paragraph_properties = current_style.is_some();
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"pPr" => {
                in_paragraph_properties = false;
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if in_paragraph_properties && current_style.is_some() =>
            {
                let (name, value) = match element.local_name().as_ref() {
                    b"pageBreakBefore" => (
                        "pageBreakBefore",
                        Some(style_xml_on_off(
                            &reader,
                            &element,
                            current_style.as_deref().unwrap_or_default(),
                        )?),
                    ),
                    b"widowControl" => (
                        "widowControl",
                        Some(style_xml_on_off(
                            &reader,
                            &element,
                            current_style.as_deref().unwrap_or_default(),
                        )?),
                    ),
                    b"bidi" => (
                        "bidi",
                        Some(style_xml_on_off(
                            &reader,
                            &element,
                            current_style.as_deref().unwrap_or_default(),
                        )?),
                    ),
                    b"adjustRightInd" => {
                        let style_id = current_style.as_deref().unwrap_or_default();
                        validate_style_xml_attributes(
                            &element,
                            &[b"val"],
                            style_id,
                            "adjustRightInd",
                        )?;
                        let raw = xml_attribute(&reader, &element, b"val")?
                            .unwrap_or_else(|| "1".to_owned());
                        let value = raw.parse::<i64>().ok().filter(|value| matches!(value, 0 | 1)).ok_or_else(|| Error::Reverse(format!("style `{style_id}` has unsupported w:adjustRightInd value `{raw}`; OOXML on/off values must be 0 or 1, so repair the property or keep the original DOCX")))?;
                        ("adjustRightIndent", Some(Value::Number(value.into())))
                    }
                    b"shd" => (
                        "shading",
                        Some(parse_style_shading(
                            &reader,
                            &element,
                            current_style.as_deref().unwrap_or_default(),
                        )?),
                    ),
                    _ => ("", None),
                };
                if let Some(value) = value {
                    let style_id = current_style.as_deref().unwrap_or_default();
                    let properties = result.entry(style_id.to_owned()).or_insert_with(Map::new);
                    if properties.insert(name.to_owned(), value).is_some() {
                        return Err(Error::Reverse(format!(
                            "style `{style_id}` contains duplicate w:{name} paragraph properties"
                        )));
                    }
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"style" => {
                current_style = None;
                in_paragraph_properties = false;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/styles.xml paragraph properties at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(result)
}

#[expect(
    clippy::too_many_lines,
    reason = "the event-state parser keeps style/run scope and strict diagnostics together"
)]
fn collect_style_run_xml_metadata(bytes: &[u8]) -> Result<HashMap<String, Map<String, Value>>> {
    let xml = read_zip_text(bytes, "word/styles.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut current_style: Option<String> = None;
    let mut in_paragraph_properties = false;
    let mut in_run_properties = false;
    let mut in_run_properties_change = false;
    let mut result = HashMap::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"style" => {
                current_style = xml_attribute(&reader, &element, b"styleId")?;
            }
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"pPr" => {
                in_paragraph_properties = current_style.is_some();
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"pPr" => {
                in_paragraph_properties = false;
            }
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"rPrChange" => {
                in_run_properties_change = true;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"rPrChange" => {
                in_run_properties_change = false;
            }
            Ok(Event::Start(element))
                if element.local_name().as_ref() == b"rPr" && !in_run_properties_change =>
            {
                in_run_properties = current_style.is_some() && !in_paragraph_properties;
            }
            Ok(Event::End(element))
                if element.local_name().as_ref() == b"rPr" && !in_run_properties_change =>
            {
                in_run_properties = false;
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if in_run_properties && !in_run_properties_change && current_style.is_some() =>
            {
                let style_id = current_style.as_deref().unwrap_or_default();
                let entry = result.entry(style_id.to_owned()).or_insert_with(Map::new);
                match element.local_name().as_ref() {
                    b"fitText" => {
                        validate_style_xml_attributes(
                            &element,
                            &[b"val", b"id"],
                            style_id,
                            "fitText",
                        )?;
                        let raw_width = xml_attribute(&reader, &element, b"val")?
                            .ok_or_else(|| Error::Reverse(format!("style `{style_id}` has w:fitText without w:val; add a positive twip width")))?;
                        let width = raw_width.parse::<i64>().ok().filter(|value| *value > 0).ok_or_else(|| Error::Reverse(format!("style `{style_id}` has invalid w:fitText width `{raw_width}`; use a positive integer twip width")))?;
                        let mut fit_text = Map::new();
                        fit_text.insert("width".to_owned(), points_value(width)?);
                        if let Some(id) = xml_attribute(&reader, &element, b"id")? {
                            let id = id.parse::<u32>().map_err(|_| Error::Reverse(format!("style `{style_id}` has invalid w:fitText id `{id}`; use an unsigned 32-bit integer")))?;
                            fit_text.insert("id".to_owned(), Value::Number(id.into()));
                        }
                        entry.insert("fitText".to_owned(), Value::Object(fit_text));
                    }
                    b"w" => {
                        validate_style_xml_attributes(&element, &[b"val"], style_id, "w")?;
                        let raw = xml_attribute(&reader, &element, b"val")?.ok_or_else(|| {
                            Error::Reverse(format!(
                                "style `{style_id}` has w:w without w:val; add a stretch percentage"
                            ))
                        })?;
                        let value = raw.parse::<u32>().ok().filter(|value| (1..=600).contains(value)).ok_or_else(|| Error::Reverse(format!("style `{style_id}` has unsupported w:w value `{raw}`; use a stretch percentage from 1 through 600")))?;
                        entry.insert("stretch".to_owned(), Value::Number(value.into()));
                    }
                    b"specVanish" => {
                        validate_style_xml_attributes(&element, &[b"val"], style_id, "specVanish")?;
                        let enabled = xml_on_off(&reader, &element)?;
                        if !enabled {
                            return Err(Error::Reverse(format!(
                                "style `{style_id}` has disabled w:specVanish, which docx-jsx cannot represent because docx-rs exposes no explicit-off writer value; remove the property or preserve it outside the JSX round trip"
                            )));
                        }
                        entry.insert("specVanish".to_owned(), Value::Bool(true));
                    }
                    b"shd" => {
                        validate_style_xml_attributes(
                            &element,
                            &[b"val", b"color", b"fill"],
                            style_id,
                            "shd",
                        )?;
                        let fill = xml_attribute(&reader, &element, b"fill")?
                            .ok_or_else(|| Error::Reverse(format!("style `{style_id}` has w:shd without w:fill; add a six-digit RGB fill")))?;
                        if fill.len() != 6 || !fill.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                            return Err(Error::Reverse(format!(
                                "style `{style_id}` has unrepresentable w:shd fill `{fill}`; use six-digit RGB"
                            )));
                        }
                        let mut shading = Map::new();
                        shading.insert("fill".to_owned(), Value::String(fill.to_ascii_uppercase()));
                        if let Some(color) = xml_attribute(&reader, &element, b"color")? {
                            if color != "auto"
                                && (color.len() != 6
                                    || !color.bytes().all(|byte| byte.is_ascii_hexdigit()))
                            {
                                return Err(Error::Reverse(format!(
                                    "style `{style_id}` has unrepresentable w:shd color `{color}`; use `auto` or six-digit RGB"
                                )));
                            }
                            shading.insert(
                                "color".to_owned(),
                                Value::String(if color == "auto" {
                                    color
                                } else {
                                    color.to_ascii_uppercase()
                                }),
                            );
                        }
                        if let Some(pattern) = xml_attribute(&reader, &element, b"val")? {
                            if !SHADING_PATTERNS.contains(&pattern.as_str()) {
                                return Err(Error::Reverse(format!(
                                    "style `{style_id}` has unsupported w:shd pattern `{pattern}`; use one of the documented ShadedText patterns"
                                )));
                            }
                            shading.insert("pattern".to_owned(), Value::String(pattern));
                        }
                        entry.insert("shading".to_owned(), Value::Object(shading));
                    }
                    _ => {}
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"style" => {
                current_style = None;
                in_paragraph_properties = false;
                in_run_properties = false;
                in_run_properties_change = false;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/styles.xml run properties at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(result)
}

fn style_xml_on_off(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    style_id: &str,
) -> Result<Value> {
    validate_style_xml_attributes(element, &[b"val"], style_id, "on/off paragraph property")?;
    let value = xml_attribute(reader, element, b"val")?.as_deref().map(|value| match value {
        "1" | "true" | "on" => Ok(true),
        "0" | "false" | "off" => Ok(false),
        _ => Err(Error::Reverse(format!("style `{style_id}` has invalid on/off value `{value}`; use true, false, 1, 0, on, or off"))),
    }).transpose()?.unwrap_or(true);
    Ok(Value::Bool(value))
}

fn parse_style_shading(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    style_id: &str,
) -> Result<Value> {
    validate_style_xml_attributes(element, &[b"val", b"color", b"fill"], style_id, "shading")?;
    let pattern = xml_attribute(reader, element, b"val")?.unwrap_or_else(|| "clear".to_owned());
    let color = xml_attribute(reader, element, b"color")?.unwrap_or_else(|| "auto".to_owned());
    let fill = xml_attribute(reader, element, b"fill")?.unwrap_or_else(|| "auto".to_owned());
    if pattern != "clear" || color != "auto" || fill == "auto" {
        return Err(Error::Reverse(format!(
            "style `{style_id}` uses non-solid w:shd (val={pattern}, color={color}, fill={fill}) that JSX cannot represent; use w:val=\"clear\", w:color=\"auto\", and a six-digit RGB w:fill"
        )));
    }
    if fill.len() != 6 || !fill.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(Error::Reverse(format!(
            "style `{style_id}` has invalid w:shd fill `{fill}`; use six hexadecimal RGB digits"
        )));
    }
    Ok(Value::String(fill.to_ascii_uppercase()))
}

fn validate_style_xml_attributes(
    element: &BytesStart<'_>,
    allowed: &[&[u8]],
    style_id: &str,
    property: &str,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::Reverse(format!(
                "style `{style_id}` has invalid w:{property} attribute: {error}"
            ))
        })?;
        let name = attribute.key.local_name();
        if !allowed.iter().any(|allowed| *allowed == name.as_ref()) {
            return Err(Error::Reverse(format!(
                "style `{style_id}` has unsupported attribute `{}` on w:{property}; remove it or extend the paragraph style spec",
                String::from_utf8_lossy(name.as_ref())
            )));
        }
    }
    Ok(())
}

fn validate_raw_character_indent(values: &BTreeMap<String, i64>, owner: &str) -> Result<()> {
    for (left, right) in [
        ("indentLeftChars", "indentStartChars"),
        ("indentRightChars", "indentEndChars"),
        ("firstLineChars", "hangingChars"),
    ] {
        if values.contains_key(left) && values.contains_key(right) {
            return Err(Error::Reverse(format!(
                "{owner} has conflicting `{left}` and `{right}` values; retain only one character-unit indentation attribute"
            )));
        }
    }
    Ok(())
}

fn collect_numbering_multi_level_types(bytes: &[u8]) -> Result<HashMap<usize, String>> {
    if !zip_entry_names(bytes)?
        .iter()
        .any(|name| name == "word/numbering.xml")
    {
        return Ok(HashMap::new());
    }
    let xml = read_zip_text(bytes, "word/numbering.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut current_id = None;
    let mut output = HashMap::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"abstractNum" => {
                if current_id.is_some() {
                    return Err(Error::Reverse(
                        "word/numbering.xml has nested `w:abstractNum` elements; repair the malformed numbering definition".to_owned(),
                    ));
                }
                current_id = xml_attribute(&reader, &element, b"abstractNumId")?
                    .and_then(|value| value.parse::<usize>().ok());
                let Some(id) = current_id else {
                    return Err(Error::Reverse(
                        "word/numbering.xml has an abstract numbering definition without a numeric `w:abstractNumId`; restore its ID".to_owned(),
                    ));
                };
                if output.insert(id, String::new()).is_some() {
                    return Err(Error::Reverse(format!(
                        "word/numbering.xml repeats abstract numbering ID `{id}`; make numbering IDs unique"
                    )));
                }
            }
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"abstractNum" => {
                let id = xml_attribute(&reader, &element, b"abstractNumId")?
                    .and_then(|value| value.parse::<usize>().ok())
                    .ok_or_else(|| Error::Reverse(
                        "word/numbering.xml has an abstract numbering definition without a numeric `w:abstractNumId`; restore its ID".to_owned(),
                    ))?;
                if output.insert(id, String::new()).is_some() {
                    return Err(Error::Reverse(format!(
                        "word/numbering.xml repeats abstract numbering ID `{id}`; make numbering IDs unique"
                    )));
                }
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"multiLevelType" =>
            {
                let Some(id) = current_id else {
                    return Err(Error::Reverse(
                        "`w:multiLevelType` is outside an abstract numbering definition".to_owned(),
                    ));
                };
                let multi_level_type = xml_attribute(&reader, &element, b"val")?.ok_or_else(|| {
                    Error::Reverse(format!(
                        "abstract numbering `{id}` has `w:multiLevelType` without `w:val`; add singleLevel, multilevel, or hybridMultilevel"
                    ))
                })?;
                let recorded_type = output.get_mut(&id).ok_or_else(|| {
                    Error::Reverse(format!(
                        "abstract numbering `{id}` has a multi-level type outside its `w:abstractNum`; place it inside the definition"
                    ))
                })?;
                if !recorded_type.is_empty() {
                    return Err(Error::Reverse(format!(
                        "abstract numbering `{id}` contains multiple `w:multiLevelType` values; keep one"
                    )));
                }
                if multi_level_type.is_empty() {
                    return Err(Error::Reverse(format!(
                        "abstract numbering `{id}` has an empty `w:multiLevelType`; use singleLevel, multilevel, or hybridMultilevel"
                    )));
                }
                *recorded_type = multi_level_type;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"abstractNum" => {
                current_id.take();
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/numbering.xml multi-level type at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    if current_id.is_some() {
        return Err(Error::Reverse(
            "word/numbering.xml has an unclosed `w:abstractNum`; close the numbering definition"
                .to_owned(),
        ));
    }
    Ok(output
        .into_iter()
        .filter(|(_, value)| !value.is_empty())
        .collect())
}

#[expect(
    clippy::too_many_lines,
    reason = "The streaming parser keeps section and column event state together to preserve XML ordering."
)]
fn collect_section_columns(bytes: &[u8]) -> Result<Vec<SectionColumnsMetadata>> {
    let xml = read_zip_text(bytes, "word/document.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut current_columns: Option<SectionColumnsMetadata> = None;
    let mut inside_columns = false;
    let mut saw_columns = false;
    let mut sections = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"sectPr" => {
                if current_columns
                    .replace(SectionColumnsMetadata::default())
                    .is_some()
                {
                    return Err(Error::Reverse(
                        "word/document.xml contains nested section properties; repair the malformed `w:sectPr` structure".to_owned(),
                    ));
                }
                saw_columns = false;
            }
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(SectionColumnsMetadata::default());
            }
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"cols" => {
                if current_columns.is_some() {
                    if saw_columns {
                        return Err(Error::Reverse(
                            "section contains multiple `w:cols` elements; keep exactly one column definition per section".to_owned(),
                        ));
                    }
                    saw_columns = true;
                    current_columns = Some(parse_section_columns(&reader, &element)?);
                    inside_columns = true;
                }
            }
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"cols" => {
                if current_columns.is_some() {
                    if saw_columns {
                        return Err(Error::Reverse(
                            "section contains multiple `w:cols` elements; keep exactly one column definition per section".to_owned(),
                        ));
                    }
                    saw_columns = true;
                    current_columns = Some(parse_section_columns(&reader, &element)?);
                    if current_columns
                        .as_ref()
                        .is_some_and(|columns| columns.widths.as_ref().is_some_and(Vec::is_empty))
                    {
                        return Err(Error::Reverse(
                            "unequal-width section columns require at least one `w:col`; add column definitions or set `w:equalWidth=1`".to_owned(),
                        ));
                    }
                }
            }
            Ok(Event::Empty(element))
                if inside_columns && element.local_name().as_ref() == b"col" =>
            {
                if let Some(columns) = current_columns.as_mut()
                    && let Some(widths) = columns.widths.as_mut()
                {
                    widths.push(parse_section_column(&reader, &element)?);
                    if widths.len() > 45 {
                        return Err(Error::Reverse(
                            "section defines more than 45 individual columns; reduce `w:col` children to at most 45".to_owned(),
                        ));
                    }
                }
            }
            Ok(Event::Start(element)) if inside_columns => {
                return Err(Error::Reverse(format!(
                    "section `w:cols` contains non-empty `<{}>`; repair it to an empty column element or keep the original DOCX",
                    String::from_utf8_lossy(element.local_name().as_ref())
                )));
            }
            Ok(Event::Empty(element)) if inside_columns => {
                return Err(Error::Reverse(format!(
                    "section `w:cols` contains unsupported `<{}>` data; remove it or extend the section column spec",
                    String::from_utf8_lossy(element.local_name().as_ref())
                )));
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"cols" => {
                if current_columns
                    .as_ref()
                    .is_some_and(|columns| columns.widths.as_ref().is_some_and(Vec::is_empty))
                {
                    return Err(Error::Reverse(
                        "unequal-width section columns require at least one `w:col`; add column definitions or set `w:equalWidth=1`".to_owned(),
                    ));
                }
                inside_columns = false;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(current_columns.take().unwrap_or_default());
                inside_columns = false;
                saw_columns = false;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/document.xml section columns at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(sections)
}

fn collect_section_types(bytes: &[u8]) -> Result<Vec<Option<String>>> {
    let xml = read_zip_text(bytes, "word/document.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut in_section = false;
    let mut saw_type = false;
    let mut current_type = None;
    let mut sections = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"sectPr" => {
                if in_section {
                    return Err(Error::Reverse(
                        "word/document.xml contains nested section properties; repair the malformed `w:sectPr` structure".to_owned(),
                    ));
                }
                in_section = true;
                saw_type = false;
                current_type = None;
            }
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(None);
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if in_section && element.local_name().as_ref() == b"type" =>
            {
                if saw_type {
                    return Err(Error::Reverse(
                        "section contains multiple `w:type` elements; keep one section type"
                            .to_owned(),
                    ));
                }
                saw_type = true;
                let value = xml_attribute(&reader, &element, b"val")?.ok_or_else(|| {
                    Error::Reverse(
                        "section `w:type` is missing `w:val`; add a supported section type or remove the element".to_owned(),
                    )
                })?;
                if ![
                    "nextPage",
                    "nextColumn",
                    "continuous",
                    "evenPage",
                    "oddPage",
                ]
                .contains(&value.as_str())
                {
                    return Err(Error::Reverse(format!(
                        "section uses unsupported w:type `{value}`; use nextPage, nextColumn, continuous, evenPage, or oddPage"
                    )));
                }
                current_type = Some(value);
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(current_type.take());
                in_section = false;
                saw_type = false;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/document.xml section types at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(sections)
}

fn collect_section_line_numbering(bytes: &[u8]) -> Result<Vec<Option<Map<String, Value>>>> {
    let xml = read_zip_text(bytes, "word/document.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut in_section = false;
    let mut saw_line_numbering = false;
    let mut current = None;
    let mut sections = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"sectPr" => {
                if in_section {
                    return Err(Error::Reverse(
                        "word/document.xml contains nested section properties; repair the malformed `w:sectPr` structure".to_owned(),
                    ));
                }
                in_section = true;
                saw_line_numbering = false;
                current = None;
            }
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(None);
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if in_section && element.local_name().as_ref() == b"lnNumType" =>
            {
                if saw_line_numbering {
                    return Err(Error::Reverse(
                        "section contains multiple `w:lnNumType` elements; keep one line-numbering definition".to_owned(),
                    ));
                }
                saw_line_numbering = true;
                current = Some(parse_section_line_numbering(&reader, &element)?);
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(current.take());
                in_section = false;
                saw_line_numbering = false;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/document.xml line numbering at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(sections)
}

fn collect_section_page_numbering(bytes: &[u8]) -> Result<Vec<Option<Map<String, Value>>>> {
    let xml = read_zip_text(bytes, "word/document.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut in_section = false;
    let mut saw_page_numbering = false;
    let mut current = None;
    let mut sections = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"sectPr" => {
                if in_section {
                    return Err(Error::Reverse(
                        "word/document.xml contains nested section properties; repair the malformed `w:sectPr` structure".to_owned(),
                    ));
                }
                in_section = true;
                saw_page_numbering = false;
                current = None;
            }
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(None);
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if in_section && element.local_name().as_ref() == b"pgNumType" =>
            {
                if saw_page_numbering {
                    return Err(Error::Reverse(
                        "section contains multiple `w:pgNumType` elements; keep one page-numbering definition".to_owned(),
                    ));
                }
                saw_page_numbering = true;
                current = parse_section_page_numbering(&reader, &element)?;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(current.take());
                in_section = false;
                saw_page_numbering = false;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/document.xml page numbering at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(sections)
}

fn parse_section_page_numbering(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<Map<String, Value>>> {
    let mut properties = Map::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Reverse(format!("invalid `w:pgNumType` attribute: {error}")))?;
        let name = attribute.key.local_name();
        if ![b"start".as_slice(), b"chapStyle", b"fmt", b"chapSep"].contains(&name.as_ref()) {
            return Err(Error::Reverse(format!(
                "`w:pgNumType` has unsupported attribute `{}`; remove it or extend the `pageNumbering` spec",
                String::from_utf8_lossy(name.as_ref())
            )));
        }
    }
    if let Some(value) = xml_attribute(reader, element, b"start")? {
        let start = value.parse::<u32>().map_err(|_| {
            Error::Reverse(
                "section `w:pgNumType/@w:start` must be a non-negative 32-bit integer; correct the value or keep the original DOCX".to_owned(),
            )
        })?;
        properties.insert("start".to_owned(), Value::from(start));
    }
    if let Some(value) = xml_attribute(reader, element, b"chapStyle")? {
        if value.is_empty() {
            return Err(Error::Reverse(
                "section `w:pgNumType/@w:chapStyle` must not be empty; set a chapter style id or remove the attribute".to_owned(),
            ));
        }
        properties.insert("chapterStyle".to_owned(), Value::String(value));
    }
    if let Some(value) = xml_attribute(reader, element, b"fmt")? {
        if !PAGE_NUMBER_FORMATS.contains(&value.as_str()) {
            return Err(Error::Reverse(format!(
                "section page-number format `{value}` is not representable; use a supported ST_NumberFormat value or keep the original DOCX"
            )));
        }
        properties.insert("format".to_owned(), Value::String(value));
    }
    if let Some(value) = xml_attribute(reader, element, b"chapSep")? {
        if !PAGE_NUMBER_CHAPTER_SEPARATORS.contains(&value.as_str()) {
            return Err(Error::Reverse(format!(
                "section chapter separator `{value}` is not representable; use hyphen, period, colon, emDash, or enDash, or keep the original DOCX"
            )));
        }
        properties.insert("chapterSeparator".to_owned(), Value::String(value));
    }
    Ok((!properties.is_empty()).then_some(properties))
}

fn parse_section_line_numbering(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Map<String, Value>> {
    let mut properties = Map::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Reverse(format!("invalid `w:lnNumType` attribute: {error}")))?;
        let name = attribute.key.local_name();
        if ![b"countBy".as_slice(), b"start", b"distance", b"restart"].contains(&name.as_ref()) {
            return Err(Error::Reverse(format!(
                "`w:lnNumType` has unsupported attribute `{}`; remove it or extend the `lineNumbering` spec",
                String::from_utf8_lossy(name.as_ref())
            )));
        }
    }
    for (xml_name, property_name) in [
        (b"countBy".as_slice(), "countBy"),
        (b"start".as_slice(), "start"),
    ] {
        if let Some(value) = xml_attribute(reader, element, xml_name)? {
            let number = value.parse::<u64>().ok().filter(|number| *number > 0).ok_or_else(|| {
                Error::Reverse(format!(
                    "`w:lnNumType/@w:{property_name}` must be a positive integer; fix it to a value greater than zero"
                ))
            })?;
            properties.insert(property_name.to_owned(), Value::from(number));
        }
    }
    if let Some(value) = xml_attribute(reader, element, b"distance")? {
        let twips = value.parse::<u64>().map_err(|_| {
            Error::Reverse("`w:lnNumType/@w:distance` must be a non-negative twips measure; use a non-negative distance or remove the attribute".to_owned())
        })?;
        let points = if twips % 20 == 0 {
            Value::from(twips / 20)
        } else {
            let points = format_points(twips).parse::<f64>().map_err(|error| {
                Error::Reverse(format!("cannot convert line-number distance: {error}"))
            })?;
            let number = Number::from_f64(points).ok_or_else(|| {
                Error::Reverse(
                    "line-number distance is outside the supported numeric range".to_owned(),
                )
            })?;
            Value::Number(number)
        };
        properties.insert("distance".to_owned(), points);
    }
    if let Some(value) = xml_attribute(reader, element, b"restart")? {
        if !["newPage", "newSection", "continuous"].contains(&value.as_str()) {
            return Err(Error::Reverse(format!(
                "`w:lnNumType/@w:restart` uses unsupported value `{value}`; use newPage, newSection, or continuous"
            )));
        }
        properties.insert("restart".to_owned(), Value::String(value));
    }
    Ok(properties)
}

fn parse_section_columns(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<SectionColumnsMetadata> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Reverse(format!("invalid `w:cols` attribute: {error}")))?;
        if ![b"num".as_slice(), b"space", b"sep", b"equalWidth"]
            .contains(&attribute.key.local_name().as_ref())
        {
            return Err(Error::Reverse(format!(
                "section `w:cols` has unsupported attribute `{}`; remove it or extend the columns spec",
                String::from_utf8_lossy(attribute.key.local_name().as_ref())
            )));
        }
    }
    let count = xml_attribute(reader, element, b"num")?
        .map(|value| parse_section_column_integer(&value, "w:cols/@w:num"))
        .transpose()?
        .unwrap_or(1);
    if !(1..=45).contains(&count) {
        return Err(Error::Reverse(format!(
            "section column count `{count}` is outside the supported range 1–45; set `w:cols/@w:num` to a supported count"
        )));
    }
    let equal_width = xml_attribute(reader, element, b"equalWidth")?
        .as_deref()
        .map(parse_on_off)
        .transpose()?
        .unwrap_or(true);
    let separator = xml_attribute(reader, element, b"sep")?
        .as_deref()
        .map(parse_on_off)
        .transpose()?
        .unwrap_or(false);
    let spacing_twips = xml_attribute(reader, element, b"space")?
        .map(|value| parse_section_column_integer(&value, "w:cols/@w:space"))
        .transpose()?;
    Ok(SectionColumnsMetadata {
        count: usize::try_from(count)
            .map_err(|_| Error::Reverse("section column count is out of range".to_owned()))?,
        spacing_twips,
        separator,
        widths: (!equal_width).then(Vec::new),
    })
}

fn parse_section_column(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<SectionColumnWidthMetadata> {
    let mut column = SectionColumnWidthMetadata::default();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Reverse(format!("invalid `w:col` attribute: {error}")))?;
        let name = attribute.key.local_name();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::Reverse(format!(
                    "cannot decode `w:col/@w:{}`: {error}",
                    String::from_utf8_lossy(name.as_ref())
                ))
            })?;
        let value = parse_section_column_integer(
            &value,
            &format!("w:col/@w:{}", String::from_utf8_lossy(name.as_ref())),
        )?;
        match name.as_ref() {
            b"w" if column.width_twips.replace(value).is_none() => {}
            b"space" if column.spacing_twips.replace(value).is_none() => {}
            b"w" | b"space" => {
                return Err(Error::Reverse(
                    "`w:col` contains duplicate width/spacing attributes".to_owned(),
                ));
            }
            _ => {
                return Err(Error::Reverse(format!(
                    "`w:col` has unsupported attribute `{}`; remove it or extend the columns spec",
                    String::from_utf8_lossy(name.as_ref())
                )));
            }
        }
    }
    Ok(column)
}

fn parse_section_column_integer(value: &str, name: &str) -> Result<u64> {
    value.parse::<u64>().map_err(|_| {
        Error::Reverse(format!(
            "{name} must be a non-negative integer twips value; repair the column measure"
        ))
    })
}

fn section_jsx_attrs(
    property: &docx_rs::SectionProperty,
    columns: &SectionColumnsMetadata,
    section_type: Option<&str>,
    line_numbering: Option<&Map<String, Value>>,
    page_numbering: Option<&Map<String, Value>>,
) -> Result<String> {
    let data = json(property)?;
    let mut attrs = String::new();
    if let Some(section_type) = section_type {
        attrs.push_str(&attr("type", section_type));
    }
    if let Some(line_numbering) = line_numbering {
        let mut fields = Vec::new();
        for key in ["countBy", "start", "distance", "restart"] {
            if let Some(value) = line_numbering.get(key) {
                let encoded = match value {
                    Value::String(value) => format!("{value:?}"),
                    Value::Number(value) => value.to_string(),
                    _ => {
                        return Err(Error::Reverse(format!(
                            "cannot render invalid `lineNumbering.{key}` metadata"
                        )));
                    }
                };
                fields.push(format!("{key}:{encoded}"));
            }
        }
        let _ = write!(attrs, " lineNumbering={{{{{}}}}}", fields.join(","));
    }
    if let Some(page_numbering) = page_numbering {
        let fields = ["start", "chapterStyle", "format", "chapterSeparator"]
            .into_iter()
            .filter_map(|key| {
                page_numbering.get(key).map(|value| {
                    let encoded = match value {
                        Value::String(value) => format!("{value:?}"),
                        Value::Number(value) => value.to_string(),
                        _ => String::new(),
                    };
                    format!("{key}:{encoded}")
                })
            })
            .collect::<Vec<_>>();
        let _ = write!(attrs, " pageNumbering={{{{{}}}}}", fields.join(","));
    }
    if let Some(columns) = section_columns_value(columns)? {
        let encoded = serde_json::to_string(&columns).map_err(|error| {
            Error::Reverse(format!("cannot serialize section columns: {error}"))
        })?;
        let _ = write!(attrs, " columns={{{encoded}}}");
    }
    if let Some(size) = data.get("pageSize") {
        let width = size.get("w").and_then(Value::as_u64).unwrap_or(0);
        let height = size.get("h").and_then(Value::as_u64).unwrap_or(0);
        match (width, height) {
            (11_906, 16_838) | (16_838, 11_906) => attrs.push_str(r#" pageSize="A4""#),
            (12_240, 15_840) | (15_840, 12_240) => attrs.push_str(r#" pageSize="Letter""#),
            (0, 0) => {}
            _ => {
                let width_pt = format_points(width);
                let height_pt = format_points(height);
                let _ = write!(
                    attrs,
                    " pageSize={{{{width:{width_pt},height:{height_pt}}}}}"
                );
            }
        }
        let landscape = size
            .get("orient")
            .and_then(Value::as_str)
            .is_some_and(|value| value.eq_ignore_ascii_case("landscape"))
            || width > height;
        if landscape {
            attrs.push_str(r#" orientation="landscape""#);
        }
    }
    if let Some(margin) = data.get("pageMargin") {
        let mut fields = Vec::new();
        for key in [
            "top", "right", "bottom", "left", "header", "footer", "gutter",
        ] {
            if let Some(value) = margin.get(key).and_then(Value::as_i64) {
                fields.push(format!("{key}:{}", format_points(value.unsigned_abs())));
            }
        }
        if !fields.is_empty() {
            let _ = write!(attrs, " margins={{{{{}}}}}", fields.join(","));
        }
    }
    if let Some(grid) = data.get("docGrid") {
        let mut fields = Vec::new();
        if let Some(value) = grid.get("gridType").and_then(Value::as_str) {
            fields.push(format!("type:{value:?}"));
        }
        if let Some(value) = grid.get("linePitch").and_then(Value::as_u64) {
            fields.push(format!("linePitch:{}", format_points(value)));
        }
        if let Some(value) = grid.get("charSpace").and_then(Value::as_i64) {
            fields.push(format!("charSpace:{value}"));
        }
        if !fields.is_empty() {
            let _ = write!(attrs, " documentGrid={{{{{}}}}}", fields.join(","));
        }
    }
    Ok(attrs)
}

fn section_columns_value(columns: &SectionColumnsMetadata) -> Result<Option<Value>> {
    if let Some(widths) = &columns.widths {
        let widths = widths
            .iter()
            .map(|column| {
                let mut value = Map::new();
                if let Some(width) = column.width_twips {
                    let width = i64::try_from(width).map_err(|_| {
                        Error::Reverse(
                            "column width exceeds the supported integer range".to_owned(),
                        )
                    })?;
                    value.insert("width".to_owned(), points_value(width)?);
                }
                if let Some(spacing) = column.spacing_twips {
                    let spacing = i64::try_from(spacing).map_err(|_| {
                        Error::Reverse(
                            "column spacing exceeds the supported integer range".to_owned(),
                        )
                    })?;
                    value.insert("spacing".to_owned(), points_value(spacing)?);
                }
                Ok(Value::Object(value))
            })
            .collect::<Result<Vec<_>>>()?;
        let mut value = Map::new();
        value.insert("widths".to_owned(), Value::Array(widths));
        if columns.separator {
            value.insert("separator".to_owned(), Value::Bool(true));
        }
        return Ok(Some(Value::Object(value)));
    }
    if columns.count == 1 && !columns.separator {
        return Ok(None);
    }
    if columns.spacing_twips == Some(425) && !columns.separator {
        return Ok(Some(Value::from(columns.count)));
    }
    let mut value = Map::new();
    value.insert("count".to_owned(), Value::from(columns.count));
    let spacing = columns.spacing_twips.unwrap_or(0);
    let spacing = i64::try_from(spacing).map_err(|_| {
        Error::Reverse("column spacing exceeds the supported integer range".to_owned())
    })?;
    value.insert("spacing".to_owned(), points_value(spacing)?);
    if columns.separator {
        value.insert("separator".to_owned(), Value::Bool(true));
    }
    Ok(Some(Value::Object(value)))
}

fn collect_package_images(bytes: &[u8]) -> Result<ImageCatalog> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    let rels_parts = (0..archive.len())
        .filter_map(|index| {
            let file = archive.by_index(index).ok()?;
            let name = file.name().replace('\\', "/");
            is_word_rels_part(&name).then_some(name)
        })
        .collect::<Vec<_>>();
    let mut images = ImageCatalog::default();
    let mut media_src: HashMap<String, String> = HashMap::new();
    for rels_part in rels_parts {
        let Some(owner) = owner_part_from_rels(&rels_part) else {
            continue;
        };
        let Ok(mut rels_file) = archive.by_name(&rels_part) else {
            continue;
        };
        let mut rels = String::new();
        rels_file.read_to_string(&mut rels).map_err(|error| {
            Error::Reverse(format!("cannot read relationships `{rels_part}`: {error}"))
        })?;
        drop(rels_file);
        for (id, target) in relationships_ending_with(&rels, "/image")? {
            let part = resolve_word_target(&target);
            let Ok(mut file) = archive.by_name(&part) else {
                continue;
            };
            let mut image_bytes = Vec::new();
            file.read_to_end(&mut image_bytes).map_err(|error| {
                Error::Reverse(format!("cannot read image part `{part}`: {error}"))
            })?;
            drop(file);
            let src = if let Some(src) = media_src.get(&part) {
                src.clone()
            } else {
                let file_name = Path::new(&target)
                    .file_name()
                    .and_then(|value| value.to_str())
                    .filter(|value| !value.is_empty())
                    .unwrap_or("image.bin");
                let unique = unique_asset_name(file_name, &images.used_names);
                images.used_names.insert(unique.clone());
                let src = format!("media/{unique}");
                media_src.insert(part, src.clone());
                src
            };
            images.insert(
                &owner,
                id,
                ImageAsset {
                    src,
                    bytes: image_bytes,
                },
            );
        }
    }
    Ok(images)
}

fn collect_header_footer_parts(bytes: &[u8]) -> Result<HashMap<String, String>> {
    let Ok(rels) = read_zip_text(bytes, "word/_rels/document.xml.rels") else {
        return Ok(HashMap::new());
    };
    let mut parts = HashMap::new();
    for suffix in ["/header", "/footer"] {
        for (id, target) in relationships_ending_with(&rels, suffix)? {
            parts.insert(id, resolve_word_target(&target));
        }
    }
    Ok(parts)
}

fn is_word_rels_part(name: &str) -> bool {
    name.starts_with("word/_rels/")
        && Path::new(name)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("rels"))
}

fn owner_part_from_rels(rels_path: &str) -> Option<String> {
    let file_name = Path::new(rels_path).file_name()?.to_str()?;
    let owner_name = file_name.strip_suffix(".rels")?;
    Some(format!("word/{owner_name}"))
}

fn resolve_word_target(target: &str) -> String {
    let target = target.trim_start_matches('/').replace('\\', "/");
    if target.starts_with("word/") {
        target
    } else {
        format!("word/{target}")
    }
}

fn relationships_ending_with(rels: &str, type_suffix: &str) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(rels);
    reader.config_mut().trim_text(true);
    let mut output = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"Relationship" =>
            {
                let rel_type = xml_attribute(&reader, &element, b"Type")?;
                if rel_type
                    .as_deref()
                    .is_some_and(|value| value.ends_with(type_suffix))
                    && let (Some(id), Some(target)) = (
                        xml_attribute(&reader, &element, b"Id")?,
                        xml_attribute(&reader, &element, b"Target")?,
                    )
                {
                    output.push((id, target.replace('\\', "/")));
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse package relationships: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(output)
}

fn collect_images(
    docx: &docx_rs::Docx,
    used_names: &BTreeSet<String>,
) -> HashMap<String, ImageAsset> {
    let mut images = HashMap::new();
    let mut used_names = used_names.clone();
    for (id, path, image, _) in &docx.images {
        let file_name = Path::new(path)
            .file_name()
            .and_then(|value| value.to_str())
            .filter(|value| !value.is_empty())
            .unwrap_or("image.bin");
        let unique = unique_asset_name(file_name, &used_names);
        used_names.insert(unique.clone());
        images.insert(
            id.clone(),
            ImageAsset {
                src: format!("media/{unique}"),
                bytes: image.0.clone(),
            },
        );
    }
    images
}

fn unique_asset_name(file_name: &str, used: &BTreeSet<String>) -> String {
    if !used.contains(file_name) {
        return file_name.to_owned();
    }
    let path = Path::new(file_name);
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("image");
    let ext = path
        .extension()
        .and_then(|value| value.to_str())
        .map_or_else(String::new, |value| format!(".{value}"));
    (2usize..=used.len().saturating_add(2))
        .map(|index| format!("{stem}-{index}{ext}"))
        .find(|candidate| !used.contains(candidate))
        .expect("image name space should not be exhausted")
}

struct ThemeColorPlan {
    styles: HashMap<String, ThemeColorInfo>,
    runs: Vec<ThemeColorInfo>,
}

fn parse_theme_colors(bytes: &[u8]) -> Result<ThemeColorPlan> {
    Ok(ThemeColorPlan {
        styles: parse_style_theme_colors(bytes)?,
        runs: parse_document_run_theme_colors(bytes)?,
    })
}

fn parse_style_theme_colors(bytes: &[u8]) -> Result<HashMap<String, ThemeColorInfo>> {
    let Ok(xml) = read_zip_text(bytes, "word/styles.xml") else {
        return Ok(HashMap::new());
    };
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);
    let mut output = HashMap::new();
    let mut current_id: Option<String> = None;
    let mut in_style_run = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                let name = element.local_name();
                match name.as_ref() {
                    b"style" => {
                        current_id = xml_attribute(&reader, &element, b"styleId")?;
                        in_style_run = false;
                    }
                    b"rPr" if current_id.is_some() => in_style_run = true,
                    b"color" if in_style_run => {
                        if let Some(id) = current_id.clone()
                            && let Some(theme) = theme_color_from_element(&reader, &element)?
                        {
                            output.insert(id, theme);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(element)) => {
                let name = element.local_name();
                if name.as_ref() == b"rPr" {
                    in_style_run = false;
                }
                if name.as_ref() == b"style" {
                    current_id = None;
                    in_style_run = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse theme colors in word/styles.xml: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(output)
}

fn parse_document_run_theme_colors(bytes: &[u8]) -> Result<Vec<ThemeColorInfo>> {
    let Ok(xml) = read_zip_text(bytes, "word/document.xml") else {
        return Ok(Vec::new());
    };
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);
    let mut output = Vec::new();
    let mut in_run = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                let name = element.local_name();
                match name.as_ref() {
                    b"r" => in_run = true,
                    b"color" if in_run => {
                        if let Some(theme) = theme_color_from_element(&reader, &element)? {
                            output.push(theme);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"r" => {
                in_run = false;
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse theme colors in word/document.xml: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(output)
}

fn theme_color_from_element(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<ThemeColorInfo>> {
    let theme_color = xml_attribute(reader, element, b"themeColor")?;
    let theme_shade = xml_attribute(reader, element, b"themeShade")?;
    let theme_tint = xml_attribute(reader, element, b"themeTint")?;
    if theme_color.is_none() && theme_shade.is_none() && theme_tint.is_none() {
        return Ok(None);
    }
    Ok(Some(ThemeColorInfo {
        val: xml_attribute(reader, element, b"val")?,
        theme_color,
        theme_shade,
        theme_tint,
    }))
}

fn reverse_image_anchor_attrs(picture: &Pic) -> String {
    let mut attrs = format!(" layoutInCell={{{}}}", picture.layout_in_cell);
    if picture.simple_pos {
        let _ = write!(
            attrs,
            " simplePosition={{{{ x: {}, y: {} }}}}",
            picture.simple_pos_x, picture.simple_pos_y
        );
    }
    if picture.allow_overlap {
        attrs.push_str(" allowOverlap");
    }
    attrs.push_str(&reverse_drawing_position("positionH", picture.position_h));
    attrs.push_str(&reverse_drawing_position("positionV", picture.position_v));
    let from_h = match picture.relative_from_h {
        RelativeFromHType::Character => "character",
        RelativeFromHType::Column => "column",
        RelativeFromHType::InsideMargin => "insideMargin",
        RelativeFromHType::LeftMargin => "leftMargin",
        RelativeFromHType::OutsizeMargin => "outsideMargin",
        RelativeFromHType::Page => "page",
        RelativeFromHType::RightMargin => "rightMargin",
        RelativeFromHType::Margin => "margin",
    };
    if from_h != "margin" {
        attrs.push_str(&attr("relativeFromH", from_h));
    }
    let from_v = match picture.relative_from_v {
        RelativeFromVType::BottomMargin => "bottomMargin",
        RelativeFromVType::InsideMargin => "insideMargin",
        RelativeFromVType::Line => "line",
        RelativeFromVType::OutsizeMargin => "outsideMargin",
        RelativeFromVType::Page => "page",
        RelativeFromVType::Paragraph => "paragraph",
        RelativeFromVType::TopMargin => "topMargin",
        RelativeFromVType::Margin => "margin",
    };
    if from_v != "margin" {
        attrs.push_str(&attr("relativeFromV", from_v));
    }
    if picture.dist_t != 0 {
        attrs.push_str(&emu_point_attr("distanceTop", i64::from(picture.dist_t)));
    }
    if picture.dist_b != 0 {
        attrs.push_str(&emu_point_attr("distanceBottom", i64::from(picture.dist_b)));
    }
    if picture.dist_l != 0 {
        attrs.push_str(&emu_point_attr("distanceLeft", i64::from(picture.dist_l)));
    }
    if picture.dist_r != 0 {
        attrs.push_str(&emu_point_attr("distanceRight", i64::from(picture.dist_r)));
    }
    if picture.relative_height != 0 && picture.relative_height != 190_500 {
        let _ = write!(attrs, " relativeHeight={{{}}}", picture.relative_height);
    }
    attrs
}

fn reverse_drawing_position(name: &str, position: DrawingPosition) -> String {
    match position {
        DrawingPosition::Offset(0) => String::new(),
        DrawingPosition::Offset(value) => emu_point_attr(name, i64::from(value)),
        DrawingPosition::Align(alignment) => {
            let token = match alignment {
                PicAlign::Left => "left",
                PicAlign::Right => "right",
                PicAlign::Center => "center",
                PicAlign::Top => "top",
                PicAlign::Bottom => "bottom",
            };
            attr(name, token)
        }
    }
}

fn emu_point_attr(name: &str, emu: i64) -> String {
    const EMU_PER_POINT: i64 = 12_700;
    let sign = if emu < 0 { "-" } else { "" };
    let absolute = emu.unsigned_abs();
    let whole = absolute / EMU_PER_POINT as u64;
    let remainder = absolute % EMU_PER_POINT as u64;
    if remainder == 0 {
        format!(" {name}={{{sign}{whole}}}")
    } else {
        let hundredths = remainder * 100 / EMU_PER_POINT as u64;
        format!(" {name}={{{sign}{whole}.{hundredths:02}}}")
    }
}

fn paragraph_num_id(paragraph: &docx_rs::Paragraph) -> Option<usize> {
    let property = json(&paragraph.property).ok()?;
    property
        .get("numberingProperty")
        .and_then(|value| value.get("id"))
        .and_then(Value::as_u64)
        .and_then(|value| usize::try_from(value).ok())
}

fn nested_block_num_id(child: &NestedBlock<'_>) -> Option<usize> {
    nested_block_paragraph(child)
        .ok()
        .and_then(paragraph_num_id)
}

fn nested_block_paragraph<'a>(child: &'a NestedBlock<'a>) -> Result<&'a docx_rs::Paragraph> {
    match child {
        NestedBlock::Child(DocumentChild::Paragraph(paragraph)) => Ok(paragraph),
        _ => Err(Error::Reverse(
            "numbered list items must be paragraphs; split non-paragraph children out of the list"
                .to_owned(),
        )),
    }
}

fn reverse_list_definition(
    definition: &AbstractNumbering,
    multi_level_type: Option<&str>,
) -> Result<String> {
    let mut attrs = if definition.levels.is_empty()
        && (definition.style_link.is_some() || definition.num_style_link.is_some())
    {
        String::new()
    } else if let Some(attrs) = collapse_list_shorthand(definition) {
        attrs
    } else {
        let mut levels = Vec::new();
        for level in &definition.levels {
            levels.push(reverse_level_definition(level)?);
        }
        let encoded = serde_json::to_string(&levels)
            .map_err(|error| Error::Reverse(format!("cannot serialize list levels: {error}")))?;
        format!(" levels={{{encoded}}}")
    };
    if let Some(multi_level_type) = multi_level_type {
        if !["singleLevel", "multilevel", "hybridMultilevel"].contains(&multi_level_type) {
            return Err(Error::Reverse(format!(
                "numbering uses unsupported w:multiLevelType `{multi_level_type}`; use singleLevel, multilevel, or hybridMultilevel"
            )));
        }
        attrs.push_str(&attr("multiLevelType", multi_level_type));
    }
    if let Some(style_link) = &definition.style_link {
        attrs.push_str(&attr("styleLink", style_link));
    }
    if let Some(num_style_link) = &definition.num_style_link {
        attrs.push_str(&attr("numStyleLink", num_style_link));
    }
    Ok(attrs)
}

fn collapse_list_shorthand(definition: &AbstractNumbering) -> Option<String> {
    if definition.levels.len() != 9 {
        return None;
    }
    let first = definition.levels.first()?;
    let start = json(&first.start)
        .ok()
        .and_then(|value| value.as_u64())
        .and_then(|value| usize::try_from(value).ok())
        .unwrap_or(1);
    let format = json(&first.format).ok()?;
    let format = format.as_str()?;
    let kind = match format {
        "decimal" => "ordered",
        "bullet" => "bullet",
        _ => return None,
    };
    for (index, level) in definition.levels.iter().enumerate() {
        if level.level != index {
            return None;
        }
        let level_start = json(&level.start)
            .ok()
            .and_then(|value| value.as_u64())
            .and_then(|value| usize::try_from(value).ok())
            .unwrap_or(1);
        if level_start != start {
            return None;
        }
        let level_format = json(&level.format).ok()?;
        if level_format.as_str()? != format {
            return None;
        }
        let text = json(&level.text).ok()?;
        let expected = if kind == "ordered" {
            format!("%{}.", index + 1)
        } else {
            "•".to_owned()
        };
        if text.as_str()? != expected {
            return None;
        }
        let jc = json(&level.jc).ok()?;
        if jc.as_str()? != "left" {
            return None;
        }
        if json(&level.suffix).ok()?.as_str()? != "tab" {
            return None;
        }
        if level.pstyle.is_some() || level.level_restart.is_some() || level.is_lgl.is_some() {
            return None;
        }
        let property = json(&level.paragraph_property).ok()?;
        let indent = property.get("indent")?;
        let left = indent.get("start").and_then(Value::as_i64)?;
        if left != i64::try_from((index + 1) * 720).ok()? {
            return None;
        }
        let special = indent.get("specialIndent")?;
        if special.get("type").and_then(Value::as_str)? != "hanging"
            || special.get("val").and_then(Value::as_i64)? != 360
        {
            return None;
        }
        let run = json(&level.run_property).ok()?;
        if run
            .as_object()
            .is_none_or(|object| object.values().any(|value| !value.is_null()))
        {
            return None;
        }
    }
    let mut attrs = attr("type", kind);
    if kind == "ordered" && start != 1 {
        let _ = write!(attrs, " start={{{start}}}");
    }
    Some(attrs)
}

fn reverse_level_definition(level: &Level) -> Result<Value> {
    let source = json(level)?;
    let mut output = Map::new();
    if let Some(value) = source.get("format").and_then(Value::as_str) {
        output.insert("format".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source.get("text").and_then(Value::as_str) {
        output.insert("text".to_owned(), Value::String(value.to_owned()));
    }
    let start = source
        .get("start")
        .and_then(Value::as_u64)
        .and_then(|value| usize::try_from(value).ok());
    if let Some(start) = start {
        output.insert("start".to_owned(), Value::Number(start.into()));
    }
    if let Some(value) = source.get("jc").and_then(Value::as_str)
        && value != "left"
    {
        output.insert("align".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source.get("suffix").and_then(Value::as_str)
        && value != "tab"
    {
        output.insert("suffix".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source
        .get("pstyle")
        .and_then(scalar_string)
        .or_else(|| nested_string(&source, &["pstyle", "val"]))
    {
        output.insert("paragraphStyle".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source
        .get("levelRestart")
        .and_then(Value::as_u64)
        .or_else(|| nested(&source, &["levelRestart"]).and_then(Value::as_u64))
    {
        output.insert("restart".to_owned(), Value::Number(value.into()));
    }
    if source.get("isLgl").is_some() {
        output.insert("legal".to_owned(), Value::Bool(true));
    }
    if let Some(indent) = nested(&source, &["paragraphProperty", "indent"]) {
        if let Some(value) = indent.get("start").and_then(Value::as_i64) {
            output.insert("indentLeft".to_owned(), points_value(value)?);
        }
        if let Some(value) = indent.get("end").and_then(Value::as_i64) {
            output.insert("indentRight".to_owned(), points_value(value)?);
        }
        if let Some(special) = indent.get("specialIndent")
            && let (Some(kind), Some(value)) = (
                special.get("type").and_then(Value::as_str),
                special.get("val").and_then(Value::as_i64),
            )
        {
            let key = if kind == "hanging" {
                "hanging"
            } else {
                "firstLine"
            };
            output.insert(key.to_owned(), points_value(value)?);
        }
    }
    if let Some(run) = reverse_run_properties(source.get("runProperty").unwrap_or(&Value::Null))? {
        for (key, value) in run {
            output.insert(key, value);
        }
    }
    Ok(Value::Object(output))
}

fn format_points(twips: u64) -> String {
    let whole = twips / 20;
    let remainder = twips % 20;
    if remainder == 0 {
        whole.to_string()
    } else {
        let fraction = format!("{:02}", remainder * 5);
        format!("{whole}.{}", fraction.trim_end_matches('0'))
    }
}

fn json(value: &impl Serialize) -> Result<Value> {
    serde_json::to_value(value).map_err(|error| Error::Reverse(error.to_string()))
}

#[derive(Default)]
struct StyleMetadata {
    based_on: Option<String>,
    next: Option<String>,
    link: Option<String>,
    quick_format: Option<bool>,
    ui_priority: Option<usize>,
    semi_hidden: Option<bool>,
    unhide_when_used: Option<bool>,
}

fn reverse_document_attributes(
    docx: &docx_rs::Docx,
    metadata: &HashMap<String, StyleMetadata>,
    writer: &Writer,
    bytes: &[u8],
) -> Result<String> {
    let mut attributes = String::new();
    attributes.push_str(&reverse_default_font_attrs(docx)?);
    attributes.push_str(&reverse_doc_default_metrics(docx)?);
    attributes.push_str(&reverse_core_and_custom_properties(docx, bytes)?);
    attributes.push_str(&reverse_settings_attributes(&docx.settings, bytes)?);
    if let Some(items) = collect_custom_xml_items(bytes)? {
        attributes.push_str(&jsx_prop("customXmlItems", &items)?);
    }
    if let Some(extensions) = collect_web_extensions(bytes)? {
        attributes.push_str(&jsx_prop("webExtensions", &extensions)?);
    }
    attributes.push_str(&reverse_styles_attribute(docx, metadata, writer)?);
    Ok(attributes)
}

fn reverse_default_font_attrs(docx: &docx_rs::Docx) -> Result<String> {
    let defaults = json(&docx.styles.doc_defaults)?;
    let Some(fonts) = nested(&defaults, &["runPropertyDefault", "runProperty", "fonts"])
        .and_then(Value::as_object)
    else {
        return Ok(String::new());
    };
    let mut output = Map::new();
    for key in [
        "ascii",
        "hiAnsi",
        "eastAsia",
        "cs",
        "asciiTheme",
        "hiAnsiTheme",
        "eastAsiaTheme",
        "csTheme",
        "hint",
    ] {
        if let Some(value) = fonts.get(key).and_then(Value::as_str) {
            output.insert(key.to_owned(), Value::String(value.to_owned()));
        }
    }
    if output.is_empty() {
        return Ok(String::new());
    }
    let physical =
        ["ascii", "hiAnsi", "eastAsia", "cs"].map(|key| output.get(key).and_then(Value::as_str));
    let can_collapse = physical[0].is_some()
        && physical.iter().all(|value| *value == physical[0])
        && !output
            .keys()
            .any(|key| key.ends_with("Theme") || key == "hint");
    if can_collapse {
        return Ok(attr("defaultFont", physical[0].unwrap_or_default()));
    }
    jsx_prop("defaultFonts", &Value::Object(output))
}

fn reverse_doc_default_metrics(docx: &docx_rs::Docx) -> Result<String> {
    let defaults = json(&docx.styles.doc_defaults)?;
    let mut attributes = String::new();
    if let Some(run) = nested(&defaults, &["runPropertyDefault", "runProperty"]) {
        if let Some(size) = run.get("sz").and_then(scalar_u64) {
            let size = i64::try_from(size).map_err(|_| {
                Error::Reverse("default font size exceeds supported range".to_owned())
            })?;
            attributes.push_str(&jsx_prop("defaultSize", &scaled_decimal_value(size, 2)?)?);
        }
        if let Some(spacing) = run.get("characterSpacing").and_then(scalar_i64) {
            attributes.push_str(&jsx_prop(
                "defaultCharacterSpacing",
                &points_value(spacing)?,
            )?);
        }
    }
    if let Some(spacing) = nested(
        &defaults,
        &[
            "paragraphPropertyDefault",
            "paragraphProperty",
            "lineSpacing",
        ],
    ) {
        let mut output = Map::new();
        for (source_key, target_key) in [("before", "before"), ("after", "after"), ("line", "line")]
        {
            if let Some(value) = spacing.get(source_key).and_then(scalar_i64) {
                output.insert(target_key.to_owned(), points_value(value)?);
            }
        }
        for (source_key, target_key) in
            [("beforeLines", "beforeLines"), ("afterLines", "afterLines")]
        {
            if let Some(value) = spacing.get(source_key).and_then(Value::as_u64) {
                output.insert(target_key.to_owned(), Value::Number(value.into()));
            }
        }
        copy_string(spacing, &mut output, "lineRule", "lineRule");
        if !output.is_empty() {
            attributes.push_str(&jsx_prop("defaultLineSpacing", &Value::Object(output))?);
        }
    }
    Ok(attributes)
}

fn reverse_core_and_custom_properties(docx: &docx_rs::Docx, bytes: &[u8]) -> Result<String> {
    let mut attributes = String::new();
    for (name, value) in parse_core_properties(bytes)? {
        match value {
            Value::String(value) => attributes.push_str(&attr(&name, &value)),
            value => attributes.push_str(&jsx_prop(&name, &value)?),
        }
    }
    if !docx.doc_props.custom.properties.is_empty() {
        let mut properties = Map::new();
        let mut names = docx
            .doc_props
            .custom
            .properties
            .keys()
            .cloned()
            .collect::<Vec<_>>();
        names.sort();
        for name in names {
            let Some(value) = docx.doc_props.custom.properties.get(&name) else {
                continue;
            };
            if name.is_empty() {
                return Err(Error::Reverse(
                    "custom property name is empty; give every `docProps/custom.xml` property a name"
                        .to_owned(),
                ));
            }
            properties.insert(name, Value::String(value.clone()));
        }
        attributes.push_str(&jsx_prop("customProperties", &Value::Object(properties))?);
    }
    Ok(attributes)
}

fn reverse_settings_attributes(settings: &docx_rs::Settings, bytes: &[u8]) -> Result<String> {
    let source = json(settings)?;
    let mut attributes = String::new();
    if let Some(id) = source
        .get("docId")
        .and_then(Value::as_str)
        .filter(|id| !id.is_empty())
    {
        attributes.push_str(&attr("documentId", id));
    }
    if let Some(tab_stop) = source
        .get("defaultTabStop")
        .and_then(Value::as_u64)
        .filter(|value| *value != 840)
    {
        let tab_stop = i64::try_from(tab_stop)
            .map_err(|_| Error::Reverse("default tab stop exceeds supported range".to_owned()))?;
        attributes.push_str(&jsx_prop("defaultTabStop", &points_value(tab_stop)?)?);
    }
    if let Some(protection) = collect_document_protection(bytes)? {
        attributes.push_str(&jsx_prop("documentProtection", &protection)?);
    }
    if let Some(variables) = source.get("docVars").and_then(Value::as_array) {
        let mut output = Map::new();
        for variable in variables {
            let Some(name) = variable
                .get("name")
                .and_then(Value::as_str)
                .filter(|name| !name.is_empty())
            else {
                return Err(Error::Reverse(
                    "document variable is missing a name; restore `w:docVar/@w:name` or remove the empty variable"
                        .to_owned(),
                ));
            };
            let value = variable
                .get("val")
                .and_then(Value::as_str)
                .unwrap_or_default();
            output.insert(name.to_owned(), Value::String(value.to_owned()));
        }
        if !output.is_empty() {
            attributes.push_str(&jsx_prop("documentVariables", &Value::Object(output))?);
        }
    }
    if source.get("evenAndOddHeaders").and_then(Value::as_bool) == Some(true) {
        attributes.push_str(" evenAndOddHeaders");
    }
    if source
        .get("adjustLineHeightInTable")
        .and_then(Value::as_bool)
        == Some(true)
    {
        attributes.push_str(" adjustLineHeightInTable");
    }
    if let Some(value) = source
        .get("characterSpacingControl")
        .and_then(Value::as_str)
        .filter(|value| *value != "unsupported")
    {
        attributes.push_str(&attr("characterSpacingControl", value));
    } else if source.get("characterSpacingControl").is_some() {
        return Err(Error::Reverse(
            "settings use an unsupported `w:characterSpacingControl`; set it to `doNotCompress`, `compressPunctuation`, or `compressPunctuationAndJapaneseKana`"
                .to_owned(),
        ));
    }
    Ok(attributes)
}

fn collect_document_protection(bytes: &[u8]) -> Result<Option<Value>> {
    let Ok(xml) = read_zip_text(bytes, "word/settings.xml") else {
        return Ok(None);
    };
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut protection = None;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"documentProtection" => {
                if protection.is_some() {
                    return Err(Error::Reverse(
                        "word/settings.xml contains multiple `w:documentProtection` elements; keep exactly one or remove the duplicates".to_owned(),
                    ));
                }
                protection = Some(parse_document_protection_attributes(&reader, &element)?);
            }
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"documentProtection" => {
                return Err(Error::Reverse(
                    "`w:documentProtection` must be an empty leaf element; remove its child content".to_owned(),
                ));
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/settings.xml document protection: {error}"
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    protection
        .map(|properties| {
            if properties.is_empty() {
                return Err(Error::Reverse(
                    "`w:documentProtection` has no attributes; add supported protection settings or remove the empty element".to_owned(),
                ));
            }
            let value = Value::Object(properties);
            crate::ir::validate_document_protection(&value, "Document").map_err(|error| {
                Error::Reverse(format!(
                    "`w:documentProtection` cannot be represented by the DSL: {error}; repair the attributes to match `docx-jsx spec` or extend the DSL"
                ))
            })?;
            Ok(value)
        })
        .transpose()
}

fn parse_document_protection_attributes(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Map<String, Value>> {
    const SUPPORTED: &[&str] = &[
        "edit",
        "formatting",
        "enforcement",
        "cryptProviderType",
        "cryptAlgorithmClass",
        "cryptAlgorithmType",
        "cryptAlgorithmSid",
        "cryptSpinCount",
        "cryptProvider",
        "algIdExt",
        "algIdExtSource",
        "cryptProviderTypeExt",
        "cryptProviderTypeExtSource",
        "hash",
        "salt",
        "algorithmName",
        "hashValue",
        "saltValue",
        "spinCount",
    ];
    let mut properties = Map::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::Reverse(format!("invalid `w:documentProtection` attribute: {error}"))
        })?;
        let name = String::from_utf8_lossy(attribute.key.local_name().as_ref()).into_owned();
        if !SUPPORTED.contains(&name.as_str()) {
            return Err(Error::Reverse(format!(
                "`w:documentProtection` attribute `{name}` is not represented by documentProtection DSL; remove it or extend the DSL before reversing"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::Reverse(format!(
                    "cannot decode `w:documentProtection/@{name}`: {error}"
                ))
            })?
            .into_owned();
        let value = match name.as_str() {
            "formatting" | "enforcement" => Value::Bool(match value.as_str() {
                "1" | "true" | "on" => true,
                "0" | "false" | "off" => false,
                _ => {
                    return Err(Error::Reverse(format!(
                        "`w:documentProtection/@{name}` has invalid on/off value `{value}`; use true, false, 1, 0, on, or off"
                    )));
                }
            }),
            "cryptAlgorithmSid" | "cryptSpinCount" | "spinCount" => {
                let number = value.parse::<u64>().map_err(|_| {
                    Error::Reverse(format!(
                        "`w:documentProtection/@{name}` must be a non-negative integer; repair the value or remove the attribute"
                    ))
                })?;
                Value::Number(number.into())
            }
            _ => Value::String(value),
        };
        properties.insert(name, value);
    }
    Ok(properties)
}

fn reverse_styles_attribute(
    docx: &docx_rs::Docx,
    metadata: &HashMap<String, StyleMetadata>,
    writer: &Writer,
) -> Result<String> {
    let styles = docx
        .styles
        .styles
        .iter()
        .map(|style| {
            reverse_style_definition(
                style,
                metadata.get(&style.style_id),
                writer.style_theme_colors.get(&style.style_id),
                writer.style_character_indents.get(&style.style_id),
                writer.style_suppress_line_numbers.get(&style.style_id),
                writer.style_paragraph_properties.get(&style.style_id),
                writer.style_run_properties.get(&style.style_id),
            )
        })
        .collect::<Result<Vec<_>>>()?
        .into_iter()
        .flatten()
        .collect::<Vec<_>>();
    if styles.is_empty() {
        return Ok(String::new());
    }
    jsx_prop("styles", &Value::Array(styles))
}

fn parse_core_properties(bytes: &[u8]) -> Result<Map<String, Value>> {
    let Ok(xml) = read_zip_text(bytes, "docProps/core.xml") else {
        return Ok(Map::new());
    };
    let properties = read_core_property_values(&xml)?;
    normalize_core_properties(properties)
}

fn read_core_property_values(xml: &str) -> Result<Map<String, Value>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let property_names = [
        (b"created".as_slice(), "createdAt"),
        (b"modified".as_slice(), "updatedAt"),
        (b"creator".as_slice(), "creator"),
        (b"lastModifiedBy".as_slice(), "lastModifiedBy"),
        (b"description".as_slice(), "description"),
        (b"language".as_slice(), "language"),
        (b"revision".as_slice(), "revision"),
        (b"subject".as_slice(), "subject"),
        (b"title".as_slice(), "title"),
    ];
    let mut properties = Map::new();
    let mut current = None;
    let mut current_value = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => {
                current = property_names
                    .iter()
                    .find(|(local_name, _)| *local_name == element.local_name().as_ref())
                    .map(|(_, property)| *property);
                current_value.clear();
            }
            Ok(Event::Text(text)) => {
                let decoded = text.decode().map_err(|error| {
                    Error::Reverse(format!("cannot decode docProps/core.xml: {error}"))
                })?;
                let value = quick_xml::escape::unescape(&decoded).map_err(|error| {
                    Error::Reverse(format!("cannot unescape docProps/core.xml text: {error}"))
                })?;
                current_value.push_str(&value);
            }
            Ok(Event::GeneralRef(reference)) if current.is_some() => {
                let name = reference.decode().map_err(|error| {
                    Error::Reverse(format!("cannot decode docProps/core.xml entity: {error}"))
                })?;
                if let Some(character) = reference.resolve_char_ref().map_err(|error| {
                    Error::Reverse(format!(
                        "cannot resolve docProps/core.xml character: {error}"
                    ))
                })? {
                    current_value.push(character);
                } else {
                    let character = match name.as_ref() {
                        "amp" => '&',
                        "lt" => '<',
                        "gt" => '>',
                        "apos" => '\'',
                        "quot" => '"',
                        _ => {
                            return Err(Error::Reverse(format!(
                                "docProps/core.xml contains unsupported entity `&{name};`; replace it with a predefined or numeric XML entity"
                            )));
                        }
                    };
                    current_value.push(character);
                }
            }
            Ok(Event::End(_)) => {
                if let Some(property) = current.take() {
                    properties.insert(
                        property.to_owned(),
                        Value::String(std::mem::take(&mut current_value)),
                    );
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse docProps/core.xml: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(properties)
}

fn normalize_core_properties(mut properties: Map<String, Value>) -> Result<Map<String, Value>> {
    for (property, default) in [
        ("createdAt", "1970-01-01T00:00:00Z"),
        ("updatedAt", "1970-01-01T00:00:00Z"),
        ("creator", "unknown"),
        ("lastModifiedBy", "unknown"),
        ("revision", "1"),
    ] {
        if properties.get(property).and_then(Value::as_str) == Some(default) {
            properties.remove(property);
        }
    }
    if let Some(revision) = properties.remove("revision") {
        let text = revision.as_str().unwrap_or_default();
        let parsed = text.parse::<u64>().map_err(|_| {
            Error::Reverse(format!(
                "docProps/core.xml has invalid `cp:revision` value `{text}`; use a positive integer or repair the core properties"
            ))
        })?;
        if parsed == 0 {
            return Err(Error::Reverse(
                "docProps/core.xml has invalid `cp:revision` value `0`; use a positive integer"
                    .to_owned(),
            ));
        }
        properties.insert("revision".to_owned(), Value::Number(parsed.into()));
    }
    for property in [
        "createdAt",
        "updatedAt",
        "creator",
        "lastModifiedBy",
        "description",
        "language",
        "subject",
        "title",
    ] {
        if properties
            .get(property)
            .and_then(Value::as_str)
            .is_some_and(str::is_empty)
        {
            return Err(Error::Reverse(format!(
                "docProps/core.xml contains an empty `{property}` value; provide non-empty core metadata or remove the empty element"
            )));
        }
    }
    Ok(properties)
}

fn collect_custom_xml_items(bytes: &[u8]) -> Result<Option<Value>> {
    let names = zip_entry_names(bytes)?;
    let mut item_parts = names
        .iter()
        .filter(|name| is_custom_xml_item_part(name))
        .cloned()
        .collect::<Vec<_>>();
    let props_parts = names
        .iter()
        .filter(|name| is_custom_xml_props_part(name))
        .cloned()
        .collect::<Vec<_>>();
    if item_parts.is_empty() && props_parts.is_empty() {
        return Ok(None);
    }
    if item_parts.is_empty() {
        return Err(Error::Reverse(format!(
            "custom XML properties `{}` have no item payload; restore the matching `customXml/itemN.xml` or remove the orphan properties part",
            props_parts[0]
        )));
    }
    item_parts.sort();
    let mut items = Vec::new();
    let mut used_props = BTreeSet::new();
    for item_part in item_parts {
        let Some(props_part) = custom_xml_props_part_for(&item_part) else {
            return Err(Error::Reverse(format!(
                "custom XML part `{item_part}` does not use the `itemN.xml` name; rename it or remove the unrepresentable store item"
            )));
        };
        let props_xml = read_zip_text(bytes, &props_part).map_err(|_| {
            Error::Reverse(format!(
                "custom XML item `{item_part}` is missing `{props_part}`; restore the properties part with a `ds:itemID` or remove the orphan item"
            ))
        })?;
        used_props.insert(props_part);
        let (id, schema_refs) = parse_custom_xml_item_id(&props_xml, &item_part)?;
        let xml = normalize_custom_xml_payload(&read_zip_text(bytes, &item_part)?);
        if xml.is_empty() {
            return Err(Error::Reverse(format!(
                "custom XML item `{item_part}` is empty; restore well-formed XML or remove the part"
            )));
        }
        if xml.parse::<docx_rs::CustomItem>().is_err() {
            return Err(Error::Reverse(format!(
                "custom XML item `{item_part}` is not well-formed XML; fix the payload or remove the part"
            )));
        }
        let mut item = Map::new();
        item.insert("id".to_owned(), Value::String(id));
        item.insert("xml".to_owned(), Value::String(xml));
        if !schema_refs.is_empty() {
            item.insert(
                "schemaRefs".to_owned(),
                Value::Array(
                    schema_refs
                        .into_iter()
                        .map(Value::String)
                        .collect::<Vec<_>>(),
                ),
            );
        }
        items.push(Value::Object(item));
    }
    if let Some(orphan) = props_parts.iter().find(|part| !used_props.contains(*part)) {
        return Err(Error::Reverse(format!(
            "custom XML properties `{orphan}` have no item payload; restore the matching `customXml/itemN.xml` or remove the orphan properties part"
        )));
    }
    Ok(Some(Value::Array(items)))
}

fn collect_web_extensions(bytes: &[u8]) -> Result<Option<Value>> {
    const TASKPANES: &str = "word/webextensions/taskpanes.xml";
    const RELS: &str = "word/webextensions/_rels/taskpanes.xml.rels";
    let names = zip_entry_names(bytes)?;
    if !names.iter().any(|name| name == TASKPANES) {
        if names
            .iter()
            .any(|name| name.starts_with("word/webextensions/"))
        {
            return Err(Error::Reverse(format!(
                "web extension package parts exist without `{TASKPANES}`; restore the task-pane catalog or remove the orphan parts"
            )));
        }
        return Ok(None);
    }
    let refs = parse_taskpane_refs(&read_zip_text(bytes, TASKPANES)?)?;
    let rels_xml = read_zip_text(bytes, RELS).map_err(|_| {
        Error::Reverse(format!(
            "task-pane catalog has no `{RELS}` relationship part; restore the web-extension relationships or remove the task-pane catalog"
        ))
    })?;
    let relationships = parse_taskpane_relationships(&rels_xml)?;
    if refs.is_empty() {
        return Err(Error::Reverse(
            "task-pane catalog contains no web extension references; add a `wetp:webextensionref` or remove the empty task-pane part".to_owned(),
        ));
    }
    let mut used_targets = BTreeSet::new();
    let mut extensions = Vec::with_capacity(refs.len());
    for id in &refs {
        let target = relationships.get(id).ok_or_else(|| {
            Error::Reverse(format!(
                "task-pane reference `{id}` has no web-extension relationship; restore its relationship in `{RELS}`"
            ))
        })?;
        if !used_targets.insert(target.clone()) {
            return Err(Error::Reverse(format!(
                "multiple task-pane references target `{target}`; assign each extension a distinct relationship and part"
            )));
        }
        let part = resolve_webextension_target(target)?;
        let xml = read_zip_text(bytes, &part).map_err(|_| {
            Error::Reverse(format!(
                "web-extension relationship `{id}` targets missing `{part}`; restore the part or remove the relationship"
            ))
        })?;
        extensions.push(parse_webextension(&xml, &part)?);
    }
    if relationships.len() != refs.len() {
        return Err(Error::Reverse(
            "task-pane relationships and references do not match; remove orphan web-extension relationships or restore matching references".to_owned(),
        ));
    }
    Ok(Some(Value::Array(extensions)))
}

fn parse_taskpane_refs(xml: &str) -> Result<Vec<String>> {
    parse_xml_elements(
        xml,
        "word/webextensions/taskpanes.xml",
        |reader, element| {
            if element.local_name().as_ref() != b"webextensionref" {
                if element.local_name().as_ref() == b"taskpane" {
                    for (name, expected) in [
                        (b"dockstate".as_slice(), ""),
                        (b"visibility", "1"),
                        (b"width", "350"),
                        (b"row", "1"),
                    ] {
                        if let Some(value) = xml_attribute(reader, element, name)?
                            && value != expected
                        {
                            return Err(Error::Reverse(format!(
                                "task-pane layout `{}` is not representable in JSX; restore the default `{}` value or keep the original DOCX",
                                String::from_utf8_lossy(name),
                                expected
                            )));
                        }
                    }
                }
                return Ok(None);
            }
            xml_attribute(reader, element, b"id")?.map(Some).ok_or_else(|| {
            Error::Reverse(
                "task-pane web-extension reference is missing `r:id`; restore the relationship id or remove the malformed reference".to_owned(),
            )
        })
        },
    )
}

fn parse_taskpane_relationships(xml: &str) -> Result<HashMap<String, String>> {
    let mut relationships = HashMap::new();
    for (id, target) in parse_xml_elements(
        xml,
        "word/webextensions/_rels/taskpanes.xml.rels",
        |reader, element| {
            if element.local_name().as_ref() != b"Relationship"
                || xml_attribute(reader, element, b"Type")?.as_deref()
                    != Some("http://schemas.microsoft.com/office/2011/relationships/webextension")
            {
                return Ok(None);
            }
            let id = xml_attribute(reader, element, b"Id")?.ok_or_else(|| {
            Error::Reverse("web-extension relationship is missing `Id`; restore its relationship identifier".to_owned())
        })?;
            let target = xml_attribute(reader, element, b"Target")?.ok_or_else(|| {
            Error::Reverse(format!("web-extension relationship `{id}` is missing `Target`; restore the target part path"))
        })?;
            Ok(Some((id, target)))
        },
    )? {
        if relationships.insert(id.clone(), target).is_some() {
            return Err(Error::Reverse(format!(
                "duplicate web-extension relationship id `{id}`; assign unique relationship ids"
            )));
        }
    }
    Ok(relationships)
}

fn resolve_webextension_target(target: &str) -> Result<String> {
    let target = target.trim_start_matches('/');
    if target.starts_with("../") || target.contains("/../") || target.is_empty() {
        return Err(Error::Reverse(format!(
            "web-extension relationship target `{target}` is not a safe package-relative path; repair it to point inside `word/webextensions/`"
        )));
    }
    let part = if target.starts_with("word/") {
        target.to_owned()
    } else {
        format!("word/webextensions/{target}")
    };
    if !part.starts_with("word/webextensions/") {
        return Err(Error::Reverse(format!(
            "web-extension target `{target}` is outside `word/webextensions/`; move the part into that directory or remove the relationship"
        )));
    }
    Ok(part)
}

fn parse_webextension(xml: &str, part: &str) -> Result<Value> {
    let mut extension = Map::new();
    let mut properties = Map::new();
    let mut root_seen = false;
    let mut reference_seen = false;
    let mut inside_properties = false;
    let mut unsupported_children = Vec::new();
    let _: Vec<()> = parse_xml_elements(xml, part, |reader, element| {
        let name = element.local_name();
        match name.as_ref() {
            b"webextension" => {
                if root_seen {
                    return Err(Error::Reverse(format!(
                        "`{part}` contains multiple web-extension roots; keep exactly one `we:webextension` element"
                    )));
                }
                root_seen = true;
                let id = xml_attribute(reader, element, b"id")?.ok_or_else(|| Error::Reverse(format!("`{part}` is missing web-extension `id`; restore the `we:webextension/@id` attribute")))?;
                let id = id.trim_matches(['{', '}']).to_owned();
                if id.is_empty() {
                    return Err(Error::Reverse(format!(
                        "`{part}` has an empty web-extension id; restore its `we:webextension/@id` value"
                    )));
                }
                extension.insert("id".to_owned(), Value::String(id));
            }
            b"reference" => {
                if reference_seen {
                    return Err(Error::Reverse(format!(
                        "`{part}` contains multiple web-extension references; keep exactly one `we:reference` element"
                    )));
                }
                reference_seen = true;
                for (xml_name, key) in [
                    (b"id".as_slice(), "referenceId"),
                    (b"version", "version"),
                    (b"store", "store"),
                    (b"storeType", "storeType"),
                ] {
                    let value = xml_attribute(reader, element, xml_name)?.filter(|value| !value.is_empty()).ok_or_else(|| Error::Reverse(format!("`{part}` web-extension reference is missing or has an empty `{key}`; restore the corresponding `we:reference` attribute")))?;
                    extension.insert(key.to_owned(), Value::String(value));
                }
            }
            b"properties" => inside_properties = true,
            b"property" if inside_properties => {
                let name = xml_attribute(reader, element, b"name")?.ok_or_else(|| Error::Reverse(format!("`{part}` has a web-extension property without `name`; add a unique property name")))?;
                let value = xml_attribute(reader, element, b"value")?.ok_or_else(|| Error::Reverse(format!("`{part}` web-extension property `{name}` has no `value`; restore its string value")))?;
                // docx-rs 0.4.22 encodes property strings as JSON string values.
                let value = if value.starts_with('"') && value.ends_with('"') {
                    serde_json::from_str::<String>(&value).map_err(|error| {
                        Error::Reverse(format!(
                            "`{part}` web-extension property `{name}` is not a valid JSON string (`{error}`); repair its value or remove the property"
                        ))
                    })?
                } else {
                    value
                };
                if properties
                    .insert(name.clone(), Value::String(value))
                    .is_some()
                {
                    return Err(Error::Reverse(format!(
                        "`{part}` contains duplicate web-extension property `{name}`; rename or remove the duplicate"
                    )));
                }
            }
            b"alternateReferences" | b"bindings" | b"snapshot" => {}
            _ => unsupported_children.push(String::from_utf8_lossy(name.as_ref()).into_owned()),
        }
        Ok(None)
    })?;
    if !root_seen || !reference_seen {
        return Err(Error::Reverse(format!(
            "`{part}` is missing its web-extension root or reference; restore the required `we:webextension` and `we:reference` elements"
        )));
    }
    if !unsupported_children.is_empty() {
        return Err(Error::Reverse(format!(
            "`{part}` contains unsupported web-extension structures {}; these have no JSX representation, so remove them or retain the original DOCX",
            unsupported_children.join(", ")
        )));
    }
    if !properties.is_empty() {
        extension.insert("properties".to_owned(), Value::Object(properties));
    }
    Ok(Value::Object(extension))
}

fn parse_xml_elements<T>(
    xml: &str,
    part: &str,
    mut visit: impl FnMut(&Reader<&[u8]>, &BytesStart<'_>) -> Result<Option<T>>,
) -> Result<Vec<T>> {
    let mut reader = Reader::from_str(xml);
    let mut buffer = Vec::new();
    let mut values = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                if let Some(value) = visit(&reader, &element)? {
                    values.push(value);
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse {part} at byte {}: {error}",
                    reader.error_position()
                )));
            }
            _ => {}
        }
        buffer.clear();
    }
    Ok(values)
}

fn zip_entry_names(bytes: &[u8]) -> Result<Vec<String>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    Ok((0..archive.len())
        .filter_map(|index| {
            archive
                .by_index(index)
                .ok()
                .map(|file| file.name().replace('\\', "/"))
        })
        .collect())
}

fn is_custom_xml_item_part(name: &str) -> bool {
    Path::new(name)
        .file_name()
        .and_then(|value| value.to_str())
        .is_some_and(|file_name| {
            file_name.starts_with("item")
                && !file_name.starts_with("itemProps")
                && Path::new(file_name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        })
        && name.replace('\\', "/").starts_with("customXml/")
        && !name.contains("_rels/")
}

fn is_custom_xml_props_part(name: &str) -> bool {
    Path::new(name)
        .file_name()
        .and_then(|value| value.to_str())
        .is_some_and(|file_name| {
            file_name.starts_with("itemProps")
                && Path::new(file_name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        })
        && name.replace('\\', "/").starts_with("customXml/")
}

fn custom_xml_props_part_for(item_part: &str) -> Option<String> {
    let file_name = Path::new(item_part).file_name()?.to_str()?;
    let number = file_name
        .strip_prefix("item")?
        .strip_suffix(".xml")
        .or_else(|| file_name.strip_prefix("item")?.strip_suffix(".XML"))?;
    Some(format!("customXml/itemProps{number}.xml"))
}

fn parse_custom_xml_item_id(props_xml: &str, item_part: &str) -> Result<(String, Vec<String>)> {
    let mut reader = Reader::from_str(props_xml);
    reader.config_mut().trim_text(true);
    let mut item_id = None;
    let mut schema_refs = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                match element.local_name().as_ref() {
                    b"datastoreItem" => {
                        item_id = xml_attribute(&reader, &element, b"itemID")?;
                    }
                    b"schemaRef" => {
                        if let Some(uri) = xml_attribute(&reader, &element, b"uri")?
                            .filter(|value| !value.is_empty())
                        {
                            schema_refs.push(uri);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse custom XML properties for `{item_part}`: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    let id = item_id
        .as_deref()
        .map(|value| value.trim_matches(|ch| ch == '{' || ch == '}'))
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            Error::Reverse(format!(
                "custom XML item `{item_part}` is missing `ds:itemID`; restore the item id on `itemProps` or remove the store item"
            ))
        })?;
    Ok((id.to_owned(), schema_refs))
}

fn normalize_custom_xml_payload(xml: &str) -> String {
    let trimmed = xml.trim();
    let Some(rest) = trimmed.strip_prefix("<?xml") else {
        return trimmed.to_owned();
    };
    rest.find("?>").map_or_else(
        || trimmed.to_owned(),
        |index| rest[index + 2..].trim().to_owned(),
    )
}

fn scalar_i64(value: &Value) -> Option<i64> {
    value
        .as_i64()
        .or_else(|| value.as_u64().and_then(|inner| i64::try_from(inner).ok()))
        .or_else(|| {
            value.get("val").and_then(|inner| {
                inner
                    .as_i64()
                    .or_else(|| inner.as_u64().and_then(|number| i64::try_from(number).ok()))
            })
        })
}

fn parse_on_off(value: &str) -> Result<bool> {
    match value {
        "1" | "true" | "on" => Ok(true),
        "0" | "false" | "off" => Ok(false),
        _ => Err(Error::Reverse(format!(
            "invalid on/off value `{value}` for w:suppressLineNumbers; use true, false, 1, 0, on, or off"
        ))),
    }
}

fn reverse_style_definition(
    style: &docx_rs::Style,
    metadata: Option<&StyleMetadata>,
    theme: Option<&ThemeColorInfo>,
    raw_character_indentation: Option<&BTreeMap<String, i64>>,
    raw_suppress_line_numbers: Option<&bool>,
    raw_paragraph_properties: Option<&Map<String, Value>>,
    raw_run_properties: Option<&Map<String, Value>>,
) -> Result<Option<Value>> {
    let source = json(style)?;
    let mut definition = Map::new();
    definition.insert("id".to_owned(), Value::String(style.style_id.clone()));
    definition.insert("name".to_owned(), source["name"].clone());
    definition.insert(
        "type".to_owned(),
        Value::String(style.style_type.to_string()),
    );
    copy_optional_string(
        metadata.and_then(|value| value.based_on.as_deref()),
        &source,
        &mut definition,
        "basedOn",
    );
    copy_optional_string(
        metadata.and_then(|value| value.next.as_deref()),
        &source,
        &mut definition,
        "next",
    );
    copy_optional_string(
        metadata.and_then(|value| value.link.as_deref()),
        &source,
        &mut definition,
        "link",
    );
    definition.insert(
        "quickFormat".to_owned(),
        Value::Bool(
            metadata
                .and_then(|value| value.quick_format)
                .unwrap_or(style.q_format),
        ),
    );
    if let Some(value) = metadata
        .and_then(|value| value.ui_priority)
        .or(style.ui_priority)
    {
        definition.insert("uiPriority".to_owned(), Value::Number(value.into()));
    }
    if metadata
        .and_then(|value| value.semi_hidden)
        .unwrap_or(style.semi_hidden)
    {
        definition.insert("semiHidden".to_owned(), Value::Bool(true));
    }
    if metadata
        .and_then(|value| value.unhide_when_used)
        .unwrap_or(style.unhide_when_used)
    {
        definition.insert("unhideWhenUsed".to_owned(), Value::Bool(true));
    }
    let mut run_properties = reverse_run_properties(&source["runProperty"])?.unwrap_or_default();
    if let Some(raw_run_properties) = raw_run_properties {
        run_properties.extend(raw_run_properties.clone());
    }
    if !run_properties.is_empty() {
        let mut run = run_properties;
        apply_theme_color_to_map(&mut run, theme);
        definition.insert("run".to_owned(), Value::Object(run));
    } else if let Some(theme) = theme.filter(|theme| {
        theme.theme_color.is_some() || theme.theme_shade.is_some() || theme.theme_tint.is_some()
    }) {
        let mut run = Map::new();
        apply_theme_color_to_map(&mut run, Some(theme));
        if !run.is_empty() {
            definition.insert("run".to_owned(), Value::Object(run));
        }
    }
    let mut paragraph =
        reverse_paragraph_properties(&source["paragraphProperty"])?.unwrap_or_default();
    if let Some(raw_character_indentation) = raw_character_indentation {
        for (name, value) in raw_character_indentation {
            paragraph.insert(name.clone(), Value::Number((*value).into()));
        }
    }
    if let Some(value) = raw_suppress_line_numbers {
        paragraph.insert("suppressLineNumbers".to_owned(), Value::Bool(*value));
    }
    if let Some(raw_properties) = raw_paragraph_properties {
        paragraph.extend(raw_properties.clone());
    }
    if !paragraph.is_empty() {
        definition.insert("paragraph".to_owned(), Value::Object(paragraph));
    }
    if style.style_type == docx_rs::StyleType::Table {
        if let Some(table) = reverse_table_style_properties(&source["tableProperty"])? {
            definition.insert("table".to_owned(), Value::Object(table));
        }
        if let Some(cell) = reverse_table_cell_properties(&source["tableCellProperty"])? {
            definition.insert("cell".to_owned(), Value::Object(cell));
        }
    }
    if is_stock_normal(&definition) {
        return Ok(None);
    }
    Ok(Some(Value::Object(definition)))
}

fn reverse_table_style_properties(source: &Value) -> Result<Option<Map<String, Value>>> {
    let mut output = Map::new();
    if let Some(width) = source.get("width").filter(|value| !value.is_null()) {
        let width_value = width
            .get("width")
            .and_then(Value::as_u64)
            .unwrap_or_default();
        let width_type = width
            .get("widthType")
            .and_then(Value::as_str)
            .unwrap_or_default();
        if width_value != 0 || width_type != "auto" {
            return Err(Error::Reverse(
                "table style contains a width property that this JSX style schema cannot represent; remove `w:tblW` from the style or apply width to a `Table` instance".to_owned(),
            ));
        }
    }
    if let Some(value) = source
        .get("justification")
        .and_then(scalar_string)
        .filter(|value| *value != "left")
    {
        output.insert("align".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source.get("indent").filter(|value| !value.is_null()) {
        let width_type = value
            .get("widthType")
            .and_then(Value::as_str)
            .unwrap_or_default();
        if width_type != "dxa" {
            return Err(Error::Reverse(format!(
                "table style indent uses `{width_type}`, which JSX cannot represent; convert it to dxa or remove the style indent"
            )));
        }
        let twips = value.get("width").and_then(Value::as_i64).ok_or_else(|| {
            Error::Reverse("table style indent is malformed; repair `w:tblInd`".to_owned())
        })?;
        output.insert("indent".to_owned(), points_value(twips)?);
    }
    if source.get("style").is_some_and(|value| !value.is_null())
        || source.get("layout").is_some_and(|value| !value.is_null())
        || source.get("position").is_some_and(|value| !value.is_null())
    {
        return Err(Error::Reverse(
            "table style contains nested style, layout, or positioning properties not representable by JSX; remove them from `w:tblPr` or use the corresponding property on a `Table` instance".to_owned(),
        ));
    }
    if let Some(value) = source.get("margins").filter(|value| !value.is_null()) {
        output.insert("margins".to_owned(), table_cell_margins_jsx(value)?);
    }
    if let Some(value) = source.get("borders").filter(|value| !value.is_null()) {
        output.insert("border".to_owned(), table_cell_border_jsx(value)?);
    }
    Ok((!output.is_empty()).then_some(output))
}

fn reverse_table_cell_properties(source: &Value) -> Result<Option<Map<String, Value>>> {
    let output = table_cell_jsx_props(source)?;
    Ok((!output.is_empty()).then_some(output))
}

fn is_stock_normal(definition: &Map<String, Value>) -> bool {
    definition.get("id").and_then(Value::as_str) == Some("Normal")
        && definition.get("name").and_then(Value::as_str) == Some("Normal")
        && definition.get("type").and_then(Value::as_str) == Some("paragraph")
        && definition.get("quickFormat").and_then(Value::as_bool) != Some(false)
        && ![
            "basedOn",
            "next",
            "link",
            "uiPriority",
            "run",
            "paragraph",
            "table",
            "cell",
        ]
        .iter()
        .any(|key| definition.contains_key(*key))
        && definition.get("semiHidden").and_then(Value::as_bool) != Some(true)
        && definition.get("unhideWhenUsed").and_then(Value::as_bool) != Some(true)
}

fn copy_optional_string(
    preferred: Option<&str>,
    source: &Value,
    target: &mut Map<String, Value>,
    key: &str,
) {
    if let Some(value) = preferred.or_else(|| source.get(key).and_then(scalar_string)) {
        target.insert(key.to_owned(), Value::String(value.to_owned()));
    }
}

fn read_style_metadata(bytes: &[u8]) -> Result<HashMap<String, StyleMetadata>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    let Ok(mut file) = archive.by_name("word/styles.xml") else {
        return Ok(HashMap::new());
    };
    let mut xml = Vec::new();
    file.read_to_end(&mut xml)
        .map_err(|error| Error::Reverse(format!("cannot read word/styles.xml: {error}")))?;
    parse_style_metadata(&xml)
}

fn parse_style_metadata(xml: &[u8]) -> Result<HashMap<String, StyleMetadata>> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut output = HashMap::new();
    let mut current: Option<(String, StyleMetadata)> = None;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"style" => {
                if let Some(id) = xml_attribute(&reader, &element, b"styleId")? {
                    current = Some((id, StyleMetadata::default()));
                }
            }
            Ok(Event::Empty(element)) => {
                apply_style_metadata_element(&reader, &element, &mut current)?;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"style" => {
                if let Some((id, metadata)) = current.take() {
                    output.insert(id, metadata);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(Error::Reverse(format!(
                    "cannot parse word/styles.xml at byte {}: {error}",
                    reader.error_position()
                )));
            }
        }
        buffer.clear();
    }
    Ok(output)
}

fn apply_style_metadata_element(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    current: &mut Option<(String, StyleMetadata)>,
) -> Result<()> {
    let Some((_, metadata)) = current.as_mut() else {
        return Ok(());
    };
    let name = element.local_name();
    let name = name.as_ref();
    match name {
        b"basedOn" => metadata.based_on = xml_attribute(reader, element, b"val")?,
        b"next" => metadata.next = xml_attribute(reader, element, b"val")?,
        b"link" => metadata.link = xml_attribute(reader, element, b"val")?,
        b"qFormat" => metadata.quick_format = Some(xml_on_off(reader, element)?),
        b"uiPriority" => {
            metadata.ui_priority =
                xml_attribute(reader, element, b"val")?.and_then(|value| value.parse().ok());
        }
        b"semiHidden" => metadata.semi_hidden = Some(xml_on_off(reader, element)?),
        b"unhideWhenUsed" => metadata.unhide_when_used = Some(xml_on_off(reader, element)?),
        _ => {}
    }
    Ok(())
}

fn xml_on_off(reader: &Reader<&[u8]>, element: &BytesStart<'_>) -> Result<bool> {
    Ok(xml_attribute(reader, element, b"val")?
        .is_none_or(|value| !matches!(value.as_str(), "0" | "false" | "off")))
}

fn validate_run_visibility_xml(bytes: &[u8]) -> Result<()> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    for index in 0..archive.len() {
        let mut file = archive
            .by_index(index)
            .map_err(|error| Error::Reverse(format!("cannot read DOCX part: {error}")))?;
        let part = file.name().to_owned();
        if !part.starts_with("word/")
            || !Path::new(&part)
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        {
            continue;
        }
        let mut xml = String::new();
        file.read_to_string(&mut xml)
            .map_err(|error| Error::Reverse(format!("cannot read `{part}`: {error}")))?;
        let mut reader = Reader::from_str(&xml);
        loop {
            match reader.read_event() {
                Ok(Event::Start(element) | Event::Empty(element))
                    if matches!(element.local_name().as_ref(), b"vanish" | b"specVanish") =>
                {
                    let name = if element.local_name().as_ref() == b"vanish" {
                        "w:vanish"
                    } else {
                        "w:specVanish"
                    };
                    if let Some(value) = xml_attribute(&reader, &element, b"val")? {
                        match value.as_str() {
                            "0" | "false" | "off" => {
                                return Err(Error::Reverse(format!(
                                    "`{part}` contains `{name} w:val=\"{value}\"`, but docx-rs 0.4.22's reader treats this explicit off value as on; preserving it as JSX would change the document. Keep the original DOCX or remove the explicit off override before reversing"
                                )));
                            }
                            "1" | "true" | "on" => {}
                            _ => {
                                return Err(Error::Reverse(format!(
                                    "`{part}` contains unsupported `{name} w:val=\"{value}\"`; use true, false, 1, 0, on, or off"
                                )));
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(error) => {
                    return Err(Error::Reverse(format!(
                        "invalid `{part}` while checking hidden-text overrides: {error}"
                    )));
                }
            }
        }
    }
    Ok(())
}

fn validate_external_xml_features(bytes: &[u8]) -> Result<()> {
    validate_run_visibility_xml(bytes)?;
    validate_character_indent_xml(bytes)
}

fn xml_attribute(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| Error::Reverse(format!("invalid XML attribute: {error}")))?;
        if attribute.key.local_name().as_ref() == name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::Reverse(format!("invalid XML attribute value: {error}")));
        }
    }
    Ok(None)
}

fn reverse_run_properties(source: &Value) -> Result<Option<Map<String, Value>>> {
    let mut output = Map::new();
    for (source_key, target_key) in [
        ("bold", "bold"),
        ("italic", "italic"),
        ("strike", "strike"),
        ("dstrike", "doubleStrike"),
        ("caps", "caps"),
        ("vanish", "hidden"),
    ] {
        if source.get(source_key).is_some() {
            output.insert(
                target_key.to_owned(),
                Value::Bool(property_enabled(source, source_key)),
            );
        }
    }
    if let Some(value) = source
        .get("color")
        .and_then(scalar_string)
        .filter(|value| value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        output.insert("color".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source.get("sz").and_then(scalar_u64) {
        let value = i64::try_from(value)
            .map_err(|_| Error::Reverse("style font size exceeds supported range".to_owned()))?;
        output.insert("size".to_owned(), scaled_decimal_value(value, 2)?);
    }
    if let Some((name, value)) = reverse_fonts_property(source, "font", "fonts") {
        output.insert(name, value);
    }
    if let Some(value) = source.get("underline").and_then(scalar_string) {
        output.insert("underline".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source.get("vertAlign").and_then(scalar_string) {
        if !["baseline", "superscript", "subscript"].contains(&value) {
            return Err(Error::Reverse(format!(
                "style run has unsupported vertical alignment `{value}`; use baseline, superscript, or subscript"
            )));
        }
        output.insert("verticalAlign".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source.get("characterSpacing").and_then(scalar_i64) {
        output.insert("characterSpacing".to_owned(), points_value(value)?);
    }
    Ok((!output.is_empty()).then_some(output))
}

fn append_paragraph_mark_revision_attributes(
    attributes: &mut Vec<String>,
    paragraph_property: &Value,
) -> Result<()> {
    let run_property = paragraph_property
        .get("runProperty")
        .unwrap_or(&Value::Null);
    append_run_revision_attributes(attributes, run_property, "paragraph mark")
}

fn append_run_revision_attributes(
    attributes: &mut Vec<String>,
    run_property: &Value,
    subject: &str,
) -> Result<()> {
    let inserted = run_property.get("ins").filter(|value| !value.is_null());
    let deleted = run_property.get("del").filter(|value| !value.is_null());
    if inserted.is_some() && deleted.is_some() {
        return Err(Error::Reverse(format!(
            "{subject} has both insertion and deletion revision markers, which cannot be represented together; keep only one `w:ins` or `w:del` revision"
        )));
    }
    for (source_key, target_key, revision) in
        [("ins", "inserted", inserted), ("del", "deleted", deleted)]
    {
        let Some(revision) = revision else {
            continue;
        };
        let revision = revision.as_object().ok_or_else(|| {
            Error::Reverse(format!(
                "{subject} `{source_key}` revision metadata is malformed; repair its author/date or remove the revision"
            ))
        })?;
        if revision
            .get("children")
            .and_then(Value::as_array)
            .is_some_and(|children| !children.is_empty())
        {
            return Err(Error::Reverse(format!(
                "{subject} `{source_key}` contains nested revision content not modeled by JSX; preserve the original DOCX or remove the nested revision content"
            )));
        }
        let mut metadata = Map::new();
        for key in ["author", "date"] {
            if let Some(value) = revision.get(key) {
                let value = value.as_str().filter(|value| !value.is_empty()).ok_or_else(|| {
                    Error::Reverse(format!(
                        "{subject} `{source_key}.{key}` must be a non-empty string; repair the revision metadata or remove the revision"
                    ))
                })?;
                metadata.insert(key.to_owned(), Value::String(value.to_owned()));
            }
        }
        attributes.push(jsx_prop(target_key, &Value::Object(metadata))?);
    }
    Ok(())
}

fn apply_theme_color_to_map(target: &mut Map<String, Value>, theme: Option<&ThemeColorInfo>) {
    let Some(theme) = theme else {
        return;
    };
    if let Some(value) = theme
        .val
        .as_deref()
        .filter(|value| value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        target
            .entry("color".to_owned())
            .or_insert_with(|| Value::String(value.to_owned()));
    }
    if let Some(value) = &theme.theme_color {
        target.insert("themeColor".to_owned(), Value::String(value.clone()));
    }
    if let Some(value) = &theme.theme_shade {
        target.insert("themeShade".to_owned(), Value::String(value.clone()));
    }
    if let Some(value) = &theme.theme_tint {
        target.insert("themeTint".to_owned(), Value::String(value.clone()));
    }
}

fn theme_color_attrs(theme: &ThemeColorInfo) -> Vec<String> {
    let mut attrs = Vec::new();
    if let Some(value) = &theme.theme_color {
        attrs.push(attr("themeColor", value));
    }
    if let Some(value) = &theme.theme_shade {
        attrs.push(attr("themeShade", value));
    }
    if let Some(value) = &theme.theme_tint {
        attrs.push(attr("themeTint", value));
    }
    attrs
}

fn reverse_fonts_property(source: &Value, singular: &str, plural: &str) -> Option<(String, Value)> {
    let fonts = source.get("fonts").and_then(Value::as_object)?;
    let mut output = Map::new();
    for key in [
        "ascii",
        "hiAnsi",
        "eastAsia",
        "cs",
        "asciiTheme",
        "hiAnsiTheme",
        "eastAsiaTheme",
        "csTheme",
        "hint",
    ] {
        if let Some(value) = fonts.get(key).and_then(Value::as_str) {
            output.insert(key.to_owned(), Value::String(value.to_owned()));
        }
    }
    if output.is_empty() {
        return None;
    }
    let physical =
        ["ascii", "hiAnsi", "eastAsia", "cs"].map(|key| output.get(key).and_then(Value::as_str));
    let can_collapse = physical[0].is_some()
        && physical.iter().all(|value| *value == physical[0])
        && !output
            .keys()
            .any(|key| key.ends_with("Theme") || key == "hint");
    if can_collapse {
        return Some((
            singular.to_owned(),
            Value::String(physical[0].unwrap_or_default().to_owned()),
        ));
    }
    Some((plural.to_owned(), Value::Object(output)))
}

fn jsx_prop(name: &str, value: &Value) -> Result<String> {
    match value {
        Value::String(value) => Ok(attr(name, value)),
        value => serde_json::to_string(value)
            .map(|value| format!(" {name}={{{value}}}"))
            .map_err(|error| Error::Reverse(error.to_string())),
    }
}

fn reverse_paragraph_properties(source: &Value) -> Result<Option<Map<String, Value>>> {
    let mut output = Map::new();
    if let Some(value) = source.get("alignment").and_then(scalar_string) {
        output.insert("align".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(spacing) = source.get("lineSpacing") {
        for (source_key, target_key) in [
            ("before", "spacingBefore"),
            ("after", "spacingAfter"),
            ("line", "lineSpacing"),
        ] {
            if let Some(value) = spacing.get(source_key).and_then(Value::as_i64) {
                output.insert(target_key.to_owned(), points_value(value)?);
            }
        }
        for (source_key, target_key) in [
            ("beforeLines", "spacingBeforeLines"),
            ("afterLines", "spacingAfterLines"),
        ] {
            if let Some(value) = spacing.get(source_key).and_then(Value::as_u64) {
                output.insert(target_key.to_owned(), Value::Number(value.into()));
            }
        }
        copy_string(spacing, &mut output, "lineRule", "lineRule");
    }
    if let Some(indent) = source.get("indent") {
        for (source_key, target_key) in [("start", "indentLeft"), ("end", "indentRight")] {
            if let Some(value) = indent.get(source_key).and_then(Value::as_i64) {
                output.insert(target_key.to_owned(), points_value(value)?);
            }
        }
        for key in ["hangingChars", "firstLineChars"] {
            if let Some(value) = indent.get(key).and_then(Value::as_i64) {
                output.insert(key.to_owned(), Value::Number(value.into()));
            }
        }
        if let Some(special) = indent.get("specialIndent")
            && let (Some(kind), Some(value)) = (
                special.get("type").and_then(Value::as_str),
                special.get("val").and_then(Value::as_i64),
            )
        {
            let key = if kind == "hanging" {
                "hanging"
            } else {
                "firstLine"
            };
            output.insert(key.to_owned(), points_value(value)?);
        }
    }
    reverse_paragraph_boolean_properties(source, &mut output);
    reverse_paragraph_right_indent(source, &mut output)?;
    reverse_paragraph_shading(source, &mut output)?;
    if let Some(value) = source
        .get("outlineLvl")
        .and_then(Value::as_u64)
        .or_else(|| source.get("outlineLvl").and_then(scalar_u64))
    {
        output.insert("outlineLevel".to_owned(), Value::Number(value.into()));
    }
    Ok((!output.is_empty()).then_some(output))
}

fn reverse_paragraph_boolean_properties(source: &Value, output: &mut Map<String, Value>) {
    if source.get("keepNext").and_then(Value::as_bool) == Some(true) {
        output.insert("keepNext".to_owned(), Value::Bool(true));
    }
    if source.get("keepLines").and_then(Value::as_bool) == Some(true) {
        output.insert("keepLines".to_owned(), Value::Bool(true));
    }
    for key in ["pageBreakBefore", "widowControl", "bidi"] {
        if let Some(value) = source.get(key).and_then(Value::as_bool) {
            output.insert(key.to_owned(), Value::Bool(value));
        }
    }
}

fn reverse_paragraph_right_indent(source: &Value, output: &mut Map<String, Value>) -> Result<()> {
    if let Some(value) = source.get("adjustRightInd") {
        let adjustment = scalar_i64(value).ok_or_else(|| {
            Error::Reverse(
                "paragraph style `w:adjustRightInd` is malformed; repair its integer `w:val` or keep the original DOCX".to_owned(),
            )
        })?;
        output.insert(
            "adjustRightIndent".to_owned(),
            Value::Number(adjustment.into()),
        );
    }
    Ok(())
}

fn reverse_paragraph_shading(source: &Value, output: &mut Map<String, Value>) -> Result<()> {
    if let Some(value) = source.get("shading").filter(|value| !value.is_null()) {
        let fill = value.get("fill").and_then(Value::as_str).ok_or_else(|| {
            Error::Reverse(
                "paragraph style shading has no representable fill color; set a solid `w:shd/@w:fill` or keep the original DOCX".to_owned(),
            )
        })?;
        if fill.len() != 6 || !fill.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(Error::Reverse(
                "paragraph style shading fill is not six-digit RGB; repair `w:shd/@w:fill` or keep the original DOCX".to_owned(),
            ));
        }
        if value.get("shdType").and_then(Value::as_str) != Some("clear")
            || value.get("color").and_then(Value::as_str) != Some("auto")
        {
            return Err(Error::Reverse(
                "paragraph style uses patterned shading not represented by JSX; use a solid fill or keep the original DOCX".to_owned(),
            ));
        }
        output.insert(
            "shading".to_owned(),
            Value::String(fill.to_ascii_uppercase()),
        );
    }
    Ok(())
}

fn copy_string(source: &Value, target: &mut Map<String, Value>, from: &str, to: &str) {
    if let Some(value) = source.get(from).and_then(scalar_string) {
        target.insert(to.to_owned(), Value::String(value.to_owned()));
    }
}

fn points_value(twips: i64) -> Result<Value> {
    scaled_decimal_value(twips, 20)
}

fn scaled_decimal_value(value: i64, divisor: i64) -> Result<Value> {
    let whole = value / divisor;
    let remainder = value % divisor;
    let number = if remainder == 0 {
        whole.to_string()
    } else if divisor == 2 {
        format!("{whole}.5")
    } else {
        let sign = if value < 0 { "-" } else { "" };
        let absolute = value.unsigned_abs();
        format!("{sign}{}.{:02}", absolute / 20, absolute % 20 * 5)
    };
    number
        .parse::<Number>()
        .map(Value::Number)
        .map_err(|error| {
            Error::Reverse(format!(
                "cannot represent numeric value `{number}`: {error}"
            ))
        })
}

fn nested<'a>(value: &'a Value, path: &[&str]) -> Option<&'a Value> {
    path.iter().try_fold(value, |value, key| value.get(key))
}

fn nested_string<'a>(value: &'a Value, path: &[&str]) -> Option<&'a str> {
    nested(value, path).and_then(Value::as_str)
}

fn scalar_string(value: &Value) -> Option<&str> {
    value
        .as_str()
        .or_else(|| value.get("val").and_then(Value::as_str))
}

fn scalar_u64(value: &Value) -> Option<u64> {
    value
        .as_u64()
        .or_else(|| value.get("val").and_then(Value::as_u64))
}

fn property_enabled(value: &Value, key: &str) -> bool {
    value.get(key).is_some_and(|value| {
        value
            .as_bool()
            .unwrap_or_else(|| value.get("val").and_then(Value::as_bool).unwrap_or(true))
    })
}

fn attr(name: &str, value: &str) -> String {
    format!(" {name}=\"{}\"", escape_attr(value))
}

fn point_attr(name: &str, twips: i64) -> String {
    let sign = if twips < 0 { "-" } else { "" };
    let absolute = twips.unsigned_abs();
    let whole = absolute / 20;
    let remainder = absolute % 20;
    if remainder == 0 {
        format!(" {name}={{{sign}{whole}}}")
    } else {
        let hundredths = remainder * 5;
        format!(" {name}={{{sign}{whole}.{hundredths:02}}}")
    }
}

fn escape_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn jsx_string(value: &str) -> Result<String> {
    serde_json::to_string(value)
        .map(|value| format!("{{{value}}}"))
        .map_err(|error| Error::Reverse(error.to_string()))
}

fn unsupported(context: &str, value: &impl std::fmt::Debug) -> Error {
    Error::Reverse(format!("unsupported {context} structure: {value:?}"))
}

fn text_box_reverse_error(context: &str) -> Error {
    Error::Reverse(format!(
        "{context} text box cannot be represented as JSX without losing position and geometry; replace it with ordinary paragraphs and images, or use a docx-rs version with a text-box writer before round-tripping"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use docx_rs::{
        AbstractNumbering, BorderType, CellMargins, CharacterSpacingValues, Comment, Delete,
        DocGrid, Docx, Footer, Header, HeightRule, Hyperlink, HyperlinkType, IndentLevel, Insert,
        InstrToC, Level, LevelJc, LevelOverride, LevelSuffixType, LevelText, LineSpacing,
        LineSpacingType, NumberFormat, Numbering, NumberingId, PageMargin, PageNumType,
        PageOrientationType, PageSize, Paragraph, ParagraphProperty, Pic, Run, RunFonts, Section,
        Settings, Shading, ShdType, SpecialIndentType, Start, StructuredDataTag, Style, StyleType,
        Table, TableAlignmentType, TableBorder, TableBorderPosition, TableBorders, TableCell,
        TableCellBorder, TableCellBorderPosition, TableCellBorders, TableCellMargins,
        TableCellProperty, TableLayoutType, TableOfContents, TableOfContentsItem,
        TablePositionProperty, TableRow, TextDirectionType, ThemeColor, VAlignType, VMergeType,
        WidthType,
    };
    use std::io::Write;

    fn external_docx(paragraph: Paragraph) -> Vec<u8> {
        let mut bytes = Cursor::new(Vec::new());
        Docx::new()
            .add_paragraph(paragraph)
            .build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");
        bytes.into_inner()
    }

    #[test]
    fn reverse_external_docx_should_preserve_distinct_default_font_slots() {
        let mut bytes = Cursor::new(Vec::new());
        Docx::new()
            .default_fonts(
                RunFonts::new()
                    .ascii("Times New Roman")
                    .hi_ansi("Arial")
                    .east_asia("宋体")
                    .cs("Traditional Arabic")
                    .ascii_theme("majorHAnsi")
                    .hi_ansi_theme("minorHAnsi")
                    .east_asia_theme("majorEastAsia")
                    .cs_theme("minorBidi")
                    .hint("eastAsia"),
            )
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body")))
            .build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");

        let jsx = reverse_document(&bytes.into_inner()).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#"defaultFonts={{"ascii":"Times New Roman","hiAnsi":"Arial","eastAsia":"宋体","cs":"Traditional Arabic","asciiTheme":"majorHAnsi","hiAnsiTheme":"minorHAnsi","eastAsiaTheme":"majorEastAsia","csTheme":"minorBidi","hint":"eastAsia"}}"#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_collapse_equal_default_font_slots() {
        let mut bytes = Cursor::new(Vec::new());
        Docx::new()
            .default_fonts(
                RunFonts::new()
                    .ascii("Arial")
                    .hi_ansi("Arial")
                    .east_asia("Arial")
                    .cs("Arial"),
            )
            .build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");

        let jsx = reverse_document(&bytes.into_inner()).expect("external DOCX should reverse");

        assert!(jsx.contains(r#"defaultFont="Arial""#), "{jsx}");
        assert!(!jsx.contains("defaultFonts="), "{jsx}");
    }

    #[test]
    fn reverse_external_docx_should_preserve_exact_line_spacing_and_zero_paragraph_spacing() {
        let bytes = external_docx(
            Paragraph::new().line_spacing(
                LineSpacing::new()
                    .before(0)
                    .after(0)
                    .line(560)
                    .line_rule(LineSpacingType::Exact),
            ),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains(
                r#"<Paragraph spacingBefore={0} spacingAfter={0} lineSpacing={28} lineRule="exact">"#
            ),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_first_line_indent_in_twentieth_points() {
        let bytes = external_docx(Paragraph::new().indent(
            Some(0),
            Some(SpecialIndentType::FirstLine(643)),
            Some(0),
            None,
        ));

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r"<Paragraph indentLeft={0} indentRight={0} firstLine={32.15}>"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_explicit_zero_first_line_indent() {
        let bytes = external_docx(Paragraph::new().indent(
            None,
            Some(SpecialIndentType::FirstLine(0)),
            None,
            None,
        ));

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(jsx.contains(" firstLine={0}"), "{jsx}");
    }

    #[test]
    fn reverse_external_docx_should_reject_conflicting_character_unit_indentation() {
        let packed =
            external_docx(Paragraph::new().add_run(Run::new().add_text("character indent")));
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        let document = document.replace(
            "<w:pPr><w:rPr /></w:pPr>",
            r#"<w:pPr><w:ind w:firstLineChars="200" w:hangingChars="100" /></w:pPr>"#,
        );
        assert!(document.contains("w:firstLineChars"), "{document}");
        let modified =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let error = reverse_document(&modified)
            .expect_err("conflicting character-unit indentation must be rejected");
        assert!(
            error.to_string().contains("firstLineChars")
                && error.to_string().contains("hangingChars")
                && error.to_string().contains("retain only one"),
            "diagnostic should identify the conflicting values and fix: {error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_explain_malformed_character_unit_value() {
        let packed = external_docx(Paragraph::new().add_run(Run::new().add_text("malformed")));
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        let document = document.replace(
            "<w:pPr><w:rPr /></w:pPr>",
            r#"<w:pPr><w:ind w:firstLineChars="oops" /><w:rPr /></w:pPr>"#,
        );
        let modified =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let error = reverse_document(&modified).expect_err("malformed indentation must fail");
        assert!(
            error.to_string().contains("non-integer")
                && error
                    .to_string()
                    .contains("repair the character-unit indentation"),
            "diagnostic should explain the malformed value and fix: {error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_paragraph_character_indentation() {
        let packed = pack_docx(
            Docx::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("first")))
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("second"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        document = document.replacen(
            "<w:pPr><w:rPr /></w:pPr>",
            r#"<w:pPr><w:ind w:left="120" w:leftChars="-100" w:rightChars="200" w:firstLineChars="640" /><w:rPr /></w:pPr>"#,
            1,
        );
        document = document.replacen(
            "<w:pPr><w:rPr /></w:pPr>",
            r#"<w:pPr><w:ind w:startChars="120" w:endChars="80" w:hangingChars="20" /><w:rPr /></w:pPr>"#,
            1,
        );
        let modified =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&modified).expect("character indentation should reverse");
        for attr in [
            "indentLeftChars={-100}",
            "indentLeft={6}",
            "indentRightChars={200}",
            "firstLineChars={640}",
            "indentStartChars={120}",
            "indentEndChars={80}",
            "hangingChars={20}",
        ] {
            assert!(jsx.contains(attr), "missing {attr}: {jsx}");
        }
    }

    #[test]
    fn reverse_external_docx_should_preserve_style_character_indentation() {
        let packed = pack_docx(
            Docx::new()
                .add_style(
                    Style::new("CharacterIndent", StyleType::Paragraph).name("Character Indent"),
                )
                .add_style(
                    Style::new("StartEndIndent", StyleType::Paragraph).name("Start End Indent"),
                )
                .add_style(
                    Style::new("FirstLineIndent", StyleType::Paragraph).name("First Line Indent"),
                ),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part")
            .read_to_string(&mut styles)
            .expect("styles XML");
        for (style_id, indent) in [
            ("CharacterIndent", r#"<w:ind w:hangingChars="20" />"#),
            (
                "StartEndIndent",
                r#"<w:ind w:startChars="120" w:endChars="80" />"#,
            ),
            (
                "FirstLineIndent",
                r#"<w:ind w:leftChars="-100" w:rightChars="200" w:firstLineChars="640" />"#,
            ),
        ] {
            let style_marker = styles
                .find(&format!(r#"w:styleId="{style_id}""#))
                .expect("custom style should be present");
            let style_end = styles[style_marker..]
                .find("</w:style>")
                .map(|offset| style_marker + offset)
                .expect("style end should be present");
            styles.insert_str(style_end, &format!("<w:pPr>{indent}</w:pPr>"));
        }
        let modified = rewrite_package_parts(&packed, &[("word/styles.xml", styles.as_bytes())]);
        let jsx = reverse_document(&modified).expect("style character indentation should reverse");
        for property in [
            r#""hangingChars":20"#,
            r#""indentStartChars":120"#,
            r#""indentEndChars":80"#,
            r#""indentLeftChars":-100"#,
            r#""indentRightChars":200"#,
            r#""firstLineChars":640"#,
        ] {
            assert!(jsx.contains(property), "missing {property}: {jsx}");
        }
    }

    #[test]
    fn reverse_external_docx_should_preserve_paragraph_and_run_style_references() {
        let bytes = external_docx(
            Paragraph::new()
                .style("Body")
                .add_run(Run::new().style("Emphasis").add_text("styled")),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#"<Paragraph style="Body">"#)
                && jsx.contains(r#"<Run style="Emphasis">"#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_font_slots_on_runs_paragraphs_and_styles() {
        let mut bytes = Cursor::new(Vec::new());
        Docx::new()
            .add_style(
                Style::new("Body", StyleType::Paragraph)
                    .name("Body")
                    .fonts(RunFonts::new().ascii("Style Latin").east_asia("样式中文")),
            )
            .add_paragraph(
                Paragraph::new()
                    .fonts(
                        RunFonts::new()
                            .ascii("Paragraph Latin")
                            .east_asia("段落中文"),
                    )
                    .add_run(
                        Run::new()
                            .fonts(
                                RunFonts::new()
                                    .ascii("Run Latin")
                                    .hi_ansi("Run ANSI")
                                    .east_asia("运行中文")
                                    .cs("Run CS")
                                    .hint("eastAsia"),
                            )
                            .add_text("mixed"),
                    ),
            )
            .build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");

        let jsx = reverse_document(&bytes.into_inner()).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#"run":{"fonts":{"ascii":"Style Latin","eastAsia":"样式中文"}}"#),
            "{jsx}"
        );
        assert!(
            jsx.contains(
                r#"<Paragraph fonts={{"ascii":"Paragraph Latin","eastAsia":"段落中文"}}>"#
            ),
            "{jsx}"
        );
        assert!(jsx.contains(r#"<Run fonts={{"ascii":"Run Latin","hiAnsi":"Run ANSI","eastAsia":"运行中文","cs":"Run CS","hint":"eastAsia"}}>"#), "{jsx}");
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_style_reference() {
        let mut bytes = Cursor::new(Vec::new());
        Docx::new()
            .add_table(
                Table::new(vec![TableRow::new(vec![TableCell::new().add_paragraph(
                    Paragraph::new().add_run(Run::new().add_text("cell")),
                )])])
                .style("Grid"),
            )
            .build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");

        let jsx = reverse_document(&bytes.into_inner()).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#"<Table style="Grid" border={{"top":{"style":"single""#),
            "table style reference and explicit table borders must reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_cell_properties() {
        let mut cell = TableCell::new()
            .width(1600, WidthType::Dxa)
            .grid_span(2)
            .vertical_merge(VMergeType::Restart)
            .vertical_align(VAlignType::Center)
            .text_direction(TextDirectionType::TbRl)
            .shading(Shading::new().fill("D9EAF7"))
            .set_borders(
                TableCellBorders::with_empty().set(
                    TableCellBorder::new(TableCellBorderPosition::Top)
                        .border_type(BorderType::Double)
                        .size(8)
                        .color("112233"),
                ),
            );
        cell.property = cell.property.margins(
            CellMargins::new()
                .margin_top(20, WidthType::Dxa)
                .margin_right(40, WidthType::Dxa)
                .margin_bottom(60, WidthType::Dxa)
                .margin_left(80, WidthType::Dxa),
        );
        let bytes = pack_docx(Docx::new().add_table(Table::new(vec![TableRow::new(vec![cell])])));

        let jsx = reverse_document(&bytes).expect("external table cell should reverse");

        assert!(
            jsx.contains(r#"<TableCell width={80} colSpan={2} verticalAlign="center" verticalMerge="restart" textDirection="tbRl" margins={{"#)
                && jsx.contains(r#"shading="D9EAF7""#)
                && jsx.contains(r#"border={{"top":{"style":"double","size":1,"color":"112233"}}}"#),
            "cell properties must survive external reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_properties() {
        let table = Table::new(vec![])
            .width(2400, WidthType::Dxa)
            .align(TableAlignmentType::Center)
            .layout(TableLayoutType::Fixed)
            .indent(120)
            .set_grid(vec![1200, 1200])
            .margins(
                TableCellMargins::new()
                    .margin_top(20, WidthType::Dxa)
                    .margin_right(40, WidthType::Dxa)
                    .margin_bottom(60, WidthType::Dxa)
                    .margin_left(80, WidthType::Dxa),
            )
            .set_borders(
                TableBorders::with_empty().set(
                    TableBorder::new(TableBorderPosition::Top)
                        .border_type(BorderType::Double)
                        .size(8)
                        .color("112233"),
                ),
            )
            .position(
                TablePositionProperty::new()
                    .left_from_text(40)
                    .right_from_text(60)
                    .vertical_anchor("text")
                    .horizontal_anchor("margin")
                    .position_x_alignment("center")
                    .position_y(80),
            );
        let bytes = pack_docx(Docx::new().add_table(table));

        let jsx = reverse_document(&bytes).expect("external table should reverse");

        assert!(
            jsx.contains(r#"<Table width={120} align="center" layout="fixed" columnWidths={[60,60]}"#)
                && jsx.contains(r#"position={{"leftFromText":2,"rightFromText":3,"y":4,"verticalAnchor":"text","horizontalAnchor":"margin","xAlign":"center"}}"#)
                && jsx.contains(r#"margins={{"top":1,"right":2,"bottom":3,"left":4}}"#)
                && jsx.contains(r#""top":{"style":"double","size":1,"color":"112233"}"#),
            "table properties must survive external reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_percent_table_width_and_autofit_layout() {
        let table = Table::new(vec![])
            .width(2500, WidthType::Pct)
            .layout(TableLayoutType::Autofit);
        let bytes = pack_docx(Docx::new().add_table(table));

        let jsx = reverse_document(&bytes).expect("external table should reverse");

        assert!(
            jsx.contains(r#"<Table widthPercent={50} layout="auto""#),
            "percentage width and autofit layout must survive reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_unsupported_table_cell_width() {
        let cell = TableCell::new().width(120, WidthType::Auto);
        let bytes = pack_docx(Docx::new().add_table(Table::new(vec![TableRow::new(vec![cell])])));

        let error = reverse_document(&bytes)
            .expect_err("percent cell width is not represented by JSX point width");

        assert!(
            error.to_string().contains("auto")
                && error.to_string().contains("convert")
                && error.to_string().contains("retain the DOCX"),
            "diagnostic should explain cell width conversion: {error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_percent_table_cell_width() {
        let cell = TableCell::new().width(2500, WidthType::Pct);
        let bytes = pack_docx(Docx::new().add_table(Table::new(vec![TableRow::new(vec![cell])])));

        let jsx = reverse_document(&bytes).expect("percent cell width should reverse");

        assert!(jsx.contains("<TableCell widthPercent={50}>"), "{jsx}");
    }

    #[test]
    fn reverse_external_docx_should_reject_partial_table_cell_margins() {
        let mut cell = TableCell::new();
        cell.property = cell.property.margin_top(20, WidthType::Dxa);
        let bytes = pack_docx(Docx::new().add_table(Table::new(vec![TableRow::new(vec![cell])])));

        let error = reverse_document(&bytes)
            .expect_err("partial margins cannot be represented by the four-edge JSX object");

        assert!(
            error.to_string().contains("only some edges")
                && error.to_string().contains("add explicit"),
            "diagnostic should explain how to repair partial margins: {error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_patterned_table_cell_shading() {
        let cell = TableCell::new().shading(
            Shading::new()
                .shd_type(ShdType::Pct10)
                .color("112233")
                .fill("D9EAF7"),
        );
        let bytes = pack_docx(Docx::new().add_table(Table::new(vec![TableRow::new(vec![cell])])));

        let error = reverse_document(&bytes)
            .expect_err("patterned shading cannot be represented by solid JSX fill");

        assert!(
            error.to_string().contains("pattern or foreground color")
                && error.to_string().contains("solid fill"),
            "diagnostic should explain the unsupported shading and repair: {error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_style_table_and_cell_properties() {
        let style = Style::new("ReportTable", StyleType::Table)
            .name("Report Table")
            .table_align(TableAlignmentType::Center)
            .table_indent(120)
            .table_cell_property(
                TableCellProperty::new()
                    .vertical_align(VAlignType::Center)
                    .shading(Shading::new().fill("D9EAF7")),
            );
        let mut bytes = Cursor::new(Vec::new());
        Docx::new()
            .add_style(style)
            .build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");

        let jsx = reverse_document(&bytes.into_inner()).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#""align":"center""#)
                && jsx.contains(r#""indent":6"#)
                && jsx.contains(r#""verticalAlign":"center""#)
                && jsx.contains(r#""shading":"D9EAF7""#)
                && !jsx.contains(r#""id":"Normal""#),
            "table style properties must survive external reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_style_definitions_and_inheritance() {
        let mut bytes = Cursor::new(Vec::new());
        Docx::new()
            .add_style(
                Style::new("BodyBase", StyleType::Paragraph)
                    .name("Body Base")
                    .size(24)
                    .line_spacing(
                        LineSpacing::new()
                            .line(560)
                            .line_rule(LineSpacingType::Exact),
                    ),
            )
            .add_style(
                Style::new("Body", StyleType::Paragraph)
                    .name("Body")
                    .based_on("BodyBase")
                    .next("Body")
                    .align(docx_rs::AlignmentType::Both),
            )
            .add_paragraph(
                Paragraph::new()
                    .style("Body")
                    .add_run(Run::new().add_text("text")),
            )
            .build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");

        let jsx = reverse_document(&bytes.into_inner()).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#"{"id":"BodyBase","name":"Body Base","type":"paragraph","quickFormat":true,"run":{"size":12},"paragraph":{"lineSpacing":28,"lineRule":"exact"}}"#)
                && jsx.contains(r#"{"id":"Body","name":"Body","type":"paragraph","basedOn":"BodyBase","next":"Body","quickFormat":true,"paragraph":{"align":"both"}}"#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_extended_paragraph_style_properties() {
        let mut style = Style::new("Body", StyleType::Paragraph).name("Body");
        style.paragraph_property = ParagraphProperty::new()
            .page_break_before(true)
            .widow_control(false)
            .bidi(true)
            .shading(Shading::new().fill("DDEEFF"))
            .adjust_right_ind(1);
        let bytes = pack_docx(
            Docx::new()
                .add_style(style)
                .add_paragraph(Paragraph::new().style("Body")),
        );

        let jsx = reverse_document(&bytes).expect("external style should reverse");
        assert!(
            jsx.contains(r#""adjustRightIndent":1"#)
                && jsx.contains(r#""bidi":true"#)
                && jsx.contains(r#""pageBreakBefore":true"#)
                && jsx.contains(r#""shading":"DDEEFF""#)
                && jsx.contains(r#""widowControl":false"#),
            "paragraph style properties must survive external DOCX reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_extended_style_run_properties() {
        let style = Style::new("Emphasis", StyleType::Character).name("Emphasis");
        let mut style = style;
        style.run_property = docx_rs::RunProperty::new()
            .disable_bold()
            .disable_italic()
            .strike()
            .disable_dstrike()
            .caps()
            .vert_align(docx_rs::VertAlignType::SuperScript)
            .spacing(-30);
        let bytes = pack_docx(Docx::new().add_style(style));

        let jsx = reverse_document(&bytes).expect("external style should reverse");
        for expected in [
            r#""bold":false"#,
            r#""italic":false"#,
            r#""strike":true"#,
            r#""doubleStrike":false"#,
            r#""caps":true"#,
            r#""verticalAlign":"superscript""#,
            r#""characterSpacing":-1.5"#,
        ] {
            assert!(jsx.contains(expected), "missing {expected} in {jsx}");
        }
    }

    #[test]
    fn reverse_external_docx_should_preserve_extended_style_run_effects() {
        let packed = pack_docx(
            Docx::new().add_style(Style::new("Effects", StyleType::Character).name("Effects")),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part")
            .read_to_string(&mut styles)
            .expect("styles XML");
        let style_marker = styles
            .find(r#"w:styleId="Effects""#)
            .expect("custom style should be present");
        let style_end = styles[style_marker..]
            .find("</w:style>")
            .map(|offset| style_marker + offset)
            .expect("style end should be present");
        styles.insert_str(
            style_end,
            r#"<w:pPr><w:rPr><w:w w:val="200"/></w:rPr></w:pPr><w:rPr><w:rPrChange><w:rPr><w:w w:val="250"/></w:rPr></w:rPrChange><w:fitText w:val="840" w:id="9"/><w:w w:val="125"/><w:specVanish/><w:shd w:val="pct20" w:color="336699" w:fill="FFF2CC"/></w:rPr>"#,
        );
        let modified = rewrite_package_parts(&packed, &[("word/styles.xml", styles.as_bytes())]);
        let jsx = reverse_document(&modified).expect("external style effects should reverse");
        for expected in [
            r#""fitText":{"width":42,"id":9}"#,
            r#""stretch":125"#,
            r#""specVanish":true"#,
            r#""shading":{"fill":"FFF2CC","color":"336699","pattern":"pct20"}"#,
        ] {
            assert!(jsx.contains(expected), "missing {expected} in {jsx}");
        }
    }

    #[test]
    fn reverse_should_explain_unrepresentable_style_run_effects() {
        for (properties, expected) in [
            (
                r#"<w:w w:val="700"/>"#,
                "stretch percentage from 1 through 600",
            ),
            (
                r#"<w:specVanish w:val="false"/>"#,
                "explicit off value as on",
            ),
            (
                r#"<w:shd w:val="waves" w:fill="FFF2CC"/>"#,
                "unsupported w:shd pattern",
            ),
        ] {
            let packed = pack_docx(
                Docx::new().add_style(Style::new("Effects", StyleType::Character).name("Effects")),
            );
            let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
            let mut styles = String::new();
            archive
                .by_name("word/styles.xml")
                .expect("styles part")
                .read_to_string(&mut styles)
                .expect("styles XML");
            let style_marker = styles
                .find(r#"w:styleId="Effects""#)
                .expect("custom style should be present");
            let style_end = styles[style_marker..]
                .find("</w:style>")
                .map(|offset| style_marker + offset)
                .expect("style end should be present");
            styles.insert_str(style_end, &format!("<w:rPr>{properties}</w:rPr>"));
            let modified =
                rewrite_package_parts(&packed, &[("word/styles.xml", styles.as_bytes())]);
            let error = reverse_document(&modified)
                .expect_err("unrepresentable external style must be rejected");
            assert!(error.to_string().contains(expected), "{error}");
            if !properties.contains("specVanish") {
                assert!(error.to_string().contains("style `Effects`"), "{error}");
            }
        }
    }

    #[test]
    fn reverse_external_docx_should_preserve_default_footer_content() {
        let mut bytes = Cursor::new(Vec::new());
        Docx::new()
            .footer(
                Footer::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text("footer text"))),
            )
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body")))
            .build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");

        let jsx = reverse_document(&bytes.into_inner()).expect("external DOCX should reverse");

        assert!(
            jsx.contains("<Footer type=\"default\">") && jsx.contains(r#"{"footer text"}"#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_footer_text_box_should_error_instead_of_flattening_geometry() {
        let text_box = docx_rs::TextBox {
            children: vec![docx_rs::TextBoxContentChild::Paragraph(Box::new(
                Paragraph::new().add_run(Run::new().add_text("boxed footer")),
            ))],
            ..docx_rs::TextBox::new()
        };
        let mut run = Run::new();
        run.children.push(RunChild::Drawing(Box::new(
            docx_rs::Drawing::new().text_box(text_box),
        )));
        let paragraph = Paragraph::new().add_run(run);
        let error = Writer::default()
            .footer_paragraph(&paragraph, 0)
            .expect_err("text box geometry cannot be represented by current JSX");

        assert!(
            error.to_string().contains("footer text box")
                && error.to_string().contains("position and geometry"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_footer_text_box_should_report_unrepresentable_geometry() {
        let packed = pack_docx(
            Docx::new()
                .footer(Footer::new().add_paragraph(Paragraph::new()))
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut footer = String::new();
        archive
            .by_name("word/footer1.xml")
            .expect("default footer part")
            .read_to_string(&mut footer)
            .expect("footer XML should be readable");
        let text_box = r#"<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>boxed footer</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;
        footer = footer.replacen("</w:p>", &format!("{text_box}</w:p>"), 1);
        let external = rewrite_package_parts(&packed, &[("word/footer1.xml", footer.as_bytes())]);

        let error = reverse_document(&external)
            .expect_err("text box shape data cannot be represented without loss");

        assert!(
            error.to_string().contains("footer text box")
                && error.to_string().contains("position and geometry"),
            "{error}"
        );
    }

    fn pack_docx(docx: Docx) -> Vec<u8> {
        let mut bytes = Cursor::new(Vec::new());
        docx.build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");
        bytes.into_inner()
    }

    fn test_png(width: u32, height: u32, pixel: [u8; 4]) -> Vec<u8> {
        let mut image = image::DynamicImage::new_rgba8(width, height);
        if let Some(buffer) = image.as_mut_rgba8() {
            for px in buffer.pixels_mut() {
                *px = image::Rgba(pixel);
            }
        }
        let mut encoded = Cursor::new(Vec::new());
        image
            .write_to(&mut encoded, image::ImageFormat::Png)
            .expect("png should encode");
        encoded.into_inner()
    }

    fn header_jsx<'a>(jsx: &'a str, kind: &str) -> &'a str {
        let marker = format!("<Header type=\"{kind}\">");
        jsx.split_once(&marker)
            .and_then(|(_, rest)| rest.split_once("</Header>"))
            .map_or("", |(inner, _)| inner)
    }

    fn footer_jsx<'a>(jsx: &'a str, kind: &str) -> &'a str {
        let marker = format!("<Footer type=\"{kind}\">");
        jsx.split_once(&marker)
            .and_then(|(_, rest)| rest.split_once("</Footer>"))
            .map_or("", |(inner, _)| inner)
    }

    fn jsx_after_header(jsx: &str) -> &str {
        jsx.split_once("</Header>").map_or(jsx, |(_, rest)| rest)
    }

    fn image_src(jsx: &str) -> Option<String> {
        let start = jsx.find(r#"src=""#)? + 5;
        let end = jsx[start..].find('"')? + start;
        Some(jsx[start..end].to_owned())
    }

    fn collide_header_and_body_image_rids(bytes: &[u8]) -> Vec<u8> {
        let mut source =
            zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("DOCX should be a ZIP");
        let names = (0..source.len())
            .map(|index| source.by_index(index).expect("entry").name().to_owned())
            .collect::<Vec<_>>();
        let header_rels_name = names
            .iter()
            .find(|name| name.starts_with("word/_rels/header") && is_word_rels_part(name))
            .cloned()
            .expect("header relationships should exist");
        let header_xml_name = header_rels_name
            .replacen("_rels/", "", 1)
            .trim_end_matches(".rels")
            .to_owned();
        let mut read = |name: &str| {
            let mut file = source.by_name(name).expect("part");
            let mut data = String::new();
            file.read_to_string(&mut data).expect("part should read");
            data
        };
        let document_rels = read("word/_rels/document.xml.rels");
        let header_rels = read(&header_rels_name);
        let header_xml = read(&header_xml_name);
        let document_id = relationships_ending_with(&document_rels, "/image")
            .expect("document rels")
            .into_iter()
            .next()
            .expect("document image relationship")
            .0;
        let header_id = relationships_ending_with(&header_rels, "/image")
            .expect("header rels")
            .into_iter()
            .next()
            .expect("header image relationship")
            .0;
        if document_id == header_id {
            return bytes.to_vec();
        }
        let header_rels = header_rels.replace(
            &format!(r#"Id="{header_id}""#),
            &format!(r#"Id="{document_id}""#),
        );
        let header_xml = header_xml.replace(
            &format!(r#"r:embed="{header_id}""#),
            &format!(r#"r:embed="{document_id}""#),
        );
        drop(source);
        rewrite_package_parts(
            bytes,
            &[
                (header_rels_name.as_str(), header_rels.as_bytes()),
                (header_xml_name.as_str(), header_xml.as_bytes()),
            ],
        )
    }

    fn rewrite_package_parts(bytes: &[u8], replacements: &[(&str, &[u8])]) -> Vec<u8> {
        let mut source =
            zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("DOCX should be a ZIP");
        let output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(output);
        for index in 0..source.len() {
            let mut file = source.by_index(index).expect("entry");
            let name = file.name().to_owned();
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            writer
                .start_file(name.clone(), options)
                .expect("replacement entry");
            if let Some((_, data)) = replacements.iter().find(|(part, _)| *part == name) {
                writer.write_all(data).expect("replacement bytes");
            } else {
                let mut data = Vec::new();
                file.read_to_end(&mut data).expect("entry bytes");
                writer.write_all(&data).expect("copied bytes");
            }
        }
        writer.finish().expect("rewritten ZIP").into_inner()
    }

    fn strip_ir_manifest(bytes: &[u8]) -> Vec<u8> {
        let mut source = zip::ZipArchive::new(Cursor::new(bytes.to_vec()))
            .expect("compiled DOCX should be a ZIP");
        let output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(output);
        for index in 0..source.len() {
            let mut file = source.by_index(index).expect("DOCX entry should open");
            let name = file.name().to_owned();
            if name == "docx-jsx/ir-v1.json" || name.ends_with('/') {
                continue;
            }
            let mut data = Vec::new();
            file.read_to_end(&mut data).expect("DOCX entry should read");
            writer
                .start_file(
                    name,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )
                .expect("DOCX entry should start");
            writer.write_all(&data).expect("DOCX entry should write");
        }
        writer
            .finish()
            .map(Cursor::into_inner)
            .expect("stripped DOCX should finalize")
    }

    #[test]
    fn reverse_external_docx_should_nest_inline_bookmark_around_marked_runs() {
        let bytes = external_docx(
            Paragraph::new()
                .add_bookmark_start(1, "intro")
                .add_run(Run::new().add_text("marked"))
                .add_bookmark_end(1),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains("<InlineBookmark name=\"intro\">")
                && jsx.contains(r#"{"marked"}"#)
                && jsx.contains("</InlineBookmark>")
                && !jsx.contains("<InlineBookmark name=\"intro\" />"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_nest_document_bookmark_around_paragraphs() {
        let bytes = pack_docx(
            Docx::new()
                .add_bookmark_start(1, "chapter")
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body")))
                .add_bookmark_end(1),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains("<Bookmark name=\"chapter\">")
                && jsx.contains(r#"{"body"}"#)
                && jsx.contains("</Bookmark>")
                && !jsx.contains("<Bookmark name=\"chapter\" />"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_nest_comment_around_selected_runs() {
        let bytes = external_docx(
            Paragraph::new()
                .add_comment_start(
                    Comment::new(1)
                        .author("Grace")
                        .date("2026-08-14T00:00:00Z")
                        .add_paragraph(Paragraph::new().add_run(Run::new().add_text("review"))),
                )
                .add_run(Run::new().add_text("commented"))
                .add_comment_end(1),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains(
                r#"<Comment text="review" author="Grace" date="2026-08-14T00:00:00Z" id={1} done={false}>"#
            ) && jsx.contains(r#"{"commented"}"#)
                && jsx.contains("</Comment>"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_nest_comment_inside_deleted_text() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Deleted","props":{"author":"Ada"},"children":[{"type":"Comment","props":{"text":"Typo","author":"Grace"},"children":["removed"]}]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled =
            crate::compiler::compile_document(&ir, Path::new(".")).expect("fixture should compile");
        let bytes = strip_ir_manifest(&compiled);

        let jsx = reverse_document(&bytes).expect("manifest-free DOCX should reverse");
        assert!(
            jsx.contains(r#"<Deleted author="Ada""#)
                && jsx.contains(r#"<Comment text="Typo" author="Grace""#)
                && jsx.contains(r#"{"removed"}"#)
                && jsx.contains("</Comment>")
                && jsx.contains("</Deleted>"),
            "comment range should remain nested inside deleted text: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_deleted_field_instruction() {
        let packed = pack_docx(Docx::new().add_paragraph(Paragraph::new().add_delete(
            Delete::new().add_run(Run::new().add_delete_text("HYPERLINK \"https://example.com\"")),
        )));
        let baseline = docx_rs::read_docx(&packed).expect("baseline DOCX should parse");
        assert!(!baseline.document.children.is_empty(), "{baseline:?}");
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        document = document.replace("w:delText", "w:delInstrText");
        let external =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let (normalized, markers) =
            normalize_deleted_instruction_parts(&external).expect("normalization should work");
        assert_eq!(markers.len(), 1, "{markers:?}");
        let normalized_document =
            read_zip_text(&normalized, "word/document.xml").expect("normalized part should read");
        assert!(
            normalized_document.contains("<w:delText xml:space=\"preserve\">Z0Q</w:delText>"),
            "{normalized_document}"
        );
        let parsed = docx_rs::read_docx(&normalized).expect("normalized DOCX should parse");
        assert!(format!("{parsed:?}").contains("Z0Q"), "{parsed:?}");
        let jsx = reverse_document(&external).expect("deleted field instruction should reverse");
        assert!(
            jsx.contains(
                "<DeletedInstruction instruction=\"HYPERLINK &quot;https://example.com&quot;\" />"
            ),
            "{jsx}\n{document}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_hyperlink_composite_children() {
        let bytes =
            pack_docx(
                Docx::new().add_paragraph(
                    Paragraph::new().add_hyperlink(
                        Hyperlink::new("https://example.com", HyperlinkType::External)
                            .add_structured_data_tag(
                                StructuredDataTag::new()
                                    .alias("LinkData")
                                    .add_run(Run::new().add_text("bound")),
                            )
                            .add_insert(Insert::new(Run::new().add_text("new")).author("Ada"))
                            .add_delete(
                                Delete::new()
                                    .author("Lin")
                                    .add_run(Run::new().add_delete_text("old")),
                            )
                            .add_bookmark_start(1, "insideLink")
                            .add_run(Run::new().add_text("marked"))
                            .add_bookmark_end(1)
                            .add_comment_start(Comment::new(1).author("Grace").add_paragraph(
                                Paragraph::new().add_run(Run::new().add_text("review")),
                            ))
                            .add_run(Run::new().add_text("commented"))
                            .add_comment_end(1),
                    ),
                ),
            );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#"<Hyperlink href="https://example.com""#)
                && jsx.contains(r#"<ContentControl alias="LinkData">"#)
                && jsx.contains(r#"{"bound"}"#)
                && jsx.contains(r#"<Inserted author="Ada""#)
                && jsx.contains(r#"{"new"}"#)
                && jsx.contains(r#"<Deleted author="Lin""#)
                && jsx.contains(r#"{"old"}"#)
                && jsx.contains(r#"<InlineBookmark name="insideLink">"#)
                && jsx.contains(r#"{"marked"}"#)
                && jsx.contains("</InlineBookmark>")
                && jsx.contains(r#"<Comment text="review" author="Grace""#)
                && jsx.contains(r#"{"commented"}"#)
                && jsx.contains("</Comment>")
                && jsx.contains("</Hyperlink>"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_keep_body_hyperlink_sdt_when_headers_also_have_hyperlinks() {
        let bytes = pack_docx(
            Docx::new()
                .header(
                    Header::new().add_paragraph(
                        Paragraph::new().add_hyperlink(
                            Hyperlink::new("headerMark", HyperlinkType::Anchor)
                                .add_run(Run::new().add_text("header link")),
                        ),
                    ),
                )
                .footer(
                    Footer::new().add_paragraph(
                        Paragraph::new().add_hyperlink(
                            Hyperlink::new("footerMark", HyperlinkType::Anchor)
                                .add_run(Run::new().add_text("footer link")),
                        ),
                    ),
                )
                .add_paragraph(
                    Paragraph::new().add_hyperlink(
                        Hyperlink::new("https://body.example", HyperlinkType::External)
                            .add_structured_data_tag(
                                StructuredDataTag::new()
                                    .alias("BodyData")
                                    .add_run(Run::new().add_text("bound")),
                            ),
                    ),
                ),
        );
        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");
        let header = jsx
            .split_once("<Header type=\"default\">")
            .and_then(|(_, rest)| rest.split_once("</Header>"))
            .map(|(content, _)| content)
            .expect("header JSX should exist");
        let footer = jsx
            .split_once("<Footer type=\"default\">")
            .and_then(|(_, rest)| rest.split_once("</Footer>"))
            .map(|(content, _)| content)
            .expect("footer JSX should exist");
        let body = jsx
            .split_once("</Footer>")
            .map(|(_, rest)| rest)
            .expect("body JSX should follow the footer");

        assert!(
            header.contains(r#"<Hyperlink anchor="headerMark""#)
                && header.contains(r#"{"header link"}"#)
                && !header.contains("ContentControl"),
            "header hyperlink stole the body SDT plan: {header}"
        );
        assert!(
            footer.contains(r#"<Hyperlink anchor="footerMark""#)
                && footer.contains(r#"{"footer link"}"#)
                && !footer.contains("ContentControl"),
            "footer hyperlink stole the body SDT plan: {footer}"
        );
        assert!(
            body.contains(r#"<Hyperlink href="https://body.example""#)
                && body.contains(r#"<ContentControl alias="BodyData">"#)
                && body.contains(r#"{"bound"}"#)
                && body.contains("</ContentControl>"),
            "body hyperlink lost its ContentControl: {body}"
        );
    }

    #[test]
    fn reverse_compiled_docx_without_manifest_should_preserve_hyperlink_composite_children() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"href":"https://example.com"},"children":[{"type":"ContentControl","props":{"alias":"LinkData"},"children":["bound"]},{"type":"Inserted","props":{"author":"Ada"},"children":["new"]},{"type":"Deleted","props":{"author":"Lin"},"children":["old"]},{"type":"InlineBookmark","props":{"name":"insideLink"},"children":["marked"]},{"type":"Comment","props":{"text":"review","author":"Grace"},"children":["commented"]}]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, std::path::Path::new("."))
            .expect("compile should work");
        let bytes = strip_ir_manifest(&compiled);

        let jsx = reverse_document(&bytes).expect("manifest-free compiled DOCX should reverse");

        assert!(
            jsx.contains(r#"<Hyperlink href="https://example.com""#)
                && jsx.contains(r#"<ContentControl alias="LinkData">"#)
                && jsx.contains(r#"<Inserted author="Ada""#)
                && jsx.contains(r#"<Deleted author="Lin""#)
                && jsx.contains(r#"<InlineBookmark name="insideLink">"#)
                && jsx.contains(r#"<Comment text="review" author="Grace""#)
                && !jsx.contains("<InlineBookmark name=\"insideLink\" />"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_should_preserve_false_hyperlink_history_without_manifest() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"href":"https://example.com","history":false},"children":["link"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, std::path::Path::new("."))
            .expect("compile should work");
        let bytes = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&bytes).expect("manifest-free DOCX should reverse");
        assert!(jsx.contains("history={false}"), "{jsx}");
    }

    #[test]
    fn reverse_should_reject_non_boolean_hyperlink_history() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"href":"https://example.com"},"children":["link"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, std::path::Path::new("."))
            .expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(compiled.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document.xml")
            .read_to_string(&mut document)
            .expect("document XML");
        let changed = document.replace("w:history=\"1\"", "w:history=\"2\"");
        assert_ne!(
            changed, document,
            "test fixture has no explicit w:history: {document}"
        );
        document = changed;
        let without_manifest = strip_ir_manifest(&compiled);
        let external = rewrite_package_parts(
            &without_manifest,
            &[("word/document.xml", document.as_bytes())],
        );
        let error = reverse_document(&external).expect_err("invalid history value must fail");
        assert!(
            error
                .to_string()
                .contains("use true, false, 1, 0, on, or off"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_two_sections() {
        let bytes = pack_docx(
            Docx::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("first")))
                .add_section(Section::new().page_size(PageSize::new().size(11_906, 16_838)))
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("second")))
                .page_size(15_840, 12_240)
                .page_orient(PageOrientationType::Landscape),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");
        assert_eq!(
            jsx.matches("<Section").count(),
            2,
            "expected two Section components: {jsx}"
        );
        assert!(
            jsx.contains(r#"pageSize="A4""#)
                && jsx.contains(r#"pageSize="Letter""#)
                && jsx.contains(r#"orientation="landscape""#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_row_layout() {
        let row = TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("row"))),
        ])
        .row_height(560.0)
        .height_rule(HeightRule::Exact)
        .cant_split();
        let bytes = pack_docx(Docx::new().add_table(Table::new(vec![row])));
        let jsx = reverse_document(&bytes).expect("external row layout should reverse");
        assert!(
            jsx.contains(r#"<TableRow height={28} heightRule="exact" cantSplit>"#),
            "row height, rule, and cantSplit must survive external reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_paragraph_mark_insertion() {
        let paragraph = Paragraph::new()
            .insert("Ada", "2026-08-14T00:00:00Z")
            .add_run(Run::new().add_text("inserted paragraph"));
        let bytes = pack_docx(Docx::new().add_paragraph(paragraph));

        let jsx =
            reverse_document(&bytes).expect("external paragraph mark insertion should reverse");

        assert!(
            jsx.contains(r#"inserted={{"author":"Ada","date":"2026-08-14T00:00:00Z"}}"#),
            "paragraph mark insertion metadata should survive external reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_paragraph_mark_deletion() {
        let paragraph = Paragraph::new()
            .delete("Lin", "2026-08-15T00:00:00Z")
            .add_run(Run::new().add_text("deleted paragraph"));
        let bytes = pack_docx(Docx::new().add_paragraph(paragraph));

        let jsx =
            reverse_document(&bytes).expect("external paragraph mark deletion should reverse");

        assert!(
            jsx.contains(r#"deleted={{"author":"Lin","date":"2026-08-15T00:00:00Z"}}"#),
            "paragraph mark deletion metadata should survive external reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_schema_invalid_run_property_revisions() {
        let mut run = Run::new().add_text("invalid revision");
        run.run_property = run
            .run_property
            .delete(Delete::new().author("Lin").date("2026-08-15T00:00:00Z"));
        let bytes = external_docx(Paragraph::new().add_run(run));

        let error = reverse_document(&bytes)
            .expect_err("w:del inside w:rPr is invalid and must not be silently dropped");

        assert!(
            error
                .to_string()
                .contains("schema-invalid in this location"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_explicit_off_visibility_flags() {
        for (run_property, element, property_name) in [
            (
                docx_rs::RunProperty::new().vanish(),
                "<w:vanish />",
                "w:vanish",
            ),
            (
                docx_rs::RunProperty::new().spec_vanish(),
                "<w:specVanish />",
                "w:specVanish",
            ),
        ] {
            let mut run = Run::new().add_text("hidden");
            run.run_property = run_property;
            let packed = external_docx(Paragraph::new().add_run(run));
            let mut archive = zip::ZipArchive::new(Cursor::new(&packed)).expect("valid ZIP");
            let mut document = String::new();
            archive
                .by_name("word/document.xml")
                .expect("document part")
                .read_to_string(&mut document)
                .expect("document XML");
            let changed = document.replace(element, &element.replace(" />", " w:val=\"0\" />"));
            assert_ne!(changed, document, "fixture should contain {element}");
            let external =
                rewrite_package_parts(&packed, &[("word/document.xml", changed.as_bytes())]);

            let error = reverse_document(&external)
                .expect_err("docx-rs reader loses explicit off visibility values");

            assert!(
                error
                    .to_string()
                    .contains(&format!("{property_name} w:val=\"0\"")),
                "{error}"
            );
        }
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_row_revisions() {
        let cell =
            || TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("row")));
        let inserted = TableRow::new(vec![cell()]).insert(
            Insert::new_with_empty()
                .author("Ada")
                .date("2026-08-14T00:00:00Z"),
        );
        let deleted = TableRow::new(vec![cell()])
            .delete(Delete::new().author("Lin").date("2026-08-15T00:00:00Z"));
        let bytes = pack_docx(Docx::new().add_table(Table::new(vec![inserted, deleted])));
        let jsx = reverse_document(&bytes).expect("external row revisions should reverse");
        assert!(
            jsx.contains(r#"inserted={{"author":"Ada","date":"2026-08-14T00:00:00Z"}}"#)
                && jsx.contains(r#"deleted={{"author":"Lin","date":"2026-08-15T00:00:00Z"}}"#),
            "row revisions must survive external reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_row_grid_data() {
        let row = TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("row"))),
        ]);
        let packed = pack_docx(Docx::new().add_table(Table::new(vec![row])));
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        let document = document.replace(
            "<w:trPr />",
            r#"<w:trPr><w:gridBefore w:val="2" /><w:gridAfter w:val="1" /><w:wBefore w:w="1200" w:type="dxa" /><w:wAfter w:w="1500" w:type="dxa" /></w:trPr>"#,
        );
        assert!(document.contains("w:gridBefore"));
        let modified =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&modified).expect("external row grid data should reverse");
        assert!(
            jsx.contains("gridBefore={2}")
                && jsx.contains("gridAfter={1}")
                && jsx.contains("widthBefore={60}")
                && jsx.contains("widthAfter={75}"),
            "row grid properties must survive reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_without_manifest_should_preserve_section_columns() {
        let ir: crate::ir::IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"columns":2},"children":[{"type":"Paragraph","props":{},"children":["first"]}]},{"type":"Section","props":{"columns":3},"children":[{"type":"Paragraph","props":{},"children":["second"]}]}]}}"#,
        )
        .expect("fixture IR should parse");
        let compiled =
            crate::compiler::compile_document(&ir, Path::new(".")).expect("fixture should compile");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external DOCX should reverse");
        assert!(
            jsx.contains("columns={2}") && jsx.contains("columns={3}"),
            "section column counts must survive structural reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_without_manifest_should_preserve_section_page_number_format() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"pageNumbering":{"start":4,"chapterStyle":"1","format":"upperRoman","chapterSeparator":"emDash"}},"children":[{"type":"Paragraph","props":{},"children":["chapter page numbers"]}]},{"type":"Section","props":{"pageNumbering":{"format":"lowerLetter"}},"children":[{"type":"Paragraph","props":{},"children":["letter page numbers"]}]}]}}"#,
        )
        .expect("fixture IR should parse");
        let compiled =
            crate::compiler::compile_document(&ir, Path::new(".")).expect("fixture compiles");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external DOCX should reverse");
        assert!(
            jsx.contains(
                r#"pageNumbering={{start:4,chapterStyle:"1",format:"upperRoman",chapterSeparator:"emDash"}}"#
            ) && jsx.contains(r#"pageNumbering={{format:"lowerLetter"}}"#),
            "section page-number format should survive manifest-free reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_read_page_number_format_from_package_xml() {
        let packed = pack_docx(
            Docx::new()
                .page_num_type(PageNumType::new().start(4).chap_style("1"))
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("page numbers"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        let document = document.replace(
            r#"<w:pgNumType w:start="4" w:chapStyle="1" />"#,
            r#"<w:pgNumType w:start="4" w:chapStyle="1" w:fmt="upperRoman" w:chapSep="emDash" />"#,
        );
        assert!(document.contains(r#"w:fmt="upperRoman""#));
        let external =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&external).expect("external DOCX should reverse");
        assert!(
            jsx.contains(
                r#"pageNumbering={{start:4,chapterStyle:"1",format:"upperRoman",chapterSeparator:"emDash"}}"#
            ),
            "all page-numbering attributes should reverse without an IR manifest: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_unknown_page_number_format_with_repair() {
        let packed = pack_docx(
            Docx::new()
                .page_num_type(PageNumType::new().start(1))
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("page numbers"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        let document = document.replace(
            r#"<w:pgNumType w:start="1" />"#,
            r#"<w:pgNumType w:start="1" w:fmt="romanX" />"#,
        );
        let external =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let error = reverse_document(&external)
            .expect_err("unsupported external page-number format must not be dropped");
        assert!(
            error.to_string().contains("romanX") && error.to_string().contains("supported"),
            "error should identify unsupported format and repair: {error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_custom_section_columns_without_manifest() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"columns":{"count":2,"spacing":18,"separator":true}},"children":[{"type":"Paragraph","props":{},"children":["first"]}]},{"type":"Section","props":{"columns":{"widths":[{"width":144,"spacing":24},{"width":180,"spacing":18},{"width":216}],"separator":true}},"children":[{"type":"Paragraph","props":{},"children":["second"]}]}]}}"#,
        ).expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, Path::new("."))
            .expect("document should compile");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external columns should reverse");
        assert!(
            jsx.contains(r#"columns={{"count":2,"spacing":18,"separator":true}}"#),
            "{jsx}"
        );
        assert!(
            jsx.contains(r#"columns={{"widths":[{"width":144,"spacing":24},{"width":180,"spacing":18},{"width":216}],"separator":true}}"#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_should_preserve_custom_section_column_spacing() {
        let packed = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("columns"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        let document = document.replace(
            r#"<w:cols w:space="425" w:num="1" />"#,
            r#"<w:cols w:space="500" w:num="2" />"#,
        );
        assert!(document.contains(r#"w:space="500" w:num="2""#));
        let modified =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&modified).expect("custom column spacing should reverse");
        assert!(
            jsx.contains(r#"columns={{"count":2,"spacing":25}}"#),
            "custom column spacing must survive reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_of_contents_instead_of_content_control() {
        let bytes = pack_docx(
            Docx::new()
                .add_table_of_contents(TableOfContents::new().heading_styles_range(1, 3).auto())
                .add_structured_data_tag(
                    StructuredDataTag::new()
                        .alias("BlockData")
                        .add_paragraph(Paragraph::new().add_run(Run::new().add_text("inside"))),
                ),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part")
            .read_to_string(&mut styles)
            .expect("styles XML");
        // Isolate TOC recognition from the backend's unsupported character-unit
        // indentation in its built-in TOC styles.
        let mut styles_without_character_indent = String::new();
        let mut remaining_styles = styles.as_str();
        while let Some((before, after)) = remaining_styles.split_once(r#" w:leftChars=""#) {
            styles_without_character_indent.push_str(before);
            let Some((_, remainder)) = after.split_once('"') else {
                break;
            };
            remaining_styles = remainder;
        }
        styles_without_character_indent.push_str(remaining_styles);
        let modified = rewrite_package_parts(
            &bytes,
            &[(
                "word/styles.xml",
                styles_without_character_indent.as_bytes(),
            )],
        );

        let jsx = reverse_document(&modified).expect("external DOCX should reverse");

        assert!(
            jsx.contains("<TableOfContents")
                && jsx.contains(r"startLevel={1}")
                && jsx.contains(r"endLevel={3}"),
            "TOC SDT should reverse as TableOfContents: {jsx}"
        );
        assert!(
            jsx.contains(r#"<ContentControl alias="BlockData">"#) && jsx.contains(r#"{"inside"}"#),
            "block control should keep its alias beside the TOC: {jsx}"
        );
        let toc_as_control = jsx.contains("<ContentControl")
            && jsx.matches("<ContentControl").count() == 1
            && !jsx.contains("alias=\"BlockData\"");
        assert!(
            !toc_as_control,
            "TOC must not flatten to a nameless ContentControl: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_restore_deleted_index_field_metadata() {
        let mut figures = TableOfContents::new();
        figures.instr = InstrToC::new().caption_label("Figure");
        let mut entries = TableOfContents::new();
        entries.instr = InstrToC::new().tc_field_identifier(Some("legal".to_owned()));
        let deletion_date = "2026-10-12T00:00:00Z";
        let bytes = pack_docx(
            Docx::new()
                .add_table_of_contents(
                    TableOfContents::new()
                        .heading_styles_range(1, 3)
                        .delete("Editor", deletion_date),
                )
                .add_table_of_contents(figures.delete("Editor", deletion_date))
                .add_table_of_contents(entries.delete("Editor", deletion_date)),
        );
        let external = strip_ir_manifest(&bytes);
        let jsx = reverse_document(&external).expect("deleted index field should reverse");
        assert!(
            jsx.matches(r#"deletion={{"author":"Editor","date":"2026-10-12T00:00:00Z"}}"#)
                .count()
                == 3
                && jsx.contains("<TableOfContents")
                && jsx.contains("<TableOfFigures")
                && jsx.contains("<TableOfEntries"),
            "all deleted index types should retain metadata: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_restore_index_field_paragraph_properties() {
        let toc = TableOfContents::new().paragraph_property(
            ParagraphProperty::new()
                .align(docx_rs::AlignmentType::Center)
                .line_spacing(
                    LineSpacing::new()
                        .before(120)
                        .line(560)
                        .line_rule(docx_rs::LineSpacingType::Exact),
                )
                .indent(
                    Some(80),
                    Some(SpecialIndentType::FirstLine(640)),
                    Some(60),
                    None,
                )
                .keep_next(true),
        );
        let external = strip_ir_manifest(&pack_docx(Docx::new().add_table_of_contents(toc)));
        let jsx = reverse_document(&external).expect("field paragraph properties should reverse");
        assert!(
            jsx.contains(
                r#"fieldParagraph={{"align":"center","indentLeft":4,"indentRight":3,"spacingBefore":6,"lineSpacing":28,"lineRule":"exact","keepNext":true,"firstLine":32}}"#
            ),
            "index paragraph formatting should be restored as fieldParagraph: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_restore_index_field_line_number_suppression() {
        let packed = pack_docx(
            Docx::new().add_table_of_contents(TableOfContents::new().heading_styles_range(1, 3)),
        );
        let mut document =
            read_zip_text(&packed, "word/document.xml").expect("document part should be readable");
        let instruction = document
            .find("<w:instrText>TOC ")
            .expect("TOC instruction should exist");
        let paragraph_start = ["<w:p>", "<w:p "]
            .iter()
            .filter_map(|marker| document[..instruction].rfind(marker))
            .max()
            .expect("field paragraph should start");
        let paragraph_end = document[instruction..]
            .find("</w:p>")
            .map(|end| instruction + end)
            .expect("field paragraph should end");
        let ppr_end = document[paragraph_start..paragraph_end]
            .find("</w:pPr>")
            .map(|end| paragraph_start + end)
            .expect("field paragraph properties should exist");
        document.insert_str(
            ppr_end,
            "<w:ind w:leftChars=\"-100\" w:rightChars=\"200\" w:firstLineChars=\"640\" /><w:suppressLineNumbers />",
        );
        let external = strip_ir_manifest(&rewrite_package_parts(
            &packed,
            &[("word/document.xml", document.as_bytes())],
        ));
        let jsx = reverse_document(&external).expect("external field property should reverse");
        assert!(
            jsx.contains(r#""suppressLineNumbers":true"#)
                && jsx.contains(r#""indentLeftChars":-100"#)
                && jsx.contains(r#""indentRightChars":200"#)
                && jsx.contains(r#""firstLineChars":640"#),
            "index field paragraph should preserve line-number suppression: {jsx}"
        );
        let document = document.replace(
            "<w:suppressLineNumbers />",
            "<w:suppressLineNumbers w:val=\"0\" />",
        );
        let external = strip_ir_manifest(&rewrite_package_parts(
            &packed,
            &[("word/document.xml", document.as_bytes())],
        ));
        let jsx = reverse_document(&external).expect("explicit false should reverse");
        assert!(
            jsx.contains(r#""suppressLineNumbers":false"#),
            "index field paragraph should preserve explicit false: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_conflicting_index_field_indents() {
        let packed = pack_docx(
            Docx::new().add_table_of_contents(TableOfContents::new().heading_styles_range(1, 3)),
        );
        let mut document =
            read_zip_text(&packed, "word/document.xml").expect("document part should be readable");
        let instruction = document
            .find("<w:instrText>TOC ")
            .expect("TOC instruction should exist");
        let paragraph_start = ["<w:p>", "<w:p "]
            .iter()
            .filter_map(|marker| document[..instruction].rfind(marker))
            .max()
            .expect("field paragraph should start");
        let paragraph_end = document[instruction..]
            .find("</w:p>")
            .map(|end| instruction + end)
            .expect("field paragraph should end");
        let ppr_end = document[paragraph_start..paragraph_end]
            .find("</w:pPr>")
            .map(|end| paragraph_start + end)
            .expect("field paragraph properties should exist");
        document.insert_str(
            ppr_end,
            "<w:ind w:leftChars=\"100\" w:startChars=\"120\" />",
        );
        let external = strip_ir_manifest(&rewrite_package_parts(
            &packed,
            &[("word/document.xml", document.as_bytes())],
        ));
        let error = reverse_document(&external)
            .expect_err("conflicting indentation units cannot be represented losslessly");
        assert!(error.to_string().contains("conflicting `indentLeftChars`"));
    }

    #[test]
    fn reverse_external_docx_should_reject_unrepresentable_index_paragraph_properties() {
        let toc = TableOfContents::new().paragraph_property(
            ParagraphProperty::new().run_property(docx_rs::RunProperty::new().caps()),
        );
        let external = strip_ir_manifest(&pack_docx(Docx::new().add_table_of_contents(toc)));
        let error = reverse_document(&external)
            .expect_err("unsupported field paragraph formatting must not be discarded");
        assert!(error.to_string().contains("field paragraph property"));
        assert!(error.to_string().contains("fieldParagraph"));
    }

    #[test]
    fn reverse_external_docx_should_restore_unwrapped_dynamic_toc_without_manifest() {
        let mut figures = TableOfContents::new();
        figures.instr = InstrToC::new().caption_label("Figure");
        let mut entries = TableOfContents::new();
        entries.instr = InstrToC::new().tc_field_identifier(Some("legal".to_owned()));
        let bytes = pack_docx(
            Docx::new()
                .add_table_of_contents(
                    TableOfContents::new()
                        .heading_styles_range(1, 3)
                        .auto()
                        .paragraph_property(
                            ParagraphProperty::new().align(docx_rs::AlignmentType::Center),
                        )
                        .without_sdt(),
                )
                .add_table_of_contents(figures.without_sdt())
                .add_table_of_contents(entries.without_sdt()),
        );
        let external = strip_ir_manifest(&bytes);
        let jsx = reverse_document(&external).expect("unwrapped external TOC should reverse");
        assert!(
            jsx.contains("<TableOfContents")
                && jsx.contains("<TableOfFigures")
                && jsx.contains("<TableOfEntries")
                && jsx.matches("withoutSdt").count() == 3
                && jsx.contains(r#"fieldParagraph={{"align":"center"}}"#),
            "unwrapped field paragraphs should reconstruct all index components: {jsx}"
        );
        assert!(
            !jsx.contains("<Paragraph"),
            "the native field marker paragraphs should not leak as user paragraphs: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_restore_static_toc_items() {
        let packed = pack_docx(
            Docx::new().add_table_of_contents(
                TableOfContents::new().heading_styles_range(1, 3).add_item(
                    TableOfContentsItem::new()
                        .text("Overview")
                        .toc_key("overview")
                        .level(2)
                        .page_ref("7"),
                ),
            ),
        );
        let mut document = read_zip_text(&packed, "word/document.xml")
            .expect("generated document XML should be readable");
        document = document.replace("w:pos=\"80000\"", "w:pos=\"8400\"");
        let bytes = rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&bytes).expect("external static TOC should reverse");
        assert!(
            jsx.contains(
                r#"<TocItem text="Overview" bookmark="overview" level={2} pageRef="7" />"#
            ),
            "static TOC entry must be restored: {jsx}"
        );
    }

    #[test]
    fn reverse_should_restore_toc_before_and_after_paragraphs_without_manifest() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[{"type":"TocBefore","props":{},"children":[{"type":"Paragraph","props":{},"children":["before index"]},{"type":"Table","props":{},"children":[{"type":"TableRow","props":{},"children":[{"type":"TableCell","props":{},"children":[{"type":"Paragraph","props":{},"children":["before table"]}]}]}]}]},{"type":"TocItem","props":{"text":"Entry","bookmark":"entry","pageRef":"3"},"children":[]},{"type":"TocAfter","props":{},"children":[{"type":"Table","props":{},"children":[{"type":"TableRow","props":{},"children":[{"type":"TableCell","props":{},"children":[{"type":"Paragraph","props":{},"children":["after table"]}]}]}]},{"type":"Paragraph","props":{},"children":["after index"]}]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled =
            crate::compiler::compile_document(&ir, Path::new(".")).expect("fixture should compile");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external index should reverse");
        assert!(
            jsx.contains("<TocBefore>") && jsx.contains("before index"),
            "before content should reverse: {jsx}"
        );
        assert!(
            jsx.contains("<TocAfter>") && jsx.contains("after index"),
            "after content should reverse: {jsx}"
        );
        assert!(
            jsx.contains("before table") && jsx.contains("after table"),
            "before/after tables should reverse: {jsx}"
        );
        let after_content = jsx
            .split("<TocAfter>")
            .nth(1)
            .and_then(|content| content.split("</TocAfter>").next())
            .expect("TocAfter wrapper should be present");
        assert!(
            after_content.find("after table") < after_content.find("after index"),
            "after-table order should be preserved: {jsx}"
        );
        assert_eq!(
            after_content.matches("\n        <Paragraph").count(),
            1,
            "docx-rs table-spacing paragraph should not become user content: {jsx}"
        );
        assert!(
            jsx.contains("text=\"Entry\"") && jsx.contains("bookmark=\"entry\""),
            "static index item should remain between wrappers: {jsx}"
        );
    }

    #[test]
    fn reverse_should_preserve_custom_static_toc_tab_position_without_manifest() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[{"type":"TocItem","props":{"text":"Overview","bookmark":"overview","tabPosition":431.25},"children":[]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, Path::new("."))
            .expect("document should compile");
        let jsx = reverse_document(&strip_ir_manifest(&compiled))
            .expect("external document should reverse without its IR manifest");
        assert!(
            jsx.contains("tabPosition={431.25}"),
            "custom tab position should round trip: {jsx}"
        );
    }

    #[test]
    fn reverse_without_manifest_should_preserve_extended_toc_instructions() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{"startLevel":1,"endLevel":3,"tcFieldLevelRange":{"start":2,"end":4},"omitPageNumbersLevelRange":{"start":5,"end":7},"entryBookmarkName":"contents","entryAndPageNumberSeparator":" — ","sequenceAndPageNumbersSeparator":" · ","seqFieldIdentifierForPrefix":"Figure","stylesWithLevels":[{"style":"Report Heading","level":2}],"preserveTab":true,"preserveNewLine":true,"useAppliedParagraphLineLevel":true,"hideTabAndPageNumbersInWebview":true,"hyperlinks":false,"dirty":false,"alias":"Contents"},"children":[]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, std::path::Path::new("."))
            .expect("compile should work");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external TOC should reverse");
        for expected in [
            r#"tcFieldLevelRange={{"start":2,"end":4}}"#,
            r#"omitPageNumbersLevelRange={{"start":5,"end":7}}"#,
            r#"entryBookmarkName="contents""#,
            r#"entryAndPageNumberSeparator=" — ""#,
            r#"sequenceAndPageNumbersSeparator=" · ""#,
            r#"seqFieldIdentifierForPrefix="Figure""#,
            r#"stylesWithLevels={[{"style":"Report Heading","level":2}]}"#,
            "preserveTab",
            "preserveNewLine",
            "useAppliedParagraphLineLevel",
            "hideTabAndPageNumbersInWebview",
            "hyperlinks={false}",
            "dirty={false}",
            r#"alias="Contents""#,
        ] {
            assert!(jsx.contains(expected), "missing `{expected}` in {jsx}");
        }
    }

    #[test]
    fn reverse_without_manifest_should_preserve_figure_and_entry_instructions() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfFigures","props":{"label":"Figure","includeLabelAndNumber":false,"separator":" — ","entryAndPageNumberSeparator":" / ","seqFieldIdentifierForPrefix":"Figure","stylesWithLevels":[{"style":"Caption Custom","level":2}],"preserveNewLine":true,"hyperlinks":false},"children":[]},{"type":"TableOfEntries","props":{"identifier":"legal","tcFieldLevelRange":{"start":1,"end":4},"omitPageNumbersLevelRange":{"start":3,"end":4},"entryBookmarkName":"terms","sequenceAndPageNumbersSeparator":" · ","preserveTab":true},"children":[]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, std::path::Path::new("."))
            .expect("compile should work");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external indexes should reverse");
        for expected in [
            r#"label="Figure""#,
            "includeLabelAndNumber={false}",
            r#"separator=" — ""#,
            r#"entryAndPageNumberSeparator=" / ""#,
            r#"seqFieldIdentifierForPrefix="Figure""#,
            r#"stylesWithLevels={[{"style":"Caption Custom","level":2}]}"#,
            "preserveNewLine",
            r#"identifier="legal""#,
            r#"tcFieldLevelRange={{"start":1,"end":4}}"#,
            r#"omitPageNumbersLevelRange={{"start":3,"end":4}}"#,
            r#"entryBookmarkName="terms""#,
            r#"sequenceAndPageNumbersSeparator=" · ""#,
            "preserveTab",
        ] {
            assert!(jsx.contains(expected), "missing `{expected}` in {jsx}");
        }
    }

    #[test]
    fn reverse_external_docx_should_reject_unknown_toc_switches() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, std::path::Path::new("."))
            .expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(compiled.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document.xml")
            .read_to_string(&mut document)
            .expect("document XML");
        let changed = document.replace(
            r"TOC \o &quot;1-3&quot; \h",
            r"TOC \o &quot;1-3&quot; \h \q",
        );
        assert_ne!(changed, document, "fixture TOC instruction was not found");
        let without_manifest = strip_ir_manifest(&compiled);
        let external = rewrite_package_parts(
            &without_manifest,
            &[("word/document.xml", changed.as_bytes())],
        );
        let error = reverse_document(&external).expect_err("unknown TOC switches must fail");
        assert!(
            error.to_string().contains("unsupported `\\q` switch"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_table_cell_and_body_content_controls() {
        let bytes = pack_docx(
            Docx::new()
                .add_table(Table::new(vec![TableRow::new(vec![
                    TableCell::new().add_structured_data_tag(
                        StructuredDataTag::new()
                            .alias("CellData")
                            .add_run(Run::new().add_text("cell")),
                    ),
                ])]))
                .add_structured_data_tag(
                    StructuredDataTag::new()
                        .alias("BlockData")
                        .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
                ),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#"<ContentControl alias="CellData">"#) && jsx.contains(r#"{"cell"}"#),
            "cell control should reverse: {jsx}"
        );
        assert!(
            jsx.contains(r#"<ContentControl alias="BlockData">"#) && jsx.contains(r#"{"body"}"#),
            "body control alias must not be stolen by the cell SDT: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_header_footer_content_controls() {
        let bytes = pack_docx(
            Docx::new()
                .header(Header {
                    children: vec![HeaderChild::StructuredDataTag(Box::new(
                        StructuredDataTag::new()
                            .alias("HeaderData")
                            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("head"))),
                    ))],
                    ..Header::new()
                })
                .footer(Footer {
                    children: vec![FooterChild::StructuredDataTag(Box::new(
                        StructuredDataTag::new()
                            .alias("FooterData")
                            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("foot"))),
                    ))],
                    ..Footer::new()
                }),
        );
        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");
        assert!(
            jsx.contains(r#"<ContentControl alias="HeaderData">"#)
                && jsx.contains(r#"<ContentControl alias="FooterData">"#)
                && jsx.contains(r#"{"head"}"#)
                && jsx.contains(r#"{"foot"}"#),
            "header/footer block controls should retain their placement and aliases: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_keep_body_cell_sdt_when_header_table_also_has_one() {
        let bytes = pack_docx(
            Docx::new()
                .header(Header::new().add_table(Table::new(vec![TableRow::new(vec![
                    TableCell::new().add_structured_data_tag(
                        StructuredDataTag::new()
                            .alias("HeaderCell")
                            .add_run(Run::new().add_text("header-cell")),
                    ),
                ])])))
                .add_table(Table::new(vec![TableRow::new(vec![
                    TableCell::new().add_structured_data_tag(
                        StructuredDataTag::new()
                            .alias("BodyCell")
                            .add_run(Run::new().add_text("body-cell")),
                    ),
                ])])),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");
        let header = jsx
            .split_once("<Header type=\"default\">")
            .and_then(|(_, rest)| rest.split_once("</Header>"))
            .map(|(content, _)| content)
            .expect("header JSX should exist");
        let body = jsx
            .split_once("</Header>")
            .map(|(_, rest)| rest)
            .expect("body JSX should follow the header");

        assert!(
            header.contains(r#"{"header-cell"}"#) && !header.contains(r#"alias="BodyCell""#),
            "header table cell stole the body cell SDT alias: {header}"
        );
        assert!(
            body.contains(r#"<ContentControl alias="BodyCell">"#)
                && body.contains(r#"{"body-cell"}"#),
            "body table cell lost its ContentControl alias: {body}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_document_level_content_control() {
        let bytes = pack_docx(
            Docx::new()
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("before")))
                .add_structured_data_tag(
                    StructuredDataTag::new()
                        .alias("BlockData")
                        .add_paragraph(Paragraph::new().add_run(Run::new().add_text("inside"))),
                )
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("after"))),
        );

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

        assert!(
            jsx.contains(r#"{"before"}"#)
                && jsx.contains(r#"<ContentControl alias="BlockData">"#)
                && jsx.contains("<Paragraph>")
                && jsx.contains(r#"{"inside"}"#)
                && jsx.contains("</ContentControl>")
                && jsx.contains(r#"{"after"}"#),
            "{jsx}"
        );
        let control = jsx
            .split_once(r#"<ContentControl alias="BlockData">"#)
            .and_then(|(_, rest)| rest.split_once("</ContentControl>"))
            .map(|(content, _)| content)
            .expect("block ContentControl should nest");
        assert!(
            control.contains("<Paragraph>") && control.contains(r#"{"inside"}"#),
            "block control should wrap its paragraph: {control}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_unmatched_bookmark_end() {
        let bytes = external_docx(Paragraph::new().add_bookmark_end(7));

        let error = reverse_document(&bytes).expect_err("orphan bookmark end must fail");

        assert!(
            error.to_string().contains("unmatched bookmark end")
                && error.to_string().contains("pair each"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_section_margins_and_document_grid() {
        let mut docx = Docx::new()
            .page_size(11_907, 16_840)
            .page_margin(PageMargin {
                top: 1440,
                right: 1440,
                bottom: 1440,
                left: 1440,
                header: 709,
                footer: 709,
                gutter: 0,
            })
            .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body")));
        docx.document = docx
            .document
            .doc_grid(DocGrid::with_empty().line_pitch(360));
        let jsx = reverse_document(&pack_docx(docx)).expect("external DOCX should reverse");
        assert!(
            jsx.contains(
                "margins={{top:72,right:72,bottom:72,left:72,header:35.45,footer:35.45,gutter:0}}"
            ) && jsx.contains(r#"documentGrid={{type:"default",linePitch:18}}"#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_section_type_without_manifest() {
        let bytes = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        document = document.replace("<w:sectPr>", r#"<w:sectPr><w:type w:val="continuous" />"#);
        let external = rewrite_package_parts(&bytes, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&external).expect("external DOCX should reverse");
        assert!(
            jsx.contains(r#"<Section type="continuous" pageSize="A4""#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_unknown_section_type_with_repair() {
        let bytes = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        document = document.replace("<w:sectPr>", r#"<w:sectPr><w:type w:val="customBreak" />"#);
        let external = rewrite_package_parts(&bytes, &[("word/document.xml", document.as_bytes())]);
        let error = reverse_document(&external)
            .expect_err("unsupported section type must not be collapsed");
        assert!(
            error.to_string().contains("customBreak") && error.to_string().contains("nextPage"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_line_numbering_without_manifest() {
        let bytes = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        document = document.replace("<w:cols", r#"<w:lnNumType w:countBy="5" w:start="3" w:distance="361" w:restart="newSection" /><w:cols"#);
        let external = rewrite_package_parts(&bytes, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&external).expect("external DOCX should reverse");
        assert!(
            jsx.contains(
                r#"lineNumbering={{countBy:5,start:3,distance:18.05,restart:"newSection"}}"#
            ),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_paragraph_line_number_suppression_without_manifest() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"lineNumbering":{}},"children":[{"type":"Paragraph","props":{"suppressLineNumbers":true},"children":["excluded"]},{"type":"Paragraph","props":{"suppressLineNumbers":false},"children":["included"]},{"type":"List","props":{"type":"ordered"},"children":[{"type":"ListItem","props":{"suppressLineNumbers":true},"children":["list item"]}]}]}]}}"#,
        ).expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, Path::new("."))
            .expect("valid document should compile");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external DOCX should reverse");
        assert!(jsx.contains("suppressLineNumbers={true}"), "{jsx}");
        assert!(jsx.contains("suppressLineNumbers={false}"), "{jsx}");
        assert!(
            jsx.contains("<ListItem suppressLineNumbers={true}>"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_explain_invalid_line_number_suppression_value() {
        let bytes =
            external_docx(Paragraph::new().add_run(Run::new().add_text("malformed suppression")));
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        document = document.replace(
            "<w:pPr><w:rPr /></w:pPr>",
            r#"<w:pPr><w:suppressLineNumbers w:val="maybe" /><w:rPr /></w:pPr>"#,
        );
        let external = rewrite_package_parts(&bytes, &[("word/document.xml", document.as_bytes())]);
        let error = reverse_document(&external).expect_err("invalid on/off must be rejected");
        assert!(
            error.to_string().contains("invalid on/off value")
                && error.to_string().contains("use true, false"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_style_line_number_suppression_without_manifest() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"NoLineNumbers","name":"No Line Numbers","type":"paragraph","paragraph":{"suppressLineNumbers":true}},{"id":"KeepLineNumbers","name":"Keep Line Numbers","type":"paragraph","paragraph":{"suppressLineNumbers":false}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        ).expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, Path::new("."))
            .expect("valid document should compile");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external DOCX should reverse");
        assert!(jsx.contains("\"suppressLineNumbers\":true"), "{jsx}");
        assert!(jsx.contains("\"suppressLineNumbers\":false"), "{jsx}");
    }

    #[test]
    fn reverse_external_docx_should_keep_default_line_numbering_enabled() {
        let bytes = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        document = document.replace("<w:cols", "<w:lnNumType /><w:cols");
        let external = rewrite_package_parts(&bytes, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&external).expect("external DOCX should reverse");
        assert!(jsx.contains("lineNumbering={{}}"), "{jsx}");
    }

    #[test]
    fn reverse_external_docx_should_reject_unknown_line_number_restart() {
        let bytes = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        document = document.replace("<w:cols", r#"<w:lnNumType w:restart="everyLine" /><w:cols"#);
        let external = rewrite_package_parts(&bytes, &[("word/document.xml", document.as_bytes())]);
        let error =
            reverse_document(&external).expect_err("unsupported line-number restart must fail");
        assert!(
            error.to_string().contains("everyLine") && error.to_string().contains("newPage"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_paragraph_keep_and_outline() {
        let bytes = external_docx(
            Paragraph::new()
                .keep_next(true)
                .keep_lines(true)
                .outline_lvl(2)
                .add_run(Run::new().add_text("kept")),
        );
        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");
        assert!(
            jsx.contains("<Paragraph keepNext keepLines outlineLevel={2}>"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_style_keep_and_outline() {
        let bytes = pack_docx(
            Docx::new()
                .add_style({
                    let mut style = Style::new("HeadingLike", StyleType::Paragraph)
                        .name("Heading Like")
                        .outline_lvl(0);
                    style.paragraph_property =
                        style.paragraph_property.keep_next(true).keep_lines(true);
                    style
                })
                .add_paragraph(
                    Paragraph::new()
                        .style("HeadingLike")
                        .add_run(Run::new().add_text("title")),
                ),
        );
        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");
        assert!(
            jsx.contains(r#""id":"HeadingLike""#)
                && jsx.contains(r#""keepNext":true"#)
                && jsx.contains(r#""keepLines":true"#)
                && jsx.contains(r#""outlineLevel":0"#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_run_theme_color() {
        let bytes = external_docx(
            Paragraph::new().add_run(
                Run::new()
                    .color("548DD4")
                    .theme_color(ThemeColor::Text2)
                    .theme_tint("99")
                    .add_text("blue"),
            ),
        );
        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");
        assert!(
            jsx.contains(r#"color="548DD4""#)
                && jsx.contains(r#"themeColor="text2""#)
                && jsx.contains(r#"themeTint="99""#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_omit_stock_normal_style() {
        let bytes = pack_docx(
            Docx::new()
                .add_style(Style::new("Normal", StyleType::Paragraph).name("Normal"))
                .add_style(
                    Style::new("Body", StyleType::Paragraph)
                        .name("Body")
                        .based_on("Normal"),
                )
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");
        assert!(
            jsx.contains(r#""id":"Body""#) && !jsx.contains(r#""id":"Normal""#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_raster_image() {
        let png = test_png(2, 1, [0, 0, 0, 255]);
        let bytes = pack_docx(Docx::new().add_paragraph(
            Paragraph::new().add_run(Run::new().add_image(Pic::new_with_dimensions(png, 2, 1))),
        ));
        let reversed = reverse_package(&bytes).expect("external image DOCX should reverse");
        assert!(
            reversed.jsx.contains("<Image")
                && reversed.jsx.contains(r#"src="media/"#)
                && reversed.jsx.contains(" width={")
                && reversed.jsx.contains(" height={"),
            "{}",
            reversed.jsx
        );
        assert!(
            reversed
                .assets
                .iter()
                .any(|(path, bytes)| path.starts_with("media/") && !bytes.is_empty()),
            "expected extracted raster bytes, got {:?}",
            reversed.assets.keys().collect::<Vec<_>>()
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_image_layout_in_cell() {
        let mut picture = Pic::new_with_dimensions(test_png(2, 1, [0, 0, 0, 255]), 2, 1).floating();
        picture.layout_in_cell = true;
        let bytes = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_image(picture))),
        );

        let jsx = reverse_document(&bytes).expect("external image DOCX should reverse");

        assert!(jsx.contains("layoutInCell={true}"), "{jsx}");
    }

    #[test]
    fn reverse_external_docx_should_preserve_disabled_image_layout_in_cell() {
        let picture = Pic::new_with_dimensions(test_png(2, 1, [0, 0, 0, 255]), 2, 1).floating();
        let bytes = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_image(picture))),
        );

        let jsx = reverse_document(&bytes).expect("external image DOCX should reverse");

        assert!(jsx.contains("layoutInCell={false}"), "{jsx}");
    }

    #[test]
    fn reverse_external_docx_should_preserve_simple_positioned_images() {
        let picture = Pic::new_with_dimensions(test_png(2, 1, [0, 0, 0, 255]), 2, 1).floating();
        let packed = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_image(picture))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        let document = document
            .replacen("simplePos=\"0\"", "simplePos=\"1\"", 1)
            .replacen(
                r#"<wp:simplePos x="0" y="0" />"#,
                r#"<wp:simplePos x="25400" y="-12700" />"#,
                1,
            );
        assert!(document.contains("simplePos=\"1\""));
        let modified =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);

        let jsx = reverse_document(&modified).expect("simple-position coordinates reverse");

        assert!(
            jsx.contains("simplePosition={{ x: 25400, y: -12700 }}"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_header_raster_image() {
        let png = test_png(3, 1, [255, 0, 0, 255]);
        let bytes = pack_docx(
            Docx::new()
                .header(
                    Header::new().add_paragraph(
                        Paragraph::new()
                            .add_run(Run::new().add_image(Pic::new_with_dimensions(png, 3, 1))),
                    ),
                )
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let reversed = reverse_package(&bytes).expect("header image DOCX should reverse");
        let header = header_jsx(&reversed.jsx, "default");
        assert!(
            header.contains("<Image")
                && header.contains(r#"src="media/"#)
                && header.contains(" width={")
                && header.contains(" height={"),
            "header JSX should keep the raster image: {header}"
        );
        assert!(
            reversed
                .assets
                .iter()
                .any(|(path, bytes)| path.starts_with("media/") && !bytes.is_empty()),
            "expected extracted header raster bytes, got {:?}",
            reversed.assets.keys().collect::<Vec<_>>()
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_footer_raster_image() {
        let png = test_png(1, 3, [0, 255, 0, 255]);
        let bytes = pack_docx(
            Docx::new()
                .footer(
                    Footer::new().add_paragraph(
                        Paragraph::new()
                            .add_run(Run::new().add_image(Pic::new_with_dimensions(png, 1, 3))),
                    ),
                )
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let reversed = reverse_package(&bytes).expect("footer image DOCX should reverse");
        let footer = footer_jsx(&reversed.jsx, "default");
        assert!(
            footer.contains("<Image") && footer.contains(r#"src="media/"#),
            "footer JSX should keep the raster image: {footer}"
        );
        assert!(
            reversed
                .assets
                .iter()
                .any(|(path, bytes)| path.starts_with("media/") && !bytes.is_empty()),
            "expected extracted footer raster bytes, got {:?}",
            reversed.assets.keys().collect::<Vec<_>>()
        );
    }

    #[test]
    fn reverse_external_docx_should_keep_header_and_body_images_distinct() {
        let header_png = test_png(2, 2, [255, 0, 0, 255]);
        let body_png = test_png(2, 2, [0, 0, 255, 255]);
        let bytes = pack_docx(
            Docx::new()
                .header(Header::new().add_paragraph(Paragraph::new().add_run(
                    Run::new().add_image(Pic::new_with_dimensions(header_png.clone(), 2, 2)),
                )))
                .add_paragraph(Paragraph::new().add_run(
                    Run::new().add_image(Pic::new_with_dimensions(body_png.clone(), 2, 2)),
                )),
        );
        let reversed =
            reverse_package(&bytes).expect("header and body images should reverse independently");
        let header = header_jsx(&reversed.jsx, "default");
        let body = jsx_after_header(&reversed.jsx);
        assert!(
            header.contains("<Image") && body.contains("<Image"),
            "both header and body should keep Image nodes:\n{}",
            reversed.jsx
        );
        let header_src = image_src(header).expect("header Image src");
        let body_src = image_src(body).expect("body Image src");
        assert_ne!(
            header_src, body_src,
            "header and body images must not share one src when their bytes differ"
        );
        assert_eq!(
            reversed.assets.get(&header_src).map(Vec::as_slice),
            Some(header_png.as_slice()),
            "header asset bytes should match the header PNG"
        );
        assert_eq!(
            reversed.assets.get(&body_src).map(Vec::as_slice),
            Some(body_png.as_slice()),
            "body asset bytes should match the body PNG"
        );
    }

    #[test]
    fn reverse_external_docx_should_resolve_header_image_when_rids_collide() {
        let header_png = test_png(2, 2, [255, 128, 0, 255]);
        let body_png = test_png(2, 2, [0, 128, 255, 255]);
        let packed = pack_docx(
            Docx::new()
                .header(Header::new().add_paragraph(Paragraph::new().add_run(
                    Run::new().add_image(Pic::new_with_dimensions(header_png.clone(), 2, 2)),
                )))
                .add_paragraph(Paragraph::new().add_run(
                    Run::new().add_image(Pic::new_with_dimensions(body_png.clone(), 2, 2)),
                )),
        );
        let bytes = collide_header_and_body_image_rids(&packed);
        let reversed = reverse_package(&bytes).expect(
            "colliding header/body image relationship ids should reverse from their own rels",
        );
        let header = header_jsx(&reversed.jsx, "default");
        let body = jsx_after_header(&reversed.jsx);
        let header_src = image_src(header).expect("header Image src");
        let body_src = image_src(body).expect("body Image src");
        assert_ne!(header_src, body_src);
        assert_eq!(
            reversed.assets.get(&header_src).map(Vec::as_slice),
            Some(header_png.as_slice()),
            "header must resolve from headerN.xml.rels, not the body rId"
        );
        assert_eq!(
            reversed.assets.get(&body_src).map(Vec::as_slice),
            Some(body_png.as_slice()),
            "body must keep its own media bytes"
        );
    }

    #[test]
    fn reverse_external_docx_without_annotation_parts_should_not_invent_them() {
        let packed = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("plain"))),
        );
        let stripped = strip_package_parts(
            &packed,
            &[
                "word/comments.xml",
                "word/commentsExtended.xml",
                "word/footnotes.xml",
                "word/numbering.xml",
            ],
        );
        let jsx = reverse_document(&stripped).expect("external DOCX should reverse");
        assert!(
            !jsx.contains("<Comment") && !jsx.contains("<Footnote") && !jsx.contains("<List"),
            "missing annotation parts must not invent components: {jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_custom_xml_items() {
        let bytes = pack_docx(
            Docx::new()
                .add_custom_item(
                    "06AC5857-5C65-A94A-BCEC-37356A209BC3",
                    "<customer><name>Ada</name></customer>",
                )
                .add_custom_item(
                    "11111111-AAAA-BBBB-CCCC-222222222222",
                    r#"<order id="42"/>"#,
                )
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let jsx = reverse_document(&bytes).expect("custom XML items should reverse");
        assert!(
            jsx.contains("customXmlItems={")
                && jsx.contains(r#""id":"06AC5857-5C65-A94A-BCEC-37356A209BC3""#)
                && jsx.contains(r#""xml":"<customer><name>Ada</name></customer>""#)
                && jsx.contains(r#""id":"11111111-AAAA-BBBB-CCCC-222222222222""#)
                && jsx.contains(r#""xml":"<order id=\"42\""#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_without_manifest_should_preserve_web_extensions() {
        let ir: crate::ir::IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"webExtensions":[{"id":"7f33b723-fb58-4524-8733-dbedc4b7c095","referenceId":"office-addin","version":"1.0.0.0","store":"developer","storeType":"Registry","properties":{"mode":"review"}},{"id":"11111111-2222-3333-4444-555555555555","referenceId":"second","version":"2.0","store":"OMEX","storeType":"Marketplace"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture IR should parse");
        let compiled =
            crate::compiler::compile_document(&ir, Path::new(".")).expect("fixture should compile");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external task panes should reverse");
        assert!(
            jsx.contains("webExtensions={")
                && jsx.contains(r#""id":"7f33b723-fb58-4524-8733-dbedc4b7c095""#)
                && jsx.contains(r#""referenceId":"office-addin""#)
                && jsx.contains(r#""mode":"review""#)
                && jsx.contains(r#""id":"11111111-2222-3333-4444-555555555555""#),
            "web extensions must survive structural reverse: {jsx}"
        );
    }

    #[test]
    fn reverse_should_explain_web_extension_reference_without_relationship() {
        let ir: crate::ir::IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"webExtensions":[{"id":"11111111-2222-3333-4444-555555555555","referenceId":"addin","version":"1","store":"developer","storeType":"Registry"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture IR should parse");
        let compiled =
            crate::compiler::compile_document(&ir, Path::new(".")).expect("fixture should compile");
        let external = strip_ir_manifest(&compiled);
        let stripped =
            strip_package_parts(&external, &["word/webextensions/_rels/taskpanes.xml.rels"]);
        let error = reverse_document(&stripped)
            .expect_err("missing task-pane relationship must not be ignored");
        assert!(
            error.to_string().contains("relationship") && error.to_string().contains("restore"),
            "diagnostic should explain the cause and repair: {error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_custom_xml_without_item_properties() {
        let packed = pack_docx(
            Docx::new()
                .add_custom_item("06AC5857-5C65-A94A-BCEC-37356A209BC3", "<customer/>")
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let stripped = strip_package_parts(&packed, &["customXml/itemProps1.xml"]);
        let error =
            reverse_document(&stripped).expect_err("custom XML without item properties must fail");
        assert!(
            error.to_string().contains("itemProps") && error.to_string().contains("itemID"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_custom_xml_schema_refs() {
        let packed = pack_docx(
            Docx::new()
                .add_custom_item("06AC5857-5C65-A94A-BCEC-37356A209BC3", "<customer/>")
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let rewritten = rewrite_package_parts(
            &packed,
            &[(
                "customXml/itemProps1.xml",
                br#"<?xml version="1.0"?><ds:datastoreItem ds:itemID="{06AC5857-5C65-A94A-BCEC-37356A209BC3}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="https://example.com"/></ds:schemaRefs></ds:datastoreItem>"#,
            )],
        );
        let jsx = reverse_document(&rewritten).expect("schema refs should reverse");
        assert!(
            jsx.contains("schemaRefs") && jsx.contains("https://example.com"),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_positional_tab() {
        let bytes = pack_docx(Docx::new().add_paragraph(Paragraph::new().add_run(
            Run::new().add_ptab(docx_rs::PositionalTab::new(
                docx_rs::PositionalTabAlignmentType::Center,
                docx_rs::PositionalTabRelativeTo::Margin,
                docx_rs::TabLeaderType::None,
            )),
        )));
        let jsx = reverse_document(&bytes).expect("positional tab should reverse");
        assert!(
            jsx.contains("<PositionalTab") && jsx.contains(r#"align="center""#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_document_settings_and_metadata() {
        let bytes = pack_docx(
            Docx::new()
                .created_at("2026-08-14T00:00:00Z")
                .updated_at("2026-08-15T00:00:00Z")
                .custom_property("Project", "Apollo")
                .default_size(22)
                .default_spacing(10)
                .default_line_spacing(
                    LineSpacing::new()
                        .before(40)
                        .after(120)
                        .line(280)
                        .line_rule(LineSpacingType::AtLeast),
                )
                .settings(
                    Settings::new()
                        .doc_id("01234567-89AB-CDEF-0123-456789ABCDEF")
                        .default_tab_stop(720)
                        .add_doc_var("Customer", "Ada")
                        .even_and_odd_headers()
                        .adjust_line_height_in_table()
                        .character_spacing_control(CharacterSpacingValues::CompressPunctuation),
                )
                .add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let jsx = reverse_document(&bytes).expect("document settings should reverse");
        assert!(
            jsx.contains(r#"createdAt="2026-08-14T00:00:00Z""#)
                && jsx.contains(r#"updatedAt="2026-08-15T00:00:00Z""#)
                && jsx.contains(r#""Project":"Apollo""#)
                && jsx.contains(" defaultSize={11}")
                && jsx.contains(" defaultCharacterSpacing={0.5}")
                && jsx.contains("defaultLineSpacing={")
                && jsx.contains(r#""before":2"#)
                && jsx.contains(r#""after":6"#)
                && jsx.contains(r#""line":14"#)
                && jsx.contains(r#""lineRule":"atLeast""#)
                && jsx.contains(r#"documentId="01234567-89AB-CDEF-0123-456789ABCDEF""#)
                && jsx.contains(" defaultTabStop={36}")
                && jsx.contains(r#""Customer":"Ada""#)
                && jsx.contains(" evenAndOddHeaders")
                && jsx.contains(" adjustLineHeightInTable")
                && jsx.contains(r#"characterSpacingControl="compressPunctuation""#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_all_core_properties_without_manifest() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"createdAt":"2026-08-14T00:00:00Z","updatedAt":"2026-08-15T00:00:00Z","creator":"Ada & Co","lastModifiedBy":"Lin","description":"Weekly <report>","language":"zh-CN","revision":7,"subject":"Economy","title":"Status"},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["body"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, Path::new("."))
            .expect("core properties should compile");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external core properties should reverse");
        for expected in [
            r#"creator="Ada &amp; Co""#,
            r#"lastModifiedBy="Lin""#,
            r#"description="Weekly &lt;report&gt;""#,
            r#"language="zh-CN""#,
            r"revision={7}",
            r#"subject="Economy""#,
            r#"title="Status""#,
        ] {
            assert!(jsx.contains(expected), "missing {expected} in {jsx}");
        }
    }

    #[test]
    fn reverse_external_docx_should_preserve_document_protection_without_manifest() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"documentProtection":{"edit":"readOnly","formatting":false,"enforcement":true,"cryptProviderType":"rsaAES","cryptAlgorithmClass":"hash","cryptAlgorithmType":"typeAny","cryptAlgorithmSid":14,"cryptSpinCount":100000,"cryptProvider":"RSA AES Provider","algIdExt":"A1B2C3D4","algIdExtSource":"urn:example:alg","cryptProviderTypeExt":"01020304","cryptProviderTypeExtSource":"urn:example:provider","hash":"YWJjZA==","salt":"AQIDBA==","algorithmName":"SHA-512","hashValue":"YWJjZA==","saltValue":"AQIDBA==","spinCount":100000}},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["body"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, Path::new("."))
            .expect("document protection should compile");
        let external = strip_ir_manifest(&compiled);
        let jsx = reverse_document(&external).expect("external protection should reverse");
        for expected in [
            r#""edit":"readOnly""#,
            r#""formatting":false"#,
            r#""enforcement":true"#,
            r#""cryptAlgorithmSid":14"#,
            r#""cryptSpinCount":100000"#,
            r#""hash":"YWJjZA==""#,
            r#""salt":"AQIDBA==""#,
            r#""algorithmName":"SHA-512""#,
            r#""hashValue":"YWJjZA==""#,
            r#""saltValue":"AQIDBA==""#,
            r#""spinCount":100000"#,
        ] {
            assert!(jsx.contains(expected), "missing {expected} in {jsx}");
        }
    }

    #[test]
    fn reverse_external_docx_should_reject_unsupported_document_protection_attribute() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"documentProtection":{"edit":"readOnly"}},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["body"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, Path::new("."))
            .expect("document protection should compile");
        let external = strip_ir_manifest(&compiled);
        let mut archive = zip::ZipArchive::new(Cursor::new(external.clone())).expect("DOCX ZIP");
        let mut settings = String::new();
        archive
            .by_name("word/settings.xml")
            .expect("settings")
            .read_to_string(&mut settings)
            .expect("settings should be UTF-8");
        settings = settings.replace(
            "w:edit=\"readOnly\"",
            "w:edit=\"readOnly\" w:futureProtectionFlag=\"true\"",
        );
        let external =
            rewrite_package_parts(&external, &[("word/settings.xml", settings.as_bytes())]);
        let error = reverse_document(&external).expect_err("unsupported metadata must be rejected");
        assert!(
            error
                .to_string()
                .contains("attribute `futureProtectionFlag`")
                && error.to_string().contains("remove it or extend the DSL"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_incomplete_protection_hash_pair() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"documentProtection":{"edit":"readOnly","hash":"YWJjZA==","salt":"AQIDBA=="}},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["body"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let compiled = crate::compiler::compile_document(&ir, Path::new("."))
            .expect("document protection should compile");
        let external = strip_ir_manifest(&compiled);
        let mut archive = zip::ZipArchive::new(Cursor::new(external.clone())).expect("DOCX ZIP");
        let mut settings = String::new();
        archive
            .by_name("word/settings.xml")
            .expect("settings")
            .read_to_string(&mut settings)
            .expect("settings should be UTF-8");
        settings = settings.replace(" w:salt=\"AQIDBA==\"", "");
        let external =
            rewrite_package_parts(&external, &[("word/settings.xml", settings.as_bytes())]);
        let error = reverse_document(&external).expect_err("partial hash pair must be rejected");
        assert!(
            error
                .to_string()
                .contains("`hash` and `salt` must be supplied together"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_reject_non_integer_core_revision() {
        let bytes = pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("body"))),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut core = String::new();
        archive
            .by_name("docProps/core.xml")
            .expect("core properties")
            .read_to_string(&mut core)
            .expect("core properties should be UTF-8");
        core = core.replace(
            "<cp:revision>1</cp:revision>",
            "<cp:revision>latest</cp:revision>",
        );
        let malformed = rewrite_package_parts(&bytes, &[("docProps/core.xml", core.as_bytes())]);
        let error = reverse_document(&malformed).expect_err("invalid revision should be rejected");
        assert!(
            error.to_string().contains("invalid `cp:revision`")
                && error.to_string().contains("positive integer"),
            "{error}"
        );
    }

    #[test]
    fn reverse_external_docx_should_omit_stock_default_tab_stop_and_epoch_timestamps() {
        let jsx = reverse_document(&pack_docx(
            Docx::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("plain"))),
        ))
        .expect("plain document should reverse");
        assert!(
            !jsx.contains("defaultTabStop")
                && !jsx.contains("createdAt")
                && !jsx.contains("updatedAt")
                && !jsx.contains("customXmlItems")
                && !jsx.contains("documentVariables")
                && !jsx.contains("evenAndOddHeaders"),
            "backend defaults must not be invented: {jsx}"
        );
    }

    fn default_list_abstract(id: usize, ordered: bool) -> AbstractNumbering {
        let mut definition = AbstractNumbering::new(id);
        for level in 0..=8 {
            let text = if ordered {
                format!("%{}.", level + 1)
            } else {
                "•".to_owned()
            };
            definition = definition.add_level(
                Level::new(
                    level,
                    Start::new(1),
                    NumberFormat::new(if ordered { "decimal" } else { "bullet" }),
                    LevelText::new(text),
                    LevelJc::new("left"),
                )
                .indent(
                    Some(i32::try_from((level + 1) * 720).expect("indent")),
                    Some(SpecialIndentType::Hanging(360)),
                    None,
                    None,
                ),
            );
        }
        definition
    }

    fn numbered(id: usize, level: usize, text: &str) -> Paragraph {
        Paragraph::new()
            .numbering(NumberingId::new(id), IndentLevel::new(level))
            .add_run(Run::new().add_text(text))
    }

    #[test]
    fn reverse_external_docx_should_preserve_simple_ordered_list() {
        let packed = pack_docx(
            Docx::new()
                .add_abstract_numbering(default_list_abstract(2, true))
                .add_numbering(Numbering::new(2, 2))
                .add_paragraph(numbered(2, 0, "one"))
                .add_paragraph(numbered(2, 1, "two")),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(packed.clone())).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        document = document.replacen(
            "<w:pPr>",
            r#"<w:pPr><w:ind w:startChars="120" w:endChars="80" w:hangingChars="20" />"#,
            1,
        );
        let bytes = rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let jsx = reverse_document(&bytes).expect("external list should reverse");
        assert!(
            jsx.contains(r#"<List type="ordered">"#)
                && jsx.contains("<ListItem ")
                && jsx.contains("indentStartChars={120}")
                && jsx.contains("indentEndChars={80}")
                && jsx.contains("hangingChars={20}")
                && jsx.contains("level={1}")
                && jsx.contains(r#"{"one"}"#)
                && jsx.contains(r#"{"two"}"#)
                && !jsx.contains("levels="),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_custom_list_levels() {
        let definition = AbstractNumbering::new(4)
            .add_level(
                Level::new(
                    0,
                    Start::new(3),
                    NumberFormat::new("decimal"),
                    LevelText::new("%1."),
                    LevelJc::new("left"),
                )
                .suffix(LevelSuffixType::Space)
                .is_lgl()
                .level_restart(0)
                .bold(),
            )
            .add_level(
                Level::new(
                    1,
                    Start::new(1),
                    NumberFormat::new("lowerLetter"),
                    LevelText::new("%2)"),
                    LevelJc::new("right"),
                )
                .paragraph_style("ListParagraph"),
            );
        let mut definition = definition;
        definition.multi_level_type = Some("hybridMultilevel".to_owned());
        let bytes = pack_docx(
            Docx::new()
                .add_abstract_numbering(definition)
                .add_numbering(Numbering::new(4, 4))
                .add_paragraph(numbered(4, 0, "one").style("ListParagraph"))
                .add_paragraph(numbered(4, 1, "two")),
        );
        let jsx = reverse_document(&bytes).expect("custom levels should reverse");
        assert!(
            jsx.contains("levels={")
                && jsx.contains(r#"multiLevelType="hybridMultilevel""#)
                && jsx.contains(r#""format":"decimal""#)
                && jsx.contains(r#""suffix":"space""#)
                && jsx.contains(r#""legal":true"#)
                && jsx.contains(r#""format":"lowerLetter""#)
                && jsx.contains(r#""paragraphStyle":"ListParagraph""#)
                && jsx.contains(r#"style="ListParagraph""#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_instance_level_replacement() {
        let replacement = Level::new(
            0,
            Start::new(1),
            NumberFormat::new("lowerLetter"),
            LevelText::new("%1)"),
            LevelJc::new("left"),
        );
        let bytes = pack_docx(
            Docx::new()
                .add_abstract_numbering(default_list_abstract(7, true))
                .add_numbering(
                    Numbering::new(7, 7)
                        .add_override(LevelOverride::new(0).level(replacement))
                        .add_override(LevelOverride::new(1).start(4)),
                )
                .add_paragraph(numbered(7, 0, "item")),
        );
        let jsx = reverse_document(&bytes).expect("level replacement should reverse");
        assert!(
            jsx.contains(r#"overrides={[{"level":0,"definition":{"format":"lowerLetter""#,)
                && jsx.contains(r#"{"level":1,"start":4}"#),
            "{jsx}"
        );
    }

    #[test]
    fn reverse_external_docx_should_preserve_numbering_style_links() {
        let bytes = pack_docx(
            Docx::new()
                .add_abstract_numbering(AbstractNumbering::new(5))
                .add_numbering(Numbering::new(5, 5))
                .add_paragraph(numbered(5, 0, "item")),
        );
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes.clone())).expect("DOCX ZIP");
        let mut numbering = String::new();
        archive
            .by_name("word/numbering.xml")
            .expect("numbering")
            .read_to_string(&mut numbering)
            .expect("UTF-8");
        let marker = r#"<w:abstractNum w:abstractNumId="5" />"#;
        assert!(numbering.contains(marker), "{numbering}");
        numbering = numbering.replace(marker, r#"<w:abstractNum w:abstractNumId="5"><w:styleLink w:val="ParagraphList" /><w:numStyleLink w:val="ListNumber" /></w:abstractNum>"#);
        let external =
            rewrite_package_parts(&bytes, &[("word/numbering.xml", numbering.as_bytes())]);
        let jsx = reverse_document(&external).expect("linked numbering should reverse");
        assert!(jsx.contains(r#"styleLink="ParagraphList""#), "{jsx}");
        assert!(jsx.contains(r#"numStyleLink="ListNumber""#), "{jsx}");
    }

    fn strip_package_parts(bytes: &[u8], omit: &[&str]) -> Vec<u8> {
        let mut source =
            zip::ZipArchive::new(Cursor::new(bytes.to_vec())).expect("DOCX should be a ZIP");
        let output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(output);
        for index in 0..source.len() {
            let mut file = source.by_index(index).expect("entry");
            let name = file.name().to_owned();
            if omit.iter().any(|part| *part == name) {
                continue;
            }
            let mut data = Vec::new();
            file.read_to_end(&mut data).expect("read entry");
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            writer.start_file(&name, options).expect("start file");
            writer.write_all(&data).expect("write file");
        }
        writer.finish().expect("finish zip").into_inner()
    }
}
