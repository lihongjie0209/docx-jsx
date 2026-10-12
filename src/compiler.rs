use std::collections::{BTreeMap, BTreeSet};
use std::fmt::Write as _;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use docx_rs::{
    AbstractNumbering, AlignmentType, BorderType, BreakType, Caps, CellMargins,
    CharacterSpacingValues, Comment, DataBinding, Delete, DeleteInstrText, DocGrid, DocGridType,
    Docx, DrawingPosition, FieldCharType, Footer, FooterChild, Footnote, Header, HeaderChild,
    HeightRule, Hyperlink, HyperlinkType, IndentLevel, Insert, InstrPAGEREF, InstrTC, InstrText,
    InstrToC, Level, LevelJc, LevelOverride, LevelSuffixType, LevelText, LineSpacing,
    LineSpacingType, MoveFrom, MoveTo, NumPages, NumberFormat, Numbering, NumberingId, PageMargin,
    PageNum, PageNumType, PageOrientationType, PageSize, Paragraph, ParagraphBorder,
    ParagraphBorderPosition, ParagraphBorders, ParagraphPropertyChange, Pic, PicAlign,
    PositionalTab, PositionalTabAlignmentType, PositionalTabRelativeTo, RelativeFromHType,
    RelativeFromVType, Run, RunFonts, Section, Settings, Shading, ShdType, SpecialIndentType,
    Start, StructuredDataTag, Style, StyleType, StyleWithLevel, Sym, Tab as DocxTab, TabLeaderType,
    TabValueType, Table, TableAlignmentType, TableBorder, TableBorderPosition, TableBorders,
    TableCell, TableCellBorder, TableCellBorderPosition, TableCellBorders, TableCellMargins,
    TableCellProperty, TableLayoutType, TableOfContents, TableOfContentsItem,
    TablePositionProperty, TableRow, TextAlignmentType, TextBorder, TextDirectionType, ThemeColor,
    VAlignType, VMergeType, VertAlignType, WebExtension, WidthType,
};
use image::ImageFormat;
use num_traits::ToPrimitive;
use serde_json::{Map, Value};

use crate::error::{Error, Result};
use crate::ir::{Child, IrEnvelope, Node, NodeKind};

const TWIPS_PER_POINT: f64 = 20.0;
const HALF_POINTS_PER_POINT: f64 = 2.0;
const EMU_PER_POINT: f64 = 12_700.0;

struct CompileContext {
    next_numbering_id: usize,
    next_bookmark_id: usize,
    next_comment_id: usize,
    numberings: Vec<(AbstractNumbering, Numbering)>,
    character_indent_patches: Vec<ParagraphPropertyPatch>,
    numbering_style_links: Vec<(usize, Option<String>, Option<String>)>,
    used_paragraph_ids: BTreeSet<String>,
    next_generated_paragraph_id: u32,
    deleted_comment_dependencies: Vec<Comment>,
    toc_item_tab_patches: Vec<(String, u32)>,
    toc_marker_occupied_text: String,
}

struct ParagraphPropertyPatch {
    paragraph_id: String,
    indentation: BTreeMap<String, i64>,
    temporary_id: bool,
    suppress_line_numbers: Option<bool>,
}

#[derive(Clone, Debug)]
struct SectionColumnsConfig {
    count: usize,
    spacing: Option<u32>,
    separator: bool,
    widths: Vec<SectionColumnWidth>,
}

#[derive(Clone, Debug)]
struct SectionColumnWidth {
    width: Option<u32>,
    spacing: Option<u32>,
}

impl Default for SectionColumnsConfig {
    fn default() -> Self {
        Self {
            count: 1,
            spacing: Some(425),
            separator: false,
            widths: Vec::new(),
        }
    }
}

impl Default for CompileContext {
    fn default() -> Self {
        Self {
            // docx-rs always injects abstract numbering 1; user lists start at 2.
            next_numbering_id: 1,
            next_bookmark_id: 0,
            next_comment_id: 0,
            numberings: Vec::new(),
            character_indent_patches: Vec::new(),
            numbering_style_links: Vec::new(),
            used_paragraph_ids: BTreeSet::new(),
            next_generated_paragraph_id: 0xf000_0000,
            deleted_comment_dependencies: Vec::new(),
            toc_item_tab_patches: Vec::new(),
            toc_marker_occupied_text: String::new(),
        }
    }
}

/// Compiles validated IR into a DOCX archive held in memory.
///
/// # Errors
///
/// Returns an error when validation fails, an image cannot be read or decoded,
/// a scalar cannot be converted, or `docx-rs` cannot package the archive.
pub fn compile_document(ir: &IrEnvelope, entry_dir: &Path) -> Result<Vec<u8>> {
    ir.validate()?;
    let mut docx = Docx::new();
    if let Some(font) = string_prop(&ir.document.props, "defaultFont", "Document")? {
        let fonts = RunFonts::new()
            .ascii(font)
            .hi_ansi(font)
            .east_asia(font)
            .cs(font);
        docx = docx.default_fonts(fonts);
    }
    if let Some(fonts) = object_prop(&ir.document.props, "defaultFonts", "Document")? {
        docx = docx.default_fonts(compile_run_fonts(fonts, "Document/defaultFonts")?);
    }
    if let Some(size) = number_prop(&ir.document.props, "defaultSize", "Document")? {
        docx = docx.default_size(to_half_points(size, "Document/defaultSize")?);
    }
    if let Some(spacing) = number_prop(&ir.document.props, "defaultCharacterSpacing", "Document")? {
        docx = docx.default_spacing(to_twips_i32(spacing, "Document/defaultCharacterSpacing")?);
    }
    if let Some(value) = ir.document.props.get("defaultLineSpacing") {
        docx = docx.default_line_spacing(compile_default_line_spacing(value, "Document")?);
    }
    if let Some(created_at) = string_prop(&ir.document.props, "createdAt", "Document")? {
        docx = docx.created_at(created_at);
    }
    if let Some(updated_at) = string_prop(&ir.document.props, "updatedAt", "Document")? {
        docx = docx.updated_at(updated_at);
    }
    if let Some(properties) = object_prop(&ir.document.props, "customProperties", "Document")? {
        for (name, value) in properties {
            let value = value.as_str().ok_or_else(|| {
                validation("Document", "`customProperties` values must be strings")
            })?;
            docx = docx.custom_property(name, value);
        }
    }
    docx = docx.settings(compile_document_settings(&ir.document.props)?);
    if let Some(styles) = ir.document.props.get("styles").and_then(Value::as_array) {
        for (index, value) in styles.iter().enumerate() {
            docx = docx.add_style(compile_style(value, &format!("Document/styles[{index}]"))?);
        }
    }
    docx = compile_package_parts(docx, &ir.document.props)?;
    let mut context = CompileContext {
        toc_marker_occupied_text: serde_json::to_string(&ir.document)
            .map_err(|error| Error::Compile(format!("cannot inspect document text: {error}")))?,
        ..CompileContext::default()
    };
    collect_existing_paragraph_ids(&ir.document, &mut context.used_paragraph_ids);
    let section_count = ir.document.children.len();
    for (index, child) in ir.document.children.iter().enumerate() {
        let Child::Node(section_node) = child else {
            return Err(validation("Document", "expected Section"));
        };
        docx = compile_section_into_docx(
            docx,
            section_node,
            entry_dir,
            index,
            index + 1 == section_count,
            &mut context,
        )?;
    }
    for (abstract_numbering, numbering) in std::mem::take(&mut context.numberings) {
        docx = docx
            .add_abstract_numbering(abstract_numbering)
            .add_numbering(numbering);
    }
    let mut comment_seed_paragraph_ids = Vec::new();
    for comment in std::mem::take(&mut context.deleted_comment_dependencies) {
        let paragraph_id = next_generated_paragraph_id(&mut context);
        comment_seed_paragraph_ids.push(paragraph_id.clone());
        let comment_id = comment.id();
        docx = docx.add_paragraph(
            Paragraph::new()
                .id(paragraph_id)
                .add_comment_start(comment)
                .add_comment_end(comment_id),
        );
    }
    package_document(
        docx,
        ir,
        &context.character_indent_patches,
        &context.numbering_style_links,
        &comment_seed_paragraph_ids,
        &context.toc_item_tab_patches,
    )
}

fn compile_section_into_docx(
    mut docx: Docx,
    section_node: &Node,
    entry_dir: &Path,
    index: usize,
    is_last: bool,
    context: &mut CompileContext,
) -> Result<Docx> {
    for (child_index, child) in section_node.children.iter().enumerate() {
        if let Child::Node(child) = child
            && matches!(
                child.kind,
                NodeKind::TableOfContents | NodeKind::TableOfFigures | NodeKind::TableOfEntries
            )
        {
            let child_path = format!(
                "Document/Section[{index}]/{}[{child_index}]",
                child.kind.name()
            );
            let compiled_index = match child.kind {
                NodeKind::TableOfFigures => {
                    compile_table_of_figures(child, &child_path, entry_dir, context)?
                }
                NodeKind::TableOfEntries => {
                    compile_table_of_entries(child, &child_path, entry_dir, context)?
                }
                _ => compile_table_of_contents(child, &child_path, entry_dir, context)?,
            };
            docx = docx.add_table_of_contents(compiled_index);
        }
    }
    for (child_index, child) in section_node.children.iter().enumerate() {
        let Child::Node(child) = child else {
            return Err(validation(
                format!("Document/Section[{index}]"),
                "Section only accepts structural children",
            ));
        };
        let child_path = format!(
            "Document/Section[{index}]/{}[{child_index}]",
            child.kind.name()
        );
        docx = add_section_body_to_docx(docx, child, entry_dir, &child_path, context)?;
    }
    let path = format!("Document/Section[{index}]");
    if is_last {
        docx = apply_section_properties(docx, section_node, &path)?;
        for (child_index, child) in section_node.children.iter().enumerate() {
            let Child::Node(child) = child else { continue };
            let child_path = format!("{path}/{}[{child_index}]", child.kind.name());
            docx = match child.kind {
                NodeKind::Header => attach_header(docx, child, entry_dir, &child_path, context)?,
                NodeKind::Footer => attach_footer(docx, child, entry_dir, &child_path, context)?,
                _ => docx,
            };
        }
        Ok(docx)
    } else {
        let mut section = compile_section_element(section_node, &path)?;
        for (child_index, child) in section_node.children.iter().enumerate() {
            let Child::Node(child) = child else { continue };
            let child_path = format!("{path}/{}[{child_index}]", child.kind.name());
            section = match child.kind {
                NodeKind::Header => {
                    attach_header_to_section(section, child, entry_dir, &child_path, context)?
                }
                NodeKind::Footer => {
                    attach_footer_to_section(section, child, entry_dir, &child_path, context)?
                }
                _ => section,
            };
        }
        Ok(docx.add_section(section))
    }
}

fn package_document(
    docx: Docx,
    ir: &IrEnvelope,
    paragraph_character_indents: &[ParagraphPropertyPatch],
    numbering_style_links: &[(usize, Option<String>, Option<String>)],
    comment_seed_paragraph_ids: &[String],
    toc_item_tab_patches: &[(String, u32)],
) -> Result<Vec<u8>> {
    let mut cursor = Cursor::new(Vec::new());
    docx.pack(&mut cursor)
        .map_err(|error| Error::Compile(error.to_string()))?;
    let columns =
        ir.document
            .children
            .iter()
            .enumerate()
            .filter_map(|(index, child)| match child {
                Child::Node(node) if node.kind == NodeKind::Section => Some(
                    section_columns_config(node, &format!("Document/Section[{index}]")),
                ),
                Child::Node(_) | Child::String(_) | Child::Number(_) => None,
            })
            .collect::<Result<Vec<_>>>()?;
    let bytes = strip_comment_seed_paragraphs(cursor.into_inner(), comment_seed_paragraph_ids)?;
    let bytes = patch_section_columns(bytes, &columns)?;
    let section_types = ir
        .document
        .children
        .iter()
        .filter_map(|child| match child {
            Child::Node(node) if node.kind == NodeKind::Section => Some(
                node.props
                    .get("type")
                    .and_then(Value::as_str)
                    .map(ToOwned::to_owned),
            ),
            _ => None,
        })
        .collect::<Vec<_>>();
    let bytes = patch_section_types(bytes, &section_types)?;
    let section_line_numbering = ir
        .document
        .children
        .iter()
        .filter_map(|child| match child {
            Child::Node(node) if node.kind == NodeKind::Section => Some(
                node.props
                    .get("lineNumbering")
                    .and_then(Value::as_object)
                    .cloned(),
            ),
            _ => None,
        })
        .collect::<Vec<_>>();
    let bytes = patch_section_line_numbering(bytes, &section_line_numbering)?;
    let section_page_numbering = ir
        .document
        .children
        .iter()
        .filter_map(|child| match child {
            Child::Node(node) if node.kind == NodeKind::Section => Some(
                node.props
                    .get("pageNumbering")
                    .and_then(Value::as_object)
                    .cloned(),
            ),
            _ => None,
        })
        .collect::<Vec<_>>();
    let bytes = patch_section_page_numbering(bytes, &section_page_numbering)?;
    let bytes = patch_paragraph_character_indentation(bytes, paragraph_character_indents)?;
    let bytes = patch_core_properties(bytes, &ir.document.props)?;
    let bytes = patch_document_protection(bytes, &ir.document.props)?;
    let row_grid = collect_row_grid_properties(&ir.document);
    let mut bytes = bytes;
    for (part, rows) in row_grid {
        bytes = patch_table_row_grid_properties(bytes, &part, &rows)?;
    }
    let extension_count = ir
        .document
        .props
        .get("webExtensions")
        .and_then(Value::as_array)
        .map_or(0, Vec::len);
    let bytes = patch_taskpane_references(bytes, extension_count)?;
    let bytes = patch_external_hyperlink_relationships(bytes, &ir.document)?;
    let custom_xml_count = ir
        .document
        .props
        .get("customXmlItems")
        .and_then(Value::as_array)
        .map_or(0, Vec::len);
    let bytes = patch_custom_xml_content_types(bytes, custom_xml_count)?;
    let bytes = patch_custom_xml_schema_refs(bytes, &ir.document.props)?;
    let bytes = patch_duplicate_normal_style(bytes, &ir.document.props)?;
    let bytes = patch_style_character_indentation(bytes, &ir.document.props)?;
    let bytes = omit_unused_default_parts(bytes, &ir.document)?;
    let bytes = normalize_ooxml_element_order(bytes)?;
    let bytes = patch_index_dirty_values(bytes, &ir.document)?;
    let bytes = patch_toc_item_tabs(bytes, toc_item_tab_patches)?;
    let bytes = patch_numbering_style_links(bytes, numbering_style_links)?;
    let bytes = patch_simple_position_images(bytes)?;
    embed_ir_manifest(bytes, ir)
}

fn patch_index_dirty_values(bytes: Vec<u8>, document: &Node) -> Result<Vec<u8>> {
    let mut dirty_values = Vec::new();
    collect_index_dirty_values(document, &mut dirty_values);
    if !dirty_values.iter().any(|dirty| !dirty) {
        return Ok(bytes);
    }

    let mut xml = read_zip_text_entry(&bytes, "word/document.xml")?;
    let mut search_from = 0;
    let mut index_number = 0;
    while let Some(relative_start) = xml[search_from..].find("<w:sdt>") {
        let start = search_from + relative_start;
        let Some(relative_end) = xml[start..].find("</w:sdt>") else {
            return Err(Error::Compile(
                "word/document.xml contains an unclosed `w:sdt` while applying TOC dirty settings"
                    .to_owned(),
            ));
        };
        let end = start + relative_end + "</w:sdt>".len();
        let control = &xml[start..end];
        if control.contains("<w:instrText>TOC ") {
            let Some(dirty) = dirty_values.get(index_number) else {
                return Err(Error::Compile(
                    "more serialized TOC controls exist than index components in the DSL; keep the DOCX structure aligned with the document tree".to_owned(),
                ));
            };
            if !dirty {
                let Some(begin) = control.find("<w:fldChar w:fldCharType=\"begin\"") else {
                    return Err(Error::Compile(
                        "TOC control has no begin field character; repair the generated field structure".to_owned(),
                    ));
                };
                let tag_end = control[begin..]
                    .find('>')
                    .map(|offset| begin + offset + 1)
                    .ok_or_else(|| {
                        Error::Compile(
                            "TOC begin field character is malformed; repair the generated field structure".to_owned(),
                        )
                    })?;
                let tag = &control[begin..tag_end];
                if !tag.contains("w:dirty=\"true\"") {
                    return Err(Error::Compile(
                        "TOC begin field character does not contain its expected dirty flag; report this docx-rs serialization change".to_owned(),
                    ));
                }
                let patched = tag.replace("w:dirty=\"true\"", "w:dirty=\"false\"");
                let absolute_begin = start + begin;
                let absolute_end = start + tag_end;
                xml.replace_range(absolute_begin..absolute_end, &patched);
            }
            index_number += 1;
        }
        search_from = end;
    }
    if index_number != dirty_values.len() {
        return Err(Error::Compile(format!(
            "DSL defines {} TOC components, but only {index_number} were serialized; repair index placement before compiling",
            dirty_values.len()
        )));
    }
    let archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    replace_zip_text_entry(archive, "word/document.xml", xml.as_bytes())
}

fn toc_item_marker(context: &mut CompileContext) -> String {
    let mut sequence = context.toc_item_tab_patches.len();
    loop {
        let marker = format!("DOCXJSX_INTERNAL_TOC_TAB_{sequence:08X}");
        if !context.toc_marker_occupied_text.contains(&marker) {
            context.toc_marker_occupied_text.push_str(&marker);
            return marker;
        }
        sequence += 1;
    }
}

fn patch_toc_item_tabs(bytes: Vec<u8>, patches: &[(String, u32)]) -> Result<Vec<u8>> {
    if patches.is_empty() {
        return Ok(bytes);
    }
    let mut xml = read_zip_text_entry(&bytes, "word/document.xml")?;
    for (marker, tab_position) in patches {
        let Some(marker_offset) = xml.find(marker) else {
            return Err(Error::Compile(format!(
                "could not locate generated TOC item marker `{marker}`; ensure the index item remains in the document body"
            )));
        };
        let paragraph_start = xml[..marker_offset]
            .rfind("<w:p ")
            .ok_or_else(|| Error::Compile("TOC item has no containing paragraph".to_owned()))?;
        let paragraph_end = xml[marker_offset..]
            .find("</w:p>")
            .map(|offset| marker_offset + offset + "</w:p>".len())
            .ok_or_else(|| Error::Compile("TOC item paragraph is unterminated".to_owned()))?;
        let paragraph = &xml[paragraph_start..paragraph_end];
        let invalid = "w:pos=\"80000\"";
        if paragraph.matches(invalid).count() != 1 {
            return Err(Error::Compile(
                "TOC item paragraph did not contain exactly one expected invalid tab position; docx-rs serialization changed, inspect its `w:tab` output".to_owned(),
            ));
        }
        let replacement = format!("w:pos=\"{tab_position}\"");
        let mut patched_paragraph = paragraph.replace(invalid, &replacement);
        patched_paragraph = patched_paragraph.replace(marker, "");
        xml.replace_range(paragraph_start..paragraph_end, &patched_paragraph);
    }
    let archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    replace_zip_text_entry(archive, "word/document.xml", xml.as_bytes())
}

fn collect_index_dirty_values(node: &Node, values: &mut Vec<bool>) {
    if matches!(
        node.kind,
        NodeKind::TableOfContents | NodeKind::TableOfFigures | NodeKind::TableOfEntries
    ) {
        values.push(
            node.props
                .get("dirty")
                .and_then(Value::as_bool)
                .unwrap_or(true),
        );
    }
    for child in &node.children {
        if let Child::Node(child) = child {
            collect_index_dirty_values(child, values);
        }
    }
}

fn patch_document_protection(bytes: Vec<u8>, props: &Map<String, Value>) -> Result<Vec<u8>> {
    let Some(protection) = props.get("documentProtection").and_then(Value::as_object) else {
        return Ok(bytes);
    };
    let mut attributes = String::new();
    for (property, attribute) in [
        ("edit", "edit"),
        ("formatting", "formatting"),
        ("enforcement", "enforcement"),
        ("cryptProviderType", "cryptProviderType"),
        ("cryptAlgorithmClass", "cryptAlgorithmClass"),
        ("cryptAlgorithmType", "cryptAlgorithmType"),
        ("cryptAlgorithmSid", "cryptAlgorithmSid"),
        ("cryptSpinCount", "cryptSpinCount"),
        ("cryptProvider", "cryptProvider"),
        ("algIdExt", "algIdExt"),
        ("algIdExtSource", "algIdExtSource"),
        ("cryptProviderTypeExt", "cryptProviderTypeExt"),
        ("cryptProviderTypeExtSource", "cryptProviderTypeExtSource"),
        ("hash", "hash"),
        ("salt", "salt"),
        ("algorithmName", "algorithmName"),
        ("hashValue", "hashValue"),
        ("saltValue", "saltValue"),
        ("spinCount", "spinCount"),
    ] {
        let Some(value) = protection.get(property) else {
            continue;
        };
        let value = match value {
            Value::Bool(value) => if *value { "1" } else { "0" }.to_owned(),
            Value::Number(value) => value.to_string(),
            Value::String(value) => value.clone(),
            Value::Null | Value::Array(_) | Value::Object(_) => {
                return Err(validation(
                    "Document/documentProtection",
                    format!("`{property}` has an unsupported value type"),
                ));
            }
        };
        let _ = write!(attributes, r#" w:{attribute}="{}""#, xml_escape(&value));
    }
    let element = format!("<w:documentProtection{attributes} />");
    let mut settings = read_zip_text_entry(&bytes, "word/settings.xml")?;
    if settings.contains("<w:documentProtection") {
        return Err(Error::Compile(
            "docx-rs unexpectedly emitted `w:documentProtection`; remove the duplicate settings element before compiling".to_owned(),
        ));
    }
    let Some(root_start) = settings.find("<w:settings") else {
        return Err(Error::Compile(
            "word/settings.xml is missing its `w:settings` root; cannot insert document protection"
                .to_owned(),
        ));
    };
    let root_end = settings[root_start..]
        .find('>')
        .map(|position| root_start + position + 1)
        .ok_or_else(|| {
            Error::Compile("word/settings.xml has an unclosed `w:settings` root".to_owned())
        })?;
    let insertion = settings.find("<w:defaultTabStop").unwrap_or(root_end);
    settings.insert_str(insertion, &element);
    let archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    replace_zip_text_entry(archive, "word/settings.xml", settings.as_bytes())
}

fn patch_core_properties(bytes: Vec<u8>, props: &Map<String, Value>) -> Result<Vec<u8>> {
    let fields = [
        ("creator", "dc:creator"),
        ("lastModifiedBy", "cp:lastModifiedBy"),
        ("description", "dc:description"),
        ("language", "dc:language"),
        ("subject", "dc:subject"),
        ("title", "dc:title"),
    ];
    let mut xml = read_zip_text_entry(&bytes, "docProps/core.xml")?;
    let mut changed = false;
    for (property, element) in fields {
        if let Some(value) = props.get(property).and_then(Value::as_str) {
            replace_xml_element_text(&mut xml, element, value)?;
            changed = true;
        }
    }
    if let Some(revision) = props.get("revision").and_then(Value::as_u64) {
        replace_xml_element_text(&mut xml, "cp:revision", &revision.to_string())?;
        changed = true;
    }
    if !changed {
        return Ok(bytes);
    }
    let archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    replace_zip_text_entry(archive, "docProps/core.xml", xml.as_bytes())
}

fn read_zip_text_entry(bytes: &[u8], name: &str) -> Result<String> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut xml = String::new();
    archive
        .by_name(name)
        .map_err(|error| Error::Compile(format!("cannot read {name}: {error}")))?
        .read_to_string(&mut xml)
        .map_err(|error| Error::Compile(format!("cannot decode {name}: {error}")))?;
    Ok(xml)
}

fn replace_xml_element_text(xml: &mut String, element: &str, value: &str) -> Result<()> {
    let open = format!("<{element}>");
    let close = format!("</{element}>");
    let Some(open_start) = xml.find(&open) else {
        let root_close = "</cp:coreProperties>";
        let insertion = xml.find(root_close).ok_or_else(|| {
            Error::Compile("docProps/core.xml is missing `cp:coreProperties`".to_owned())
        })?;
        xml.insert_str(
            insertion,
            &format!("<{element}>{}</{element}>", xml_escape(value)),
        );
        return Ok(());
    };
    let start = open_start + open.len();
    let end = xml[start..]
        .find(&close)
        .map(|end| start + end)
        .ok_or_else(|| Error::Compile(format!("docProps/core.xml has an unclosed `{element}`")))?;
    xml.replace_range(start..end, &xml_escape(value));
    Ok(())
}

fn collect_existing_paragraph_ids(node: &Node, ids: &mut BTreeSet<String>) {
    if let Some(id) = node.props.get("paragraphId").and_then(Value::as_str) {
        ids.insert(id.to_ascii_uppercase());
    }
    for child in &node.children {
        if let Child::Node(child) = child {
            collect_existing_paragraph_ids(child, ids);
        }
    }
}

fn paragraph_character_indent_attributes(props: &Map<String, Value>) -> BTreeMap<String, i64> {
    [
        ("indentLeftChars", "leftChars"),
        ("indentStartChars", "startChars"),
        ("indentRightChars", "rightChars"),
        ("indentEndChars", "endChars"),
        ("firstLineChars", "firstLineChars"),
        ("hangingChars", "hangingChars"),
    ]
    .into_iter()
    .filter_map(|(source, target)| {
        props
            .get(source)
            .and_then(Value::as_i64)
            .map(|value| (target.to_owned(), value))
    })
    .collect()
}

fn register_paragraph_character_indentation(
    mut paragraph: Paragraph,
    node: &Node,
    context: &mut CompileContext,
) -> (Paragraph, bool) {
    let attributes = paragraph_character_indent_attributes(&node.props);
    let has_character_indents = !attributes.is_empty();
    let suppress_line_numbers = node
        .props
        .get("suppressLineNumbers")
        .and_then(Value::as_bool);
    if attributes.is_empty() && suppress_line_numbers.is_none() {
        return (paragraph, false);
    }
    let (id, is_temporary) = if let Some(id) = node.props.get("paragraphId").and_then(Value::as_str)
    {
        (id.to_ascii_uppercase(), false)
    } else {
        (next_generated_paragraph_id(context), true)
    };
    paragraph = paragraph.id(id.clone());
    context
        .character_indent_patches
        .push(ParagraphPropertyPatch {
            paragraph_id: id,
            indentation: attributes,
            temporary_id: is_temporary,
            suppress_line_numbers,
        });
    (paragraph, has_character_indents)
}

fn next_generated_paragraph_id(context: &mut CompileContext) -> String {
    loop {
        let id = format!("{:08X}", context.next_generated_paragraph_id);
        context.next_generated_paragraph_id = context.next_generated_paragraph_id.wrapping_sub(1);
        if context.used_paragraph_ids.insert(id.clone()) {
            return id;
        }
    }
}

fn patch_paragraph_character_indentation(
    bytes: Vec<u8>,
    patches: &[ParagraphPropertyPatch],
) -> Result<Vec<u8>> {
    if patches.is_empty() {
        return Ok(bytes);
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut replacements = Vec::new();
    let mut matched = BTreeSet::new();
    for index in 0..archive.len() {
        let name = archive
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?
            .name()
            .to_owned();
        if !(name == "word/document.xml"
            || (name.starts_with("word/header") || name.starts_with("word/footer"))
                && Path::new(&name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
                && !name.contains("/_rels/"))
        {
            continue;
        }
        let mut xml = String::new();
        archive
            .by_name(&name)
            .map_err(|error| Error::Compile(format!("missing {name}: {error}")))?
            .read_to_string(&mut xml)
            .map_err(|error| Error::Compile(format!("cannot read {name}: {error}")))?;
        let mut changed = false;
        for patch in patches {
            if patch_paragraph_indentation_xml(
                &mut xml,
                &patch.paragraph_id,
                &patch.indentation,
                patch.temporary_id,
                patch.suppress_line_numbers,
            )? {
                matched.insert(patch.paragraph_id.as_str());
                changed = true;
            }
        }
        if changed {
            replacements.push((name, xml));
        }
    }
    for patch in patches {
        if !matched.contains(patch.paragraph_id.as_str()) {
            return Err(Error::Compile(format!(
                "docx-rs omitted the paragraph marker for paragraph properties `{}`",
                patch.paragraph_id
            )));
        }
    }
    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..archive.len() {
        let file = archive
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        if let Some((_, xml)) = replacements.iter().find(|(name, _)| name == file.name()) {
            let name = file.name().to_owned();
            let options = file.options();
            writer
                .start_file(name, options)
                .map_err(|error| Error::Compile(format!("cannot write paragraph XML: {error}")))?;
            writer
                .write_all(xml.as_bytes())
                .map_err(|error| Error::Compile(format!("cannot write paragraph XML: {error}")))?;
        } else {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
        }
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn patch_paragraph_indentation_xml(
    xml: &mut String,
    id: &str,
    attributes: &BTreeMap<String, i64>,
    temporary_id: bool,
    suppress_line_numbers: Option<bool>,
) -> Result<bool> {
    let marker = format!(r#"w14:paraId="{id}""#);
    let Some(marker_start) = xml.find(&marker) else {
        return Ok(false);
    };
    let paragraph_start = xml[..marker_start]
        .rfind("<w:p")
        .ok_or_else(|| Error::Compile(format!("invalid paragraph marker `{id}`")))?;
    let start_tag_end = xml[paragraph_start..]
        .find('>')
        .map(|at| paragraph_start + at + 1)
        .ok_or_else(|| Error::Compile(format!("invalid paragraph start for `{id}`")))?;
    let paragraph_end = xml[start_tag_end..]
        .find("</w:p>")
        .map(|at| start_tag_end + at)
        .ok_or_else(|| Error::Compile(format!("invalid paragraph end for `{id}`")))?;
    let mut attrs = String::new();
    for (name, value) in attributes {
        write!(attrs, " w:{name}=\"{value}\"").expect("String write");
    }
    let suppress_xml = suppress_line_numbers
        .map(|value| {
            if value {
                "<w:suppressLineNumbers />"
            } else {
                "<w:suppressLineNumbers w:val=\"0\" />"
            }
        })
        .unwrap_or_default();
    let paragraph = &xml[paragraph_start..paragraph_end];
    if let Some(ppr_relative) = paragraph.find("<w:pPr") {
        let ppr_start = paragraph_start + ppr_relative;
        let Some(ppr_open_end) = xml[ppr_start..paragraph_end]
            .find('>')
            .map(|at| ppr_start + at + 1)
        else {
            return Err(Error::Compile(format!(
                "invalid w:pPr for paragraph `{id}`"
            )));
        };
        let ppr_is_empty = xml[ppr_start..ppr_open_end].ends_with("/>");
        if ppr_is_empty {
            let indentation = if attrs.is_empty() {
                String::new()
            } else {
                format!("<w:ind{attrs} />")
            };
            xml.replace_range(
                ppr_start..ppr_open_end,
                &format!("<w:pPr>{indentation}{suppress_xml}</w:pPr>"),
            );
        } else if let Some(ind_relative) = xml[ppr_open_end..paragraph_end].find("<w:ind") {
            let ind_start = ppr_open_end + ind_relative;
            let mut ind_end = xml[ind_start..paragraph_end]
                .find('>')
                .map(|at| ind_start + at)
                .ok_or_else(|| Error::Compile(format!("invalid w:ind for paragraph `{id}`")))?;
            if xml[ind_end.saturating_sub(1)..=ind_end].ends_with("/>") {
                ind_end -= 1;
            }
            xml.insert_str(ind_end, &attrs);
            if !suppress_xml.is_empty() {
                let ppr_end = xml[ppr_open_end..paragraph_end]
                    .find("</w:pPr>")
                    .map(|at| ppr_open_end + at)
                    .ok_or_else(|| Error::Compile(format!("invalid w:pPr for paragraph `{id}`")))?;
                xml.insert_str(ppr_end, suppress_xml);
            }
        } else {
            let ppr_end = xml[ppr_open_end..paragraph_end]
                .find("</w:pPr>")
                .map(|at| ppr_open_end + at)
                .ok_or_else(|| Error::Compile(format!("invalid w:pPr for paragraph `{id}`")))?;
            xml.insert_str(ppr_end, &format!("<w:ind{attrs} />{suppress_xml}"));
        }
    } else {
        let indentation = if attrs.is_empty() {
            String::new()
        } else {
            format!("<w:ind{attrs} />")
        };
        xml.insert_str(
            start_tag_end,
            &format!("<w:pPr>{indentation}{suppress_xml}</w:pPr>"),
        );
    }
    if temporary_id {
        let marker_with_space = format!(" {marker}");
        if let Some(marker_start) = xml.find(&marker_with_space) {
            xml.replace_range(marker_start..marker_start + marker_with_space.len(), "");
        }
    }
    Ok(true)
}

fn collect_row_grid_properties(root: &Node) -> Vec<(String, Vec<[Option<u64>; 4]>)> {
    fn visit(node: &Node, rows: &mut Vec<[Option<u64>; 4]>) {
        if node.kind == NodeKind::TableRow {
            rows.push([
                node.props.get("gridBefore").and_then(Value::as_u64),
                node.props.get("gridAfter").and_then(Value::as_u64),
                node.props
                    .get("widthBefore")
                    .and_then(Value::as_f64)
                    .and_then(|v| (v * 20.0).round().to_u64()),
                node.props
                    .get("widthAfter")
                    .and_then(Value::as_f64)
                    .and_then(|v| (v * 20.0).round().to_u64()),
            ]);
        }
        for child in &node.children {
            if let Child::Node(child) = child {
                visit(child, rows);
            }
        }
    }
    let mut output = Vec::new();
    let mut document_rows = Vec::new();
    let (mut header_count, mut footer_count) = (0, 0);
    for child in &root.children {
        let Child::Node(section) = child else {
            continue;
        };
        if section.kind != NodeKind::Section {
            continue;
        }
        for child in &section.children {
            let Child::Node(node) = child else { continue };
            match node.kind {
                NodeKind::Header => {
                    header_count += 1;
                    let mut rows = Vec::new();
                    visit(node, &mut rows);
                    output.push((format!("word/header{header_count}.xml"), rows));
                }
                NodeKind::Footer => {
                    footer_count += 1;
                    let mut rows = Vec::new();
                    visit(node, &mut rows);
                    output.push((format!("word/footer{footer_count}.xml"), rows));
                }
                _ => visit(node, &mut document_rows),
            }
        }
    }
    output.insert(0, ("word/document.xml".to_owned(), document_rows));
    output
}

fn patch_table_row_grid_properties(
    bytes: Vec<u8>,
    part: &str,
    rows: &[[Option<u64>; 4]],
) -> Result<Vec<u8>> {
    if !rows.iter().any(|row| row.iter().any(Option::is_some)) {
        return Ok(bytes);
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut xml = String::new();
    archive
        .by_name(part)
        .map_err(|error| Error::Compile(format!("missing {part}: {error}")))?
        .read_to_string(&mut xml)
        .map_err(|error| Error::Compile(format!("cannot read {part}: {error}")))?;
    let mut starts = Vec::new();
    let mut offset = 0;
    while let Some(relative) = xml[offset..].find("<w:tr") {
        let start = offset + relative;
        if xml
            .as_bytes()
            .get(start + 5)
            .is_some_and(|byte| byte.is_ascii_whitespace() || *byte == b'>')
        {
            starts.push(start);
        }
        offset = start + 5;
    }
    if starts.len() != rows.len() {
        return Err(Error::Compile(format!(
            "table row count mismatch while applying row grid properties: JSX has {}, DOCX has {}",
            rows.len(),
            starts.len()
        )));
    }
    for (index, properties) in rows.iter().enumerate().rev() {
        if properties.iter().all(Option::is_none) {
            continue;
        }
        let end = xml[starts[index]..]
            .find("</w:tr>")
            .map(|at| starts[index] + at)
            .ok_or_else(|| {
                Error::Compile(
                    "invalid table row XML while applying row grid properties".to_owned(),
                )
            })?;
        let open_end = xml[starts[index]..end]
            .find('>')
            .map(|at| starts[index] + at + 1)
            .ok_or_else(|| Error::Compile("invalid table row start tag".to_owned()))?;
        let close = xml[open_end..end]
            .rfind("</w:trPr>")
            .map(|at| open_end + at);
        let insert_at = if let Some(close) = close {
            close
        } else {
            let marker = "<w:trPr />";
            let at = xml[open_end..end]
                .find(marker)
                .map(|at| open_end + at)
                .ok_or_else(|| {
                    Error::Compile("docx-rs omitted the expected w:trPr element".to_owned())
                })?;
            xml.replace_range(at..at + marker.len(), "<w:trPr></w:trPr>");
            at + "<w:trPr>".len()
        };
        let mut fragment = String::new();
        for (slot, name) in ["gridBefore", "gridAfter", "wBefore", "wAfter"]
            .iter()
            .enumerate()
        {
            if let Some(value) = properties[slot] {
                if slot < 2 {
                    write!(fragment, "<w:{name} w:val=\"{value}\" />").expect("String write");
                } else {
                    write!(fragment, "<w:{name} w:w=\"{value}\" w:type=\"dxa\" />")
                        .expect("String write");
                }
            }
        }
        xml.insert_str(insert_at, &fragment);
    }
    replace_zip_text_entry(archive, part, xml.as_bytes())
}

fn section_columns_config(node: &Node, path: &str) -> Result<SectionColumnsConfig> {
    let Some(value) = node.props.get("columns") else {
        return Ok(SectionColumnsConfig::default());
    };
    if let Some(count) = value.as_u64() {
        return Ok(SectionColumnsConfig {
            count: usize::try_from(count)
                .map_err(|_| validation(path, "`columns` count is out of range"))?,
            ..SectionColumnsConfig::default()
        });
    }
    let columns = value
        .as_object()
        .ok_or_else(|| validation(path, "`columns` must be an integer or object"))?;
    if let Some(count) = columns.get("count").and_then(Value::as_u64) {
        let spacing = columns
            .get("spacing")
            .and_then(Value::as_f64)
            .unwrap_or(0.0);
        return Ok(SectionColumnsConfig {
            count: usize::try_from(count)
                .map_err(|_| validation(path, "`columns.count` is out of range"))?,
            spacing: Some(to_twips_u32(spacing, &format!("{path}/columns/spacing"))?),
            separator: columns
                .get("separator")
                .and_then(Value::as_bool)
                .unwrap_or(false),
            widths: Vec::new(),
        });
    }
    let widths = columns
        .get("widths")
        .and_then(Value::as_array)
        .ok_or_else(|| validation(path, "`columns.widths` must be an array"))?;
    let column_widths = widths
        .iter()
        .enumerate()
        .map(|(index, value)| {
            let column = value.as_object().ok_or_else(|| {
                validation(path, format!("`columns.widths[{index}]` must be an object"))
            })?;
            let measure = |key: &str| -> Result<Option<u32>> {
                column
                    .get(key)
                    .and_then(Value::as_f64)
                    .map(|points| {
                        to_twips_u32(points, &format!("{path}/columns/widths[{index}]/{key}"))
                    })
                    .transpose()
            };
            Ok(SectionColumnWidth {
                width: measure("width")?,
                spacing: measure("spacing")?,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    Ok(SectionColumnsConfig {
        count: column_widths.len(),
        spacing: None,
        separator: columns
            .get("separator")
            .and_then(Value::as_bool)
            .unwrap_or(false),
        widths: column_widths,
    })
}

fn patch_section_columns(bytes: Vec<u8>, columns: &[SectionColumnsConfig]) -> Result<Vec<u8>> {
    if columns
        .iter()
        .all(|section| section.count == 1 && section.widths.is_empty() && !section.separator)
    {
        return Ok(bytes);
    }
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut document = String::new();
    source
        .by_name("word/document.xml")
        .map_err(|error| Error::Compile(format!("missing word/document.xml: {error}")))?
        .read_to_string(&mut document)
        .map_err(|error| Error::Compile(format!("cannot read word/document.xml: {error}")))?;
    let mut offset = 0;
    let mut column_index = 0;
    while let Some(relative_start) = document[offset..].find("<w:cols") {
        let start = offset + relative_start;
        let open_end = document[start..]
            .find('>')
            .map(|relative_end| start + relative_end + 1)
            .ok_or_else(|| {
                Error::Compile("invalid `w:cols` element in word/document.xml".to_owned())
            })?;
        let config = columns.get(column_index).ok_or_else(|| {
            Error::Compile("DOCX section count does not match the JSX Section count".to_owned())
        })?;
        let is_empty = document[start..open_end].ends_with("/>");
        let end = if is_empty {
            open_end
        } else {
            document[open_end..]
                .find("</w:cols>")
                .map(|relative| open_end + relative + "</w:cols>".len())
                .ok_or_else(|| {
                    Error::Compile("invalid `w:cols` element in word/document.xml".to_owned())
                })?
        };
        let replacement = section_columns_xml(config);
        document.replace_range(start..end, &replacement);
        offset = start + replacement.len();
        column_index += 1;
    }
    if column_index != columns.len() {
        return Err(Error::Compile(format!(
            "docx-rs emitted {column_index} section column elements for {} JSX Sections",
            columns.len()
        )));
    }
    replace_zip_text_entry(source, "word/document.xml", document.as_bytes())
}

fn section_columns_xml(config: &SectionColumnsConfig) -> String {
    if config.widths.is_empty() {
        let mut xml = String::from("<w:cols");
        if let Some(spacing) = config.spacing {
            write!(xml, " w:space=\"{spacing}\"").expect("String write");
        }
        write!(xml, " w:num=\"{}\"", config.count).expect("String write");
        if config.separator {
            xml.push_str(" w:sep=\"1\"");
        }
        xml.push_str(" />");
        return xml;
    }
    let mut xml = format!("<w:cols w:num=\"{}\"", config.count);
    xml.push_str(" w:equalWidth=\"0\"");
    if config.separator {
        xml.push_str(" w:sep=\"1\"");
    }
    xml.push('>');
    for column in &config.widths {
        xml.push_str("<w:col");
        if let Some(width) = column.width {
            write!(xml, " w:w=\"{width}\"").expect("String write");
        }
        if let Some(spacing) = column.spacing {
            write!(xml, " w:space=\"{spacing}\"").expect("String write");
        }
        xml.push_str(" />");
    }
    xml.push_str("</w:cols>");
    xml
}

fn patch_section_types(bytes: Vec<u8>, section_types: &[Option<String>]) -> Result<Vec<u8>> {
    if section_types.iter().all(Option::is_none) {
        return Ok(bytes);
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut xml = String::new();
    archive
        .by_name("word/document.xml")
        .map_err(|error| Error::Compile(format!("missing word/document.xml: {error}")))?
        .read_to_string(&mut xml)
        .map_err(|error| Error::Compile(format!("cannot read word/document.xml: {error}")))?;
    let mut search_from = 0;
    let mut section_index = 0;
    while let Some(relative_start) = xml[search_from..].find("<w:sectPr") {
        let start = search_from + relative_start;
        let after_name = start + "<w:sectPr".len();
        if xml
            .as_bytes()
            .get(after_name)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && *byte != b'>')
        {
            search_from = after_name;
            continue;
        }
        let close = xml[start..]
            .find("</w:sectPr>")
            .map(|offset| start + offset)
            .ok_or_else(|| Error::Compile("invalid `w:sectPr` in word/document.xml".to_owned()))?;
        let Some(section_type) = section_types.get(section_index).and_then(Option::as_deref) else {
            section_index += 1;
            search_from = close + "</w:sectPr>".len();
            continue;
        };
        if xml[start..close].contains("<w:type ") {
            return Err(Error::Compile(format!(
                "docx-rs unexpectedly emitted `w:type` for Section[{section_index}]"
            )));
        }
        let insertion = [
            "<w:pgSz",
            "<w:pgMar",
            "<w:paperSrc",
            "<w:pgBorders",
            "<w:lnNumType",
            "<w:pgNumType",
            "<w:cols",
            "<w:formProt",
            "<w:vAlign",
            "<w:noEndnote",
            "<w:titlePg",
            "<w:textDirection",
            "<w:bidi",
            "<w:rtlGutter",
            "<w:docGrid",
        ]
        .iter()
        .filter_map(|element| xml[start..close].find(element).map(|offset| start + offset))
        .min()
        .ok_or_else(|| {
            Error::Compile(format!(
                "cannot locate schema position for `w:type` in Section[{section_index}]"
            ))
        })?;
        let section_type_xml = format!(r#"<w:type w:val="{}" />"#, xml_escape(section_type));
        xml.insert_str(insertion, &section_type_xml);
        search_from = insertion + section_type_xml.len();
        section_index += 1;
    }
    if section_index < section_types.len() {
        return Err(Error::Compile(format!(
            "docx-rs emitted {section_index} section properties for {} JSX Sections",
            section_types.len()
        )));
    }
    replace_zip_text_entry(archive, "word/document.xml", xml.as_bytes())
}

fn patch_section_line_numbering(
    bytes: Vec<u8>,
    line_numbering: &[Option<Map<String, Value>>],
) -> Result<Vec<u8>> {
    if line_numbering.iter().all(Option::is_none) {
        return Ok(bytes);
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut xml = String::new();
    archive
        .by_name("word/document.xml")
        .map_err(|error| Error::Compile(format!("missing word/document.xml: {error}")))?
        .read_to_string(&mut xml)
        .map_err(|error| Error::Compile(format!("cannot read word/document.xml: {error}")))?;
    let mut search_from = 0;
    let mut section_index = 0;
    while let Some(relative_start) = xml[search_from..].find("<w:sectPr") {
        let start = search_from + relative_start;
        let after_name = start + "<w:sectPr".len();
        if xml
            .as_bytes()
            .get(after_name)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && *byte != b'>')
        {
            search_from = after_name;
            continue;
        }
        let close = xml[start..]
            .find("</w:sectPr>")
            .map(|offset| start + offset)
            .ok_or_else(|| Error::Compile("invalid `w:sectPr` in word/document.xml".to_owned()))?;
        let Some(properties) = line_numbering.get(section_index).and_then(Option::as_ref) else {
            section_index += 1;
            search_from = close + "</w:sectPr>".len();
            continue;
        };
        if xml[start..close].contains("<w:lnNumType") {
            return Err(Error::Compile(format!(
                "docx-rs unexpectedly emitted `w:lnNumType` for Section[{section_index}]"
            )));
        }
        let mut element = String::from("<w:lnNumType");
        for key in ["countBy", "start"] {
            if let Some(value) = properties.get(key).and_then(Value::as_u64) {
                let _ = write!(element, r#" w:{key}="{value}""#);
            }
        }
        if let Some(distance) = properties.get("distance").and_then(Value::as_f64) {
            let distance = to_twips_i32(
                distance,
                &format!("Section[{section_index}]/lineNumbering/distance"),
            )?;
            let _ = write!(element, r#" w:distance="{distance}""#);
        }
        if let Some(restart) = properties.get("restart").and_then(Value::as_str) {
            let _ = write!(element, r#" w:restart="{}""#, xml_escape(restart));
        }
        element.push_str(" />");
        let insertion = [
            "<w:lnNumType",
            "<w:pgNumType",
            "<w:cols",
            "<w:formProt",
            "<w:vAlign",
            "<w:noEndnote",
            "<w:titlePg",
            "<w:textDirection",
            "<w:bidi",
            "<w:rtlGutter",
            "<w:docGrid",
        ]
        .iter()
        .filter_map(|tag| xml[start..close].find(tag).map(|offset| start + offset))
        .min()
        .ok_or_else(|| {
            Error::Compile(format!(
                "cannot locate schema position for `w:lnNumType` in Section[{section_index}]"
            ))
        })?;
        xml.insert_str(insertion, &element);
        search_from = insertion + element.len();
        section_index += 1;
    }
    if section_index < line_numbering.len() {
        return Err(Error::Compile(format!(
            "docx-rs emitted {section_index} section properties for {} JSX Sections",
            line_numbering.len()
        )));
    }
    replace_zip_text_entry(archive, "word/document.xml", xml.as_bytes())
}

fn patch_section_page_numbering(
    bytes: Vec<u8>,
    page_numbering: &[Option<Map<String, Value>>],
) -> Result<Vec<u8>> {
    if page_numbering.iter().all(Option::is_none) {
        return Ok(bytes);
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut xml = String::new();
    archive
        .by_name("word/document.xml")
        .map_err(|_| Error::Compile("missing word/document.xml".to_owned()))?
        .read_to_string(&mut xml)
        .map_err(|error| Error::Compile(format!("cannot read document XML: {error}")))?;

    let mut search_from = 0;
    let mut section_index = 0;
    while let Some(relative_start) = xml[search_from..].find("<w:sectPr") {
        let section_start = search_from + relative_start;
        let after_name = section_start + "<w:sectPr".len();
        if xml
            .as_bytes()
            .get(after_name)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && *byte != b'>')
        {
            search_from = after_name;
            continue;
        }
        let section_end = xml[section_start..]
            .find("</w:sectPr>")
            .map(|offset| section_start + offset)
            .ok_or_else(|| Error::Compile("invalid `w:sectPr` in word/document.xml".to_owned()))?;
        let properties = page_numbering.get(section_index).and_then(Option::as_ref);
        if let Some(properties) = properties {
            let attributes = [("format", "fmt"), ("chapterSeparator", "chapSep")]
                .into_iter()
                .filter_map(|(property, attribute)| {
                    properties
                        .get(property)
                        .and_then(Value::as_str)
                        .map(|value| (attribute, value))
                })
                .collect::<Vec<_>>();
            if attributes.is_empty() {
                search_from = section_end + "</w:sectPr>".len();
            } else {
                let relative_element_start = xml[section_start..section_end]
                    .find("<w:pgNumType")
                    .ok_or_else(|| {
                    Error::Compile(format!(
                        "docx-rs emitted no `w:pgNumType` for Section[{section_index}]"
                    ))
                })?;
                let element_start = section_start + relative_element_start;
                let element_end = xml[element_start..]
                    .find('>')
                    .map(|offset| element_start + offset + 1)
                    .ok_or_else(|| {
                        Error::Compile(format!(
                            "malformed `w:pgNumType` for Section[{section_index}]"
                        ))
                    })?;
                let tag = &xml[element_start..element_end];
                let mut additions = String::new();
                for (attribute, value) in attributes {
                    let name = format!("w:{attribute}=");
                    if tag.contains(&name) {
                        return Err(Error::Compile(format!(
                            "docx-rs unexpectedly emitted `{name}` for Section[{section_index}]"
                        )));
                    }
                    let _ = write!(additions, " w:{attribute}=\"{}\"", xml_escape(value));
                }
                let insert_at = if tag.ends_with(" />") {
                    element_end - 3
                } else {
                    element_end - 1
                };
                xml.insert_str(insert_at, &additions);
                search_from = element_end + additions.len();
            }
        } else {
            search_from = section_end + "</w:sectPr>".len();
        }
        section_index += 1;
    }
    if section_index < page_numbering.len() {
        return Err(Error::Compile(format!(
            "docx-rs emitted {section_index} section properties for {} JSX Sections",
            page_numbering.len()
        )));
    }
    replace_zip_text_entry(archive, "word/document.xml", xml.as_bytes())
}

fn patch_taskpane_references(bytes: Vec<u8>, extension_count: usize) -> Result<Vec<u8>> {
    if extension_count <= 1 {
        return Ok(bytes);
    }
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut taskpanes = String::new();
    source
        .by_name("word/webextensions/taskpanes.xml")
        .map_err(|error| Error::Compile(format!("missing task-pane catalog: {error}")))?
        .read_to_string(&mut taskpanes)
        .map_err(|error| Error::Compile(format!("cannot read task-pane catalog: {error}")))?;
    let closing = "</wetp:taskpanes>";
    let end = taskpanes
        .rfind(closing)
        .ok_or_else(|| Error::Compile("invalid task-pane catalog XML".to_owned()))?;
    let mut additional = String::new();
    for index in 2..=extension_count {
        write!(additional, r#"<wetp:taskpane dockstate="" visibility="1" width="350" row="1"><wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId{index}" /></wetp:taskpane>"#)
            .expect("writing to String cannot fail");
    }
    taskpanes.insert_str(end, &additional);

    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        if file.name() == "word/webextensions/taskpanes.xml" {
            let options = file.options();
            writer.start_file(file.name(), options).map_err(|error| {
                Error::Compile(format!("cannot write task-pane catalog: {error}"))
            })?;
            writer.write_all(taskpanes.as_bytes()).map_err(|error| {
                Error::Compile(format!("cannot write task-pane catalog: {error}"))
            })?;
        } else {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
        }
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn compile_package_parts(mut docx: Docx, props: &Map<String, Value>) -> Result<Docx> {
    if let Some(extensions) = props.get("webExtensions").and_then(Value::as_array) {
        docx = docx.taskpanes();
        for (index, value) in extensions.iter().enumerate() {
            let path = format!("Document/webExtensions[{index}]");
            let extension = value
                .as_object()
                .ok_or_else(|| validation(&path, "web extension must be an object"))?;
            let mut output = WebExtension::new(
                required_string(extension, "id", &path)?,
                required_string(extension, "referenceId", &path)?,
                required_string(extension, "version", &path)?,
                required_string(extension, "store", &path)?,
                required_string(extension, "storeType", &path)?,
            );
            if let Some(properties) = object_prop(extension, "properties", &path)? {
                for (name, value) in properties {
                    let value = value.as_str().ok_or_else(|| {
                        validation(&path, "web extension property values must be strings")
                    })?;
                    output = output.property(name, value);
                }
            }
            docx = docx.web_extension(output);
        }
    }
    if let Some(items) = props.get("customXmlItems").and_then(Value::as_array) {
        for (index, value) in items.iter().enumerate() {
            let path = format!("Document/customXmlItems[{index}]");
            let item = value
                .as_object()
                .ok_or_else(|| validation(&path, "custom XML item must be an object"))?;
            docx = docx.add_custom_item(
                required_string(item, "id", &path)?,
                required_string(item, "xml", &path)?,
            );
        }
    }
    Ok(docx)
}

fn compile_document_settings(props: &Map<String, Value>) -> Result<Settings> {
    let path = "Document";
    let mut settings = Settings::new();
    if let Some(document_id) = string_prop(props, "documentId", path)? {
        settings = settings.doc_id(document_id);
    }
    if let Some(tab_stop) = number_prop(props, "defaultTabStop", path)? {
        settings = settings.default_tab_stop(to_twips_usize(tab_stop, path)?);
    }
    if let Some(variables) = object_prop(props, "documentVariables", path)? {
        for (name, value) in variables {
            let value = value
                .as_str()
                .ok_or_else(|| validation(path, "`documentVariables` values must be strings"))?;
            settings = settings.add_doc_var(name, value);
        }
    }
    if bool_prop(props, "evenAndOddHeaders", path)? == Some(true) {
        settings = settings.even_and_odd_headers();
    }
    if bool_prop(props, "adjustLineHeightInTable", path)? == Some(true) {
        settings = settings.adjust_line_height_in_table();
    }
    if let Some(value) = optional_enum(
        props,
        "characterSpacingControl",
        &[
            "doNotCompress",
            "compressPunctuation",
            "compressPunctuationAndJapaneseKana",
        ],
        path,
    )? {
        settings = settings.character_spacing_control(match value {
            "compressPunctuation" => CharacterSpacingValues::CompressPunctuation,
            "compressPunctuationAndJapaneseKana" => {
                CharacterSpacingValues::CompressPunctuationAndJapaneseKana
            }
            _ => CharacterSpacingValues::DoNotCompress,
        });
    }
    Ok(settings)
}

fn compile_style(value: &Value, path: &str) -> Result<Style> {
    let definition = value
        .as_object()
        .ok_or_else(|| validation(path, "style definition must be an object"))?;
    let id = required_string(definition, "id", path)?;
    let style_type = match required_string(definition, "type", path)? {
        "paragraph" => StyleType::Paragraph,
        "character" => StyleType::Character,
        "numbering" => StyleType::Numbering,
        "table" => StyleType::Table,
        value => return Err(validation(path, format!("invalid style type `{value}`"))),
    };
    let mut style = Style::new(id, style_type).name(required_string(definition, "name", path)?);
    if let Some(value) = string_prop(definition, "basedOn", path)? {
        style = style.based_on(value);
    }
    if let Some(value) = string_prop(definition, "next", path)? {
        style = style.next(value);
    }
    if let Some(value) = string_prop(definition, "link", path)? {
        style = style.link(value);
    }
    if let Some(value) = bool_prop(definition, "quickFormat", path)? {
        style = style.q_format(value);
    }
    if let Some(value) = definition.get("uiPriority") {
        style = style.ui_priority(value_to_usize(value, path, "uiPriority")?);
    }
    if bool_prop(definition, "semiHidden", path)? == Some(true) {
        style = style.semi_hidden();
    }
    if bool_prop(definition, "unhideWhenUsed", path)? == Some(true) {
        style = style.unhide_when_used();
    }
    if let Some(run) = definition.get("run") {
        style = compile_style_run(style, run, path)?;
    }
    if let Some(paragraph) = definition.get("paragraph") {
        style = compile_style_paragraph(style, paragraph, path)?;
    }
    if let Some(table) = definition.get("table") {
        style = compile_style_table(style, table, path)?;
    }
    if let Some(cell) = definition.get("cell") {
        style = compile_style_cell(style, cell, path)?;
    }
    Ok(style)
}

fn compile_style_run(mut style: Style, value: &Value, path: &str) -> Result<Style> {
    let run = value
        .as_object()
        .ok_or_else(|| validation(path, "style `run` must be an object"))?;
    if let Some(font) = string_prop(run, "font", path)? {
        style = style.fonts(
            RunFonts::new()
                .ascii(font)
                .hi_ansi(font)
                .east_asia(font)
                .cs(font),
        );
    }
    if let Some(fonts) = object_prop(run, "fonts", path)? {
        style = style.fonts(compile_run_fonts(fonts, &format!("{path}/run/fonts"))?);
    }
    if let Some(size) = number_prop(run, "size", path)? {
        style = style.size(to_half_points(size, path)?);
    }
    if let Some(color) = string_prop(run, "color", path)? {
        style = style.color(color.to_ascii_uppercase());
    }
    if let Some(value) = string_prop(run, "themeColor", path)? {
        style = style.theme_color(theme_color(value, path)?);
    }
    if let Some(value) = string_prop(run, "themeShade", path)? {
        style = style.theme_shade(value.to_ascii_uppercase());
    }
    if let Some(value) = string_prop(run, "themeTint", path)? {
        style = style.theme_tint(value.to_ascii_uppercase());
    }
    if let Some(value) = string_prop(run, "highlight", path)? {
        style = style.highlight(value);
    }
    style = compile_style_run_emphasis(style, run, path)?;
    if let Some(value) = run.get("textBorder") {
        let spec = parse_border_spec(value, path, true)?;
        let space = value
            .as_object()
            .and_then(|border| border.get("space"))
            .and_then(Value::as_u64)
            .map(usize::try_from)
            .transpose()
            .map_err(|_| validation(path, "style text border `space` is out of range"))?
            .unwrap_or(0);
        style = style.text_border(
            TextBorder::new()
                .border_type(spec.border_type)
                .size(spec.size)
                .color(spec.color)
                .space(space),
        );
    }
    style = compile_style_run_effects(style, run, path)?;
    Ok(style)
}

fn compile_style_run_effects(
    mut style: Style,
    run: &Map<String, Value>,
    path: &str,
) -> Result<Style> {
    if let Some(value) = run.get("fitText") {
        let fit_text = value
            .as_object()
            .ok_or_else(|| validation(path, "style run `fitText` must be an object"))?;
        let width = fit_text
            .get("width")
            .and_then(Value::as_f64)
            .ok_or_else(|| validation(path, "style run `fitText.width` must be a number"))?;
        let width = to_twips_usize(width, path)?;
        let id = fit_text
            .get("id")
            .and_then(Value::as_u64)
            .map(u32::try_from)
            .transpose()
            .map_err(|_| validation(path, "style run `fitText.id` is out of range"))?;
        style.run_property = std::mem::take(&mut style.run_property).fit_text(width, id);
    }
    if let Some(value) = run.get("stretch") {
        let stretch = value
            .as_i64()
            .and_then(|value| i32::try_from(value).ok())
            .ok_or_else(|| validation(path, "style run `stretch` is out of range"))?;
        style.run_property = std::mem::take(&mut style.run_property).stretch(stretch);
    }
    if run.get("specVanish").and_then(Value::as_bool) == Some(true) {
        style.run_property = std::mem::take(&mut style.run_property).spec_vanish();
    }
    if let Some(value) = run.get("shading") {
        let shading = value
            .as_object()
            .ok_or_else(|| validation(path, "style run `shading` must be an object"))?;
        let fill = string_prop(shading, "fill", path)?
            .ok_or_else(|| validation(path, "style run `shading` requires `fill`"))?;
        let mut property = Shading::new().fill(fill.to_ascii_uppercase());
        if let Some(color) = string_prop(shading, "color", path)? {
            property = property.color(if color == "auto" {
                "auto".to_owned()
            } else {
                color.to_ascii_uppercase()
            });
        }
        if let Some(pattern) = string_prop(shading, "pattern", path)? {
            let pattern = pattern
                .parse::<ShdType>()
                .map_err(|_| validation(path, "style run shading pattern is invalid"))?;
            property = property.shd_type(pattern);
        }
        style.run_property = std::mem::take(&mut style.run_property).shading(property);
    }
    Ok(style)
}

fn compile_style_run_emphasis(
    mut style: Style,
    run: &Map<String, Value>,
    path: &str,
) -> Result<Style> {
    if let Some(value) = bool_prop(run, "bold", path)? {
        style = if value {
            style.bold()
        } else {
            style.run_property = std::mem::take(&mut style.run_property).disable_bold();
            style
        };
    }
    if let Some(value) = bool_prop(run, "italic", path)? {
        style = if value {
            style.italic()
        } else {
            style.run_property = std::mem::take(&mut style.run_property).disable_italic();
            style
        };
    }
    if let Some(value) = bool_prop(run, "strike", path)? {
        style.run_property = if value {
            std::mem::take(&mut style.run_property).strike()
        } else {
            std::mem::take(&mut style.run_property).disable_strike()
        };
    }
    if let Some(value) = bool_prop(run, "doubleStrike", path)? {
        style.run_property = if value {
            std::mem::take(&mut style.run_property).dstrike()
        } else {
            std::mem::take(&mut style.run_property).disable_dstrike()
        };
    }
    if let Some(value) = bool_prop(run, "caps", path)? {
        style.run_property.caps = Some(if value {
            Caps::new()
        } else {
            Caps::new().disable()
        });
    }
    if let Some(value) = string_prop(run, "underline", path)? {
        style = style.underline(value);
    }
    if bool_prop(run, "hidden", path)? == Some(true) {
        style = style.vanish();
    }
    if let Some(value) = string_prop(run, "verticalAlign", path)? {
        let alignment = match value {
            "baseline" => VertAlignType::Baseline,
            "superscript" => VertAlignType::SuperScript,
            "subscript" => VertAlignType::SubScript,
            _ => return Err(validation(path, "invalid `verticalAlign` value")),
        };
        style.run_property = std::mem::take(&mut style.run_property).vert_align(alignment);
    }
    if let Some(spacing) = number_prop(run, "characterSpacing", path)? {
        style.run_property = std::mem::take(&mut style.run_property).spacing(to_twips_i32(
            spacing,
            &format!("{path}/run/characterSpacing"),
        )?);
    }
    Ok(style)
}

fn compile_style_table(mut style: Style, value: &Value, path: &str) -> Result<Style> {
    let table = value
        .as_object()
        .ok_or_else(|| validation(path, "style `table` must be an object"))?;
    if let Some(value) = string_prop(table, "style", path)? {
        style = style.style(value);
    }
    if let Some(value) = number_prop(table, "indent", path)? {
        style = style.table_indent(to_twips_i32(value, path)?);
    }
    if let Some(value) = number_prop(table, "width", path)? {
        style = style.width(to_twips_usize(value, path)?, WidthType::Dxa);
    }
    if let Some(value) = number_prop(table, "widthPercent", path)? {
        style = style.width(percent_to_fiftieths(value, path)?, WidthType::Pct);
    }
    if let Some(value) = string_prop(table, "align", path)? {
        style = style.table_align(match value {
            "center" => TableAlignmentType::Center,
            "right" => TableAlignmentType::Right,
            _ => TableAlignmentType::Left,
        });
    }
    if let Some(value) = string_prop(table, "layout", path)? {
        style = style.layout(if value == "fixed" {
            TableLayoutType::Fixed
        } else {
            TableLayoutType::Autofit
        });
    }
    if let Some(value) = table.get("margins") {
        let [top, right, bottom, left] = parse_box_margins(value, path)?;
        style = style.margins(TableCellMargins::new().margin(top, right, bottom, left));
    }
    if let Some(value) = table.get("border") {
        style = style.set_borders(compile_table_borders(value, path)?);
    }
    Ok(style)
}

fn compile_style_cell(mut style: Style, value: &Value, path: &str) -> Result<Style> {
    let cell = value
        .as_object()
        .ok_or_else(|| validation(path, "style `cell` must be an object"))?;
    let mut property = TableCellProperty::new();
    if let Some(value) = number_prop(cell, "width", path)? {
        property = property.width(to_twips_usize(value, path)?, WidthType::Dxa);
    }
    if let Some(value) = cell.get("colSpan") {
        property = property.grid_span(value_to_usize(value, path, "colSpan")?);
    }
    if let Some(value) = string_prop(cell, "verticalAlign", path)? {
        property = property.vertical_align(match value {
            "center" => VAlignType::Center,
            "bottom" => VAlignType::Bottom,
            _ => VAlignType::Top,
        });
    }
    if let Some(value) = string_prop(cell, "verticalMerge", path)? {
        property = property.vertical_merge(if value == "restart" {
            VMergeType::Restart
        } else {
            VMergeType::Continue
        });
    }
    if let Some(value) = string_prop(cell, "textDirection", path)? {
        let direction = value
            .parse::<TextDirectionType>()
            .map_err(|_| validation(path, format!("invalid text direction `{value}`")))?;
        property = property.text_direction(direction);
    }
    if let Some(value) = string_prop(cell, "shading", path)? {
        property = property.shading(Shading::new().fill(value.to_ascii_uppercase()));
    }
    if let Some(value) = cell.get("margins") {
        let [top, right, bottom, left] = parse_box_margins(value, path)?;
        property = property.margins(
            CellMargins::new()
                .margin_top(top, WidthType::Dxa)
                .margin_right(right, WidthType::Dxa)
                .margin_bottom(bottom, WidthType::Dxa)
                .margin_left(left, WidthType::Dxa),
        );
    }
    if let Some(value) = cell.get("border") {
        property = property.set_borders(compile_cell_borders(value, path)?);
    }
    style = style.table_cell_property(property);
    Ok(style)
}

fn compile_style_paragraph(mut style: Style, value: &Value, path: &str) -> Result<Style> {
    let paragraph = value
        .as_object()
        .ok_or_else(|| validation(path, "style `paragraph` must be an object"))?;
    if let Some(value) = string_prop(paragraph, "align", path)? {
        style = style.align(match value {
            "right" => AlignmentType::Right,
            "center" => AlignmentType::Center,
            "both" => AlignmentType::Both,
            "distribute" => AlignmentType::Distribute,
            "start" => AlignmentType::Start,
            "end" => AlignmentType::End,
            "justified" => AlignmentType::Justified,
            _ => AlignmentType::Left,
        });
    }
    if let Some(value) = string_prop(paragraph, "textAlign", path)? {
        style = style.text_alignment(match value {
            "baseline" => TextAlignmentType::Baseline,
            "bottom" => TextAlignmentType::Bottom,
            "center" => TextAlignmentType::Center,
            "top" => TextAlignmentType::Top,
            _ => TextAlignmentType::Auto,
        });
    }
    if let Some(value) = bool_prop(paragraph, "snapToGrid", path)? {
        style = style.snap_to_grid(value);
    }
    if let Some(value) = bool_prop(paragraph, "pageBreakBefore", path)? {
        style.paragraph_property =
            std::mem::take(&mut style.paragraph_property).page_break_before(value);
    }
    if let Some(value) = bool_prop(paragraph, "widowControl", path)? {
        style.paragraph_property =
            std::mem::take(&mut style.paragraph_property).widow_control(value);
    }
    if let Some(value) = bool_prop(paragraph, "bidi", path)? {
        style.paragraph_property = std::mem::take(&mut style.paragraph_property).bidi(value);
    }
    if let Some(value) = string_prop(paragraph, "shading", path)? {
        style.paragraph_property = std::mem::take(&mut style.paragraph_property)
            .shading(Shading::new().fill(value.to_ascii_uppercase()));
    }
    if let Some(value) = paragraph.get("adjustRightIndent") {
        let value = value
            .as_i64()
            .and_then(|value| isize::try_from(value).ok())
            .ok_or_else(|| validation(path, "`adjustRightIndent` must be 0 or 1"))?;
        style.paragraph_property =
            std::mem::take(&mut style.paragraph_property).adjust_right_ind(value);
    }
    style = compile_style_spacing(style, paragraph, path)?;
    style = compile_style_indent(style, paragraph, path)?;
    for (key, first_line) in [("hangingChars", false), ("firstLineChars", true)] {
        if let Some(value) = paragraph.get(key) {
            let value = value_to_i32(value, path, key)?;
            style = if first_line {
                style.first_line_chars(value)
            } else {
                style.hanging_chars(value)
            };
        }
    }
    if bool_prop(paragraph, "keepNext", path)? == Some(true) {
        style.paragraph_property = std::mem::take(&mut style.paragraph_property).keep_next(true);
    }
    if bool_prop(paragraph, "keepLines", path)? == Some(true) {
        style.paragraph_property = std::mem::take(&mut style.paragraph_property).keep_lines(true);
    }
    if let Some(value) = paragraph.get("outlineLevel") {
        style = style.outline_lvl(value_to_usize(value, path, "outlineLevel")?);
    }
    if let Some(frame) = paragraph.get("frame") {
        style = compile_style_frame(style, frame, path)?;
    }
    Ok(style)
}

fn compile_style_spacing(
    mut style: Style,
    paragraph: &Map<String, Value>,
    path: &str,
) -> Result<Style> {
    let mut spacing = LineSpacing::new();
    let mut present = false;
    if let Some(value) = number_prop(paragraph, "spacingBefore", path)? {
        spacing = spacing.before(to_twips_u32(value, path)?);
        present = true;
    }
    if let Some(value) = number_prop(paragraph, "spacingAfter", path)? {
        spacing = spacing.after(to_twips_u32(value, path)?);
        present = true;
    }
    if let Some(value) = number_prop(paragraph, "lineSpacing", path)? {
        spacing = spacing.line(to_twips_i32(value, path)?);
        present = true;
    }
    let (next, extras_present) = compile_line_spacing_extras(
        spacing,
        paragraph,
        "spacingBeforeLines",
        "spacingAfterLines",
        "lineRule",
        path,
    )?;
    spacing = next;
    present |= extras_present;
    if present {
        style = style.line_spacing(spacing);
    }
    Ok(style)
}

fn compile_default_line_spacing(value: &Value, path: &str) -> Result<LineSpacing> {
    let spacing = value
        .as_object()
        .ok_or_else(|| validation(path, "`defaultLineSpacing` must be an object"))?;
    let mut output = LineSpacing::new();
    if let Some(value) = number_prop(spacing, "before", path)? {
        output = output.before(to_twips_u32(
            value,
            &format!("{path}/defaultLineSpacing/before"),
        )?);
    }
    if let Some(value) = number_prop(spacing, "after", path)? {
        output = output.after(to_twips_u32(
            value,
            &format!("{path}/defaultLineSpacing/after"),
        )?);
    }
    if let Some(value) = number_prop(spacing, "line", path)? {
        output = output.line(to_twips_i32(
            value,
            &format!("{path}/defaultLineSpacing/line"),
        )?);
    }
    compile_line_spacing_extras(
        output,
        spacing,
        "beforeLines",
        "afterLines",
        "lineRule",
        &format!("{path}/defaultLineSpacing"),
    )
    .map(|(spacing, _)| spacing)
}

fn compile_line_spacing_extras(
    mut spacing: LineSpacing,
    props: &Map<String, Value>,
    before_lines: &str,
    after_lines: &str,
    line_rule: &str,
    path: &str,
) -> Result<(LineSpacing, bool)> {
    let mut present = false;
    if let Some(value) = props.get(before_lines) {
        spacing = spacing.before_lines(value_to_u32(value, path, before_lines)?);
        present = true;
    }
    if let Some(value) = props.get(after_lines) {
        spacing = spacing.after_lines(value_to_u32(value, path, after_lines)?);
        present = true;
    }
    if let Some(value) = string_prop(props, line_rule, path)? {
        spacing = spacing.line_rule(match value {
            "atLeast" => LineSpacingType::AtLeast,
            "exact" => LineSpacingType::Exact,
            _ => LineSpacingType::Auto,
        });
        present = true;
    }
    Ok((spacing, present))
}

fn compile_style_indent(
    mut style: Style,
    paragraph: &Map<String, Value>,
    path: &str,
) -> Result<Style> {
    let left = optional_twips_i32(paragraph, "indentLeft", path)?;
    let right = optional_twips_i32(paragraph, "indentRight", path)?;
    let special = if let Some(value) = number_prop(paragraph, "firstLine", path)? {
        Some(SpecialIndentType::FirstLine(to_twips_i32(value, path)?))
    } else if let Some(value) = number_prop(paragraph, "hanging", path)? {
        Some(SpecialIndentType::Hanging(to_twips_i32(value, path)?))
    } else {
        None
    };
    if left.is_some()
        || right.is_some()
        || special.is_some()
        || [
            "indentLeftChars",
            "indentStartChars",
            "indentRightChars",
            "indentEndChars",
            "hangingChars",
            "firstLineChars",
        ]
        .iter()
        .any(|key| paragraph.contains_key(*key))
    {
        style = style.indent(left, special, right, None);
    }
    Ok(style)
}

fn compile_style_frame(mut style: Style, value: &Value, path: &str) -> Result<Style> {
    let frame = value
        .as_object()
        .ok_or_else(|| validation(path, "style paragraph `frame` must be an object"))?;
    if let Some(value) = string_prop(frame, "wrap", path)? {
        style = style.wrap(value);
    }
    if let Some(value) = string_prop(frame, "verticalAnchor", path)? {
        style = style.v_anchor(value);
    }
    if let Some(value) = string_prop(frame, "horizontalAnchor", path)? {
        style = style.h_anchor(value);
    }
    if let Some(value) = string_prop(frame, "heightRule", path)? {
        style = style.h_rule(value);
    }
    if let Some(value) = string_prop(frame, "xAlign", path)? {
        style = style.x_align(value);
    }
    if let Some(value) = string_prop(frame, "yAlign", path)? {
        style = style.y_align(value);
    }
    for (key, apply) in [
        ("horizontalSpace", Style::h_space as fn(Style, i32) -> Style),
        ("verticalSpace", Style::v_space),
        ("x", Style::frame_x),
        ("y", Style::frame_y),
    ] {
        if let Some(value) = number_prop(frame, key, path)? {
            style = apply(style, to_twips_i32(value, path)?);
        }
    }
    if let Some(value) = number_prop(frame, "width", path)? {
        style = style.frame_width(to_twips_u32(value, path)?);
    }
    if let Some(value) = number_prop(frame, "height", path)? {
        style = style.frame_height(to_twips_u32(value, path)?);
    }
    Ok(style)
}

fn value_to_i32(value: &Value, path: &str, key: &str) -> Result<i32> {
    value
        .as_i64()
        .and_then(|value| i32::try_from(value).ok())
        .ok_or_else(|| validation(path, format!("`{key}` is out of range")))
}

fn value_to_u32(value: &Value, path: &str, key: &str) -> Result<u32> {
    value
        .as_u64()
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| validation(path, format!("`{key}` is out of range")))
}

fn value_to_usize(value: &Value, path: &str, key: &str) -> Result<usize> {
    value
        .as_u64()
        .and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| validation(path, format!("`{key}` is out of range")))
}

fn embed_ir_manifest(bytes: Vec<u8>, ir: &IrEnvelope) -> Result<Vec<u8>> {
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        writer
            .raw_copy_file(file)
            .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
    }
    writer
        .start_file(
            "docx-jsx/ir-v1.json",
            zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated),
        )
        .map_err(|error| Error::Compile(format!("cannot write IR manifest: {error}")))?;
    serde_json::to_writer(&mut writer, ir)
        .map_err(|error| Error::Compile(format!("cannot serialize IR manifest: {error}")))?;
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn patch_duplicate_normal_style(bytes: Vec<u8>, props: &Map<String, Value>) -> Result<Vec<u8>> {
    let has_user_normal = props
        .get("styles")
        .and_then(Value::as_array)
        .is_some_and(|styles| {
            styles.iter().any(|style| {
                style
                    .get("id")
                    .and_then(Value::as_str)
                    .is_some_and(|id| id == "Normal")
            })
        });
    if !has_user_normal {
        return Ok(bytes);
    }
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut styles = String::new();
    source
        .by_name("word/styles.xml")
        .map_err(|error| Error::Compile(format!("missing styles: {error}")))?
        .read_to_string(&mut styles)
        .map_err(|error| Error::Compile(format!("cannot read styles: {error}")))?;
    let Some(range) = first_normal_style_range(&styles) else {
        return replace_zip_text_entry(source, "word/styles.xml", styles.as_bytes());
    };
    if styles[range.end..].contains(r#"w:styleId="Normal""#) {
        styles.replace_range(range, "");
    }
    replace_zip_text_entry(source, "word/styles.xml", styles.as_bytes())
}

#[expect(
    clippy::too_many_lines,
    reason = "style indentation and suppression share one package-part rewrite"
)]
fn patch_style_character_indentation(
    bytes: Vec<u8>,
    props: &Map<String, Value>,
) -> Result<Vec<u8>> {
    let style_patches = props
        .get("styles")
        .and_then(Value::as_array)
        .into_iter()
        .flatten()
        .filter_map(|style| {
            let id = style.get("id")?.as_str()?;
            let paragraph = style.get("paragraph")?.as_object()?;
            let attrs = [
                ("indentLeftChars", "leftChars"),
                ("indentStartChars", "startChars"),
                ("indentRightChars", "rightChars"),
                ("indentEndChars", "endChars"),
                ("hangingChars", "hangingChars"),
                ("firstLineChars", "firstLineChars"),
            ]
            .into_iter()
            .filter_map(|(source, target)| {
                paragraph
                    .get(source)
                    .and_then(Value::as_i64)
                    .map(|value| (target, value))
            })
            .collect::<Vec<_>>();
            let suppress_line_numbers = paragraph
                .get("suppressLineNumbers")
                .and_then(Value::as_bool);
            (!attrs.is_empty() || suppress_line_numbers.is_some()).then_some((
                id.to_owned(),
                attrs,
                suppress_line_numbers,
            ))
        })
        .collect::<Vec<_>>();
    if style_patches.is_empty() {
        return Ok(bytes);
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut xml = String::new();
    archive
        .by_name("word/styles.xml")
        .map_err(|error| Error::Compile(format!("missing styles: {error}")))?
        .read_to_string(&mut xml)
        .map_err(|error| Error::Compile(format!("cannot read styles: {error}")))?;
    for (id, properties, suppress_line_numbers) in style_patches {
        let marker = format!(r#"w:styleId="{}""#, xml_escape(&id));
        let marker_start = xml.find(&marker).ok_or_else(|| {
            Error::Compile(format!("docx-rs omitted custom paragraph style `{id}`"))
        })?;
        let style_start = xml[..marker_start]
            .rfind("<w:style")
            .ok_or_else(|| Error::Compile(format!("invalid style start for `{id}`")))?;
        let style_end = xml[marker_start..]
            .find("</w:style>")
            .map(|at| marker_start + at)
            .ok_or_else(|| Error::Compile(format!("invalid style end for `{id}`")))?;
        let mut attrs = String::new();
        for (name, value) in properties {
            write!(attrs, " w:{name}=\"{value}\"").expect("String write");
        }
        let suppress_xml = suppress_line_numbers
            .map(|value| {
                if value {
                    "<w:suppressLineNumbers />"
                } else {
                    "<w:suppressLineNumbers w:val=\"0\" />"
                }
            })
            .unwrap_or_default();
        let segment = &xml[style_start..style_end];
        if let Some(ppr_relative) = segment.find("<w:pPr") {
            let ppr_start = style_start + ppr_relative;
            let ppr_open_end = xml[ppr_start..style_end]
                .find('>')
                .map(|offset| ppr_start + offset + 1)
                .ok_or_else(|| Error::Compile(format!("invalid w:pPr in style `{id}`")))?;
            if xml[ppr_start..ppr_open_end].ends_with("/>") {
                let indentation = if attrs.is_empty() {
                    String::new()
                } else {
                    format!("<w:ind{attrs} />")
                };
                let replacement = format!("<w:pPr>{indentation}{suppress_xml}</w:pPr>");
                xml.replace_range(ppr_start..ppr_open_end, &replacement);
            } else {
                let ppr_end = xml[ppr_open_end..style_end]
                    .find("</w:pPr>")
                    .map(|offset| ppr_open_end + offset)
                    .ok_or_else(|| Error::Compile(format!("invalid w:pPr in style `{id}`")))?;
                if let Some(ind_relative) = xml[ppr_open_end..ppr_end].find("<w:ind") {
                    let ind_start = ppr_open_end + ind_relative;
                    let mut ind_end = xml[ind_start..style_end]
                        .find('>')
                        .map(|at| ind_start + at)
                        .ok_or_else(|| Error::Compile(format!("invalid w:ind in style `{id}`")))?;
                    if xml[ind_end.saturating_sub(1)..=ind_end].ends_with("/>") {
                        ind_end -= 1;
                    }
                    xml.insert_str(ind_end, &attrs);
                    xml.insert_str(ppr_end, suppress_xml);
                } else {
                    let indentation = if attrs.is_empty() {
                        String::new()
                    } else {
                        format!("<w:ind{attrs} />")
                    };
                    xml.insert_str(ppr_end, &format!("{indentation}{suppress_xml}"));
                }
            }
        } else {
            let insertion = ["<w:rPr", "<w:tblPr", "<w:trPr", "<w:tcPr", "</w:style>"]
                .into_iter()
                .filter_map(|tag| segment.find(tag).map(|offset| style_start + offset))
                .min()
                .ok_or_else(|| Error::Compile(format!("invalid style body for `{id}`")))?;
            let indentation = if attrs.is_empty() {
                String::new()
            } else {
                format!("<w:ind{attrs} />")
            };
            xml.insert_str(
                insertion,
                &format!("<w:pPr>{indentation}{suppress_xml}</w:pPr>"),
            );
        }
    }
    replace_zip_text_entry(archive, "word/styles.xml", xml.as_bytes())
}

fn first_normal_style_range(xml: &str) -> Option<std::ops::Range<usize>> {
    let marker = xml.find(r#"w:styleId="Normal""#)?;
    let start = xml[..marker].rfind("<w:style")?;
    let close = xml[marker..].find("</w:style>")?;
    Some(start..marker + close + "</w:style>".len())
}

fn strip_comment_seed_paragraphs(bytes: Vec<u8>, paragraph_ids: &[String]) -> Result<Vec<u8>> {
    if paragraph_ids.is_empty() {
        return Ok(bytes);
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut xml = String::new();
    archive
        .by_name("word/document.xml")
        .map_err(|error| Error::Compile(format!("missing word/document.xml: {error}")))?
        .read_to_string(&mut xml)
        .map_err(|error| Error::Compile(format!("cannot read word/document.xml: {error}")))?;
    for id in paragraph_ids {
        let marker = format!(r#"<w:p w14:paraId="{id}""#);
        let Some(start) = xml.find(&marker) else {
            return Err(Error::Compile(format!(
                "temporary comment dependency paragraph `{id}` was not found"
            )));
        };
        let end = xml[start..]
            .find("</w:p>")
            .map(|offset| start + offset + "</w:p>".len())
            .ok_or_else(|| {
                Error::Compile(format!(
                    "temporary comment dependency paragraph `{id}` is unterminated"
                ))
            })?;
        xml.replace_range(start..end, "");
    }
    replace_zip_text_entry(archive, "word/document.xml", xml.as_bytes())
}

fn replace_zip_text_entry(
    mut source: zip::ZipArchive<Cursor<Vec<u8>>>,
    name: &str,
    contents: &[u8],
) -> Result<Vec<u8>> {
    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        if file.name() == name {
            let options = file.options();
            writer
                .start_file(file.name(), options)
                .map_err(|error| Error::Compile(format!("cannot write {name}: {error}")))?;
            writer
                .write_all(contents)
                .map_err(|error| Error::Compile(format!("cannot write {name}: {error}")))?;
        } else {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
        }
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn patch_numbering_style_links(
    bytes: Vec<u8>,
    patches: &[(usize, Option<String>, Option<String>)],
) -> Result<Vec<u8>> {
    if patches.is_empty() {
        return Ok(bytes);
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut xml = String::new();
    archive
        .by_name("word/numbering.xml")
        .map_err(|error| Error::Compile(format!("missing word/numbering.xml: {error}")))?
        .read_to_string(&mut xml)
        .map_err(|error| Error::Compile(format!("cannot read word/numbering.xml: {error}")))?;
    for (id, style_link, num_style_link) in patches {
        let marker = format!(r#"<w:abstractNum w:abstractNumId="{id}""#);
        let start = xml
            .find(&marker)
            .ok_or_else(|| Error::Compile(format!("docx-rs omitted abstract numbering `{id}`")))?;
        let tag_end = xml[start..]
            .find('>')
            .map(|offset| start + offset + 1)
            .ok_or_else(|| Error::Compile(format!("invalid abstract numbering `{id}`")))?;
        let self_closing = xml[start..tag_end].trim_end().ends_with("/>");
        let close = if self_closing {
            tag_end
        } else {
            xml[start..]
                .find("</w:abstractNum>")
                .map(|offset| start + offset)
                .ok_or_else(|| Error::Compile(format!("invalid abstract numbering `{id}`")))?
        };
        let mut links = String::new();
        if let Some(value) = style_link {
            let _ = write!(links, r#"<w:styleLink w:val="{}" />"#, xml_escape(value));
        }
        if let Some(value) = num_style_link {
            let _ = write!(links, r#"<w:numStyleLink w:val="{}" />"#, xml_escape(value));
        }
        if links.is_empty() {
            continue;
        }
        if self_closing {
            let opening = xml[start..tag_end].trim_end().trim_end_matches("/>");
            xml.replace_range(
                start..tag_end,
                &format!("{opening}>{links}</w:abstractNum>"),
            );
        } else {
            let insertion = xml[start..close]
                .find("<w:lvl ")
                .map_or(close, |offset| start + offset);
            xml.insert_str(insertion, &links);
        }
    }
    replace_zip_text_entry(archive, "word/numbering.xml", xml.as_bytes())
}

fn omit_unused_default_parts(bytes: Vec<u8>, root: &Node) -> Result<Vec<u8>> {
    let keep_comments = node_contains_kind(root, NodeKind::Comment);
    let keep_footnotes = node_contains_kind(root, NodeKind::Footnote);
    let keep_numbering = node_contains_kind(root, NodeKind::List);
    if keep_comments && keep_footnotes && keep_numbering {
        return Ok(bytes);
    }
    let mut omit = Vec::new();
    let mut rel_needles = Vec::new();
    let mut type_needles = Vec::new();
    if !keep_comments {
        omit.extend(["word/comments.xml", "word/commentsExtended.xml"]);
        rel_needles.extend(["Target=\"comments.xml\"", "Target=\"commentsExtended.xml\""]);
        type_needles.extend([
            "PartName=\"/word/comments.xml\"",
            "PartName=\"/word/commentsExtended.xml\"",
        ]);
    }
    if !keep_footnotes {
        omit.push("word/footnotes.xml");
        rel_needles.push("Target=\"footnotes.xml\"");
        type_needles.push("PartName=\"/word/footnotes.xml\"");
    }
    if !keep_numbering {
        omit.push("word/numbering.xml");
        rel_needles.push("Target=\"numbering.xml\"");
        type_needles.push("PartName=\"/word/numbering.xml\"");
    }
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let mut file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        let name = file.name().to_owned();
        if omit.iter().any(|part| *part == name) {
            continue;
        }
        let options = file.options();
        let mut data = Vec::new();
        file.read_to_end(&mut data)
            .map_err(|error| Error::Compile(format!("cannot read {name}: {error}")))?;
        if name == "word/_rels/document.xml.rels" {
            let mut rels = String::from_utf8(data).map_err(|error| {
                Error::Compile(format!("document relationships are not UTF-8: {error}"))
            })?;
            rels = remove_markup_containing(&rels, &rel_needles);
            data = rels.into_bytes();
        } else if name == "[Content_Types].xml" {
            let mut types = String::from_utf8(data)
                .map_err(|error| Error::Compile(format!("content types are not UTF-8: {error}")))?;
            types = remove_markup_containing(&types, &type_needles);
            data = types.into_bytes();
        }
        writer
            .start_file(&name, options)
            .map_err(|error| Error::Compile(format!("cannot write {name}: {error}")))?;
        writer
            .write_all(&data)
            .map_err(|error| Error::Compile(format!("cannot write {name}: {error}")))?;
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn node_contains_kind(node: &Node, kind: NodeKind) -> bool {
    node.kind == kind
        || node.children.iter().any(|child| match child {
            Child::Node(inner) => node_contains_kind(inner, kind),
            Child::String(_) | Child::Number(_) => false,
        })
}

fn remove_markup_containing(xml: &str, needles: &[&str]) -> String {
    let mut output = xml.to_owned();
    for needle in needles {
        while let Some(hit) = output.find(needle) {
            let Some(start) = output[..hit].rfind('<') else {
                break;
            };
            let Some(rel) = output[hit..].find('>') else {
                break;
            };
            output.replace_range(start..=hit + rel, "");
        }
    }
    output
}

fn patch_custom_xml_content_types(bytes: Vec<u8>, item_count: usize) -> Result<Vec<u8>> {
    if item_count == 0 {
        return Ok(bytes);
    }
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut content_types = String::new();
    source
        .by_name("[Content_Types].xml")
        .map_err(|error| Error::Compile(format!("missing content types: {error}")))?
        .read_to_string(&mut content_types)
        .map_err(|error| Error::Compile(format!("cannot read content types: {error}")))?;
    let marker = r#"<Override PartName="/customXml/itemProps"#;
    while let Some(start) = content_types.find(marker) {
        let end = content_types[start..]
            .find("/>")
            .map(|offset| start + offset + 2)
            .ok_or_else(|| Error::Compile("invalid custom XML content type override".to_owned()))?;
        content_types.replace_range(start..end, "");
    }
    let end = content_types
        .rfind("</Types>")
        .ok_or_else(|| Error::Compile("invalid [Content_Types].xml".to_owned()))?;
    let mut overrides = String::new();
    for index in 1..=item_count {
        write!(overrides, r#"<Override PartName="/customXml/itemProps{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml" />"#)
            .expect("writing to String cannot fail");
    }
    content_types.insert_str(end, &overrides);

    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        if file.name() == "[Content_Types].xml" {
            let options = file.options();
            writer
                .start_file(file.name(), options)
                .map_err(|error| Error::Compile(format!("cannot write content types: {error}")))?;
            writer
                .write_all(content_types.as_bytes())
                .map_err(|error| Error::Compile(format!("cannot write content types: {error}")))?;
        } else {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
        }
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn patch_custom_xml_schema_refs(bytes: Vec<u8>, props: &Map<String, Value>) -> Result<Vec<u8>> {
    let Some(items) = props.get("customXmlItems").and_then(Value::as_array) else {
        return Ok(bytes);
    };
    let patches = items
        .iter()
        .enumerate()
        .filter_map(|(index, value)| {
            let refs = value.get("schemaRefs")?.as_array()?;
            let uris = refs
                .iter()
                .filter_map(Value::as_str)
                .filter(|uri| !uri.is_empty())
                .collect::<Vec<_>>();
            (!uris.is_empty()).then_some((index + 1, uris))
        })
        .collect::<Vec<_>>();
    if patches.is_empty() {
        return Ok(bytes);
    }
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let replacements = patches
        .into_iter()
        .map(|(number, uris)| -> Result<(String, String)> {
            let name = format!("customXml/itemProps{number}.xml");
            let mut xml = String::new();
            source
                .by_name(&name)
                .map_err(|error| Error::Compile(format!("missing {name}: {error}")))?
                .read_to_string(&mut xml)
                .map_err(|error| Error::Compile(format!("cannot read {name}: {error}")))?;
            let mut refs = String::new();
            for uri in uris {
                write!(refs, r#"<ds:schemaRef ds:uri="{uri}"/>"#)
                    .expect("writing to String cannot fail");
            }
            let block = format!("<ds:schemaRefs>{refs}</ds:schemaRefs>");
            let xml = if xml.contains("<ds:schemaRefs></ds:schemaRefs>") {
                xml.replacen("<ds:schemaRefs></ds:schemaRefs>", &block, 1)
            } else {
                xml.replacen("<ds:schemaRefs/>", &block, 1)
            };
            if !xml.contains("<ds:schemaRef") {
                return Err(Error::Compile(format!(
                    "cannot insert schema refs into {name}"
                )));
            }
            Ok((name, xml))
        })
        .collect::<Result<Vec<_>>>()?;
    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        let name = file.name().to_owned();
        if let Some((_, xml)) = replacements.iter().find(|(part, _)| *part == name) {
            let options = file.options();
            writer.start_file(name, options).map_err(|error| {
                Error::Compile(format!("cannot write custom XML props: {error}"))
            })?;
            writer.write_all(xml.as_bytes()).map_err(|error| {
                Error::Compile(format!("cannot write custom XML props: {error}"))
            })?;
        } else {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
        }
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn normalize_ooxml_element_order(bytes: Vec<u8>) -> Result<Vec<u8>> {
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let mut file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        if file.name().starts_with("word/")
            && file.name() != "word/comments.xml"
            && file.name() != "word/commentsExtended.xml"
            && Path::new(file.name())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        {
            let name = file.name().to_owned();
            let options = file.options();
            let mut xml = String::new();
            file.read_to_string(&mut xml)
                .map_err(|error| Error::Compile(format!("cannot read {name}: {error}")))?;
            normalize_word_xml(&mut xml, &name)
                .map_err(|error| Error::Compile(format!("{name}: {error}")))?;
            writer
                .start_file(&name, options)
                .map_err(|error| Error::Compile(format!("cannot write {name}: {error}")))?;
            writer
                .write_all(xml.as_bytes())
                .map_err(|error| Error::Compile(format!("cannot write {name}: {error}")))?;
        } else {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
        }
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn normalize_word_xml(xml: &mut String, name: &str) -> Result<()> {
    remove_table_grid_column_type(xml);
    if name == "word/styles.xml" {
        remove_element_children(xml, "w:pPr", "w:rPr")?;
        remove_element_children(xml, "w:tblPr", "w:tblW")?;
        remove_element_children(xml, "w:tblPr", "w:tblLayout")?;
        prefix_element_children(
            xml,
            "w:tblPr",
            r#"<w:tblStyleRowBandSize w:val="1" /><w:tblStyleColBandSize w:val="1" />"#,
        )?;
    }
    if name == "word/numbering.xml" {
        remove_element_children(xml, "w:pPr", "w:rPr")?;
        reorder_element_children(xml, "w:lvl", numbering_level_rank)?;
        reorder_element_children(xml, "w:lvlOverride", numbering_override_rank)?;
    }
    reorder_element_children(xml, "w:numPr", numbering_property_rank)?;
    reorder_element_children(xml, "w:pPr", paragraph_property_rank)?;
    reorder_element_children(xml, "w:rPr", run_property_rank)?;
    reorder_element_children(xml, "w:tblPr", table_property_rank)?;
    reorder_element_children(xml, "w:tcPr", table_cell_property_rank)?;
    reorder_element_children(xml, "w:sectPr", section_property_rank)?;
    reorder_element_children(xml, "w:style", style_child_rank)?;
    reorder_element_children(xml, "w:settings", settings_child_rank)
}

fn remove_table_grid_column_type(xml: &mut String) {
    let element_prefix = "<w:gridCol";
    let attribute_name = "w:type=";
    let mut search_from = 0;
    while let Some(relative_start) = xml[search_from..].find(element_prefix) {
        let start = search_from + relative_start;
        let after_element_name = start + element_prefix.len();
        if xml
            .as_bytes()
            .get(after_element_name)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && *byte != b'>' && *byte != b'/')
        {
            search_from = after_element_name;
            continue;
        }
        let Some(relative_end) = xml[start..].find('>') else {
            break;
        };
        let end = start + relative_end + 1;
        let tag = &xml[start..end];
        let Some(type_attribute_start) = tag.find(attribute_name) else {
            search_from = end;
            continue;
        };
        let remove_start = tag[..type_attribute_start]
            .char_indices()
            .next_back()
            .filter(|(_, ch)| ch.is_whitespace())
            .map_or(type_attribute_start, |(index, _)| index);
        let value_start = type_attribute_start + attribute_name.len();
        let Some(quote) = tag.as_bytes().get(value_start).copied() else {
            search_from = end;
            continue;
        };
        if quote != b'\'' && quote != b'"' {
            search_from = end;
            continue;
        }
        let value_end = tag[value_start + 1..]
            .find(char::from(quote))
            .map(|relative| value_start + relative + 2);
        let Some(value_end) = value_end else {
            search_from = end;
            continue;
        };
        let mut normalized = String::with_capacity(tag.len());
        normalized.push_str(&tag[..remove_start]);
        normalized.push_str(&tag[value_end..]);
        xml.replace_range(start..end, &normalized);
        search_from = start + normalized.len();
    }
}

fn numbering_level_rank(fragment: &str) -> usize {
    rank_in(
        child_name(fragment),
        &[
            "w:start",
            "w:numFmt",
            "w:lvlRestart",
            "w:pStyle",
            "w:isLgl",
            "w:suff",
            "w:lvlText",
            "w:lvlPicBulletId",
            "w:legacy",
            "w:lvlJc",
            "w:pPr",
            "w:rPr",
        ],
    )
}

fn numbering_property_rank(fragment: &str) -> usize {
    rank_in(child_name(fragment), &["w:ilvl", "w:numId"])
}

fn numbering_override_rank(fragment: &str) -> usize {
    rank_in(child_name(fragment), &["w:startOverride", "w:lvl"])
}

fn section_property_rank(fragment: &str) -> usize {
    rank_in(
        child_name(fragment),
        &[
            "w:headerReference",
            "w:footerReference",
            "w:footnotePr",
            "w:endnotePr",
            "w:type",
            "w:pgSz",
            "w:pgMar",
            "w:paperSrc",
            "w:pgBorders",
            "w:lnNumType",
            "w:pgNumType",
            "w:cols",
            "w:formProt",
            "w:vAlign",
            "w:noEndnote",
            "w:titlePg",
            "w:textDirection",
            "w:bidi",
            "w:rtlGutter",
            "w:docGrid",
            "w:printerSettings",
            "w:sectPrChange",
        ],
    )
}

fn prefix_element_children(xml: &mut String, parent: &str, prefix: &str) -> Result<()> {
    let ranges = element_content_ranges(xml, parent)?;
    for (start, _) in ranges.into_iter().rev() {
        xml.insert_str(start, prefix);
    }
    Ok(())
}

fn remove_element_children(
    xml: &mut String,
    parent: &str,
    child_name_to_remove: &str,
) -> Result<()> {
    let ranges = element_content_ranges(xml, parent)?;
    for (start, end) in ranges.into_iter().rev() {
        let children = split_element_children(&xml[start..end])?;
        let content = children
            .into_iter()
            .filter(|child| child_name(child) != child_name_to_remove)
            .collect::<String>();
        xml.replace_range(start..end, &content);
    }
    Ok(())
}

fn reorder_element_children(xml: &mut String, tag: &str, rank: fn(&str) -> usize) -> Result<()> {
    let ranges = element_content_ranges(xml, tag)?;
    // Ranges are recorded when their closing tag is seen, so nested elements
    // precede their parents. Normalize inner elements first; sorting preserves
    // byte length and therefore keeps the remaining offsets stable.
    for (start, end) in ranges {
        let content = &xml[start..end];
        let mut children = split_element_children(content).map_err(|error| {
            Error::Compile(format!(
                "cannot normalize `<{tag}>` children `{content}`: {error}"
            ))
        })?;
        children.sort_by_key(|child| rank(child));
        xml.replace_range(start..end, &children.concat());
    }
    Ok(())
}

fn element_content_ranges(xml: &str, wanted: &str) -> Result<Vec<(usize, usize)>> {
    let mut ranges = Vec::new();
    let mut stack = Vec::new();
    let mut cursor = 0;
    while let Some(relative_start) = xml[cursor..].find('<') {
        let start = cursor + relative_start;
        let end = xml[start..]
            .find('>')
            .map(|offset| start + offset)
            .ok_or_else(|| Error::Compile("unterminated XML tag".to_owned()))?;
        let token = xml[start + 1..end].trim();
        let closing = token.strip_prefix('/');
        let body = closing.unwrap_or(token);
        let name = body
            .split(|character: char| character.is_ascii_whitespace() || character == '/')
            .next()
            .unwrap_or_default();
        if name == wanted {
            if closing.is_some() {
                let content_start = stack.pop().ok_or_else(|| {
                    Error::Compile(format!("unexpected closing `<{wanted}>` element"))
                })?;
                ranges.push((content_start, start));
            } else if !token.ends_with('/') {
                stack.push(end + 1);
            }
        }
        cursor = end + 1;
    }
    if stack.is_empty() {
        Ok(ranges)
    } else {
        Err(Error::Compile(format!("unclosed `<{wanted}>` element")))
    }
}

fn split_element_children(content: &str) -> Result<Vec<&str>> {
    let mut children = Vec::new();
    let mut depth = 0_usize;
    let mut child_start = None;
    let mut cursor = 0;
    while let Some(relative_start) = content[cursor..].find('<') {
        let start = cursor + relative_start;
        let end = content[start..]
            .find('>')
            .map(|offset| start + offset)
            .ok_or_else(|| Error::Compile("unterminated XML child tag".to_owned()))?;
        let token = content[start + 1..end].trim();
        if token.starts_with('!') || token.starts_with('?') {
            cursor = end + 1;
            continue;
        }
        if token.starts_with('/') {
            depth = depth
                .checked_sub(1)
                .ok_or_else(|| Error::Compile("unexpected XML child closing tag".to_owned()))?;
            if depth == 0 {
                let child_start = child_start
                    .take()
                    .ok_or_else(|| Error::Compile("missing XML child opening tag".to_owned()))?;
                children.push(&content[child_start..=end]);
            }
        } else if token.ends_with('/') {
            if depth == 0 {
                children.push(&content[start..=end]);
            }
        } else {
            if depth == 0 {
                child_start = Some(start);
            }
            depth += 1;
        }
        cursor = end + 1;
    }
    if depth == 0 && children.concat().len() == content.len() {
        Ok(children)
    } else {
        Err(Error::Compile(
            "property XML contains unsupported text or malformed children".to_owned(),
        ))
    }
}

fn child_name(fragment: &str) -> &str {
    fragment
        .trim_start_matches('<')
        .split(|character: char| {
            character.is_ascii_whitespace() || character == '/' || character == '>'
        })
        .next()
        .unwrap_or_default()
}

fn rank_in(name: &str, ordered: &[&str]) -> usize {
    ordered
        .iter()
        .position(|candidate| *candidate == name)
        .unwrap_or(ordered.len())
}

fn paragraph_property_rank(fragment: &str) -> usize {
    rank_in(
        child_name(fragment),
        &[
            "w:pStyle",
            "w:keepNext",
            "w:keepLines",
            "w:pageBreakBefore",
            "w:framePr",
            "w:widowControl",
            "w:numPr",
            "w:suppressLineNumbers",
            "w:pBdr",
            "w:shd",
            "w:tabs",
            "w:suppressAutoHyphens",
            "w:kinsoku",
            "w:wordWrap",
            "w:overflowPunct",
            "w:topLinePunct",
            "w:autoSpaceDE",
            "w:autoSpaceDN",
            "w:bidi",
            "w:adjustRightInd",
            "w:snapToGrid",
            "w:spacing",
            "w:ind",
            "w:contextualSpacing",
            "w:mirrorIndents",
            "w:suppressOverlap",
            "w:jc",
            "w:textDirection",
            "w:textAlignment",
            "w:textboxTightWrap",
            "w:outlineLvl",
            "w:divId",
            "w:cnfStyle",
            "w:rPr",
            "w:sectPr",
            "w:pPrChange",
        ],
    )
}

fn run_property_rank(fragment: &str) -> usize {
    rank_in(
        child_name(fragment),
        &[
            "w:rStyle",
            "w:rFonts",
            "w:b",
            "w:bCs",
            "w:i",
            "w:iCs",
            "w:caps",
            "w:smallCaps",
            "w:strike",
            "w:dstrike",
            "w:outline",
            "w:shadow",
            "w:emboss",
            "w:imprint",
            "w:noProof",
            "w:snapToGrid",
            "w:vanish",
            "w:webHidden",
            "w:color",
            "w:spacing",
            "w:w",
            "w:kern",
            "w:position",
            "w:sz",
            "w:szCs",
            "w:highlight",
            "w:u",
            "w:effect",
            "w:bdr",
            "w:shd",
            "w:fitText",
            "w:vertAlign",
            "w:rtl",
            "w:cs",
            "w:em",
            "w:lang",
            "w:eastAsianLayout",
            "w:specVanish",
            "w:oMath",
            "w:rPrChange",
        ],
    )
}

fn style_child_rank(fragment: &str) -> usize {
    rank_in(
        child_name(fragment),
        &[
            "w:name",
            "w:aliases",
            "w:basedOn",
            "w:next",
            "w:link",
            "w:autoRedefine",
            "w:hidden",
            "w:uiPriority",
            "w:semiHidden",
            "w:unhideWhenUsed",
            "w:qFormat",
            "w:locked",
            "w:personal",
            "w:personalCompose",
            "w:personalReply",
            "w:rsid",
            "w:pPr",
            "w:rPr",
            "w:tblPr",
            "w:trPr",
            "w:tcPr",
            "w:tblStylePr",
        ],
    )
}

fn table_property_rank(fragment: &str) -> usize {
    rank_in(
        child_name(fragment),
        &[
            "w:tblStyle",
            "w:tblpPr",
            "w:tblOverlap",
            "w:bidiVisual",
            "w:tblStyleRowBandSize",
            "w:tblStyleColBandSize",
            "w:tblW",
            "w:jc",
            "w:tblCellSpacing",
            "w:tblInd",
            "w:tblBorders",
            "w:shd",
            "w:tblLayout",
            "w:tblCellMar",
            "w:tblLook",
            "w:tblCaption",
            "w:tblDescription",
            "w:tblPrChange",
        ],
    )
}

fn table_cell_property_rank(fragment: &str) -> usize {
    rank_in(
        child_name(fragment),
        &[
            "w:cnfStyle",
            "w:tcW",
            "w:gridSpan",
            "w:hMerge",
            "w:vMerge",
            "w:tcBorders",
            "w:shd",
            "w:noWrap",
            "w:tcMar",
            "w:textDirection",
            "w:tcFitText",
            "w:vAlign",
            "w:hideMark",
            "w:headers",
            "w:cellIns",
            "w:cellDel",
            "w:cellMerge",
            "w:tcPrChange",
        ],
    )
}

const SETTINGS_CHILD_ORDER: &[&str] = &[
    "w:writeProtection",
    "w:view",
    "w:zoom",
    "w:removePersonalInformation",
    "w:removeDateAndTime",
    "w:doNotDisplayPageBoundaries",
    "w:displayBackgroundShape",
    "w:printPostScriptOverText",
    "w:printFractionalCharacterWidth",
    "w:printFormsData",
    "w:embedTrueTypeFonts",
    "w:embedSystemFonts",
    "w:saveSubsetFonts",
    "w:saveFormsData",
    "w:mirrorMargins",
    "w:alignBordersAndEdges",
    "w:bordersDoNotSurroundHeader",
    "w:bordersDoNotSurroundFooter",
    "w:gutterAtTop",
    "w:hideSpellingErrors",
    "w:hideGrammaticalErrors",
    "w:activeWritingStyle",
    "w:proofState",
    "w:formsDesign",
    "w:attachedTemplate",
    "w:linkStyles",
    "w:stylePaneFormatFilter",
    "w:stylePaneSortMethod",
    "w:documentType",
    "w:mailMerge",
    "w:revisionView",
    "w:trackRevisions",
    "w:doNotTrackMoves",
    "w:doNotTrackFormatting",
    "w:documentProtection",
    "w:autoFormatOverride",
    "w:styleLockTheme",
    "w:styleLockQFSet",
    "w:defaultTabStop",
    "w:autoHyphenation",
    "w:consecutiveHyphenLimit",
    "w:hyphenationZone",
    "w:doNotHyphenateCaps",
    "w:showEnvelope",
    "w:summaryLength",
    "w:clickAndTypeStyle",
    "w:defaultTableStyle",
    "w:evenAndOddHeaders",
    "w:bookFoldRevPrinting",
    "w:bookFoldPrinting",
    "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing",
    "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery",
    "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin",
    "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin",
    "w:doNotShadeFormData",
    "w:noPunctuationKerning",
    "w:characterSpacingControl",
    "w:printTwoOnOne",
    "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter",
    "w:noLineBreaksBefore",
    "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema",
    "w:saveInvalidXml",
    "w:ignoreMixedContent",
    "w:alwaysShowPlaceholderText",
    "w:doNotDemarcateInvalidXml",
    "w:saveXmlDataOnly",
    "w:useXSLTWhenSaving",
    "w:saveThroughXslt",
    "w:showXMLTags",
    "w:alwaysMergeEmptyNamespace",
    "w:updateFields",
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:docVars",
    "w:rsids",
    "m:mathPr",
    "w:uiCompat97To2003",
    "w:attachedSchema",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats",
    "w:doNotAutoCompressPictures",
    "w:forceUpgrade",
    "w:captions",
    "w:readModeInkLockDown",
    "w:smartTagType",
    "w:schemaLibrary",
    "w:shapeDefaults",
    "w:doNotEmbedSmartTags",
    "w:decimalSymbol",
    "w:listSeparator",
];

fn settings_child_rank(fragment: &str) -> usize {
    rank_in(child_name(fragment), SETTINGS_CHILD_ORDER)
}

fn patch_external_hyperlink_relationships(bytes: Vec<u8>, root: &Node) -> Result<Vec<u8>> {
    let mut targets = Vec::new();
    collect_external_targets(root, false, &mut targets);
    let mut header_targets = Vec::new();
    let mut footer_targets = Vec::new();
    collect_header_footer_targets(root, &mut header_targets, &mut footer_targets);
    let mut comments = Vec::new();
    collect_comments(root, &mut comments);
    if targets.is_empty()
        && header_targets.iter().all(Vec::is_empty)
        && footer_targets.iter().all(Vec::is_empty)
        && comments.is_empty()
    {
        return Ok(bytes);
    }
    let mut source = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Compile(format!("cannot reopen DOCX archive: {error}")))?;
    let mut replacements = std::collections::HashMap::new();
    if !targets.is_empty() {
        let relationships = patched_relationship_part(
            &mut source,
            "word/document.xml",
            "word/_rels/document.xml.rels",
            &targets,
        )?;
        replacements.insert("word/_rels/document.xml.rels".to_owned(), relationships);
    }
    for (prefix, groups) in [("header", header_targets), ("footer", footer_targets)] {
        for (index, targets) in groups.iter().enumerate() {
            if targets.is_empty() {
                continue;
            }
            let part = format!("word/{prefix}{}.xml", index + 1);
            let rel = format!("word/_rels/{prefix}{}.xml.rels", index + 1);
            let relationships = patched_relationship_part(&mut source, &part, &rel, targets)?;
            replacements.insert(rel, relationships);
        }
    }
    patch_comment_parts(&mut source, &mut replacements, &comments)?;

    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        if let Some(replacement) = replacements.get(file.name()) {
            let options = file.options();
            writer
                .start_file(file.name(), options)
                .map_err(|error| Error::Compile(format!("cannot write DOCX entry: {error}")))?;
            writer
                .write_all(replacement.as_bytes())
                .map_err(|error| Error::Compile(format!("cannot write relationships: {error}")))?;
        } else {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
        }
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize DOCX archive: {error}")))
}

fn patch_comment_parts(
    source: &mut zip::ZipArchive<Cursor<Vec<u8>>>,
    replacements: &mut std::collections::HashMap<String, String>,
    comments: &[&Node],
) -> Result<()> {
    if comments.is_empty() {
        return Ok(());
    }
    let mut comments_xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    );
    comments_xml.push('\n');
    for (index, node) in comments.iter().enumerate() {
        append_comment_xml(&mut comments_xml, index + 1, node)?;
    }
    comments_xml.push_str("</w:comments>");
    replacements.insert("word/comments.xml".to_owned(), comments_xml);

    let mut relationships = replacements
        .get("word/_rels/document.xml.rels")
        .cloned()
        .unwrap_or_default();
    if relationships.is_empty() {
        source
            .by_name("word/_rels/document.xml.rels")
            .map_err(|error| Error::Compile(format!("missing document relationships: {error}")))?
            .read_to_string(&mut relationships)
            .map_err(|error| Error::Compile(format!("cannot read relationships: {error}")))?;
    }
    if !relationships.contains(r#"Target="comments.xml""#) {
        let relationship = r#"<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml" />"#;
        relationships = relationships.replacen(
            "</Relationships>",
            &format!("{relationship}</Relationships>"),
            1,
        );
    }
    replacements.insert("word/_rels/document.xml.rels".to_owned(), relationships);
    Ok(())
}

fn append_comment_xml(output: &mut String, id: usize, node: &Node) -> Result<()> {
    let text = required_string(&node.props, "text", "Comment")?;
    let author = string_prop(&node.props, "author", "Comment")?.unwrap_or("unnamed");
    let date = string_prop(&node.props, "date", "Comment")?.unwrap_or("1970-01-01T00:00:00Z");
    write!(output, r#"<w:comment w:id="{id}" w:author="{}" w:date="{}"><w:p><w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p></w:comment>"#, xml_escape(author), xml_escape(date), xml_escape(text))
        .expect("writing to String cannot fail");
    Ok(())
}

fn collect_comments<'a>(node: &'a Node, comments: &mut Vec<&'a Node>) {
    if node.kind == NodeKind::Comment {
        comments.push(node);
    }
    if node.kind == NodeKind::Section {
        for header_footer in [false, true] {
            for child in &node.children {
                if let Child::Node(child) = child
                    && matches!(child.kind, NodeKind::Header | NodeKind::Footer) == header_footer
                {
                    collect_comments(child, comments);
                }
            }
        }
        return;
    }
    for child in &node.children {
        if let Child::Node(child) = child {
            collect_comments(child, comments);
        }
    }
}

fn patched_relationship_part(
    source: &mut zip::ZipArchive<Cursor<Vec<u8>>>,
    part_name: &str,
    relationship_name: &str,
    targets: &[&str],
) -> Result<String> {
    let mut part = String::new();
    source
        .by_name(part_name)
        .map_err(|error| Error::Compile(format!("missing {part_name}: {error}")))?
        .read_to_string(&mut part)
        .map_err(|error| Error::Compile(format!("cannot read {part_name}: {error}")))?;
    let ids = hyperlink_relationship_ids(&part);
    if ids.len() != targets.len() {
        return Err(Error::Compile(format!(
            "unexpected external hyperlink count in {part_name}"
        )));
    }
    let mut relationships = String::new();
    source
        .by_name(relationship_name)
        .map_err(|error| Error::Compile(format!("missing {relationship_name}: {error}")))?
        .read_to_string(&mut relationships)
        .map_err(|error| Error::Compile(format!("cannot read {relationship_name}: {error}")))?;
    let insertion = ids
        .iter()
        .zip(targets)
        .fold(String::new(), |mut output, (id, target)| {
            if relationships.contains(&format!(r#"Id="{id}""#)) {
                return output;
            }
            write!(output, r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{}" TargetMode="External" />"#, xml_escape(id), xml_escape(target))
                .expect("writing to String cannot fail");
            output
        });
    Ok(relationships.replacen(
        "</Relationships>",
        &format!("{insertion}</Relationships>"),
        1,
    ))
}

fn collect_header_footer_targets<'a>(
    node: &'a Node,
    headers: &mut Vec<Vec<&'a str>>,
    footers: &mut Vec<Vec<&'a str>>,
) {
    if matches!(node.kind, NodeKind::Header | NodeKind::Footer) {
        let mut targets = Vec::new();
        collect_all_external_targets(node, &mut targets);
        if node.kind == NodeKind::Header {
            headers.push(targets);
        } else {
            footers.push(targets);
        }
        return;
    }
    for child in &node.children {
        if let Child::Node(child) = child {
            collect_header_footer_targets(child, headers, footers);
        }
    }
}

fn collect_all_external_targets<'a>(node: &'a Node, targets: &mut Vec<&'a str>) {
    if node.kind == NodeKind::Hyperlink
        && let Some(target) = node.props.get("href").and_then(Value::as_str)
    {
        targets.push(target);
    }
    for child in &node.children {
        if let Child::Node(child) = child {
            collect_all_external_targets(child, targets);
        }
    }
}

fn collect_external_targets<'a>(
    node: &'a Node,
    in_header_footer: bool,
    targets: &mut Vec<&'a str>,
) {
    let in_header_footer =
        in_header_footer || matches!(node.kind, NodeKind::Header | NodeKind::Footer);
    if node.kind == NodeKind::Hyperlink
        && !in_header_footer
        && let Some(target) = node.props.get("href").and_then(Value::as_str)
    {
        targets.push(target);
    }
    for child in &node.children {
        if let Child::Node(child) = child {
            collect_external_targets(child, in_header_footer, targets);
        }
    }
}

fn hyperlink_relationship_ids(document: &str) -> Vec<&str> {
    let marker = "r:id=\"";
    document
        .match_indices(marker)
        .filter_map(|(index, _)| {
            let value = &document[index + marker.len()..];
            let end = value.find('"')?;
            let id = &value[..end];
            id.starts_with("rIdHyperlink").then_some(id)
        })
        .collect()
}

fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn compile_section_element(node: &Node, path: &str) -> Result<Section> {
    let mut section = Section::new();
    let orientation = optional_enum(&node.props, "orientation", &["portrait", "landscape"], path)?;
    if let Some(page_size) = node.props.get("pageSize") {
        let (mut width, mut height) = parse_page_size(page_size, path)?;
        if orientation == Some("landscape") && width < height {
            std::mem::swap(&mut width, &mut height);
        }
        section = section.page_size(PageSize::new().size(width, height));
    }
    if let Some(value) = orientation {
        section = section.page_orient(match value {
            "landscape" => PageOrientationType::Landscape,
            _ => PageOrientationType::Portrait,
        });
    }
    if let Some(margins) = node.props.get("margins") {
        section = section.page_margin(parse_margins(margins, path)?);
    }
    if bool_prop(&node.props, "titlePage", path)? == Some(true) {
        section = section.title_pg();
    }
    if let Some(direction) = optional_enum(
        &node.props,
        "textDirection",
        &["lrTb", "tbRl", "btLr", "lrTbV", "tbRlV"],
        path,
    )? {
        section = section.text_direction(direction.to_owned());
    }
    if let Some(value) = node.props.get("documentGrid") {
        section = section.doc_grid(parse_document_grid(value, path)?);
    }
    if let Some(value) = node.props.get("pageNumbering") {
        section = section.page_num_type(parse_page_numbering(value, path)?);
    }
    Ok(section)
}

fn apply_section_properties(mut docx: Docx, node: &Node, path: &str) -> Result<Docx> {
    let orientation = optional_enum(&node.props, "orientation", &["portrait", "landscape"], path)?;
    if let Some(page_size) = node.props.get("pageSize") {
        let (mut width, mut height) = parse_page_size(page_size, path)?;
        if orientation == Some("landscape") && width < height {
            std::mem::swap(&mut width, &mut height);
        }
        docx = docx.page_size(width, height);
    }
    if let Some(value) = orientation {
        docx = docx.page_orient(match value {
            "landscape" => PageOrientationType::Landscape,
            _ => PageOrientationType::Portrait,
        });
    }
    if let Some(margins) = node.props.get("margins") {
        docx = docx.page_margin(parse_margins(margins, path)?);
    }
    if bool_prop(&node.props, "titlePage", path)? == Some(true) {
        docx = docx.title_pg();
    }
    if let Some(direction) = optional_enum(
        &node.props,
        "textDirection",
        &["lrTb", "tbRl", "btLr", "lrTbV", "tbRlV"],
        path,
    )? {
        docx.document = docx.document.text_direction(direction.to_owned());
    }
    if let Some(value) = node.props.get("documentGrid") {
        docx.document = docx.document.doc_grid(parse_document_grid(value, path)?);
    }
    if let Some(value) = node.props.get("pageNumbering") {
        docx = docx.page_num_type(parse_page_numbering(value, path)?);
    }
    Ok(docx)
}

fn add_section_body_to_docx(
    docx: Docx,
    child: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Docx> {
    Ok(match child.kind {
        NodeKind::Paragraph => {
            docx.add_paragraph(compile_paragraph(child, entry_dir, path, context)?)
        }
        NodeKind::Heading => docx.add_paragraph(compile_heading(child, entry_dir, path, context)?),
        NodeKind::Caption => docx.add_paragraph(compile_caption(child, entry_dir, path, context)?),
        NodeKind::Index => docx.add_paragraph(compile_index(child, path)?),
        NodeKind::Table => docx.add_table(compile_table(child, entry_dir, path, context)?),
        NodeKind::List => add_list_to_docx(docx, child, entry_dir, path, context)?,
        NodeKind::Bookmark => add_bookmark_to_docx(docx, child, entry_dir, path, context)?,
        NodeKind::ContentControl => docx.add_structured_data_tag(compile_block_content_control(
            child, entry_dir, path, context,
        )?),
        NodeKind::TableOfContents
        | NodeKind::TableOfFigures
        | NodeKind::TableOfEntries
        | NodeKind::Header
        | NodeKind::Footer => docx,
        _ => return Err(validation(path, "unsupported Section child")),
    })
}

fn add_list_to_docx(
    mut docx: Docx,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Docx> {
    for paragraph in compile_list(node, entry_dir, path, context)? {
        docx = docx.add_paragraph(paragraph);
    }
    Ok(docx)
}

fn add_bookmark_to_docx(
    mut docx: Docx,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Docx> {
    context.next_bookmark_id += 1;
    let id = context.next_bookmark_id;
    docx = docx.add_bookmark_start(id, required_string(&node.props, "name", path)?);
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(child) = child else {
            return Err(validation(
                path,
                "Bookmark only accepts structural children",
            ));
        };
        let child_path = format!("{path}/{}[{index}]", child.kind.name());
        docx = add_section_body_to_docx(docx, child, entry_dir, &child_path, context)?;
    }
    Ok(docx.add_bookmark_end(id))
}

fn compile_block_content_control(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<StructuredDataTag> {
    let mut control = content_control_properties(node, path)?;
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(child) = child else {
            return Err(validation(
                path,
                "document-level ContentControl requires block children",
            ));
        };
        let child_path = format!("{path}/{}[{index}]", child.kind.name());
        control = match child.kind {
            NodeKind::Paragraph => {
                control.add_paragraph(compile_paragraph(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Heading => {
                control.add_paragraph(compile_heading(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Caption => {
                control.add_paragraph(compile_caption(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Table => {
                control.add_table(compile_table(child, entry_dir, &child_path, context)?)
            }
            NodeKind::List => {
                for paragraph in compile_list(child, entry_dir, &child_path, context)? {
                    control = control.add_paragraph(paragraph);
                }
                control
            }
            _ => return Err(validation(&child_path, "unsupported ContentControl child")),
        };
    }
    Ok(control)
}

fn content_control_properties(node: &Node, path: &str) -> Result<StructuredDataTag> {
    let mut control = StructuredDataTag::new();
    if let Some(alias) = string_prop(&node.props, "alias", path)? {
        control = control.alias(alias);
    }
    if let Some(xpath) = string_prop(&node.props, "xpath", path)? {
        let mut binding = DataBinding::new().xpath(xpath);
        if let Some(value) = string_prop(&node.props, "prefixMappings", path)? {
            binding = binding.prefix_mappings(value);
        }
        if let Some(value) = string_prop(&node.props, "storeItemId", path)? {
            binding = binding.store_item_id(value);
        }
        control = control.data_binding(binding);
    }
    Ok(control)
}

fn parse_document_grid(value: &Value, path: &str) -> Result<DocGrid> {
    let grid = value
        .as_object()
        .ok_or_else(|| validation(path, "`documentGrid` must be an object"))?;
    let grid_type = string_prop(grid, "type", path)?
        .ok_or_else(|| validation(path, "`documentGrid.type` is required"))?;
    let mut result = DocGrid::with_empty().grid_type(match grid_type {
        "default" => DocGridType::Default,
        "lines" => DocGridType::Lines,
        "linesAndChars" => DocGridType::LinesAndChars,
        "snapToChars" => DocGridType::SnapToChars,
        value => {
            return Err(validation(
                path,
                format!("invalid document grid type `{value}`"),
            ));
        }
    });
    if let Some(line_pitch) = number_prop(grid, "linePitch", path)? {
        result = result.line_pitch(to_twips_usize(line_pitch, path)?);
    }
    if let Some(char_space) = grid.get("charSpace") {
        let value = char_space
            .as_i64()
            .and_then(|value| isize::try_from(value).ok())
            .ok_or_else(|| validation(path, "`documentGrid.charSpace` is out of range"))?;
        result = result.char_space(value);
    }
    Ok(result)
}

fn parse_page_numbering(value: &Value, path: &str) -> Result<PageNumType> {
    let numbering = value
        .as_object()
        .ok_or_else(|| validation(path, "`pageNumbering` must be an object"))?;
    let mut result = PageNumType::new();
    if let Some(start) = numbering.get("start") {
        let start = start
            .as_u64()
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| validation(path, "`pageNumbering.start` is out of range"))?;
        result = result.start(start);
    }
    if let Some(chapter_style) = string_prop(numbering, "chapterStyle", path)? {
        result = result.chap_style(chapter_style);
    }
    Ok(result)
}

fn paragraph_alignment(value: &str) -> AlignmentType {
    match value {
        "center" => AlignmentType::Center,
        "right" => AlignmentType::Right,
        "both" => AlignmentType::Both,
        "distribute" => AlignmentType::Distribute,
        "start" => AlignmentType::Start,
        "end" => AlignmentType::End,
        "justified" => AlignmentType::Justified,
        _ => AlignmentType::Left,
    }
}

fn compile_paragraph(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Paragraph> {
    let (mut paragraph, has_character_indents) =
        register_paragraph_character_indentation(Paragraph::new(), node, context);
    if let Some(style) = string_prop(&node.props, "style", path)? {
        paragraph = paragraph.style(style);
    }
    if let Some(value) = optional_enum(
        &node.props,
        "align",
        &[
            "left",
            "center",
            "right",
            "both",
            "distribute",
            "start",
            "end",
            "justified",
        ],
        path,
    )? {
        paragraph = paragraph.align(paragraph_alignment(value));
    }
    let mut spacing = LineSpacing::new();
    let mut has_spacing = false;
    if let Some(value) = number_prop(&node.props, "spacingBefore", path)? {
        spacing = spacing.before(to_twips_u32(value, &format!("{path}/spacingBefore"))?);
        has_spacing = true;
    }
    if let Some(value) = number_prop(&node.props, "spacingAfter", path)? {
        spacing = spacing.after(to_twips_u32(value, &format!("{path}/spacingAfter"))?);
        has_spacing = true;
    }
    if let Some(value) = number_prop(&node.props, "lineSpacing", path)? {
        spacing = spacing.line(to_twips_i32(value, &format!("{path}/lineSpacing"))?);
        has_spacing = true;
    }
    let (next, extras_present) = compile_line_spacing_extras(
        spacing,
        &node.props,
        "spacingBeforeLines",
        "spacingAfterLines",
        "lineRule",
        path,
    )?;
    spacing = next;
    has_spacing |= extras_present;
    if has_spacing {
        paragraph = paragraph.line_spacing(spacing);
    }
    paragraph = compile_paragraph_indent(paragraph, node, path, has_character_indents)?;
    if let Some(value) = bool_prop(&node.props, "snapToGrid", path)? {
        paragraph = paragraph.snap_to_grid(value);
    }
    if let Some(value) = bool_prop(&node.props, "widowControl", path)? {
        paragraph = paragraph.widow_control(value);
    }
    paragraph = compile_paragraph_fonts(paragraph, &node.props, path)?;
    if let Some(size) = number_prop(&node.props, "size", path)? {
        paragraph = paragraph.size(to_half_points(size, &format!("{path}/size"))?);
    }
    if bool_prop(&node.props, "bold", path)?.unwrap_or(false) {
        paragraph = paragraph.bold();
    }
    if bool_prop(&node.props, "italic", path)?.unwrap_or(false) {
        paragraph = paragraph.italic();
    }
    if let Some(color) = string_prop(&node.props, "color", path)? {
        paragraph = paragraph.color(color.to_ascii_uppercase());
    }
    if let Some(spacing) = number_prop(&node.props, "characterSpacing", path)? {
        paragraph = paragraph
            .character_spacing(to_twips_i32(spacing, &format!("{path}/characterSpacing"))?);
    }
    if bool_prop(&node.props, "keepNext", path)?.unwrap_or(false) {
        paragraph = paragraph.keep_next(true);
    }
    if bool_prop(&node.props, "keepLines", path)?.unwrap_or(false) {
        paragraph = paragraph.keep_lines(true);
    }
    if bool_prop(&node.props, "pageBreakBefore", path)?.unwrap_or(false) {
        paragraph = paragraph.page_break_before(true);
    }
    paragraph = compile_advanced_paragraph_properties(paragraph, node, path)?;
    compile_paragraph_children(paragraph, &node.children, entry_dir, path, context)
}

fn compile_paragraph_indent(
    mut paragraph: Paragraph,
    node: &Node,
    path: &str,
    has_character_indents: bool,
) -> Result<Paragraph> {
    let left = optional_twips_i32(&node.props, "indentLeft", path)?;
    let right = optional_twips_i32(&node.props, "indentRight", path)?;
    let special = if let Some(value) = number_prop(&node.props, "firstLine", path)? {
        Some(SpecialIndentType::FirstLine(to_twips_i32(value, path)?))
    } else if let Some(value) = number_prop(&node.props, "hanging", path)? {
        Some(SpecialIndentType::Hanging(to_twips_i32(value, path)?))
    } else {
        None
    };
    if left.is_some() || right.is_some() || special.is_some() || has_character_indents {
        paragraph = paragraph.indent(left, special, right, None);
    }
    if let Some(value) = node.props.get("hangingChars") {
        paragraph = paragraph.hanging_chars(value_to_i32(value, path, "hangingChars")?);
    }
    if let Some(value) = node.props.get("firstLineChars") {
        paragraph = paragraph.first_line_chars(value_to_i32(value, path, "firstLineChars")?);
    }
    Ok(paragraph)
}

fn compile_paragraph_fonts(
    mut paragraph: Paragraph,
    props: &Map<String, Value>,
    path: &str,
) -> Result<Paragraph> {
    if let Some(font) = string_prop(props, "font", path)? {
        paragraph = paragraph.fonts(
            RunFonts::new()
                .ascii(font)
                .hi_ansi(font)
                .east_asia(font)
                .cs(font),
        );
    }
    if let Some(fonts) = object_prop(props, "fonts", path)? {
        paragraph = paragraph.fonts(compile_run_fonts(fonts, &format!("{path}/fonts"))?);
    }
    Ok(paragraph)
}

fn compile_advanced_paragraph_properties(
    mut paragraph: Paragraph,
    node: &Node,
    path: &str,
) -> Result<Paragraph> {
    if let Some(value) = string_prop(&node.props, "paragraphId", path)? {
        paragraph = paragraph.id(value.to_ascii_uppercase());
    }
    if let Some(value) = bool_prop(&node.props, "bidi", path)? {
        paragraph.property = paragraph.property.bidi(value);
    }
    if let Some(value) = optional_enum(
        &node.props,
        "textAlign",
        &["auto", "baseline", "bottom", "center", "top"],
        path,
    )? {
        paragraph.property = paragraph.property.text_alignment(match value {
            "baseline" => TextAlignmentType::Baseline,
            "bottom" => TextAlignmentType::Bottom,
            "center" => TextAlignmentType::Center,
            "top" => TextAlignmentType::Top,
            _ => TextAlignmentType::Auto,
        });
    }
    if let Some(value) = node.props.get("adjustRightIndent") {
        let value = value
            .as_i64()
            .and_then(|value| isize::try_from(value).ok())
            .ok_or_else(|| validation(path, "`adjustRightIndent` is out of range"))?;
        paragraph.property = paragraph.property.adjust_right_ind(value);
    }
    if let Some(value) = string_prop(&node.props, "shading", path)? {
        paragraph.property = paragraph
            .property
            .shading(Shading::new().fill(value.to_ascii_uppercase()));
    }
    if let Some(value) = node.props.get("outlineLevel") {
        paragraph = paragraph.outline_lvl(value_to_usize(value, path, "outlineLevel")?);
    }
    if let Some(value) = node.props.get("frame") {
        paragraph = compile_paragraph_frame(paragraph, value, path)?;
    }
    if let Some(value) = node.props.get("border") {
        paragraph.property = paragraph
            .property
            .set_borders(compile_paragraph_borders(value, path)?);
    }
    if let Some(value) = node.props.get("inserted") {
        paragraph.property.run_property.ins = Some(compile_row_insert(value, path)?);
    }
    if let Some(value) = node.props.get("deleted") {
        paragraph.property.run_property.del = Some(compile_row_delete(value, path)?);
    }
    if let Some(value) = node.props.get("propertyChange") {
        paragraph = compile_paragraph_property_change(paragraph, value, path)?;
    }
    Ok(paragraph)
}

fn compile_paragraph_frame(
    mut paragraph: Paragraph,
    value: &Value,
    path: &str,
) -> Result<Paragraph> {
    let frame = value
        .as_object()
        .ok_or_else(|| validation(path, "`frame` must be an object"))?;
    if let Some(value) = string_prop(frame, "wrap", path)? {
        paragraph = paragraph.wrap(value);
    }
    if let Some(value) = string_prop(frame, "verticalAnchor", path)? {
        paragraph = paragraph.v_anchor(value);
    }
    if let Some(value) = string_prop(frame, "horizontalAnchor", path)? {
        paragraph = paragraph.h_anchor(value);
    }
    if let Some(value) = string_prop(frame, "heightRule", path)? {
        paragraph = paragraph.h_rule(value);
    }
    if let Some(value) = string_prop(frame, "xAlign", path)? {
        paragraph = paragraph.x_align(value);
    }
    if let Some(value) = string_prop(frame, "yAlign", path)? {
        paragraph = paragraph.y_align(value);
    }
    for (key, apply) in [
        (
            "horizontalSpace",
            Paragraph::h_space as fn(Paragraph, i32) -> Paragraph,
        ),
        ("verticalSpace", Paragraph::v_space),
        ("x", Paragraph::frame_x),
        ("y", Paragraph::frame_y),
    ] {
        if let Some(value) = number_prop(frame, key, path)? {
            paragraph = apply(
                paragraph,
                to_twips_i32(value, &format!("{path}/frame/{key}"))?,
            );
        }
    }
    if let Some(value) = number_prop(frame, "width", path)? {
        paragraph = paragraph.frame_width(to_twips_u32(value, &format!("{path}/frame/width"))?);
    }
    if let Some(value) = number_prop(frame, "height", path)? {
        paragraph = paragraph.frame_height(to_twips_u32(value, &format!("{path}/frame/height"))?);
    }
    Ok(paragraph)
}

fn compile_paragraph_property_change(
    mut paragraph: Paragraph,
    value: &Value,
    path: &str,
) -> Result<Paragraph> {
    let change = value
        .as_object()
        .ok_or_else(|| validation(path, "`propertyChange` must be an object"))?;
    let previous = change
        .get("previous")
        .and_then(Value::as_object)
        .ok_or_else(|| validation(path, "`propertyChange.previous` must be an object"))?;
    let previous_node = Node {
        kind: NodeKind::Paragraph,
        props: previous.clone(),
        children: Vec::new(),
    };
    let previous = compile_paragraph(
        &previous_node,
        Path::new("."),
        &format!("{path}/propertyChange/previous"),
        &mut CompileContext::default(),
    )?;
    let mut revision = ParagraphPropertyChange::new().property(previous.property);
    if let Some(author) = string_prop(change, "author", path)? {
        revision = revision.author(author);
    }
    if let Some(date) = string_prop(change, "date", path)? {
        revision = revision.date(date);
    }
    paragraph.property = paragraph.property.paragraph_property_change(revision);
    Ok(paragraph)
}

fn compile_paragraph_children(
    mut paragraph: Paragraph,
    children: &[Child],
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Paragraph> {
    for (index, child) in children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                paragraph = paragraph.add_run(compile_run(run, entry_dir, &child_path)?);
            }
            Child::Node(link) if link.kind == NodeKind::Hyperlink => {
                paragraph = paragraph.add_hyperlink(compile_hyperlink(
                    link,
                    entry_dir,
                    &child_path,
                    context,
                )?);
            }
            Child::Node(field) if field.kind == NodeKind::PageNumber => {
                paragraph = paragraph.add_page_num(PageNum::new());
            }
            Child::Node(field) if field.kind == NodeKind::TotalPages => {
                paragraph = paragraph.add_num_pages(NumPages::new());
            }
            Child::Node(field) if field.kind.is_field() => {
                paragraph = add_field_to_paragraph(paragraph, field, entry_dir, &child_path)?;
            }
            Child::Node(comment) if comment.kind == NodeKind::Comment => {
                paragraph =
                    compile_comment_range(paragraph, comment, entry_dir, &child_path, context)?;
            }
            Child::Node(inserted) if inserted.kind == NodeKind::Inserted => {
                paragraph =
                    paragraph.add_insert(compile_inserted(inserted, entry_dir, &child_path)?);
            }
            Child::Node(deleted) if deleted.kind == NodeKind::Deleted => {
                paragraph = paragraph.add_delete(compile_deleted(deleted, &child_path, context)?);
            }
            Child::Node(moved) if moved.kind == NodeKind::MovedFrom => {
                paragraph =
                    paragraph.add_move_from(compile_moved_from(moved, entry_dir, &child_path)?);
            }
            Child::Node(moved) if moved.kind == NodeKind::MovedTo => {
                paragraph = paragraph.add_move_to(compile_moved_to(moved, entry_dir, &child_path)?);
            }
            Child::Node(control) if control.kind == NodeKind::ContentControl => {
                paragraph = paragraph.add_structured_data_tag(compile_content_control(
                    control,
                    entry_dir,
                    &child_path,
                )?);
            }
            Child::Node(bookmark) if bookmark.kind == NodeKind::InlineBookmark => {
                paragraph =
                    compile_inline_bookmark(paragraph, bookmark, entry_dir, &child_path, context)?;
            }
            Child::Node(tab_stop) if tab_stop.kind == NodeKind::TabStop => {
                paragraph = paragraph.add_tab(compile_tab_stop(tab_stop, &child_path)?);
            }
            Child::Node(wrapper) if wrapper.kind.is_semantic_text() => {
                paragraph =
                    paragraph.add_run(compile_semantic_text(wrapper, entry_dir, &child_path)?);
            }
            other => {
                let mut run = Run::new();
                run = compile_run_child(run, other, entry_dir, &child_path)?;
                paragraph = paragraph.add_run(run);
            }
        }
    }
    Ok(paragraph)
}

fn compile_inline_bookmark(
    mut paragraph: Paragraph,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Paragraph> {
    context.next_bookmark_id += 1;
    let id = context.next_bookmark_id;
    paragraph = paragraph.add_bookmark_start(id, required_string(&node.props, "name", path)?);
    paragraph = compile_paragraph_children(paragraph, &node.children, entry_dir, path, context)?;
    Ok(paragraph.add_bookmark_end(id))
}

fn compile_caption(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Paragraph> {
    let label = required_string(&node.props, "label", path)?;
    let identifier = string_prop(&node.props, "identifier", path)?.unwrap_or(label);
    let style = string_prop(&node.props, "style", path)?.unwrap_or("Caption");
    let number_separator = string_prop(&node.props, "numberSeparator", path)?.unwrap_or(" ");
    let text_separator = string_prop(&node.props, "textSeparator", path)?.unwrap_or(": ");
    let dirty = bool_prop(&node.props, "dirty", path)?.unwrap_or(true);
    let instruction = sequence_instruction(identifier, node, path)?;
    let placeholder = string_prop(&node.props, "placeholder", path)?.unwrap_or("1");

    let mut paragraph = Paragraph::new()
        .style(style)
        .add_run(Run::new().add_text(format!("{label}{number_separator}")))
        .add_run(
            Run::new()
                .add_field_char(FieldCharType::Begin, dirty)
                .add_instr_text(InstrText::Unsupported(instruction))
                .add_field_char(FieldCharType::Separate, false),
        )
        .add_run(Run::new().add_text(placeholder))
        .add_run(Run::new().add_field_char(FieldCharType::End, false))
        .add_run(Run::new().add_text(text_separator));
    paragraph = compile_paragraph_children(paragraph, &node.children, entry_dir, path, context)?;
    Ok(paragraph)
}

fn compile_index(node: &Node, path: &str) -> Result<Paragraph> {
    let instruction = index_instruction(node, path)?;
    let dirty = bool_prop(&node.props, "dirty", path)?.unwrap_or(true);
    let placeholder = string_prop(&node.props, "placeholder", path)?.unwrap_or("Update index");
    let mut paragraph = Paragraph::new();
    if let Some(style) = string_prop(&node.props, "style", path)? {
        paragraph = paragraph.style(style);
    }
    Ok(paragraph
        .add_run(
            Run::new()
                .add_field_char(FieldCharType::Begin, dirty)
                .add_instr_text(InstrText::Unsupported(instruction))
                .add_field_char(FieldCharType::Separate, false),
        )
        .add_run(Run::new().add_text(placeholder))
        .add_run(Run::new().add_field_char(FieldCharType::End, false)))
}

fn compile_heading(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Paragraph> {
    let level = usize::try_from(required_u64(&node.props, "level", path)?)
        .map_err(|_| validation(path, "heading level is out of range"))?;
    let mut paragraph = compile_paragraph(node, entry_dir, path, context)?.outline_lvl(level - 1);
    if !node.props.contains_key("style") {
        paragraph = paragraph.style(&format!("Heading{level}"));
    }
    Ok(paragraph)
}

fn compile_comment_range(
    mut paragraph: Paragraph,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Paragraph> {
    context.next_comment_id += 1;
    let id = context.next_comment_id;
    let body =
        Paragraph::new().add_run(Run::new().add_text(required_string(&node.props, "text", path)?));
    let mut comment = Comment::new(id).add_paragraph(body);
    if let Some(author) = string_prop(&node.props, "author", path)? {
        comment = comment.author(author);
    }
    if let Some(date) = string_prop(&node.props, "date", path)? {
        comment = comment.date(date);
    }
    paragraph = paragraph.add_comment_start(comment);
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                paragraph = paragraph.add_run(compile_run(run, entry_dir, &child_path)?);
            }
            Child::Node(link) if link.kind == NodeKind::Hyperlink => {
                paragraph = paragraph.add_hyperlink(compile_hyperlink(
                    link,
                    entry_dir,
                    &child_path,
                    context,
                )?);
            }
            child => {
                paragraph = paragraph.add_run(compile_run_child(
                    Run::new(),
                    child,
                    entry_dir,
                    &child_path,
                )?);
            }
        }
    }
    Ok(paragraph.add_comment_end(id))
}

fn compile_inserted(node: &Node, entry_dir: &Path, path: &str) -> Result<Insert> {
    let mut inserted = Insert::new_with_empty();
    if let Some(author) = string_prop(&node.props, "author", path)? {
        inserted = inserted.author(author);
    }
    if let Some(date) = string_prop(&node.props, "date", path)? {
        inserted = inserted.date(date);
    }
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        let run = match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                compile_run(run, entry_dir, &child_path)?
            }
            child => compile_run_child(Run::new(), child, entry_dir, &child_path)?,
        };
        inserted = inserted.add_run(run);
    }
    Ok(inserted)
}

fn compile_deleted(node: &Node, path: &str, context: &mut CompileContext) -> Result<Delete> {
    let mut deleted = Delete::new();
    if let Some(author) = string_prop(&node.props, "author", path)? {
        deleted = deleted.author(author);
    }
    if let Some(date) = string_prop(&node.props, "date", path)? {
        deleted = deleted.date(date);
    }
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        if let Child::Node(comment_node) = child
            && comment_node.kind == NodeKind::Comment
        {
            context.next_comment_id += 1;
            let id = context.next_comment_id;
            let text = required_string(&comment_node.props, "text", &child_path)?;
            let comment_run = Run::new().add_text(text);
            let mut comment = Comment::new(id).add_paragraph(Paragraph::new().add_run(comment_run));
            if let Some(author) = string_prop(&comment_node.props, "author", &child_path)? {
                comment = comment.author(author);
            }
            if let Some(date) = string_prop(&comment_node.props, "date", &child_path)? {
                comment = comment.date(date);
            }
            context.deleted_comment_dependencies.push(comment.clone());
            deleted = deleted.add_comment_start(comment);
            for (selected_index, selected) in comment_node.children.iter().enumerate() {
                let selected_path = format!("{child_path}/child[{selected_index}]");
                let run = match selected {
                    Child::Node(run) if run.kind == NodeKind::Run => {
                        let mut output = compile_run_properties(&run.props, &selected_path)?;
                        for text_child in &run.children {
                            output = add_deleted_text(output, text_child, &selected_path)?;
                        }
                        output
                    }
                    selected => add_deleted_text(Run::new(), selected, &selected_path)?,
                };
                deleted = deleted.add_run(run);
            }
            deleted = deleted.add_comment_end(id);
            continue;
        }
        let run = match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                let mut output = compile_run_properties(&run.props, &child_path)?;
                for child in &run.children {
                    output = add_deleted_text(output, child, &child_path)?;
                }
                output
            }
            child => add_deleted_text(Run::new(), child, &child_path)?,
        };
        deleted = deleted.add_run(run);
    }
    Ok(deleted)
}

fn add_deleted_text(mut run: Run, child: &Child, path: &str) -> Result<Run> {
    match child {
        Child::String(value) => Ok(run.add_delete_text(value)),
        Child::Number(value) => Ok(run.add_delete_text(value.to_string())),
        Child::Node(node) if node.kind == NodeKind::Text => {
            if let Some(value) = node.props.get("value") {
                run = run.add_delete_text(value_to_text(value, path)?);
            }
            for child in &node.children {
                run = add_deleted_text(run, child, path)?;
            }
            Ok(run)
        }
        Child::Node(node) if node.kind == NodeKind::DeletedInstruction => {
            let instruction = required_string(&node.props, "instruction", path)?;
            Ok(run.add_delete_instr_text(DeleteInstrText::Unsupported(instruction.to_owned())))
        }
        Child::Node(_) => Err(validation(path, "Deleted only accepts text content")),
    }
}

fn compile_hyperlink(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Hyperlink> {
    let (target, kind) = if let Some(href) = string_prop(&node.props, "href", path)? {
        (href, HyperlinkType::External)
    } else {
        (
            required_string(&node.props, "anchor", path)?,
            HyperlinkType::Anchor,
        )
    };
    let mut hyperlink = Hyperlink::new(target, kind);
    if let Some(history) = bool_prop(&node.props, "history", path)? {
        hyperlink.history = Some(usize::from(history));
    }
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        hyperlink = compile_hyperlink_child(hyperlink, child, entry_dir, &child_path, context)?;
    }
    Ok(hyperlink)
}

fn compile_hyperlink_child(
    mut hyperlink: Hyperlink,
    child: &Child,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Hyperlink> {
    let Child::Node(node) = child else {
        return Ok(hyperlink.add_run(compile_run_child(Run::new(), child, entry_dir, path)?));
    };
    match node.kind {
        NodeKind::Run => Ok(hyperlink.add_run(compile_run(node, entry_dir, path)?)),
        NodeKind::ContentControl => {
            Ok(hyperlink.add_structured_data_tag(compile_content_control(node, entry_dir, path)?))
        }
        NodeKind::Inserted => Ok(hyperlink.add_insert(compile_inserted(node, entry_dir, path)?)),
        NodeKind::Deleted => Ok(hyperlink.add_delete(compile_deleted(node, path, context)?)),
        NodeKind::InlineBookmark => {
            context.next_bookmark_id += 1;
            let id = context.next_bookmark_id;
            hyperlink =
                hyperlink.add_bookmark_start(id, required_string(&node.props, "name", path)?);
            for (index, child) in node.children.iter().enumerate() {
                hyperlink = compile_hyperlink_child(
                    hyperlink,
                    child,
                    entry_dir,
                    &format!("{path}/child[{index}]"),
                    context,
                )?;
            }
            Ok(hyperlink.add_bookmark_end(id))
        }
        NodeKind::Comment => {
            context.next_comment_id += 1;
            let id = context.next_comment_id;
            let body = Paragraph::new().add_run(Run::new().add_text(required_string(
                &node.props,
                "text",
                path,
            )?));
            let mut comment = Comment::new(id).add_paragraph(body);
            if let Some(author) = string_prop(&node.props, "author", path)? {
                comment = comment.author(author);
            }
            if let Some(date) = string_prop(&node.props, "date", path)? {
                comment = comment.date(date);
            }
            hyperlink = hyperlink.add_comment_start(comment);
            for (index, child) in node.children.iter().enumerate() {
                hyperlink = compile_hyperlink_child(
                    hyperlink,
                    child,
                    entry_dir,
                    &format!("{path}/child[{index}]"),
                    context,
                )?;
            }
            Ok(hyperlink.add_comment_end(id))
        }
        NodeKind::Hyperlink => Err(validation(
            path,
            "nested Hyperlink is not supported by docx-rs",
        )),
        _ => Ok(hyperlink.add_run(compile_run_child(Run::new(), child, entry_dir, path)?)),
    }
}

fn compile_content_control(node: &Node, entry_dir: &Path, path: &str) -> Result<StructuredDataTag> {
    let mut control = content_control_properties(node, path)?;
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        let run = match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                compile_run(run, entry_dir, &child_path)?
            }
            child => compile_run_child(Run::new(), child, entry_dir, &child_path)?,
        };
        control = control.add_run(run);
    }
    Ok(control)
}

fn compile_moved_from(node: &Node, entry_dir: &Path, path: &str) -> Result<MoveFrom> {
    let mut moved = MoveFrom::new();
    if let Some(author) = string_prop(&node.props, "author", path)? {
        moved = moved.author(author);
    }
    if let Some(date) = string_prop(&node.props, "date", path)? {
        moved = moved.date(date);
    }
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        let run = match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                compile_run(run, entry_dir, &child_path)?
            }
            child => compile_run_child(Run::new(), child, entry_dir, &child_path)?,
        };
        moved = moved.add_run(run);
    }
    Ok(moved)
}

fn compile_semantic_text(node: &Node, entry_dir: &Path, path: &str) -> Result<Run> {
    let mut run = Run::new();
    match node.kind {
        NodeKind::Bold => run = run.bold(),
        NodeKind::Italic => run = run.italic(),
        NodeKind::Underline => {
            let underline = optional_enum(
                &node.props,
                "type",
                &["single", "double", "dotted", "dash", "wave"],
                path,
            )?
            .unwrap_or("single");
            run = run.underline(underline);
        }
        NodeKind::StrikeThrough => run = run.strike(),
        NodeKind::Superscript => {
            run.run_property = run.run_property.vert_align(VertAlignType::SuperScript);
        }
        NodeKind::Subscript => {
            run.run_property = run.run_property.vert_align(VertAlignType::SubScript);
        }
        NodeKind::AllCaps => {
            run.run_property = run.run_property.caps();
        }
        NodeKind::HiddenText => run = run.vanish(),
        NodeKind::SpecialHiddenText => {
            run.run_property = run.run_property.spec_vanish();
        }
        NodeKind::DoubleStrike => run = run.dstrike(),
        NodeKind::SpacedText => {
            run = run.character_spacing(to_twips_i32(
                required_number(&node.props, "amount", path)?,
                &format!("{path}/amount"),
            )?);
        }
        NodeKind::ScaledText => {
            let percent = i32::try_from(required_u64(&node.props, "percent", path)?)
                .map_err(|_| validation(path, "ScaledText percent is out of range"))?;
            run = run.stretch(percent);
        }
        NodeKind::FitText => {
            let width = to_twips_usize(
                required_number(&node.props, "width", path)?,
                &format!("{path}/width"),
            )?;
            let id = node
                .props
                .get("id")
                .and_then(Value::as_u64)
                .map(u32::try_from)
                .transpose()
                .map_err(|_| validation(path, "FitText id is out of range"))?;
            run.run_property = run.run_property.fit_text(width, id);
        }
        NodeKind::BorderedText => run = compile_bordered_text(run, node, path)?,
        NodeKind::ShadedText => run = compile_shaded_text(run, node, path)?,
        _ => return Err(validation(path, "unsupported semantic text component")),
    }
    for (index, child) in node.children.iter().enumerate() {
        run = compile_run_child(run, child, entry_dir, &format!("{path}/child[{index}]"))?;
    }
    Ok(run)
}

fn compile_bordered_text(run: Run, node: &Node, path: &str) -> Result<Run> {
    let style = optional_enum(
        &node.props,
        "style",
        &["single", "double", "dotted", "dashed"],
        path,
    )?
    .unwrap_or("single");
    let size = number_prop(&node.props, "size", path)?.unwrap_or(0.5);
    let color = string_prop(&node.props, "color", path)?.unwrap_or("000000");
    let space = node
        .props
        .get("space")
        .and_then(Value::as_u64)
        .map(usize::try_from)
        .transpose()
        .map_err(|_| validation(path, "BorderedText space is out of range"))?
        .unwrap_or(0);
    let border = TextBorder::new()
        .border_type(match style {
            "double" => BorderType::Double,
            "dotted" => BorderType::Dotted,
            "dashed" => BorderType::Dashed,
            _ => BorderType::Single,
        })
        .size(f64_to_usize(size * 8.0, path)?)
        .color(color.to_ascii_uppercase())
        .space(space);
    Ok(run.text_border(border))
}

fn compile_shaded_text(run: Run, node: &Node, path: &str) -> Result<Run> {
    let fill = required_string(&node.props, "fill", path)?;
    let color = string_prop(&node.props, "color", path)?.unwrap_or("auto");
    let pattern = string_prop(&node.props, "pattern", path)?.unwrap_or("clear");
    let shading_type = pattern
        .parse::<ShdType>()
        .map_err(|_| validation(path, "ShadedText pattern is invalid"))?;
    Ok(run.shading(
        Shading::new()
            .shd_type(shading_type)
            .fill(fill.to_ascii_uppercase())
            .color(if color == "auto" {
                color.to_owned()
            } else {
                color.to_ascii_uppercase()
            }),
    ))
}

fn compile_moved_to(node: &Node, entry_dir: &Path, path: &str) -> Result<MoveTo> {
    let mut moved = MoveTo::new_with_empty();
    if let Some(author) = string_prop(&node.props, "author", path)? {
        moved = moved.author(author);
    }
    if let Some(date) = string_prop(&node.props, "date", path)? {
        moved = moved.date(date);
    }
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        let run = match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                compile_run(run, entry_dir, &child_path)?
            }
            child => compile_run_child(Run::new(), child, entry_dir, &child_path)?,
        };
        moved = moved.add_run(run);
    }
    Ok(moved)
}

fn compile_tab_stop(node: &Node, path: &str) -> Result<DocxTab> {
    let alignment = match optional_enum(
        &node.props,
        "align",
        &["left", "center", "right", "decimal", "bar", "clear"],
        path,
    )?
    .unwrap_or("left")
    {
        "center" => TabValueType::Center,
        "right" => TabValueType::Right,
        "decimal" => TabValueType::Decimal,
        "bar" => TabValueType::Bar,
        "clear" => TabValueType::Clear,
        _ => TabValueType::Left,
    };
    let leader = match optional_enum(
        &node.props,
        "leader",
        &["none", "dot", "heavy", "hyphen", "middleDot", "underscore"],
        path,
    )?
    .unwrap_or("none")
    {
        "dot" => TabLeaderType::Dot,
        "heavy" => TabLeaderType::Heavy,
        "hyphen" => TabLeaderType::Hyphen,
        "middleDot" => TabLeaderType::MiddleDot,
        "underscore" => TabLeaderType::Underscore,
        _ => TabLeaderType::None,
    };
    Ok(DocxTab::new()
        .val(alignment)
        .leader(leader)
        .pos(to_twips_usize(
            required_number(&node.props, "position", path)?,
            &format!("{path}/position"),
        )?))
}

fn compile_run(node: &Node, entry_dir: &Path, path: &str) -> Result<Run> {
    let mut run = compile_run_properties(&node.props, path)?;
    for (index, child) in node.children.iter().enumerate() {
        run = compile_run_child(run, child, entry_dir, &format!("{path}/child[{index}]"))?;
    }
    Ok(run)
}

fn compile_run_properties(props: &Map<String, Value>, path: &str) -> Result<Run> {
    let mut run = Run::new();
    if let Some(style) = string_prop(props, "style", path)? {
        run = run.style(style);
    }
    if let Some(font) = string_prop(props, "font", path)? {
        run = run.fonts(
            RunFonts::new()
                .ascii(font)
                .hi_ansi(font)
                .east_asia(font)
                .cs(font),
        );
    }
    if let Some(fonts) = object_prop(props, "fonts", path)? {
        run = run.fonts(compile_run_fonts(fonts, &format!("{path}/fonts"))?);
    }
    if let Some(size) = number_prop(props, "size", path)? {
        run = run.size(to_half_points(size, &format!("{path}/size"))?);
    }
    if let Some(value) = bool_prop(props, "bold", path)? {
        run = if value {
            run.bold()
        } else {
            run.disable_bold()
        };
    }
    if let Some(value) = bool_prop(props, "italic", path)? {
        run = if value {
            run.italic()
        } else {
            run.disable_italic()
        };
    }
    let strike = bool_prop(props, "strike", path)?;
    let double_strike = bool_prop(props, "doubleStrike", path)?;
    if strike == Some(true) {
        run = run.strike();
    } else if double_strike == Some(true) {
        run = run.dstrike();
    }
    if strike == Some(false) {
        run.run_property = std::mem::take(&mut run.run_property).disable_strike();
    }
    if double_strike == Some(false) {
        run.run_property = std::mem::take(&mut run.run_property).disable_dstrike();
    }
    if bool_prop(props, "underline", path)?.unwrap_or(false) {
        run = run.underline("single");
    }
    if let Some(color) = string_prop(props, "color", path)? {
        run = run.color(color.to_ascii_uppercase());
    }
    if let Some(theme) = string_prop(props, "themeColor", path)? {
        run = run.theme_color(theme_color(theme, path)?);
    }
    if let Some(shade) = string_prop(props, "themeShade", path)? {
        run = run.theme_shade(shade.to_ascii_uppercase());
    }
    if let Some(tint) = string_prop(props, "themeTint", path)? {
        run = run.theme_tint(tint.to_ascii_uppercase());
    }
    if let Some(highlight) = string_prop(props, "highlight", path)? {
        run = run.highlight(highlight);
    }
    Ok(run)
}

fn theme_color(value: &str, path: &str) -> Result<ThemeColor> {
    match value {
        "dark1" => Ok(ThemeColor::Dark1),
        "light1" => Ok(ThemeColor::Light1),
        "dark2" => Ok(ThemeColor::Dark2),
        "light2" => Ok(ThemeColor::Light2),
        "accent1" => Ok(ThemeColor::Accent1),
        "accent2" => Ok(ThemeColor::Accent2),
        "accent3" => Ok(ThemeColor::Accent3),
        "accent4" => Ok(ThemeColor::Accent4),
        "accent5" => Ok(ThemeColor::Accent5),
        "accent6" => Ok(ThemeColor::Accent6),
        "hyperlink" => Ok(ThemeColor::Hyperlink),
        "followedHyperlink" => Ok(ThemeColor::FollowedHyperlink),
        "none" => Ok(ThemeColor::None),
        "background1" => Ok(ThemeColor::Background1),
        "text1" => Ok(ThemeColor::Text1),
        "background2" => Ok(ThemeColor::Background2),
        "text2" => Ok(ThemeColor::Text2),
        _ => Err(validation(
            path,
            format!("invalid `themeColor` value `{value}`"),
        )),
    }
}

fn compile_run_child(mut run: Run, child: &Child, entry_dir: &Path, path: &str) -> Result<Run> {
    match child {
        Child::String(value) => Ok(run.add_text(value)),
        Child::Number(value) => Ok(run.add_text(value.to_string())),
        Child::Node(node) => match node.kind {
            NodeKind::Text => {
                if let Some(value) = node.props.get("value") {
                    run = run.add_text(value_to_text(value, path)?);
                }
                for child in &node.children {
                    match child {
                        Child::String(value) => run = run.add_text(value),
                        Child::Number(value) => run = run.add_text(value.to_string()),
                        Child::Node(_) => {
                            return Err(validation(path, "Text cannot contain a component"));
                        }
                    }
                }
                Ok(run)
            }
            NodeKind::Break => {
                let kind = optional_enum(&node.props, "type", &["line", "page", "column"], path)?
                    .unwrap_or("line");
                Ok(run.add_break(match kind {
                    "page" => BreakType::Page,
                    "column" => BreakType::Column,
                    _ => BreakType::TextWrapping,
                }))
            }
            NodeKind::CarriageReturn => Ok(run.add_carriage_return()),
            NodeKind::NonBreakingSpace => Ok(run.add_text("\u{00a0}")),
            NodeKind::SoftHyphen => Ok(run.add_text("\u{00ad}")),
            NodeKind::NonBreakingHyphen => Ok(run.add_text("\u{2011}")),
            NodeKind::Image => Ok(run.add_image(compile_image(node, entry_dir, path)?)),
            NodeKind::Footnote => {
                Ok(run.add_footnote_reference(compile_footnote(node, entry_dir, path)?))
            }
            NodeKind::Tab => Ok(run.add_tab()),
            NodeKind::Symbol => Ok(run.add_sym(Sym::new(
                required_string(&node.props, "font", path)?,
                required_string(&node.props, "char", path)?,
            ))),
            NodeKind::PageReference => compile_page_reference(run, node, path),
            NodeKind::PositionalTab => compile_positional_tab(run, node, path),
            NodeKind::TocEntry => compile_toc_entry(run, node, path),
            _ => Err(validation(path, "unsupported Run child")),
        },
    }
}

fn add_field_to_paragraph(
    mut paragraph: Paragraph,
    node: &Node,
    entry_dir: &Path,
    path: &str,
) -> Result<Paragraph> {
    let instruction = field_instruction(node, path)?;
    let dirty = bool_prop(&node.props, "dirty", path)?.unwrap_or(true);
    paragraph = paragraph.add_run(
        Run::new()
            .add_field_char(FieldCharType::Begin, dirty)
            .add_instr_text(InstrText::Unsupported(instruction))
            .add_field_char(FieldCharType::Separate, false),
    );
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        let run = match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                compile_run(run, entry_dir, &child_path)?
            }
            child => compile_run_child(Run::new(), child, entry_dir, &child_path)?,
        };
        paragraph = paragraph.add_run(run);
    }
    if node.children.is_empty()
        && let Some(placeholder) = string_prop(&node.props, "placeholder", path)?
    {
        paragraph = paragraph.add_run(Run::new().add_text(placeholder));
    }
    Ok(paragraph.add_run(Run::new().add_field_char(FieldCharType::End, false)))
}

fn field_instruction(node: &Node, path: &str) -> Result<String> {
    let instruction = match node.kind {
        NodeKind::Field => required_string(&node.props, "instruction", path)?.to_owned(),
        NodeKind::DateField | NodeKind::TimeField => {
            let name = if node.kind == NodeKind::DateField {
                "DATE"
            } else {
                "TIME"
            };
            string_prop(&node.props, "format", path)?.map_or_else(
                || format!(" {name} "),
                |format| format!(" {name} \\@ \"{format}\" "),
            )
        }
        NodeKind::FileNameField => {
            if bool_prop(&node.props, "fullPath", path)?.unwrap_or(false) {
                " FILENAME \\p ".to_owned()
            } else {
                " FILENAME ".to_owned()
            }
        }
        NodeKind::AuthorField => " AUTHOR ".to_owned(),
        NodeKind::TitleField => " TITLE ".to_owned(),
        NodeKind::SubjectField => " SUBJECT ".to_owned(),
        NodeKind::SequenceField => sequence_field_instruction(node, path)?,
        NodeKind::ReferenceField => reference_field_instruction(node, path)?,
        NodeKind::MergeField => merge_field_instruction(node, path)?,
        NodeKind::DocumentPropertyField => {
            let name = required_string(&node.props, "name", path)?;
            format!(" DOCPROPERTY \"{name}\" ")
        }
        NodeKind::FormulaField => formula_field_instruction(node, path)?,
        NodeKind::IndexEntry => index_entry_instruction(node, path)?,
        _ => return Err(validation(path, "unsupported field component")),
    };
    Ok(instruction)
}

fn index_instruction(node: &Node, path: &str) -> Result<String> {
    let mut instruction = " INDEX".to_owned();
    if let Some(identifier) = string_prop(&node.props, "identifier", path)? {
        write!(instruction, " \\f \"{identifier}\"")
            .map_err(|error| validation(path, error.to_string()))?;
    }
    if let Some(columns) = node.props.get("columns").and_then(Value::as_u64) {
        write!(instruction, " \\c \"{columns}\"")
            .map_err(|error| validation(path, error.to_string()))?;
    }
    if bool_prop(&node.props, "runIn", path)?.unwrap_or(false) {
        instruction.push_str(" \\r");
    }
    instruction.push(' ');
    Ok(instruction)
}

fn index_entry_instruction(node: &Node, path: &str) -> Result<String> {
    let text = escape_index_entry_part(required_string(&node.props, "text", path)?);
    let subentry = string_prop(&node.props, "subentry", path)?
        .map(escape_index_entry_part)
        .map_or_else(String::new, |value| format!(":{value}"));
    let mut instruction = format!(" XE \"{text}{subentry}\"");
    if let Some(identifier) = string_prop(&node.props, "identifier", path)? {
        write!(instruction, " \\f \"{identifier}\"")
            .map_err(|error| validation(path, error.to_string()))?;
    }
    if bool_prop(&node.props, "boldPageNumber", path)?.unwrap_or(false) {
        instruction.push_str(" \\b");
    }
    if bool_prop(&node.props, "italicPageNumber", path)?.unwrap_or(false) {
        instruction.push_str(" \\i");
    }
    if let Some(bookmark) = string_prop(&node.props, "pageRangeBookmark", path)? {
        write!(instruction, " \\r {bookmark}")
            .map_err(|error| validation(path, error.to_string()))?;
    }
    if let Some(reference) = string_prop(&node.props, "crossReference", path)? {
        write!(instruction, " \\t \"{reference}\"")
            .map_err(|error| validation(path, error.to_string()))?;
    }
    instruction.push(' ');
    Ok(instruction)
}

fn escape_index_entry_part(value: &str) -> String {
    value.replace(':', "\\:")
}

fn merge_field_instruction(node: &Node, path: &str) -> Result<String> {
    let name = required_string(&node.props, "name", path)?;
    let preserve = if bool_prop(&node.props, "preserveFormatting", path)?.unwrap_or(true) {
        " \\* MERGEFORMAT"
    } else {
        ""
    };
    Ok(format!(" MERGEFIELD {name}{preserve} "))
}

fn formula_field_instruction(node: &Node, path: &str) -> Result<String> {
    let expression = required_string(&node.props, "expression", path)?;
    let format = string_prop(&node.props, "numberFormat", path)?
        .map_or_else(String::new, |format| format!(" \\# \"{format}\""));
    Ok(format!(" = {expression}{format} "))
}

fn sequence_field_instruction(node: &Node, path: &str) -> Result<String> {
    let identifier = required_string(&node.props, "identifier", path)?;
    sequence_instruction(identifier, node, path)
}

fn sequence_instruction(identifier: &str, node: &Node, path: &str) -> Result<String> {
    let format = optional_enum(
        &node.props,
        "format",
        &["arabic", "roman", "Roman", "alphabetic", "Alphabetic"],
        path,
    )?
    .unwrap_or("arabic");
    let switch = match format {
        "roman" => "roman",
        "Roman" => "ROMAN",
        "alphabetic" => "alphabetic",
        "Alphabetic" => "ALPHABETIC",
        _ => "ARABIC",
    };
    let restart = node
        .props
        .get("restart")
        .and_then(Value::as_u64)
        .map_or_else(String::new, |value| format!(" \\r {value}"));
    Ok(format!(" SEQ {identifier} \\* {switch}{restart} "))
}

fn reference_field_instruction(node: &Node, path: &str) -> Result<String> {
    let bookmark = required_string(&node.props, "bookmark", path)?;
    let hyperlink = if bool_prop(&node.props, "hyperlink", path)?.unwrap_or(true) {
        " \\h"
    } else {
        ""
    };
    let relative = if bool_prop(&node.props, "relativePosition", path)?.unwrap_or(false) {
        " \\p"
    } else {
        ""
    };
    Ok(format!(" REF {bookmark}{hyperlink}{relative} "))
}

fn compile_toc_entry(run: Run, node: &Node, path: &str) -> Result<Run> {
    let mut entry = InstrTC::new(required_string(&node.props, "text", path)?);
    if let Some(level) = node.props.get("level").and_then(Value::as_u64) {
        entry = entry.level(
            usize::try_from(level)
                .map_err(|_| validation(path, "TocEntry level is out of range"))?,
        );
    }
    if bool_prop(&node.props, "omitPageNumber", path)?.unwrap_or(false) {
        entry = entry.omits_page_number();
    }
    if let Some(identifier) = string_prop(&node.props, "identifier", path)? {
        entry = entry.item_type_identifier(identifier);
    }
    Ok(run.add_tc(entry))
}

fn compile_page_reference(run: Run, node: &Node, path: &str) -> Result<Run> {
    let mut instruction = InstrPAGEREF::new(required_string(&node.props, "bookmark", path)?);
    if bool_prop(&node.props, "hyperlink", path)?.unwrap_or(true) {
        instruction = instruction.hyperlink();
    }
    if bool_prop(&node.props, "relativePosition", path)?.unwrap_or(false) {
        instruction = instruction.relative_position();
    }
    let placeholder = string_prop(&node.props, "placeholder", path)?.unwrap_or("1");
    let dirty = bool_prop(&node.props, "dirty", path)?.unwrap_or(true);
    Ok(run
        .add_field_char(FieldCharType::Begin, dirty)
        .add_instr_text(InstrText::PAGEREF(instruction))
        .add_field_char(FieldCharType::Separate, false)
        .add_text(placeholder)
        .add_field_char(FieldCharType::End, false))
}

fn compile_positional_tab(run: Run, node: &Node, path: &str) -> Result<Run> {
    let alignment = match optional_enum(&node.props, "align", &["left", "center", "right"], path)?
        .unwrap_or("left")
    {
        "center" => PositionalTabAlignmentType::Center,
        "right" => PositionalTabAlignmentType::Right,
        _ => PositionalTabAlignmentType::Left,
    };
    let relative_to = match optional_enum(&node.props, "relativeTo", &["margin", "indent"], path)?
        .unwrap_or("margin")
    {
        "indent" => PositionalTabRelativeTo::Indent,
        _ => PositionalTabRelativeTo::Margin,
    };
    let leader = match optional_enum(
        &node.props,
        "leader",
        &["none", "dot", "heavy", "hyphen", "middleDot", "underscore"],
        path,
    )?
    .unwrap_or("none")
    {
        "dot" => TabLeaderType::Dot,
        "heavy" => TabLeaderType::Heavy,
        "hyphen" => TabLeaderType::Hyphen,
        "middleDot" => TabLeaderType::MiddleDot,
        "underscore" => TabLeaderType::Underscore,
        _ => TabLeaderType::None,
    };
    Ok(run.add_ptab(PositionalTab::new(alignment, relative_to, leader)))
}

fn compile_footnote(node: &Node, entry_dir: &Path, path: &str) -> Result<Footnote> {
    let mut paragraph = Paragraph::new();
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
        let run = match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                compile_run(run, entry_dir, &child_path)?
            }
            child => compile_run_child(Run::new(), child, entry_dir, &child_path)?,
        };
        paragraph = paragraph.add_run(run);
    }
    Ok(Footnote::new().add_content(paragraph))
}

fn compile_image(node: &Node, entry_dir: &Path, path: &str) -> Result<Pic> {
    let src = required_string(&node.props, "src", path)?;
    let source_path = Path::new(src);
    let source_path = if source_path.is_absolute() {
        source_path.to_path_buf()
    } else {
        entry_dir.join(source_path)
    };
    let png = load_image_as_png(&source_path)?;
    let width = required_number(&node.props, "width", path)?;
    let height = required_number(&node.props, "height", path)?;
    let mut picture = Pic::new_with_dimensions(png, 1, 1).size(
        to_emu(width, &format!("{path}/width"))?,
        to_emu(height, &format!("{path}/height"))?,
    );
    if node.props.contains_key("simplePosition") {
        let final_id = node
            .props
            .get("relationshipId")
            .and_then(Value::as_str)
            .unwrap_or(&picture.id);
        let allow_overlap = node.props.get("allowOverlap").and_then(Value::as_bool) == Some(true);
        picture.id = simple_position_marker(path, final_id, allow_overlap);
    }
    compile_image_layout(picture, node, path)
}

fn simple_position_marker(path: &str, final_id: &str, allow_overlap: bool) -> String {
    format!(
        "rIdDocxJsxSimplePosition_{}_{}_{}",
        u8::from(allow_overlap),
        hex_encode(path.as_bytes()),
        hex_encode(final_id.as_bytes())
    )
}

fn hex_encode(bytes: &[u8]) -> String {
    let mut encoded = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        let _ = write!(encoded, "{byte:02X}");
    }
    encoded
}

struct SimplePositionNormalization {
    final_relationship_id: String,
    allow_overlap: bool,
}

#[expect(
    clippy::too_many_lines,
    reason = "normalizing a generated DOCX package requires keeping marker, drawing, and relationship checks together"
)]
fn patch_simple_position_images(bytes: Vec<u8>) -> Result<Vec<u8>> {
    const MARKER_PREFIX: &[u8] = b"rIdDocxJsxSimplePosition_";
    if !bytes
        .windows(MARKER_PREFIX.len())
        .any(|window| window == MARKER_PREFIX)
    {
        return Ok(bytes);
    }

    let mut source = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|error| {
        Error::Compile(format!(
            "cannot reopen DOCX for simple-position normalization: {error}"
        ))
    })?;
    let mut targets = BTreeMap::new();
    for index in 0..source.len() {
        let mut file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        let name = file.name().to_owned();
        if !name.starts_with("word/")
            || name.contains("/_rels/")
            || !Path::new(&name)
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        {
            continue;
        }
        let mut xml = String::new();
        file.read_to_string(&mut xml)
            .map_err(|error| Error::Compile(format!("cannot read {name}: {error}")))?;
        for marker in find_simple_position_markers(&xml)? {
            let normalization = parse_simple_position_marker(&marker)?;
            if targets.insert(marker.clone(), normalization).is_some() {
                return Err(Error::Compile(format!(
                    "duplicate internal simple-position relationship marker `{marker}`; use distinct image component paths"
                )));
            }
        }
    }
    if targets.is_empty() {
        return Err(Error::Compile(
            "simple-position image markers were expected but not found in the DOCX package"
                .to_owned(),
        ));
    }

    let mut drawing_markers = BTreeSet::new();
    let mut relationship_markers = BTreeSet::new();
    let output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(output);
    for index in 0..source.len() {
        let mut file = source
            .by_index(index)
            .map_err(|error| Error::Compile(format!("cannot read DOCX entry: {error}")))?;
        if file.is_dir() {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
            continue;
        }
        let name = file.name().to_owned();
        if !name.starts_with("word/")
            || !Path::new(&name).extension().is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xml") || extension.eq_ignore_ascii_case("rels")
            })
        {
            writer
                .raw_copy_file(file)
                .map_err(|error| Error::Compile(format!("cannot copy DOCX entry: {error}")))?;
            continue;
        }
        let options = file.options();
        let mut xml = String::new();
        file.read_to_string(&mut xml)
            .map_err(|error| Error::Compile(format!("cannot read {name}: {error}")))?;
        if name.contains("/_rels/") {
            for (marker, target) in &targets {
                let relationship = format!("Id=\"{marker}\"");
                let occurrence_count = xml.match_indices(&relationship).count();
                if occurrence_count > 1 {
                    return Err(Error::Compile(format!(
                        "simple-position relationship `{marker}` occurs more than once in `{name}`; use a unique relationshipId"
                    )));
                }
                if occurrence_count == 1 {
                    let replacement = format!("Id=\"{}\"", target.final_relationship_id);
                    xml = xml.replace(&relationship, &replacement);
                    relationship_markers.insert(marker.clone());
                }
            }
        } else {
            let (patched, found) = normalize_simple_position_anchors(&xml, &targets, &name)?;
            xml = patched;
            drawing_markers.extend(found);
        }
        writer
            .start_file(&name, options)
            .map_err(|error| Error::Compile(format!("cannot rewrite {name}: {error}")))?;
        writer
            .write_all(xml.as_bytes())
            .map_err(|error| Error::Compile(format!("cannot write {name}: {error}")))?;
    }
    for marker in targets.keys() {
        if !drawing_markers.contains(marker) || !relationship_markers.contains(marker) {
            return Err(Error::Compile(format!(
                "simple-position image relationship `{marker}` is missing its drawing or package relationship; restore the image relationship and try again"
            )));
        }
    }
    writer
        .finish()
        .map(Cursor::into_inner)
        .map_err(|error| Error::Compile(format!("cannot finalize simple-position DOCX: {error}")))
}

fn find_simple_position_markers(xml: &str) -> Result<Vec<String>> {
    const PREFIX: &str = "rIdDocxJsxSimplePosition_";
    let mut markers = Vec::new();
    let mut search_from = 0;
    while let Some(relative_start) = xml[search_from..].find(PREFIX) {
        let start = search_from + relative_start;
        let value_start = start + PREFIX.len();
        let value_end = xml[value_start..]
            .find('"')
            .map(|offset| value_start + offset)
            .ok_or_else(|| {
                Error::Compile(
                    "unterminated simple-position relationship marker in generated XML".to_owned(),
                )
            })?;
        markers.push(xml[start..value_end].to_owned());
        search_from = value_end + 1;
    }
    Ok(markers)
}

fn parse_simple_position_marker(marker: &str) -> Result<SimplePositionNormalization> {
    let mut parts = marker
        .strip_prefix("rIdDocxJsxSimplePosition_")
        .unwrap_or_default()
        .split('_');
    let allow_overlap = match parts.next() {
        Some("0") => false,
        Some("1") => true,
        _ => {
            return Err(Error::Compile(
                "malformed internal simple-position marker; recompile the JSX source".to_owned(),
            ));
        }
    };
    let path = parts.next().unwrap_or_default();
    let relationship_id = parts.next().unwrap_or_default();
    if path.is_empty() || relationship_id.is_empty() || parts.next().is_some() {
        return Err(Error::Compile(
            "malformed internal simple-position marker; recompile the JSX source".to_owned(),
        ));
    }
    let relationship_id = hex_decode(relationship_id)?;
    if !is_xml_ncname(&relationship_id) {
        return Err(Error::Compile("simple-position relationship id is not a valid XML NCName; use letters, numbers, `_`, `-`, or `.`".to_owned()));
    }
    Ok(SimplePositionNormalization {
        final_relationship_id: relationship_id,
        allow_overlap,
    })
}

fn hex_decode(encoded: &str) -> Result<String> {
    if !encoded.len().is_multiple_of(2) {
        return Err(Error::Compile(
            "malformed internal simple-position relationship id".to_owned(),
        ));
    }
    let bytes = (0..encoded.len())
        .step_by(2)
        .map(|index| {
            u8::from_str_radix(&encoded[index..index + 2], 16).map_err(|_| {
                Error::Compile("malformed internal simple-position relationship id".to_owned())
            })
        })
        .collect::<Result<Vec<_>>>()?;
    String::from_utf8(bytes).map_err(|_| {
        Error::Compile("internal simple-position relationship id is not UTF-8".to_owned())
    })
}

fn normalize_simple_position_anchors(
    xml: &str,
    targets: &BTreeMap<String, SimplePositionNormalization>,
    part: &str,
) -> Result<(String, BTreeSet<String>)> {
    let mut xml = xml.to_owned();
    let mut found = BTreeSet::new();
    for (marker, target) in targets {
        let embed = format!("r:embed=\"{marker}\"");
        if !xml.contains(&embed) {
            continue;
        }
        let mut search_from = 0;
        let mut occurrences = 0;
        while let Some(relative_embed) = xml[search_from..].find(&embed) {
            let embed_start = search_from + relative_embed;
            let anchor_start = xml[..embed_start].rfind("<wp:anchor").ok_or_else(|| {
                Error::Compile(format!(
                    "image relationship `{marker}` in `{part}` is not inside a floating anchor"
                ))
            })?;
            let tag_end = xml[anchor_start..]
                .find('>')
                .map(|offset| anchor_start + offset)
                .ok_or_else(|| Error::Compile(format!("unterminated wp:anchor in `{part}`")))?;
            let anchor_end = xml[embed_start..]
                .find("</wp:anchor>")
                .map(|offset| embed_start + offset + "</wp:anchor>".len())
                .ok_or_else(|| {
                    Error::Compile(format!(
                        "unterminated wp:anchor for image `{marker}` in `{part}`"
                    ))
                })?;
            if !xml[anchor_start..anchor_end].contains(&embed) {
                return Err(Error::Compile(format!(
                    "image relationship `{marker}` is not inside its expected wp:anchor in `{part}`"
                )));
            }
            let start_tag = &xml[anchor_start..=tag_end];
            let start_tag = set_xml_attribute(start_tag, "simplePos", "1")?;
            let start_tag = set_xml_attribute(
                &start_tag,
                "allowOverlap",
                if target.allow_overlap { "1" } else { "0" },
            )?;
            xml.replace_range(anchor_start..=tag_end, &start_tag);
            let patched_embed = format!("r:embed=\"{}\"", target.final_relationship_id);
            let embed_start = anchor_start + start_tag.len();
            let Some(relative_embed) = xml[embed_start..].find(&embed) else {
                return Err(Error::Compile(format!(
                    "cannot locate simple-position image relationship `{marker}` after anchor normalization"
                )));
            };
            let embed_start = embed_start + relative_embed;
            xml.replace_range(embed_start..embed_start + embed.len(), &patched_embed);
            search_from = embed_start + patched_embed.len();
            occurrences += 1;
        }
        if occurrences != 1 {
            return Err(Error::Compile(format!(
                "expected one drawing reference for simple-position image `{marker}` in `{part}`, found {occurrences}"
            )));
        }
        found.insert(marker.clone());
    }
    Ok((xml, found))
}

fn set_xml_attribute(tag: &str, name: &str, value: &str) -> Result<String> {
    let attribute = format!("{name}=\"");
    if let Some(start) = tag.find(&attribute) {
        let value_start = start + attribute.len();
        let value_end = tag[value_start..]
            .find('"')
            .map(|offset| value_start + offset)
            .ok_or_else(|| {
                Error::Compile(format!("unterminated `{name}` attribute in generated XML"))
            })?;
        let mut output = tag.to_owned();
        output.replace_range(value_start..value_end, value);
        return Ok(output);
    }
    let insert_at = tag
        .rfind('>')
        .ok_or_else(|| Error::Compile("unterminated generated XML start tag".to_owned()))?;
    let mut output = tag.to_owned();
    output.insert_str(insert_at, &format!(" {name}=\"{value}\""));
    Ok(output)
}

fn is_xml_ncname(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes
        .next()
        .is_some_and(|first| first.is_ascii_alphabetic() || first == b'_')
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
}

fn load_image_as_png(source_path: &Path) -> Result<Vec<u8>> {
    let bytes = std::fs::read(source_path).map_err(|source| Error::Resource {
        path: source_path.to_path_buf(),
        source,
    })?;
    let image = image::load_from_memory(&bytes).map_err(|error| {
        Error::Compile(format!("invalid image {}: {error}", source_path.display()))
    })?;
    let mut png = Cursor::new(Vec::new());
    image
        .write_to(&mut png, ImageFormat::Png)
        .map_err(|error| {
            Error::Compile(format!(
                "cannot encode image {}: {error}",
                source_path.display()
            ))
        })?;
    Ok(png.into_inner())
}

fn compile_image_layout(mut picture: Pic, node: &Node, path: &str) -> Result<Pic> {
    if let Some(id) = string_prop(&node.props, "relationshipId", path)?
        && !node.props.contains_key("simplePosition")
    {
        picture = picture.id(id);
    }
    if let Some(angle) = node.props.get("rotate") {
        picture = picture.rotate(
            u16::try_from(
                angle
                    .as_u64()
                    .ok_or_else(|| validation(path, "`rotate` must be a non-negative integer"))?,
            )
            .map_err(|_| validation(path, "`rotate` must not exceed 65535"))?,
        );
    }
    if bool_prop(&node.props, "floating", path)? == Some(true) {
        picture = picture.floating();
    }
    if let Some(layout_in_cell) = bool_prop(&node.props, "layoutInCell", path)? {
        picture.layout_in_cell = layout_in_cell;
    }
    if bool_prop(&node.props, "allowOverlap", path)? == Some(true) {
        // docx-rs 0.4.22 reads `simple_pos` into the `allowOverlap` anchor
        // attribute. Set both fields so the attribute and wrapping agree.
        picture = picture.overlapping().simple_pos(true);
    }
    if let Some(position) = node.props.get("simplePosition") {
        let position = position.as_object().ok_or_else(|| {
            validation(
                path,
                "`simplePosition` must be an object with `x` and `y` coordinates",
            )
        })?;
        let x = position
            .get("x")
            .and_then(Value::as_i64)
            .and_then(|value| i32::try_from(value).ok())
            .ok_or_else(|| {
                validation(path, "`simplePosition.x` must be a signed 32-bit integer")
            })?;
        let y = position
            .get("y")
            .and_then(Value::as_i64)
            .and_then(|value| i32::try_from(value).ok())
            .ok_or_else(|| {
                validation(path, "`simplePosition.y` must be a signed 32-bit integer")
            })?;
        picture = picture.simple_pos(true);
        picture.simple_pos_x = x;
        picture.simple_pos_y = y;
    }
    if let Some(position) = compile_image_position(&node.props, "positionH", path)? {
        picture = picture.position_h(position);
    }
    if let Some(position) = compile_image_position(&node.props, "positionV", path)? {
        picture = picture.position_v(position);
    }
    picture = compile_image_relative_origins(picture, &node.props, path)?;
    picture = compile_image_distances(picture, &node.props, path)?;
    if node.props.contains_key("relativeHeight") {
        picture = picture.relative_height(
            u32::try_from(required_u64(&node.props, "relativeHeight", path)?)
                .map_err(|_| validation(path, "`relativeHeight` must not exceed 4294967295"))?,
        );
    }
    Ok(picture)
}

fn compile_image_relative_origins(
    mut picture: Pic,
    props: &Map<String, Value>,
    path: &str,
) -> Result<Pic> {
    if let Some(value) = optional_enum(
        props,
        "relativeFromH",
        &[
            "character",
            "column",
            "insideMargin",
            "leftMargin",
            "margin",
            "outsideMargin",
            "page",
            "rightMargin",
        ],
        path,
    )? {
        picture = picture.relative_from_h(match value {
            "character" => RelativeFromHType::Character,
            "column" => RelativeFromHType::Column,
            "insideMargin" => RelativeFromHType::InsideMargin,
            "leftMargin" => RelativeFromHType::LeftMargin,
            "outsideMargin" => RelativeFromHType::OutsizeMargin,
            "page" => RelativeFromHType::Page,
            "rightMargin" => RelativeFromHType::RightMargin,
            _ => RelativeFromHType::Margin,
        });
    }
    if let Some(value) = optional_enum(
        props,
        "relativeFromV",
        &[
            "bottomMargin",
            "insideMargin",
            "line",
            "margin",
            "outsideMargin",
            "page",
            "paragraph",
            "topMargin",
        ],
        path,
    )? {
        picture = picture.relative_from_v(match value {
            "bottomMargin" => RelativeFromVType::BottomMargin,
            "insideMargin" => RelativeFromVType::InsideMargin,
            "line" => RelativeFromVType::Line,
            "outsideMargin" => RelativeFromVType::OutsizeMargin,
            "page" => RelativeFromVType::Page,
            "paragraph" => RelativeFromVType::Paragraph,
            "topMargin" => RelativeFromVType::TopMargin,
            _ => RelativeFromVType::Margin,
        });
    }
    Ok(picture)
}

fn compile_image_distances(
    mut picture: Pic,
    props: &Map<String, Value>,
    path: &str,
) -> Result<Pic> {
    for (key, setter) in [
        ("distanceTop", Pic::dist_t as fn(Pic, i32) -> Pic),
        ("distanceBottom", Pic::dist_b),
        ("distanceLeft", Pic::dist_l),
        ("distanceRight", Pic::dist_r),
    ] {
        if let Some(value) = number_prop(props, key, path)? {
            picture = setter(picture, to_emu_i32(value, &format!("{path}/{key}"))?);
        }
    }
    Ok(picture)
}

fn compile_image_position(
    props: &Map<String, Value>,
    key: &str,
    path: &str,
) -> Result<Option<DrawingPosition>> {
    let Some(value) = props.get(key) else {
        return Ok(None);
    };
    if let Some(offset) = value.as_f64().filter(|offset| offset.is_finite()) {
        return Ok(Some(DrawingPosition::Offset(to_emu_i32(
            offset,
            &format!("{path}/{key}"),
        )?)));
    }
    let alignment = value
        .as_str()
        .ok_or_else(|| validation(path, format!("`{key}` must be a point offset or alignment")))?;
    Ok(Some(DrawingPosition::Align(match alignment {
        "left" => PicAlign::Left,
        "right" => PicAlign::Right,
        "top" => PicAlign::Top,
        "bottom" => PicAlign::Bottom,
        "center" => PicAlign::Center,
        _ => {
            return Err(validation(
                path,
                format!("invalid `{key}` value `{alignment}`"),
            ));
        }
    })))
}

fn compile_table(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Table> {
    let mut rows = Vec::with_capacity(node.children.len());
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(row) = child else {
            return Err(validation(path, "Table only accepts TableRow"));
        };
        rows.push(compile_row(
            row,
            entry_dir,
            &format!("{path}/TableRow[{index}]"),
            context,
        )?);
    }
    let mut table = Table::new(rows);
    if let Some(style) = string_prop(&node.props, "style", path)? {
        table = table.style(style);
    }
    if let Some(indent) = number_prop(&node.props, "indent", path)? {
        table = table.indent(to_twips_i32(indent, &format!("{path}/indent"))?);
    }
    if let Some(margins) = node.props.get("margins") {
        let [top, right, bottom, left] = parse_box_margins(margins, path)?;
        table = table.margins(TableCellMargins::new().margin(top, right, bottom, left));
    }
    if let Some(position) = node.props.get("position") {
        table = table.position(parse_table_position(position, path)?);
    }
    if let Some(width) = number_prop(&node.props, "width", path)? {
        table = table.width(
            to_twips_usize(width, &format!("{path}/width"))?,
            WidthType::Dxa,
        );
    }
    if let Some(percent) = number_prop(&node.props, "widthPercent", path)? {
        table = table.width(percent_to_fiftieths(percent, path)?, WidthType::Pct);
    }
    if let Some(align) = optional_enum(&node.props, "align", &["left", "center", "right"], path)? {
        table = table.align(match align {
            "center" => TableAlignmentType::Center,
            "right" => TableAlignmentType::Right,
            _ => TableAlignmentType::Left,
        });
    }
    if let Some(layout) = optional_enum(&node.props, "layout", &["auto", "fixed"], path)? {
        table = table.layout(if layout == "fixed" {
            TableLayoutType::Fixed
        } else {
            TableLayoutType::Autofit
        });
    }
    if let Some(widths) = node.props.get("columnWidths") {
        let widths = widths
            .as_array()
            .ok_or_else(|| validation(path, "`columnWidths` must be an array"))?
            .iter()
            .enumerate()
            .map(|(index, value)| {
                let width = value.as_f64().ok_or_else(|| {
                    validation(path, format!("columnWidths[{index}] must be a number"))
                })?;
                if width <= 0.0 {
                    return Err(validation(
                        path,
                        format!("columnWidths[{index}] must be positive"),
                    ));
                }
                to_twips_usize(width, path)
            })
            .collect::<Result<Vec<_>>>()?;
        table = table.set_grid(widths);
    }
    if let Some(border) = node.props.get("border") {
        table = table.set_borders(compile_table_borders(border, path)?);
    }
    Ok(table)
}

fn compile_row(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<TableRow> {
    let mut cells = Vec::with_capacity(node.children.len());
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(cell) = child else {
            return Err(validation(path, "TableRow only accepts TableCell"));
        };
        cells.push(compile_cell(
            cell,
            entry_dir,
            &format!("{path}/TableCell[{index}]"),
            context,
        )?);
    }
    let mut row = TableRow::new(cells);
    if let Some(height) = number_prop(&node.props, "height", path)? {
        let height = to_twips_i32(height, path)?
            .to_f32()
            .ok_or_else(|| validation(path, "row height is out of range"))?;
        row = row.row_height(height);
    }
    if let Some(rule) = optional_enum(
        &node.props,
        "heightRule",
        &["auto", "atLeast", "exact"],
        path,
    )? {
        row = row.height_rule(match rule {
            "auto" => HeightRule::Auto,
            "exact" => HeightRule::Exact,
            _ => HeightRule::AtLeast,
        });
    }
    if bool_prop(&node.props, "cantSplit", path)?.unwrap_or(false) {
        row = row.cant_split();
    }
    if let Some(revision) = node.props.get("inserted") {
        row = row.insert(compile_row_insert(revision, path)?);
    }
    if let Some(revision) = node.props.get("deleted") {
        row = row.delete(compile_row_delete(revision, path)?);
    }
    Ok(row)
}

fn compile_cell(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<TableCell> {
    let mut cell = TableCell::new();
    if let Some(width) = number_prop(&node.props, "width", path)? {
        cell = cell.width(to_twips_usize(width, path)?, WidthType::Dxa);
    }
    if let Some(width) = number_prop(&node.props, "widthPercent", path)? {
        cell = cell.width(percent_to_fiftieths(width, path)?, WidthType::Pct);
    }
    if let Some(span) = node.props.get("colSpan") {
        let span = span
            .as_u64()
            .filter(|span| *span > 0)
            .and_then(|span| usize::try_from(span).ok())
            .ok_or_else(|| validation(path, "`colSpan` must be a positive integer"))?;
        cell = cell.grid_span(span);
    }
    if let Some(align) = optional_enum(
        &node.props,
        "verticalAlign",
        &["top", "center", "bottom"],
        path,
    )? {
        cell = cell.vertical_align(match align {
            "center" => VAlignType::Center,
            "bottom" => VAlignType::Bottom,
            _ => VAlignType::Top,
        });
    }
    if let Some(color) = string_prop(&node.props, "shading", path)? {
        cell = cell.shading(Shading::new().fill(color.to_ascii_uppercase()));
    }
    if let Some(merge) =
        optional_enum(&node.props, "verticalMerge", &["restart", "continue"], path)?
    {
        cell = cell.vertical_merge(if merge == "restart" {
            VMergeType::Restart
        } else {
            VMergeType::Continue
        });
    }
    if let Some(direction) = optional_enum(
        &node.props,
        "textDirection",
        &[
            "lr", "lrV", "rl", "rlV", "tb", "tbV", "tbRlV", "tbRl", "btLr", "lrTbV",
        ],
        path,
    )? {
        let direction = direction.parse::<TextDirectionType>().map_err(|_| {
            validation(path, format!("invalid `textDirection` value `{direction}`"))
        })?;
        cell = cell.text_direction(direction);
    }
    if let Some(margins) = node.props.get("margins") {
        let [top, right, bottom, left] = parse_box_margins(margins, path)?;
        cell.property = cell.property.margins(
            CellMargins::new()
                .margin_top(top, WidthType::Dxa)
                .margin_right(right, WidthType::Dxa)
                .margin_bottom(bottom, WidthType::Dxa)
                .margin_left(left, WidthType::Dxa),
        );
    }
    if let Some(border) = node.props.get("border") {
        cell = cell.set_borders(compile_cell_borders(border, path)?);
    }
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(child) = child else {
            return Err(validation(
                path,
                "TableCell only accepts Paragraph or Table",
            ));
        };
        let child_path = format!("{path}/{}[{index}]", child.kind.name());
        cell = match child.kind {
            NodeKind::Paragraph => {
                cell.add_paragraph(compile_paragraph(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Caption => {
                cell.add_paragraph(compile_caption(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Table => {
                cell.add_table(compile_table(child, entry_dir, &child_path, context)?)
            }
            NodeKind::List => add_list_to_cell(cell, child, entry_dir, &child_path, context)?,
            NodeKind::TableOfContents => cell.add_table_of_contents(compile_table_of_contents(
                child,
                &child_path,
                entry_dir,
                context,
            )?),
            NodeKind::ContentControl => cell.add_structured_data_tag(compile_content_control(
                child,
                entry_dir,
                &child_path,
            )?),
            _ => return Err(validation(&child_path, "unsupported TableCell child")),
        };
    }
    Ok(cell)
}

fn attach_header_to_section(
    section: Section,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Section> {
    let header = compile_header(node, entry_dir, path, context)?;
    Ok(
        match optional_enum(&node.props, "type", &["default", "first", "even"], path)?
            .unwrap_or("default")
        {
            "first" => section.first_header(header).title_pg(),
            "even" => section.even_header(header),
            _ => section.header(header),
        },
    )
}

fn attach_footer_to_section(
    section: Section,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Section> {
    let footer = compile_footer(node, entry_dir, path, context)?;
    Ok(
        match optional_enum(&node.props, "type", &["default", "first", "even"], path)?
            .unwrap_or("default")
        {
            "first" => section.first_footer(footer).title_pg(),
            "even" => section.even_footer(footer),
            _ => section.footer(footer),
        },
    )
}

fn attach_header(
    docx: Docx,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Docx> {
    let header = compile_header(node, entry_dir, path, context)?;
    Ok(
        match optional_enum(&node.props, "type", &["default", "first", "even"], path)?
            .unwrap_or("default")
        {
            "first" => docx.first_header(header).title_pg(),
            "even" => docx.even_header(header),
            _ => docx.header(header),
        },
    )
}

fn attach_footer(
    docx: Docx,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Docx> {
    let footer = compile_footer(node, entry_dir, path, context)?;
    Ok(
        match optional_enum(&node.props, "type", &["default", "first", "even"], path)?
            .unwrap_or("default")
        {
            "first" => docx.first_footer(footer).title_pg(),
            "even" => docx.even_footer(footer),
            _ => docx.footer(footer),
        },
    )
}

fn compile_header(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Header> {
    let mut header = Header::new();
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(child) = child else {
            return Err(validation(path, "Header only accepts structural children"));
        };
        let child_path = format!("{path}/{}[{index}]", child.kind.name());
        match child.kind {
            NodeKind::Paragraph => {
                header = header.add_paragraph(compile_paragraph(
                    child,
                    entry_dir,
                    &child_path,
                    context,
                )?);
            }
            NodeKind::Caption => {
                header =
                    header.add_paragraph(compile_caption(child, entry_dir, &child_path, context)?);
            }
            NodeKind::Table => {
                header = header.add_table(compile_table(child, entry_dir, &child_path, context)?);
            }
            NodeKind::List => {
                for paragraph in compile_list(child, entry_dir, &child_path, context)? {
                    header = header.add_paragraph(paragraph);
                }
            }
            NodeKind::ContentControl => {
                let control =
                    compile_block_content_control(child, entry_dir, &child_path, context)?;
                header.has_numbering |= control.has_numbering;
                header
                    .children
                    .push(HeaderChild::StructuredDataTag(Box::new(control)));
            }
            _ => return Err(validation(&child_path, "unsupported Header child")),
        }
    }
    Ok(header)
}

fn compile_footer(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Footer> {
    let mut footer = Footer::new();
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(child) = child else {
            return Err(validation(path, "Footer only accepts structural children"));
        };
        let child_path = format!("{path}/{}[{index}]", child.kind.name());
        match child.kind {
            NodeKind::Paragraph => {
                footer = footer.add_paragraph(compile_paragraph(
                    child,
                    entry_dir,
                    &child_path,
                    context,
                )?);
            }
            NodeKind::Caption => {
                footer =
                    footer.add_paragraph(compile_caption(child, entry_dir, &child_path, context)?);
            }
            NodeKind::Table => {
                footer = footer.add_table(compile_table(child, entry_dir, &child_path, context)?);
            }
            NodeKind::List => {
                for paragraph in compile_list(child, entry_dir, &child_path, context)? {
                    footer = footer.add_paragraph(paragraph);
                }
            }
            NodeKind::ContentControl => {
                let control =
                    compile_block_content_control(child, entry_dir, &child_path, context)?;
                footer.has_numbering |= control.has_numbering;
                footer
                    .children
                    .push(FooterChild::StructuredDataTag(Box::new(control)));
            }
            _ => return Err(validation(&child_path, "unsupported Footer child")),
        }
    }
    Ok(footer)
}

fn compile_table_of_contents(
    node: &Node,
    path: &str,
    entry_dir: &Path,
    context: &mut CompileContext,
) -> Result<TableOfContents> {
    let start = usize::try_from(
        node.props
            .get("startLevel")
            .and_then(Value::as_u64)
            .unwrap_or(1),
    )
    .map_err(|_| validation(path, "`startLevel` is out of range"))?;
    let end = usize::try_from(
        node.props
            .get("endLevel")
            .and_then(Value::as_u64)
            .unwrap_or(3),
    )
    .map_err(|_| validation(path, "`endLevel` is out of range"))?;
    let toc = TableOfContents::new()
        .heading_styles_range(start, end)
        .auto();
    configure_index(toc, node, path, entry_dir, context)
}

fn compile_table_of_figures(
    node: &Node,
    path: &str,
    entry_dir: &Path,
    context: &mut CompileContext,
) -> Result<TableOfContents> {
    let label = required_string(&node.props, "label", path)?;
    let instr = if bool_prop(&node.props, "includeLabelAndNumber", path)?.unwrap_or(true) {
        InstrToC::new().caption_label_including_numbers(label)
    } else {
        InstrToC::new().caption_label(label)
    };
    let instr = if let Some(separator) = string_prop(&node.props, "separator", path)? {
        instr.sequence_and_page_numbers_separator(separator)
    } else {
        instr
    };
    let mut toc = TableOfContents::new();
    toc.instr = instr;
    configure_index(toc, node, path, entry_dir, context)
}

fn compile_table_of_entries(
    node: &Node,
    path: &str,
    entry_dir: &Path,
    context: &mut CompileContext,
) -> Result<TableOfContents> {
    let identifier = required_string(&node.props, "identifier", path)?;
    let mut toc = TableOfContents::new();
    toc.instr = InstrToC::new().tc_field_identifier(Some(identifier.to_owned()));
    configure_index(toc, node, path, entry_dir, context)
}

fn configure_index(
    mut toc: TableOfContents,
    node: &Node,
    path: &str,
    entry_dir: &Path,
    context: &mut CompileContext,
) -> Result<TableOfContents> {
    toc = toc.auto();
    if let Some(range) = node
        .props
        .get("tcFieldLevelRange")
        .and_then(Value::as_object)
    {
        let start = usize::try_from(range["start"].as_u64().unwrap_or_default())
            .map_err(|_| validation(path, "`tcFieldLevelRange.start` is out of range"))?;
        let end = usize::try_from(range["end"].as_u64().unwrap_or_default())
            .map_err(|_| validation(path, "`tcFieldLevelRange.end` is out of range"))?;
        toc.instr = toc.instr.tc_field_level_range(start, end);
    }
    if let Some(range) = node
        .props
        .get("omitPageNumbersLevelRange")
        .and_then(Value::as_object)
    {
        let start = usize::try_from(range["start"].as_u64().unwrap_or_default())
            .map_err(|_| validation(path, "`omitPageNumbersLevelRange.start` is out of range"))?;
        let end = usize::try_from(range["end"].as_u64().unwrap_or_default())
            .map_err(|_| validation(path, "`omitPageNumbersLevelRange.end` is out of range"))?;
        toc.instr = toc.instr.omit_page_numbers_level_range(start, end);
    }
    if let Some(bookmark) = string_prop(&node.props, "entryBookmarkName", path)? {
        toc.instr = toc.instr.entry_bookmark_name(bookmark);
    }
    if let Some(separator) = string_prop(&node.props, "entryAndPageNumberSeparator", path)? {
        toc.instr = toc.instr.entry_and_page_number_separator(separator);
    }
    if let Some(separator) = string_prop(&node.props, "sequenceAndPageNumbersSeparator", path)? {
        toc.instr = toc.instr.sequence_and_page_numbers_separator(separator);
    }
    if let Some(identifier) = string_prop(&node.props, "seqFieldIdentifierForPrefix", path)? {
        toc.instr = toc.instr.seq_field_identifier_for_prefix(identifier);
    }
    if let Some(styles) = node.props.get("stylesWithLevels").and_then(Value::as_array) {
        for style in styles {
            let style_name = style
                .get("style")
                .and_then(Value::as_str)
                .ok_or_else(|| validation(path, "invalid stylesWithLevels style"))?;
            let level = style
                .get("level")
                .and_then(Value::as_u64)
                .ok_or_else(|| validation(path, "invalid stylesWithLevels level"))?;
            let level = usize::try_from(level)
                .map_err(|_| validation(path, "stylesWithLevels level is out of range"))?;
            toc.instr = toc
                .instr
                .add_style_with_level(StyleWithLevel::new(style_name, level));
        }
    }
    for (property, apply) in [
        (
            "preserveTab",
            InstrToC::preserve_tab as fn(InstrToC) -> InstrToC,
        ),
        ("preserveNewLine", InstrToC::preserve_new_line),
        (
            "useAppliedParagraphLineLevel",
            InstrToC::use_applied_paragraph_line_level,
        ),
        (
            "hideTabAndPageNumbersInWebview",
            InstrToC::hide_tab_and_page_numbers_in_webview,
        ),
    ] {
        if bool_prop(&node.props, property, path)?.unwrap_or(false) {
            toc.instr = apply(toc.instr);
        }
    }
    if bool_prop(&node.props, "hyperlinks", path)?.unwrap_or(true) {
        toc = toc.hyperlink();
    }
    if bool_prop(&node.props, "dirty", path)?.unwrap_or(true) {
        toc = toc.dirty();
    }
    if let Some(alias) = string_prop(&node.props, "alias", path)? {
        toc = toc.alias(alias);
    }
    add_index_items(toc, node, path, entry_dir, context)
}

fn add_index_items(
    mut toc: TableOfContents,
    node: &Node,
    path: &str,
    entry_dir: &Path,
    context: &mut CompileContext,
) -> Result<TableOfContents> {
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(item) = child else {
            continue;
        };
        if matches!(item.kind, NodeKind::TocBefore | NodeKind::TocAfter) {
            for (paragraph_index, child) in item.children.iter().enumerate() {
                let Child::Node(paragraph) = child else {
                    continue;
                };
                let paragraph_path =
                    format!("{path}/{}/Paragraph[{paragraph_index}]", item.kind.name());
                let compiled = compile_paragraph(paragraph, entry_dir, &paragraph_path, context)?;
                toc = if item.kind == NodeKind::TocBefore {
                    toc.add_before_paragraph(compiled)
                } else {
                    toc.add_after_paragraph(compiled)
                };
            }
            continue;
        }
        if item.kind != NodeKind::TocItem {
            continue;
        }
        let item_path = format!("{path}/TocItem[{index}]");
        let text = required_string(&item.props, "text", &item_path)?;
        let bookmark = required_string(&item.props, "bookmark", &item_path)?;
        let level = item.props.get("level").and_then(Value::as_u64).unwrap_or(1);
        let level = usize::try_from(level)
            .map_err(|_| validation(&item_path, "`level` is out of range"))?;
        let marker = toc_item_marker(context);
        let mut compiled = TableOfContentsItem::new()
            .text(format!("{marker}{text}"))
            .toc_key(bookmark)
            .level(level);
        if let Some(page_ref) = string_prop(&item.props, "pageRef", &item_path)? {
            compiled = compiled.page_ref(page_ref);
        }
        let tab_position = number_prop(&item.props, "tabPosition", &item_path)?.unwrap_or(420.0);
        let tab_twips = to_twips_u32(tab_position, &item_path)?;
        context.toc_item_tab_patches.push((marker, tab_twips));
        toc = toc.add_item(compiled);
    }
    Ok(toc)
}

fn add_list_to_cell(
    mut cell: TableCell,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<TableCell> {
    for paragraph in compile_list(node, entry_dir, path, context)? {
        cell = cell.add_paragraph(paragraph);
    }
    Ok(cell)
}

fn compile_list(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Vec<Paragraph>> {
    let list_type =
        optional_enum(&node.props, "type", &["bullet", "ordered"], path)?.unwrap_or("bullet");
    let start = usize::try_from(node.props.get("start").and_then(Value::as_u64).unwrap_or(1))
        .map_err(|_| validation(path, "`start` is out of range"))?;
    context.next_numbering_id += 1;
    let id = context.next_numbering_id;
    let mut abstract_numbering = AbstractNumbering::new(id);
    let style_link = string_prop(&node.props, "styleLink", path)?.map(ToOwned::to_owned);
    let num_style_link = string_prop(&node.props, "numStyleLink", path)?.map(ToOwned::to_owned);
    abstract_numbering.style_link.clone_from(&style_link);
    abstract_numbering
        .num_style_link
        .clone_from(&num_style_link);
    let linked_definition = style_link.is_some() || num_style_link.is_some();
    if linked_definition {
        context
            .numbering_style_links
            .push((id, style_link, num_style_link));
    }
    if let Some(multi_level_type) = string_prop(&node.props, "multiLevelType", path)? {
        abstract_numbering.multi_level_type = Some(multi_level_type.to_owned());
    }
    let explicit = node.props.get("levels").and_then(Value::as_array);
    let definition_level_count =
        explicit.map_or_else(|| if linked_definition { 0 } else { 9 }, Vec::len);
    for index in 0..definition_level_count {
        let spec = explicit.and_then(|levels| levels.get(index));
        abstract_numbering = abstract_numbering.add_level(compile_list_level(
            spec,
            index,
            list_type,
            start,
            &format!("{path}/levels[{index}]"),
        )?);
    }
    let (numbering, mut level_count) =
        compile_list_overrides(node, id, list_type, start, definition_level_count, path)?;
    if linked_definition && level_count == 0 {
        level_count = 9;
    }
    context.numberings.push((abstract_numbering, numbering));
    let mut paragraphs = Vec::with_capacity(node.children.len());
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(item) = child else {
            return Err(validation(path, "List only accepts ListItem"));
        };
        let item_path = format!("{path}/ListItem[{index}]");
        let level = usize::try_from(item.props.get("level").and_then(Value::as_u64).unwrap_or(0))
            .map_err(|_| validation(&item_path, "`level` is out of range"))?;
        if level >= level_count {
            return Err(validation(
                &item_path,
                "`level` is outside the list `levels` range",
            ));
        }
        let paragraph = Paragraph::new().numbering(NumberingId::new(id), IndentLevel::new(level));
        let (mut paragraph, has_character_indentation) =
            register_paragraph_character_indentation(paragraph, item, context);
        if has_character_indentation {
            paragraph = paragraph.indent(None, None, None, None);
        }
        if let Some(style) = string_prop(&item.props, "style", &item_path)? {
            paragraph = paragraph.style(style);
        }
        for (child_index, child) in item.children.iter().enumerate() {
            let child_path = format!("{item_path}/child[{child_index}]");
            let run = match child {
                Child::Node(run) if run.kind == NodeKind::Run => {
                    compile_run(run, entry_dir, &child_path)?
                }
                child => compile_run_child(Run::new(), child, entry_dir, &child_path)?,
            };
            paragraph = paragraph.add_run(run);
        }
        paragraphs.push(paragraph);
    }
    Ok(paragraphs)
}

fn compile_list_overrides(
    node: &Node,
    id: usize,
    list_type: &str,
    default_start: usize,
    definition_level_count: usize,
    path: &str,
) -> Result<(Numbering, usize)> {
    let mut numbering = Numbering::new(id, id);
    let mut level_count = definition_level_count;
    if let Some(overrides) = node.props.get("overrides").and_then(Value::as_array) {
        for (index, value) in overrides.iter().enumerate() {
            let properties = value
                .as_object()
                .ok_or_else(|| validation(path, "list override must be an object"))?;
            let level_index = properties
                .get("level")
                .and_then(Value::as_u64)
                .and_then(|value| usize::try_from(value).ok())
                .ok_or_else(|| {
                    validation(path, "override `level` must be an integer from 0 to 8")
                })?;
            let mut override_level = LevelOverride::new(level_index);
            if let Some(start) = properties.get("start").and_then(Value::as_u64) {
                override_level = override_level.start(
                    usize::try_from(start)
                        .map_err(|_| validation(path, "override `start` is out of range"))?,
                );
            }
            if let Some(definition) = properties.get("definition") {
                let replacement = compile_list_level(
                    Some(definition),
                    level_index,
                    list_type,
                    default_start,
                    &format!("{path}/overrides[{index}]/definition"),
                )?;
                override_level = override_level.level(replacement);
                level_count = level_count.max(level_index + 1);
            }
            numbering = numbering.add_override(override_level);
        }
    }
    Ok((numbering, level_count))
}

fn compile_list_level(
    spec: Option<&Value>,
    index: usize,
    list_type: &str,
    default_start: usize,
    path: &str,
) -> Result<Level> {
    let props = spec.and_then(Value::as_object);
    let format = props
        .and_then(|value| value.get("format"))
        .and_then(Value::as_str)
        .unwrap_or(if list_type == "ordered" {
            "decimal"
        } else {
            "bullet"
        });
    let text = props
        .and_then(|value| value.get("text"))
        .and_then(Value::as_str)
        .map_or_else(
            || {
                if list_type == "ordered" {
                    format!("%{}.", index + 1)
                } else {
                    "•".to_owned()
                }
            },
            ToOwned::to_owned,
        );
    let start = props
        .and_then(|value| value.get("start"))
        .and_then(Value::as_u64)
        .and_then(|value| usize::try_from(value).ok())
        .unwrap_or(default_start);
    let align = props
        .and_then(|value| value.get("align"))
        .and_then(Value::as_str)
        .unwrap_or("left");
    let mut level = Level::new(
        index,
        Start::new(start),
        NumberFormat::new(format),
        LevelText::new(text),
        LevelJc::new(align),
    );
    if let Some(suffix) = props
        .and_then(|value| value.get("suffix"))
        .and_then(Value::as_str)
    {
        level = level.suffix(match suffix {
            "space" => LevelSuffixType::Space,
            "nothing" => LevelSuffixType::Nothing,
            _ => LevelSuffixType::Tab,
        });
    }
    if let Some(style) = props
        .and_then(|value| value.get("paragraphStyle"))
        .and_then(Value::as_str)
    {
        level = level.paragraph_style(style);
    }
    if let Some(restart) = props
        .and_then(|value| value.get("restart"))
        .and_then(Value::as_u64)
        .and_then(|value| u32::try_from(value).ok())
    {
        level = level.level_restart(restart);
    }
    if props
        .and_then(|value| value.get("legal"))
        .and_then(Value::as_bool)
        == Some(true)
    {
        level = level.is_lgl();
    }
    level = compile_list_level_indent(level, props, path)?;
    compile_list_level_run(level, props, path)
}

fn compile_list_level_indent(
    level: Level,
    props: Option<&Map<String, Value>>,
    path: &str,
) -> Result<Level> {
    let Some(props) = props else {
        return default_list_level_indent(level);
    };
    let has_indent = ["indentLeft", "indentRight", "hanging", "firstLine"]
        .iter()
        .any(|key| props.contains_key(*key));
    if !has_indent {
        return default_list_level_indent(level);
    }
    let left = optional_twips_i32(props, "indentLeft", path)?;
    let right = optional_twips_i32(props, "indentRight", path)?;
    let special = if let Some(value) = number_prop(props, "firstLine", path)? {
        Some(SpecialIndentType::FirstLine(to_twips_i32(value, path)?))
    } else if let Some(value) = number_prop(props, "hanging", path)? {
        Some(SpecialIndentType::Hanging(to_twips_i32(value, path)?))
    } else {
        None
    };
    Ok(level.indent(left, special, right, None))
}

fn default_list_level_indent(level: Level) -> Result<Level> {
    let index = i32::try_from(level.level + 1)
        .map_err(|_| validation("List", "list indent is out of range"))?;
    Ok(level.indent(
        Some(index * 720),
        Some(SpecialIndentType::Hanging(360)),
        None,
        None,
    ))
}

fn compile_list_level_run(
    mut level: Level,
    props: Option<&Map<String, Value>>,
    path: &str,
) -> Result<Level> {
    let Some(props) = props else {
        return Ok(level);
    };
    if let Some(font) = string_prop(props, "font", path)? {
        level = level.fonts(
            RunFonts::new()
                .ascii(font)
                .hi_ansi(font)
                .east_asia(font)
                .cs(font),
        );
    }
    if let Some(fonts) = object_prop(props, "fonts", path)? {
        level = level.fonts(compile_run_fonts(fonts, &format!("{path}/fonts"))?);
    }
    if let Some(size) = number_prop(props, "size", path)? {
        level = level.size(to_half_points(size, path)?);
    }
    if let Some(color) = string_prop(props, "color", path)? {
        level = level.color(color.to_ascii_uppercase());
    }
    if let Some(value) = string_prop(props, "highlight", path)? {
        level = level.highlight(value);
    }
    if let Some(value) = bool_prop(props, "bold", path)? {
        level = if value {
            level.bold()
        } else {
            level.disable_bold()
        };
    }
    if let Some(value) = bool_prop(props, "italic", path)? {
        level = if value {
            level.italic()
        } else {
            level.disable_italic()
        };
    }
    if let Some(value) = bool_prop(props, "strike", path)? {
        level = if value {
            level.strike()
        } else {
            level.disable_strike()
        };
    }
    if let Some(value) = bool_prop(props, "doubleStrike", path)? {
        level = if value {
            level.dstrike()
        } else {
            level.disable_dstrike()
        };
    }
    if let Some(value) = string_prop(props, "underline", path)? {
        level = level.underline(value);
    }
    if bool_prop(props, "hidden", path)? == Some(true) {
        level = level.vanish();
    }
    if let Some(spacing) = number_prop(props, "characterSpacing", path)? {
        level = level.spacing(to_twips_i32(spacing, path)?);
    }
    Ok(level)
}

#[derive(Clone)]
struct BorderSpec {
    border_type: BorderType,
    size: usize,
    space: usize,
    color: String,
}

fn parse_border(value: &Value, path: &str) -> Result<BorderSpec> {
    parse_border_spec(value, path, false)
}

fn parse_border_spec(value: &Value, path: &str, allow_space: bool) -> Result<BorderSpec> {
    let object = value
        .as_object()
        .ok_or_else(|| validation(path, "`border` must be an object"))?;
    for key in object.keys() {
        if !(["style", "size", "color"].contains(&key.as_str()) || allow_space && key == "space") {
            return Err(validation(path, format!("unknown border property `{key}`")));
        }
    }
    let style = string_prop(object, "style", path)?.unwrap_or("single");
    let size_pt = number_prop(object, "size", path)?.unwrap_or(0.5);
    if size_pt < 0.0 {
        return Err(validation(path, "border size must be non-negative"));
    }
    let color = string_prop(object, "color", path)?.unwrap_or("000000");
    if color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(validation(path, "border color must be six-digit RGB"));
    }
    let space = object
        .get("space")
        .map(|value| {
            value
                .as_u64()
                .and_then(|value| usize::try_from(value).ok())
                .ok_or_else(|| validation(path, "border space must be a non-negative integer"))
        })
        .transpose()?
        .unwrap_or(0);
    Ok(BorderSpec {
        border_type: style
            .parse()
            .map_err(|_| validation(path, format!("invalid border style `{style}`")))?,
        size: f64_to_usize(size_pt * 8.0, path)?,
        space,
        color: color.to_ascii_uppercase(),
    })
}

fn paragraph_border(position: ParagraphBorderPosition, spec: &BorderSpec) -> ParagraphBorder {
    ParagraphBorder::new(position)
        .val(spec.border_type)
        .size(spec.size)
        .space(spec.space)
        .color(&spec.color)
}

fn paragraph_borders(spec: &BorderSpec) -> ParagraphBorders {
    [
        ParagraphBorderPosition::Top,
        ParagraphBorderPosition::Right,
        ParagraphBorderPosition::Bottom,
        ParagraphBorderPosition::Left,
    ]
    .into_iter()
    .fold(ParagraphBorders::with_empty(), |borders, position| {
        borders.set(paragraph_border(position, spec))
    })
}

fn compile_paragraph_borders(value: &Value, path: &str) -> Result<ParagraphBorders> {
    let object = value
        .as_object()
        .ok_or_else(|| validation(path, "`border` must be an object"))?;
    if object
        .keys()
        .all(|key| ["style", "size", "color", "space"].contains(&key.as_str()))
    {
        return Ok(paragraph_borders(&parse_border_spec(value, path, true)?));
    }
    let mut borders = ParagraphBorders::with_empty();
    if object.get("clearAll").and_then(Value::as_bool) == Some(true) {
        return Ok(borders.clear_all());
    }
    for (key, position) in [
        ("top", ParagraphBorderPosition::Top),
        ("right", ParagraphBorderPosition::Right),
        ("bottom", ParagraphBorderPosition::Bottom),
        ("left", ParagraphBorderPosition::Left),
        ("between", ParagraphBorderPosition::Between),
        ("bar", ParagraphBorderPosition::Bar),
    ] {
        let Some(edge) = object.get(key) else {
            continue;
        };
        if edge == &Value::Bool(false) {
            borders = borders.clear(position);
        } else {
            let spec = parse_border_spec(edge, &format!("{path}/border/{key}"), true)?;
            borders = borders.set(paragraph_border(position, &spec));
        }
    }
    Ok(borders)
}

fn table_borders(spec: &BorderSpec) -> TableBorders {
    [
        TableBorderPosition::Top,
        TableBorderPosition::Left,
        TableBorderPosition::Bottom,
        TableBorderPosition::Right,
        TableBorderPosition::InsideH,
        TableBorderPosition::InsideV,
    ]
    .into_iter()
    .fold(TableBorders::with_empty(), |borders, position| {
        borders.set(
            TableBorder::new(position)
                .border_type(spec.border_type)
                .size(spec.size)
                .color(&spec.color),
        )
    })
}

fn compile_table_borders(value: &Value, path: &str) -> Result<TableBorders> {
    let object = value
        .as_object()
        .ok_or_else(|| validation(path, "`border` must be an object"))?;
    if object
        .keys()
        .all(|key| ["style", "size", "color"].contains(&key.as_str()))
    {
        return Ok(table_borders(&parse_border(value, path)?));
    }
    let mut borders = TableBorders::with_empty();
    if object.get("clearAll").and_then(Value::as_bool) == Some(true) {
        return Ok(borders.clear_all());
    }
    for (key, position) in [
        ("top", TableBorderPosition::Top),
        ("right", TableBorderPosition::Right),
        ("bottom", TableBorderPosition::Bottom),
        ("left", TableBorderPosition::Left),
        ("insideHorizontal", TableBorderPosition::InsideH),
        ("insideVertical", TableBorderPosition::InsideV),
    ] {
        let Some(edge) = object.get(key) else {
            continue;
        };
        if edge == &Value::Bool(false) {
            borders = borders.clear(position);
        } else {
            let spec = parse_border(edge, &format!("{path}/border/{key}"))?;
            borders = borders.set(
                TableBorder::new(position)
                    .border_type(spec.border_type)
                    .size(spec.size)
                    .color(spec.color),
            );
        }
    }
    Ok(borders)
}

fn cell_borders(spec: &BorderSpec) -> TableCellBorders {
    [
        TableCellBorderPosition::Top,
        TableCellBorderPosition::Left,
        TableCellBorderPosition::Bottom,
        TableCellBorderPosition::Right,
    ]
    .into_iter()
    .fold(TableCellBorders::with_empty(), |borders, position| {
        borders.set(
            TableCellBorder::new(position)
                .border_type(spec.border_type)
                .size(spec.size)
                .color(&spec.color),
        )
    })
}

fn compile_cell_borders(value: &Value, path: &str) -> Result<TableCellBorders> {
    let object = value
        .as_object()
        .ok_or_else(|| validation(path, "`border` must be an object"))?;
    if object
        .keys()
        .all(|key| ["style", "size", "color"].contains(&key.as_str()))
    {
        return Ok(cell_borders(&parse_border(value, path)?));
    }
    let mut borders = TableCellBorders::with_empty();
    if object.get("clearAll").and_then(Value::as_bool) == Some(true) {
        return Ok(borders.clear_all());
    }
    for (key, position) in [
        ("top", TableCellBorderPosition::Top),
        ("right", TableCellBorderPosition::Right),
        ("bottom", TableCellBorderPosition::Bottom),
        ("left", TableCellBorderPosition::Left),
        ("insideHorizontal", TableCellBorderPosition::InsideH),
        ("insideVertical", TableCellBorderPosition::InsideV),
        ("topLeftToBottomRight", TableCellBorderPosition::Tl2br),
        ("topRightToBottomLeft", TableCellBorderPosition::Tr2bl),
    ] {
        let Some(edge) = object.get(key) else {
            continue;
        };
        if edge == &Value::Bool(false) {
            borders = borders.clear(position);
        } else {
            let spec = parse_border(edge, &format!("{path}/border/{key}"))?;
            borders = borders.set(
                TableCellBorder::new(position)
                    .border_type(spec.border_type)
                    .size(spec.size)
                    .color(spec.color),
            );
        }
    }
    Ok(borders)
}

fn parse_box_margins(value: &Value, path: &str) -> Result<[usize; 4]> {
    let margins = value
        .as_object()
        .ok_or_else(|| validation(path, "`margins` must be an object"))?;
    for key in margins.keys() {
        if !["top", "right", "bottom", "left"].contains(&key.as_str()) {
            return Err(validation(
                path,
                format!("unknown margins property `{key}`"),
            ));
        }
    }
    let value = |key: &str| {
        required_number(margins, key, path)
            .and_then(|value| to_twips_usize(value, &format!("{path}/margins/{key}")))
    };
    Ok([
        value("top")?,
        value("right")?,
        value("bottom")?,
        value("left")?,
    ])
}

fn parse_table_position(value: &Value, path: &str) -> Result<TablePositionProperty> {
    let position = value
        .as_object()
        .ok_or_else(|| validation(path, "Table `position` must be an object"))?;
    let mut output = TablePositionProperty::new();
    if let Some(value) = number_prop(position, "leftFromText", path)? {
        if value < 0.0 {
            return Err(validation(
                path,
                "position.leftFromText must be non-negative",
            ));
        }
        output = output.left_from_text(to_twips_i32(value, path)?);
    }
    if let Some(value) = number_prop(position, "rightFromText", path)? {
        if value < 0.0 {
            return Err(validation(
                path,
                "position.rightFromText must be non-negative",
            ));
        }
        output = output.right_from_text(to_twips_i32(value, path)?);
    }
    if let Some(value) = string_prop(position, "verticalAnchor", path)? {
        output = output.vertical_anchor(value);
    }
    if let Some(value) = string_prop(position, "horizontalAnchor", path)? {
        output = output.horizontal_anchor(value);
    }
    if let Some(value) = string_prop(position, "xAlign", path)? {
        output = output.position_x_alignment(value);
    }
    if let Some(value) = string_prop(position, "yAlign", path)? {
        output = output.position_y_alignment(value);
    }
    if let Some(value) = number_prop(position, "x", path)? {
        output = output.position_x(to_twips_i32(value, path)?);
    }
    if let Some(value) = number_prop(position, "y", path)? {
        output = output.position_y(to_twips_i32(value, path)?);
    }
    Ok(output)
}

fn compile_row_insert(value: &Value, path: &str) -> Result<Insert> {
    let revision = value
        .as_object()
        .ok_or_else(|| validation(path, "TableRow `inserted` must be an object"))?;
    let mut inserted = Insert::new_with_empty();
    if let Some(author) = string_prop(revision, "author", path)? {
        inserted = inserted.author(author);
    }
    if let Some(date) = string_prop(revision, "date", path)? {
        inserted = inserted.date(date);
    }
    Ok(inserted)
}

fn compile_row_delete(value: &Value, path: &str) -> Result<Delete> {
    let revision = value
        .as_object()
        .ok_or_else(|| validation(path, "TableRow `deleted` must be an object"))?;
    let mut deleted = Delete::new();
    if let Some(author) = string_prop(revision, "author", path)? {
        deleted = deleted.author(author);
    }
    if let Some(date) = string_prop(revision, "date", path)? {
        deleted = deleted.date(date);
    }
    Ok(deleted)
}

fn parse_page_size(value: &Value, path: &str) -> Result<(u32, u32)> {
    if let Some(name) = value.as_str() {
        return match name {
            "A4" => Ok((11_906, 16_838)),
            "Letter" => Ok((12_240, 15_840)),
            _ => Err(validation(
                path,
                "`pageSize` must be A4, Letter, or {width,height}",
            )),
        };
    }
    let object = value
        .as_object()
        .ok_or_else(|| validation(path, "`pageSize` must be A4, Letter, or {width,height}"))?;
    for key in object.keys() {
        if !["width", "height"].contains(&key.as_str()) {
            return Err(validation(
                path,
                format!("unknown pageSize property `{key}`"),
            ));
        }
    }
    Ok((
        to_twips_u32(required_number(object, "width", path)?, path)?,
        to_twips_u32(required_number(object, "height", path)?, path)?,
    ))
}

fn parse_margins(value: &Value, path: &str) -> Result<PageMargin> {
    let object = value
        .as_object()
        .ok_or_else(|| validation(path, "`margins` must be an object"))?;
    for key in object.keys() {
        if ![
            "top", "right", "bottom", "left", "header", "footer", "gutter",
        ]
        .contains(&key.as_str())
        {
            return Err(validation(path, format!("unknown margin property `{key}`")));
        }
    }
    Ok(PageMargin {
        top: margin_value(object, "top", 72.0, path)?,
        right: margin_value(object, "right", 72.0, path)?,
        bottom: margin_value(object, "bottom", 72.0, path)?,
        left: margin_value(object, "left", 72.0, path)?,
        header: margin_value(object, "header", 36.0, path)?,
        footer: margin_value(object, "footer", 36.0, path)?,
        gutter: margin_value(object, "gutter", 0.0, path)?,
    })
}

fn margin_value(props: &Map<String, Value>, key: &str, default: f64, path: &str) -> Result<i32> {
    let value = number_prop(props, key, path)?.unwrap_or(default);
    if value < 0.0 {
        return Err(validation(
            path,
            format!("margin `{key}` must be non-negative"),
        ));
    }
    to_twips_i32(value, path)
}

fn value_to_text(value: &Value, path: &str) -> Result<String> {
    match value {
        Value::String(value) => Ok(value.clone()),
        Value::Number(value) => Ok(value.to_string()),
        _ => Err(validation(path, "Text `value` must be a string or number")),
    }
}

fn required_string<'a>(props: &'a Map<String, Value>, key: &str, path: &str) -> Result<&'a str> {
    let value = string_prop(props, key, path)?
        .ok_or_else(|| validation(path, format!("missing `{key}`")))?;
    if value.is_empty() {
        return Err(validation(path, format!("`{key}` must not be empty")));
    }
    Ok(value)
}

fn compile_run_fonts(props: &Map<String, Value>, path: &str) -> Result<RunFonts> {
    let mut fonts = RunFonts::new();
    if let Some(value) = string_prop(props, "ascii", path)? {
        fonts = fonts.ascii(value);
    }
    if let Some(value) = string_prop(props, "hiAnsi", path)? {
        fonts = fonts.hi_ansi(value);
    }
    if let Some(value) = string_prop(props, "eastAsia", path)? {
        fonts = fonts.east_asia(value);
    }
    if let Some(value) = string_prop(props, "cs", path)? {
        fonts = fonts.cs(value);
    }
    if let Some(value) = string_prop(props, "asciiTheme", path)? {
        fonts = fonts.ascii_theme(value);
    }
    if let Some(value) = string_prop(props, "hiAnsiTheme", path)? {
        fonts = fonts.hi_ansi_theme(value);
    }
    if let Some(value) = string_prop(props, "eastAsiaTheme", path)? {
        fonts = fonts.east_asia_theme(value);
    }
    if let Some(value) = string_prop(props, "csTheme", path)? {
        fonts = fonts.cs_theme(value);
    }
    if let Some(value) = string_prop(props, "hint", path)? {
        fonts = fonts.hint(value);
    }
    Ok(fonts)
}

fn string_prop<'a>(
    props: &'a Map<String, Value>,
    key: &str,
    path: &str,
) -> Result<Option<&'a str>> {
    props
        .get(key)
        .map(|value| {
            value
                .as_str()
                .ok_or_else(|| validation(path, format!("`{key}` must be a string")))
        })
        .transpose()
}

fn object_prop<'a>(
    props: &'a Map<String, Value>,
    key: &str,
    path: &str,
) -> Result<Option<&'a Map<String, Value>>> {
    props
        .get(key)
        .map(|value| {
            value
                .as_object()
                .ok_or_else(|| validation(path, format!("`{key}` must be an object")))
        })
        .transpose()
}

fn required_number(props: &Map<String, Value>, key: &str, path: &str) -> Result<f64> {
    number_prop(props, key, path)?.ok_or_else(|| validation(path, format!("missing `{key}`")))
}

fn required_u64(props: &Map<String, Value>, key: &str, path: &str) -> Result<u64> {
    props
        .get(key)
        .and_then(Value::as_u64)
        .ok_or_else(|| validation(path, format!("`{key}` must be a non-negative integer")))
}

fn number_prop(props: &Map<String, Value>, key: &str, path: &str) -> Result<Option<f64>> {
    props
        .get(key)
        .map(|value| {
            value
                .as_f64()
                .filter(|number| number.is_finite())
                .ok_or_else(|| validation(path, format!("`{key}` must be a finite number")))
        })
        .transpose()
}

fn bool_prop(props: &Map<String, Value>, key: &str, path: &str) -> Result<Option<bool>> {
    props
        .get(key)
        .map(|value| {
            value
                .as_bool()
                .ok_or_else(|| validation(path, format!("`{key}` must be a boolean")))
        })
        .transpose()
}

fn optional_enum<'a>(
    props: &'a Map<String, Value>,
    key: &str,
    allowed: &[&str],
    path: &str,
) -> Result<Option<&'a str>> {
    optional_enum_value(props, key, allowed, path)
}

fn optional_enum_value<'a>(
    props: &'a Map<String, Value>,
    key: &str,
    allowed: &[&str],
    path: &str,
) -> Result<Option<&'a str>> {
    let Some(value) = string_prop(props, key, path)? else {
        return Ok(None);
    };
    if allowed.contains(&value) {
        Ok(Some(value))
    } else {
        Err(validation(path, format!("invalid `{key}` value `{value}`")))
    }
}

fn optional_twips_i32(props: &Map<String, Value>, key: &str, path: &str) -> Result<Option<i32>> {
    number_prop(props, key, path)?
        .map(|value| to_twips_i32(value, &format!("{path}/{key}")))
        .transpose()
}

fn to_half_points(value: f64, path: &str) -> Result<usize> {
    f64_to_usize(value * HALF_POINTS_PER_POINT, path)
}

fn to_twips_u32(value: f64, path: &str) -> Result<u32> {
    f64_to_u32(value * TWIPS_PER_POINT, path)
}

fn to_twips_i32(value: f64, path: &str) -> Result<i32> {
    let rounded = (value * TWIPS_PER_POINT).round();
    if rounded < f64::from(i32::MIN) || rounded > f64::from(i32::MAX) {
        return Err(validation(path, "value is out of range"));
    }
    rounded
        .to_i32()
        .ok_or_else(|| validation(path, "value is out of range"))
}

fn to_twips_usize(value: f64, path: &str) -> Result<usize> {
    f64_to_usize(value * TWIPS_PER_POINT, path)
}

fn to_emu(value: f64, path: &str) -> Result<u32> {
    f64_to_u32(value * EMU_PER_POINT, path)
}

fn to_emu_i32(value: f64, path: &str) -> Result<i32> {
    let rounded = (value * EMU_PER_POINT).round();
    if !rounded.is_finite() || rounded < f64::from(i32::MIN) || rounded > f64::from(i32::MAX) {
        return Err(validation(path, "value is out of range"));
    }
    rounded
        .to_i32()
        .ok_or_else(|| validation(path, "value is out of range"))
}

fn percent_to_fiftieths(value: f64, path: &str) -> Result<usize> {
    f64_to_usize(value * 50.0, path)
}

fn f64_to_u32(value: f64, path: &str) -> Result<u32> {
    let rounded = value.round();
    if !rounded.is_finite() || rounded < 0.0 || rounded > f64::from(u32::MAX) {
        return Err(validation(path, "value is out of range"));
    }
    rounded
        .to_u32()
        .ok_or_else(|| validation(path, "value is out of range"))
}

fn f64_to_usize(value: f64, path: &str) -> Result<usize> {
    let rounded = value.round();
    if !rounded.is_finite() || rounded < 0.0 {
        return Err(validation(path, "value is out of range"));
    }
    rounded
        .to_usize()
        .ok_or_else(|| validation(path, "value is out of range"))
}

fn validation(path: impl Into<String>, message: impl Into<String>) -> Error {
    Error::Validation {
        path: path.into(),
        message: message.into(),
    }
}

#[cfg(test)]
mod tests {
    use std::io::Read;

    use super::*;

    fn minimal_ir() -> IrEnvelope {
        serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"defaultFont":"Arial"},"children":[{"type":"Section","props":{"pageSize":"A4"},"children":[{"type":"Paragraph","props":{"align":"center"},"children":[{"type":"Run","props":{"bold":true,"size":18},"children":["Hello"]}]}]}]}}"#,
        )
        .expect("fixture should parse")
    }

    #[test]
    fn normalize_should_order_table_cell_properties_for_openxml() {
        let mut xml =
            r"<w:tcPr><w:tcW/><w:vAlign/><w:shd/><w:tcMar/><w:tcBorders/><w:vMerge/></w:tcPr>"
                .to_owned();

        normalize_word_xml(&mut xml, "word/document.xml")
            .expect("table cell properties should normalize");

        assert_eq!(
            xml,
            r"<w:tcPr><w:tcW/><w:vMerge/><w:tcBorders/><w:shd/><w:tcMar/><w:vAlign/></w:tcPr>"
        );
    }

    #[test]
    fn normalize_should_remove_unsupported_table_grid_column_type() {
        let mut xml =
            r#"<w:tblGrid><w:gridCol w:w="4400" w:type="dxa" /><w:gridCol w:w="3200" /></w:tblGrid>"#.to_owned();

        normalize_word_xml(&mut xml, "word/document.xml")
            .expect("table grid columns should normalize");

        assert_eq!(
            xml,
            r#"<w:tblGrid><w:gridCol w:w="4400" /><w:gridCol w:w="3200" /></w:tblGrid>"#
        );
    }

    #[test]
    fn compile_should_create_valid_docx_with_text_and_styles() {
        let bytes = compile_document(&minimal_ir(), Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");
        assert!(document.contains("Hello") && document.contains("w:b"));
    }

    #[test]
    fn compile_should_render_distinct_default_font_slots() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"defaultFonts":{"ascii":"Times New Roman","hiAnsi":"Arial","eastAsia":"宋体","cs":"Traditional Arabic","asciiTheme":"majorHAnsi","hiAnsiTheme":"minorHAnsi","eastAsiaTheme":"majorEastAsia","csTheme":"minorBidi","hint":"eastAsia"}},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part should exist")
            .read_to_string(&mut styles)
            .expect("styles XML should be UTF-8");

        for attribute in [
            r#"w:ascii="Times New Roman""#,
            r#"w:hAnsi="Arial""#,
            r#"w:eastAsia="宋体""#,
            r#"w:cs="Traditional Arabic""#,
            r#"w:asciiTheme="majorHAnsi""#,
            r#"w:hAnsiTheme="minorHAnsi""#,
            r#"w:eastAsiaTheme="majorEastAsia""#,
            r#"w:cstheme="minorBidi""#,
            r#"w:hint="eastAsia""#,
        ] {
            assert!(
                styles.contains(attribute),
                "missing {attribute} in {styles}"
            );
        }
    }

    #[test]
    fn point_conversions_should_match_ooxml_units() {
        assert_eq!(to_half_points(12.0, "test").expect("valid"), 24);
        assert_eq!(to_twips_u32(72.0, "test").expect("valid"), 1440);
        assert_eq!(to_emu(1.0, "test").expect("valid"), 12_700);
    }

    #[test]
    fn compile_should_render_two_sections_with_distinct_page_sizes() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"pageSize":"A4","orientation":"portrait"},"children":[{"type":"Paragraph","props":{},"children":["first"]}]},{"type":"Section","props":{"pageSize":"Letter","orientation":"landscape"},"children":[{"type":"Paragraph","props":{},"children":["second"]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");

        let first = document.find(r#"<w:pgSz w:w="11906" w:h="16838""#);
        let second = document.find(r#"<w:pgSz w:w="15840" w:h="12240""#);
        assert!(
            first.is_some() && second.is_some() && first < second,
            "expected distinct A4 then Letter-landscape page sizes in {document}"
        );
        assert!(
            document.matches("<w:sectPr>").count() >= 2
                || document.matches("<w:sectPr ").count() + document.matches("<w:sectPr>").count()
                    >= 2,
            "expected two section properties in {document}"
        );
    }

    #[test]
    fn compile_should_render_section_page_configuration() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"titlePage":true,"textDirection":"tbRl","documentGrid":{"type":"linesAndChars","linePitch":18,"charSpace":-10},"pageNumbering":{"start":3,"chapterStyle":"1","format":"upperRoman","chapterSeparator":"emDash"}},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");

        assert!(document.contains(r"<w:titlePg />"));
        assert!(document.contains(r#"<w:textDirection w:val="tbRl" />"#));
        assert!(document.contains(
            r#"<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="-10" />"#
        ));
        assert!(document.contains(
            r#"<w:pgNumType w:start="3" w:chapStyle="1" w:fmt="upperRoman" w:chapSep="emDash" />"#
        ));
    }

    #[test]
    fn compile_should_support_page_number_format_without_start_or_chapter_style() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"pageNumbering":{"format":"lowerLetter"}},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");

        assert!(
            document.contains(r#"<w:pgNumType w:fmt="lowerLetter" />"#),
            "format-only page numbering must emit a valid page-numbering element: {document}"
        );
    }

    #[test]
    fn compile_should_render_columns_for_each_section() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"columns":2},"children":[{"type":"Paragraph","props":{},"children":["two columns"]}]},{"type":"Section","props":{"columns":3},"children":[{"type":"Paragraph","props":{},"children":["three columns"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");
        assert!(
            document.contains(r#"<w:cols w:space="425" w:num="2" />"#),
            "{document}"
        );
        assert!(
            document.contains(r#"<w:cols w:space="425" w:num="3" />"#),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_custom_section_column_layouts() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"columns":{"count":2,"spacing":18,"separator":true}},"children":[]},{"type":"Section","props":{"columns":{"widths":[{"width":144,"spacing":24},{"width":180,"spacing":18},{"width":216}]}},"children":[]}]}}"#,
        ).expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        assert!(
            document.contains(r#"<w:cols w:space="360" w:num="2" w:sep="1" />"#),
            "{document}"
        );
        assert!(
            document.contains(r#"<w:cols w:num="3" w:equalWidth="0"><w:col w:w="2880" w:space="480" /><w:col w:w="3600" w:space="360" /><w:col w:w="4320" /></w:cols>"#),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_section_type_in_schema_order() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"type":"continuous"},"children":[{"type":"Paragraph","props":{},"children":["body"]}]}]}}"#,
        ).expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        assert!(
            document.contains(r#"<w:sectPr><w:type w:val="continuous" /><w:pgSz"#),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_section_line_numbering() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"lineNumbering":{"countBy":5,"start":3,"distance":18,"restart":"newSection"}},"children":[{"type":"Paragraph","props":{},"children":["body"]}]}]}}"#,
        ).expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        assert!(document.contains(r#"<w:lnNumType w:countBy="5" w:start="3" w:distance="360" w:restart="newSection" />"#), "{document}");
    }

    #[test]
    fn compile_should_render_paragraph_line_number_suppression() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"lineNumbering":{}},"children":[{"type":"Paragraph","props":{"suppressLineNumbers":true},"children":["body"]},{"type":"Paragraph","props":{"suppressLineNumbers":false},"children":["numbered"]}]}]}}"#,
        ).expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        assert!(document.contains("<w:suppressLineNumbers />"), "{document}");
        assert!(
            document.contains("<w:suppressLineNumbers w:val=\"0\" />"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_enable_default_section_line_numbering_for_empty_object() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"lineNumbering":{}},"children":[]}]}}"#,
        ).expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("UTF-8");
        assert!(document.contains("<w:lnNumType />"), "{document}");
    }

    #[test]
    fn compile_should_render_document_settings_and_metadata() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"defaultCharacterSpacing":0.5,"createdAt":"2026-08-14T00:00:00Z","updatedAt":"2026-08-15T00:00:00Z","creator":"Ada & Co","lastModifiedBy":"Lin","description":"Weekly <report>","language":"zh-CN","revision":7,"subject":"Economy","title":"Status","documentProtection":{"edit":"readOnly","formatting":false,"enforcement":true,"cryptProviderType":"rsaAES","cryptAlgorithmClass":"hash","cryptAlgorithmType":"typeAny","cryptAlgorithmSid":14,"cryptSpinCount":100000,"cryptProvider":"RSA AES Provider","algIdExt":"A1B2C3D4","algIdExtSource":"urn:example:alg","cryptProviderTypeExt":"01020304","cryptProviderTypeExtSource":"urn:example:provider","hash":"YWJjZA==","salt":"AQIDBA==","algorithmName":"SHA-512","hashValue":"YWJjZA==","saltValue":"AQIDBA==","spinCount":100000},"customProperties":{"Project":"Apollo"},"documentId":"01234567-89AB-CDEF-0123-456789ABCDEF","defaultTabStop":36,"documentVariables":{"Customer":"Ada"},"evenAndOddHeaders":true,"adjustLineHeightInTable":true,"characterSpacingControl":"compressPunctuation"},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut core = String::new();
        archive
            .by_name("docProps/core.xml")
            .expect("core properties should exist")
            .read_to_string(&mut core)
            .expect("core properties should be UTF-8");
        let mut custom = String::new();
        archive
            .by_name("docProps/custom.xml")
            .expect("custom properties should exist")
            .read_to_string(&mut custom)
            .expect("custom properties should be UTF-8");
        let mut settings = String::new();
        archive
            .by_name("word/settings.xml")
            .expect("settings should exist")
            .read_to_string(&mut settings)
            .expect("settings should be UTF-8");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles should exist")
            .read_to_string(&mut styles)
            .expect("styles should be UTF-8");

        assert!(core.contains("2026-08-14T00:00:00Z") && core.contains("2026-08-15T00:00:00Z"));
        assert!(core.contains("Ada &amp; Co") && core.contains("Weekly &lt;report&gt;"));
        assert!(core.contains("zh-CN") && core.contains("<cp:revision>7</cp:revision>"));
        assert!(core.contains("Economy") && core.contains("Status"));
        assert!(custom.contains(r#"name="Project""#) && custom.contains("Apollo"));
        assert!(settings.contains(r#"w:defaultTabStop w:val="720""#));
        assert!(settings.contains(r#"w15:docId w15:val="{01234567-89AB-CDEF-0123-456789ABCDEF}""#));
        assert!(settings.contains(r#"w:name="Customer" w:val="Ada""#));
        assert!(settings.contains(
            r#"<w:documentProtection w:edit="readOnly" w:formatting="0" w:enforcement="1""#
        ));
        assert!(settings.contains(r#"w:hash="YWJjZA==" w:salt="AQIDBA==""#));
        assert!(settings.contains(r#"w:algorithmName="SHA-512""#));
        assert!(
            settings
                .contains(r#"w:hashValue="YWJjZA==" w:saltValue="AQIDBA==" w:spinCount="100000""#)
        );
        assert!(settings.contains("<w:evenAndOddHeaders />"));
        assert!(settings.contains("<w:adjustLineHeightInTable />"));
        assert!(settings.contains(r#"w:characterSpacingControl w:val="compressPunctuation""#));
        assert!(styles.contains(r#"<w:spacing w:val="10" />"#));
    }

    #[test]
    fn compile_should_render_complete_line_spacing_properties() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"defaultLineSpacing":{"before":2,"after":6,"line":14,"beforeLines":100,"afterLines":200,"lineRule":"atLeast"},"styles":[{"id":"Dense","name":"Dense","type":"paragraph","paragraph":{"spacingBeforeLines":50,"spacingAfterLines":75,"lineRule":"exact"}}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"spacingBefore":3,"spacingAfter":4,"lineSpacing":12,"spacingBeforeLines":125,"spacingAfterLines":250,"lineRule":"auto"},"children":["text"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part")
            .read_to_string(&mut styles)
            .expect("UTF-8 XML");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            styles.contains("<w:spacing w:before=\"40\" w:after=\"120\" w:beforeLines=\"100\" w:afterLines=\"200\" w:line=\"280\" w:lineRule=\"atLeast\" />")
                && styles.contains("<w:spacing w:beforeLines=\"50\" w:afterLines=\"75\" w:lineRule=\"exact\" />")
                && document.contains("<w:spacing w:before=\"60\" w:after=\"80\" w:beforeLines=\"125\" w:afterLines=\"250\" w:line=\"240\" w:lineRule=\"auto\" />"),
            "styles={styles}\ndocument={document}"
        );
    }

    #[test]
    fn compile_should_render_character_unit_indentation_without_point_conversion() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"indentLeft":6,"indentLeftChars":-100,"indentRightChars":200,"firstLineChars":640},"children":["first"]},{"type":"Paragraph","props":{"indentStartChars":120,"indentEndChars":80,"hangingChars":20},"children":["second"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");
        for attr in [
            r#"w:leftChars="-100""#,
            r#"w:left="120""#,
            r#"w:rightChars="200""#,
            r#"w:firstLineChars="640""#,
            r#"w:startChars="120""#,
            r#"w:endChars="80""#,
            r#"w:hangingChars="20""#,
        ] {
            assert!(document.contains(attr), "missing {attr}: {document}");
        }
        assert!(
            !document.contains("F0000000"),
            "temporary marker leaked: {document}"
        );
    }

    #[test]
    fn compile_should_package_web_extensions_and_multiple_custom_xml_items() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"webExtensions":[{"id":"7f33b723-fb58-4524-8733-dbedc4b7c095","referenceId":"office-addin","version":"1.0.0.0","store":"developer","storeType":"Registry","properties":{"mode":"review"}},{"id":"11111111-2222-3333-4444-555555555555","referenceId":"second","version":"2.0","store":"OMEX","storeType":"Marketplace"}],"customXmlItems":[{"id":"06AC5857-5C65-A94A-BCEC-37356A209BC3","xml":"<customer><name>Ada</name></customer>"},{"id":"11111111-AAAA-BBBB-CCCC-222222222222","xml":"<order id=\"42\"/>"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        for name in [
            "word/webextensions/taskpanes.xml",
            "word/webextensions/_rels/taskpanes.xml.rels",
            "word/webextensions/webextension1.xml",
            "word/webextensions/webextension2.xml",
            "customXml/item1.xml",
            "customXml/item2.xml",
            "customXml/itemProps1.xml",
            "customXml/itemProps2.xml",
            "customXml/_rels/item1.xml.rels",
            "customXml/_rels/item2.xml.rels",
        ] {
            assert!(archive.by_name(name).is_ok(), "missing package part {name}");
        }
        let mut read = |name: &str| {
            let mut output = String::new();
            archive
                .by_name(name)
                .unwrap_or_else(|_| panic!("missing {name}"))
                .read_to_string(&mut output)
                .expect("UTF-8 XML");
            output
        };
        let extension = read("word/webextensions/webextension1.xml");
        assert!(extension.contains("office-addin") && extension.contains("name=\"mode\""));
        let taskpanes = read("word/webextensions/taskpanes.xml");
        assert!(taskpanes.contains(r#"r:id="rId1""#) && taskpanes.contains(r#"r:id="rId2""#));
        assert!(read("customXml/item1.xml").contains("<name>Ada</name>"));
        assert!(read("customXml/item2.xml").contains("id=\"42\""));
        let content_types = read("[Content_Types].xml");
        assert!(
            content_types.contains("/customXml/itemProps1.xml")
                && content_types.contains("/customXml/itemProps2.xml"),
            "{content_types}"
        );
    }

    #[test]
    fn compile_should_render_custom_style_definition() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"ReportTitle","name":"Report Title","type":"paragraph","basedOn":"Normal","next":"Normal","quickFormat":false,"uiPriority":5,"semiHidden":true,"unhideWhenUsed":true,"run":{"font":"Noto Sans CJK SC","size":18,"color":"336699","themeColor":"accent1","themeTint":"99","bold":true,"italic":true,"underline":"single","hidden":true,"textBorder":{"style":"double","size":1,"color":"336699","space":2}},"paragraph":{"align":"center","textAlign":"baseline","snapToGrid":false,"spacingAfter":12,"indentLeft":6,"hangingChars":20,"outlineLevel":1,"frame":{"wrap":"around","horizontalAnchor":"margin","verticalAnchor":"text","xAlign":"center","y":12,"horizontalSpace":3,"width":240,"height":48}}},{"id":"ReportFirst","name":"Report First","type":"paragraph","paragraph":{"firstLineChars":640}},{"id":"ReportTable","name":"Report Table","type":"table","table":{"indent":6,"align":"center","margins":{"top":1,"right":2,"bottom":3,"left":4},"border":{"style":"double","size":1,"color":"336699"}},"cell":{"width":72,"colSpan":2,"verticalAlign":"center","verticalMerge":"restart","textDirection":"tbRl","shading":"FFF2CC","margins":{"top":1,"right":2,"bottom":3,"left":4},"border":{"style":"dotted","size":0.5,"color":"993366"}}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles should exist")
            .read_to_string(&mut styles)
            .expect("styles should be UTF-8");

        assert!(styles.contains(r#"w:type="paragraph" w:styleId="ReportTitle""#));
        assert!(styles.contains(r#"<w:name w:val="Report Title" />"#));
        assert!(styles.contains(r#"w:ascii="Noto Sans CJK SC""#));
        assert!(styles.contains(r#"w:val="336699" w:themeColor="accent1" w:themeTint="99""#));
        assert!(styles.contains("<w:b />") && styles.contains("<w:i />"));
        assert!(
            styles.contains(r#"<w:bdr w:val="double" w:sz="8" w:space="2" w:color="336699" />"#)
        );
        assert!(styles.contains(r#"<w:snapToGrid w:val="false" />"#));
        assert!(styles.contains(r#"w:after="240""#));
        assert!(styles.contains(r#"w:left="120" w:right="0""#));
        assert!(styles.contains(r#"w:hangingChars="20""#));
        let report_first_style = styles
            .split(r#"w:styleId="ReportFirst""#)
            .nth(1)
            .and_then(|value| value.split("</w:style>").next())
            .expect("ReportFirst style body should exist");
        assert!(report_first_style.contains(r#"w:firstLineChars="640""#));
        assert!(styles.contains(r#"w:wrap="around""#) && styles.contains(r#"w:y="240""#));
        assert!(styles.contains(r#"<w:uiPriority w:val="5" />"#));
        assert!(styles.contains("<w:semiHidden />") && styles.contains("<w:unhideWhenUsed />"));
        let report_style = styles
            .split(r#"w:styleId="ReportTitle""#)
            .nth(1)
            .and_then(|value| value.split("</w:style>").next())
            .expect("ReportTitle style body should exist");
        assert!(!report_style.contains("<w:qFormat />"));
        assert!(styles.contains(r#"w:type="table" w:styleId="ReportTable""#));
        assert!(styles.contains(r#"<w:jc w:val="center" />"#));
        assert!(styles.contains(r#"<w:gridSpan w:val="2" />"#));
        assert!(styles.contains(r#"<w:textDirection w:val="tbRl" />"#));
        assert!(styles.contains(r#"<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC" />"#));
    }

    #[test]
    fn compile_should_render_extended_paragraph_style_properties() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","paragraph":{"pageBreakBefore":true,"widowControl":false,"bidi":true,"shading":"DDEEFF","adjustRightIndent":1}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part should exist")
            .read_to_string(&mut styles)
            .expect("styles XML should be UTF-8");

        let body = styles
            .split(r#"w:styleId="Body""#)
            .nth(1)
            .and_then(|value| value.split("</w:style>").next())
            .expect("Body style should exist");
        assert!(body.contains("<w:pageBreakBefore />"), "{body}");
        assert!(body.contains(r#"<w:widowControl w:val="0" />"#), "{body}");
        assert!(body.contains("<w:bidi />"), "{body}");
        assert!(body.contains(r#"<w:adjustRightInd w:val="1" />"#), "{body}");
        assert!(body.contains(r#"w:fill="DDEEFF""#), "{body}");
    }

    #[test]
    fn compile_should_render_extended_style_run_properties() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Emphasis","name":"Emphasis","type":"character","run":{"bold":false,"italic":false,"strike":true,"doubleStrike":false,"caps":true,"verticalAlign":"superscript","characterSpacing":-1.5}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part should exist")
            .read_to_string(&mut styles)
            .expect("styles XML should be UTF-8");

        let emphasis = styles
            .split(r#"w:styleId="Emphasis""#)
            .nth(1)
            .and_then(|value| value.split("</w:style>").next())
            .expect("Emphasis style should exist");
        for expected in [
            r#"<w:b w:val="false" />"#,
            r#"<w:i w:val="false" />"#,
            "<w:strike />",
            r#"<w:dstrike w:val="false" />"#,
            r#"<w:caps w:val="true" />"#,
            r#"<w:vertAlign w:val="superscript" />"#,
            r#"<w:spacing w:val="-30" />"#,
        ] {
            assert!(
                emphasis.contains(expected),
                "missing {expected} in {emphasis}"
            );
        }
    }

    #[test]
    fn compile_should_render_style_line_number_suppression() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"NoLineNumbers","name":"No Line Numbers","type":"paragraph","paragraph":{"suppressLineNumbers":true}},{"id":"KeepLineNumbers","name":"Keep Line Numbers","type":"paragraph","paragraph":{"suppressLineNumbers":false}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        ).expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles")
            .read_to_string(&mut styles)
            .expect("UTF-8");
        assert!(styles.contains("<w:suppressLineNumbers />"), "{styles}");
        assert!(
            styles.contains("<w:suppressLineNumbers w:val=\"0\" />"),
            "{styles}"
        );
    }

    #[test]
    fn compile_should_render_all_style_character_indent_sides() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"PhysicalIndent","name":"Physical Indent","type":"paragraph","paragraph":{"indentLeftChars":-100,"indentRightChars":200}},{"id":"LogicalIndent","name":"Logical Indent","type":"paragraph","paragraph":{"indentStartChars":120,"indentEndChars":80}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles should exist")
            .read_to_string(&mut styles)
            .expect("styles XML should be UTF-8");
        for attr in [
            r#"w:leftChars="-100""#,
            r#"w:rightChars="200""#,
            r#"w:startChars="120""#,
            r#"w:endChars="80""#,
        ] {
            assert!(styles.contains(attr), "missing {attr}: {styles}");
        }
    }

    #[test]
    fn compile_should_emit_style_keep_flags_and_a_single_normal() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Normal","name":"Normal","type":"paragraph","quickFormat":true},{"id":"Kept","name":"Kept","type":"paragraph","paragraph":{"keepNext":true,"keepLines":true,"outlineLevel":1}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles should exist")
            .read_to_string(&mut styles)
            .expect("styles should be UTF-8");
        assert_eq!(
            styles.matches(r#"w:styleId="Normal""#).count(),
            1,
            "{styles}"
        );
        assert!(
            styles.contains("<w:keepNext")
                && styles.contains("<w:keepLines")
                && styles.contains(r#"<w:outlineLvl w:val="1""#),
            "{styles}"
        );
    }

    #[test]
    fn compile_should_render_list_level_writer_properties() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{"type":"ordered","multiLevelType":"hybridMultilevel","levels":[{"format":"decimal","text":"%1.","suffix":"space","legal":true,"restart":0,"indentLeft":36,"hanging":18,"bold":true,"size":12},{"format":"lowerLetter","text":"%2)","align":"right","paragraphStyle":"ListParagraph"}]},"children":[{"type":"ListItem","props":{"style":"ListParagraph"},"children":["one"]},{"type":"ListItem","props":{"level":1},"children":["two"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut numbering = String::new();
        archive
            .by_name("word/numbering.xml")
            .expect("numbering part")
            .read_to_string(&mut numbering)
            .expect("numbering UTF-8");
        assert!(
            numbering.contains(r#"w:val="decimal""#)
                && numbering.contains(r#"w:val="lowerLetter""#)
                && numbering.contains(r#"<w:suff w:val="space" />"#)
                && numbering.contains("<w:isLgl")
                && numbering.contains(r#"<w:lvlRestart w:val="0""#)
                && numbering.contains(r#"w:val="ListParagraph""#)
                && numbering.contains("<w:b")
                && numbering.contains(r#"w:val="right""#)
                && numbering.contains(r#"<w:multiLevelType w:val="hybridMultilevel" />"#),
            "{numbering}"
        );
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document")
            .read_to_string(&mut document)
            .expect("document UTF-8");
        assert!(
            document.contains("<w:numPr>") && document.contains(r#"w:val="ListParagraph""#),
            "{document}"
        );
    }

    #[test]
    fn compile_should_preserve_numbering_style_links_without_default_levels() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{"styleLink":"ParagraphList","numStyleLink":"ListNumber","multiLevelType":"multilevel"},"children":[{"type":"ListItem","props":{},"children":["item"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut numbering = String::new();
        archive
            .by_name("word/numbering.xml")
            .expect("numbering part")
            .read_to_string(&mut numbering)
            .expect("numbering UTF-8");
        let linked = numbering
            .split(r#"<w:abstractNum w:abstractNumId="2">"#)
            .nth(1)
            .expect("linked definition should be present");
        assert!(
            linked.contains(r#"<w:numStyleLink w:val="ListNumber" />"#),
            "{linked}"
        );
        assert!(
            linked.contains(r#"<w:styleLink w:val="ParagraphList" /><w:numStyleLink"#),
            "link children must be in schema order: {linked}"
        );
        assert!(
            !linked.contains("<w:lvl "),
            "linked definition must not synthesize levels: {linked}"
        );
    }

    #[test]
    fn compile_should_render_list_instance_level_override_in_schema_order() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{"type":"ordered","start":3,"overrides":[{"level":0,"start":4,"definition":{"format":"lowerLetter","text":"%1)"}}]},"children":[{"type":"ListItem","props":{},"children":["item"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX ZIP");
        let mut numbering = String::new();
        archive
            .by_name("word/numbering.xml")
            .expect("numbering part")
            .read_to_string(&mut numbering)
            .expect("numbering UTF-8");
        assert!(
            numbering.contains(
                r#"<w:lvlOverride w:ilvl="0"><w:startOverride w:val="4" /><w:lvl w:ilvl="0">"#
            ) && numbering.contains(r#"w:val="lowerLetter""#),
            "{numbering}"
        );
        assert!(
            numbering.contains(r#"<w:lvl w:ilvl="0"><w:start w:val="3" />"#),
            "replacement level should inherit List.start when `start` is omitted: {numbering}"
        );
    }

    #[test]
    fn compile_should_allow_a_replacement_to_add_a_list_level() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{"levels":[{"format":"decimal","text":"%1."}],"overrides":[{"level":8,"definition":{"format":"upperRoman","text":"%9."}}]},"children":[{"type":"ListItem","props":{"level":8},"children":["item"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document UTF-8");
        assert!(document.contains(r#"<w:ilvl w:val="8" />"#), "{document}");
    }

    #[test]
    fn compile_should_omit_unused_comments_footnotes_and_numbering_parts() {
        let ir = minimal_ir();
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let names: Vec<String> = archive.file_names().map(ToOwned::to_owned).collect();
        for part in [
            "word/comments.xml",
            "word/commentsExtended.xml",
            "word/footnotes.xml",
            "word/numbering.xml",
        ] {
            assert!(
                !names.iter().any(|name| name == part),
                "unused part {part} must be omitted, got {names:?}"
            );
        }
        let mut rels = String::new();
        let mut archive = archive;
        archive
            .by_name("word/_rels/document.xml.rels")
            .expect("document rels")
            .read_to_string(&mut rels)
            .expect("rels UTF-8");
        assert!(
            !rels.contains("comments.xml")
                && !rels.contains("commentsExtended.xml")
                && !rels.contains("footnotes.xml")
                && !rels.contains("numbering.xml"),
            "{rels}"
        );
        let mut types = String::new();
        archive
            .by_name("[Content_Types].xml")
            .expect("content types")
            .read_to_string(&mut types)
            .expect("types UTF-8");
        assert!(
            !types.contains("/word/comments.xml")
                && !types.contains("/word/commentsExtended.xml")
                && !types.contains("/word/footnotes.xml")
                && !types.contains("/word/numbering.xml"),
            "{types}"
        );
    }

    #[test]
    fn compile_should_keep_comments_footnotes_and_numbering_when_used() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Comment","props":{"text":"review","author":"Ada"},"children":["marked"]},{"type":"Footnote","props":{},"children":["note"]}]},{"type":"List","props":{"type":"ordered"},"children":[{"type":"ListItem","props":{},"children":["one"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        for part in [
            "word/comments.xml",
            "word/footnotes.xml",
            "word/numbering.xml",
        ] {
            archive
                .by_name(part)
                .unwrap_or_else(|_| panic!("{part} should exist"));
        }
        let mut comments = String::new();
        archive
            .by_name("word/comments.xml")
            .expect("comments")
            .read_to_string(&mut comments)
            .expect("comments UTF-8");
        assert!(comments.contains("review"), "{comments}");
        let mut footnotes = String::new();
        archive
            .by_name("word/footnotes.xml")
            .expect("footnotes")
            .read_to_string(&mut footnotes)
            .expect("footnotes UTF-8");
        assert!(footnotes.contains("note"), "{footnotes}");
        let mut numbering = String::new();
        archive
            .by_name("word/numbering.xml")
            .expect("numbering")
            .read_to_string(&mut numbering)
            .expect("numbering UTF-8");
        assert!(numbering.contains("w:abstractNum"), "{numbering}");
    }

    #[test]
    fn compile_should_render_style_inheritance_links_and_typed_references() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"BodyBase","name":"Body Base","type":"paragraph"},{"id":"Body","name":"Body","type":"paragraph","basedOn":"BodyBase","next":"Body","link":"BodyChar"},{"id":"BodyChar","name":"Body Char","type":"character","link":"Body"},{"id":"TableBase","name":"Table Base","type":"table"},{"id":"ReportTable","name":"Report Table","type":"table","basedOn":"TableBase"}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"style":"Body"},"children":[{"type":"Run","props":{"style":"BodyChar"},"children":["styled"]}]},{"type":"Table","props":{"style":"ReportTable"},"children":[{"type":"TableRow","props":{},"children":[{"type":"TableCell","props":{},"children":[{"type":"Paragraph","props":{},"children":["cell"]}]}]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles should exist")
            .read_to_string(&mut styles)
            .expect("styles should be UTF-8");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document should exist")
            .read_to_string(&mut document)
            .expect("document should be UTF-8");

        for expected in [
            r#"<w:basedOn w:val="BodyBase" />"#,
            r#"<w:next w:val="Body" />"#,
            r#"<w:link w:val="BodyChar" />"#,
            r#"<w:link w:val="Body" />"#,
            r#"<w:basedOn w:val="TableBase" />"#,
        ] {
            assert!(styles.contains(expected), "missing {expected} in {styles}");
        }
        for expected in [
            r#"<w:pStyle w:val="Body" />"#,
            r#"<w:rStyle w:val="BodyChar" />"#,
            r#"<w:tblStyle w:val="ReportTable" />"#,
        ] {
            assert!(
                document.contains(expected),
                "missing {expected} in {document}"
            );
        }
    }

    #[test]
    fn compile_should_render_paragraph_run_defaults() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"snapToGrid":false,"widowControl":true,"font":"Noto Sans CJK SC","size":12,"bold":true,"italic":true,"color":"1a2B3c","characterSpacing":0.5},"children":["body"]},{"type":"Heading","props":{"level":2,"font":"Noto Sans CJK SC","size":16},"children":["title"]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");

        assert!(document.contains(r#"<w:snapToGrid w:val="false" />"#));
        assert!(document.contains(r#"<w:widowControl w:val="1" />"#));
        assert!(document.contains(r#"<w:rFonts w:ascii="Noto Sans CJK SC""#));
        assert!(document.contains(r#"<w:sz w:val="24" />"#));
        assert!(document.contains(r#"<w:color w:val="1A2B3C" />"#));
        assert!(document.contains(r#"<w:spacing w:val="10" />"#));
        assert!(document.contains("<w:b />") && document.contains("<w:i />"));
    }

    #[test]
    fn compile_should_render_font_slots_on_runs_paragraphs_and_styles() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","run":{"fonts":{"ascii":"Style Latin","eastAsia":"样式中文"}}}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"fonts":{"ascii":"Paragraph Latin","eastAsia":"段落中文"}},"children":[{"type":"Run","props":{"fonts":{"ascii":"Run Latin","hiAnsi":"Run ANSI","eastAsia":"运行中文","cs":"Run CS","hint":"eastAsia"}},"children":["mixed"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part should exist")
            .read_to_string(&mut styles)
            .expect("styles XML should be UTF-8");

        assert!(document.contains(r#"w:ascii="Paragraph Latin""#));
        assert!(document.contains(r#"w:eastAsia="段落中文""#));
        assert!(document.contains(r#"w:ascii="Run Latin""#));
        assert!(document.contains(r#"w:hAnsi="Run ANSI""#));
        assert!(document.contains(r#"w:eastAsia="运行中文""#));
        assert!(document.contains(r#"w:cs="Run CS""#));
        assert!(document.contains(r#"w:hint="eastAsia""#));
        assert!(styles.contains(r#"w:ascii="Style Latin""#));
        assert!(styles.contains(r#"w:eastAsia="样式中文""#));
    }

    #[test]
    fn compile_should_render_hyperlink_composite_children() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"href":"https://example.com"},"children":[{"type":"ContentControl","props":{"alias":"LinkData"},"children":["bound"]},{"type":"Inserted","props":{"author":"Ada"},"children":["new"]},{"type":"Deleted","props":{"author":"Lin"},"children":["old"]},{"type":"InlineBookmark","props":{"name":"insideLink"},"children":["marked"]},{"type":"Comment","props":{"text":"review","author":"Grace"},"children":["commented"]}]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");
        let hyperlink = document
            .split_once("<w:hyperlink")
            .and_then(|(_, rest)| rest.split_once("</w:hyperlink>"))
            .map(|(content, _)| content)
            .expect("hyperlink XML should exist");

        for marker in [
            "<w:sdt>",
            r#"<w:alias w:val="LinkData""#,
            "<w:ins ",
            "<w:del ",
            r#"w:name="insideLink""#,
            "<w:bookmarkEnd ",
            "<w:commentRangeStart ",
            "<w:commentRangeEnd ",
        ] {
            assert!(
                hyperlink.contains(marker),
                "missing {marker} in {hyperlink}"
            );
        }
        assert!(
            hyperlink.find("<w:bookmarkStart").expect("start")
                < hyperlink.find("insideLink").expect("name")
                && hyperlink.find("marked").expect("marked")
                    < hyperlink.find("<w:bookmarkEnd").expect("end"),
            "bookmark markers should wrap marked text: {hyperlink}"
        );
        let mut comments = String::new();
        archive
            .by_name("word/comments.xml")
            .expect("comments part should exist")
            .read_to_string(&mut comments)
            .expect("comments XML should be UTF-8");
        assert!(
            comments.contains("review") && comments.contains("Grace"),
            "{comments}"
        );
    }

    #[test]
    fn compile_should_render_document_level_content_control() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["before"]},{"type":"ContentControl","props":{"alias":"BlockData"},"children":[{"type":"Paragraph","props":{},"children":["inside"]}]},{"type":"Paragraph","props":{},"children":["after"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");
        let sdt = document
            .split_once("<w:sdt>")
            .and_then(|(_, rest)| rest.split_once("</w:sdt>"))
            .map(|(content, _)| content)
            .expect("document-level sdt should exist");
        assert!(
            sdt.contains(r#"<w:alias w:val="BlockData""#) && sdt.contains("inside"),
            "missing block control markup in {sdt}"
        );
        assert!(
            document.find("before").expect("before") < document.find("<w:sdt>").expect("sdt")
                && document.find("</w:sdt>").expect("sdt end")
                    < document.find("after").expect("after"),
            "block sdt should sit between surrounding paragraphs: {document}"
        );
    }

    #[test]
    fn compile_should_render_comment_range_inside_deleted_text() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Deleted","props":{"author":"Ada"},"children":[{"type":"Comment","props":{"text":"Typo","author":"Grace"},"children":["removed"]}]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");
        let deleted = document
            .split_once("<w:del ")
            .and_then(|(_, rest)| rest.split_once("</w:del>"))
            .map(|(content, _)| content)
            .expect("deleted range should exist");
        assert!(
            deleted.contains("<w:commentRangeStart w:id=\"1\" />")
                && deleted.contains("<w:delText")
                && deleted.contains("removed")
                && deleted.contains("<w:commentRangeEnd w:id=\"1\" />")
                && deleted.contains("<w:commentReference w:id=\"1\" />"),
            "comment range should enclose deleted text: {deleted}"
        );
    }

    #[test]
    fn compile_should_render_deleted_field_instruction() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [{
                        "type": "Deleted", "props": {"author": "Editor"}, "children": [{
                            "type": "Run", "props": {}, "children": [
                            {"type": "DeletedInstruction", "props": {
                                "instruction": "HYPERLINK \"https://example.com\""
                            }, "children": []}]
                        }]
                    }]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("<w:delInstrText>HYPERLINK \"https://example.com\"</w:delInstrText>"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_header_and_footer_block_content_controls() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Header","props":{},"children":[{"type":"ContentControl","props":{"alias":"HeaderData","xpath":"/root/header","storeItemId":"{ABC}"},"children":[{"type":"Paragraph","props":{},"children":["header"]}]}]},{"type":"Footer","props":{},"children":[{"type":"ContentControl","props":{"alias":"FooterData","xpath":"/root/footer","storeItemId":"{ABC}"},"children":[{"type":"Paragraph","props":{},"children":["footer"]}]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        for (part, alias, text) in [
            ("word/header1.xml", "HeaderData", "header"),
            ("word/footer1.xml", "FooterData", "footer"),
        ] {
            let mut xml = String::new();
            archive
                .by_name(part)
                .expect("header/footer part should exist")
                .read_to_string(&mut xml)
                .expect("XML should be UTF-8");
            assert!(
                xml.contains(&format!(r#"<w:alias w:val="{alias}""#))
                    && xml.contains(&format!(r#"w:xpath="/root/{text}""#))
                    && xml.contains(text),
                "block content control missing from {part}: {xml}"
            );
        }
    }

    #[test]
    fn compile_should_render_advanced_paragraph_properties_and_revisions() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"bidi":true,"textAlign":"baseline","adjustRightIndent":1,"shading":"DDEEFF","outlineLevel":3,"frame":{"wrap":"around","horizontalAnchor":"page","x":12,"yAlign":"top","width":144,"height":72},"inserted":{"author":"Ada","date":"2026-08-14T00:00:00Z"},"propertyChange":{"author":"Lin","date":"2026-08-13T00:00:00Z","previous":{"align":"right","spacingAfter":6,"bidi":false}}},"children":["body"]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("<w:bidi />")
                && document.contains("<w:textAlignment w:val=\"baseline\" />")
                && document.contains("<w:adjustRightInd w:val=\"1\" />")
                && document
                    .contains("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"DDEEFF\" />")
                && document.contains("<w:outlineLvl w:val=\"3\" />")
                && document.contains("<w:framePr w:wrap=\"around\"")
                && document.contains("w:hAnchor=\"page\"")
                && document.contains("w:x=\"240\"")
                && document.contains("w:yAlign=\"top\"")
                && document.contains("w:w=\"2880\"")
                && document.contains("w:h=\"1440\"")
                && document.contains("<w:ins w:id=")
                && document.contains("w:author=\"Ada\"")
                && document.contains("<w:pPrChange w:id=")
                && document.contains("w:author=\"Lin\"")
                && document.contains("<w:jc w:val=\"right\" />")
                && document.contains("<w:spacing w:after=\"120\" />"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_paragraph_id_and_borders() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"paragraphId":"a1b2c3d4","border":{"top":{"style":"double","size":1,"color":"336699","space":2},"between":false,"bar":{"style":"babyRattle"}}},"children":["one"]},{"type":"Paragraph","props":{"border":{"style":"dashed","size":0.5,"color":"993366","space":1}},"children":["two"]},{"type":"Paragraph","props":{"border":{"clearAll":true}},"children":[]}]}]}}"#,
        )
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("w14:paraId=\"A1B2C3D4\"")
                && document.contains(
                    "<w:top w:val=\"double\" w:space=\"2\" w:sz=\"8\" w:color=\"336699\" />"
                )
                && document.contains("<w:between w:val=\"nil\"")
                && document.contains("<w:bar w:val=\"babyRattle\"")
                && document.contains(
                    "<w:left w:val=\"dashed\" w:space=\"1\" w:sz=\"4\" w:color=\"993366\" />"
                )
                && document.matches("w:val=\"nil\"").count() >= 7,
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_run_theme_and_explicit_formatting_flags() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Run","props":{"color":"2e74b5","themeColor":"accent1","themeShade":"bf","themeTint":"99","bold":false,"italic":false,"strike":false,"doubleStrike":true},"children":["themed"]}]}]}]}}"#,
        )
        .expect("fixture should parse");
        ir.validate().expect("fixture should validate");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");

        assert!(document.contains(
            r#"<w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF" w:themeTint="99" />"#
        ));
        assert!(document.contains(r#"<w:b w:val="false" />"#));
        assert!(document.contains(r#"<w:i w:val="false" />"#));
        assert!(document.contains(r#"<w:strike w:val="false" />"#));
        assert!(document.contains("<w:dstrike />"));
    }

    #[test]
    fn compile_should_render_table_properties() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {
                "type": "Document", "props": {}, "children": [{
                    "type": "Section", "props": {}, "children": [{
                        "type": "Table",
                        "props": {
                            "widthPercent": 100,
                            "style": "GridTable4",
                            "indent": 12,
                            "margins": {"top": 2, "right": 3, "bottom": 4, "left": 5},
                            "position": {"leftFromText": 7.1, "rightFromText": 7.1, "verticalAnchor": "text", "horizontalAnchor": "margin", "xAlign": "right", "y": 25.5},
                            "layout": "fixed",
                            "columnWidths": [100, 100],
                            "border": {"top": {"style": "double", "size": 1, "color": "112233"}, "insideHorizontal": false}
                        },
                        "children": [{
                            "type": "TableRow", "props": {"cantSplit": true, "gridBefore": 1, "gridAfter": 2, "widthBefore": 60, "widthAfter": 75, "inserted": {"author": "Ada", "date": "2026-08-14T00:00:00Z"}}, "children": [{
                                "type": "TableCell",
                                "props": {"width": 120, "colSpan": 2, "verticalAlign": "center", "verticalMerge": "restart", "textDirection": "tbRl", "margins": {"top": 1, "right": 2, "bottom": 3, "left": 4}, "shading": "EEEEEE", "border": {"left": false, "topLeftToBottomRight": {"style": "dotted", "size": 0.5, "color": "993366"}}},
                                "children": [
                                    {"type": "Paragraph", "props": {}, "children": ["cell"]},
                                    {"type": "TableOfContents", "props": {}, "children": []},
                                    {"type": "ContentControl", "props": {"alias": "Cell"}, "children": ["value"]}
                                ]
                            }]
                        }, {
                            "type": "TableRow", "props": {"deleted": {"author": "Linus"}}, "children": [{
                                "type": "TableCell", "props": {"widthPercent": 50}, "children": []
                            }]
                        }]
                    }]
                }]
            }
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part should exist")
            .read_to_string(&mut document)
            .expect("document XML should be UTF-8");
        assert!(
            document.contains("<w:tbl>")
                && document.contains("w:gridSpan")
                && document.contains("w:fill=\"EEEEEE\"")
                && document.contains(r#"<w:tblStyle w:val="GridTable4" />"#)
                && document.contains(r#"<w:gridBefore w:val="1" />"#)
                && document.contains(r#"<w:gridAfter w:val="2" />"#)
                && document.contains(r#"<w:wBefore w:w="1200" w:type="dxa" />"#)
                && document.contains(r#"<w:wAfter w:w="1500" w:type="dxa" />"#)
                && document.contains(r#"<w:tblInd w:w="240" w:type="dxa" />"#)
                && document.contains("<w:tblCellMar>")
                && document.contains("<w:tcMar>")
                && document.contains(r#"<w:tcW w:w="2400" w:type="dxa" />"#)
                && document.contains(r#"<w:tcW w:w="2500" w:type="pct" />"#)
                && document.contains(r#"<w:vMerge w:val="restart" />"#)
                && document.contains(r#"<w:textDirection w:val="tbRl" />"#)
                && document
                    .contains(r#"<w:top w:val="double" w:sz="8" w:space="0" w:color="112233" />"#)
                && document.contains(r#"<w:insideH w:val="nil""#)
                && document.contains(r#"<w:left w:val="nil""#)
                && document.contains(
                    r#"<w:tl2br w:val="dotted" w:sz="4" w:space="0" w:color="993366" />"#
                )
                && document.contains("<w:sdt>")
                && document.contains("TOC")
                && document.contains(r#"w:leftFromText="142""#)
                && document.contains(r#"w:rightFromText="142""#)
                && document.contains(r#"w:vertAnchor="text""#)
                && document.contains(r#"w:horzAnchor="margin""#)
                && document.contains(r#"w:tblpXSpec="right""#)
                && document.contains(r#"w:tblpY="510""#)
                && document.contains(r"<w:ins w:id=")
                && document.contains(r#"w:author="Ada""#)
                && document.contains(r"<w:del w:id=")
                && document.contains(r#"w:author="Linus""#)
        );
    }

    #[test]
    fn compile_should_embed_image_media() {
        let directory = tempfile::tempdir().expect("tempdir should work");
        let image_path = directory.path().join("pixel.png");
        image::DynamicImage::new_rgba8(1, 1)
            .save(&image_path)
            .expect("image should write");
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {
                "type": "Document", "props": {}, "children": [{
                    "type": "Section", "props": {}, "children": [{
                        "type": "Paragraph", "props": {}, "children": [{
                            "type": "Image", "props": {
                                "src": "pixel.png", "width": 12, "height": 12,
                                "relationshipId": "rIdHero", "rotate": 45,
                                "floating": true, "allowOverlap": true,
                                "simplePosition": {"x": 25_400, "y": -12_700},
                                "layoutInCell": true,
                                "positionH": "right", "positionV": 18.5,
                                "relativeFromH": "page", "relativeFromV": "paragraph",
                                "distanceTop": 2, "distanceBottom": 3,
                                "distanceLeft": 4, "distanceRight": 5,
                                "relativeHeight": 251_658_240
                            }, "children": []
                        }]
                    }]
                }]
            }
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, directory.path()).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let has_media = archive
            .file_names()
            .any(|name| name.starts_with("word/media/"));
        assert!(has_media);
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("document XML");
        assert!(
            document.contains(r#"<wp:anchor distT="25400" distB="38100" distL="50800" distR="63500" simplePos="1" allowOverlap="1""#)
                && document.contains(r#"<wp:simplePos x="25400" y="-12700" />"#)
                && document.contains(r#"relativeHeight="251658240""#)
                && document.contains(r#"layoutInCell="1""#)
                && document.contains(r#"<wp:positionH relativeFrom="page"><wp:align>right</wp:align>"#)
                && document.contains(r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>234950</wp:posOffset>"#)
                && document.contains(r#"<a:xfrm rot="2700000">"#),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_bound_content_control() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [{
                        "type": "ContentControl",
                        "props": {
                            "alias": "Customer",
                            "xpath": "/root/customer",
                            "prefixMappings": "xmlns:x='urn:test'",
                            "storeItemId": "{ABC}"
                        },
                        "children": [{"type": "Run", "props": {"bold": true}, "children": ["Ada"]}]
                    }]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("<w:sdt>")
                && document.contains("<w:alias w:val=\"Customer\"")
                && document.contains("w:xpath=\"/root/customer\"")
                && document.contains("w:prefixMappings=\"xmlns:x='urn:test'\"")
                && document.contains("w:storeItemID=\"{ABC}\"")
                && document.contains("<w:sdtContent><w:r>"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_generic_complex_field() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [{
                        "type": "Field",
                        "props": {"instruction": " DATE \\@ \"yyyy-MM-dd\" ", "dirty": false},
                        "children": [{"type": "Run", "props": {"bold": true}, "children": ["2026-08-14"]}]
                    }]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("w:fldCharType=\"begin\"")
                && document.contains("w:dirty=\"false\"")
                && document.contains(" DATE \\@ \"yyyy-MM-dd\" ")
                && document.contains("w:fldCharType=\"separate\"")
                && document.contains("2026-08-14")
                && document.contains("w:fldCharType=\"end\""),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_typed_fields() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [
                        {"type": "DateField", "props": {"format": "yyyy-MM-dd"}, "children": ["date"]},
                        {"type": "TimeField", "props": {"format": "HH:mm", "dirty": false}, "children": []},
                        {"type": "FileNameField", "props": {"fullPath": true}, "children": ["file"]},
                        {"type": "AuthorField", "props": {}, "children": []},
                        {"type": "TitleField", "props": {}, "children": []},
                        {"type": "SubjectField", "props": {}, "children": []},
                        {"type": "MergeField", "props": {"name": "CustomerName", "preserveFormatting": true, "placeholder": "Ada"}, "children": []},
                        {"type": "DocumentPropertyField", "props": {"name": "Project Name"}, "children": ["Apollo"]},
                        {"type": "FormulaField", "props": {"expression": "SUM(ABOVE)", "numberFormat": "#,##0.00"}, "children": ["42.00"]}
                    ]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains(" DATE \\@ \"yyyy-MM-dd\" ")
                && document.contains(" TIME \\@ \"HH:mm\" ")
                && document.contains(" FILENAME \\p ")
                && document.contains(" AUTHOR ")
                && document.contains(" TITLE ")
                && document.contains(" SUBJECT ")
                && document.contains(" MERGEFIELD CustomerName \\* MERGEFORMAT ")
                && document.contains(" DOCPROPERTY \"Project Name\" ")
                && document.contains(" = SUM(ABOVE) \\# \"#,##0.00\" "),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_sequence_and_reference_fields() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [
                    {"type": "Bookmark", "props": {"name": "target"}, "children": []},
                    {"type": "Paragraph", "props": {}, "children": [
                        {"type": "SequenceField", "props": {"identifier": "Figure", "format": "Roman", "restart": 3, "placeholder": "III"}, "children": []},
                        {"type": "ReferenceField", "props": {"bookmark": "target", "hyperlink": true, "relativePosition": true, "placeholder": "target text"}, "children": []}
                    ]}
                ]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains(" SEQ Figure \\* ROMAN \\r 3 ")
                && document.contains(" REF target \\h \\p ")
                && document.contains("III")
                && document.contains("target text"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_caption_as_styled_sequence_paragraph() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Caption",
                    "props": {
                        "label": "Figure",
                        "identifier": "Diagram",
                        "format": "Roman",
                        "restart": 3,
                        "placeholder": "III",
                        "dirty": false,
                        "style": "FigureCaption",
                        "numberSeparator": " ",
                        "textSeparator": " — "
                    },
                    "children": [{"type": "Run", "props": {"bold": true}, "children": ["Architecture"]}]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("w:val=\"FigureCaption\"")
                && document.contains("Figure ")
                && document.contains(" SEQ Diagram \\* ROMAN \\r 3 ")
                && document.contains("w:dirty=\"false\"")
                && document.contains("III")
                && document.contains(" — ")
                && document.contains("Architecture"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_native_index_fields() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [
                    {"type": "Paragraph", "props": {}, "children": [{
                        "type": "IndexEntry",
                        "props": {"text": "Rust", "subentry": "Ownership", "identifier": "topics", "boldPageNumber": true},
                        "children": []
                    }]},
                    {"type": "Index", "props": {"identifier": "topics", "columns": 2, "runIn": true, "placeholder": "Update index", "dirty": false, "style": "IndexBody"}, "children": []}
                ]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains(" XE \"Rust:Ownership\" \\f \"topics\" \\b ")
                && document.contains(" INDEX \\f \"topics\" \\c \"2\" \\r ")
                && document.contains("w:val=\"IndexBody\"")
                && document.contains("w:dirty=\"false\"")
                && document.contains("Update index"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_carriage_return_and_custom_tab_stop() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [
                        {"type": "TabStop", "props": {"position": 72, "align": "right", "leader": "dot"}, "children": []},
                        "Label",
                        {"type": "Tab", "props": {}, "children": []},
                        "Value",
                        {"type": "CarriageReturn", "props": {}, "children": []}
                    ]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains(
                "<w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"1440\" /></w:tabs>"
            ) && document.contains("<w:tab />")
                && document.contains("<w:cr />"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_non_breaking_and_soft_characters() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [
                        "left", {"type": "NonBreakingSpace", "props": {}, "children": []},
                        "soft", {"type": "SoftHyphen", "props": {}, "children": []},
                        "break", {"type": "NonBreakingHyphen", "props": {}, "children": []}, "right"
                    ]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains('\u{00a0}')
                && document.contains('\u{00ad}')
                && document.contains('\u{2011}'),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_toc_entry_and_move_revisions() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [
                        {"type": "TocEntry", "props": {"text": "Appendix", "level": 2, "omitPageNumber": true, "identifier": "figures"}, "children": []},
                        {"type": "MovedFrom", "props": {"author": "Ada", "date": "2026-08-14T00:00:00Z"}, "children": ["old"]},
                        {"type": "MovedTo", "props": {"author": "Ada", "date": "2026-08-14T00:00:00Z"}, "children": [
                            {"type": "Run", "props": {"bold": true}, "children": ["new"]}
                        ]}
                    ]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("<w:instrText>TC \"Appendix\" \\f figures \\l 2 \\n</w:instrText>")
                && document.contains("<w:moveFrom ")
                && document.contains("w:author=\"Ada\"")
                && document.contains(">old</w:t>")
                && document.contains("<w:moveTo ")
                && document.contains("<w:b />")
                && document.contains(">new</w:t>"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_figure_and_custom_entry_indexes() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [
                    {"type": "TableOfContents", "props": {
                        "startLevel": 1, "endLevel": 3,
                        "tcFieldLevelRange": {"start": 2, "end": 4},
                        "omitPageNumbersLevelRange": {"start": 5, "end": 7},
                        "entryBookmarkName": "contents",
                        "entryAndPageNumberSeparator": " — ",
                        "stylesWithLevels": [{"style": "Report Heading", "level": 2}],
                        "preserveTab": true, "preserveNewLine": true,
                        "useAppliedParagraphLineLevel": true,
                        "hideTabAndPageNumbersInWebview": true
                    }, "children": []},
                    {"type": "TableOfFigures", "props": {
                        "label": "Figure", "includeLabelAndNumber": true,
                        "separator": " — ", "hyperlinks": true, "dirty": false,
                        "alias": "Figures"
                    }, "children": []},
                    {"type": "TableOfEntries", "props": {
                        "identifier": "legal", "hyperlinks": true, "dirty": true,
                        "alias": "Legal entries"
                    }, "children": []},
                    {"type": "Paragraph", "props": {}, "children": ["body"]}
                ]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("TOC \\c &quot;Figure&quot; \\d &quot; — &quot; \\h")
                && document.contains("TOC \\b &quot;contents&quot; \\l &quot;2-4&quot; \\n &quot;5-7&quot; \\o &quot;1-3&quot; \\p &quot; — &quot; \\t &quot;Report Heading,2&quot; \\h \\u \\w \\x \\z")
                && document.contains("<w:alias w:val=\"Figures\"")
                && document.contains("TOC \\f &quot;legal&quot; \\h")
                && document.contains("<w:alias w:val=\"Legal entries\"")
                && document.contains("<w:fldChar w:fldCharType=\"begin\" w:dirty=\"false\" />")
                && document.matches("w:dirty=\"true\"").count() >= 2,
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_static_toc_item_with_page_reference() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "TableOfContents", "props": {"alias": "Contents"},
                    "children": [{"type": "TocItem", "props": {
                        "text": "Overview", "bookmark": "overview", "level": 2,
                        "pageRef": "7"
                    }, "children": []}]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(document.contains("w:anchor=\"overview\""), "{document}");
        assert!(document.contains("PAGEREF overview"), "{document}");
        assert!(document.contains(">7</w:t>"), "{document}");
        assert!(document.contains("w:pos=\"8400\""), "{document}");
        assert!(!document.contains("w:pos=\"80000\""), "{document}");
    }

    #[test]
    fn compile_should_place_toc_before_and_after_paragraphs_around_field() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "TableOfContents", "props": {}, "children": [
                        {"type": "TocBefore", "props": {}, "children": [
                            {"type": "Paragraph", "props": {}, "children": ["before index"]}
                        ]},
                        {"type": "TocAfter", "props": {}, "children": [
                            {"type": "Paragraph", "props": {}, "children": ["after index"]}
                        ]}
                    ]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        let before = document.find("before index").expect("before paragraph");
        let field = document.find("TOC ").expect("TOC field");
        let after = document.find("after index").expect("after paragraph");
        assert!(before < field && field < after, "{document}");
    }

    #[test]
    fn compile_should_render_semantic_text_wrappers_as_independent_runs() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [
                        "H",
                        {"type": "Subscript", "props": {}, "children": [2]},
                        "O x",
                        {"type": "Superscript", "props": {}, "children": [2]},
                        {"type": "AllCaps", "props": {}, "children": ["draft"]},
                        {"type": "HiddenText", "props": {}, "children": ["internal"]}
                    ]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("<w:vertAlign w:val=\"subscript\" />")
                && document.contains("<w:vertAlign w:val=\"superscript\" />")
                && document.contains("<w:caps w:val=\"true\" />")
                && document.contains("<w:vanish />")
                && document.contains("</w:r><w:r><w:rPr><w:vertAlign"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_basic_formatting_wrappers() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [
                        {"type": "Bold", "props": {}, "children": ["bold"]},
                        {"type": "Italic", "props": {}, "children": ["italic"]},
                        {"type": "Underline", "props": {"type": "wave"}, "children": ["underlined"]},
                        {"type": "StrikeThrough", "props": {}, "children": ["removed"]}
                    ]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("<w:b />")
                && document.contains("<w:i />")
                && document.contains("<w:u w:val=\"wave\"")
                && document.contains("<w:strike />"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_advanced_text_effects() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {}, "children": [{
                "type": "Section", "props": {}, "children": [{
                    "type": "Paragraph", "props": {}, "children": [
                        {"type": "DoubleStrike", "props": {}, "children": ["obsolete"]},
                        {"type": "SpacedText", "props": {"amount": 1.5}, "children": ["wide"]},
                        {"type": "SpacedText", "props": {"amount": -0.5}, "children": ["tight"]},
                        {"type": "ScaledText", "props": {"percent": 125}, "children": ["scaled"]},
                        {"type": "FitText", "props": {"width": 42, "id": 7}, "children": ["fitted"]},
                        {"type": "BorderedText", "props": {"style": "double", "size": 1, "color": "336699", "space": 2}, "children": ["bordered"]},
                        {"type": "ShadedText", "props": {"fill": "FFF2CC", "color": "336699", "pattern": "pct20"}, "children": ["shaded"]},
                        {"type": "SpecialHiddenText", "props": {}, "children": ["metadata"]}
                    ]
                }]
            }]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut document = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document part")
            .read_to_string(&mut document)
            .expect("UTF-8 XML");
        assert!(
            document.contains("<w:dstrike />")
                && document.contains("<w:spacing w:val=\"30\" />")
                && document.contains("<w:spacing w:val=\"-10\" />")
                && document.contains("<w:w w:val=\"125\" />")
                && document.contains("<w:fitText w:val=\"840\" w:id=\"7\" />")
                && document.contains(
                    "<w:bdr w:val=\"double\" w:sz=\"8\" w:space=\"2\" w:color=\"336699\" />"
                )
                && document
                    .contains("<w:shd w:val=\"pct20\" w:color=\"336699\" w:fill=\"FFF2CC\" />")
                && document.contains("<w:specVanish />"),
            "{document}"
        );
    }

    #[test]
    fn compile_should_render_extended_style_run_effects() {
        let ir: IrEnvelope = serde_json::from_value(serde_json::json!({
            "version": 1,
            "document": {"type": "Document", "props": {"styles": [{
                "id": "Effects", "name": "Effects", "type": "character",
                "run": {"fitText": {"width": 42, "id": 9}, "stretch": 125,
                    "specVanish": true,
                    "shading": {"fill": "FFF2CC", "color": "336699", "pattern": "pct20"}}
            }]}, "children": [{"type": "Section", "props": {}, "children": []}]}
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, Path::new(".")).expect("compile should work");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles part")
            .read_to_string(&mut styles)
            .expect("UTF-8 XML");
        assert!(
            styles.contains("<w:fitText w:val=\"840\" w:id=\"9\" />")
                && styles.contains("<w:w w:val=\"125\" />")
                && styles.contains("<w:specVanish />")
                && styles
                    .contains("<w:shd w:val=\"pct20\" w:color=\"336699\" w:fill=\"FFF2CC\" />"),
            "{styles}"
        );
    }
}
