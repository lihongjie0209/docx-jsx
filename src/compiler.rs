use std::fmt::Write as _;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use docx_rs::{
    AbstractNumbering, AlignmentType, BorderType, BreakType, CellMargins, CharacterSpacingValues,
    Comment, DataBinding, Delete, DocGrid, DocGridType, Docx, FieldCharType, Footer, Footnote,
    Header, HeightRule, Hyperlink, HyperlinkType, IndentLevel, Insert, InstrPAGEREF, InstrTC,
    InstrText, InstrToC, Level, LevelJc, LevelText, LineSpacing, MoveFrom, MoveTo, NumPages,
    NumberFormat, Numbering, NumberingId, PageMargin, PageNum, PageNumType, PageOrientationType,
    PageSize, Paragraph, ParagraphBorder, ParagraphBorderPosition, ParagraphBorders,
    ParagraphPropertyChange, Pic, PositionalTab, PositionalTabAlignmentType,
    PositionalTabRelativeTo, Run, RunFonts, Section, Settings, Shading, ShdType, SpecialIndentType,
    Start, StructuredDataTag, Style, StyleType, Sym, Tab as DocxTab, TabLeaderType, TabValueType,
    Table, TableAlignmentType, TableBorder, TableBorderPosition, TableBorders, TableCell,
    TableCellBorder, TableCellBorderPosition, TableCellBorders, TableCellMargins,
    TableCellProperty, TableLayoutType, TableOfContents, TablePositionProperty, TableRow,
    TextAlignmentType, TextBorder, TextDirectionType, ThemeColor, VAlignType, VMergeType,
    VertAlignType, WidthType,
};
use image::ImageFormat;
use num_traits::ToPrimitive;
use serde_json::{Map, Value};

use crate::error::{Error, Result};
use crate::ir::{Child, IrEnvelope, Node, NodeKind};

const TWIPS_PER_POINT: f64 = 20.0;
const HALF_POINTS_PER_POINT: f64 = 2.0;
const EMU_PER_POINT: f64 = 12_700.0;

#[derive(Default)]
struct CompileContext {
    next_numbering_id: usize,
    next_bookmark_id: usize,
    next_comment_id: usize,
    numberings: Vec<(AbstractNumbering, Numbering)>,
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
    if let Some(size) = number_prop(&ir.document.props, "defaultSize", "Document")? {
        docx = docx.default_size(to_half_points(size, "Document/defaultSize")?);
    }
    if let Some(spacing) = number_prop(&ir.document.props, "defaultCharacterSpacing", "Document")? {
        docx = docx.default_spacing(to_twips_i32(spacing, "Document/defaultCharacterSpacing")?);
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
    let mut context = CompileContext::default();
    for (index, child) in ir.document.children.iter().enumerate() {
        let Child::Node(section_node) = child else {
            return Err(validation("Document", "expected Section"));
        };
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
                    NodeKind::TableOfFigures => compile_table_of_figures(child, &child_path)?,
                    NodeKind::TableOfEntries => compile_table_of_entries(child, &child_path)?,
                    _ => compile_table_of_contents(child, &child_path)?,
                };
                docx = docx.add_table_of_contents(compiled_index);
            }
        }
        docx = docx.add_section(compile_section(
            section_node,
            entry_dir,
            &format!("Document/Section[{index}]"),
            &mut context,
        )?);
        for (child_index, child) in section_node.children.iter().enumerate() {
            let Child::Node(child) = child else { continue };
            let child_path = format!(
                "Document/Section[{index}]/{}[{child_index}]",
                child.kind.name()
            );
            docx = match child.kind {
                NodeKind::Header => {
                    attach_header(docx, child, entry_dir, &child_path, &mut context)?
                }
                NodeKind::Footer => {
                    attach_footer(docx, child, entry_dir, &child_path, &mut context)?
                }
                _ => docx,
            };
        }
    }
    for (abstract_numbering, numbering) in context.numberings {
        docx = docx
            .add_abstract_numbering(abstract_numbering)
            .add_numbering(numbering);
    }
    let mut cursor = Cursor::new(Vec::new());
    docx.pack(&mut cursor)
        .map_err(|error| Error::Compile(error.to_string()))?;
    let bytes = patch_external_hyperlink_relationships(cursor.into_inner(), &ir.document)?;
    embed_ir_manifest(bytes, ir)
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
    if bool_prop(run, "bold", path)? == Some(true) {
        style = style.bold();
    }
    if bool_prop(run, "italic", path)? == Some(true) {
        style = style.italic();
    }
    if let Some(value) = string_prop(run, "underline", path)? {
        style = style.underline(value);
    }
    if bool_prop(run, "hidden", path)? == Some(true) {
        style = style.vanish();
    }
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
    if present {
        style = style.line_spacing(spacing);
    }
    Ok(style)
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
    if left.is_some() || right.is_some() || special.is_some() {
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
    let mut comments_xml = String::new();
    source
        .by_name("word/comments.xml")
        .map_err(|error| Error::Compile(format!("missing comments.xml: {error}")))?
        .read_to_string(&mut comments_xml)
        .map_err(|error| Error::Compile(format!("cannot read comments.xml: {error}")))?;
    let end = comments_xml
        .rfind("/>")
        .ok_or_else(|| Error::Compile("invalid empty comments.xml".to_owned()))?;
    comments_xml.replace_range(end.., ">\n");
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
    let relationship = r#"<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml" />"#;
    relationships = relationships.replacen(
        "</Relationships>",
        &format!("{relationship}</Relationships>"),
        1,
    );
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

fn compile_section(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Section> {
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
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(child) = child else {
            return Err(validation(path, "Section only accepts structural children"));
        };
        let child_path = format!("{path}/{}[{index}]", child.kind.name());
        section = match child.kind {
            NodeKind::Paragraph => {
                section.add_paragraph(compile_paragraph(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Heading => {
                section.add_paragraph(compile_heading(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Caption => {
                section.add_paragraph(compile_caption(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Index => section.add_paragraph(compile_index(child, &child_path)?),
            NodeKind::Table => {
                section.add_table(compile_table(child, entry_dir, &child_path, context)?)
            }
            NodeKind::List => add_list_to_section(section, child, entry_dir, &child_path, context)?,
            NodeKind::Bookmark => {
                add_bookmark_to_section(section, child, entry_dir, &child_path, context)?
            }
            NodeKind::TableOfContents
            | NodeKind::TableOfFigures
            | NodeKind::TableOfEntries
            | NodeKind::Header
            | NodeKind::Footer => section,
            _ => return Err(validation(&child_path, "unsupported Section child")),
        };
    }
    Ok(section)
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

fn compile_paragraph(
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Paragraph> {
    let mut paragraph = Paragraph::new();
    if let Some(style) = string_prop(&node.props, "style", path)? {
        paragraph = paragraph.style(style);
    }
    if let Some(value) = optional_enum(
        &node.props,
        "align",
        &["left", "center", "right", "both"],
        path,
    )? {
        paragraph = paragraph.align(match value {
            "center" => AlignmentType::Center,
            "right" => AlignmentType::Right,
            "both" => AlignmentType::Both,
            _ => AlignmentType::Left,
        });
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
    if has_spacing {
        paragraph = paragraph.line_spacing(spacing);
    }
    let left = optional_twips_i32(&node.props, "indentLeft", path)?;
    let right = optional_twips_i32(&node.props, "indentRight", path)?;
    let special = if let Some(value) = number_prop(&node.props, "firstLine", path)? {
        Some(SpecialIndentType::FirstLine(to_twips_i32(value, path)?))
    } else if let Some(value) = number_prop(&node.props, "hanging", path)? {
        Some(SpecialIndentType::Hanging(to_twips_i32(value, path)?))
    } else {
        None
    };
    if left.is_some() || right.is_some() || special.is_some() {
        paragraph = paragraph.indent(left, special, right, None);
    }
    if let Some(value) = bool_prop(&node.props, "snapToGrid", path)? {
        paragraph = paragraph.snap_to_grid(value);
    }
    if let Some(value) = bool_prop(&node.props, "widowControl", path)? {
        paragraph = paragraph.widow_control(value);
    }
    if let Some(font) = string_prop(&node.props, "font", path)? {
        paragraph = paragraph.fonts(
            RunFonts::new()
                .ascii(font)
                .hi_ansi(font)
                .east_asia(font)
                .cs(font),
        );
    }
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
                paragraph =
                    paragraph.add_hyperlink(compile_hyperlink(link, entry_dir, &child_path)?);
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
                paragraph = paragraph.add_delete(compile_deleted(deleted, &child_path)?);
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
                paragraph =
                    paragraph.add_hyperlink(compile_hyperlink(link, entry_dir, &child_path)?);
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

fn compile_deleted(node: &Node, path: &str) -> Result<Delete> {
    let mut deleted = Delete::new();
    if let Some(author) = string_prop(&node.props, "author", path)? {
        deleted = deleted.author(author);
    }
    if let Some(date) = string_prop(&node.props, "date", path)? {
        deleted = deleted.date(date);
    }
    for (index, child) in node.children.iter().enumerate() {
        let child_path = format!("{path}/child[{index}]");
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
        Child::Node(_) => Err(validation(path, "Deleted only accepts text content")),
    }
}

fn compile_hyperlink(node: &Node, entry_dir: &Path, path: &str) -> Result<Hyperlink> {
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
        let run = match child {
            Child::Node(run) if run.kind == NodeKind::Run => {
                compile_run(run, entry_dir, &child_path)?
            }
            child => compile_run_child(Run::new(), child, entry_dir, &child_path)?,
        };
        hyperlink = hyperlink.add_run(run);
    }
    Ok(hyperlink)
}

fn compile_content_control(node: &Node, entry_dir: &Path, path: &str) -> Result<StructuredDataTag> {
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
    let bytes = std::fs::read(&source_path).map_err(|source| Error::Resource {
        path: source_path.clone(),
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
    let width = required_number(&node.props, "width", path)?;
    let height = required_number(&node.props, "height", path)?;
    Ok(Pic::new_with_dimensions(png.into_inner(), 1, 1).size(
        to_emu(width, &format!("{path}/width"))?,
        to_emu(height, &format!("{path}/height"))?,
    ))
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
            NodeKind::TableOfContents => {
                cell.add_table_of_contents(compile_table_of_contents(child, &child_path)?)
            }
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
            _ => return Err(validation(&child_path, "unsupported Footer child")),
        }
    }
    Ok(footer)
}

fn add_list_to_section(
    mut section: Section,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Section> {
    for paragraph in compile_list(node, entry_dir, path, context)? {
        section = section.add_paragraph(paragraph);
    }
    Ok(section)
}

fn add_bookmark_to_section(
    mut section: Section,
    node: &Node,
    entry_dir: &Path,
    path: &str,
    context: &mut CompileContext,
) -> Result<Section> {
    context.next_bookmark_id += 1;
    let id = context.next_bookmark_id;
    section = section.add_bookmark_start(id, required_string(&node.props, "name", path)?);
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(child) = child else {
            return Err(validation(
                path,
                "Bookmark only accepts structural children",
            ));
        };
        let child_path = format!("{path}/{}[{index}]", child.kind.name());
        section = match child.kind {
            NodeKind::Paragraph => {
                section.add_paragraph(compile_paragraph(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Heading => {
                section.add_paragraph(compile_heading(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Caption => {
                section.add_paragraph(compile_caption(child, entry_dir, &child_path, context)?)
            }
            NodeKind::Table => {
                section.add_table(compile_table(child, entry_dir, &child_path, context)?)
            }
            NodeKind::List => add_list_to_section(section, child, entry_dir, &child_path, context)?,
            _ => return Err(validation(&child_path, "unsupported Bookmark child")),
        };
    }
    Ok(section.add_bookmark_end(id))
}

fn compile_table_of_contents(node: &Node, path: &str) -> Result<TableOfContents> {
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
    let mut toc = TableOfContents::new()
        .heading_styles_range(start, end)
        .auto();
    if bool_prop(&node.props, "hyperlinks", path)?.unwrap_or(true) {
        toc = toc.hyperlink();
    }
    if bool_prop(&node.props, "dirty", path)?.unwrap_or(true) {
        toc = toc.dirty();
    }
    if let Some(alias) = string_prop(&node.props, "alias", path)? {
        toc = toc.alias(alias);
    }
    Ok(toc)
}

fn compile_table_of_figures(node: &Node, path: &str) -> Result<TableOfContents> {
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
    configure_index(toc, node, path)
}

fn compile_table_of_entries(node: &Node, path: &str) -> Result<TableOfContents> {
    let identifier = required_string(&node.props, "identifier", path)?;
    let mut toc = TableOfContents::new();
    toc.instr = InstrToC::new().tc_field_identifier(Some(identifier.to_owned()));
    configure_index(toc, node, path)
}

fn configure_index(mut toc: TableOfContents, node: &Node, path: &str) -> Result<TableOfContents> {
    toc = toc.auto();
    if bool_prop(&node.props, "hyperlinks", path)?.unwrap_or(true) {
        toc = toc.hyperlink();
    }
    if bool_prop(&node.props, "dirty", path)?.unwrap_or(true) {
        toc = toc.dirty();
    }
    if let Some(alias) = string_prop(&node.props, "alias", path)? {
        toc = toc.alias(alias);
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
    for level in 0..=8 {
        let text = if list_type == "ordered" {
            format!("%{}.", level + 1)
        } else {
            "•".to_owned()
        };
        abstract_numbering = abstract_numbering.add_level(
            Level::new(
                level,
                Start::new(start),
                NumberFormat::new(if list_type == "ordered" {
                    "decimal"
                } else {
                    "bullet"
                }),
                LevelText::new(text),
                LevelJc::new("left"),
            )
            .indent(
                Some(
                    i32::try_from((level + 1) * 720)
                        .map_err(|_| validation(path, "list indent is out of range"))?,
                ),
                Some(SpecialIndentType::Hanging(360)),
                None,
                None,
            ),
        );
    }
    context
        .numberings
        .push((abstract_numbering, Numbering::new(id, id)));
    let mut paragraphs = Vec::with_capacity(node.children.len());
    for (index, child) in node.children.iter().enumerate() {
        let Child::Node(item) = child else {
            return Err(validation(path, "List only accepts ListItem"));
        };
        let item_path = format!("{path}/ListItem[{index}]");
        let level = usize::try_from(item.props.get("level").and_then(Value::as_u64).unwrap_or(0))
            .map_err(|_| validation(&item_path, "`level` is out of range"))?;
        let mut paragraph =
            Paragraph::new().numbering(NumberingId::new(id), IndentLevel::new(level));
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
    fn point_conversions_should_match_ooxml_units() {
        assert_eq!(to_half_points(12.0, "test").expect("valid"), 24);
        assert_eq!(to_twips_u32(72.0, "test").expect("valid"), 1440);
        assert_eq!(to_emu(1.0, "test").expect("valid"), 12_700);
    }

    #[test]
    fn compile_should_render_section_page_configuration() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"titlePage":true,"textDirection":"tbRl","documentGrid":{"type":"linesAndChars","linePitch":18,"charSpace":-10},"pageNumbering":{"start":3,"chapterStyle":"1"}},"children":[]}]}}"#,
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
        assert!(document.contains(r#"<w:pgNumType w:start="3" w:chapStyle="1" />"#));
    }

    #[test]
    fn compile_should_render_document_settings_and_metadata() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"defaultCharacterSpacing":0.5,"createdAt":"2026-08-14T00:00:00Z","updatedAt":"2026-08-15T00:00:00Z","customProperties":{"Project":"Apollo"},"documentId":"01234567-89AB-CDEF-0123-456789ABCDEF","defaultTabStop":36,"documentVariables":{"Customer":"Ada"},"evenAndOddHeaders":true,"adjustLineHeightInTable":true,"characterSpacingControl":"compressPunctuation"},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
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
        assert!(custom.contains(r#"name="Project""#) && custom.contains("Apollo"));
        assert!(settings.contains(r#"w:defaultTabStop w:val="720""#));
        assert!(settings.contains(r#"w15:docId w15:val="{01234567-89AB-CDEF-0123-456789ABCDEF}""#));
        assert!(settings.contains(r#"w:name="Customer" w:val="Ada""#));
        assert!(settings.contains("<w:evenAndOddHeaders />"));
        assert!(settings.contains("<w:adjustLineHeightInTable />"));
        assert!(settings.contains(r#"w:characterSpacingControl w:val="compressPunctuation""#));
        assert!(styles.contains(r#"<w:spacing w:val="10" />"#));
    }

    #[test]
    fn compile_should_render_custom_style_definition() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"ReportTitle","name":"Report Title","type":"paragraph","basedOn":"Normal","next":"Normal","quickFormat":false,"uiPriority":5,"semiHidden":true,"unhideWhenUsed":true,"run":{"font":"Noto Sans CJK SC","size":18,"color":"336699","themeColor":"accent1","themeTint":"99","bold":true,"italic":true,"underline":"single","hidden":true,"textBorder":{"style":"double","size":1,"color":"336699","space":2}},"paragraph":{"align":"center","textAlign":"baseline","snapToGrid":false,"spacingAfter":12,"indentLeft":6,"firstLine":2,"hangingChars":20,"outlineLevel":1,"frame":{"wrap":"around","horizontalAnchor":"margin","verticalAnchor":"text","xAlign":"center","y":12,"horizontalSpace":3,"width":240,"height":48}}},{"id":"ReportTable","name":"Report Table","type":"table","table":{"style":"BaseTable","indent":6,"widthPercent":80,"align":"center","layout":"fixed","margins":{"top":1,"right":2,"bottom":3,"left":4},"border":{"style":"double","size":1,"color":"336699"}},"cell":{"width":72,"colSpan":2,"verticalAlign":"center","verticalMerge":"restart","textDirection":"tbRl","shading":"FFF2CC","margins":{"top":1,"right":2,"bottom":3,"left":4},"border":{"style":"dotted","size":0.5,"color":"993366"}}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
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
        assert!(styles.contains(r#"w:left="120" w:right="0" w:firstLine="40""#));
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
        assert!(styles.contains(r#"<w:tblStyle w:val="BaseTable" />"#));
        assert!(styles.contains(r#"<w:tblW w:w="4000" w:type="pct" />"#));
        assert!(styles.contains(r#"<w:tblLayout w:type="fixed" />"#));
        assert!(styles.contains(r#"<w:gridSpan w:val="2" />"#));
        assert!(styles.contains(r#"<w:textDirection w:val="tbRl" />"#));
        assert!(styles.contains(r#"<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC" />"#));
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
    fn compile_should_render_advanced_paragraph_properties_and_revisions() {
        let ir: IrEnvelope = serde_json::from_str(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"bidi":true,"textAlign":"baseline","adjustRightIndent":-2,"shading":"DDEEFF","outlineLevel":3,"frame":{"wrap":"around","horizontalAnchor":"page","x":12,"yAlign":"top","width":144,"height":72},"inserted":{"author":"Ada","date":"2026-08-14T00:00:00Z"},"propertyChange":{"author":"Lin","date":"2026-08-13T00:00:00Z","previous":{"align":"right","spacingAfter":6,"bidi":false}}},"children":["body"]}]}]}}"#,
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
                && document.contains("<w:adjustRightInd w:val=\"-2\" />")
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
                            "type": "TableRow", "props": {"cantSplit": true, "inserted": {"author": "Ada", "date": "2026-08-14T00:00:00Z"}}, "children": [{
                                "type": "TableCell",
                                "props": {"colSpan": 2, "verticalAlign": "center", "verticalMerge": "restart", "textDirection": "tbRl", "margins": {"top": 1, "right": 2, "bottom": 3, "left": 4}, "shading": "EEEEEE", "border": {"left": false, "topLeftToBottomRight": {"style": "dotted", "size": 0.5, "color": "993366"}}},
                                "children": [
                                    {"type": "Paragraph", "props": {}, "children": ["cell"]},
                                    {"type": "TableOfContents", "props": {}, "children": []},
                                    {"type": "ContentControl", "props": {"alias": "Cell"}, "children": ["value"]}
                                ]
                            }]
                        }, {
                            "type": "TableRow", "props": {"deleted": {"author": "Linus"}}, "children": [{
                                "type": "TableCell", "props": {}, "children": []
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
                && document.contains(r#"<w:tblInd w:w="240" w:type="dxa" />"#)
                && document.contains("<w:tblCellMar>")
                && document.contains("<w:tcMar>")
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
                            "type": "Image", "props": {"src": "pixel.png", "width": 12, "height": 12}, "children": []
                        }]
                    }]
                }]
            }
        }))
        .expect("fixture should parse");
        let bytes = compile_document(&ir, directory.path()).expect("compile should work");
        let archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("DOCX should be ZIP");
        let has_media = archive
            .file_names()
            .any(|name| name.starts_with("word/media/"));
        assert!(has_media);
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
                    {"type": "TableOfFigures", "props": {
                        "label": "Figure", "includeLabelAndNumber": true,
                        "separator": " — ", "hyperlinks": true, "dirty": true,
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
                && document.contains("<w:alias w:val=\"Figures\"")
                && document.contains("TOC \\f &quot;legal&quot; \\h")
                && document.contains("<w:alias w:val=\"Legal entries\"")
                && document.matches("w:dirty=\"true\"").count() >= 2,
            "{document}"
        );
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
}
