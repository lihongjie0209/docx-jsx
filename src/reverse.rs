//! DOCX to deterministic, recompilable JSX conversion.

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::fmt::Write as _;
use std::io::{Cursor, Read};
use std::path::Path;

use docx_rs::{
    Comment, CommentChild, Delete, DeleteChild, DocumentChild, DrawingData, DrawingPosition,
    DrawingPositionType, FooterChild, HeaderChild, Hyperlink, HyperlinkData, Insert, InsertChild,
    MoveFrom, MoveFromChild, MoveTo, MoveToChild, ParagraphChild, Pic, PicAlign, RelativeFromHType,
    RelativeFromVType, RunChild, StructuredDataTag, StructuredDataTagChild, TableCellContent,
    TableChild, TableRowChild, TextBoxContentChild,
};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use serde::Serialize;
use serde_json::{Map, Number, Value};

use crate::ir::{Child, IrEnvelope, Node};
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
    let docx = docx_rs::read_docx(bytes).map_err(|error| Error::Reverse(error.to_string()))?;
    let style_metadata = read_style_metadata(bytes)?;
    let theme_colors = parse_theme_colors(bytes)?;
    let mut images = collect_package_images(bytes)?;
    for (id, asset) in collect_images(&docx) {
        images.entry(id).or_insert(asset);
    }
    let mut writer = Writer {
        images,
        style_theme_colors: theme_colors.styles,
        run_theme_colors: theme_colors.runs,
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
    let document_attrs = reverse_document_attributes(&docx, &style_metadata, &writer)?;
    writer.line(0, &format!("<Document{document_attrs}>"));
    writer.emit_sections(&docx)?;
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
    hyperlink_targets: HashMap<String, String>,
    sdt_plan: SdtPlan,
    paragraph_index: usize,
    hyperlink_index: usize,
    body_sdt_index: usize,
    cell_sdt_index: usize,
    apply_sdt: bool,
    images: HashMap<String, ImageAsset>,
    assets: BTreeMap<String, Vec<u8>>,
    style_theme_colors: HashMap<String, ThemeColorInfo>,
    run_theme_colors: Vec<ThemeColorInfo>,
    run_theme_index: usize,
}

#[derive(Clone, Debug)]
struct ImageAsset {
    src: String,
    bytes: Vec<u8>,
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
        for (index, (children, break_property)) in groups.into_iter().enumerate() {
            let property = break_property.unwrap_or(&docx.document.section_property);
            self.line(1, &format!("<Section{}>", section_jsx_attrs(property)?));
            if index == 0 {
                self.apply_sdt = false;
                self.headers_and_footers(property, 2)?;
                self.apply_sdt = true;
            }
            for child in children {
                if section_break_of(child).is_some() && is_section_break_marker(child) {
                    continue;
                }
                self.emit_nested_block(child, 2)?;
            }
            self.line(1, "</Section>");
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
                for child in children {
                    self.emit_nested_block(child, depth + 1)?;
                }
                self.line(depth, "</Bookmark>");
                Ok(())
            }
            NestedBlock::ContentControl { tag } => self.body_structured_tag(tag, depth),
        }
    }

    fn body_structured_tag(&mut self, tag: &StructuredDataTag, depth: usize) -> Result<()> {
        if let Some(index) = index_from_sdt(tag) {
            if self.apply_sdt {
                self.body_sdt_index += 1;
            }
            self.advance_paragraphs_in_sdt(tag);
            self.index_component(&index, depth);
            return Ok(());
        }
        self.block_content_control(tag, depth)
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

    fn index_component(&mut self, index: &IndexField, depth: usize) {
        match index {
            IndexField::Contents { start, end } => {
                self.empty(
                    "TableOfContents",
                    depth,
                    &format!(" startLevel={{{start}}} endLevel={{{end}}}"),
                );
            }
            IndexField::Figures { label } => {
                self.empty("TableOfFigures", depth, &attr("label", label));
            }
            IndexField::Entries { identifier } => {
                self.empty("TableOfEntries", depth, &attr("identifier", identifier));
            }
        }
    }

    fn block_content_control(&mut self, tag: &StructuredDataTag, depth: usize) -> Result<()> {
        self.components.insert("ContentControl");
        let mut props = sdt_props_from_tag(tag);
        if self.apply_sdt {
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
        for child in &tag.children {
            match child {
                StructuredDataTagChild::Paragraph(paragraph) => {
                    self.paragraph(paragraph, depth + 1)?;
                }
                StructuredDataTagChild::Table(table) => self.table(table, depth + 1)?,
                StructuredDataTagChild::StructuredDataTag(inner) => {
                    self.block_content_control(inner, depth + 1)?;
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
            if let Some((_, header)) = value {
                self.components.insert("Header");
                self.line(depth, &format!(r#"<Header type="{kind}">"#));
                for child in &header.children {
                    match child {
                        HeaderChild::Paragraph(value) => self.paragraph(value, depth + 1)?,
                        HeaderChild::Table(value) => self.table(value, depth + 1)?,
                        value @ HeaderChild::StructuredDataTag(_) => {
                            return Err(unsupported("header", value));
                        }
                    }
                }
                self.line(depth, "</Header>");
            }
        }
        for (kind, value) in [
            ("default", property.footer.as_ref()),
            ("first", property.first_footer.as_ref()),
            ("even", property.even_footer.as_ref()),
        ] {
            if let Some((_, footer)) = value {
                self.components.insert("Footer");
                self.line(depth, &format!(r#"<Footer type="{kind}">"#));
                for child in &footer.children {
                    match child {
                        FooterChild::Paragraph(value) => {
                            self.footer_paragraph(value, depth + 1)?;
                        }
                        FooterChild::Table(value) => self.table(value, depth + 1)?,
                        value @ FooterChild::StructuredDataTag(_) => {
                            return Err(unsupported("footer", value));
                        }
                    }
                }
                self.line(depth, "</Footer>");
            }
        }
        Ok(())
    }

    fn footer_paragraph(&mut self, paragraph: &docx_rs::Paragraph, depth: usize) -> Result<()> {
        let mut text_box_found = false;
        for child in &paragraph.children {
            let ParagraphChild::Run(run) = child else {
                continue;
            };
            for child in &run.children {
                let RunChild::Drawing(drawing) = child else {
                    continue;
                };
                let Some(DrawingData::TextBox(text_box)) = &drawing.data else {
                    continue;
                };
                text_box_found = true;
                for child in &text_box.children {
                    match child {
                        TextBoxContentChild::Paragraph(value) => self.paragraph(value, depth)?,
                        TextBoxContentChild::Table(value) => self.table(value, depth)?,
                    }
                }
            }
        }
        if !text_box_found {
            self.paragraph(paragraph, depth)?;
        }
        Ok(())
    }

    fn paragraph(&mut self, paragraph: &docx_rs::Paragraph, depth: usize) -> Result<()> {
        self.components.insert("Paragraph");
        let property = json(&paragraph.property)?;
        let mut attrs = Vec::new();
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
        if let Some((name, value)) = reverse_fonts_property(
            property.get("runProperty").unwrap_or(&Value::Null),
            "font",
            "fonts",
        ) {
            attrs.push(jsx_prop(&name, &value)?);
        }
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
                "<Comment{}{}{}>",
                attr("text", &text),
                attr("author", &comment.author),
                attr("date", &comment.date)
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

    fn hyperlink_attrs(&self, value: &Hyperlink) -> Result<String> {
        let data = json(value)?;
        let link = match &value.link {
            HyperlinkData::Anchor { anchor } => attr("anchor", anchor),
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
                attr("href", href)
            }
        };
        let history = data
            .get("history")
            .and_then(Value::as_u64)
            .filter(|value| *value != 0)
            .map_or("", |_| " history");
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
        for child in &value.children {
            match child {
                DeleteChild::Run(run) => self.run(run, depth + 1)?,
                child => return Err(unsupported("deleted revision", child)),
            }
        }
        self.line(depth, "</Deleted>");
        Ok(())
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

    fn run(&mut self, run: &docx_rs::Run, depth: usize) -> Result<()> {
        self.components.insert("Run");
        let property = json(&run.run_property)?;
        let mut attrs = Vec::new();
        for (key, name) in [("bold", "bold"), ("italic", "italic"), ("strike", "strike")] {
            if property_enabled(&property, key) {
                attrs.push(format!(" {name}"));
            }
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
        if let Some((name, value)) = reverse_fonts_property(&property, "font", "fonts") {
            attrs.push(jsx_prop(&name, &value)?);
        }
        if let Some(value) = property
            .get("style")
            .and_then(scalar_string)
            .or_else(|| nested_string(&property, &["style", "val"]))
        {
            attrs.push(attr("style", value));
        }
        self.line(depth, &format!("<Run{}>", attrs.concat()));
        for child in &run.children {
            match child {
                RunChild::Text(value) => self.line(depth + 1, &jsx_string(&value.text)?),
                RunChild::DeleteText(value) => {
                    let data = json(value)?;
                    let text = data.get("text").and_then(Value::as_str).unwrap_or_default();
                    self.line(depth + 1, &jsx_string(text)?);
                }
                RunChild::Tab(_) => self.empty("Tab", depth + 1, ""),
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

    fn drawing(&mut self, drawing: &docx_rs::Drawing, depth: usize) -> Result<()> {
        let Some(data) = &drawing.data else {
            return Err(Error::Reverse(
                "drawing has no representable picture; restore the `w:drawing` content or replace it with a raster image"
                    .to_owned(),
            ));
        };
        match data {
            DrawingData::Pic(picture) => self.image(picture, depth),
            DrawingData::TextBox(_) => Err(unsupported("run", data)),
        }
    }

    fn image(&mut self, picture: &Pic, depth: usize) -> Result<()> {
        let asset = self.images.get(&picture.id).cloned().ok_or_else(|| {
            Error::Reverse(format!(
                "image relationship `{}` is missing media bytes; restore the image part in word/media and its document relationship",
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
        let style = property
            .get("style")
            .and_then(scalar_string)
            .map_or_else(String::new, |value| attr("style", value));
        self.line(depth, &format!("<Table{style}>"));
        for row in &table.rows {
            let TableChild::TableRow(row) = row;
            self.components.insert("TableRow");
            self.line(depth + 1, "<TableRow>");
            for cell in &row.cells {
                let TableRowChild::TableCell(cell) = cell;
                self.components.insert("TableCell");
                self.line(depth + 2, "<TableCell>");
                for child in &cell.children {
                    match child {
                        TableCellContent::Paragraph(value) => self.paragraph(value, depth + 3)?,
                        TableCellContent::Table(value) => self.table(value, depth + 3)?,
                        TableCellContent::StructuredDataTag(tag) => {
                            if let Some(index) = index_from_sdt(tag) {
                                self.index_component(&index, depth + 3);
                            } else {
                                self.cell_content_control(tag, depth + 3)?;
                            }
                        }
                        TableCellContent::TableOfContents(_) => {
                            self.empty("TableOfContents", depth + 3, "");
                        }
                    }
                }
                self.line(depth + 2, "</TableCell>");
            }
            self.line(depth + 1, "</TableRow>");
        }
        self.line(depth, "</Table>");
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
    Contents { start: u64, end: u64 },
    Figures { label: String },
    Entries { identifier: String },
}

fn index_from_sdt(tag: &StructuredDataTag) -> Option<IndexField> {
    let instruction = collect_sdt_instruction(tag);
    parse_toc_instruction(&instruction)
}

fn collect_sdt_instruction(tag: &StructuredDataTag) -> String {
    let mut instruction = String::new();
    for child in &tag.children {
        let StructuredDataTagChild::Paragraph(paragraph) = child else {
            continue;
        };
        for child in &paragraph.children {
            let ParagraphChild::Run(run) = child else {
                continue;
            };
            for child in &run.children {
                if let RunChild::InstrTextString(value) = child {
                    instruction.push_str(value);
                }
            }
        }
    }
    instruction
}

fn parse_toc_instruction(instruction: &str) -> Option<IndexField> {
    let trimmed = instruction.trim();
    if !trimmed.starts_with("TOC") {
        return None;
    }
    if let Some(label) = toc_switch(trimmed, 'c') {
        return Some(IndexField::Figures { label });
    }
    if let Some(identifier) = toc_switch(trimmed, 'f') {
        return Some(IndexField::Entries { identifier });
    }
    let (start, end) = toc_switch(trimmed, 'o')
        .as_deref()
        .and_then(parse_level_range)
        .unwrap_or((1, 3));
    Some(IndexField::Contents { start, end })
}

fn toc_switch(instruction: &str, flag: char) -> Option<String> {
    let marker = format!("\\{flag}");
    let after = instruction.split_once(&marker)?.1.trim_start();
    if after.starts_with('\\') || after.is_empty() {
        return None;
    }
    if let Some(rest) = after.strip_prefix('"') {
        return rest.split_once('"').map(|(value, _)| value.to_owned());
    }
    Some(
        after
            .split(|character: char| character.is_whitespace() || character == '\\')
            .next()
            .unwrap_or_default()
            .to_owned(),
    )
}

fn parse_level_range(value: &str) -> Option<(u64, u64)> {
    let (start, end) = value.split_once('-')?;
    Some((start.parse().ok()?, end.parse().ok()?))
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

fn section_jsx_attrs(property: &docx_rs::SectionProperty) -> Result<String> {
    let data = json(property)?;
    let mut attrs = String::new();
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

fn collect_package_images(bytes: &[u8]) -> Result<HashMap<String, ImageAsset>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| Error::Reverse(format!("invalid DOCX archive: {error}")))?;
    let Ok(mut rels_file) = archive.by_name("word/_rels/document.xml.rels") else {
        return Ok(HashMap::new());
    };
    let mut rels = String::new();
    rels_file
        .read_to_string(&mut rels)
        .map_err(|error| Error::Reverse(format!("cannot read document relationships: {error}")))?;
    drop(rels_file);
    let mut images = HashMap::new();
    let mut used_names = BTreeSet::new();
    for (id, target) in image_relationships(&rels)? {
        let part = if target.starts_with("word/") {
            target.clone()
        } else {
            format!("word/{target}")
        };
        let Ok(mut file) = archive.by_name(&part) else {
            continue;
        };
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)
            .map_err(|error| Error::Reverse(format!("cannot read image part `{part}`: {error}")))?;
        drop(file);
        let file_name = Path::new(&target)
            .file_name()
            .and_then(|value| value.to_str())
            .filter(|value| !value.is_empty())
            .unwrap_or("image.bin");
        let unique = unique_asset_name(file_name, &used_names);
        used_names.insert(unique.clone());
        images.insert(
            id,
            ImageAsset {
                src: format!("media/{unique}"),
                bytes,
            },
        );
    }
    Ok(images)
}

fn image_relationships(rels: &str) -> Result<Vec<(String, String)>> {
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
                    .is_some_and(|value| value.ends_with("/image"))
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
                    "cannot parse document relationships: {error}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(output)
}

fn collect_images(docx: &docx_rs::Docx) -> HashMap<String, ImageAsset> {
    let mut images = HashMap::new();
    let mut used_names = BTreeSet::new();
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
    let mut attrs = String::new();
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

fn format_points(twips: u64) -> String {
    let whole = twips / 20;
    let remainder = twips % 20;
    if remainder == 0 {
        whole.to_string()
    } else {
        format!("{}.{}", whole, remainder * 5)
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
) -> Result<String> {
    let mut attributes = String::new();
    let defaults = json(&docx.styles.doc_defaults)?;
    if let Some(fonts) = nested(&defaults, &["runPropertyDefault", "runProperty", "fonts"])
        .and_then(Value::as_object)
    {
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
        if !output.is_empty() {
            let physical = ["ascii", "hiAnsi", "eastAsia", "cs"]
                .map(|key| output.get(key).and_then(Value::as_str));
            let can_collapse = physical[0].is_some()
                && physical.iter().all(|value| *value == physical[0])
                && !output
                    .keys()
                    .any(|key| key.ends_with("Theme") || key == "hint");
            if can_collapse {
                attributes.push_str(&attr("defaultFont", physical[0].unwrap_or_default()));
            } else {
                let fonts = serde_json::to_string(&output).map_err(|error| {
                    Error::Reverse(format!("cannot serialize default font slots: {error}"))
                })?;
                write!(attributes, " defaultFonts={{{fonts}}}")
                    .expect("writing to a String cannot fail");
            }
        }
    }
    let styles = docx
        .styles
        .styles
        .iter()
        .map(|style| {
            reverse_style_definition(
                style,
                metadata.get(&style.style_id),
                writer.style_theme_colors.get(&style.style_id),
            )
        })
        .collect::<Result<Vec<_>>>()?
        .into_iter()
        .flatten()
        .collect::<Vec<_>>();
    if styles.is_empty() {
        return Ok(attributes);
    }
    let styles = serde_json::to_string(&styles)
        .map_err(|error| Error::Reverse(format!("cannot serialize style definitions: {error}")))?;
    write!(attributes, " styles={{{styles}}}").expect("writing to a String cannot fail");
    Ok(attributes)
}

fn reverse_style_definition(
    style: &docx_rs::Style,
    metadata: Option<&StyleMetadata>,
    theme: Option<&ThemeColorInfo>,
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
    if let Some(mut run) = reverse_run_properties(&source["runProperty"])? {
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
    if let Some(paragraph) = reverse_paragraph_properties(&source["paragraphProperty"])? {
        definition.insert("paragraph".to_owned(), Value::Object(paragraph));
    }
    if is_stock_normal(&definition) {
        return Ok(None);
    }
    Ok(Some(Value::Object(definition)))
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

fn xml_attribute(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::Reverse(format!("invalid attribute in word/styles.xml: {error}"))
        })?;
        if attribute.key.local_name().as_ref() == name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::Reverse(format!("invalid text in word/styles.xml: {error}"))
                });
        }
    }
    Ok(None)
}

fn reverse_run_properties(source: &Value) -> Result<Option<Map<String, Value>>> {
    let mut output = Map::new();
    for (source_key, target_key) in [("bold", "bold"), ("italic", "italic"), ("vanish", "hidden")] {
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
    Ok((!output.is_empty()).then_some(output))
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
    if source.get("keepNext").and_then(Value::as_bool) == Some(true) {
        output.insert("keepNext".to_owned(), Value::Bool(true));
    }
    if source.get("keepLines").and_then(Value::as_bool) == Some(true) {
        output.insert("keepLines".to_owned(), Value::Bool(true));
    }
    if let Some(value) = source
        .get("outlineLvl")
        .and_then(Value::as_u64)
        .or_else(|| source.get("outlineLvl").and_then(scalar_u64))
    {
        output.insert("outlineLevel".to_owned(), Value::Number(value.into()));
    }
    Ok((!output.is_empty()).then_some(output))
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
        let sign = if value < 0 && whole == 0 { "-" } else { "" };
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

#[cfg(test)]
mod tests {
    use super::*;
    use docx_rs::{
        Comment, Delete, DocGrid, Docx, Footer, Header, Hyperlink, HyperlinkType, Insert,
        LineSpacing, LineSpacingType, PageMargin, PageOrientationType, PageSize, Paragraph, Pic,
        Run, RunFonts, Section, SpecialIndentType, StructuredDataTag, Style, StyleType, Table,
        TableCell, TableOfContents, TableRow, ThemeColor,
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

        assert!(jsx.contains(r#"<Table style="Grid">"#), "{jsx}");
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

    fn pack_docx(docx: Docx) -> Vec<u8> {
        let mut bytes = Cursor::new(Vec::new());
        docx.build()
            .pack(&mut bytes)
            .expect("external DOCX fixture should pack");
        bytes.into_inner()
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
            jsx.contains(r#"<Comment text="review" author="Grace" date="2026-08-14T00:00:00Z">"#)
                && jsx.contains(r#"{"commented"}"#)
                && jsx.contains("</Comment>"),
            "{jsx}"
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

        let jsx = reverse_document(&bytes).expect("external DOCX should reverse");

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
        let png = {
            let mut encoded = Cursor::new(Vec::new());
            image::DynamicImage::new_rgba8(2, 1)
                .write_to(&mut encoded, image::ImageFormat::Png)
                .expect("png should encode");
            encoded.into_inner()
        };
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
