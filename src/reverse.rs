//! DOCX to deterministic, recompilable JSX conversion.

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::fmt::Write as _;
use std::io::{Cursor, Read};
use std::path::Path;

use docx_rs::{
    AbstractNumbering, Comment, CommentChild, Delete, DeleteChild, DocumentChild, DrawingData,
    DrawingPosition, DrawingPositionType, FooterChild, HeaderChild, Hyperlink, HyperlinkData,
    Insert, InsertChild, Level, MoveFrom, MoveFromChild, MoveTo, MoveToChild, Numbering,
    ParagraphChild, Pic, PicAlign, RelativeFromHType, RelativeFromVType, RunChild,
    StructuredDataTag, StructuredDataTagChild, TableCellContent, TableChild, TableRowChild,
    TextBoxContentChild,
};
use num_traits::ToPrimitive;
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
    validate_character_indent_xml(bytes)?;
    validate_simple_positioned_images(bytes)?;
    let style_metadata = read_style_metadata(bytes)?;
    let theme_colors = parse_theme_colors(bytes)?;
    let section_columns = collect_section_columns(bytes)?;
    let table_row_xml = collect_table_row_xml_metadata(bytes)?;
    let table_xml = collect_table_xml_metadata(bytes)?;
    let mut images = collect_package_images(bytes)?;
    for (id, asset) in collect_images(&docx, &images.used_names) {
        images.by_id.entry(id).or_insert(asset);
    }
    let mut writer = Writer {
        images,
        image_scope: Some("word/document.xml".to_owned()),
        header_footer_parts: collect_header_footer_parts(bytes)?,
        style_theme_colors: theme_colors.styles,
        run_theme_colors: theme_colors.runs,
        abstract_nums: docx.numberings.abstract_nums.clone(),
        numbering_instances: docx.numberings.numberings.clone(),
        section_columns,
        table_row_xml,
        table_xml,
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
    let document_attrs = reverse_document_attributes(&docx, &style_metadata, &writer, bytes)?;
    writer.line(0, &format!("<Document{document_attrs}>"));
    writer.emit_sections(&docx)?;
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
    images: ImageCatalog,
    image_scope: Option<String>,
    header_footer_parts: HashMap<String, String>,
    assets: BTreeMap<String, Vec<u8>>,
    style_theme_colors: HashMap<String, ThemeColorInfo>,
    run_theme_colors: Vec<ThemeColorInfo>,
    run_theme_index: usize,
    abstract_nums: Vec<AbstractNumbering>,
    numbering_instances: Vec<Numbering>,
    section_columns: Vec<usize>,
    table_row_xml: HashMap<String, Vec<TableRowXmlMetadata>>,
    table_row_indices: HashMap<String, usize>,
    table_xml: HashMap<String, Vec<TableXmlMetadata>>,
    table_indices: HashMap<String, usize>,
}

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
        for (index, (children, break_property)) in groups.into_iter().enumerate() {
            let property = break_property.unwrap_or(&docx.document.section_property);
            let columns = self.section_columns.get(index).copied().unwrap_or(1);
            self.line(
                1,
                &format!("<Section{}>", section_jsx_attrs(property, columns)?),
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
            DrawingData::TextBox(_) => Err(unsupported("run", data)),
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
                value @ HeaderChild::StructuredDataTag(_) => {
                    return Err(unsupported("header", value));
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
                value @ FooterChild::StructuredDataTag(_) => {
                    return Err(unsupported("footer", value));
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
                    if let Some(kind) = index_from_sdt(tag) {
                        self.index_component(&kind, depth);
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
        if instance
            .level_overrides
            .iter()
            .any(|value| value.override_level.is_some())
        {
            return Err(Error::Reverse(
                "numbering uses a full level override; flatten the override into the abstract definition or simplify the list before reversing"
                    .to_owned(),
            ));
        }
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
        if definition.style_link.is_some()
            || definition.num_style_link.is_some()
            || definition.levels.is_empty()
        {
            return Err(Error::Reverse(
                "numbering is a style-linked or empty abstract definition; expand `numStyleLink`/`styleLink` to explicit `w:lvl` entries before reversing"
                    .to_owned(),
            ));
        }
        let starts = instance
            .level_overrides
            .iter()
            .filter_map(|value| value.override_start.map(|start| (value.level, start)))
            .collect::<HashMap<_, _>>();
        reverse_list_definition(definition, &starts)
    }

    fn list_item(&mut self, paragraph: &docx_rs::Paragraph, depth: usize) -> Result<()> {
        let property = json(&paragraph.property)?;
        let mut attrs = String::new();
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

fn validate_simple_positioned_images(bytes: &[u8]) -> Result<()> {
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
                    if element.local_name().as_ref() == b"anchor" =>
                {
                    if xml_attribute(&reader, &element, b"simplePos")?
                        .is_some_and(|value| !matches!(value.as_str(), "0" | "false" | "off"))
                    {
                        return Err(Error::Reverse(format!(
                            "{part} contains `wp:anchor/@simplePos=1`, which the current JSX model cannot represent; convert the placement to `positionH`/`positionV` or retain the DOCX"
                        )));
                    }
                }
                Ok(Event::Eof) => break,
                Err(error) => {
                    return Err(Error::Reverse(format!(
                        "cannot parse {part} drawing placement at byte {}: {error}",
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

fn collect_section_columns(bytes: &[u8]) -> Result<Vec<usize>> {
    let xml = read_zip_text(bytes, "word/document.xml")?;
    let mut reader = Reader::from_str(&xml);
    let mut buffer = Vec::new();
    let mut current_columns = None;
    let mut inside_columns = false;
    let mut saw_columns = false;
    let mut sections = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"sectPr" => {
                if current_columns.replace(1).is_some() {
                    return Err(Error::Reverse(
                        "word/document.xml contains nested section properties; repair the malformed `w:sectPr` structure".to_owned(),
                    ));
                }
                saw_columns = false;
            }
            Ok(Event::Empty(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(1);
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
                }
            }
            Ok(Event::Start(element) | Event::Empty(element)) if inside_columns => {
                let name = element.local_name();
                let error = if name.as_ref() == b"col" {
                    "section uses individually sized columns, which cannot be represented by the equal-width `columns` property; use equal-width columns or keep the original DOCX".to_owned()
                } else {
                    format!(
                        "section `w:cols` contains unsupported `<{}>` data; remove the custom column data or keep the original DOCX",
                        String::from_utf8_lossy(name.as_ref())
                    )
                };
                return Err(Error::Reverse(error));
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"cols" => {
                inside_columns = false;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sectPr" => {
                sections.push(current_columns.take().unwrap_or(1));
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

fn parse_section_columns(reader: &Reader<&[u8]>, element: &BytesStart<'_>) -> Result<usize> {
    let columns = xml_attribute(reader, element, b"num")?
        .map(|value| value.parse::<usize>())
        .transpose()
        .map_err(|_| {
            Error::Reverse(
                "section column count is not an integer; repair `w:cols/@w:num`".to_owned(),
            )
        })?
        .unwrap_or(1);
    if !(1..=45).contains(&columns) {
        return Err(Error::Reverse(format!(
            "section column count `{columns}` is outside the supported range 1–45; set `w:cols/@w:num` to a supported count"
        )));
    }
    let separator = xml_attribute(reader, element, b"sep")?;
    if separator
        .as_deref()
        .is_some_and(|value| !matches!(value, "0" | "false" | "off"))
    {
        return Err(Error::Reverse(
            "section columns use a separator line, which is not represented by the `columns` property; remove the separator or keep the original DOCX".to_owned(),
        ));
    }
    let equal_width = xml_attribute(reader, element, b"equalWidth")?;
    if equal_width
        .as_deref()
        .is_some_and(|value| matches!(value, "0" | "false" | "off"))
    {
        return Err(Error::Reverse(
            "section columns are not equal-width, which cannot be represented by `columns`; use equal-width columns or keep the original DOCX".to_owned(),
        ));
    }
    if columns > 1 {
        let space = xml_attribute(reader, element, b"space")?;
        if space.as_deref() != Some("425") {
            return Err(Error::Reverse(format!(
                "section column spacing `{}` twips is not representable; use the backend default of 425 twips or keep the original DOCX",
                space.as_deref().unwrap_or("unspecified")
            )));
        }
    }
    Ok(columns)
}

fn section_jsx_attrs(property: &docx_rs::SectionProperty, columns: usize) -> Result<String> {
    let data = json(property)?;
    let mut attrs = String::new();
    if columns != 1 {
        let _ = write!(attrs, " columns={{{columns}}}");
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
    start_overrides: &HashMap<usize, usize>,
) -> Result<String> {
    if let Some(attrs) = collapse_list_shorthand(definition, start_overrides) {
        return Ok(attrs);
    }
    let mut levels = Vec::new();
    for level in &definition.levels {
        levels.push(reverse_level_definition(level, start_overrides)?);
    }
    let encoded = serde_json::to_string(&levels)
        .map_err(|error| Error::Reverse(format!("cannot serialize list levels: {error}")))?;
    Ok(format!(" levels={{{encoded}}}"))
}

fn collapse_list_shorthand(
    definition: &AbstractNumbering,
    start_overrides: &HashMap<usize, usize>,
) -> Option<String> {
    if definition.levels.len() != 9 {
        return None;
    }
    let first = definition.levels.first()?;
    let start = start_overrides
        .get(&first.level)
        .copied()
        .or_else(|| {
            json(&first.start)
                .ok()
                .and_then(|value| value.as_u64())
                .and_then(|value| usize::try_from(value).ok())
        })
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
        let level_start = start_overrides
            .get(&level.level)
            .copied()
            .or_else(|| {
                json(&level.start)
                    .ok()
                    .and_then(|value| value.as_u64())
                    .and_then(|value| usize::try_from(value).ok())
            })
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

fn reverse_level_definition(
    level: &Level,
    start_overrides: &HashMap<usize, usize>,
) -> Result<Value> {
    let source = json(level)?;
    let mut output = Map::new();
    if let Some(value) = source.get("format").and_then(Value::as_str) {
        output.insert("format".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source.get("text").and_then(Value::as_str) {
        output.insert("text".to_owned(), Value::String(value.to_owned()));
    }
    let start = start_overrides.get(&level.level).copied().or_else(|| {
        source
            .get("start")
            .and_then(Value::as_u64)
            .and_then(|value| usize::try_from(value).ok())
    });
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
    bytes: &[u8],
) -> Result<String> {
    let mut attributes = String::new();
    attributes.push_str(&reverse_default_font_attrs(docx)?);
    attributes.push_str(&reverse_doc_default_metrics(docx)?);
    attributes.push_str(&reverse_core_and_custom_properties(docx, bytes)?);
    attributes.push_str(&reverse_settings_attributes(&docx.settings)?);
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
    let (created, updated) = parse_core_timestamps(bytes)?;
    if let Some(created) = created {
        attributes.push_str(&attr("createdAt", &created));
    }
    if let Some(updated) = updated {
        attributes.push_str(&attr("updatedAt", &updated));
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

fn reverse_settings_attributes(settings: &docx_rs::Settings) -> Result<String> {
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

fn parse_core_timestamps(bytes: &[u8]) -> Result<(Option<String>, Option<String>)> {
    let Ok(xml) = read_zip_text(bytes, "docProps/core.xml") else {
        return Ok((None, None));
    };
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);
    let mut created = None;
    let mut modified = None;
    let mut current = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"created" => current = Some("created"),
                b"modified" => current = Some("modified"),
                _ => {}
            },
            Ok(Event::Text(text)) => {
                let value = text.decode().map_err(|error| {
                    Error::Reverse(format!("cannot decode docProps/core.xml: {error}"))
                })?;
                match current {
                    Some("created") => created = Some(value.into_owned()),
                    Some("modified") => modified = Some(value.into_owned()),
                    _ => {}
                }
            }
            Ok(Event::End(element))
                if matches!(element.local_name().as_ref(), b"created" | b"modified") =>
            {
                current = None;
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
    Ok((
        created.filter(|value| value != "1970-01-01T00:00:00Z"),
        modified.filter(|value| value != "1970-01-01T00:00:00Z"),
    ))
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
        AbstractNumbering, BorderType, CellMargins, CharacterSpacingValues, Comment, Delete,
        DocGrid, Docx, Footer, Header, HeightRule, Hyperlink, HyperlinkType, IndentLevel, Insert,
        Level, LevelJc, LevelSuffixType, LevelText, LineSpacing, LineSpacingType, NumberFormat,
        Numbering, NumberingId, PageMargin, PageOrientationType, PageSize, Paragraph, Pic, Run,
        RunFonts, Section, Settings, Shading, ShdType, SpecialIndentType, Start, StructuredDataTag,
        Style, StyleType, Table, TableAlignmentType, TableBorder, TableBorderPosition,
        TableBorders, TableCell, TableCellBorder, TableCellBorderPosition, TableCellBorders,
        TableCellMargins, TableCellProperty, TableLayoutType, TableOfContents,
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
    fn reverse_external_docx_should_reject_character_unit_indentation() {
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
            r#"<w:pPr><w:ind w:firstLineChars="200" /></w:pPr>"#,
        );
        assert!(document.contains("w:firstLineChars"), "{document}");
        let modified =
            rewrite_package_parts(&packed, &[("word/document.xml", document.as_bytes())]);
        let error = reverse_document(&modified)
            .expect_err("character-unit indentation is not representable by JSX points");
        assert!(
            error.to_string().contains("firstLineChars")
                && error.to_string().contains("points")
                && error.to_string().contains("replace it"),
            "diagnostic should identify the unsupported unit and a repair: {error}"
        );
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
    fn reverse_should_reject_unrepresented_section_column_spacing() {
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
        let error =
            reverse_document(&modified).expect_err("custom column spacing must not be dropped");
        assert!(
            error.to_string().contains("spacing") && error.to_string().contains("425 twips"),
            "diagnostic should name the unsupported value and repair: {error}"
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
    fn reverse_external_docx_should_reject_simple_positioned_images() {
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

        let error = reverse_document(&modified)
            .expect_err("simple-position coordinates are not representable in JSX");

        assert!(
            error.to_string().contains("simplePos")
                && error.to_string().contains("positionH")
                && error.to_string().contains("retain the DOCX"),
            "diagnostic should name the unsupported placement and repair: {error}"
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
        let bytes = pack_docx(
            Docx::new()
                .add_abstract_numbering(default_list_abstract(2, true))
                .add_numbering(Numbering::new(2, 2))
                .add_paragraph(numbered(2, 0, "one"))
                .add_paragraph(numbered(2, 1, "two")),
        );
        let jsx = reverse_document(&bytes).expect("external list should reverse");
        assert!(
            jsx.contains(r#"<List type="ordered">"#)
                && jsx.contains("<ListItem>")
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
    fn reverse_external_docx_should_reject_style_linked_numbering() {
        let bytes = pack_docx(
            Docx::new()
                .add_abstract_numbering(AbstractNumbering::new(5).num_style_link("ListNumber"))
                .add_numbering(Numbering::new(5, 5))
                .add_paragraph(numbered(5, 0, "item")),
        );
        let error = reverse_document(&bytes).expect_err("style-linked numbering must fail");
        assert!(
            error.to_string().contains("style-linked")
                && error.to_string().contains("explicit `w:lvl`"),
            "{error}"
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
