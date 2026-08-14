//! DOCX to deterministic, recompilable JSX conversion.

use std::collections::{BTreeSet, HashMap};
use std::io::{Cursor, Read};

use docx_rs::{
    Delete, DeleteChild, DocumentChild, DrawingData, FooterChild, HeaderChild, Insert, InsertChild,
    MoveFrom, MoveFromChild, MoveTo, MoveToChild, ParagraphChild, RunChild, TableCellContent,
    TableChild, TableRowChild, TextBoxContentChild,
};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use serde::Serialize;
use serde_json::{Map, Number, Value};

use crate::ir::{Child, IrEnvelope, Node};
use crate::{Error, Result};

/// Converts DOCX package bytes into a deterministic JSX module.
///
/// # Errors
///
/// Returns an error when the package is invalid or contains a structure that
/// the JSX component model cannot currently represent without data loss.
pub fn reverse_document(bytes: &[u8]) -> Result<String> {
    if let Some(ir) = read_ir_manifest(bytes)? {
        ir.validate()?;
        return ir_to_jsx(&ir);
    }
    let docx = docx_rs::read_docx(bytes).map_err(|error| Error::Reverse(error.to_string()))?;
    let style_metadata = read_style_metadata(bytes)?;
    let mut writer = Writer::default();
    writer.components.insert("Document");
    writer.components.insert("Section");
    let document_attrs = reverse_document_attributes(&docx, &style_metadata)?;
    writer.line(0, &format!("<Document{document_attrs}>"));
    writer.line(1, "<Section>");
    writer.headers_and_footers(&docx.document.section_property, 2)?;
    for child in &docx.document.children {
        writer.document_child(child, 2)?;
    }
    writer.line(1, "</Section>");
    writer.line(0, "</Document>");

    let imports = writer
        .components
        .iter()
        .copied()
        .collect::<Vec<_>>()
        .join(", ");
    Ok(format!(
        "import {{ {imports} }} from \"docx-jsx\";\n\nexport default (\n{});\n",
        writer.output
    ))
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
    ignored_bookmark_ids: BTreeSet<usize>,
}

impl Writer {
    fn line(&mut self, depth: usize, value: &str) {
        self.output.push_str(&"  ".repeat(depth));
        self.output.push_str(value);
        self.output.push('\n');
    }

    fn document_child(&mut self, child: &DocumentChild, depth: usize) -> Result<()> {
        match child {
            DocumentChild::Paragraph(value) => self.paragraph(value, depth),
            DocumentChild::Table(value) => self.table(value, depth),
            DocumentChild::BookmarkStart(value) => {
                self.components.insert("Bookmark");
                self.ignored_bookmark_ids.insert(value.id);
                self.line(depth, &format!("<Bookmark{} />", attr("name", &value.name)));
                Ok(())
            }
            DocumentChild::BookmarkEnd(value) if self.ignored_bookmark_ids.remove(&value.id) => {
                Ok(())
            }
            value => Err(unsupported("document", value)),
        }
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
        self.line(depth, &format!("<Paragraph{}>", attrs.concat()));
        for child in &paragraph.children {
            self.paragraph_child(child, depth + 1)?;
        }
        self.line(depth, "</Paragraph>");
        Ok(())
    }

    fn paragraph_child(&mut self, child: &ParagraphChild, depth: usize) -> Result<()> {
        match child {
            ParagraphChild::Run(value) => self.run(value, depth),
            ParagraphChild::Insert(value) => self.inserted(value, depth),
            ParagraphChild::Delete(value) => self.deleted(value, depth),
            ParagraphChild::MoveFrom(value) => self.moved_from(value, depth),
            ParagraphChild::MoveTo(value) => self.moved_to(value, depth),
            ParagraphChild::BookmarkStart(value) if value.name == "_GoBack" => {
                self.ignored_bookmark_ids.insert(value.id);
                Ok(())
            }
            ParagraphChild::BookmarkStart(value) => {
                self.components.insert("InlineBookmark");
                self.ignored_bookmark_ids.insert(value.id);
                self.line(
                    depth,
                    &format!("<InlineBookmark{} />", attr("name", &value.name)),
                );
                Ok(())
            }
            ParagraphChild::BookmarkEnd(value) if self.ignored_bookmark_ids.remove(&value.id) => {
                Ok(())
            }
            ParagraphChild::Hyperlink(value) => {
                self.components.insert("Hyperlink");
                let data = json(value)?;
                let link = if let Some(href) = data.get("path").and_then(Value::as_str) {
                    attr("href", href)
                } else if let Some(anchor) = data.get("anchor").and_then(Value::as_str) {
                    attr("anchor", anchor)
                } else {
                    return Err(Error::Reverse("hyperlink target is missing".to_owned()));
                };
                let history = data
                    .get("history")
                    .and_then(Value::as_u64)
                    .filter(|value| *value != 0)
                    .map_or("", |_| " history");
                self.line(depth, &format!("<Hyperlink{link}{history}>"));
                for child in &value.children {
                    self.paragraph_child(child, depth + 1)?;
                }
                self.line(depth, "</Hyperlink>");
                Ok(())
            }
            ParagraphChild::PageNum(_) => {
                self.empty("PageNumber", depth, "");
                Ok(())
            }
            ParagraphChild::NumPages(_) => {
                self.empty("TotalPages", depth, "");
                Ok(())
            }
            value => Err(unsupported("paragraph", value)),
        }
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
        }
        if let Some(value) = property.get("sz").and_then(scalar_u64) {
            let size = if value % 2 == 0 {
                (value / 2).to_string()
            } else {
                format!("{}.5", value / 2)
            };
            attrs.push(format!(" size={{{size}}}"));
        }
        if let Some(value) = nested_string(&property, &["fonts", "ascii"])
            .or_else(|| nested_string(&property, &["fonts", "eastAsia"]))
        {
            attrs.push(attr("font", value));
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
                value => return Err(unsupported("run", value)),
            }
        }
        self.line(depth, "</Run>");
        Ok(())
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
                        value => return Err(unsupported("table cell", value)),
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
) -> Result<String> {
    let styles = docx
        .styles
        .styles
        .iter()
        .map(|style| reverse_style_definition(style, metadata.get(&style.style_id)))
        .collect::<Result<Vec<_>>>()?;
    if styles.is_empty() {
        return Ok(String::new());
    }
    let styles = serde_json::to_string(&styles)
        .map_err(|error| Error::Reverse(format!("cannot serialize style definitions: {error}")))?;
    Ok(format!(" styles={{{styles}}}"))
}

fn reverse_style_definition(
    style: &docx_rs::Style,
    metadata: Option<&StyleMetadata>,
) -> Result<Value> {
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
    if let Some(run) = reverse_run_properties(&source["runProperty"])? {
        definition.insert("run".to_owned(), Value::Object(run));
    }
    if let Some(paragraph) = reverse_paragraph_properties(&source["paragraphProperty"])? {
        definition.insert("paragraph".to_owned(), Value::Object(paragraph));
    }
    Ok(Value::Object(definition))
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
    if let Some(value) = nested_string(source, &["fonts", "ascii"])
        .or_else(|| nested_string(source, &["fonts", "eastAsia"]))
    {
        output.insert("font".to_owned(), Value::String(value.to_owned()));
    }
    if let Some(value) = source.get("underline").and_then(scalar_string) {
        output.insert("underline".to_owned(), Value::String(value.to_owned()));
    }
    Ok((!output.is_empty()).then_some(output))
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
        Docx, Footer, LineSpacing, LineSpacingType, Paragraph, Run, SpecialIndentType, Style,
        StyleType, Table, TableCell, TableRow,
    };

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
}
