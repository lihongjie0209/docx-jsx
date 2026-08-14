//! DOCX to deterministic, recompilable JSX conversion.

use std::collections::BTreeSet;
use std::io::{Cursor, Read};

use docx_rs::{
    DocumentChild, ParagraphChild, RunChild, TableCellContent, TableChild, TableRowChild,
};
use serde::Serialize;
use serde_json::Value;

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
    let mut writer = Writer::default();
    writer.components.insert("Document");
    writer.components.insert("Section");
    writer.line(0, "<Document>");
    writer.line(1, "<Section>");
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
            value => Err(unsupported("document", value)),
        }
    }

    fn paragraph(&mut self, paragraph: &docx_rs::Paragraph, depth: usize) -> Result<()> {
        self.components.insert("Paragraph");
        let property = json(&paragraph.property)?;
        let mut attrs = Vec::new();
        if let Some(value) = property.get("alignment").and_then(scalar_string) {
            attrs.push(attr("align", value));
        }
        if let Some(value) = nested_string(&property, &["style", "val"])
            .or_else(|| nested_string(&property, &["style", "styleId"]))
        {
            attrs.push(attr("style", value));
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
        if let Some(value) = property.get("color").and_then(scalar_string) {
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
        if let Some(value) = nested_string(&property, &["style", "val"]) {
            attrs.push(attr("style", value));
        }
        self.line(depth, &format!("<Run{}>", attrs.concat()));
        for child in &run.children {
            match child {
                RunChild::Text(value) => self.line(depth + 1, &escape_jsx(&value.text)),
                RunChild::DeleteText(value) => {
                    let data = json(value)?;
                    let text = data.get("text").and_then(Value::as_str).unwrap_or_default();
                    self.line(depth + 1, &escape_jsx(text));
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
                value => return Err(unsupported("run", value)),
            }
        }
        self.line(depth, "</Run>");
        Ok(())
    }

    fn table(&mut self, table: &docx_rs::Table, depth: usize) -> Result<()> {
        self.components.insert("Table");
        self.line(depth, "<Table>");
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

fn escape_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn escape_jsx(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('{', "&#123;")
        .replace('}', "&#125;")
}

fn unsupported(context: &str, value: &impl std::fmt::Debug) -> Error {
    Error::Reverse(format!("unsupported {context} structure: {value:?}"))
}
