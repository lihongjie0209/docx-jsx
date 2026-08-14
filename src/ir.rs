use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};
use std::collections::HashSet;

use crate::error::{Error, Result};

/// Versioned data exchanged between the JavaScript runtime and Rust compiler.
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct IrEnvelope {
    pub version: u8,
    pub document: Node,
}

/// A normalized JSX node.
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct Node {
    #[serde(rename = "type")]
    pub kind: NodeKind,
    #[serde(default)]
    pub props: Map<String, Value>,
    #[serde(default)]
    pub children: Vec<Child>,
}

/// Built-in component kinds supported by IR v1.
#[derive(Debug, Clone, Copy, Deserialize, Serialize, PartialEq, Eq)]
pub enum NodeKind {
    Document,
    Section,
    Paragraph,
    Run,
    Text,
    Break,
    CarriageReturn,
    NonBreakingSpace,
    SoftHyphen,
    NonBreakingHyphen,
    Image,
    Table,
    TableRow,
    TableCell,
    Header,
    Footer,
    Hyperlink,
    PageNumber,
    TotalPages,
    List,
    ListItem,
    Heading,
    Caption,
    Index,
    Bookmark,
    TableOfContents,
    TableOfFigures,
    TableOfEntries,
    TocEntry,
    IndexEntry,
    Comment,
    Footnote,
    Tab,
    TabStop,
    Symbol,
    Bold,
    Italic,
    Underline,
    StrikeThrough,
    Superscript,
    Subscript,
    AllCaps,
    HiddenText,
    SpecialHiddenText,
    DoubleStrike,
    SpacedText,
    ScaledText,
    FitText,
    BorderedText,
    ShadedText,
    Inserted,
    Deleted,
    MovedFrom,
    MovedTo,
    PageReference,
    PositionalTab,
    ContentControl,
    Field,
    DateField,
    TimeField,
    FileNameField,
    AuthorField,
    TitleField,
    SubjectField,
    SequenceField,
    ReferenceField,
    MergeField,
    DocumentPropertyField,
    FormulaField,
}

impl NodeKind {
    #[must_use]
    pub fn name(self) -> &'static str {
        match self {
            Self::Document => "Document",
            Self::Section => "Section",
            Self::Paragraph => "Paragraph",
            Self::Run => "Run",
            Self::Text => "Text",
            Self::Break => "Break",
            Self::CarriageReturn => "CarriageReturn",
            Self::NonBreakingSpace => "NonBreakingSpace",
            Self::SoftHyphen => "SoftHyphen",
            Self::NonBreakingHyphen => "NonBreakingHyphen",
            Self::Image => "Image",
            Self::Table => "Table",
            Self::TableRow => "TableRow",
            Self::TableCell => "TableCell",
            Self::Header => "Header",
            Self::Footer => "Footer",
            Self::Hyperlink => "Hyperlink",
            Self::PageNumber => "PageNumber",
            Self::TotalPages => "TotalPages",
            Self::List => "List",
            Self::ListItem => "ListItem",
            Self::Heading => "Heading",
            Self::Caption => "Caption",
            Self::Index => "Index",
            Self::Bookmark => "Bookmark",
            Self::TableOfContents => "TableOfContents",
            Self::TableOfFigures => "TableOfFigures",
            Self::TableOfEntries => "TableOfEntries",
            Self::TocEntry => "TocEntry",
            Self::IndexEntry => "IndexEntry",
            Self::Comment => "Comment",
            Self::Footnote => "Footnote",
            Self::Tab => "Tab",
            Self::TabStop => "TabStop",
            Self::Symbol => "Symbol",
            Self::Bold => "Bold",
            Self::Italic => "Italic",
            Self::Underline => "Underline",
            Self::StrikeThrough => "StrikeThrough",
            Self::Superscript => "Superscript",
            Self::Subscript => "Subscript",
            Self::AllCaps => "AllCaps",
            Self::HiddenText => "HiddenText",
            Self::SpecialHiddenText => "SpecialHiddenText",
            Self::DoubleStrike => "DoubleStrike",
            Self::SpacedText => "SpacedText",
            Self::ScaledText => "ScaledText",
            Self::FitText => "FitText",
            Self::BorderedText => "BorderedText",
            Self::ShadedText => "ShadedText",
            Self::Inserted => "Inserted",
            Self::Deleted => "Deleted",
            Self::MovedFrom => "MovedFrom",
            Self::MovedTo => "MovedTo",
            Self::PageReference => "PageReference",
            Self::PositionalTab => "PositionalTab",
            Self::ContentControl => "ContentControl",
            Self::Field => "Field",
            Self::DateField => "DateField",
            Self::TimeField => "TimeField",
            Self::FileNameField => "FileNameField",
            Self::AuthorField => "AuthorField",
            Self::TitleField => "TitleField",
            Self::SubjectField => "SubjectField",
            Self::SequenceField => "SequenceField",
            Self::ReferenceField => "ReferenceField",
            Self::MergeField => "MergeField",
            Self::DocumentPropertyField => "DocumentPropertyField",
            Self::FormulaField => "FormulaField",
        }
    }

    pub(crate) fn is_semantic_text(self) -> bool {
        matches!(
            self,
            Self::Bold
                | Self::Italic
                | Self::Underline
                | Self::StrikeThrough
                | Self::Superscript
                | Self::Subscript
                | Self::AllCaps
                | Self::HiddenText
                | Self::SpecialHiddenText
                | Self::DoubleStrike
                | Self::SpacedText
                | Self::ScaledText
                | Self::FitText
                | Self::BorderedText
                | Self::ShadedText
        )
    }

    pub(crate) fn is_field(self) -> bool {
        matches!(
            self,
            Self::Field
                | Self::DateField
                | Self::TimeField
                | Self::FileNameField
                | Self::AuthorField
                | Self::TitleField
                | Self::SubjectField
                | Self::SequenceField
                | Self::IndexEntry
                | Self::ReferenceField
                | Self::MergeField
                | Self::DocumentPropertyField
                | Self::FormulaField
        )
    }

    fn is_special_character(self) -> bool {
        matches!(
            self,
            Self::NonBreakingSpace | Self::SoftHyphen | Self::NonBreakingHyphen
        )
    }
}

/// A normalized component or primitive text child.
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(untagged)]
pub enum Child {
    Node(Node),
    String(String),
    Number(serde_json::Number),
}

impl IrEnvelope {
    /// Validates the IR version, root, child matrix, properties, and scalar values.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Validation`] for an unsupported version, invalid root,
    /// invalid nesting, unknown property, or invalid scalar constraint.
    pub fn validate(&self) -> Result<()> {
        if self.version != 1 {
            return Err(validation("Document", "unsupported IR version"));
        }
        if self.document.kind != NodeKind::Document {
            return Err(validation("Document", "root must be Document"));
        }
        validate_node(&self.document, "Document")?;
        validate_bookmark_references(&self.document)
    }
}

fn validate_bookmark_references(root: &Node) -> Result<()> {
    fn collect<'a>(node: &'a Node, names: &mut Vec<&'a str>, anchors: &mut Vec<&'a str>) {
        if node.kind == NodeKind::Bookmark
            && let Some(name) = node.props.get("name").and_then(Value::as_str)
        {
            names.push(name);
        }
        if node.kind == NodeKind::Hyperlink
            && let Some(anchor) = node.props.get("anchor").and_then(Value::as_str)
        {
            anchors.push(anchor);
        }
        if matches!(
            node.kind,
            NodeKind::PageReference | NodeKind::ReferenceField
        ) && let Some(bookmark) = node.props.get("bookmark").and_then(Value::as_str)
        {
            anchors.push(bookmark);
        }
        for child in &node.children {
            if let Child::Node(child) = child {
                collect(child, names, anchors);
            }
        }
    }

    let mut names = Vec::new();
    let mut anchors = Vec::new();
    collect(root, &mut names, &mut anchors);
    let unique = names.iter().copied().collect::<HashSet<_>>();
    if unique.len() != names.len() {
        return Err(validation("Document", "Bookmark names must be unique"));
    }
    if let Some(anchor) = anchors.into_iter().find(|anchor| !unique.contains(anchor)) {
        return Err(validation(
            "Document",
            format!("bookmark reference `{anchor}` has no matching Bookmark"),
        ));
    }
    Ok(())
}

fn validate_node(node: &Node, path: &str) -> Result<()> {
    validate_props(node, path)?;
    if node.kind == NodeKind::Document && node.children.is_empty() {
        return Err(validation(path, "Document requires at least one Section"));
    }
    if node.kind == NodeKind::Section {
        let mut saw_body = false;
        let mut toc_count = 0;
        let mut figures_count = 0;
        let mut entries_count = 0;
        for child in &node.children {
            let Child::Node(child) = child else { continue };
            if matches!(
                child.kind,
                NodeKind::TableOfContents | NodeKind::TableOfFigures | NodeKind::TableOfEntries
            ) {
                match child.kind {
                    NodeKind::TableOfContents => toc_count += 1,
                    NodeKind::TableOfFigures => figures_count += 1,
                    NodeKind::TableOfEntries => entries_count += 1,
                    _ => {}
                }
                if saw_body || toc_count > 1 || figures_count > 1 || entries_count > 1 {
                    return Err(validation(
                        path,
                        "Section permits one of each index component before body content",
                    ));
                }
            } else if !matches!(child.kind, NodeKind::Header | NodeKind::Footer) {
                saw_body = true;
            }
        }
    }
    if node.kind == NodeKind::Deleted {
        for child in &node.children {
            if let Child::Node(run) = child
                && run.kind == NodeKind::Run
                && run
                    .children
                    .iter()
                    .any(|child| matches!(child, Child::Node(node) if node.kind != NodeKind::Text))
            {
                return Err(validation(path, "Deleted Run only accepts text content"));
            }
        }
    }

    for (index, child) in node.children.iter().enumerate() {
        match child {
            Child::Node(child_node) => {
                if !allows(node.kind, child_node.kind) {
                    return Err(validation(
                        path,
                        format!(
                            "{} cannot contain {}",
                            node.kind.name(),
                            child_node.kind.name()
                        ),
                    ));
                }
                let child_path = format!("{path}/{}[{index}]", child_node.kind.name());
                validate_node(child_node, &child_path)?;
            }
            Child::String(_) | Child::Number(_) => {
                if !matches!(
                    node.kind,
                    NodeKind::Paragraph
                        | NodeKind::Run
                        | NodeKind::Text
                        | NodeKind::Hyperlink
                        | NodeKind::ListItem
                        | NodeKind::Heading
                        | NodeKind::Caption
                        | NodeKind::Comment
                        | NodeKind::Footnote
                        | NodeKind::Inserted
                        | NodeKind::Deleted
                        | NodeKind::MovedFrom
                        | NodeKind::MovedTo
                        | NodeKind::ContentControl
                        | NodeKind::Field
                        | NodeKind::DateField
                        | NodeKind::TimeField
                        | NodeKind::FileNameField
                        | NodeKind::AuthorField
                        | NodeKind::TitleField
                        | NodeKind::SubjectField
                        | NodeKind::SequenceField
                        | NodeKind::ReferenceField
                        | NodeKind::MergeField
                        | NodeKind::DocumentPropertyField
                        | NodeKind::FormulaField
                ) && !node.kind.is_semantic_text()
                {
                    return Err(validation(path, "text is not allowed here"));
                }
            }
        }
    }
    Ok(())
}

fn allows(parent: NodeKind, child: NodeKind) -> bool {
    match parent {
        NodeKind::Document => child == NodeKind::Section,
        NodeKind::Section => matches!(
            child,
            NodeKind::Paragraph
                | NodeKind::Table
                | NodeKind::Header
                | NodeKind::Footer
                | NodeKind::List
                | NodeKind::Heading
                | NodeKind::Caption
                | NodeKind::Index
                | NodeKind::Bookmark
                | NodeKind::TableOfContents
                | NodeKind::TableOfFigures
                | NodeKind::TableOfEntries
        ),
        NodeKind::Paragraph | NodeKind::Heading | NodeKind::Caption => {
            allows_paragraph_child(child)
        }
        NodeKind::Run => allows_run_child(child),
        NodeKind::Text
        | NodeKind::Break
        | NodeKind::CarriageReturn
        | NodeKind::NonBreakingSpace
        | NodeKind::SoftHyphen
        | NodeKind::NonBreakingHyphen
        | NodeKind::Image
        | NodeKind::PageNumber
        | NodeKind::TotalPages
        | NodeKind::TableOfContents
        | NodeKind::TableOfFigures
        | NodeKind::TableOfEntries
        | NodeKind::TocEntry
        | NodeKind::Index
        | NodeKind::IndexEntry
        | NodeKind::Tab
        | NodeKind::TabStop
        | NodeKind::Symbol
        | NodeKind::PageReference
        | NodeKind::PositionalTab => false,
        NodeKind::Bold
        | NodeKind::Italic
        | NodeKind::Underline
        | NodeKind::StrikeThrough
        | NodeKind::Superscript
        | NodeKind::Subscript
        | NodeKind::AllCaps
        | NodeKind::HiddenText
        | NodeKind::SpecialHiddenText
        | NodeKind::DoubleStrike
        | NodeKind::SpacedText
        | NodeKind::ScaledText
        | NodeKind::FitText
        | NodeKind::BorderedText
        | NodeKind::ShadedText => child == NodeKind::Text,
        NodeKind::Table => child == NodeKind::TableRow,
        NodeKind::TableRow => child == NodeKind::TableCell,
        NodeKind::TableCell | NodeKind::Header | NodeKind::Footer => matches!(
            child,
            NodeKind::Paragraph | NodeKind::Caption | NodeKind::Table | NodeKind::List
        ),
        NodeKind::Hyperlink
        | NodeKind::ListItem
        | NodeKind::Inserted
        | NodeKind::MovedFrom
        | NodeKind::MovedTo
        | NodeKind::ContentControl
        | NodeKind::Field
        | NodeKind::DateField
        | NodeKind::TimeField
        | NodeKind::FileNameField
        | NodeKind::AuthorField
        | NodeKind::TitleField
        | NodeKind::SubjectField
        | NodeKind::SequenceField
        | NodeKind::ReferenceField
        | NodeKind::MergeField
        | NodeKind::DocumentPropertyField
        | NodeKind::FormulaField => child == NodeKind::Run || allows_run_child(child),
        NodeKind::Comment => {
            child == NodeKind::Run || allows_run_child(child) || child == NodeKind::Hyperlink
        }
        NodeKind::List => child == NodeKind::ListItem,
        NodeKind::Bookmark => matches!(
            child,
            NodeKind::Paragraph
                | NodeKind::Heading
                | NodeKind::Caption
                | NodeKind::Table
                | NodeKind::List
        ),
        NodeKind::Footnote => {
            child == NodeKind::Run || (allows_run_child(child) && child != NodeKind::Footnote)
        }
        NodeKind::Deleted => matches!(child, NodeKind::Run | NodeKind::Text),
    }
}

fn allows_run_child(child: NodeKind) -> bool {
    child.is_special_character()
        || matches!(
            child,
            NodeKind::Text
                | NodeKind::Break
                | NodeKind::CarriageReturn
                | NodeKind::Image
                | NodeKind::Footnote
                | NodeKind::Tab
                | NodeKind::Symbol
                | NodeKind::PageReference
                | NodeKind::PositionalTab
                | NodeKind::TocEntry
        )
}

fn allows_paragraph_child(child: NodeKind) -> bool {
    child == NodeKind::Run
        || allows_run_child(child)
        || child.is_semantic_text()
        || matches!(
            child,
            NodeKind::Hyperlink
                | NodeKind::Comment
                | NodeKind::PageNumber
                | NodeKind::TotalPages
                | NodeKind::Inserted
                | NodeKind::Deleted
                | NodeKind::MovedFrom
                | NodeKind::MovedTo
                | NodeKind::ContentControl
                | NodeKind::Field
                | NodeKind::DateField
                | NodeKind::TimeField
                | NodeKind::FileNameField
                | NodeKind::AuthorField
                | NodeKind::TitleField
                | NodeKind::SubjectField
                | NodeKind::SequenceField
                | NodeKind::ReferenceField
                | NodeKind::MergeField
                | NodeKind::DocumentPropertyField
                | NodeKind::FormulaField
                | NodeKind::IndexEntry
                | NodeKind::TabStop
        )
}

fn validate_props(node: &Node, path: &str) -> Result<()> {
    let allowed = allowed_props(node.kind);
    for key in node.props.keys() {
        if !allowed.contains(&key.as_str()) {
            return Err(validation(path, format!("unknown property `{key}`")));
        }
    }
    validate_semantics(node, path)
}

fn allowed_props(kind: NodeKind) -> &'static [&'static str] {
    match kind {
        NodeKind::Document => &["defaultFont", "defaultSize"],
        NodeKind::Section => &["pageSize", "orientation", "margins"],
        NodeKind::Paragraph => paragraph_props(),
        NodeKind::Heading => heading_props(),
        NodeKind::Caption => caption_props(),
        NodeKind::Run => &[
            "style",
            "font",
            "size",
            "bold",
            "italic",
            "strike",
            "underline",
            "color",
            "highlight",
            "themeColor",
            "themeShade",
            "themeTint",
            "doubleStrike",
        ],
        NodeKind::Text => &["value"],
        NodeKind::Break | NodeKind::Header | NodeKind::Footer => &["type"],
        NodeKind::Image => &["src", "width", "height"],
        NodeKind::Table => table_props(),
        NodeKind::TableRow => &["height", "heightRule", "cantSplit"],
        NodeKind::TableCell => &["width", "colSpan", "verticalAlign", "shading", "border"],
        NodeKind::Hyperlink => &["href", "anchor", "history"],
        NodeKind::PageNumber
        | NodeKind::TotalPages
        | NodeKind::Footnote
        | NodeKind::Tab
        | NodeKind::CarriageReturn
        | NodeKind::NonBreakingSpace
        | NodeKind::SoftHyphen
        | NodeKind::NonBreakingHyphen => &[],
        kind @ (NodeKind::Bold
        | NodeKind::Italic
        | NodeKind::Underline
        | NodeKind::StrikeThrough
        | NodeKind::Superscript
        | NodeKind::Subscript
        | NodeKind::AllCaps
        | NodeKind::HiddenText
        | NodeKind::SpecialHiddenText
        | NodeKind::DoubleStrike
        | NodeKind::SpacedText
        | NodeKind::ScaledText
        | NodeKind::FitText
        | NodeKind::BorderedText
        | NodeKind::ShadedText) => semantic_text_props(kind),
        NodeKind::TabStop => &["position", "align", "leader"],
        NodeKind::List => &["type", "start"],
        NodeKind::ListItem => &["level"],
        NodeKind::Bookmark => &["name"],
        NodeKind::TableOfContents => &["startLevel", "endLevel", "hyperlinks", "dirty", "alias"],
        NodeKind::TableOfFigures => &[
            "label",
            "includeLabelAndNumber",
            "separator",
            "hyperlinks",
            "dirty",
            "alias",
        ],
        NodeKind::TableOfEntries => &["identifier", "hyperlinks", "dirty", "alias"],
        NodeKind::TocEntry => &["text", "level", "omitPageNumber", "identifier"],
        NodeKind::Index => index_props(),
        NodeKind::IndexEntry => index_entry_props(),
        NodeKind::Comment => &["text", "author", "date"],
        NodeKind::Symbol => &["font", "char"],
        NodeKind::Inserted | NodeKind::Deleted | NodeKind::MovedFrom | NodeKind::MovedTo => {
            &["author", "date"]
        }
        NodeKind::PageReference | NodeKind::ReferenceField => reference_props(),
        NodeKind::PositionalTab => &["align", "relativeTo", "leader"],
        NodeKind::ContentControl => &["alias", "xpath", "prefixMappings", "storeItemId"],
        NodeKind::Field => &["instruction", "dirty"],
        NodeKind::DateField | NodeKind::TimeField => &["format", "dirty"],
        NodeKind::FileNameField => &["fullPath", "dirty"],
        NodeKind::AuthorField | NodeKind::TitleField | NodeKind::SubjectField => &["dirty"],
        NodeKind::SequenceField => &["identifier", "format", "restart", "placeholder", "dirty"],
        NodeKind::MergeField => &["name", "preserveFormatting", "placeholder", "dirty"],
        NodeKind::DocumentPropertyField => &["name", "placeholder", "dirty"],
        NodeKind::FormulaField => &["expression", "numberFormat", "placeholder", "dirty"],
    }
}

fn caption_props() -> &'static [&'static str] {
    &[
        "label",
        "identifier",
        "format",
        "restart",
        "placeholder",
        "dirty",
        "style",
        "numberSeparator",
        "textSeparator",
    ]
}

fn index_props() -> &'static [&'static str] {
    &[
        "identifier",
        "columns",
        "runIn",
        "placeholder",
        "dirty",
        "style",
    ]
}

fn index_entry_props() -> &'static [&'static str] {
    &[
        "text",
        "subentry",
        "identifier",
        "boldPageNumber",
        "italicPageNumber",
        "pageRangeBookmark",
        "crossReference",
    ]
}

fn reference_props() -> &'static [&'static str] {
    &[
        "bookmark",
        "hyperlink",
        "relativePosition",
        "placeholder",
        "dirty",
    ]
}

fn paragraph_props() -> &'static [&'static str] {
    &[
        "style",
        "align",
        "spacingBefore",
        "spacingAfter",
        "lineSpacing",
        "indentLeft",
        "indentRight",
        "firstLine",
        "hanging",
        "keepNext",
        "keepLines",
        "pageBreakBefore",
        "snapToGrid",
        "widowControl",
        "font",
        "size",
        "bold",
        "italic",
        "color",
        "characterSpacing",
    ]
}

fn heading_props() -> &'static [&'static str] {
    &[
        "level",
        "style",
        "align",
        "spacingBefore",
        "spacingAfter",
        "lineSpacing",
        "indentLeft",
        "indentRight",
        "firstLine",
        "hanging",
        "keepNext",
        "keepLines",
        "pageBreakBefore",
        "snapToGrid",
        "widowControl",
        "font",
        "size",
        "bold",
        "italic",
        "color",
        "characterSpacing",
    ]
}

fn semantic_text_props(kind: NodeKind) -> &'static [&'static str] {
    match kind {
        NodeKind::Underline => &["type"],
        NodeKind::SpacedText => &["amount"],
        NodeKind::ScaledText => &["percent"],
        NodeKind::FitText => &["width", "id"],
        NodeKind::BorderedText => &["style", "size", "color", "space"],
        NodeKind::ShadedText => &["fill", "color", "pattern"],
        _ => &[],
    }
}

fn table_props() -> &'static [&'static str] {
    &[
        "width",
        "widthPercent",
        "align",
        "layout",
        "columnWidths",
        "border",
    ]
}

fn validate_semantics(node: &Node, path: &str) -> Result<()> {
    for key in ["defaultSize", "size", "width", "height"] {
        if let Some(value) = node.props.get(key) {
            require_number(value, path, key, true)?;
        }
    }
    for key in [
        "spacingBefore",
        "spacingAfter",
        "lineSpacing",
        "indentLeft",
        "indentRight",
        "firstLine",
        "hanging",
    ] {
        if let Some(value) = node.props.get(key) {
            require_number(value, path, key, false)?;
        }
    }
    for key in ["color", "shading"] {
        if let Some(value) = node.props.get(key) {
            let color = value
                .as_str()
                .ok_or_else(|| validation(path, format!("`{key}` must be a color string")))?;
            if color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(validation(path, format!("`{key}` must be six-digit RGB")));
            }
        }
    }
    if node.props.contains_key("firstLine") && node.props.contains_key("hanging") {
        return Err(validation(
            path,
            "`firstLine` and `hanging` are mutually exclusive",
        ));
    }
    validate_paragraph_defaults(node, path)?;
    if node.props.contains_key("width") && node.props.contains_key("widthPercent") {
        return Err(validation(
            path,
            "`width` and `widthPercent` are mutually exclusive",
        ));
    }
    if node.kind == NodeKind::Image {
        for key in ["src", "width", "height"] {
            if !node.props.contains_key(key) {
                return Err(validation(path, format!("Image requires `{key}`")));
            }
        }
    }
    if node.kind == NodeKind::Hyperlink {
        let targets = usize::from(node.props.contains_key("href"))
            + usize::from(node.props.contains_key("anchor"));
        if targets != 1 {
            return Err(validation(
                path,
                "Hyperlink requires exactly one of `href` or `anchor`",
            ));
        }
        let key = if node.props.contains_key("href") {
            "href"
        } else {
            "anchor"
        };
        if node
            .props
            .get(key)
            .and_then(Value::as_str)
            .is_none_or(str::is_empty)
        {
            return Err(validation(
                path,
                format!("Hyperlink `{key}` must be a non-empty string"),
            ));
        }
    }
    if node.kind == NodeKind::List && node.children.is_empty() {
        return Err(validation(path, "List requires at least one ListItem"));
    }
    validate_advanced_semantics(node, path)?;
    if let Some(value) = node.props.get("widthPercent") {
        let number = require_number(value, path, "widthPercent", true)?;
        if number > 100.0 {
            return Err(validation(path, "`widthPercent` must be at most 100"));
        }
    }
    Ok(())
}

fn validate_paragraph_defaults(node: &Node, path: &str) -> Result<()> {
    if matches!(node.kind, NodeKind::Paragraph | NodeKind::Heading) {
        for key in [
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "snapToGrid",
            "widowControl",
            "bold",
            "italic",
        ] {
            if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
                return Err(validation(path, format!("`{key}` must be a boolean")));
            }
        }
        if node
            .props
            .get("font")
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(path, "`font` must be a non-empty string"));
        }
        if let Some(value) = node.props.get("characterSpacing") {
            value
                .as_f64()
                .filter(|number| number.is_finite())
                .ok_or_else(|| validation(path, "`characterSpacing` must be a finite number"))?;
        }
    }
    Ok(())
}

fn validate_advanced_semantics(node: &Node, path: &str) -> Result<()> {
    validate_run_defaults(node, path)?;
    validate_field_semantics(node, path)?;
    validate_index_semantics(node, path)?;
    validate_text_effect_semantics(node, path)?;
    validate_revision_and_control_semantics(node, path)?;
    validate_annotation_semantics(node, path)?;
    validate_structure_semantics(node, path)
}

fn validate_run_defaults(node: &Node, path: &str) -> Result<()> {
    if node.kind != NodeKind::Run {
        return Ok(());
    }
    if node.props.contains_key("themeColor") {
        validate_optional_enum_prop(
            node,
            path,
            "themeColor",
            &[
                "dark1",
                "light1",
                "dark2",
                "light2",
                "accent1",
                "accent2",
                "accent3",
                "accent4",
                "accent5",
                "accent6",
                "hyperlink",
                "followedHyperlink",
                "none",
                "background1",
                "text1",
                "background2",
                "text2",
            ],
        )?;
    }
    for key in ["themeShade", "themeTint"] {
        if let Some(value) = node.props.get(key) {
            let modifier = value
                .as_str()
                .ok_or_else(|| validation(path, format!("`{key}` must be a hex byte")))?;
            if modifier.len() != 2 || !modifier.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(validation(path, format!("`{key}` must be a hex byte")));
            }
        }
    }
    for key in ["bold", "italic", "strike", "doubleStrike"] {
        if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(path, format!("`{key}` must be a boolean")));
        }
    }
    if node.props.get("strike").and_then(Value::as_bool) == Some(true)
        && node.props.get("doubleStrike").and_then(Value::as_bool) == Some(true)
    {
        return Err(validation(
            path,
            "enabled `strike` and `doubleStrike` are mutually exclusive",
        ));
    }
    Ok(())
}

fn validate_text_effect_semantics(node: &Node, path: &str) -> Result<()> {
    if node.kind.is_semantic_text() && !has_text_content(node) {
        return Err(validation(
            path,
            format!("{} requires text content", node.kind.name()),
        ));
    }
    if node.kind == NodeKind::Underline {
        validate_optional_enum_prop(
            node,
            path,
            "type",
            &["single", "double", "dotted", "dash", "wave"],
        )?;
    }
    if node.kind == NodeKind::SpacedText {
        let amount = node
            .props
            .get("amount")
            .and_then(Value::as_f64)
            .filter(|amount| amount.is_finite());
        if amount.is_none() {
            return Err(validation(
                path,
                "SpacedText requires finite numeric `amount`",
            ));
        }
    }
    if node.kind == NodeKind::ScaledText {
        let percent = node.props.get("percent").and_then(Value::as_u64);
        if percent.is_none_or(|percent| !(1..=600).contains(&percent)) {
            return Err(validation(
                path,
                "ScaledText `percent` must be an integer from 1 to 600",
            ));
        }
    }
    if node.kind == NodeKind::FitText {
        if node.props.get("width").is_none() {
            return Err(validation(path, "FitText requires positive `width`"));
        }
        if node.props.get("id").is_some_and(|value| {
            value
                .as_u64()
                .and_then(|value| u32::try_from(value).ok())
                .is_none()
        }) {
            return Err(validation(
                path,
                "FitText `id` must be an unsigned 32-bit integer",
            ));
        }
    }
    if node.kind == NodeKind::BorderedText {
        if let Some(style) = node.props.get("style").and_then(Value::as_str)
            && !["single", "double", "dotted", "dashed"].contains(&style)
        {
            return Err(validation(path, "BorderedText `style` is invalid"));
        }
        if node
            .props
            .get("style")
            .is_some_and(|value| !value.is_string())
        {
            return Err(validation(path, "BorderedText `style` must be a string"));
        }
        if let Some(space) = node.props.get("space")
            && space
                .as_u64()
                .and_then(|value| usize::try_from(value).ok())
                .is_none()
        {
            return Err(validation(
                path,
                "BorderedText `space` must be a non-negative integer",
            ));
        }
    }
    validate_shaded_text(node, path)?;
    Ok(())
}

fn validate_shaded_text(node: &Node, path: &str) -> Result<()> {
    if node.kind != NodeKind::ShadedText {
        return Ok(());
    }
    let fill = node.props.get("fill").and_then(Value::as_str);
    if fill
        .is_none_or(|value| value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()))
    {
        return Err(validation(path, "ShadedText requires six-digit RGB `fill`"));
    }
    if let Some(pattern) = node.props.get("pattern") {
        let pattern = pattern
            .as_str()
            .ok_or_else(|| validation(path, "ShadedText `pattern` must be a string"))?;
        if !SHADING_PATTERNS.contains(&pattern) {
            return Err(validation(path, "ShadedText `pattern` is invalid"));
        }
    }
    Ok(())
}

const SHADING_PATTERNS: &[&str] = &[
    "nil",
    "clear",
    "solid",
    "horzStripe",
    "vertStripe",
    "reverseDiagStripe",
    "diagStripe",
    "horzCross",
    "diagCross",
    "thinHorzStripe",
    "thinVertStripe",
    "thinReverseDiagStripe",
    "thinDiagStripe",
    "thinHorzCross",
    "thinDiagCross",
    "pct5",
    "pct10",
    "pct12",
    "pct15",
    "pct20",
    "pct25",
    "pct30",
    "pct35",
    "pct37",
    "pct40",
    "pct45",
    "pct50",
    "pct55",
    "pct60",
    "pct62",
    "pct65",
    "pct70",
    "pct75",
    "pct80",
    "pct85",
    "pct87",
    "pct90",
    "pct95",
];

fn validate_revision_and_control_semantics(node: &Node, path: &str) -> Result<()> {
    if matches!(
        node.kind,
        NodeKind::Inserted | NodeKind::Deleted | NodeKind::MovedFrom | NodeKind::MovedTo
    ) {
        if node.children.is_empty() {
            return Err(validation(path, "tracked revision requires content"));
        }
        for key in ["author", "date"] {
            if node.props.get(key).is_some_and(|value| !value.is_string()) {
                return Err(validation(
                    path,
                    format!("revision `{key}` must be a string"),
                ));
            }
        }
    }
    if node.kind == NodeKind::ContentControl {
        if node.children.is_empty() {
            return Err(validation(path, "ContentControl requires content"));
        }
        for key in ["alias", "xpath", "prefixMappings", "storeItemId"] {
            if node.props.get(key).is_some_and(|value| !value.is_string()) {
                return Err(validation(
                    path,
                    format!("ContentControl `{key}` must be a string"),
                ));
            }
        }
        let has_binding = ["xpath", "prefixMappings", "storeItemId"]
            .iter()
            .any(|key| node.props.contains_key(*key));
        if has_binding
            && node
                .props
                .get("xpath")
                .and_then(Value::as_str)
                .is_none_or(str::is_empty)
        {
            return Err(validation(
                path,
                "ContentControl data binding requires non-empty `xpath`",
            ));
        }
    }
    if node.kind == NodeKind::Field
        && node
            .props
            .get("instruction")
            .and_then(Value::as_str)
            .is_none_or(str::is_empty)
    {
        return Err(validation(path, "Field requires non-empty `instruction`"));
    }
    validate_typed_field(node, path)?;
    Ok(())
}

fn validate_typed_field(node: &Node, path: &str) -> Result<()> {
    if !node.kind.is_field() {
        return Ok(());
    }
    if node
        .props
        .get("dirty")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(validation(path, "field `dirty` must be a boolean"));
    }
    if matches!(node.kind, NodeKind::DateField | NodeKind::TimeField)
        && let Some(format) = node.props.get("format")
    {
        let valid = format.as_str().is_some_and(|value| {
            !value.is_empty() && !value.contains('"') && !value.chars().any(char::is_control)
        });
        if !valid {
            return Err(validation(
                path,
                "date/time field `format` must be non-empty and contain no quotes or controls",
            ));
        }
    }
    if node.kind == NodeKind::FileNameField
        && node
            .props
            .get("fullPath")
            .is_some_and(|value| !value.is_boolean())
    {
        return Err(validation(
            path,
            "FileNameField `fullPath` must be a boolean",
        ));
    }
    if node.kind == NodeKind::SequenceField {
        let identifier = node.props.get("identifier").and_then(Value::as_str);
        if identifier.is_none_or(|value| !is_word_identifier(value)) {
            return Err(validation(
                path,
                "SequenceField requires a valid `identifier`",
            ));
        }
        validate_optional_enum_prop(
            node,
            path,
            "format",
            &["arabic", "roman", "Roman", "alphabetic", "Alphabetic"],
        )?;
        if node
            .props
            .get("restart")
            .is_some_and(|value| value.as_u64().is_none())
        {
            return Err(validation(
                path,
                "SequenceField `restart` must be a non-negative integer",
            ));
        }
    }
    if node.kind == NodeKind::ReferenceField {
        let bookmark = node.props.get("bookmark").and_then(Value::as_str);
        if bookmark.is_none_or(str::is_empty) {
            return Err(validation(
                path,
                "ReferenceField requires non-empty `bookmark`",
            ));
        }
        for key in ["hyperlink", "relativePosition"] {
            if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
                return Err(validation(
                    path,
                    format!("ReferenceField `{key}` must be a boolean"),
                ));
            }
        }
    }
    if matches!(
        node.kind,
        NodeKind::SequenceField
            | NodeKind::ReferenceField
            | NodeKind::MergeField
            | NodeKind::DocumentPropertyField
            | NodeKind::FormulaField
    ) && node
        .props
        .get("placeholder")
        .is_some_and(|value| !value.is_string())
    {
        return Err(validation(path, "field `placeholder` must be a string"));
    }
    validate_merge_and_formula_fields(node, path)
}

fn validate_merge_and_formula_fields(node: &Node, path: &str) -> Result<()> {
    if node.kind == NodeKind::MergeField {
        let name = node.props.get("name").and_then(Value::as_str);
        if name.is_none_or(|value| !is_word_identifier(value)) {
            return Err(validation(path, "MergeField requires a valid `name`"));
        }
        if node
            .props
            .get("preserveFormatting")
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(validation(
                path,
                "MergeField `preserveFormatting` must be a boolean",
            ));
        }
    }
    if node.kind == NodeKind::DocumentPropertyField {
        validate_quoted_field_value(node, path, "name", "DocumentPropertyField")?;
    }
    if node.kind == NodeKind::FormulaField {
        let expression = node.props.get("expression").and_then(Value::as_str);
        if expression.is_none_or(|value| value.is_empty() || value.chars().any(char::is_control)) {
            return Err(validation(
                path,
                "FormulaField requires non-empty control-free `expression`",
            ));
        }
        if node.props.contains_key("numberFormat") {
            validate_quoted_field_value(node, path, "numberFormat", "FormulaField")?;
        }
    }
    Ok(())
}

fn validate_quoted_field_value(node: &Node, path: &str, key: &str, component: &str) -> Result<()> {
    let valid = node
        .props
        .get(key)
        .and_then(Value::as_str)
        .is_some_and(|value| {
            !value.is_empty() && !value.contains('"') && !value.chars().any(char::is_control)
        });
    if !valid {
        return Err(validation(
            path,
            format!("{component} `{key}` must be non-empty and contain no quotes or controls"),
        ));
    }
    Ok(())
}

fn is_word_identifier(value: &str) -> bool {
    let mut chars = value.chars();
    chars
        .next()
        .is_some_and(|first| first.is_ascii_alphabetic() || first == '_')
        && chars.all(|character| character.is_ascii_alphanumeric() || character == '_')
}

fn has_text_content(node: &Node) -> bool {
    node.children.iter().any(|child| match child {
        Child::String(value) => !value.is_empty(),
        Child::Node(text) if text.kind == NodeKind::Text => {
            text.props.get("value").is_some_and(|value| match value {
                Value::String(value) => !value.is_empty(),
                Value::Number(_) => true,
                _ => false,
            }) || text.children.iter().any(|child| match child {
                Child::String(value) => !value.is_empty(),
                Child::Number(_) => true,
                Child::Node(_) => false,
            })
        }
        Child::Number(_) | Child::Node(_) => true,
    })
}

fn validate_index_semantics(node: &Node, path: &str) -> Result<()> {
    if node.kind == NodeKind::Index {
        validate_optional_field_argument(node, path, "identifier")?;
        if let Some(columns) = node.props.get("columns")
            && columns
                .as_u64()
                .is_none_or(|columns| !(1..=4).contains(&columns))
        {
            return Err(validation(
                path,
                "Index `columns` must be an integer from 1 to 4",
            ));
        }
        for key in ["runIn", "dirty"] {
            if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
                return Err(validation(path, format!("Index `{key}` must be boolean")));
            }
        }
        for key in ["placeholder", "style"] {
            if node.props.get(key).is_some_and(|value| !value.is_string()) {
                return Err(validation(path, format!("Index `{key}` must be a string")));
            }
        }
        if node.props.get("style").and_then(Value::as_str) == Some("") {
            return Err(validation(path, "Index `style` must be non-empty"));
        }
    }
    if node.kind == NodeKind::TableOfFigures {
        let label = node.props.get("label").and_then(Value::as_str);
        if label.is_none_or(str::is_empty) {
            return Err(validation(
                path,
                "TableOfFigures requires non-empty `label`",
            ));
        }
        if node
            .props
            .get("includeLabelAndNumber")
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(validation(
                path,
                "TableOfFigures `includeLabelAndNumber` must be boolean",
            ));
        }
    }
    if node.kind == NodeKind::TableOfEntries {
        let identifier = node.props.get("identifier").and_then(Value::as_str);
        if identifier.is_none_or(str::is_empty) {
            return Err(validation(
                path,
                "TableOfEntries requires non-empty `identifier`",
            ));
        }
    }
    if matches!(
        node.kind,
        NodeKind::TableOfFigures | NodeKind::TableOfEntries
    ) {
        for key in ["hyperlinks", "dirty"] {
            if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
                return Err(validation(path, format!("`{key}` must be boolean")));
            }
        }
        for key in ["alias", "separator"] {
            if node.props.get(key).is_some_and(|value| !value.is_string()) {
                return Err(validation(path, format!("`{key}` must be a string")));
            }
        }
    }
    Ok(())
}

fn validate_field_semantics(node: &Node, path: &str) -> Result<()> {
    if node.kind == NodeKind::IndexEntry {
        validate_index_entry_semantics(node, path)?;
    }
    if node.kind == NodeKind::TocEntry {
        let text = node.props.get("text").and_then(Value::as_str);
        if text.is_none_or(str::is_empty) {
            return Err(validation(path, "TocEntry requires non-empty `text`"));
        }
        if let Some(level) = node.props.get("level")
            && level.as_u64().is_none_or(|level| !(1..=9).contains(&level))
        {
            return Err(validation(
                path,
                "TocEntry `level` must be an integer from 1 to 9",
            ));
        }
        if node
            .props
            .get("omitPageNumber")
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(validation(
                path,
                "TocEntry `omitPageNumber` must be boolean",
            ));
        }
        if node
            .props
            .get("identifier")
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(
                path,
                "TocEntry `identifier` must be a non-empty string",
            ));
        }
    }
    if node.kind == NodeKind::TabStop {
        let position = node
            .props
            .get("position")
            .ok_or_else(|| validation(path, "TabStop requires non-negative numeric `position`"))?;
        require_number(position, path, "position", false)?;
        validate_optional_enum_prop(
            node,
            path,
            "align",
            &["left", "center", "right", "decimal", "bar", "clear"],
        )?;
        validate_optional_enum_prop(
            node,
            path,
            "leader",
            &["none", "dot", "heavy", "hyphen", "middleDot", "underscore"],
        )?;
    }
    if node.kind == NodeKind::PageReference {
        let bookmark = node.props.get("bookmark").and_then(Value::as_str);
        if bookmark.is_none_or(str::is_empty) {
            return Err(validation(
                path,
                "PageReference requires non-empty `bookmark`",
            ));
        }
        if node
            .props
            .get("placeholder")
            .is_some_and(|value| !value.is_string())
        {
            return Err(validation(
                path,
                "PageReference `placeholder` must be a string",
            ));
        }
        for key in ["hyperlink", "relativePosition", "dirty"] {
            if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
                return Err(validation(
                    path,
                    format!("PageReference `{key}` must be boolean"),
                ));
            }
        }
    }
    if node.kind == NodeKind::PositionalTab {
        validate_optional_enum_prop(node, path, "align", &["left", "center", "right"])?;
        validate_optional_enum_prop(node, path, "relativeTo", &["margin", "indent"])?;
        validate_optional_enum_prop(
            node,
            path,
            "leader",
            &["none", "dot", "heavy", "hyphen", "middleDot", "underscore"],
        )?;
    }
    Ok(())
}

fn validate_index_entry_semantics(node: &Node, path: &str) -> Result<()> {
    let text = node.props.get("text").and_then(Value::as_str);
    if text.is_none_or(|value| value.is_empty() || !is_valid_field_argument(value)) {
        return Err(validation(
            path,
            "IndexEntry requires valid non-empty `text`",
        ));
    }
    for key in ["subentry", "identifier", "crossReference"] {
        validate_optional_field_argument(node, path, key)?;
    }
    if node.props.get("pageRangeBookmark").is_some_and(|value| {
        value
            .as_str()
            .is_none_or(|value| !is_word_identifier(value))
    }) {
        return Err(validation(
            path,
            "IndexEntry `pageRangeBookmark` must be a valid Word identifier",
        ));
    }
    for key in ["boldPageNumber", "italicPageNumber"] {
        if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(
                path,
                format!("IndexEntry `{key}` must be boolean"),
            ));
        }
    }
    if node.props.contains_key("pageRangeBookmark") && node.props.contains_key("crossReference") {
        return Err(validation(
            path,
            "IndexEntry `pageRangeBookmark` and `crossReference` are mutually exclusive",
        ));
    }
    Ok(())
}

fn validate_optional_field_argument(node: &Node, path: &str, key: &str) -> Result<()> {
    if node.props.get(key).is_some_and(|value| {
        value
            .as_str()
            .is_none_or(|value| value.is_empty() || !is_valid_field_argument(value))
    }) {
        return Err(validation(
            path,
            format!("`{key}` must be a valid non-empty field argument"),
        ));
    }
    Ok(())
}

fn is_valid_field_argument(value: &str) -> bool {
    !value.contains('"') && !value.contains('\\') && !value.chars().any(char::is_control)
}

fn validate_annotation_semantics(node: &Node, path: &str) -> Result<()> {
    if node.kind == NodeKind::Comment {
        let text = node.props.get("text").and_then(Value::as_str);
        if text.is_none_or(str::is_empty) || node.children.is_empty() {
            return Err(validation(
                path,
                "Comment requires non-empty `text` and selected children",
            ));
        }
        for key in ["author", "date"] {
            if node.props.get(key).is_some_and(|value| !value.is_string()) {
                return Err(validation(
                    path,
                    format!("Comment `{key}` must be a string"),
                ));
            }
        }
    }
    if node.kind == NodeKind::Footnote && node.children.is_empty() {
        return Err(validation(path, "Footnote requires content"));
    }
    if node.kind == NodeKind::Symbol {
        for key in ["font", "char"] {
            let value = node.props.get(key).and_then(Value::as_str);
            if value.is_none_or(str::is_empty) {
                return Err(validation(
                    path,
                    format!("Symbol requires non-empty `{key}`"),
                ));
            }
        }
        let char_code = node
            .props
            .get("char")
            .and_then(Value::as_str)
            .unwrap_or_default();
        if char_code.len() != 4 || !char_code.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(validation(
                path,
                "Symbol `char` must be four hexadecimal digits",
            ));
        }
    }
    Ok(())
}

fn validate_structure_semantics(node: &Node, path: &str) -> Result<()> {
    if node.kind == NodeKind::Heading {
        let level = node.props.get("level").and_then(Value::as_u64);
        if level.is_none_or(|level| !(1..=9).contains(&level)) {
            return Err(validation(
                path,
                "Heading requires integer `level` from 1 to 9",
            ));
        }
    }
    if node.kind == NodeKind::Bookmark {
        let name = node.props.get("name").and_then(Value::as_str);
        if name.is_none_or(str::is_empty) {
            return Err(validation(path, "Bookmark requires non-empty `name`"));
        }
    }
    if node.kind == NodeKind::Caption {
        let label = node.props.get("label").and_then(Value::as_str);
        if label.is_none_or(|value| !is_word_identifier(value)) {
            return Err(validation(path, "Caption requires a valid `label`"));
        }
        if node.props.get("identifier").is_some_and(|value| {
            value
                .as_str()
                .is_none_or(|value| !is_word_identifier(value))
        }) {
            return Err(validation(path, "Caption `identifier` must be valid"));
        }
        validate_optional_enum_prop(
            node,
            path,
            "format",
            &["arabic", "roman", "Roman", "alphabetic", "Alphabetic"],
        )?;
        if node
            .props
            .get("restart")
            .is_some_and(|value| value.as_u64().is_none())
        {
            return Err(validation(
                path,
                "Caption `restart` must be a non-negative integer",
            ));
        }
        for key in ["placeholder", "style", "numberSeparator", "textSeparator"] {
            if node.props.get(key).is_some_and(|value| !value.is_string()) {
                return Err(validation(
                    path,
                    format!("Caption `{key}` must be a string"),
                ));
            }
        }
        if node.props.get("style").and_then(Value::as_str) == Some("") {
            return Err(validation(path, "Caption `style` must be non-empty"));
        }
        if node
            .props
            .get("dirty")
            .is_some_and(|value| !value.is_boolean())
        {
            return Err(validation(path, "Caption `dirty` must be a boolean"));
        }
    }
    if node.kind == NodeKind::TableOfContents {
        let start = node
            .props
            .get("startLevel")
            .and_then(Value::as_u64)
            .unwrap_or(1);
        let end = node
            .props
            .get("endLevel")
            .and_then(Value::as_u64)
            .unwrap_or(3);
        if !(1..=9).contains(&start) || !(start..=9).contains(&end) {
            return Err(validation(
                path,
                "TOC levels must satisfy 1 <= startLevel <= endLevel <= 9",
            ));
        }
    }
    if node.kind == NodeKind::ListItem
        && let Some(value) = node.props.get("level")
        && value.as_u64().is_none_or(|level| level > 8)
    {
        return Err(validation(path, "`level` must be an integer from 0 to 8"));
    }
    if let Some(value) = node.props.get("start")
        && value.as_u64().is_none_or(|start| start == 0)
    {
        return Err(validation(path, "`start` must be a positive integer"));
    }
    Ok(())
}

fn validate_optional_enum_prop(node: &Node, path: &str, key: &str, allowed: &[&str]) -> Result<()> {
    let Some(value) = node.props.get(key) else {
        return Ok(());
    };
    let value = value
        .as_str()
        .ok_or_else(|| validation(path, format!("`{key}` must be a string")))?;
    if !allowed.contains(&value) {
        return Err(validation(path, format!("invalid `{key}` value `{value}`")));
    }
    Ok(())
}

fn require_number(value: &Value, path: &str, key: &str, positive: bool) -> Result<f64> {
    let number = value
        .as_f64()
        .filter(|number| number.is_finite())
        .ok_or_else(|| validation(path, format!("`{key}` must be a finite number")))?;
    if (positive && number <= 0.0) || (!positive && number < 0.0) {
        let requirement = if positive { "positive" } else { "non-negative" };
        return Err(validation(path, format!("`{key}` must be {requirement}")));
    }
    Ok(number)
}

fn validation(path: impl Into<String>, message: impl Into<String>) -> Error {
    Error::Validation {
        path: path.into(),
        message: message.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(source: &str) -> IrEnvelope {
        serde_json::from_str(source).expect("test IR should parse")
    }

    #[test]
    fn validate_should_accept_minimal_document() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["hello"]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_nesting() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Paragraph","props":{},"children":[]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("paragraph under document must fail");
        assert!(
            error
                .to_string()
                .contains("Document cannot contain Paragraph")
        );
    }

    #[test]
    fn validate_should_reject_unknown_property() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"wat":true},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        let error = ir.validate().expect_err("unknown property must fail");
        assert!(error.to_string().contains("unknown property `wat`"));
    }

    #[test]
    fn validate_should_accept_advanced_components() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Header","props":{},"children":[{"type":"Paragraph","props":{},"children":["head"]}]},{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"href":"https://example.com"},"children":["link"]},{"type":"PageNumber","props":{},"children":[]}]},{"type":"List","props":{"type":"ordered","start":2},"children":[{"type":"ListItem","props":{"level":1},"children":["item"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok());
    }

    #[test]
    fn validate_should_reject_ambiguous_hyperlink_and_invalid_list_level() {
        let hyperlink = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"href":"x","anchor":"y"},"children":[]}]}]}]}}"#,
        );
        assert!(
            hyperlink
                .validate()
                .expect_err("ambiguous target must fail")
                .to_string()
                .contains("exactly one")
        );

        let list = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{},"children":[{"type":"ListItem","props":{"level":9},"children":["item"]}]}]}]}}"#,
        );
        assert!(
            list.validate()
                .expect_err("level 9 must fail")
                .to_string()
                .contains("0 to 8")
        );
    }

    #[test]
    fn validate_should_accept_heading_bookmark_and_toc() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{"startLevel":1,"endLevel":3},"children":[]},{"type":"Bookmark","props":{"name":"intro"},"children":[{"type":"Heading","props":{"level":1},"children":["Intro"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_paragraph_and_heading_run_defaults() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"snapToGrid":false,"widowControl":true,"font":"Noto Sans CJK SC","size":12,"bold":true,"italic":false,"color":"1a2B3c","characterSpacing":0.5},"children":["body"]},{"type":"Heading","props":{"level":2,"font":"Noto Sans CJK SC","size":16},"children":["title"]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_run_theme_color_and_explicit_formatting_flags() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Run","props":{"color":"2E74B5","themeColor":"accent1","themeShade":"BF","themeTint":"99","bold":false,"italic":false,"strike":false,"doubleStrike":true},"children":["themed"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_run_theme_and_conflicting_strikes() {
        let invalid_theme = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Run","props":{"themeColor":"accent9"},"children":["x"]}]}]}]}}"#,
        );
        assert!(
            invalid_theme
                .validate()
                .expect_err("unknown theme must fail")
                .to_string()
                .contains("themeColor")
        );

        let conflicting = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Run","props":{"strike":true,"doubleStrike":true},"children":["x"]}]}]}]}}"#,
        );
        assert!(
            conflicting
                .validate()
                .expect_err("strike modes must be exclusive")
                .to_string()
                .contains("mutually exclusive")
        );
    }

    #[test]
    fn validate_should_reject_invalid_paragraph_run_defaults() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"snapToGrid":"yes","font":""},"children":[]}]}]}}"#,
        );
        let error = ir.validate().expect_err("invalid defaults must fail");
        assert!(error.to_string().contains("snapToGrid"));

        let empty_font = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Heading","props":{"level":1,"font":""},"children":[]}]}]}}"#,
        );
        assert!(
            empty_font
                .validate()
                .expect_err("empty font must fail")
                .to_string()
                .contains("non-empty string")
        );
    }

    #[test]
    fn validate_should_reject_invalid_heading_and_toc_range() {
        let heading = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Heading","props":{"level":0},"children":[]}]}]}}"#,
        );
        assert!(
            heading
                .validate()
                .expect_err("level zero must fail")
                .to_string()
                .contains("1 to 9")
        );

        let toc = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{"startLevel":4,"endLevel":2},"children":[]}]}]}}"#,
        );
        assert!(
            toc.validate()
                .expect_err("reversed range must fail")
                .to_string()
                .contains("startLevel")
        );
    }

    #[test]
    fn validate_should_reject_unresolved_or_duplicate_bookmarks() {
        let unresolved = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"anchor":"missing"},"children":["jump"]}]}]}]}}"#,
        );
        assert!(
            unresolved
                .validate()
                .expect_err("unresolved anchor must fail")
                .to_string()
                .contains("no matching Bookmark")
        );

        let duplicate = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Bookmark","props":{"name":"same"},"children":[]},{"type":"Bookmark","props":{"name":"same"},"children":[]}]}]}}"#,
        );
        assert!(
            duplicate
                .validate()
                .expect_err("duplicate bookmark must fail")
                .to_string()
                .contains("must be unique")
        );
    }

    #[test]
    fn validate_should_accept_review_and_inline_reference_components() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Comment","props":{"text":"review","author":"Ada"},"children":["selected"]},{"type":"Tab","props":{},"children":[]},{"type":"Symbol","props":{"font":"Wingdings","char":"F0A7"},"children":[]},{"type":"Footnote","props":{},"children":["note"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_complex_field() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Field","props":{"instruction":" DATE \\@ \"yyyy-MM-dd\" ","dirty":false},"children":[{"type":"Run","props":{"bold":true},"children":["2026-08-14"]}]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_field_properties() {
        let missing = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Field","props":{},"children":[]}]}]}]}}"#,
        );
        assert!(
            missing
                .validate()
                .expect_err("missing instruction must fail")
                .to_string()
                .contains("non-empty `instruction`")
        );

        let dirty = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Field","props":{"instruction":"DATE","dirty":"yes"},"children":[]}]}]}]}}"#,
        );
        assert!(
            dirty
                .validate()
                .expect_err("non-boolean dirty must fail")
                .to_string()
                .contains("must be a boolean")
        );
    }

    #[test]
    fn validate_should_accept_typed_field_components() {
        let ir = parse(
            r##"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"DateField","props":{"format":"yyyy-MM-dd"},"children":["date"]},{"type":"TimeField","props":{"format":"HH:mm","dirty":false},"children":[]},{"type":"FileNameField","props":{"fullPath":true},"children":["file"]},{"type":"AuthorField","props":{},"children":[]},{"type":"TitleField","props":{},"children":[]},{"type":"SubjectField","props":{},"children":[]},{"type":"MergeField","props":{"name":"CustomerName","preserveFormatting":true,"placeholder":"Ada"},"children":[]},{"type":"DocumentPropertyField","props":{"name":"Project Name"},"children":["Apollo"]},{"type":"FormulaField","props":{"expression":"SUM(ABOVE)","numberFormat":"#,##0.00"},"children":["42.00"]}]}]}]}}"##,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_sequence_and_reference_fields() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Bookmark","props":{"name":"target"},"children":[]},{"type":"Paragraph","props":{},"children":[{"type":"SequenceField","props":{"identifier":"Figure","format":"Roman","restart":3,"placeholder":"III"},"children":[]},{"type":"ReferenceField","props":{"bookmark":"target","hyperlink":true,"relativePosition":true,"placeholder":"target text"},"children":[]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_caption_with_inline_content() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Caption","props":{"label":"Figure","format":"Roman","restart":3,"placeholder":"III","dirty":false,"style":"FigureCaption","numberSeparator":" ","textSeparator":" — "},"children":[{"type":"Run","props":{"bold":true},"children":["Architecture"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_index_and_index_entry() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"IndexEntry","props":{"text":"Rust","subentry":"Ownership","identifier":"topics","boldPageNumber":true},"children":[]}]},{"type":"Index","props":{"identifier":"topics","columns":2,"runIn":true,"placeholder":"Update index","dirty":false,"style":"IndexBody"},"children":[]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_ambiguous_index_entry() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"IndexEntry","props":{"text":"Rust","pageRangeBookmark":"range","crossReference":"see Ownership"},"children":[]}]}]}]}}"#,
        );
        assert!(
            ir.validate()
                .expect_err("ambiguous index entry must fail")
                .to_string()
                .contains("mutually exclusive")
        );
    }

    #[test]
    fn validate_should_reject_caption_with_invalid_identifier() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Caption","props":{"label":"Figure 1"},"children":["Architecture"]}]}]}}"#,
        );
        assert!(
            ir.validate()
                .expect_err("invalid caption identifier must fail")
                .to_string()
                .contains("valid `label`")
        );
    }

    #[test]
    fn validate_should_reject_invalid_sequence_and_reference_fields() {
        let sequence = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"SequenceField","props":{"identifier":"bad name","restart":-1},"children":[]}]}]}]}}"#,
        );
        assert!(
            sequence
                .validate()
                .expect_err("invalid identifier must fail")
                .to_string()
                .contains("valid `identifier`")
        );

        let reference = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"ReferenceField","props":{"bookmark":"missing"},"children":[]}]}]}]}}"#,
        );
        assert!(
            reference
                .validate()
                .expect_err("unresolved reference must fail")
                .to_string()
                .contains("no matching Bookmark")
        );
    }

    #[test]
    fn validate_should_reject_invalid_typed_field_properties() {
        let format = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"DateField","props":{"format":"yyyy \"year\""},"children":[]}]}]}]}}"#,
        );
        assert!(
            format
                .validate()
                .expect_err("quoted format must fail")
                .to_string()
                .contains("contain no quotes")
        );

        let full_path = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"FileNameField","props":{"fullPath":"yes"},"children":[]}]}]}]}}"#,
        );
        assert!(
            full_path
                .validate()
                .expect_err("non-boolean fullPath must fail")
                .to_string()
                .contains("fullPath")
        );

        let formula = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"FormulaField","props":{"expression":"","numberFormat":"0 \"items\""},"children":[]}]}]}]}}"#,
        );
        assert!(
            formula
                .validate()
                .expect_err("empty formula must fail")
                .to_string()
                .contains("expression")
        );
    }

    #[test]
    fn validate_should_reject_empty_comment_footnote_and_invalid_symbol() {
        let comment = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Comment","props":{"text":""},"children":[]}]}]}]}}"#,
        );
        assert!(
            comment
                .validate()
                .expect_err("empty comment must fail")
                .to_string()
                .contains("Comment requires")
        );

        let footnote = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Footnote","props":{},"children":[]}]}]}]}}"#,
        );
        assert!(
            footnote
                .validate()
                .expect_err("empty footnote must fail")
                .to_string()
                .contains("Footnote requires")
        );

        let symbol = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Symbol","props":{"font":"Wingdings","char":"XYZ"},"children":[]}]}]}]}}"#,
        );
        assert!(
            symbol
                .validate()
                .expect_err("invalid symbol must fail")
                .to_string()
                .contains("hexadecimal")
        );
    }

    #[test]
    fn validate_should_accept_tracked_revisions() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Deleted","props":{"author":"Ada"},"children":[{"type":"Run","props":{"bold":true},"children":["old"]}]},{"type":"Inserted","props":{"author":"Ada"},"children":["new"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_empty_revision_and_non_text_deleted_run() {
        let empty = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Inserted","props":{},"children":[]}]}]}]}}"#,
        );
        assert!(
            empty
                .validate()
                .expect_err("empty revision must fail")
                .to_string()
                .contains("requires content")
        );

        let invalid = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Deleted","props":{},"children":[{"type":"Run","props":{},"children":[{"type":"Break","props":{},"children":[]}]}]}]}]}]}}"#,
        );
        assert!(
            invalid
                .validate()
                .expect_err("deleted break must fail")
                .to_string()
                .contains("only accepts text")
        );
    }

    #[test]
    fn validate_should_accept_page_reference_and_positional_tab() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Bookmark","props":{"name":"target"},"children":[]},{"type":"Paragraph","props":{},"children":[{"type":"PageReference","props":{"bookmark":"target","relativePosition":true},"children":[]},{"type":"PositionalTab","props":{"align":"right","relativeTo":"margin","leader":"dot"},"children":[]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_unresolved_page_reference_and_invalid_positional_tab() {
        let reference = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"PageReference","props":{"bookmark":"missing"},"children":[]}]}]}]}}"#,
        );
        assert!(
            reference
                .validate()
                .expect_err("unresolved page reference must fail")
                .to_string()
                .contains("no matching Bookmark")
        );

        let tab = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"PositionalTab","props":{"leader":"stars"},"children":[]}]}]}]}}"#,
        );
        assert!(
            tab.validate()
                .expect_err("invalid tab leader must fail")
                .to_string()
                .contains("invalid `leader`")
        );
    }

    #[test]
    fn validate_should_accept_bound_content_control() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"ContentControl","props":{"alias":"Customer","xpath":"/root/customer","prefixMappings":"xmlns:x='urn:test'","storeItemId":"{ABC}"},"children":["Ada"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_empty_or_incomplete_content_control() {
        let empty = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"ContentControl","props":{},"children":[]}]}]}]}}"#,
        );
        assert!(
            empty
                .validate()
                .expect_err("empty control must fail")
                .to_string()
                .contains("requires content")
        );

        let incomplete = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"ContentControl","props":{"storeItemId":"{ABC}"},"children":["Ada"]}]}]}]}}"#,
        );
        assert!(
            incomplete
                .validate()
                .expect_err("binding without xpath must fail")
                .to_string()
                .contains("requires non-empty `xpath`")
        );
    }

    #[test]
    fn validate_should_accept_carriage_return_and_tab_stop() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"TabStop","props":{"position":72,"align":"right","leader":"dot"},"children":[]},"Label",{"type":"Tab","props":{},"children":[]},"Value",{"type":"CarriageReturn","props":{},"children":[]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_non_breaking_and_soft_characters() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["A",{"type":"NonBreakingSpace","props":{},"children":[]},{"type":"Run","props":{},"children":["soft",{"type":"SoftHyphen","props":{},"children":[]},"hyphen",{"type":"NonBreakingHyphen","props":{},"children":[]}]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_character_component_children() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"SoftHyphen","props":{},"children":["invalid"]}]}]}]}}"#,
        );
        assert!(
            ir.validate()
                .expect_err("character component children must fail")
                .to_string()
                .contains("text is not allowed")
        );
    }

    #[test]
    fn validate_should_reject_invalid_tab_stop() {
        let missing = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"TabStop","props":{},"children":[]}]}]}]}}"#,
        );
        assert!(
            missing
                .validate()
                .expect_err("missing position must fail")
                .to_string()
                .contains("requires non-negative numeric `position`")
        );

        let invalid = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"TabStop","props":{"position":-1,"align":"justify"},"children":[]}]}]}]}}"#,
        );
        assert!(
            invalid
                .validate()
                .expect_err("negative position must fail")
                .to_string()
                .contains("must be non-negative")
        );
    }

    #[test]
    fn validate_should_accept_toc_entry_and_move_revisions() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"TocEntry","props":{"text":"Appendix","level":2,"omitPageNumber":true,"identifier":"figures"},"children":[]},{"type":"MovedFrom","props":{"author":"Ada","date":"2026-08-14T00:00:00Z"},"children":["old"]},{"type":"MovedTo","props":{"author":"Ada","date":"2026-08-14T00:00:00Z"},"children":[{"type":"Run","props":{"bold":true},"children":["new"]}]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_toc_entry_and_empty_move_revision() {
        let toc_entry = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"TocEntry","props":{"text":"","level":10},"children":[]}]}]}]}}"#,
        );
        assert!(
            toc_entry
                .validate()
                .expect_err("empty entry text must fail")
                .to_string()
                .contains("requires non-empty `text`")
        );

        let moved = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"MovedTo","props":{},"children":[]}]}]}]}}"#,
        );
        assert!(
            moved
                .validate()
                .expect_err("empty move revision must fail")
                .to_string()
                .contains("requires content")
        );
    }

    #[test]
    fn validate_should_accept_figure_and_custom_entry_indexes() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfFigures","props":{"label":"Figure","includeLabelAndNumber":true,"separator":" — "},"children":[]},{"type":"TableOfEntries","props":{"identifier":"legal","hyperlinks":true},"children":[]},{"type":"Paragraph","props":{},"children":[{"type":"TocEntry","props":{"text":"Terms","identifier":"legal"},"children":[]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_or_misplaced_index_components() {
        let missing = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfFigures","props":{},"children":[]}]}]}}"#,
        );
        assert!(
            missing
                .validate()
                .expect_err("missing figure label must fail")
                .to_string()
                .contains("requires non-empty `label`")
        );

        let misplaced = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["body"]},{"type":"TableOfEntries","props":{"identifier":"legal"},"children":[]}]}]}}"#,
        );
        assert!(
            misplaced
                .validate()
                .expect_err("index after body must fail")
                .to_string()
                .contains("before body content")
        );
    }

    #[test]
    fn validate_should_accept_semantic_text_wrappers() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":["H",{"type":"Subscript","props":{},"children":[2]},"O x",{"type":"Superscript","props":{},"children":[{"type":"Text","props":{"value":2},"children":[]}]},{"type":"AllCaps","props":{},"children":["draft"]},{"type":"HiddenText","props":{},"children":["internal"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_basic_formatting_wrappers() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Bold","props":{},"children":["bold"]},{"type":"Italic","props":{},"children":["italic"]},{"type":"Underline","props":{"type":"wave"},"children":["underlined"]},{"type":"StrikeThrough","props":{},"children":["removed"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_underline_type() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Underline","props":{"type":"invalid"},"children":["text"]}]}]}]}}"#,
        );
        assert!(
            ir.validate()
                .expect_err("invalid underline type must fail")
                .to_string()
                .contains("invalid `type`")
        );
    }

    #[test]
    fn validate_should_reject_empty_or_nested_semantic_text_wrapper() {
        let empty = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Superscript","props":{},"children":[]}]}]}]}}"#,
        );
        assert!(
            empty
                .validate()
                .expect_err("empty superscript must fail")
                .to_string()
                .contains("requires text content")
        );

        let nested = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Subscript","props":{},"children":[{"type":"Run","props":{},"children":["2"]}]}]}]}]}}"#,
        );
        assert!(
            nested
                .validate()
                .expect_err("Run inside subscript must fail")
                .to_string()
                .contains("Subscript cannot contain Run")
        );
    }

    #[test]
    fn validate_should_accept_advanced_semantic_text_wrappers() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"DoubleStrike","props":{},"children":["obsolete"]},{"type":"SpacedText","props":{"amount":1.5},"children":["wide"]},{"type":"SpacedText","props":{"amount":-0.5},"children":["tight"]},{"type":"ScaledText","props":{"percent":125},"children":["scaled"]},{"type":"FitText","props":{"width":42,"id":7},"children":["fitted"]},{"type":"BorderedText","props":{"style":"double","size":1,"color":"336699","space":2},"children":["bordered"]},{"type":"ShadedText","props":{"fill":"FFF2CC","color":"336699","pattern":"pct20"},"children":["shaded"]},{"type":"SpecialHiddenText","props":{},"children":["metadata"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_spacing_and_scale() {
        let spacing = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"SpacedText","props":{"amount":"wide"},"children":["text"]}]}]}]}}"#,
        );
        assert!(
            spacing
                .validate()
                .expect_err("non-numeric spacing must fail")
                .to_string()
                .contains("finite numeric `amount`")
        );

        let scale = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"ScaledText","props":{"percent":601},"children":["text"]}]}]}]}}"#,
        );
        assert!(
            scale
                .validate()
                .expect_err("scale above 600 must fail")
                .to_string()
                .contains("integer from 1 to 600")
        );
    }

    #[test]
    fn validate_should_reject_invalid_fit_and_bordered_text() {
        let fit = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"FitText","props":{"id":4294967296},"children":["text"]}]}]}]}}"#,
        );
        assert!(
            fit.validate()
                .expect_err("missing width must fail")
                .to_string()
                .contains("positive `width`")
        );

        let border = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"BorderedText","props":{"style":"triple","space":-1},"children":["text"]}]}]}]}}"#,
        );
        assert!(
            border
                .validate()
                .expect_err("invalid border must fail")
                .to_string()
                .contains("style")
        );
    }

    #[test]
    fn validate_should_reject_invalid_shaded_text() {
        let missing_fill = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"ShadedText","props":{},"children":["text"]}]}]}]}}"#,
        );
        assert!(
            missing_fill
                .validate()
                .expect_err("missing fill must fail")
                .to_string()
                .contains("six-digit RGB `fill`")
        );

        let pattern = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"ShadedText","props":{"fill":"FFF2CC","pattern":"waves"},"children":["text"]}]}]}]}}"#,
        );
        assert!(
            pattern
                .validate()
                .expect_err("unknown pattern must fail")
                .to_string()
                .contains("pattern")
        );
    }
}
