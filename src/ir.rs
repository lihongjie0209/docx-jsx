use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};
use std::collections::{HashMap, HashSet};

use crate::error::{Error, Result};

const BORDER_STYLES: &[&str] = &[
    "nil",
    "none",
    "single",
    "thick",
    "double",
    "dotted",
    "dashed",
    "dotDash",
    "dotDotDash",
    "triple",
    "thinThickSmallGap",
    "thickThinSmallGap",
    "thinThickThinSmallGap",
    "thinThickMediumGap",
    "thickThinMediumGap",
    "thinThickThinMediumGap",
    "thinThickLargeGap",
    "thickThinLargeGap",
    "thinThickThinLargeGap",
    "wave",
    "doubleWave",
    "dashSmallGap",
    "dashDotStroked",
    "threeDEmboss",
    "threeDEngrave",
    "outset",
    "inset",
    "apples",
    "archedScallops",
    "babyPacifier",
    "babyRattle",
];

pub(crate) const PAGE_NUMBER_FORMATS: &[&str] = &[
    "decimal",
    "upperRoman",
    "lowerRoman",
    "upperLetter",
    "lowerLetter",
    "ordinal",
    "cardinalText",
    "ordinalText",
    "hex",
    "chicago",
    "ideographDigital",
    "japaneseCounting",
    "aiueo",
    "iroha",
    "decimalFullWidth",
    "decimalHalfWidth",
    "japaneseLegal",
    "japaneseDigitalTenThousand",
    "decimalEnclosedCircle",
    "decimalFullWidth2",
    "aiueoFullWidth",
    "irohaFullWidth",
    "decimalZero",
    "bullet",
    "ganada",
    "chosung",
    "decimalEnclosedFullstop",
    "decimalEnclosedParen",
    "decimalEnclosedCircleChinese",
    "ideographEnclosedCircle",
    "ideographTraditional",
    "ideographZodiac",
    "ideographZodiacTraditional",
    "taiwaneseCounting",
    "ideographLegalTraditional",
    "taiwaneseCountingThousand",
    "taiwaneseDigital",
    "chineseCounting",
    "chineseLegalSimplified",
    "chineseCountingThousand",
    "koreanDigital",
    "koreanCounting",
    "koreanLegal",
    "koreanDigital2",
    "vietnameseCounting",
    "russianLower",
    "russianUpper",
    "none",
    "numberInDash",
    "hebrew1",
    "hebrew2",
    "arabicAlpha",
    "arabicAbjad",
    "hindiVowels",
    "hindiConsonants",
    "hindiNumbers",
    "hindiCounting",
    "thaiLetters",
    "thaiNumbers",
    "thaiCounting",
    "bahtText",
    "dollarText",
    "custom",
];

pub(crate) const PAGE_NUMBER_CHAPTER_SEPARATORS: &[&str] =
    &["hyphen", "period", "colon", "emDash", "enDash"];

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
    InlineBookmark,
    TableOfContents,
    TableOfFigures,
    TableOfEntries,
    TocItem,
    TocBefore,
    TocAfter,
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
    DeletedInstruction,
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
            Self::InlineBookmark => "InlineBookmark",
            Self::TableOfContents => "TableOfContents",
            Self::TableOfFigures => "TableOfFigures",
            Self::TableOfEntries => "TableOfEntries",
            Self::TocItem => "TocItem",
            Self::TocBefore => "TocBefore",
            Self::TocAfter => "TocAfter",
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
            Self::DeletedInstruction => "DeletedInstruction",
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
        validate_deleted_instruction_context(&self.document, false, "Document")?;
        validate_bookmark_references(&self.document)?;
        validate_style_relationships(&self.document)
    }
}

fn validate_style_relationships(root: &Node) -> Result<()> {
    let Some(styles) = root.props.get("styles").and_then(Value::as_array) else {
        return Ok(());
    };
    let mut types = HashMap::new();
    let mut based_on = HashMap::new();
    let mut next_styles = HashMap::new();
    let mut linked_styles = HashMap::new();
    for style in styles {
        let Some(style) = style.as_object() else {
            continue;
        };
        let (Some(id), Some(style_type)) = (
            style.get("id").and_then(Value::as_str),
            style.get("type").and_then(Value::as_str),
        ) else {
            continue;
        };
        types.insert(id, style_type);
        if let Some(base) = style.get("basedOn").and_then(Value::as_str) {
            based_on.insert(id, base);
            if let Some(base_type) = types.get(base).copied() {
                ensure_matching_style_types(id, style_type, base, base_type)?;
            }
        }
        if let Some(next) = style.get("next").and_then(Value::as_str) {
            next_styles.insert(id, next);
        }
        if let Some(link) = style.get("link").and_then(Value::as_str) {
            linked_styles.insert(id, link);
        }
    }
    // A base style can be declared after its derived style.
    for (&id, &base) in &based_on {
        if let (Some(style_type), Some(base_type)) = (types.get(id), types.get(base)) {
            ensure_matching_style_types(id, style_type, base, base_type)?;
        }
    }
    for &start in types.keys() {
        let mut seen = HashSet::new();
        let mut current = start;
        while let Some(&base) = based_on.get(current) {
            if !seen.insert(current) {
                return Err(validation(
                    "Document/styles",
                    format!(
                        "style inheritance cycle includes `{current}`; remove or redirect one `basedOn` reference"
                    ),
                ));
            }
            if !types.contains_key(base) {
                break;
            }
            current = base;
        }
    }
    validate_next_style_relationships(&types, &next_styles)?;
    validate_linked_style_relationships(&types, &linked_styles)?;

    validate_style_references(root, "Document", &types)
}

fn validate_next_style_relationships(
    types: &HashMap<&str, &str>,
    next_styles: &HashMap<&str, &str>,
) -> Result<()> {
    for (&id, &next) in next_styles {
        if types.get(id).copied() != Some("paragraph") {
            return Err(validation(
                "Document/styles",
                format!(
                    "style `{id}`: `next` is only valid on paragraph styles; remove `next` or change the style type"
                ),
            ));
        }
        if let Some(next_type) = types.get(next)
            && *next_type != "paragraph"
        {
            return Err(validation(
                "Document/styles",
                format!(
                    "style `{id}`: `next` must reference a paragraph style, but `{next}` is {next_type}; choose a paragraph style id"
                ),
            ));
        }
    }
    Ok(())
}

fn validate_linked_style_relationships(
    types: &HashMap<&str, &str>,
    linked_styles: &HashMap<&str, &str>,
) -> Result<()> {
    for (&id, &link) in linked_styles {
        let Some(style_type) = types.get(id).copied() else {
            continue;
        };
        if !matches!(style_type, "paragraph" | "character") {
            return Err(validation(
                "Document/styles",
                format!(
                    "style `{id}`: `link` is only valid for a paragraph and character style pair; remove `link`"
                ),
            ));
        }
        let Some(link_type) = types.get(link).copied() else {
            continue;
        };
        if style_type == link_type || !matches!(link_type, "paragraph" | "character") {
            return Err(validation(
                "Document/styles",
                format!(
                    "styles `{id}` and `{link}` must form a paragraph and character style pair; change one style type"
                ),
            ));
        }
        if linked_styles.get(link).copied() != Some(id) {
            return Err(validation(
                "Document/styles",
                format!(
                    "linked style `{link}` must link back to `{id}`; set its `link` property to `{id}`"
                ),
            ));
        }
    }
    Ok(())
}

fn validate_style_references(node: &Node, path: &str, types: &HashMap<&str, &str>) -> Result<()> {
    if let Some(style) = node.props.get("style").and_then(Value::as_str)
        && let Some(actual) = types.get(style).copied()
    {
        let expected = match node.kind {
            NodeKind::Run => Some("character"),
            NodeKind::Table => Some("table"),
            NodeKind::Paragraph | NodeKind::Heading | NodeKind::Caption | NodeKind::Index => {
                Some("paragraph")
            }
            _ => None,
        };
        if let Some(expected) = expected
            && actual != expected
        {
            return Err(validation(
                path,
                format!(
                    "{} `style` requires a {expected} style, but `{style}` is {actual}; reference a matching style id",
                    node.kind.name()
                ),
            ));
        }
    }
    for (index, child) in node.children.iter().enumerate() {
        if let Child::Node(child) = child {
            validate_style_references(
                child,
                &format!("{path}/{}[{index}]", child.kind.name()),
                types,
            )?;
        }
    }
    Ok(())
}

fn ensure_matching_style_types(
    id: &str,
    style_type: &str,
    base: &str,
    base_type: &str,
) -> Result<()> {
    if style_type == base_type {
        return Ok(());
    }
    Err(validation(
        "Document/styles",
        format!(
            "style `{id}` and its base `{base}` must have the same type; use a {style_type} base style"
        ),
    ))
}

fn validate_bookmark_references(root: &Node) -> Result<()> {
    fn collect<'a>(node: &'a Node, names: &mut Vec<&'a str>, anchors: &mut Vec<&'a str>) {
        if matches!(node.kind, NodeKind::Bookmark | NodeKind::InlineBookmark)
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
    validate_toc_wrappers(node, path)?;
    if node.kind == NodeKind::Deleted {
        validate_deleted_children(node, path)?;
    }
    if node.kind == NodeKind::DeletedInstruction {
        if node
            .props
            .get("instruction")
            .and_then(Value::as_str)
            .is_none_or(str::is_empty)
        {
            return Err(validation(
                path,
                "DeletedInstruction requires a non-empty `instruction` string",
            ));
        }
        if !node.children.is_empty() {
            return Err(validation(path, "DeletedInstruction cannot have children"));
        }
    }
    validate_toc_item(node, path)?;

    for (index, child) in node.children.iter().enumerate() {
        match child {
            Child::Node(child_node) => {
                validate_child_node(node.kind, child_node, path, index)?;
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
                        | NodeKind::InlineBookmark
                ) && !node.kind.is_semantic_text()
                {
                    return Err(validation(path, "text is not allowed here"));
                }
            }
        }
    }
    Ok(())
}

fn validate_deleted_instruction_context(
    node: &Node,
    inside_deleted: bool,
    path: &str,
) -> Result<()> {
    let inside_deleted = inside_deleted || node.kind == NodeKind::Deleted;
    if node.kind == NodeKind::DeletedInstruction && !inside_deleted {
        return Err(validation(
            path,
            "DeletedInstruction must be inside a Run nested within Deleted",
        ));
    }
    for (index, child) in node.children.iter().enumerate() {
        if let Child::Node(child_node) = child {
            validate_deleted_instruction_context(
                child_node,
                inside_deleted,
                &format!("{path}/{}[{index}]", child_node.kind.name()),
            )?;
        }
    }
    Ok(())
}

fn validate_deleted_children(node: &Node, path: &str) -> Result<()> {
    for child in &node.children {
        if let Child::Node(run) = child
            && run.kind == NodeKind::Run
            && run
                .children
                .iter()
                .any(|child| matches!(child, Child::Node(node) if !matches!(node.kind, NodeKind::Text | NodeKind::DeletedInstruction)))
        {
            return Err(validation(path, "Deleted Run only accepts text content"));
        }
        let Child::Node(comment) = child else {
            continue;
        };
        if comment.kind != NodeKind::Comment {
            continue;
        }
        if comment.children.iter().any(|selected| match selected {
            Child::String(_) | Child::Number(_) => false,
            Child::Node(text) if text.kind == NodeKind::Text => false,
            Child::Node(run) if run.kind == NodeKind::Run => run.children.iter().any(|child| {
                !matches!(child, Child::String(_) | Child::Number(_))
                    && !matches!(child, Child::Node(node) if node.kind == NodeKind::Text)
            }),
            Child::Node(_) => true,
        }) {
            return Err(validation(
                path,
                "Comment inside Deleted can select only text (`Text`, `Run`, string, or number); move formatting to the `Run` or keep the comment outside the deletion",
            ));
        }
    }
    Ok(())
}

fn validate_toc_wrappers(node: &Node, path: &str) -> Result<()> {
    if matches!(
        node.kind,
        NodeKind::TableOfContents | NodeKind::TableOfFigures | NodeKind::TableOfEntries
    ) {
        for wrapper in [NodeKind::TocBefore, NodeKind::TocAfter] {
            let count = node
                .children
                .iter()
                .filter(|child| matches!(child, Child::Node(child) if child.kind == wrapper))
                .count();
            if count > 1 {
                return Err(validation(
                    path,
                    format!("index allows at most one `{}` wrapper", wrapper.name()),
                ));
            }
        }
    }
    if matches!(node.kind, NodeKind::TocBefore | NodeKind::TocAfter) && node.children.is_empty() {
        return Err(validation(
            path,
            format!("{} requires at least one Paragraph", node.kind.name()),
        ));
    }
    Ok(())
}

fn validate_toc_item(node: &Node, path: &str) -> Result<()> {
    if node.kind != NodeKind::TocItem {
        return Ok(());
    }
    if node
        .props
        .get("text")
        .and_then(Value::as_str)
        .is_none_or(str::is_empty)
    {
        return Err(validation(path, "TocItem requires non-empty `text`"));
    }
    let bookmark = node.props.get("bookmark").and_then(Value::as_str);
    if bookmark.is_none_or(str::is_empty) {
        return Err(validation(path, "TocItem requires non-empty `bookmark`"));
    }
    if bookmark.is_some_and(|value| {
        value
            .chars()
            .any(|character| character.is_whitespace() || matches!(character, '"' | '\\'))
    }) {
        return Err(validation(
            path,
            "TocItem `bookmark` cannot contain whitespace, quotes, or backslashes because it is an unquoted PAGEREF field argument",
        ));
    }
    if node
        .props
        .get("level")
        .is_some_and(|value| value.as_u64().is_none_or(|level| !(1..=9).contains(&level)))
    {
        return Err(validation(
            path,
            "TocItem `level` must be an integer from 1 to 9",
        ));
    }
    if node
        .props
        .get("pageRef")
        .is_some_and(|value| !value.is_string())
    {
        return Err(validation(path, "TocItem `pageRef` must be a string"));
    }
    if node.props.get("tabPosition").is_some_and(|value| {
        value
            .as_f64()
            .is_none_or(|position| !position.is_finite() || position <= 0.0 || position > 1584.0)
    }) {
        return Err(validation(
            path,
            "TocItem `tabPosition` must be a positive point value no greater than 1584 (the OOXML 31680-twip limit); reduce the position to keep the page-number tab inside the page width",
        ));
    }
    Ok(())
}

fn validate_child_node(parent: NodeKind, child: &Node, path: &str, index: usize) -> Result<()> {
    if !allows(parent, child.kind) {
        return Err(validation(
            path,
            format!("{} cannot contain {}", parent.name(), child.kind.name()),
        ));
    }
    let child_path = format!("{path}/{}[{index}]", child.kind.name());
    if child.kind == NodeKind::ContentControl {
        validate_content_control_placement(parent, child, &child_path)?;
    }
    validate_node(child, &child_path)
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
                | NodeKind::ContentControl
        ),
        NodeKind::TableOfContents | NodeKind::TableOfFigures | NodeKind::TableOfEntries => {
            allows_index_child(child)
        }
        NodeKind::TocBefore | NodeKind::TocAfter => child == NodeKind::Paragraph,
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
        | NodeKind::TocEntry
        | NodeKind::Index
        | NodeKind::IndexEntry
        | NodeKind::Tab
        | NodeKind::TabStop
        | NodeKind::Symbol
        | NodeKind::PageReference
        | NodeKind::PositionalTab
        | NodeKind::DeletedInstruction
        | NodeKind::TocItem => false,
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
        NodeKind::TableCell => allows_table_cell_child(child),
        NodeKind::Header | NodeKind::Footer => allows_header_footer_child(child),
        NodeKind::Hyperlink => allows_hyperlink_child(child),
        NodeKind::ListItem
        | NodeKind::Inserted
        | NodeKind::MovedFrom
        | NodeKind::MovedTo
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
        NodeKind::ContentControl => allows_content_control_child(child),
        NodeKind::List => child == NodeKind::ListItem,
        NodeKind::Bookmark => matches!(
            child,
            NodeKind::Paragraph
                | NodeKind::Heading
                | NodeKind::Caption
                | NodeKind::Table
                | NodeKind::List
        ),
        NodeKind::InlineBookmark => allows_paragraph_child(child),
        NodeKind::Footnote => {
            child == NodeKind::Run || (allows_run_child(child) && child != NodeKind::Footnote)
        }
        NodeKind::Deleted => matches!(child, NodeKind::Run | NodeKind::Text | NodeKind::Comment),
    }
}

fn allows_index_child(child: NodeKind) -> bool {
    matches!(
        child,
        NodeKind::TocItem | NodeKind::TocBefore | NodeKind::TocAfter
    )
}

fn allows_header_footer_child(child: NodeKind) -> bool {
    matches!(
        child,
        NodeKind::Paragraph
            | NodeKind::Caption
            | NodeKind::Table
            | NodeKind::List
            | NodeKind::ContentControl
    )
}

fn allows_hyperlink_child(child: NodeKind) -> bool {
    child == NodeKind::Run
        || allows_run_child(child)
        || matches!(
            child,
            NodeKind::ContentControl
                | NodeKind::Inserted
                | NodeKind::Deleted
                | NodeKind::InlineBookmark
                | NodeKind::Comment
        )
}

fn allows_content_control_child(child: NodeKind) -> bool {
    child == NodeKind::Run || allows_run_child(child) || is_block_content_control_child(child)
}

fn is_block_content_control_child(child: NodeKind) -> bool {
    matches!(
        child,
        NodeKind::Paragraph
            | NodeKind::Heading
            | NodeKind::Caption
            | NodeKind::Table
            | NodeKind::List
    )
}

fn validate_content_control_placement(parent: NodeKind, control: &Node, path: &str) -> Result<()> {
    let has_block = control.children.iter().any(|child| match child {
        Child::Node(node) => is_block_content_control_child(node.kind),
        Child::String(_) | Child::Number(_) => false,
    });
    let has_inline = control.children.iter().any(|child| match child {
        Child::Node(node) => !is_block_content_control_child(node.kind),
        Child::String(_) | Child::Number(_) => true,
    });
    if has_block && has_inline {
        return Err(validation(
            path,
            "ContentControl cannot mix block and inline children; use only paragraphs/tables or only inline runs",
        ));
    }
    if matches!(
        parent,
        NodeKind::Section | NodeKind::Header | NodeKind::Footer
    ) {
        if !has_block {
            let placement = if parent == NodeKind::Section {
                "document-level"
            } else if parent == NodeKind::Header {
                "header"
            } else {
                "footer"
            };
            return Err(validation(
                path,
                format!(
                    "{placement} ContentControl requires block children (`Paragraph`, `Heading`, `Caption`, `Table`, or `List`)"
                ),
            ));
        }
    } else if has_block {
        return Err(validation(
            path,
            "inline ContentControl cannot contain Paragraph; move the control to the Section or unwrap the paragraph",
        ));
    }
    Ok(())
}

fn allows_table_cell_child(child: NodeKind) -> bool {
    matches!(
        child,
        NodeKind::Paragraph
            | NodeKind::Caption
            | NodeKind::Table
            | NodeKind::List
            | NodeKind::TableOfContents
            | NodeKind::ContentControl
    )
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
                | NodeKind::DeletedInstruction
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
                | NodeKind::InlineBookmark
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

#[allow(clippy::too_many_lines)]
fn allowed_props(kind: NodeKind) -> &'static [&'static str] {
    match kind {
        NodeKind::Document => document_props(),
        NodeKind::Section => section_props(),
        NodeKind::Paragraph => paragraph_props(),
        NodeKind::Heading => heading_props(),
        NodeKind::Caption => caption_props(),
        NodeKind::Run => &[
            "style",
            "font",
            "fonts",
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
        NodeKind::Image => image_props(),
        NodeKind::Table => table_props(),
        NodeKind::TableRow => &[
            "height",
            "heightRule",
            "cantSplit",
            "gridBefore",
            "gridAfter",
            "widthBefore",
            "widthAfter",
            "inserted",
            "deleted",
        ],
        NodeKind::TableCell => &[
            "width",
            "widthPercent",
            "colSpan",
            "verticalAlign",
            "verticalMerge",
            "textDirection",
            "margins",
            "shading",
            "border",
        ],
        NodeKind::Hyperlink => &["href", "anchor", "history"],
        NodeKind::PageNumber
        | NodeKind::TotalPages
        | NodeKind::Footnote
        | NodeKind::Tab
        | NodeKind::CarriageReturn
        | NodeKind::NonBreakingSpace
        | NodeKind::SoftHyphen
        | NodeKind::NonBreakingHyphen
        | NodeKind::TocBefore
        | NodeKind::TocAfter => &[],
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
        NodeKind::List => &[
            "type",
            "start",
            "multiLevelType",
            "levels",
            "overrides",
            "styleLink",
            "numStyleLink",
        ],
        NodeKind::ListItem => &[
            "level",
            "style",
            "suppressLineNumbers",
            "indentLeftChars",
            "indentStartChars",
            "indentRightChars",
            "indentEndChars",
            "firstLineChars",
            "hangingChars",
        ],
        NodeKind::Bookmark | NodeKind::InlineBookmark => &["name"],
        NodeKind::DeletedInstruction => &["instruction"],
        NodeKind::TocItem => &["text", "bookmark", "level", "pageRef", "tabPosition"],
        NodeKind::TableOfContents => &[
            "startLevel",
            "endLevel",
            "tcFieldLevelRange",
            "omitPageNumbersLevelRange",
            "entryBookmarkName",
            "entryAndPageNumberSeparator",
            "sequenceAndPageNumbersSeparator",
            "seqFieldIdentifierForPrefix",
            "stylesWithLevels",
            "preserveTab",
            "preserveNewLine",
            "useAppliedParagraphLineLevel",
            "hideTabAndPageNumbersInWebview",
            "hyperlinks",
            "dirty",
            "alias",
        ],
        NodeKind::TableOfFigures => &[
            "label",
            "includeLabelAndNumber",
            "separator",
            "entryBookmarkName",
            "entryAndPageNumberSeparator",
            "sequenceAndPageNumbersSeparator",
            "seqFieldIdentifierForPrefix",
            "stylesWithLevels",
            "preserveTab",
            "preserveNewLine",
            "useAppliedParagraphLineLevel",
            "hideTabAndPageNumbersInWebview",
            "hyperlinks",
            "dirty",
            "alias",
        ],
        NodeKind::TableOfEntries => &[
            "identifier",
            "tcFieldLevelRange",
            "omitPageNumbersLevelRange",
            "entryBookmarkName",
            "entryAndPageNumberSeparator",
            "sequenceAndPageNumbersSeparator",
            "seqFieldIdentifierForPrefix",
            "stylesWithLevels",
            "preserveTab",
            "preserveNewLine",
            "useAppliedParagraphLineLevel",
            "hideTabAndPageNumbersInWebview",
            "hyperlinks",
            "dirty",
            "alias",
        ],
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

fn image_props() -> &'static [&'static str] {
    &[
        "src",
        "width",
        "height",
        "relationshipId",
        "rotate",
        "floating",
        "allowOverlap",
        "layoutInCell",
        "simplePosition",
        "positionH",
        "positionV",
        "relativeFromH",
        "relativeFromV",
        "distanceTop",
        "distanceBottom",
        "distanceLeft",
        "distanceRight",
        "relativeHeight",
    ]
}

fn section_props() -> &'static [&'static str] {
    &[
        "pageSize",
        "orientation",
        "margins",
        "titlePage",
        "textDirection",
        "documentGrid",
        "pageNumbering",
        "lineNumbering",
        "columns",
        "type",
    ]
}

fn document_props() -> &'static [&'static str] {
    &[
        "defaultFont",
        "defaultFonts",
        "defaultSize",
        "defaultCharacterSpacing",
        "defaultLineSpacing",
        "webExtensions",
        "customXmlItems",
        "createdAt",
        "updatedAt",
        "creator",
        "lastModifiedBy",
        "description",
        "language",
        "revision",
        "subject",
        "title",
        "documentProtection",
        "customProperties",
        "documentId",
        "defaultTabStop",
        "documentVariables",
        "evenAndOddHeaders",
        "adjustLineHeightInTable",
        "characterSpacingControl",
        "styles",
    ]
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
        "spacingBeforeLines",
        "spacingAfterLines",
        "lineRule",
        "indentLeft",
        "indentRight",
        "indentLeftChars",
        "indentStartChars",
        "indentRightChars",
        "indentEndChars",
        "firstLine",
        "hanging",
        "firstLineChars",
        "hangingChars",
        "keepNext",
        "keepLines",
        "pageBreakBefore",
        "snapToGrid",
        "widowControl",
        "suppressLineNumbers",
        "font",
        "fonts",
        "size",
        "bold",
        "italic",
        "color",
        "characterSpacing",
        "bidi",
        "textAlign",
        "adjustRightIndent",
        "shading",
        "outlineLevel",
        "frame",
        "inserted",
        "deleted",
        "propertyChange",
        "paragraphId",
        "border",
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
        "spacingBeforeLines",
        "spacingAfterLines",
        "lineRule",
        "indentLeft",
        "indentRight",
        "indentLeftChars",
        "indentStartChars",
        "indentRightChars",
        "indentEndChars",
        "firstLine",
        "hanging",
        "firstLineChars",
        "hangingChars",
        "keepNext",
        "keepLines",
        "pageBreakBefore",
        "snapToGrid",
        "widowControl",
        "suppressLineNumbers",
        "font",
        "fonts",
        "size",
        "bold",
        "italic",
        "color",
        "characterSpacing",
        "bidi",
        "textAlign",
        "adjustRightIndent",
        "shading",
        "frame",
        "inserted",
        "deleted",
        "propertyChange",
        "paragraphId",
        "border",
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
        "style",
        "indent",
        "margins",
        "position",
        "border",
    ]
}

fn validate_semantics(node: &Node, path: &str) -> Result<()> {
    for key in ["defaultSize", "size", "width", "height"] {
        if let Some(value) = node.props.get(key) {
            require_number(value, path, key, true)?;
        }
    }
    for key in ["indentLeft", "indentRight"] {
        if let Some(value) = node.props.get(key) {
            value
                .as_f64()
                .filter(|number| number.is_finite())
                .ok_or_else(|| validation(path, format!("`{key}` must be a finite number")))?;
        }
    }
    for key in [
        "spacingBefore",
        "spacingAfter",
        "lineSpacing",
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
        validate_image_semantics(node, path)?;
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
    if node.kind == NodeKind::List {
        validate_list_semantics(node, path)?;
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

#[expect(
    clippy::too_many_lines,
    reason = "image properties are validated together to keep the DSL contract explicit"
)]
fn validate_image_semantics(node: &Node, path: &str) -> Result<()> {
    node.props
        .get("src")
        .and_then(Value::as_str)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| validation(path, "Image `src` must be a non-empty string"))?;
    if node
        .props
        .get("relationshipId")
        .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
    {
        return Err(validation(
            path,
            "`relationshipId` must be a non-empty string",
        ));
    }
    if let Some(value) = node.props.get("simplePosition") {
        let position = value.as_object().ok_or_else(|| {
            validation(path, "`simplePosition` must be an object with signed integer `x` and `y` EMU coordinates")
        })?;
        validate_object_keys(position, &format!("{path}.simplePosition"), &["x", "y"])?;
        for axis in ["x", "y"] {
            position
                .get(axis)
                .and_then(Value::as_i64)
                .and_then(|value| i32::try_from(value).ok())
                .ok_or_else(|| {
                    validation(
                        path,
                        format!("`simplePosition.{axis}` must be a signed 32-bit integer in EMUs"),
                    )
                })?;
        }
        if node.props.get("floating").and_then(Value::as_bool) != Some(true) {
            return Err(validation(
                path,
                "`simplePosition` requires `floating={true}`",
            ));
        }
        if node
            .props
            .get("relationshipId")
            .and_then(Value::as_str)
            .is_some_and(|id| !is_xml_ncname(id) || id.starts_with("rIdDocxJsxSimplePosition_"))
        {
            return Err(validation(
                path,
                "`relationshipId` used with `simplePosition` must be an XML NCName and must not use the reserved `rIdDocxJsxSimplePosition_` prefix",
            ));
        }
    }
    if let Some(value) = node.props.get("rotate") {
        value
            .as_u64()
            .filter(|angle| u16::try_from(*angle).is_ok())
            .ok_or_else(|| validation(path, "`rotate` must be an integer from 0 through 65535"))?;
    }
    for key in ["floating", "allowOverlap", "layoutInCell"] {
        if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(path, format!("`{key}` must be a boolean")));
        }
    }
    validate_image_position(node, path, "positionH", &["left", "center", "right"])?;
    validate_image_position(node, path, "positionV", &["top", "center", "bottom"])?;
    validate_optional_enum_prop(
        node,
        path,
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
    )?;
    validate_optional_enum_prop(
        node,
        path,
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
    )?;
    for key in [
        "distanceTop",
        "distanceBottom",
        "distanceLeft",
        "distanceRight",
    ] {
        if let Some(value) = node.props.get(key) {
            require_number(value, path, key, false)?;
        }
    }
    if node
        .props
        .get("relativeHeight")
        .is_some_and(|value| value.as_u64().is_none())
    {
        return Err(validation(
            path,
            "`relativeHeight` must be a non-negative integer",
        ));
    }
    let anchor_props = [
        "allowOverlap",
        "layoutInCell",
        "simplePosition",
        "positionH",
        "positionV",
        "relativeFromH",
        "relativeFromV",
        "distanceTop",
        "distanceBottom",
        "distanceLeft",
        "distanceRight",
        "relativeHeight",
    ];
    if anchor_props.iter().any(|key| node.props.contains_key(*key))
        && node.props.get("floating").and_then(Value::as_bool) != Some(true)
    {
        return Err(validation(
            path,
            "advanced image anchor properties require `floating={true}`",
        ));
    }
    Ok(())
}

fn validate_list_semantics(node: &Node, path: &str) -> Result<()> {
    if node.children.is_empty() {
        return Err(validation(path, "List requires at least one ListItem"));
    }
    validate_optional_enum_prop(
        node,
        path,
        "multiLevelType",
        &["singleLevel", "multilevel", "hybridMultilevel"],
    )?;
    for key in ["styleLink", "numStyleLink"] {
        if let Some(value) = node.props.get(key) {
            value
                .as_str()
                .filter(|value| is_xml_ncname(value))
                .ok_or_else(|| {
                    validation(
                        path,
                        format!("`{key}` must be a non-empty XML style identifier"),
                    )
                })?;
        }
    }
    if node.props.contains_key("numStyleLink")
        && ["type", "start", "levels"]
            .iter()
            .any(|key| node.props.contains_key(*key))
    {
        return Err(validation(
            path,
            "`numStyleLink` supplies the numbering definition; remove local `type`, `start`, and `levels`, or use `styleLink` for a base definition",
        ));
    }
    validate_list_levels(node, path)
}

fn is_xml_ncname(value: &str) -> bool {
    let mut bytes = value.bytes();
    bytes
        .next()
        .is_some_and(|first| first.is_ascii_alphabetic() || first == b'_')
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
}

fn validate_image_position(node: &Node, path: &str, key: &str, alignments: &[&str]) -> Result<()> {
    let Some(value) = node.props.get(key) else {
        return Ok(());
    };
    if value.as_f64().is_some_and(f64::is_finite) {
        return Ok(());
    }
    let alignment = value
        .as_str()
        .ok_or_else(|| validation(path, format!("`{key}` must be a point offset or alignment")))?;
    if !alignments.contains(&alignment) {
        return Err(validation(
            path,
            format!("invalid `{key}` value `{alignment}`"),
        ));
    }
    Ok(())
}

fn validate_paragraph_defaults(node: &Node, path: &str) -> Result<()> {
    if matches!(node.kind, NodeKind::Paragraph | NodeKind::Heading) {
        validate_character_indent_props(&node.props, path, "")?;
        for key in [
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "snapToGrid",
            "widowControl",
            "suppressLineNumbers",
            "bold",
            "italic",
            "bidi",
        ] {
            if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
                return Err(validation(path, format!("`{key}` must be a boolean")));
            }
        }
        validate_font_choice(&node.props, path, "font", "fonts")?;
        if let Some(value) = node.props.get("characterSpacing") {
            value
                .as_f64()
                .filter(|number| number.is_finite())
                .ok_or_else(|| validation(path, "`characterSpacing` must be a finite number"))?;
        }
        validate_line_spacing_extras(
            &node.props,
            path,
            "spacingBeforeLines",
            "spacingAfterLines",
            "lineRule",
        )?;
        validate_optional_enum_prop(
            node,
            path,
            "align",
            &[
                "left",
                "right",
                "center",
                "both",
                "distribute",
                "start",
                "end",
                "justified",
            ],
        )?;
        validate_optional_enum_prop(
            node,
            path,
            "textAlign",
            &["auto", "baseline", "bottom", "center", "top"],
        )?;
        if let Some(value) = node.props.get("adjustRightIndent") {
            let number = value
                .as_i64()
                .and_then(|number| isize::try_from(number).ok())
                .ok_or_else(|| validation(path, "`adjustRightIndent` must be 0 or 1"))?;
            if !matches!(number, 0 | 1) {
                return Err(validation(
                    path,
                    "`adjustRightIndent` must be 0 or 1 because OOXML defines w:adjustRightInd as on/off",
                ));
            }
        }
        if let Some(value) = node.props.get("outlineLevel")
            && !matches!(value.as_u64(), Some(0..=9))
        {
            return Err(validation(
                path,
                "`outlineLevel` must be an integer from 0 through 9",
            ));
        }
        if let Some(value) = node.props.get("frame") {
            validate_style_frame(value, path)?;
        }
        if let Some(value) = node.props.get("paragraphId") {
            let valid = value.as_str().is_some_and(|value| {
                value.len() == 8 && value.bytes().all(|byte| byte.is_ascii_hexdigit())
            });
            if !valid {
                return Err(validation(
                    path,
                    "`paragraphId` must contain exactly eight hexadecimal characters",
                ));
            }
        }
        if let Some(value) = node.props.get("border") {
            validate_paragraph_border_model(value, path)?;
        }
        validate_paragraph_revisions(node, path)?;
    }
    Ok(())
}

fn validate_paragraph_revisions(node: &Node, path: &str) -> Result<()> {
    if node.props.contains_key("inserted") && node.props.contains_key("deleted") {
        return Err(validation(
            path,
            "`inserted` and `deleted` are mutually exclusive",
        ));
    }
    for key in ["inserted", "deleted"] {
        if let Some(value) = node.props.get(key) {
            validate_paragraph_revision_metadata(value, path, key)?;
        }
    }
    let Some(value) = node.props.get("propertyChange") else {
        return Ok(());
    };
    let change = value
        .as_object()
        .ok_or_else(|| validation(path, "`propertyChange` must be an object"))?;
    validate_object_keys(change, path, &["author", "date", "previous"])?;
    for key in ["author", "date"] {
        if change
            .get(key)
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(
                path,
                format!("`propertyChange.{key}` must be a non-empty string"),
            ));
        }
    }
    let previous = change
        .get("previous")
        .and_then(Value::as_object)
        .ok_or_else(|| validation(path, "`propertyChange.previous` must be an object"))?;
    if previous.is_empty() {
        return Err(validation(
            path,
            "`propertyChange.previous` must not be empty",
        ));
    }
    if let Some(key) = previous
        .keys()
        .find(|key| ["inserted", "deleted", "propertyChange"].contains(&key.as_str()))
    {
        return Err(validation(
            path,
            format!("`propertyChange.previous` cannot contain `{key}`"),
        ));
    }
    validate_props(
        &Node {
            kind: NodeKind::Paragraph,
            props: previous.clone(),
            children: Vec::new(),
        },
        &format!("{path}/propertyChange/previous"),
    )
}

fn validate_paragraph_revision_metadata(value: &Value, path: &str, key: &str) -> Result<()> {
    let revision = value
        .as_object()
        .ok_or_else(|| validation(path, format!("`{key}` must be an object")))?;
    validate_object_keys(revision, path, &["author", "date"])?;
    for field in ["author", "date"] {
        if revision
            .get(field)
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(
                path,
                format!("`{key}.{field}` must be a non-empty string"),
            ));
        }
    }
    Ok(())
}

fn validate_advanced_semantics(node: &Node, path: &str) -> Result<()> {
    validate_document_semantics(node, path)?;
    validate_section_semantics(node, path)?;
    validate_run_defaults(node, path)?;
    validate_field_semantics(node, path)?;
    validate_index_semantics(node, path)?;
    validate_text_effect_semantics(node, path)?;
    validate_revision_and_control_semantics(node, path)?;
    validate_annotation_semantics(node, path)?;
    validate_table_semantics(node, path)?;
    validate_structure_semantics(node, path)
}

fn validate_line_spacing_extras(
    object: &Map<String, Value>,
    path: &str,
    before_lines: &str,
    after_lines: &str,
    line_rule: &str,
) -> Result<()> {
    for key in [before_lines, after_lines] {
        if object.get(key).is_some_and(|value| {
            value
                .as_u64()
                .and_then(|value| u32::try_from(value).ok())
                .is_none()
        }) {
            return Err(validation(
                path,
                format!("`{key}` must be a non-negative integer"),
            ));
        }
    }
    if let Some(value) = object.get(line_rule) {
        let valid = value
            .as_str()
            .is_some_and(|value| ["auto", "atLeast", "exact"].contains(&value));
        if !valid {
            return Err(validation(
                path,
                format!("invalid `{line_rule}`; expected auto, atLeast, or exact"),
            ));
        }
    }
    Ok(())
}

fn validate_document_semantics(node: &Node, path: &str) -> Result<()> {
    if node.kind != NodeKind::Document {
        return Ok(());
    }
    validate_default_fonts(node, path)?;
    if let Some(value) = node.props.get("defaultCharacterSpacing") {
        value
            .as_f64()
            .filter(|number| number.is_finite())
            .ok_or_else(|| validation(path, "`defaultCharacterSpacing` must be finite"))?;
    }
    if let Some(value) = node.props.get("defaultLineSpacing") {
        let spacing = value
            .as_object()
            .ok_or_else(|| validation(path, "`defaultLineSpacing` must be an object"))?;
        validate_object_keys(
            spacing,
            path,
            &[
                "before",
                "after",
                "line",
                "beforeLines",
                "afterLines",
                "lineRule",
            ],
        )?;
        if spacing.is_empty() {
            return Err(validation(path, "`defaultLineSpacing` must not be empty"));
        }
        for key in ["before", "after"] {
            if let Some(value) = spacing.get(key) {
                require_number(value, path, &format!("defaultLineSpacing.{key}"), false)?;
                if value.as_f64().is_some_and(|value| value < 0.0) {
                    return Err(validation(
                        path,
                        format!("`defaultLineSpacing.{key}` must be non-negative"),
                    ));
                }
            }
        }
        if let Some(value) = spacing.get("line") {
            require_number(value, path, "defaultLineSpacing.line", false)?;
        }
        validate_line_spacing_extras(spacing, path, "beforeLines", "afterLines", "lineRule")?;
    }
    if let Some(value) = node.props.get("webExtensions") {
        validate_web_extensions(value, path)?;
    }
    if let Some(value) = node.props.get("customXmlItems") {
        validate_custom_xml_items(value, path)?;
    }
    if let Some(value) = node.props.get("defaultTabStop") {
        require_number(value, path, "defaultTabStop", true)?;
    }
    validate_document_core_metadata(&node.props, path)?;
    if let Some(value) = node.props.get("documentProtection") {
        validate_document_protection(value, path)?;
    }
    for key in ["customProperties", "documentVariables"] {
        if let Some(value) = node.props.get(key) {
            validate_string_map(value, path, key)?;
        }
    }
    for key in ["evenAndOddHeaders", "adjustLineHeightInTable"] {
        if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(path, format!("`{key}` must be a boolean")));
        }
    }
    validate_optional_enum_prop(
        node,
        path,
        "characterSpacingControl",
        &[
            "doNotCompress",
            "compressPunctuation",
            "compressPunctuationAndJapaneseKana",
        ],
    )?;
    if let Some(value) = node.props.get("styles") {
        validate_style_definitions(value, path)?;
    }
    Ok(())
}

fn validate_document_core_metadata(props: &Map<String, Value>, path: &str) -> Result<()> {
    for key in [
        "createdAt",
        "updatedAt",
        "documentId",
        "creator",
        "lastModifiedBy",
        "description",
        "language",
        "subject",
        "title",
    ] {
        if props
            .get(key)
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(
                path,
                format!("`{key}` must be a non-empty string"),
            ));
        }
    }
    if let Some(value) = props.get("revision") {
        let revision = value
            .as_u64()
            .ok_or_else(|| validation(path, "`revision` must be a positive integer"))?;
        if revision == 0 {
            return Err(validation(path, "`revision` must be a positive integer"));
        }
    }
    Ok(())
}

pub(crate) fn validate_document_protection(value: &Value, path: &str) -> Result<()> {
    let protection_path = format!("{path}.documentProtection");
    let protection = value
        .as_object()
        .ok_or_else(|| validation(path, "`documentProtection` must be an object"))?;
    validate_object_keys(
        protection,
        &protection_path,
        &[
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
        ],
    )?;
    if protection.is_empty() {
        return Err(validation(
            &protection_path,
            "must configure at least one protection property",
        ));
    }
    validate_protection_enums(protection, &protection_path)?;
    validate_protection_numeric_and_boolean(protection, &protection_path)?;
    validate_protection_strings(protection, &protection_path)?;
    validate_protection_hash_and_salt(protection, &protection_path)
}

fn validate_protection_enums(protection: &Map<String, Value>, path: &str) -> Result<()> {
    for (key, values) in [
        (
            "edit",
            &["none", "readOnly", "comments", "trackedChanges", "forms"][..],
        ),
        ("cryptProviderType", &["rsaFull", "rsaAES"][..]),
        ("cryptAlgorithmClass", &["hash", "encryption", "custom"][..]),
        ("cryptAlgorithmType", &["typeAny", "custom"][..]),
    ] {
        if let Some(value) = protection.get(key) {
            let actual = value
                .as_str()
                .ok_or_else(|| validation(path, format!("`{key}` must be a string")))?;
            if !values.contains(&actual) {
                return Err(validation(
                    path,
                    format!(
                        "invalid `{key}` value `{actual}`; expected {}",
                        values.join(", ")
                    ),
                ));
            }
        }
    }
    Ok(())
}

fn validate_protection_numeric_and_boolean(
    protection: &Map<String, Value>,
    path: &str,
) -> Result<()> {
    for key in ["formatting", "enforcement"] {
        if protection.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(path, format!("`{key}` must be a boolean")));
        }
    }
    for key in ["cryptAlgorithmSid", "cryptSpinCount", "spinCount"] {
        if protection.get(key).is_some_and(|value| {
            value
                .as_u64()
                .is_none_or(|number| number > u64::from(u32::MAX))
        }) {
            return Err(validation(
                path,
                format!("`{key}` must be a non-negative 32-bit integer"),
            ));
        }
    }
    Ok(())
}

fn validate_protection_strings(protection: &Map<String, Value>, path: &str) -> Result<()> {
    for key in [
        "cryptProvider",
        "algIdExtSource",
        "cryptProviderTypeExtSource",
        "hash",
        "algorithmName",
        "hashValue",
        "saltValue",
        "salt",
    ] {
        if protection
            .get(key)
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(
                path,
                format!("`{key}` must be a non-empty string"),
            ));
        }
    }
    for key in ["algIdExt", "cryptProviderTypeExt"] {
        if let Some(encoded) = protection.get(key).and_then(Value::as_str)
            && !is_hex_binary_value(encoded)
        {
            return Err(validation(
                path,
                format!("`{key}` must be exactly 4 bytes (8 hexadecimal characters)"),
            ));
        }
    }
    Ok(())
}

fn validate_protection_hash_and_salt(protection: &Map<String, Value>, path: &str) -> Result<()> {
    if protection.contains_key("hash") != protection.contains_key("salt") {
        return Err(validation(
            path,
            "`hash` and `salt` must be supplied together; provide both precomputed values or omit both",
        ));
    }
    if protection.contains_key("hashValue") != protection.contains_key("saltValue") {
        return Err(validation(
            path,
            "`hashValue` and `saltValue` must be supplied together; provide both precomputed values or omit both",
        ));
    }
    for key in ["hash", "salt", "hashValue", "saltValue"] {
        if let Some(encoded) = protection.get(key).and_then(Value::as_str)
            && !is_base64_value(encoded)
        {
            return Err(validation(
                path,
                format!("`{key}` must be a non-empty base64 value"),
            ));
        }
    }
    Ok(())
}

fn is_base64_value(value: &str) -> bool {
    if value.is_empty() || !value.len().is_multiple_of(4) {
        return false;
    }
    let padding = value.bytes().rev().take_while(|byte| *byte == b'=').count();
    padding <= 2
        && value[..value.len() - padding]
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/'))
        && !value[..value.len() - padding].contains('=')
}

fn is_hex_binary_value(value: &str) -> bool {
    value.len() == 8 && value.bytes().all(|byte| byte.is_ascii_hexdigit())
}

fn validate_default_fonts(node: &Node, path: &str) -> Result<()> {
    if node.props.contains_key("defaultFont") && node.props.contains_key("defaultFonts") {
        return Err(validation(
            path,
            "`defaultFont` and `defaultFonts` cannot be used together; use `defaultFonts` when font slots differ",
        ));
    }
    if let Some(value) = node.props.get("defaultFonts") {
        let fonts = value
            .as_object()
            .ok_or_else(|| validation(path, "`defaultFonts` must be an object"))?;
        validate_object_keys(
            fonts,
            path,
            &[
                "ascii",
                "hiAnsi",
                "eastAsia",
                "cs",
                "asciiTheme",
                "hiAnsiTheme",
                "eastAsiaTheme",
                "csTheme",
                "hint",
            ],
        )?;
        if fonts.is_empty() {
            return Err(validation(path, "`defaultFonts` must not be empty"));
        }
        for (key, value) in fonts {
            if value.as_str().is_none_or(str::is_empty) {
                return Err(validation(
                    path,
                    format!("`defaultFonts.{key}` must be a non-empty string"),
                ));
            }
        }
    }
    Ok(())
}

fn validate_web_extensions(value: &Value, path: &str) -> Result<()> {
    let extensions = value
        .as_array()
        .ok_or_else(|| validation(path, "`webExtensions` must be an array"))?;
    if extensions.is_empty() {
        return Err(validation(path, "`webExtensions` must not be empty"));
    }
    let mut ids = HashSet::new();
    for (index, value) in extensions.iter().enumerate() {
        let extension_path = format!("{path}/webExtensions[{index}]");
        let extension = value
            .as_object()
            .ok_or_else(|| validation(&extension_path, "web extension must be an object"))?;
        validate_object_keys(
            extension,
            &extension_path,
            &[
                "id",
                "referenceId",
                "version",
                "store",
                "storeType",
                "properties",
            ],
        )?;
        for key in ["id", "referenceId", "version", "store", "storeType"] {
            if extension
                .get(key)
                .and_then(Value::as_str)
                .is_none_or(str::is_empty)
            {
                return Err(validation(
                    &extension_path,
                    format!("web extension requires non-empty `{key}`"),
                ));
            }
        }
        let id = extension
            .get("id")
            .and_then(Value::as_str)
            .unwrap_or_default();
        if !ids.insert(id) {
            return Err(validation(
                &extension_path,
                format!("duplicate web extension id `{id}`"),
            ));
        }
        if let Some(properties) = extension.get("properties") {
            let properties = properties
                .as_object()
                .ok_or_else(|| validation(&extension_path, "`properties` must be an object"))?;
            if properties.values().any(|value| !value.is_string()) {
                return Err(validation(
                    &extension_path,
                    "web extension property values must be strings",
                ));
            }
        }
    }
    Ok(())
}

fn validate_custom_xml_items(value: &Value, path: &str) -> Result<()> {
    let items = value
        .as_array()
        .ok_or_else(|| validation(path, "`customXmlItems` must be an array"))?;
    if items.is_empty() {
        return Err(validation(path, "`customXmlItems` must not be empty"));
    }
    let mut ids = HashSet::new();
    for (index, value) in items.iter().enumerate() {
        let item_path = format!("{path}/customXmlItems[{index}]");
        let item = value
            .as_object()
            .ok_or_else(|| validation(&item_path, "custom XML item must be an object"))?;
        validate_object_keys(item, &item_path, &["id", "xml", "schemaRefs"])?;
        let id = item
            .get("id")
            .and_then(Value::as_str)
            .filter(|value| !value.is_empty())
            .ok_or_else(|| validation(&item_path, "custom XML item requires non-empty `id`"))?;
        if !ids.insert(id) {
            return Err(validation(
                &item_path,
                format!("duplicate custom XML id `{id}`"),
            ));
        }
        let xml = item
            .get("xml")
            .and_then(Value::as_str)
            .filter(|value| !value.is_empty())
            .ok_or_else(|| validation(&item_path, "custom XML item requires non-empty `xml`"))?;
        if xml.parse::<docx_rs::CustomItem>().is_err() {
            return Err(validation(
                &item_path,
                "custom XML item `xml` must be well-formed XML",
            ));
        }
        if let Some(value) = item.get("schemaRefs") {
            let refs = value.as_array().ok_or_else(|| {
                validation(&item_path, "custom XML `schemaRefs` must be an array")
            })?;
            if refs.is_empty() {
                return Err(validation(
                    &item_path,
                    "custom XML `schemaRefs` must not be empty",
                ));
            }
            for uri in refs {
                if uri.as_str().is_none_or(str::is_empty) {
                    return Err(validation(
                        &item_path,
                        "custom XML `schemaRefs` entries must be non-empty strings",
                    ));
                }
            }
        }
    }
    Ok(())
}

fn validate_style_definitions(value: &Value, path: &str) -> Result<()> {
    let styles = value
        .as_array()
        .filter(|styles| !styles.is_empty())
        .ok_or_else(|| validation(path, "`styles` must be a non-empty array"))?;
    let mut ids = HashSet::new();
    for (index, value) in styles.iter().enumerate() {
        let style_path = format!("{path}/styles[{index}]");
        let style = value
            .as_object()
            .ok_or_else(|| validation(&style_path, "style definition must be an object"))?;
        validate_object_keys(
            style,
            &style_path,
            &[
                "id",
                "name",
                "type",
                "basedOn",
                "next",
                "link",
                "quickFormat",
                "uiPriority",
                "semiHidden",
                "unhideWhenUsed",
                "run",
                "paragraph",
                "table",
                "cell",
            ],
        )?;
        let (id, style_type) = validate_style_metadata(style, &style_path)?;
        if !ids.insert(id) {
            return Err(validation(
                &style_path,
                format!("duplicate style id `{id}`"),
            ));
        }
        if let Some(run) = style.get("run") {
            validate_style_run(run, &style_path)?;
        }
        if let Some(paragraph) = style.get("paragraph") {
            validate_style_paragraph(paragraph, &style_path)?;
        }
        if style_type != "table" && (style.contains_key("table") || style.contains_key("cell")) {
            return Err(validation(
                &style_path,
                "style `table` and `cell` properties require type `table`",
            ));
        }
        if let Some(table) = style.get("table") {
            validate_style_table(table, &style_path)?;
        }
        if let Some(cell) = style.get("cell") {
            validate_style_cell(cell, &style_path)?;
        }
    }
    Ok(())
}

fn validate_style_metadata<'a>(
    style: &'a Map<String, Value>,
    path: &str,
) -> Result<(&'a str, &'a str)> {
    for key in ["id", "name", "type"] {
        if style
            .get(key)
            .and_then(Value::as_str)
            .is_none_or(str::is_empty)
        {
            return Err(validation(
                path,
                format!("style requires non-empty `{key}`"),
            ));
        }
    }
    let id = style.get("id").and_then(Value::as_str).unwrap_or_default();
    let style_type = style
        .get("type")
        .and_then(Value::as_str)
        .unwrap_or_default();
    if !["paragraph", "character", "numbering", "table"].contains(&style_type) {
        return Err(validation(
            path,
            format!("invalid style type `{style_type}`"),
        ));
    }
    for key in ["basedOn", "next", "link"] {
        if style
            .get(key)
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(path, format!("`{key}` must be non-empty")));
        }
    }
    for key in ["quickFormat", "semiHidden", "unhideWhenUsed"] {
        if style.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(path, format!("`{key}` must be a boolean")));
        }
    }
    if style
        .get("uiPriority")
        .is_some_and(|value| value.as_u64().is_none())
    {
        return Err(validation(
            path,
            "`uiPriority` must be a non-negative integer",
        ));
    }
    Ok((id, style_type))
}

fn validate_style_run(value: &Value, path: &str) -> Result<()> {
    let run = value
        .as_object()
        .ok_or_else(|| validation(path, "style `run` must be an object"))?;
    validate_object_keys(
        run,
        path,
        &[
            "font",
            "fonts",
            "size",
            "color",
            "themeColor",
            "themeShade",
            "themeTint",
            "highlight",
            "bold",
            "italic",
            "strike",
            "doubleStrike",
            "caps",
            "underline",
            "hidden",
            "verticalAlign",
            "characterSpacing",
            "textBorder",
            "fitText",
            "stretch",
            "specVanish",
            "shading",
        ],
    )?;
    if let Some(value) = run.get("size") {
        require_number(value, path, "run.size", true)?;
    }
    validate_font_choice(run, path, "font", "fonts")?;
    for key in ["highlight"] {
        if run
            .get(key)
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(
                path,
                format!("style run `{key}` must be non-empty"),
            ));
        }
    }
    if let Some(value) = run.get("color") {
        validate_rgb(value, path, "style run `color`")?;
    }
    validate_theme_fields(run, path)?;
    validate_map_enum(
        run,
        path,
        "underline",
        &["single", "double", "dotted", "dash", "wave"],
    )?;
    for key in ["bold", "italic", "strike", "doubleStrike", "caps", "hidden"] {
        if run.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(
                path,
                format!("style run `{key}` must be a boolean"),
            ));
        }
    }
    if run.get("strike").and_then(Value::as_bool) == Some(true)
        && run.get("doubleStrike").and_then(Value::as_bool) == Some(true)
    {
        return Err(validation(
            path,
            "enabled style run `strike` and `doubleStrike` are mutually exclusive",
        ));
    }
    validate_map_enum(
        run,
        path,
        "verticalAlign",
        &["baseline", "superscript", "subscript"],
    )?;
    if let Some(value) = run.get("characterSpacing") {
        value
            .as_f64()
            .filter(|number| number.is_finite())
            .ok_or_else(|| {
                validation(path, "style run `characterSpacing` must be a finite number")
            })?;
    }
    validate_style_run_effects(run, path)?;
    if let Some(border) = run.get("textBorder") {
        validate_style_border(border, path, true)?;
    }
    Ok(())
}

fn validate_style_run_effects(run: &Map<String, Value>, path: &str) -> Result<()> {
    if let Some(value) = run.get("fitText") {
        let fit_text = value
            .as_object()
            .ok_or_else(|| validation(path, "style run `fitText` must be an object"))?;
        validate_object_keys(fit_text, path, &["width", "id"])?;
        let width = fit_text
            .get("width")
            .ok_or_else(|| validation(path, "style run `fitText` requires `width`"))?;
        require_number(width, path, "style run fitText.width", true)?;
        if let Some(id) = fit_text.get("id")
            && id.as_u64().is_none_or(|value| value > u64::from(u32::MAX))
        {
            return Err(validation(
                path,
                "style run `fitText.id` must be an unsigned 32-bit integer",
            ));
        }
    }
    if let Some(value) = run.get("stretch")
        && value
            .as_u64()
            .is_none_or(|value| !(1..=600).contains(&value))
    {
        return Err(validation(
            path,
            "style run `stretch` must be an integer from 1..=600",
        ));
    }
    if let Some(value) = run.get("specVanish")
        && value.as_bool() != Some(true)
    {
        return Err(validation(
            path,
            "style run `specVanish` only accepts `true`; docx-rs has no explicit-off writer value",
        ));
    }
    if let Some(value) = run.get("shading") {
        let shading = value
            .as_object()
            .ok_or_else(|| validation(path, "style run `shading` must be an object"))?;
        validate_object_keys(shading, path, &["fill", "color", "pattern"])?;
        let fill = shading
            .get("fill")
            .ok_or_else(|| validation(path, "style run `shading` requires six-digit RGB `fill`"))?;
        validate_rgb(fill, path, "style run shading `fill`")?;
        if let Some(color) = shading.get("color") {
            if color.as_str() == Some("auto") {
                // `auto` is a valid OOXML foreground color.
            } else {
                validate_rgb(color, path, "style run shading `color`")?;
            }
        }
        if let Some(pattern) = shading.get("pattern") {
            let pattern = pattern
                .as_str()
                .ok_or_else(|| validation(path, "style run shading `pattern` must be a string"))?;
            if !SHADING_PATTERNS.contains(&pattern) {
                return Err(validation(path, "style run shading `pattern` is invalid"));
            }
        }
    }
    Ok(())
}

fn validate_style_table(value: &Value, path: &str) -> Result<()> {
    let table = value
        .as_object()
        .ok_or_else(|| validation(path, "style `table` must be an object"))?;
    validate_object_keys(
        table,
        path,
        &[
            "style",
            "indent",
            "width",
            "widthPercent",
            "align",
            "layout",
            "margins",
            "border",
        ],
    )?;
    for key in ["style", "width", "widthPercent", "layout"] {
        if table.contains_key(key) {
            return Err(validation(
                path,
                format!(
                    "style table `{key}` is not legal in OOXML style properties; apply `{key}` to Table instead"
                ),
            ));
        }
    }
    if let Some(value) = table.get("indent") {
        value
            .as_f64()
            .filter(|number| number.is_finite())
            .ok_or_else(|| validation(path, "style table `indent` must be finite"))?;
    }
    validate_map_enum(table, path, "align", &["left", "center", "right"])?;
    if let Some(value) = table.get("margins") {
        validate_box_margins(value, path)?;
    }
    if let Some(value) = table.get("border") {
        validate_border_model(value, path, false)?;
    }
    Ok(())
}

fn validate_style_cell(value: &Value, path: &str) -> Result<()> {
    let cell = value
        .as_object()
        .ok_or_else(|| validation(path, "style `cell` must be an object"))?;
    validate_object_keys(
        cell,
        path,
        &[
            "width",
            "colSpan",
            "verticalAlign",
            "verticalMerge",
            "textDirection",
            "shading",
            "margins",
            "border",
        ],
    )?;
    if let Some(value) = cell.get("width") {
        require_number(value, path, "cell.width", true)?;
    }
    if cell
        .get("colSpan")
        .is_some_and(|value| value.as_u64().is_none_or(|span| span == 0))
    {
        return Err(validation(
            path,
            "style cell `colSpan` must be a positive integer",
        ));
    }
    validate_map_enum(cell, path, "verticalAlign", &["top", "center", "bottom"])?;
    validate_map_enum(cell, path, "verticalMerge", &["restart", "continue"])?;
    validate_map_enum(
        cell,
        path,
        "textDirection",
        &[
            "lr", "lrV", "rl", "rlV", "tb", "tbV", "tbRlV", "tbRl", "btLr", "lrTbV",
        ],
    )?;
    if let Some(value) = cell.get("shading") {
        validate_rgb(value, path, "style cell `shading`")?;
    }
    if let Some(value) = cell.get("margins") {
        validate_box_margins(value, path)?;
    }
    if let Some(value) = cell.get("border") {
        validate_border_model(value, path, true)?;
    }
    Ok(())
}

fn validate_style_border(value: &Value, path: &str, allow_space: bool) -> Result<()> {
    let border = value
        .as_object()
        .ok_or_else(|| validation(path, "style border must be an object"))?;
    let allowed = if allow_space {
        &["style", "size", "color", "space"][..]
    } else {
        &["style", "size", "color"][..]
    };
    validate_object_keys(border, path, allowed)?;
    validate_map_enum(border, path, "style", BORDER_STYLES)?;
    if let Some(value) = border.get("size") {
        require_number(value, path, "border.size", true)?;
    }
    if let Some(value) = border.get("color") {
        validate_rgb(value, path, "style border `color`")?;
    }
    if border
        .get("space")
        .is_some_and(|value| value.as_u64().is_none())
    {
        return Err(validation(
            path,
            "border `space` must be a non-negative integer",
        ));
    }
    Ok(())
}

fn validate_border_model(value: &Value, path: &str, cell: bool) -> Result<()> {
    let mut positions = vec![
        "top",
        "right",
        "bottom",
        "left",
        "insideHorizontal",
        "insideVertical",
    ];
    if cell {
        positions.extend(["topLeftToBottomRight", "topRightToBottomLeft"]);
    }
    validate_positioned_border_model(value, path, &positions, false)
}

fn validate_paragraph_border_model(value: &Value, path: &str) -> Result<()> {
    validate_positioned_border_model(
        value,
        path,
        &["top", "right", "bottom", "left", "between", "bar"],
        true,
    )
}

fn validate_positioned_border_model(
    value: &Value,
    path: &str,
    positions: &[&str],
    allow_space: bool,
) -> Result<()> {
    let border = value
        .as_object()
        .ok_or_else(|| validation(path, "`border` must be an object"))?;
    let mut uniform = vec!["style", "size", "color"];
    if allow_space {
        uniform.push("space");
    }
    let uses_uniform = border.keys().any(|key| uniform.contains(&key.as_str()));
    let uses_advanced = border
        .keys()
        .any(|key| key == "clearAll" || positions.contains(&key.as_str()));
    if uses_uniform && uses_advanced {
        return Err(validation(
            path,
            "uniform and positioned border properties are mutually exclusive",
        ));
    }
    if !uses_advanced {
        return validate_style_border(value, path, allow_space);
    }
    let mut allowed = positions.to_vec();
    allowed.push("clearAll");
    validate_object_keys(border, path, &allowed)?;
    if border
        .get("clearAll")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(validation(path, "border `clearAll` must be a boolean"));
    }
    if border.get("clearAll").and_then(Value::as_bool) == Some(true)
        && positions.iter().any(|key| border.contains_key(*key))
    {
        return Err(validation(
            path,
            "border `clearAll` cannot be combined with position entries",
        ));
    }
    if border.get("clearAll").and_then(Value::as_bool) != Some(true)
        && !positions.iter().any(|key| border.contains_key(*key))
    {
        return Err(validation(
            path,
            "advanced `border` must configure a position",
        ));
    }
    for position in positions {
        let Some(value) = border.get(*position) else {
            continue;
        };
        if value == &Value::Bool(false) {
            continue;
        }
        if value.is_boolean() {
            return Err(validation(
                path,
                format!("border `{position}` must be an object or false"),
            ));
        }
        validate_style_border(value, path, allow_space)?;
    }
    Ok(())
}

fn validate_style_paragraph(value: &Value, path: &str) -> Result<()> {
    let paragraph = value
        .as_object()
        .ok_or_else(|| validation(path, "style `paragraph` must be an object"))?;
    validate_object_keys(
        paragraph,
        path,
        &[
            "align",
            "textAlign",
            "snapToGrid",
            "suppressLineNumbers",
            "pageBreakBefore",
            "widowControl",
            "bidi",
            "shading",
            "adjustRightIndent",
            "keepNext",
            "keepLines",
            "spacingBefore",
            "spacingAfter",
            "lineSpacing",
            "spacingBeforeLines",
            "spacingAfterLines",
            "lineRule",
            "indentLeft",
            "indentRight",
            "indentLeftChars",
            "indentStartChars",
            "indentRightChars",
            "indentEndChars",
            "firstLine",
            "hanging",
            "hangingChars",
            "firstLineChars",
            "outlineLevel",
            "frame",
        ],
    )?;
    validate_character_indent_props(paragraph, path, "style paragraph ")?;
    validate_map_enum(
        paragraph,
        path,
        "align",
        &[
            "left",
            "right",
            "center",
            "both",
            "distribute",
            "start",
            "end",
            "justified",
        ],
    )?;
    validate_map_enum(
        paragraph,
        path,
        "textAlign",
        &["auto", "baseline", "bottom", "center", "top"],
    )?;
    validate_style_paragraph_layout(paragraph, path)?;
    if paragraph.contains_key("indentLeftChars") && paragraph.contains_key("indentStartChars")
        || paragraph.contains_key("indentRightChars") && paragraph.contains_key("indentEndChars")
    {
        return Err(validation(
            path,
            "style paragraph character indentation aliases are mutually exclusive; choose left or start and right or end",
        ));
    }
    if paragraph
        .get("outlineLevel")
        .is_some_and(|value| value.as_u64().is_none_or(|level| level > 9))
    {
        return Err(validation(
            path,
            "style paragraph `outlineLevel` must be an integer from 0 to 9",
        ));
    }
    if let Some(frame) = paragraph.get("frame") {
        validate_style_frame(frame, path)?;
    }
    Ok(())
}

fn validate_character_indent_props(
    props: &Map<String, Value>,
    path: &str,
    owner: &str,
) -> Result<()> {
    for key in [
        "indentLeftChars",
        "indentStartChars",
        "indentRightChars",
        "indentEndChars",
        "firstLineChars",
        "hangingChars",
    ] {
        if props.get(key).is_some_and(|value| value.as_i64().is_none()) {
            return Err(validation(
                path,
                format!("{owner}`{key}` must be an integer"),
            ));
        }
    }
    if props.contains_key("indentLeftChars") && props.contains_key("indentStartChars")
        || props.contains_key("indentRightChars") && props.contains_key("indentEndChars")
    {
        return Err(validation(
            path,
            format!(
                "{owner}character indentation aliases are mutually exclusive; choose left or start and right or end"
            ),
        ));
    }
    if (props.contains_key("hanging") || props.contains_key("hangingChars"))
        && (props.contains_key("firstLine") || props.contains_key("firstLineChars"))
        || props.contains_key("hanging") && props.contains_key("hangingChars")
        || props.contains_key("firstLine") && props.contains_key("firstLineChars")
    {
        return Err(validation(
            path,
            format!(
                "{owner}`hanging`/`hangingChars` and `firstLine`/`firstLineChars` are mutually exclusive; choose only one indentation mode"
            ),
        ));
    }
    Ok(())
}

fn validate_style_paragraph_layout(paragraph: &Map<String, Value>, path: &str) -> Result<()> {
    if paragraph
        .get("snapToGrid")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(validation(
            path,
            "style paragraph `snapToGrid` must be a boolean",
        ));
    }
    for key in [
        "keepNext",
        "keepLines",
        "suppressLineNumbers",
        "pageBreakBefore",
        "widowControl",
        "bidi",
    ] {
        if paragraph.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(
                path,
                format!("style paragraph `{key}` must be a boolean"),
            ));
        }
    }
    if paragraph
        .get("adjustRightIndent")
        .is_some_and(|value| !matches!(value.as_i64(), Some(0 | 1)))
    {
        return Err(validation(
            path,
            "style paragraph `adjustRightIndent` must be 0 or 1 because OOXML defines w:adjustRightInd as on/off",
        ));
    }
    if let Some(value) = paragraph.get("shading") {
        validate_rgb(value, path, "style paragraph `shading`")?;
    }
    for key in ["indentLeft", "indentRight"] {
        if let Some(value) = paragraph.get(key) {
            value
                .as_f64()
                .filter(|number| number.is_finite())
                .ok_or_else(|| validation(path, format!("`{key}` must be a finite number")))?;
        }
    }
    for key in [
        "spacingBefore",
        "spacingAfter",
        "lineSpacing",
        "firstLine",
        "hanging",
    ] {
        if let Some(value) = paragraph.get(key) {
            require_number(value, path, key, false)?;
        }
    }
    validate_line_spacing_extras(
        paragraph,
        path,
        "spacingBeforeLines",
        "spacingAfterLines",
        "lineRule",
    )?;
    if paragraph.contains_key("firstLine") && paragraph.contains_key("hanging") {
        return Err(validation(
            path,
            "style paragraph `firstLine` and `hanging` are mutually exclusive",
        ));
    }
    Ok(())
}

fn validate_style_frame(value: &Value, path: &str) -> Result<()> {
    let frame = value
        .as_object()
        .ok_or_else(|| validation(path, "`frame` must be an object"))?;
    validate_object_keys(
        frame,
        path,
        &[
            "wrap",
            "verticalAnchor",
            "horizontalAnchor",
            "heightRule",
            "xAlign",
            "yAlign",
            "horizontalSpace",
            "verticalSpace",
            "x",
            "y",
            "width",
            "height",
        ],
    )?;
    if frame.is_empty() {
        return Err(validation(path, "`frame` must not be empty"));
    }
    validate_map_enum(
        frame,
        path,
        "wrap",
        &["around", "auto", "none", "notBeside", "through", "tight"],
    )?;
    for key in ["verticalAnchor", "horizontalAnchor"] {
        validate_map_enum(frame, path, key, &["margin", "page", "text"])?;
    }
    validate_map_enum(frame, path, "heightRule", &["atLeast", "auto", "exact"])?;
    validate_map_enum(
        frame,
        path,
        "xAlign",
        &["center", "inside", "left", "outside", "right"],
    )?;
    validate_map_enum(
        frame,
        path,
        "yAlign",
        &["bottom", "center", "inline", "inside", "outside", "top"],
    )?;
    for key in ["horizontalSpace", "verticalSpace", "x", "y"] {
        if let Some(value) = frame.get(key) {
            value
                .as_f64()
                .filter(|number| number.is_finite())
                .ok_or_else(|| validation(path, format!("frame `{key}` must be finite")))?;
        }
    }
    for key in ["width", "height"] {
        if let Some(value) = frame.get(key) {
            require_number(value, path, &format!("frame.{key}"), true)?;
        }
    }
    for (coordinate, alignment) in [("x", "xAlign"), ("y", "yAlign")] {
        if frame.contains_key(coordinate) && frame.contains_key(alignment) {
            return Err(validation(
                path,
                format!("frame `{coordinate}` and `{alignment}` are mutually exclusive"),
            ));
        }
    }
    Ok(())
}

fn validate_object_keys(object: &Map<String, Value>, path: &str, allowed: &[&str]) -> Result<()> {
    if let Some(key) = object.keys().find(|key| !allowed.contains(&key.as_str())) {
        return Err(validation(path, format!("unknown property `{key}`")));
    }
    Ok(())
}

fn validate_map_enum(
    object: &Map<String, Value>,
    path: &str,
    key: &str,
    allowed: &[&str],
) -> Result<()> {
    let Some(value) = object.get(key) else {
        return Ok(());
    };
    let value = value
        .as_str()
        .ok_or_else(|| validation(path, format!("`{key}` must be a string")))?;
    if !allowed.contains(&value) {
        return Err(validation(
            path,
            format!(
                "invalid `{key}` value `{value}`; choose one of: {}",
                allowed.join(", ")
            ),
        ));
    }
    Ok(())
}

fn validate_rgb(value: &Value, path: &str, label: &str) -> Result<()> {
    let color = value
        .as_str()
        .ok_or_else(|| validation(path, format!("{label} must be a color string")))?;
    if color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(validation(path, format!("{label} must be six-digit RGB")));
    }
    Ok(())
}

fn validate_theme_fields(object: &Map<String, Value>, path: &str) -> Result<()> {
    validate_map_enum(
        object,
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
    for key in ["themeShade", "themeTint"] {
        if let Some(value) = object.get(key) {
            let modifier = value
                .as_str()
                .ok_or_else(|| validation(path, format!("`{key}` must be a hex byte")))?;
            if modifier.len() != 2 || !modifier.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return Err(validation(path, format!("`{key}` must be a hex byte")));
            }
        }
    }
    Ok(())
}

fn validate_string_map(value: &Value, path: &str, key: &str) -> Result<()> {
    let entries = value
        .as_object()
        .ok_or_else(|| validation(path, format!("`{key}` must be an object")))?;
    if entries.is_empty() {
        return Err(validation(path, format!("`{key}` must not be empty")));
    }
    if entries.keys().any(String::is_empty) {
        return Err(validation(path, format!("`{key}` names must be non-empty")));
    }
    if entries.values().any(|value| !value.is_string()) {
        return Err(validation(
            path,
            format!("`{key}` must contain string values"),
        ));
    }
    Ok(())
}

fn validate_section_semantics(node: &Node, path: &str) -> Result<()> {
    if node.kind != NodeKind::Section {
        return Ok(());
    }
    if node
        .props
        .get("titlePage")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(validation(path, "`titlePage` must be a boolean"));
    }
    validate_optional_enum_prop(
        node,
        path,
        "textDirection",
        &["lrTb", "tbRl", "btLr", "lrTbV", "tbRlV"],
    )?;
    validate_optional_enum_prop(
        node,
        path,
        "type",
        &[
            "nextPage",
            "nextColumn",
            "continuous",
            "evenPage",
            "oddPage",
        ],
    )?;
    if let Some(value) = node.props.get("documentGrid") {
        validate_document_grid(value, path)?;
    }
    if let Some(value) = node.props.get("pageNumbering") {
        validate_page_numbering(value, path)?;
    }
    if let Some(value) = node.props.get("lineNumbering") {
        validate_line_numbering(value, path)?;
    }
    if let Some(value) = node.props.get("columns") {
        validate_section_columns(value, path)?;
    }
    Ok(())
}

fn validate_section_columns(value: &Value, path: &str) -> Result<()> {
    if let Some(columns) = value.as_u64() {
        if (1..=45).contains(&columns) {
            return Ok(());
        }
        return Err(validation(
            path,
            "`columns` must be an integer from 1 to 45",
        ));
    }
    let columns = value.as_object().ok_or_else(|| {
        validation(
            path,
            "`columns` must be an integer from 1 to 45 or an object with `count` or `widths`",
        )
    })?;
    validate_object_keys(
        columns,
        &format!("{path}.columns"),
        &["count", "spacing", "separator", "widths"],
    )?;
    if columns
        .get("separator")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(validation(path, "`columns.separator` must be a boolean"));
    }
    match (columns.get("count"), columns.get("widths")) {
        (Some(_), Some(_)) => Err(validation(
            path,
            "`columns.count` and `columns.widths` are mutually exclusive; choose equal-width count or individual widths",
        )),
        (Some(count), None) => {
            if count
                .as_u64()
                .is_none_or(|count| !(1..=45).contains(&count))
            {
                return Err(validation(
                    path,
                    "`columns.count` must be an integer from 1 to 45",
                ));
            }
            if let Some(spacing) = columns.get("spacing") {
                validate_non_negative_column_measure(spacing, path, "columns.spacing")?;
            }
            Ok(())
        }
        (None, Some(widths)) => {
            if columns.contains_key("spacing") {
                return Err(validation(
                    path,
                    "`columns.spacing` cannot be combined with `columns.widths`; set `spacing` on each width entry",
                ));
            }
            let widths = widths.as_array().ok_or_else(|| {
                validation(path, "`columns.widths` must be an array of column objects")
            })?;
            if widths.is_empty() || widths.len() > 45 {
                return Err(validation(
                    path,
                    "`columns.widths` must contain between 1 and 45 column objects",
                ));
            }
            for (index, column) in widths.iter().enumerate() {
                let column = column.as_object().ok_or_else(|| {
                    validation(path, format!("`columns.widths[{index}]` must be an object"))
                })?;
                let column_path = format!("{path}.columns.widths[{index}]");
                validate_object_keys(column, &column_path, &["width", "spacing"])?;
                for key in ["width", "spacing"] {
                    if let Some(value) = column.get(key) {
                        validate_non_negative_column_measure(value, &column_path, key)?;
                    }
                }
            }
            Ok(())
        }
        (None, None) => Err(validation(
            path,
            "`columns` object requires either `count` or `widths`",
        )),
    }
}

fn validate_non_negative_column_measure(value: &Value, path: &str, key: &str) -> Result<()> {
    if value
        .as_f64()
        .is_some_and(|number| number.is_finite() && number >= 0.0)
    {
        Ok(())
    } else {
        Err(validation(
            path,
            format!("`{key}` must be a finite non-negative point measure"),
        ))
    }
}

fn validate_document_grid(value: &Value, path: &str) -> Result<()> {
    let grid = value
        .as_object()
        .ok_or_else(|| validation(path, "`documentGrid` must be an object"))?;
    if grid.is_empty() {
        return Err(validation(path, "`documentGrid` must not be empty"));
    }
    for key in grid.keys() {
        if !["type", "linePitch", "charSpace"].contains(&key.as_str()) {
            return Err(validation(
                path,
                format!("unknown documentGrid property `{key}`"),
            ));
        }
    }
    let grid_type = grid
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| validation(path, "`documentGrid.type` must be a string"))?;
    if !["default", "lines", "linesAndChars", "snapToChars"].contains(&grid_type) {
        return Err(validation(
            path,
            format!("invalid `documentGrid.type` value `{grid_type}`"),
        ));
    }
    if let Some(value) = grid.get("linePitch") {
        require_number(value, path, "documentGrid.linePitch", true)?;
    }
    if grid
        .get("charSpace")
        .is_some_and(|value| value.as_i64().is_none())
    {
        return Err(validation(
            path,
            "`documentGrid.charSpace` must be a signed integer",
        ));
    }
    Ok(())
}

fn validate_page_numbering(value: &Value, path: &str) -> Result<()> {
    let numbering = value
        .as_object()
        .ok_or_else(|| validation(path, "`pageNumbering` must be an object"))?;
    if numbering.is_empty() {
        return Err(validation(path, "`pageNumbering` must not be empty"));
    }
    for key in numbering.keys() {
        if !["start", "chapterStyle", "format", "chapterSeparator"].contains(&key.as_str()) {
            return Err(validation(
                path,
                format!("unknown pageNumbering property `{key}`"),
            ));
        }
    }
    if numbering
        .get("start")
        .is_some_and(|value| value.as_u64().is_none())
    {
        return Err(validation(
            path,
            "`pageNumbering.start` must be a non-negative integer",
        ));
    }
    if numbering
        .get("chapterStyle")
        .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
    {
        return Err(validation(
            path,
            "`pageNumbering.chapterStyle` must be a non-empty string",
        ));
    }
    if numbering.get("format").is_some_and(|value| {
        value
            .as_str()
            .is_none_or(|format| !PAGE_NUMBER_FORMATS.contains(&format))
    }) {
        return Err(validation(
            path,
            format!(
                "`pageNumbering.format` must be one of: {}",
                PAGE_NUMBER_FORMATS.join(", ")
            ),
        ));
    }
    if numbering.get("chapterSeparator").is_some_and(|value| {
        value
            .as_str()
            .is_none_or(|separator| !PAGE_NUMBER_CHAPTER_SEPARATORS.contains(&separator))
    }) {
        return Err(validation(
            path,
            "`pageNumbering.chapterSeparator` must be hyphen, period, colon, emDash, or enDash",
        ));
    }
    Ok(())
}

fn validate_line_numbering(value: &Value, path: &str) -> Result<()> {
    let numbering = value
        .as_object()
        .ok_or_else(|| validation(path, "`lineNumbering` must be an object"))?;
    for key in numbering.keys() {
        if !["countBy", "start", "distance", "restart"].contains(&key.as_str()) {
            return Err(validation(
                path,
                format!("unknown lineNumbering property `{key}`"),
            ));
        }
    }
    for key in ["countBy", "start"] {
        if numbering
            .get(key)
            .is_some_and(|value| value.as_u64().is_none_or(|value| value == 0))
        {
            return Err(validation(
                path,
                format!("`lineNumbering.{key}` must be a positive integer"),
            ));
        }
    }
    if let Some(value) = numbering.get("distance") {
        require_number(value, path, "lineNumbering.distance", false)?;
    }
    if let Some(restart) = numbering.get("restart") {
        let restart = restart
            .as_str()
            .ok_or_else(|| validation(path, "`lineNumbering.restart` must be a string"))?;
        if !["newPage", "newSection", "continuous"].contains(&restart) {
            return Err(validation(
                path,
                format!(
                    "invalid `lineNumbering.restart` value `{restart}`; use newPage, newSection, or continuous"
                ),
            ));
        }
    }
    Ok(())
}

fn validate_table_semantics(node: &Node, path: &str) -> Result<()> {
    if node.kind == NodeKind::Table {
        if node
            .props
            .get("style")
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(path, "Table `style` must be a non-empty string"));
        }
        if let Some(value) = node.props.get("indent") {
            value
                .as_f64()
                .filter(|number| number.is_finite())
                .ok_or_else(|| validation(path, "Table `indent` must be a finite number"))?;
        }
        if let Some(value) = node.props.get("position") {
            validate_table_position(value, path)?;
        }
        if let Some(value) = node.props.get("border") {
            validate_border_model(value, path, false)?;
        }
    }
    if matches!(node.kind, NodeKind::Table | NodeKind::TableCell)
        && let Some(value) = node.props.get("margins")
    {
        validate_box_margins(value, path)?;
    }
    if node.kind == NodeKind::TableCell {
        for (key, allowed) in [
            ("verticalMerge", &["restart", "continue"][..]),
            (
                "textDirection",
                &[
                    "lr", "lrV", "rl", "rlV", "tb", "tbV", "tbRlV", "tbRl", "btLr", "lrTbV",
                ][..],
            ),
        ] {
            if node.props.contains_key(key) {
                validate_optional_enum_prop(node, path, key, allowed)?;
            }
        }
        if let Some(value) = node.props.get("border") {
            validate_border_model(value, path, true)?;
        }
    }
    if node.kind == NodeKind::TableRow {
        for key in ["gridBefore", "gridAfter"] {
            if let Some(value) = node.props.get(key)
                && value.as_u64().is_none()
            {
                return Err(validation(
                    path,
                    format!("TableRow `{key}` must be a non-negative integer"),
                ));
            }
        }
        for key in ["widthBefore", "widthAfter"] {
            if let Some(value) = node.props.get(key) {
                require_number(value, path, key, false)?;
            }
        }
        if node.props.contains_key("inserted") && node.props.contains_key("deleted") {
            return Err(validation(
                path,
                "TableRow `inserted` and `deleted` are mutually exclusive",
            ));
        }
        for key in ["inserted", "deleted"] {
            if let Some(value) = node.props.get(key) {
                validate_row_revision(value, path, key)?;
            }
        }
    }
    Ok(())
}

fn validate_table_position(value: &Value, path: &str) -> Result<()> {
    let position = value
        .as_object()
        .ok_or_else(|| validation(path, "Table `position` must be an object"))?;
    let allowed = [
        "leftFromText",
        "rightFromText",
        "verticalAnchor",
        "horizontalAnchor",
        "xAlign",
        "yAlign",
        "x",
        "y",
    ];
    if position.is_empty() {
        return Err(validation(path, "Table `position` must not be empty"));
    }
    for key in position.keys() {
        if !allowed.contains(&key.as_str()) {
            return Err(validation(
                path,
                format!("unknown position property `{key}`"),
            ));
        }
    }
    for key in ["leftFromText", "rightFromText"] {
        if let Some(value) = position.get(key) {
            require_number(value, path, &format!("position.{key}"), false)?;
        }
    }
    for key in ["x", "y"] {
        if let Some(value) = position.get(key) {
            value
                .as_f64()
                .filter(|number| number.is_finite())
                .ok_or_else(|| validation(path, format!("position.{key} must be finite")))?;
        }
    }
    validate_position_enum(
        position,
        path,
        "verticalAnchor",
        &["margin", "page", "text"],
    )?;
    validate_position_enum(
        position,
        path,
        "horizontalAnchor",
        &["margin", "page", "text"],
    )?;
    validate_position_enum(
        position,
        path,
        "xAlign",
        &["center", "inside", "left", "outside", "right"],
    )?;
    validate_position_enum(
        position,
        path,
        "yAlign",
        &["bottom", "center", "inline", "inside", "outside", "top"],
    )?;
    for (coordinate, alignment) in [("x", "xAlign"), ("y", "yAlign")] {
        if position.contains_key(coordinate) && position.contains_key(alignment) {
            return Err(validation(
                path,
                format!("position `{coordinate}` and `{alignment}` are mutually exclusive"),
            ));
        }
    }
    Ok(())
}

fn validate_position_enum(
    position: &Map<String, Value>,
    path: &str,
    key: &str,
    allowed: &[&str],
) -> Result<()> {
    let Some(value) = position.get(key) else {
        return Ok(());
    };
    let value = value
        .as_str()
        .ok_or_else(|| validation(path, format!("position `{key}` must be a string")))?;
    if !allowed.contains(&value) {
        return Err(validation(
            path,
            format!("invalid position `{key}` value `{value}`"),
        ));
    }
    Ok(())
}

fn validate_row_revision(value: &Value, path: &str, key: &str) -> Result<()> {
    let revision = value
        .as_object()
        .ok_or_else(|| validation(path, format!("TableRow `{key}` must be an object")))?;
    for field in revision.keys() {
        if !["author", "date"].contains(&field.as_str()) {
            return Err(validation(
                path,
                format!("unknown {key} property `{field}`"),
            ));
        }
    }
    for field in ["author", "date"] {
        if revision.get(field).is_some_and(|value| !value.is_string()) {
            return Err(validation(
                path,
                format!("TableRow `{key}.{field}` must be a string"),
            ));
        }
    }
    Ok(())
}

fn validate_box_margins(value: &Value, path: &str) -> Result<()> {
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
    for key in ["top", "right", "bottom", "left"] {
        let value = margins
            .get(key)
            .ok_or_else(|| validation(path, format!("`margins` requires `{key}`")))?;
        require_number(value, path, &format!("margins.{key}"), false)?;
    }
    Ok(())
}

fn validate_run_defaults(node: &Node, path: &str) -> Result<()> {
    if node.kind != NodeKind::Run {
        return Ok(());
    }
    validate_font_choice(&node.props, path, "font", "fonts")?;
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

fn validate_font_choice(
    props: &Map<String, Value>,
    path: &str,
    shorthand_key: &str,
    slots_key: &str,
) -> Result<()> {
    if props.contains_key(shorthand_key) && props.contains_key(slots_key) {
        return Err(validation(
            path,
            format!(
                "`{shorthand_key}` and `{slots_key}` cannot be used together; use `{slots_key}` when font slots differ"
            ),
        ));
    }
    if props
        .get(shorthand_key)
        .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
    {
        return Err(validation(
            path,
            format!("`{shorthand_key}` must be a non-empty string"),
        ));
    }
    let Some(value) = props.get(slots_key) else {
        return Ok(());
    };
    let fonts = value
        .as_object()
        .ok_or_else(|| validation(path, format!("`{slots_key}` must be an object")))?;
    validate_object_keys(
        fonts,
        path,
        &[
            "ascii",
            "hiAnsi",
            "eastAsia",
            "cs",
            "asciiTheme",
            "hiAnsiTheme",
            "eastAsiaTheme",
            "csTheme",
            "hint",
        ],
    )?;
    if fonts.is_empty() {
        return Err(validation(path, format!("`{slots_key}` must not be empty")));
    }
    for (key, value) in fonts {
        if value.as_str().is_none_or(str::is_empty) {
            return Err(validation(
                path,
                format!("`{slots_key}.{key}` must be a non-empty string"),
            ));
        }
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

pub(crate) const SHADING_PATTERNS: &[&str] = &[
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
        if has_binding
            && node
                .props
                .get("storeItemId")
                .and_then(Value::as_str)
                .is_none_or(str::is_empty)
        {
            return Err(validation(
                path,
                "ContentControl data binding requires non-empty `storeItemId` (the custom XML data-store item GUID)",
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

fn validate_index_instruction_options(node: &Node, path: &str) -> Result<()> {
    for key in ["tcFieldLevelRange", "omitPageNumbersLevelRange"] {
        if let Some(range) = node.props.get(key) {
            validate_index_level_range(range, path, key)?;
        }
    }
    for key in [
        "preserveTab",
        "preserveNewLine",
        "useAppliedParagraphLineLevel",
        "hideTabAndPageNumbersInWebview",
    ] {
        if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(path, format!("`{key}` must be a boolean")));
        }
    }
    if matches!(
        node.kind,
        NodeKind::TableOfContents | NodeKind::TableOfFigures | NodeKind::TableOfEntries
    ) {
        for key in ["hyperlinks", "dirty"] {
            if node.props.get(key).is_some_and(|value| !value.is_boolean()) {
                return Err(validation(path, format!("`{key}` must be a boolean")));
            }
        }
        if let Some(alias) = node.props.get("alias")
            && alias.as_str().is_none_or(str::is_empty)
        {
            return Err(validation(path, "`alias` must be a non-empty string"));
        }
    }
    for key in ["entryBookmarkName", "seqFieldIdentifierForPrefix"] {
        if node
            .props
            .get(key)
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(
                path,
                format!("`{key}` must be a non-empty string"),
            ));
        }
        if node
            .props
            .get(key)
            .and_then(Value::as_str)
            .is_some_and(|value| value.contains('"'))
        {
            return Err(validation(
                path,
                format!("`{key}` cannot contain a quote because OOXML field arguments are quoted"),
            ));
        }
    }
    for key in [
        "entryAndPageNumberSeparator",
        "sequenceAndPageNumbersSeparator",
    ] {
        if node.props.get(key).is_some_and(|value| !value.is_string()) {
            return Err(validation(path, format!("`{key}` must be a string")));
        }
        if node
            .props
            .get(key)
            .and_then(Value::as_str)
            .is_some_and(|value| value.contains('"'))
        {
            return Err(validation(
                path,
                format!("`{key}` cannot contain a quote because OOXML field arguments are quoted"),
            ));
        }
    }
    for key in ["separator", "label", "identifier"] {
        if node
            .props
            .get(key)
            .and_then(Value::as_str)
            .is_some_and(|value| value.contains('"'))
        {
            return Err(validation(
                path,
                format!("`{key}` cannot contain a quote because OOXML field arguments are quoted"),
            ));
        }
    }
    if node.props.contains_key("separator")
        && node.props.contains_key("sequenceAndPageNumbersSeparator")
    {
        return Err(validation(
            path,
            "`separator` and `sequenceAndPageNumbersSeparator` both set the same TOC switch; keep only one",
        ));
    }
    if let Some(styles) = node.props.get("stylesWithLevels") {
        validate_styles_with_levels(styles, path)?;
    }
    Ok(())
}

fn validate_styles_with_levels(value: &Value, path: &str) -> Result<()> {
    let styles = value
        .as_array()
        .filter(|styles| !styles.is_empty())
        .ok_or_else(|| validation(path, "`stylesWithLevels` must be a non-empty array"))?;
    for (index, style) in styles.iter().enumerate() {
        let style_path = format!("{path}.stylesWithLevels[{index}]");
        let style = style
            .as_object()
            .ok_or_else(|| validation(&style_path, "must be an object"))?;
        validate_object_keys(style, &style_path, &["style", "level"])?;
        let valid_name = style
            .get("style")
            .and_then(Value::as_str)
            .is_some_and(|name| {
                !name.is_empty() && !name.chars().any(|c| matches!(c, ',' | '"' | '\\'))
            });
        if !valid_name {
            return Err(validation(
                &style_path,
                "`style` must be non-empty and cannot contain commas, quotes, or backslashes",
            ));
        }
        if style
            .get("level")
            .and_then(Value::as_u64)
            .is_none_or(|level| !(1..=9).contains(&level))
        {
            return Err(validation(
                &style_path,
                "`level` must be an integer from 1 to 9",
            ));
        }
    }
    Ok(())
}

fn validate_index_semantics(node: &Node, path: &str) -> Result<()> {
    validate_index_instruction_options(node, path)?;
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

fn validate_index_level_range(value: &Value, path: &str, key: &str) -> Result<()> {
    let range_path = format!("{path}.{key}");
    let range = value
        .as_object()
        .ok_or_else(|| validation(&range_path, "must be an object with `start` and `end`"))?;
    validate_object_keys(range, &range_path, &["start", "end"])?;
    let start = range
        .get("start")
        .and_then(Value::as_u64)
        .filter(|value| (1..=9).contains(value))
        .ok_or_else(|| validation(&range_path, "`start` must be an integer from 1 to 9"))?;
    let end = range
        .get("end")
        .and_then(Value::as_u64)
        .filter(|value| (1..=9).contains(value))
        .ok_or_else(|| validation(&range_path, "`end` must be an integer from 1 to 9"))?;
    if start > end {
        return Err(validation(
            &range_path,
            "`start` must be less than or equal to `end`",
        ));
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
    if matches!(node.kind, NodeKind::Header | NodeKind::Footer) {
        validate_optional_enum_prop(node, path, "type", &["default", "first", "even"])?;
    }
    if node.kind == NodeKind::Heading {
        let level = node.props.get("level").and_then(Value::as_u64);
        if level.is_none_or(|level| !(1..=9).contains(&level)) {
            return Err(validation(
                path,
                "Heading requires integer `level` from 1 to 9",
            ));
        }
    }
    if matches!(node.kind, NodeKind::Bookmark | NodeKind::InlineBookmark) {
        let name = node.props.get("name").and_then(Value::as_str);
        if name.is_none_or(str::is_empty) {
            return Err(validation(
                path,
                format!("{} requires non-empty `name`", node.kind.name()),
            ));
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
    validate_list_item_props(node, path)?;
    if let Some(value) = node.props.get("start")
        && value.as_u64().is_none_or(|start| start == 0)
    {
        return Err(validation(path, "`start` must be a positive integer"));
    }
    Ok(())
}

fn validate_list_item_props(node: &Node, path: &str) -> Result<()> {
    if node.kind != NodeKind::ListItem {
        return Ok(());
    }
    if node
        .props
        .get("suppressLineNumbers")
        .is_some_and(|value| !value.is_boolean())
    {
        return Err(validation(path, "`suppressLineNumbers` must be a boolean"));
    }
    if let Some(value) = node.props.get("level")
        && value.as_u64().is_none_or(|level| level > 8)
    {
        return Err(validation(path, "`level` must be an integer from 0 to 8"));
    }
    if node
        .props
        .get("style")
        .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
    {
        return Err(validation(
            path,
            "ListItem `style` must be a non-empty string",
        ));
    }
    for key in [
        "indentLeftChars",
        "indentStartChars",
        "indentRightChars",
        "indentEndChars",
        "firstLineChars",
        "hangingChars",
    ] {
        if node
            .props
            .get(key)
            .is_some_and(|value| value.as_i64().is_none())
        {
            return Err(validation(
                path,
                format!("ListItem `{key}` must be an integer"),
            ));
        }
    }
    if node.props.contains_key("indentLeftChars") && node.props.contains_key("indentStartChars")
        || node.props.contains_key("indentRightChars") && node.props.contains_key("indentEndChars")
        || (node.props.contains_key("hangingChars") && node.props.contains_key("firstLineChars"))
    {
        return Err(validation(
            path,
            "ListItem character indentation values are mutually exclusive by axis; choose one value for each axis",
        ));
    }
    Ok(())
}

fn validate_list_levels(node: &Node, path: &str) -> Result<()> {
    if let Some(value) = node.props.get("levels") {
        let levels = value
            .as_array()
            .ok_or_else(|| validation(path, "`levels` must be an array"))?;
        if levels.is_empty() {
            return Err(validation(path, "`levels` must not be empty"));
        }
        if levels.len() > 9 {
            return Err(validation(path, "`levels` must contain at most 9 entries"));
        }
        for (index, level) in levels.iter().enumerate() {
            validate_list_level(level, &format!("{path}/levels[{index}]"))?;
        }
    }
    validate_list_overrides(node, path)
}

fn validate_list_overrides(node: &Node, path: &str) -> Result<()> {
    let Some(value) = node.props.get("overrides") else {
        return Ok(());
    };
    let overrides = value
        .as_array()
        .ok_or_else(|| validation(path, "`overrides` must be an array"))?;
    let mut levels = std::collections::BTreeSet::new();
    for (index, value) in overrides.iter().enumerate() {
        let item_path = format!("{path}/overrides[{index}]");
        let item = value
            .as_object()
            .ok_or_else(|| validation(&item_path, "list override must be an object"))?;
        validate_object_keys(item, &item_path, &["level", "start", "definition"])?;
        let level = item
            .get("level")
            .and_then(Value::as_u64)
            .filter(|level| *level <= 8)
            .ok_or_else(|| {
                validation(
                    &item_path,
                    "override `level` must be an integer from 0 to 8",
                )
            })?;
        if !levels.insert(level) {
            return Err(validation(
                &item_path,
                format!("duplicate list override for level {level}"),
            ));
        }
        if !item.contains_key("start") && !item.contains_key("definition") {
            return Err(validation(
                &item_path,
                "list override requires `start` or `definition`",
            ));
        }
        if let Some(start) = item.get("start")
            && start.as_u64().is_none_or(|start| start == 0)
        {
            return Err(validation(
                &item_path,
                "override `start` must be a positive integer",
            ));
        }
        if let Some(definition) = item.get("definition") {
            validate_list_level(definition, &format!("{item_path}/definition"))?;
        }
    }
    Ok(())
}

fn validate_list_level(value: &Value, path: &str) -> Result<()> {
    let level = value
        .as_object()
        .ok_or_else(|| validation(path, "list level must be an object"))?;
    validate_object_keys(level, path, LIST_LEVEL_KEYS)?;
    validate_list_level_scalars(level, path)?;
    validate_font_choice(level, path, "font", "fonts")?;
    Ok(())
}

const LIST_LEVEL_KEYS: &[&str] = &[
    "format",
    "text",
    "start",
    "align",
    "suffix",
    "paragraphStyle",
    "restart",
    "legal",
    "indentLeft",
    "indentRight",
    "hanging",
    "firstLine",
    "font",
    "fonts",
    "size",
    "color",
    "highlight",
    "bold",
    "italic",
    "strike",
    "doubleStrike",
    "underline",
    "hidden",
    "characterSpacing",
];

fn validate_list_level_scalars(level: &Map<String, Value>, path: &str) -> Result<()> {
    for key in ["format", "text", "paragraphStyle", "highlight"] {
        if level
            .get(key)
            .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
        {
            return Err(validation(
                path,
                format!("list level `{key}` must be a non-empty string"),
            ));
        }
    }
    if let Some(value) = level.get("start")
        && value.as_u64().is_none_or(|start| start == 0)
    {
        return Err(validation(
            path,
            "list level `start` must be a positive integer",
        ));
    }
    if let Some(value) = level.get("restart")
        && value.as_u64().is_none()
    {
        return Err(validation(
            path,
            "list level `restart` must be a non-negative integer",
        ));
    }
    validate_map_enum(level, path, "align", &["left", "center", "right"])?;
    validate_map_enum(level, path, "suffix", &["tab", "space", "nothing"])?;
    for key in [
        "bold",
        "italic",
        "strike",
        "doubleStrike",
        "hidden",
        "legal",
    ] {
        if level.get(key).is_some_and(|value| !value.is_boolean()) {
            return Err(validation(
                path,
                format!("list level `{key}` must be a boolean"),
            ));
        }
    }
    if level.contains_key("firstLine") && level.contains_key("hanging") {
        return Err(validation(
            path,
            "list level `firstLine` and `hanging` are mutually exclusive",
        ));
    }
    for key in ["indentLeft", "indentRight", "hanging", "firstLine", "size"] {
        if let Some(value) = level.get(key) {
            require_number(
                value,
                path,
                key,
                key != "indentLeft" && key != "indentRight",
            )?;
        }
    }
    if let Some(value) = level.get("characterSpacing") {
        value
            .as_f64()
            .filter(|number| number.is_finite())
            .ok_or_else(|| validation(path, "`characterSpacing` must be a finite number"))?;
    }
    if let Some(value) = level.get("color") {
        validate_rgb(value, path, "list level `color`")?;
    }
    if let Some(value) = level.get("underline")
        && value.as_str().is_none_or(str::is_empty)
    {
        return Err(validation(
            path,
            "list level `underline` must be a non-empty string",
        ));
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
        return Err(validation(
            path,
            format!(
                "invalid `{key}` value `{value}`; choose one of: {}",
                allowed.join(", ")
            ),
        ));
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
        ir.validate().expect("fixture should validate");
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
    fn validate_should_accept_list_level_definitions() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{"type":"ordered","start":3,"levels":[{"format":"decimal","text":"%1.","start":3,"suffix":"space","legal":true,"restart":0,"indentLeft":36,"hanging":18,"bold":true,"size":12},{"format":"lowerLetter","text":"%2)","align":"right","paragraphStyle":"ListParagraph","italic":false}]},"children":[{"type":"ListItem","props":{"style":"ListParagraph"},"children":["one"]},{"type":"ListItem","props":{"level":1},"children":["two"]}]}]}]}}"#,
        );
        ir.validate().expect("custom list levels should validate");
    }

    #[test]
    fn validate_should_reject_invalid_list_level_definitions() {
        for (levels, expected) in [
            ("[]", "must not be empty"),
            (r#"[{"format":""}]"#, "format"),
            (r#"[{"firstLine":1,"hanging":1}]"#, "mutually exclusive"),
            (r#"[{"suffix":"gap"}]"#, "suffix"),
            (r#"[{"unknown":true}]"#, "unknown"),
        ] {
            let source = format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{}},"children":[{{"type":"Section","props":{{}},"children":[{{"type":"List","props":{{"levels":{levels}}},"children":[{{"type":"ListItem","props":{{}},"children":["x"]}}]}}]}}]}}}}"#
            );
            let error = parse(&source)
                .validate()
                .expect_err("invalid list levels must fail");
            assert!(
                error.to_string().contains(expected),
                "expected {expected} in {error}"
            );
        }
    }

    #[test]
    fn validate_should_accept_list_instance_overrides() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{"type":"ordered","overrides":[{"level":0,"start":4},{"level":1,"definition":{"format":"lowerLetter","text":"%2)"}}]},"children":[{"type":"ListItem","props":{},"children":["item"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_duplicate_or_empty_list_overrides() {
        for (overrides, expected) in [
            (r#"[{"level":0}]"#, "requires `start` or `definition`"),
            (
                r#"[{"level":0,"start":2},{"level":0,"start":3}]"#,
                "duplicate list override",
            ),
            (r#"[{"level":9,"start":1}]"#, "0 to 8"),
            (r#"[{"level":0,"start":0}]"#, "positive integer"),
        ] {
            let ir = parse(&format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{}},"children":[{{"type":"Section","props":{{}},"children":[{{"type":"List","props":{{"overrides":{overrides}}},"children":[{{"type":"ListItem","props":{{}},"children":["item"]}}]}}]}}]}}}}"#
            ));
            let error = ir.validate().expect_err("invalid overrides must fail");
            assert!(
                error.to_string().contains(expected),
                "expected {expected:?} for {overrides}, got {error}"
            );
        }
    }

    #[test]
    fn validate_should_reject_unknown_multi_level_type() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{"multiLevelType":"multilevelPlus"},"children":[{"type":"ListItem","props":{},"children":["item"]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("unknown multi-level type must fail");
        assert!(
            error.to_string().contains("multiLevelType")
                && error.to_string().contains("singleLevel")
                && error.to_string().contains("hybridMultilevel"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_local_levels_with_num_style_link() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"List","props":{"numStyleLink":"ListNumber","type":"ordered"},"children":[{"type":"ListItem","props":{},"children":["item"]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("numStyleLink cannot define local list properties");
        assert!(
            error.to_string().contains("numStyleLink") && error.to_string().contains("type"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_unknown_section_type() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"type":"continuousPlus"},"children":[]}]}}"#,
        );
        let error = ir.validate().expect_err("unknown section type must fail");
        assert!(
            error.to_string().contains("type") && error.to_string().contains("nextPage"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_invalid_section_line_numbering() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"lineNumbering":{"countBy":0,"restart":"eachPage"}},"children":[]}]}}"#,
        );
        let error = ir.validate().expect_err("invalid line numbering must fail");
        assert!(
            error.to_string().contains("lineNumbering.countBy")
                || error.to_string().contains("lineNumbering.restart"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_unknown_section_line_number_restart() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"lineNumbering":{"restart":"eachPage"}},"children":[]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("unknown line-number restart must fail");
        assert!(
            error.to_string().contains("eachPage") && error.to_string().contains("newSection"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_accept_hyperlink_composite_children() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"href":"https://example.com"},"children":[{"type":"ContentControl","props":{"alias":"LinkData"},"children":["bound"]},{"type":"Inserted","props":{"author":"Ada"},"children":["new"]},{"type":"Deleted","props":{"author":"Lin"},"children":["old"]},{"type":"InlineBookmark","props":{"name":"insideLink"},"children":["marked"]},{"type":"Comment","props":{"text":"review","author":"Grace"},"children":["commented"]}]}]}]}]}}"#,
        );
        ir.validate()
            .expect("hyperlink composite children should validate");
    }

    #[test]
    fn validate_should_accept_section_block_content_control() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"ContentControl","props":{"alias":"BlockData"},"children":[{"type":"Paragraph","props":{},"children":["inside"]}]}]}]}}"#,
        );
        ir.validate()
            .expect("section-level ContentControl with a paragraph should validate");
    }

    #[test]
    fn validate_should_accept_header_block_content_control() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Header","props":{},"children":[{"type":"ContentControl","props":{"alias":"HeaderData"},"children":[{"type":"Paragraph","props":{},"children":["header"]}]}]}]}]}}"#,
        );
        ir.validate()
            .expect("header block ContentControl should validate");
    }

    #[test]
    fn validate_should_reject_section_content_control_with_inline_children() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"ContentControl","props":{"alias":"BlockData"},"children":["loose"]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("section ContentControl must reject inline text");
        assert!(
            error
                .to_string()
                .contains("document-level ContentControl requires block children"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_mixed_block_and_inline_content_control() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"ContentControl","props":{"alias":"Mixed"},"children":[{"type":"Paragraph","props":{},"children":["block"]},{"type":"Run","props":{},"children":["inline"]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("mixed ContentControl children must fail");
        assert!(
            error
                .to_string()
                .contains("cannot mix block and inline children"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_nested_hyperlink() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Hyperlink","props":{"href":"https://example.com"},"children":[{"type":"Hyperlink","props":{"anchor":"intro"},"children":["inner"]}]}]}]}]}}"#,
        );
        let error = ir.validate().expect_err("nested Hyperlink must fail");
        assert!(
            error
                .to_string()
                .contains("Hyperlink cannot contain Hyperlink"),
            "{error}"
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
    fn validate_should_accept_toc_item_with_bookmark_and_page_placeholder() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[{"type":"TocItem","props":{"text":"Overview","bookmark":"overview","level":2,"pageRef":"7"},"children":[]}]}]}]}}"#,
        );
        ir.validate()
            .expect("TOC item with a bookmark should validate");
    }

    #[test]
    fn validate_should_reject_toc_item_without_bookmark() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[{"type":"TocItem","props":{"text":"Overview"},"children":[]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("TocItem requires a target bookmark");
        assert!(
            error.to_string().contains("requires non-empty `bookmark`"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_accept_toc_before_and_after_paragraph_wrappers() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[{"type":"TocBefore","props":{},"children":[{"type":"Paragraph","props":{},"children":["before"]}]},{"type":"TocAfter","props":{},"children":[{"type":"Paragraph","props":{},"children":["after"]}]}]}]}]}}"#,
        );
        ir.validate()
            .expect("TOC before/after paragraph wrappers should validate");
    }

    #[test]
    fn validate_should_reject_duplicate_toc_before_wrapper() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[{"type":"TocBefore","props":{},"children":[]},{"type":"TocBefore","props":{},"children":[]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("only one before wrapper may map to the API");
        assert!(error.to_string().contains("at most one `TocBefore`"));
    }

    #[test]
    fn validate_should_reject_empty_toc_after_wrapper() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[{"type":"TocAfter","props":{},"children":[]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("empty after wrapper has no content to serialize");
        assert!(
            error
                .to_string()
                .contains("TocAfter requires at least one Paragraph")
        );
    }

    #[test]
    fn validate_should_reject_toc_item_tab_position_outside_openxml_range() {
        let error = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{},"children":[{"type":"TocItem","props":{"text":"Overview","bookmark":"overview","tabPosition":1600},"children":[]}]}]}]}}"#,
        )
        .validate()
        .expect_err("out-of-range tab stop must fail validation");
        assert!(error.to_string().contains("31680-twip limit"));
        assert!(error.to_string().contains("reduce the position"));
    }

    #[test]
    fn validate_should_accept_paragraph_and_heading_run_defaults() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"snapToGrid":false,"widowControl":true,"font":"Noto Sans CJK SC","size":12,"bold":true,"italic":false,"color":"1a2B3c","characterSpacing":0.5},"children":["body"]},{"type":"Heading","props":{"level":2,"font":"Noto Sans CJK SC","size":16},"children":["title"]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_suppress_line_numbers_on_paragraph_and_list_item() {
        let input = r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"suppressLineNumbers":true},"children":["body"]},{"type":"List","props":{},"children":[{"type":"ListItem","props":{"suppressLineNumbers":false},"children":["item"]}]}]}]}}"#;
        let ir = parse(input);
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_non_boolean_suppress_line_numbers() {
        let input = r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"suppressLineNumbers":"yes"},"children":[]}]}]}}"#;
        let error = parse(input).validate().unwrap_err();
        assert!(
            error
                .to_string()
                .contains("suppressLineNumbers` must be a boolean")
        );
    }

    #[test]
    fn validate_should_accept_style_line_number_suppression() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"NoLineNumbers","name":"No Line Numbers","type":"paragraph","paragraph":{"suppressLineNumbers":true}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_non_boolean_style_line_number_suppression() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"NoLineNumbers","name":"No Line Numbers","type":"paragraph","paragraph":{"suppressLineNumbers":"yes"}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        let error = ir.validate().expect_err("style property must be boolean");
        assert!(error.to_string().contains("suppressLineNumbers"), "{error}");
    }

    #[test]
    fn validate_should_accept_font_slots_on_runs_paragraphs_and_styles() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","run":{"fonts":{"ascii":"Times New Roman","eastAsia":"宋体","hint":"eastAsia"}}}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"fonts":{"ascii":"Arial","eastAsia":"黑体"}},"children":[{"type":"Run","props":{"fonts":{"hiAnsi":"Calibri","csTheme":"minorBidi"}},"children":["mixed"]}]},{"type":"Heading","props":{"level":1,"fonts":{"eastAsiaTheme":"majorEastAsia"}},"children":["title"]}]}]}}"#,
        );
        ir.validate()
            .expect("font slots should validate everywhere");
    }

    #[test]
    fn validate_should_reject_ambiguous_font_slots_everywhere() {
        for source in [
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Run","props":{"font":"Arial","fonts":{"eastAsia":"宋体"}},"children":[]}]}]}]}}"#,
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"font":"Arial","fonts":{"eastAsia":"宋体"}},"children":[]}]}]}}"#,
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","run":{"font":"Arial","fonts":{"eastAsia":"宋体"}}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        ] {
            let error = parse(source)
                .validate()
                .expect_err("ambiguous fonts must fail");
            assert!(
                error.to_string().contains("cannot be used together"),
                "{error}"
            );
        }
    }

    #[test]
    fn validate_should_accept_advanced_paragraph_properties_and_revisions() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"bidi":true,"textAlign":"baseline","adjustRightIndent":1,"shading":"DDEEFF","outlineLevel":3,"frame":{"wrap":"around","horizontalAnchor":"page","x":12,"yAlign":"top","width":144,"height":72},"inserted":{"author":"Ada","date":"2026-08-14T00:00:00Z"},"propertyChange":{"author":"Lin","date":"2026-08-13T00:00:00Z","previous":{"align":"right","spacingAfter":6,"bidi":false}}},"children":["body"]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_paragraph_id_and_borders() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"paragraphId":"A1B2C3D4","border":{"top":{"style":"double","size":1,"color":"336699","space":2},"between":false,"bar":{"style":"babyRattle"}}},"children":["one"]},{"type":"Paragraph","props":{"border":{"style":"dashed","size":0.5,"color":"993366","space":1}},"children":["two"]},{"type":"Paragraph","props":{"border":{"clearAll":true}},"children":[]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_paragraph_id_and_borders() {
        for (props, expected) in [
            (r#"{"paragraphId":"123"}"#, "eight hexadecimal"),
            (r#"{"paragraphId":"ZZZZZZZZ"}"#, "eight hexadecimal"),
            (
                r#"{"border":{"style":"single","top":false}}"#,
                "uniform and positioned",
            ),
            (
                r#"{"border":{"between":{"space":-1}}}"#,
                "non-negative integer",
            ),
        ] {
            let ir = parse(&format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{}},"children":[{{"type":"Section","props":{{}},"children":[{{"type":"Paragraph","props":{props},"children":[]}}]}}]}}}}"#
            ));
            let error = ir.validate().expect_err("fixture must fail");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_reject_invalid_advanced_paragraph_properties() {
        for (source, expected) in [
            (
                r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"outlineLevel":10},"children":[]}]}]}}"#,
                "`outlineLevel` must be an integer from 0 through 9",
            ),
            (
                r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"inserted":{},"deleted":{}},"children":[]}]}]}}"#,
                "`inserted` and `deleted` are mutually exclusive",
            ),
            (
                r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"propertyChange":{"previous":{}}},"children":[]}]}]}}"#,
                "`propertyChange.previous` must not be empty",
            ),
            (
                r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Heading","props":{"level":2,"outlineLevel":7},"children":[]}]}]}}"#,
                "unknown property `outlineLevel`",
            ),
        ] {
            let error = parse(source).validate().expect_err("fixture must fail");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_accept_run_theme_color_and_explicit_formatting_flags() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Run","props":{"color":"2E74B5","themeColor":"accent1","themeShade":"BF","themeTint":"99","bold":false,"italic":false,"strike":false,"doubleStrike":true},"children":["themed"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_extended_table_and_cell_properties() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Table","props":{"style":"GridTable4","indent":12,"margins":{"top":2,"right":3,"bottom":4,"left":5}},"children":[{"type":"TableRow","props":{},"children":[{"type":"TableCell","props":{"widthPercent":50,"verticalMerge":"restart","textDirection":"tbRl","margins":{"top":1,"right":2,"bottom":3,"left":4}},"children":[{"type":"TableOfContents","props":{},"children":[]},{"type":"ContentControl","props":{"alias":"Cell"},"children":["value"]}]}]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_advanced_image_properties() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Image","props":{"src":"figure.png","width":120,"height":80,"relationshipId":"rIdHero","rotate":45,"floating":true,"allowOverlap":true,"layoutInCell":true,"positionH":"right","positionV":18.5,"relativeFromH":"page","relativeFromV":"paragraph","simplePosition":{"x":25400,"y":-12700},"distanceTop":2,"distanceBottom":3,"distanceLeft":4,"distanceRight":5,"relativeHeight":251658240},"children":[]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_advanced_image_properties() {
        for (props, expected) in [
            (
                r#"{"src":"x.png","width":1,"height":1,"positionH":"right"}"#,
                "require `floating={true}`",
            ),
            (
                r#"{"src":"x.png","width":1,"height":1,"floating":true,"positionV":"left"}"#,
                "positionV",
            ),
            (
                r#"{"src":"x.png","width":1,"height":1,"rotate":1.5}"#,
                "integer",
            ),
            (
                r#"{"src":"x.png","width":1,"height":1,"floating":true,"distanceTop":-1}"#,
                "non-negative",
            ),
            (
                r#"{"src":"x.png","width":1,"height":1,"layoutInCell":true}"#,
                "require `floating={true}`",
            ),
            (
                r#"{"src":"x.png","width":1,"height":1,"floating":true,"layoutInCell":"yes"}"#,
                "layoutInCell` must be a boolean",
            ),
            (
                r#"{"src":"x.png","width":1,"height":1,"simplePosition":{"x":0,"y":0}}"#,
                "`simplePosition` requires `floating={true}`",
            ),
            (
                r#"{"src":"x.png","width":1,"height":1,"floating":true,"simplePosition":{"x":2147483648,"y":0}}"#,
                "signed 32-bit integer",
            ),
        ] {
            let ir = parse(&format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{}},"children":[{{"type":"Section","props":{{}},"children":[{{"type":"Paragraph","props":{{}},"children":[{{"type":"Image","props":{props},"children":[]}}]}}]}}]}}}}"#
            ));
            let error = ir.validate().expect_err("fixture must fail");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_accept_positioned_and_cleared_borders() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Table","props":{"border":{"top":{"style":"double","size":1,"color":"336699"},"insideHorizontal":false}},"children":[{"type":"TableRow","props":{},"children":[{"type":"TableCell","props":{"border":{"left":false,"topLeftToBottomRight":{"style":"dotted","size":0.5,"color":"993366"}}},"children":[]}]}]},{"type":"Table","props":{"border":{"clearAll":true}},"children":[]}]}]}}"#,
        );
        ir.validate().expect("advanced borders should validate");
    }

    #[test]
    fn validate_should_reject_ambiguous_or_invalid_positioned_borders() {
        for (border, expected) in [
            (
                r#"{"style":"single","top":{"style":"double"}}"#,
                "mutually exclusive",
            ),
            (r#"{"clearAll":true,"left":false}"#, "cannot be combined"),
            (r#"{"top":true}"#, "object or false"),
            (
                r#"{"topLeftToBottomRight":{"style":"single"}}"#,
                "unknown property",
            ),
        ] {
            let source = format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{}},"children":[{{"type":"Section","props":{{}},"children":[{{"type":"Table","props":{{"border":{border}}},"children":[]}}]}}]}}}}"#
            );
            let error = parse(&source)
                .validate()
                .expect_err("invalid border should be rejected");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_accept_floating_table_and_row_revisions() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Table","props":{"position":{"leftFromText":7.1,"rightFromText":7.1,"verticalAnchor":"text","horizontalAnchor":"margin","xAlign":"right","y":25.5}},"children":[{"type":"TableRow","props":{"gridBefore":1,"gridAfter":2,"widthBefore":60,"widthAfter":75,"inserted":{"author":"Ada","date":"2026-08-14T00:00:00Z"}},"children":[{"type":"TableCell","props":{},"children":[]}]},{"type":"TableRow","props":{"deleted":{"author":"Linus"}},"children":[{"type":"TableCell","props":{},"children":[]}]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_table_row_grid_properties() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Table","props":{},"children":[{"type":"TableRow","props":{"gridBefore":-1},"children":[]}]}]}]}}"#,
        );

        assert!(
            ir.validate()
                .expect_err("gridBefore must be a non-negative integer")
                .to_string()
                .contains("gridBefore"),
        );
    }

    #[test]
    fn validate_should_reject_conflicting_table_position_and_row_revisions() {
        let position = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Table","props":{"position":{"x":10,"xAlign":"right"}},"children":[]}]}]}}"#,
        );
        assert!(
            position
                .validate()
                .expect_err("x position modes must be exclusive")
                .to_string()
                .contains("mutually exclusive")
        );

        let revisions = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Table","props":{},"children":[{"type":"TableRow","props":{"inserted":{},"deleted":{}},"children":[]}]}]}]}}"#,
        );
        assert!(
            revisions
                .validate()
                .expect_err("row revisions must be exclusive")
                .to_string()
                .contains("mutually exclusive")
        );
    }

    #[test]
    fn validate_should_reject_invalid_table_margins_and_text_direction() {
        let margins = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Table","props":{"margins":{"top":1,"right":1,"bottom":1}},"children":[]}]}]}}"#,
        );
        assert!(
            margins
                .validate()
                .expect_err("incomplete margins must fail")
                .to_string()
                .contains("margins")
        );

        let direction = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Table","props":{},"children":[{"type":"TableRow","props":{},"children":[{"type":"TableCell","props":{"textDirection":"sideways"},"children":[]}]}]}]}]}}"#,
        );
        assert!(
            direction
                .validate()
                .expect_err("unknown direction must fail")
                .to_string()
                .contains("textDirection")
        );
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
    fn validate_should_accept_comment_range_inside_deleted_text() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Deleted","props":{"author":"Ada"},"children":[{"type":"Comment","props":{"text":"Typo","author":"Grace"},"children":["removed"]}]}]}]}]}}"#,
        );
        ir.validate()
            .expect("a comment range over deleted text should validate");
    }

    #[test]
    fn validate_should_reject_formatted_wrapper_inside_deleted_comment_range() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Deleted","props":{},"children":[{"type":"Comment","props":{"text":"Typo"},"children":[{"type":"Hyperlink","props":{"href":"https://example.com"},"children":["removed"]}]}]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("comment range in deleted text must not accept unsupported wrappers");
        assert!(
            error
                .to_string()
                .contains("Comment inside Deleted can select only text"),
            "{error}"
        );
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
    fn validate_should_reject_bound_content_control_without_store_item_id() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"ContentControl","props":{"xpath":"/root/customer"},"children":["Ada"]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("data binding without storeItemId is invalid OOXML");
        assert!(
            error
                .to_string()
                .contains("requires non-empty `storeItemId`"),
            "{error}"
        );
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
    fn validate_should_accept_extended_toc_instruction_options() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{"tcFieldLevelRange":{"start":2,"end":4},"omitPageNumbersLevelRange":{"start":5,"end":7},"entryBookmarkName":"contents","entryAndPageNumberSeparator":" — ","stylesWithLevels":[{"style":"Report Heading","level":2}],"preserveTab":true,"preserveNewLine":true,"useAppliedParagraphLineLevel":true,"hideTabAndPageNumbersInWebview":true},"children":[]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_toc_instruction_options() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{"tcFieldLevelRange":{"start":5,"end":2},"preserveTab":"yes","stylesWithLevels":[{"style":"Bad,Name","level":0}]},"children":[]}]}]}}"#,
        );
        let error = ir.validate().expect_err("invalid TOC options must fail");
        assert!(
            error.to_string().contains("tcFieldLevelRange")
                && error.to_string().contains("start")
                && error.to_string().contains("end"),
            "{error}"
        );

        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{"preserveTab":"yes"},"children":[]}]}]}}"#,
        );
        assert!(
            ir.validate()
                .expect_err("TOC switches must be booleans")
                .to_string()
                .contains("`preserveTab` must be a boolean")
        );

        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfContents","props":{"stylesWithLevels":[{"style":"Bad,Name","level":2}]},"children":[]}]}]}}"#,
        );
        assert!(
            ir.validate()
                .expect_err("unrepresentable style names must fail")
                .to_string()
                .contains("cannot contain commas")
        );

        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"TableOfEntries","props":{"identifier":"bad\"id"},"children":[]}]}]}}"#,
        );
        assert!(
            ir.validate()
                .expect_err("quoted switch values must be rejected")
                .to_string()
                .contains("cannot contain a quote")
        );
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

    #[test]
    fn validate_should_accept_section_page_configuration() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"titlePage":true,"textDirection":"tbRl","documentGrid":{"type":"linesAndChars","linePitch":18,"charSpace":-10},"pageNumbering":{"start":3,"chapterStyle":"1","format":"upperRoman","chapterSeparator":"emDash"},"columns":2},"children":[]}]}}"#,
        );
        ir.validate()
            .expect("section configuration should validate");
    }

    #[test]
    fn validate_should_accept_equal_and_individually_sized_section_columns() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"columns":{"count":3,"spacing":18,"separator":true}},"children":[]},{"type":"Section","props":{"columns":{"widths":[{"width":180,"spacing":24},{"width":240,"spacing":18},{"width":300}],"separator":false}},"children":[]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_individually_sized_section_columns() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{"columns":{"widths":[{"width":10},{"width":-1}] }},"children":[]}]}}"#,
        );
        let error = ir.validate().expect_err("negative column width must fail");
        assert!(
            error.to_string().contains("columns.widths[1]")
                && error
                    .to_string()
                    .contains("finite non-negative point measure"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_accept_document_settings_and_metadata() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"defaultCharacterSpacing":0.5,"createdAt":"2026-08-14T00:00:00Z","updatedAt":"2026-08-15T00:00:00Z","creator":"Ada","lastModifiedBy":"Lin","description":"Report","language":"zh-CN","revision":2,"subject":"Weekly","title":"Status","documentProtection":{"edit":"readOnly","formatting":false,"enforcement":true,"cryptProviderType":"rsaAES","cryptAlgorithmClass":"hash","cryptAlgorithmType":"typeAny","cryptAlgorithmSid":14,"cryptSpinCount":100000,"cryptProvider":"RSA AES Provider","algIdExt":"A1B2C3D4","algIdExtSource":"urn:example:alg","cryptProviderTypeExt":"01020304","cryptProviderTypeExtSource":"urn:example:provider","hash":"YWJjZA==","salt":"AQIDBA=="},"customProperties":{"Project":"Apollo"},"documentId":"01234567-89AB-CDEF-0123-456789ABCDEF","defaultTabStop":36,"documentVariables":{"Customer":"Ada"},"evenAndOddHeaders":true,"adjustLineHeightInTable":true,"characterSpacingControl":"compressPunctuation"},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        ir.validate()
            .expect("document settings and metadata should validate");
    }

    #[test]
    fn validate_should_accept_distinct_default_font_slots() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"defaultFonts":{"ascii":"Times New Roman","hiAnsi":"Times New Roman","eastAsia":"宋体","cs":"Times New Roman","asciiTheme":"minorHAnsi","eastAsiaTheme":"minorEastAsia","hint":"eastAsia"}},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        ir.validate().expect("default font slots should validate");
    }

    #[test]
    fn validate_should_reject_invalid_or_ambiguous_default_fonts() {
        for (props, expected) in [
            (r#"{"defaultFonts":{}}"#, "must not be empty"),
            (
                r#"{"defaultFonts":{"eastAsia":""}}"#,
                "must be a non-empty string",
            ),
            (
                r#"{"defaultFonts":{"unknown":"Arial"}}"#,
                "unknown property `unknown`",
            ),
            (
                r#"{"defaultFont":"Arial","defaultFonts":{"eastAsia":"宋体"}}"#,
                "cannot be used together",
            ),
        ] {
            let ir = parse(&format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{props},"children":[{{"type":"Section","props":{{}},"children":[]}}]}}}}"#
            ));
            let error = ir.validate().expect_err("fixture must fail");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_accept_complete_line_spacing_properties() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"defaultLineSpacing":{"before":2,"after":6,"line":14,"beforeLines":100,"afterLines":200,"lineRule":"atLeast"},"styles":[{"id":"Dense","name":"Dense","type":"paragraph","paragraph":{"spacingBeforeLines":50,"spacingAfterLines":75,"lineRule":"exact"}}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"spacingBefore":3,"spacingAfter":4,"lineSpacing":12,"spacingBeforeLines":125,"spacingAfterLines":250,"lineRule":"auto"},"children":["text"]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_accept_web_extensions_and_custom_xml_items() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"webExtensions":[{"id":"7f33b723-fb58-4524-8733-dbedc4b7c095","referenceId":"office-addin","version":"1.0.0.0","store":"developer","storeType":"Registry","properties":{"mode":"review"}},{"id":"11111111-2222-3333-4444-555555555555","referenceId":"second","version":"2.0","store":"OMEX","storeType":"Marketplace"}],"customXmlItems":[{"id":"06AC5857-5C65-A94A-BCEC-37356A209BC3","xml":"<customer><name>Ada</name></customer>"},{"id":"11111111-AAAA-BBBB-CCCC-222222222222","xml":"<order id=\"42\"/>"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_web_extensions_and_custom_xml_items() {
        for (props, expected) in [
            (
                r#"{"webExtensions":[]}"#,
                "`webExtensions` must not be empty",
            ),
            (
                r#"{"webExtensions":[{"id":"same","referenceId":"a","version":"1","store":"s","storeType":"t"},{"id":"same","referenceId":"b","version":"1","store":"s","storeType":"t"}]}"#,
                "duplicate web extension id `same`",
            ),
            (
                r#"{"customXmlItems":[{"id":"x","xml":"<broken>"}]}"#,
                "well-formed XML",
            ),
            (
                r#"{"customXmlItems":[{"id":"x","xml":"<a/>"},{"id":"x","xml":"<b/>"}]}"#,
                "duplicate custom XML id `x`",
            ),
        ] {
            let ir = parse(&format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{props},"children":[{{"type":"Section","props":{{}},"children":[]}}]}}}}"#
            ));
            let error = ir.validate().expect_err("fixture must fail");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_reject_invalid_complete_line_spacing_properties() {
        for (source, expected) in [
            (
                r#"{"version":1,"document":{"type":"Document","props":{"defaultLineSpacing":{}},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
                "`defaultLineSpacing` must not be empty",
            ),
            (
                r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"spacingBeforeLines":-1},"children":[]}]}]}}"#,
                "`spacingBeforeLines` must be a non-negative integer",
            ),
            (
                r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"lineRule":"sometimes"},"children":[]}]}]}}"#,
                "invalid `lineRule`",
            ),
        ] {
            let error = parse(source).validate().expect_err("fixture must fail");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_accept_paragraph_character_unit_indentation() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"indentLeftChars":-100,"indentRightChars":200,"firstLineChars":640},"children":["paragraph"]},{"type":"Heading","props":{"level":2,"indentStartChars":100,"indentEndChars":-20,"hangingChars":200},"children":["heading"]}] }]}}"#,
        );
        ir.validate()
            .expect("valid character-unit indentation should pass validation");
    }

    #[test]
    fn validate_should_reject_conflicting_paragraph_character_indentation() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"firstLine":12,"firstLineChars":120},"children":[]}] }]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("the same indentation mode cannot use both units");
        assert!(error.to_string().contains("mutually exclusive"), "{error}");
    }

    #[test]
    fn validate_should_accept_custom_style_definitions() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"ReportTitle","name":"Report Title","type":"paragraph","basedOn":"Normal","next":"Normal","quickFormat":true,"uiPriority":5,"run":{"font":"Noto Sans CJK SC","size":18,"color":"336699","themeColor":"accent1","themeTint":"99","bold":true,"italic":false,"underline":"single","textBorder":{"style":"double","size":1,"color":"336699","space":2}},"paragraph":{"align":"center","textAlign":"baseline","snapToGrid":false,"keepNext":true,"keepLines":true,"pageBreakBefore":true,"widowControl":false,"bidi":true,"shading":"DDEEFF","adjustRightIndent":1,"spacingAfter":12,"indentLeft":6,"firstLine":2,"outlineLevel":1,"frame":{"wrap":"around","horizontalAnchor":"margin","verticalAnchor":"text","xAlign":"center","y":12,"horizontalSpace":3,"width":240,"height":48}}},{"id":"ReportTable","name":"Report Table","type":"table","table":{"indent":6,"align":"center","margins":{"top":1,"right":2,"bottom":3,"left":4},"border":{"style":"double","size":1,"color":"336699"}},"cell":{"width":72,"colSpan":2,"verticalAlign":"center","verticalMerge":"restart","textDirection":"tbRl","shading":"FFF2CC","margins":{"top":1,"right":2,"bottom":3,"left":4},"border":{"style":"dotted","size":0.5,"color":"993366"}}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        ir.validate().expect("custom style should validate");
    }

    #[test]
    fn validate_should_accept_extended_style_run_effects() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Effects","name":"Effects","type":"character","run":{"fitText":{"width":42,"id":9},"stretch":125,"specVanish":true,"shading":{"fill":"FFF2CC","color":"336699","pattern":"pct20"}}}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        ir.validate()
            .expect("extended style run effects should validate");
    }

    #[test]
    fn validate_should_reject_invalid_extended_style_run_effects() {
        for (run, expected) in [
            (r#"{"fitText":{"width":0}}"#, "positive"),
            (r#"{"fitText":{"width":42,"extra":1}}"#, "unknown property"),
            (r#"{"stretch":0}"#, "1..=600"),
            (r#"{"stretch":601}"#, "1..=600"),
            (r#"{"specVanish":false}"#, "only accepts `true`"),
            (
                r#"{"shading":{"color":"336699"}}"#,
                "requires six-digit RGB `fill`",
            ),
            (
                r#"{"shading":{"fill":"FFF2CC","pattern":"waves"}}"#,
                "invalid",
            ),
        ] {
            let ir = parse(&format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{"styles":[{{"id":"Effects","name":"Effects","type":"character","run":{run}}}]}},"children":[{{"type":"Section","props":{{}},"children":[]}}]}}}}"#
            ));
            let error = ir
                .validate()
                .expect_err("invalid style run effect must fail");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_reject_invalid_custom_style_definitions() {
        for (styles, expected) in [
            (r"[]", "non-empty array"),
            (r"[{}]", "requires non-empty `id`"),
            (
                r#"[{"id":"A","name":"A","type":"paragraph"},{"id":"A","name":"B","type":"character"}]"#,
                "duplicate style id",
            ),
            (
                r#"[{"id":"A","name":"A","type":"page"}]"#,
                "invalid style type",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","run":{"color":"red"}}]"#,
                "six-digit RGB",
            ),
            (
                r#"[{"id":"A","name":"A","type":"character","run":{"strike":true,"doubleStrike":true}}]"#,
                "mutually exclusive",
            ),
            (
                r#"[{"id":"A","name":"A","type":"character","run":{"verticalAlign":"middle"}}]"#,
                "invalid `verticalAlign`",
            ),
            (
                r#"[{"id":"A","name":"A","type":"character","run":{"characterSpacing":"wide"}}]"#,
                "finite number",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","paragraph":{"firstLine":1,"hanging":1}}]"#,
                "mutually exclusive",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","paragraph":{"firstLine":1,"hangingChars":20}}]"#,
                "mutually exclusive",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","paragraph":{"indentLeftChars":2,"indentStartChars":3}}]"#,
                "aliases are mutually exclusive",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","paragraph":{"keepNext":"yes"}}]"#,
                "keepNext",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","paragraph":{"widowControl":"yes"}}]"#,
                "widowControl",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","paragraph":{"shading":"red"}}]"#,
                "six-digit RGB",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","paragraph":{"adjustRightIndent":-2}}]"#,
                "0 or 1",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","paragraph":{"frame":{"x":1,"xAlign":"left"}}}]"#,
                "mutually exclusive",
            ),
            (
                r#"[{"id":"A","name":"A","type":"paragraph","table":{"align":"center"}}]"#,
                "require type `table`",
            ),
            (
                r#"[{"id":"A","name":"A","type":"table","table":{"width":10}}]"#,
                "not legal in OOXML style properties",
            ),
        ] {
            let source = format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{"styles":{styles}}},"children":[{{"type":"Section","props":{{}},"children":[]}}]}}}}"#
            );
            let error = parse(&source)
                .validate()
                .expect_err("fixture should be rejected");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_accept_deleted_field_instruction() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Deleted","props":{},"children":[{"type":"Run","props":{},"children":[{"type":"DeletedInstruction","props":{"instruction":"HYPERLINK \"https://example.com\""},"children":[]}]}]}]}]}]}}"#,
        );
        ir.validate()
            .expect("deleted field instruction should validate");
    }

    #[test]
    fn validate_should_reject_deleted_instruction_outside_revision() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"DeletedInstruction","props":{"instruction":"HYPERLINK target"},"children":[]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("DeletedInstruction must be inside Deleted");
        assert!(
            error
                .to_string()
                .contains("must be inside a Run nested within Deleted"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_empty_deleted_instruction_with_repair() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Deleted","props":{},"children":[{"type":"Run","props":{},"children":[{"type":"DeletedInstruction","props":{"instruction":""},"children":[]}]}]}]}]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("empty field instructions should be rejected");
        assert!(
            error
                .to_string()
                .contains("requires a non-empty `instruction` string"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_invalid_document_settings_and_metadata() {
        for (props, expected) in [
            (r#""createdAt":"""#, "createdAt"),
            (r#""creator":"""#, "creator"),
            (r#""lastModifiedBy":"""#, "lastModifiedBy"),
            (r#""revision":0"#, "positive integer"),
            (r#""revision":1.5"#, "positive integer"),
            (r#""documentProtection":{}"#, "at least one protection"),
            (
                r#""documentProtection":{"edit":"unlock"}"#,
                "invalid `edit` value",
            ),
            (
                r#""documentProtection":{"enforcement":"yes"}"#,
                "must be a boolean",
            ),
            (
                r#""documentProtection":{"hash":"YWJjZA=="}"#,
                "must be supplied together",
            ),
            (
                r#""documentProtection":{"hashValue":"YWJjZA=="}"#,
                "`hashValue` and `saltValue` must be supplied together",
            ),
            (
                r#""documentProtection":{"hash":"not-base64!","salt":"AQIDBA=="}"#,
                "base64",
            ),
            (
                r#""documentProtection":{"cryptSpinCount":4294967296}"#,
                "32-bit integer",
            ),
            (
                r#""documentProtection":{"spinCount":4294967296}"#,
                "32-bit integer",
            ),
            (r#""documentProtection":{"algIdExt":"A1B2"}"#, "4 bytes"),
            (r#""customProperties":{}"#, "must not be empty"),
            (r#""customProperties":{"Project":1}"#, "string values"),
            (
                r#""documentVariables":{"":"Ada"}"#,
                "names must be non-empty",
            ),
            (r#""defaultTabStop":0"#, "positive"),
            (r#""evenAndOddHeaders":"yes""#, "boolean"),
            (
                r#""characterSpacingControl":"compressAll""#,
                "characterSpacingControl",
            ),
        ] {
            let source = format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{{props}}},"children":[{{"type":"Section","props":{{}},"children":[]}}]}}}}"#
            );
            let error = parse(&source)
                .validate()
                .expect_err("fixture should be rejected");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_accept_office_2010_document_protection_attributes() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"documentProtection":{"algorithmName":"SHA-512","hashValue":"YWJjZA==","saltValue":"AQIDBA==","spinCount":100000}},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_invalid_section_page_configuration() {
        for (props, expected) in [
            (r#""documentGrid":{}"#, "must not be empty"),
            (r#""documentGrid":{"type":"diagonal"}"#, "documentGrid.type"),
            (
                r#""documentGrid":{"type":"lines","linePitch":0}"#,
                "linePitch",
            ),
            (r#""pageNumbering":{}"#, "must not be empty"),
            (r#""pageNumbering":{"start":1.5}"#, "non-negative integer"),
            (
                r#""pageNumbering":{"format":"roman"}"#,
                "pageNumbering.format",
            ),
            (
                r#""pageNumbering":{"chapterSeparator":"slash"}"#,
                "chapterSeparator",
            ),
            (
                r#""columns":0"#,
                "`columns` must be an integer from 1 to 45",
            ),
            (
                r#""columns":46"#,
                "`columns` must be an integer from 1 to 45",
            ),
            (
                r#""columns":2.5"#,
                "`columns` must be an integer from 1 to 45",
            ),
            (r#""titlePage":"yes""#, "boolean"),
        ] {
            let source = format!(
                r#"{{"version":1,"document":{{"type":"Document","props":{{}},"children":[{{"type":"Section","props":{{{props}}},"children":[]}}]}}}}"#
            );
            let error = parse(&source)
                .validate()
                .expect_err("fixture should be rejected");
            assert!(error.to_string().contains(expected), "{error}");
        }
    }

    #[test]
    fn validate_should_accept_typed_style_references_and_inheritance() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"BodyBase","name":"Body Base","type":"paragraph","paragraph":{"spacingAfter":6}},{"id":"Body","name":"Body","type":"paragraph","basedOn":"BodyBase","run":{"font":"Arial"}},{"id":"Emphasis","name":"Emphasis","type":"character","run":{"italic":true}},{"id":"ReportTable","name":"Report Table","type":"table","table":{"align":"center"}}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"style":"Body"},"children":[{"type":"Run","props":{"style":"Emphasis"},"children":["text"]}]},{"type":"Table","props":{"style":"ReportTable"},"children":[]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_style_reference_type_mismatch() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Emphasis","name":"Emphasis","type":"character"}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"style":"Emphasis"},"children":[]}]}]}}"#,
        );
        let error = ir.validate().expect_err("style type mismatch must fail");
        assert!(
            error.to_string().contains("requires a paragraph style"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_cyclic_style_inheritance() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"A","name":"A","type":"paragraph","basedOn":"B"},{"id":"B","name":"B","type":"paragraph","basedOn":"A"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        let error = ir.validate().expect_err("style cycle must fail");
        assert!(error.to_string().contains("inheritance cycle"), "{error}");
    }

    #[test]
    fn validate_should_reject_inherited_style_type_mismatch() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","basedOn":"Emphasis"},{"id":"Emphasis","name":"Emphasis","type":"character"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("inherited style type mismatch must fail");
        assert!(error.to_string().contains("same type"), "{error}");
    }

    #[test]
    fn validate_should_accept_linked_styles_and_paragraph_next_style() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","next":"Body","link":"BodyChar"},{"id":"BodyChar","name":"Body Char","type":"character","link":"Body"}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"style":"Body"},"children":[{"type":"Run","props":{"style":"BodyChar"},"children":["text"]}]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }

    #[test]
    fn validate_should_reject_next_on_non_paragraph_style() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Strong","name":"Strong","type":"character","next":"Normal"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("next on character style must fail");
        assert!(
            error
                .to_string()
                .contains("`next` is only valid on paragraph styles"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_next_style_type_mismatch() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","next":"Strong"},{"id":"Strong","name":"Strong","type":"character"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        let error = ir.validate().expect_err("next target must be paragraph");
        assert!(
            error
                .to_string()
                .contains("`next` must reference a paragraph style"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_linked_style_type_mismatch() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","link":"Quote"},{"id":"Quote","name":"Quote","type":"paragraph"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("linked styles must pair paragraph and character styles");
        assert!(
            error.to_string().contains("paragraph and character style"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_reject_non_reciprocal_local_style_link() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","link":"BodyChar"},{"id":"BodyChar","name":"Body Char","type":"character"}]},"children":[{"type":"Section","props":{},"children":[]}]}}"#,
        );
        let error = ir
            .validate()
            .expect_err("local style links must be reciprocal");
        assert!(error.to_string().contains("must link back"), "{error}");
    }

    #[test]
    fn validate_should_reject_run_style_reference_to_table_style() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Grid","name":"Grid","type":"table"}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{},"children":[{"type":"Run","props":{"style":"Grid"},"children":["text"]}]}]}]}}"#,
        );
        let error = ir.validate().expect_err("Run requires a character style");
        assert!(
            error.to_string().contains("requires a character style"),
            "{error}"
        );
    }

    #[test]
    fn validate_should_allow_external_template_style_references() {
        let ir = parse(
            r#"{"version":1,"document":{"type":"Document","props":{"styles":[{"id":"Body","name":"Body","type":"paragraph","basedOn":"TemplateBody","next":"TemplateNext","link":"TemplateBodyChar"}]},"children":[{"type":"Section","props":{},"children":[{"type":"Paragraph","props":{"style":"TemplateBody"},"children":[]}]}]}}"#,
        );
        assert!(ir.validate().is_ok(), "{:?}", ir.validate());
    }
}
