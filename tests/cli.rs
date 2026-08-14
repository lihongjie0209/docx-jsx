use std::fs;
use std::io::Read;

use assert_cmd::Command;
use predicates::prelude::*;
use tempfile::tempdir;

fn embedded_ir(path: &std::path::Path) -> serde_json::Value {
    let bytes = fs::read(path).expect("DOCX should exist");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid ZIP");
    let mut json = String::new();
    archive
        .by_name("docx-jsx/ir-v1.json")
        .expect("embedded IR manifest")
        .read_to_string(&mut json)
        .expect("IR should read");
    serde_json::from_str(&json).expect("IR should be JSON")
}

#[test]
fn spec_should_output_agent_readable_markdown() {
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg("spec")
        .assert()
        .success()
        .stdout(predicate::str::contains("# docx-jsx v1 specification"))
        .stdout(predicate::str::contains("## Component tree"))
        .stdout(predicate::str::contains("`IndexEntry`"));
}

#[test]
fn spec_should_output_machine_readable_json_schema() {
    let output = Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["spec", "--format", "json-schema"])
        .output()
        .expect("spec command should run");
    assert!(output.status.success());
    let schema: serde_json::Value =
        serde_json::from_slice(&output.stdout).expect("stdout should be JSON");
    assert_eq!(schema["title"], "docx-jsx IR v1");
}

#[test]
fn validation_error_should_include_explanation_and_fix_suggestion() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("invalid.jsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph } from "docx-jsx";
export default <Document><Section><Paragraph unknownProp /></Section></Document>;"#,
    )
    .expect("input should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("explanation:"))
        .stderr(predicate::str::contains("suggestion:"))
        .stderr(predicate::str::contains("docx-jsx spec"));
}

#[test]
fn argument_error_should_include_fix_suggestion() {
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["spec", "--format", "xml"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("explanation:"))
        .stderr(predicate::str::contains("suggestion:"))
        .stderr(predicate::str::contains("docx-jsx spec --help"));
}

#[test]
fn validate_should_accept_valid_tsx_with_json_data_without_writing_docx() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("valid.tsx");
    let data = directory.path().join("data.json");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph } from "docx-jsx";
export default data => <Document><Section><Paragraph>{data.title}</Paragraph></Section></Document>;"#,
    )
    .expect("input should write");
    fs::write(&data, r#"{"title":"Validated"}"#).expect("data should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            "validate",
            input.to_str().expect("UTF-8 path"),
            "--data",
            data.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success()
        .stdout(predicate::str::contains(format!(
            "valid: {}",
            input.display()
        )));
    assert!(!input.with_extension("docx").exists());
}

#[test]
fn validate_should_reject_invalid_dsl_with_actionable_diagnostic() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("invalid.jsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph } from "docx-jsx";
export default <Document><Section><Paragraph unknownProp /></Section></Document>;"#,
    )
    .expect("input should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["validate", input.to_str().expect("UTF-8 path")])
        .assert()
        .failure()
        .stderr(predicate::str::contains("unknown property `unknownProp`"))
        .stderr(predicate::str::contains("explanation:"))
        .stderr(predicate::str::contains("suggestion:"))
        .stderr(predicate::str::contains("docx-jsx spec"));
}

#[test]
fn cli_should_compile_tsx_with_json_data() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("report.tsx");
    let data = directory.path().join("data.json");
    let output = directory.path().join("report.docx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, Run } from "docx-jsx";
export default data => <Document><Section><Paragraph><Run bold>Hello {data.name}</Run></Paragraph></Section></Document>;"#,
    )
    .expect("input should write");
    fs::write(&data, r#"{"name":"Ada"}"#).expect("data should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            input.to_str().expect("UTF-8 path"),
            "--data",
            data.to_str().expect("UTF-8 path"),
            "-o",
            output.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success();
    let bytes = fs::read(&output).expect("output should exist");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid ZIP");
    let mut xml = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document XML")
        .read_to_string(&mut xml)
        .expect("XML should read");
    assert!(xml.contains("Hello ") && xml.contains("Ada"));
}

#[test]
fn cli_should_refuse_to_overwrite_without_force() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("report.jsx");
    let output = directory.path().join("report.docx");
    fs::write(
        &input,
        r#"import { Document, Section } from "docx-jsx"; export default <Document><Section /></Document>;"#,
    )
    .expect("input should write");
    fs::write(&output, "existing").expect("output should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            input.to_str().expect("UTF-8 path"),
            "-o",
            output.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .failure()
        .stderr(predicate::str::contains("--force"));
}

#[test]
fn cli_should_reverse_docx_to_recompilable_jsx() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("source.tsx");
    let docx = directory.path().join("source.docx");
    let jsx = directory.path().join("reversed.jsx");
    let roundtrip = directory.path().join("roundtrip.docx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, Run, Break, Table, TableRow, TableCell } from "docx-jsx";
export default <Document><Section><Paragraph align="center"><Run bold italic color="336699">Hello &amp; world<Break type="page" /></Run></Paragraph><Table><TableRow><TableCell><Paragraph>Cell</Paragraph></TableCell></TableRow></Table></Section></Document>;"#,
    )
    .expect("input should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            input.to_str().expect("UTF-8 path"),
            "-o",
            docx.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success();
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            "reverse",
            docx.to_str().expect("UTF-8 path"),
            "-o",
            jsx.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success();

    let source = fs::read_to_string(&jsx).expect("JSX output should exist");
    assert!(source.contains("<Paragraph align=\"center\">"));
    assert!(
        source.contains("<Run bold")
            && source.contains("color=\"336699\"")
            && source.contains(" italic")
    );
    assert!(source.contains(r#"{"Hello & world"}"#));
    assert!(source.contains("<Break type=\"page\" />"));
    assert!(source.contains("<TableCell>"));

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            jsx.to_str().expect("UTF-8 path"),
            "-o",
            roundtrip.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success();
    assert_eq!(embedded_ir(&docx), embedded_ir(&roundtrip));
}

#[test]
fn reverse_should_refuse_to_overwrite_without_force() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("source.jsx");
    let docx = directory.path().join("source.docx");
    let output = directory.path().join("output.jsx");
    fs::write(&input, r#"import { Document, Section } from "docx-jsx"; export default <Document><Section /></Document>;"#)
        .expect("input should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            input.to_str().expect("UTF-8 path"),
            "-o",
            docx.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success();
    fs::write(&output, "existing").expect("output should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            "reverse",
            docx.to_str().expect("UTF-8 path"),
            "-o",
            output.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .failure()
        .stderr(predicate::str::contains("--force"));
}

#[test]
#[allow(clippy::too_many_lines)]
fn cli_should_compile_advanced_document_components() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("advanced.tsx");
    let output = directory.path().join("advanced.docx");
    let reversed = directory.path().join("advanced.reversed.jsx");
    let recompiled = directory.path().join("advanced.recompiled.docx");
    fs::write(
        &input,
        r##"import { Document, Section, Header, Footer, Heading, Caption, Index, IndexEntry, Paragraph, Hyperlink, PageNumber, TotalPages, List, ListItem, Run, Bookmark, Table, TableRow, TableCell, TableOfContents, TableOfFigures, TableOfEntries, TocEntry, Comment, Footnote, Tab, TabStop, CarriageReturn, NonBreakingSpace, SoftHyphen, NonBreakingHyphen, Symbol, Bold, Italic, Underline, StrikeThrough, Superscript, Subscript, AllCaps, HiddenText, SpecialHiddenText, DoubleStrike, SpacedText, ScaledText, FitText, BorderedText, ShadedText, Inserted, Deleted, MovedFrom, MovedTo, PageReference, PositionalTab, ContentControl, Field, DateField, TimeField, FileNameField, AuthorField, TitleField, SubjectField, SequenceField, ReferenceField, MergeField, DocumentPropertyField, FormulaField } from "docx-jsx";
export default <Document defaultCharacterSpacing={0.5} createdAt="2026-08-14T00:00:00Z" updatedAt="2026-08-15T00:00:00Z" customProperties={{Project: "Apollo"}} documentId="01234567-89AB-CDEF-0123-456789ABCDEF" defaultTabStop={36} documentVariables={{Customer: "Ada"}} evenAndOddHeaders adjustLineHeightInTable characterSpacingControl="compressPunctuation"><Section titlePage textDirection="tbRl" documentGrid={{type: "linesAndChars", linePitch: 18, charSpace: -10}} pageNumbering={{start: 3, chapterStyle: "1"}}>
  <Header type="first"><Paragraph>Report header</Paragraph></Header>
  <Footer type="even"><Paragraph>Page <PageNumber /> of <TotalPages /></Paragraph></Footer>
  <TableOfContents startLevel={1} endLevel={2} alias="Contents" />
  <TableOfFigures label="Figure" alias="Figures" />
  <TableOfEntries identifier="manual" alias="Manual entries" />
  <Paragraph><Hyperlink href="https://example.com" history><Run underline color="2E74B5" themeColor="accent1" themeShade="BF" bold={false} italic={false} strike={false} doubleStrike>Example</Run></Hyperlink></Paragraph>
  <Bookmark name="intro"><Heading level={1} font="Noto Sans CJK SC" size={16}>Introduction</Heading><Paragraph style="Quote" snapToGrid={false} widowControl font="Noto Sans CJK SC" size={12} bold italic={false} color="1a2B3c" characterSpacing={0.5}>Bookmarked text</Paragraph></Bookmark>
  <Paragraph><Hyperlink anchor="intro">Jump to introduction</Hyperlink></Paragraph>
  <Paragraph><TabStop position={72} align="right" leader="dot" /><TocEntry text="Manual entry" level={2} identifier="manual" />Introduction is on page <PageReference bookmark="intro" placeholder="?" relativePosition />.<Tab /><ContentControl alias="Customer" xpath="/root/customer">Ada</ContentControl><CarriageReturn /><MovedFrom author="Ada">old</MovedFrom><MovedTo author="Ada"><Run bold>new</Run></MovedTo><PositionalTab align="right" relativeTo="margin" leader="dot" /></Paragraph>
  <Paragraph><Comment text="Please verify" author="Ada">Reviewed text</Comment><Tab /><Symbol font="Wingdings" char="F0A7" /><Footnote>Footnote content</Footnote></Paragraph>
  <Paragraph>keep<NonBreakingSpace />together soft<SoftHyphen />hyphen non<NonBreakingHyphen />breaking</Paragraph>
  <Paragraph>H<Subscript>2</Subscript>O x<Superscript>2</Superscript> <AllCaps>draft</AllCaps><HiddenText>internal</HiddenText></Paragraph>
  <Paragraph><Bold>bold</Bold><Italic>italic</Italic><Underline type="wave">underlined</Underline><StrikeThrough>removed</StrikeThrough></Paragraph>
  <Paragraph><DoubleStrike>obsolete</DoubleStrike><SpacedText amount={1.5}>wide</SpacedText><ScaledText percent={125}>scaled</ScaledText><FitText width={42} id={7}>fitted</FitText><BorderedText style="double" size={1} color="336699" space={2}>bordered</BorderedText><ShadedText fill="FFF2CC" color="336699" pattern="pct20">shaded</ShadedText><SpecialHiddenText>metadata</SpecialHiddenText></Paragraph>
  <Paragraph>Generated: <Field instruction={' DATE \\@ "yyyy-MM-dd" '} dirty={false}><Run bold>2026-08-14</Run></Field></Paragraph>
  <Paragraph><DateField format="yyyy-MM-dd">2026-08-14</DateField><TimeField format="HH:mm" /><FileNameField fullPath>report.docx</FileNameField><AuthorField /><TitleField /><SubjectField /></Paragraph>
  <Paragraph>Figure <SequenceField identifier="Figure" format="Roman" restart={3} placeholder="III" />; see <ReferenceField bookmark="intro" relativePosition placeholder="Introduction" />.</Paragraph>
  <Caption label="Figure" identifier="Diagram" format="Roman" restart={3} placeholder="III" style="FigureCaption" textSeparator=" — "><Run bold>Architecture</Run></Caption>
  <Paragraph><IndexEntry text="Rust" subentry="Ownership" identifier="topics" boldPageNumber /></Paragraph>
  <Index identifier="topics" columns={2} runIn placeholder="Update index" style="IndexBody" />
  <Paragraph><MergeField name="CustomerName" placeholder="Ada" /><DocumentPropertyField name="Project Name">Apollo</DocumentPropertyField><FormulaField expression="SUM(ABOVE)" numberFormat="#,##0.00">42.00</FormulaField></Paragraph>
  <Paragraph><Deleted author="Ada" date="2026-08-14T00:00:00Z"><Run bold>Old text</Run></Deleted><Inserted author="Ada"><Run bold>New text</Run></Inserted></Paragraph>
  <List type="ordered" start={3}><ListItem>Third</ListItem><ListItem level={1}>Nested</ListItem></List>
  <Table style="GridTable4" indent={12} margins={{top: 2, right: 3, bottom: 4, left: 5}} position={{leftFromText: 7.1, rightFromText: 7.1, verticalAnchor: "text", horizontalAnchor: "margin", xAlign: "right", y: 25.5}}><TableRow inserted={{author: "Ada", date: "2026-08-14T00:00:00Z"}}><TableCell verticalMerge="restart" textDirection="tbRl" margins={{top: 1, right: 2, bottom: 3, left: 4}}><Paragraph>Cell</Paragraph><TableOfContents startLevel={1} endLevel={2} /><ContentControl alias="CellValue">value</ContentControl></TableCell></TableRow><TableRow deleted={{author: "Linus"}}><TableCell><Paragraph>Removed row</Paragraph></TableCell></TableRow></Table>
</Section></Document>;"##,
    )
    .expect("input should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            input.to_str().expect("UTF-8 path"),
            "-o",
            output.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success();

    let bytes = fs::read(&output).expect("output should exist");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid ZIP");
    let header_name = archive
        .file_names()
        .find(|name| {
            name.starts_with("word/header")
                && std::path::Path::new(name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        })
        .expect("header XML should exist")
        .to_owned();
    let footer_name = archive
        .file_names()
        .find(|name| {
            name.starts_with("word/footer")
                && std::path::Path::new(name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        })
        .expect("footer XML should exist")
        .to_owned();
    let mut read_part = |name: &str| {
        let mut xml = String::new();
        archive
            .by_name(name)
            .unwrap_or_else(|_| panic!("{name} should exist"))
            .read_to_string(&mut xml)
            .expect("XML should read");
        xml
    };
    assert!(read_part(&header_name).contains("Report header"));
    let footer = read_part(&footer_name);
    assert!(footer.contains("PAGE") && footer.contains("NUMPAGES"));
    let document = read_part("word/document.xml");
    assert!(document.contains("w:hyperlink") && document.contains("w:numPr"));
    assert!(document.contains("w:bookmarkStart") && document.contains("w:bookmarkEnd"));
    assert!(document.contains("Heading1") && document.contains("w:outlineLvl"));
    assert!(document.contains("TOC") && document.contains("Contents"));
    assert!(document.contains("w:commentRangeStart") && document.contains("w:commentReference"));
    assert!(document.contains("w:footnoteReference") && document.contains("w:tab"));
    assert!(document.contains("w:sym") && document.contains("F0A7"));
    assert!(document.contains("w:ins") && document.contains("w:del"));
    assert!(document.contains("w:delText") && document.contains("Old text"));
    assert!(document.contains("PAGEREF intro") && document.contains("w:dirty=\"true\""));
    assert!(document.contains("w:ptab") && document.contains("w:leader=\"dot\""));
    let relationships = read_part("word/_rels/document.xml.rels");
    assert!(relationships.contains("https://example.com"));
    assert!(relationships.contains("relationships/comments"));
    assert!(relationships.contains("relationships/footnotes"));
    let numbering = read_part("word/numbering.xml");
    assert!(numbering.contains("decimal") && numbering.contains("w:start w:val=\"3\""));
    assert!(read_part("word/comments.xml").contains("Please verify"));
    assert!(read_part("word/footnotes.xml").contains("Footnote content"));

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            "reverse",
            output.to_str().expect("UTF-8 path"),
            "-o",
            reversed.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success();
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            reversed.to_str().expect("UTF-8 path"),
            "-o",
            recompiled.to_str().expect("UTF-8 path"),
        ])
        .assert()
        .success();
    assert_eq!(embedded_ir(&output), embedded_ir(&recompiled));
}
