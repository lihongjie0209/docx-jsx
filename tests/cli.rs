use std::fs;
use std::io::{Read, Write};

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

fn strip_ir_manifest(bytes: &[u8]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid ZIP");
    let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let entry = archive.by_index(index).expect("ZIP entry should exist");
        if entry.name() != "docx-jsx/ir-v1.json" {
            writer.raw_copy_file(entry).expect("entry should copy");
        }
    }
    writer
        .finish()
        .expect("DOCX archive should finish")
        .into_inner()
}

fn first_index_field_paragraph_properties(path: &std::path::Path) -> String {
    let bytes = fs::read(path).expect("DOCX should exist");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid ZIP");
    let mut xml = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part")
        .read_to_string(&mut xml)
        .expect("XML should read");
    let control = xml.find("<w:sdtContent>").expect("TOC SDT content");
    let start = xml[control..]
        .find("<w:pPr")
        .map(|index| index + control)
        .expect("field paragraph properties");
    let end = xml[start..]
        .find("</w:pPr>")
        .map(|index| index + start + "</w:pPr>".len())
        .expect("field paragraph property end");
    xml[start..end].to_owned()
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
        .stdout(predicate::str::contains("`IndexEntry`"))
        .stdout(predicate::str::contains("`TocItem`"))
        .stdout(predicate::str::contains("`TocBefore`"))
        .stdout(predicate::str::contains("`TocAfter`"))
        .stdout(predicate::str::contains("`creator?: string`"))
        .stdout(predicate::str::contains("`revision?: integer`"))
        .stdout(predicate::str::contains("`documentProtection?:"))
        .stdout(predicate::str::contains("`trackRevisions?: boolean`"))
        .stdout(predicate::str::contains(
            "`underline?: boolean | UnderlineStyle`",
        ))
        .stdout(predicate::str::contains("`dottedHeavy`"))
        .stdout(predicate::str::contains("`tcFieldLevelRange?:"))
        .stdout(predicate::str::contains("`stylesWithLevels?:"))
        .stdout(predicate::str::contains(
            "`hideTabAndPageNumbersInWebview?:",
        ))
        .stdout(predicate::str::contains("type?: \"nextPage\""))
        .stdout(predicate::str::contains("lineNumbering?:"))
        .stdout(predicate::str::contains("docx-jsx pdf"))
        .stdout(predicate::str::contains("-o <OUTPUT.docx|pdf>"))
        .stdout(predicate::str::contains("inferred from the `-o` suffix"))
        .stdout(predicate::str::contains("DOCX_JSX_SOFFICE"))
        .stdout(predicate::str::contains("--engine"));
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
    assert!(
        schema["$defs"]["node"]["properties"]["type"]["enum"]
            .as_array()
            .is_some_and(|components| components.contains(&serde_json::json!("TocItem")))
            && schema["$defs"]["node"]["properties"]["type"]["enum"]
                .as_array()
                .is_some_and(|components| components.contains(&serde_json::json!("TocBefore")))
            && schema["$defs"]["node"]["properties"]["type"]["enum"]
                .as_array()
                .is_some_and(|components| components.contains(&serde_json::json!("TocAfter")))
    );
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
fn validate_should_reject_style_cycle_with_targeted_repair_suggestion() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("style-cycle.jsx");
    fs::write(
        &input,
        r#"import { Document, Section } from "docx-jsx";
export default <Document styles={[{id:"A",name:"A",type:"paragraph",basedOn:"B"},{id:"B",name:"B",type:"paragraph",basedOn:"A"}]}><Section /></Document>;"#,
    )
    .expect("input should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["validate", input.to_str().expect("UTF-8 path")])
        .assert()
        .failure()
        .stderr(predicate::str::contains("style inheritance cycle"))
        .stderr(predicate::str::contains(
            "Remove or redirect one `basedOn` reference",
        ));
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
export default <Document><Section type="continuous" lineNumbering={{countBy:5,restart:"newSection",distance:18}}><Paragraph align="center"><Run bold italic color="336699">Hello &amp; world<Break type="page" /></Run></Paragraph><Table><TableRow><TableCell><Paragraph>Cell</Paragraph></TableCell></TableRow></Table></Section></Document>;"#,
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
    assert!(source.contains(r#"type="continuous""#));
    assert!(
        source.contains(r#"lineNumbering={{"countBy":5,"restart":"newSection","distance":18}}"#),
        "{source}"
    );

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
fn cli_should_reverse_and_recompile_external_index_field_paragraph_formatting() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("index.tsx");
    let compiled = directory.path().join("index.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    let rebuilt = directory.path().join("rebuilt.docx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, TableOfContents, TableOfFigures } from "docx-jsx";
export default <Document><Section><TableOfContents fieldParagraph={{align:"center",spacingBefore:6,lineSpacing:28,lineRule:"exact",keepNext:true,bidi:true,shading:"EEF2F7",suppressLineNumbers:true,indentLeftChars:-100,indentRightChars:200,firstLineChars:640,font:"Arial",size:12,bold:false,italic:true,color:"336699",characterSpacing:1.25}} /><TableOfFigures label="Figure" fieldParagraph={{suppressLineNumbers:false,indentStartChars:120,indentEndChars:80,hangingChars:20,fonts:{ascii:"A",hiAnsi:"H",eastAsia:"E",cs:"C",asciiTheme:"majorAscii",hiAnsiTheme:"majorHAnsi",eastAsiaTheme:"majorEastAsia",csTheme:"majorBidi",hint:"eastAsia"}}} /><Paragraph bold={false} italic={false}>ordinary paragraph</Paragraph></Section></Document>;"#,
    )
    .expect("input should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    let field_paragraph = first_index_field_paragraph_properties(&compiled);
    assert!(
        field_paragraph.contains("<w:b w:val=\"false\" /><w:bCs w:val=\"0\" />")
            && field_paragraph.contains("<w:i />")
            && field_paragraph.contains("<w:iCs />")
            && field_paragraph.contains("<w:sz w:val=\"24\" />")
            && field_paragraph.contains("<w:spacing w:val=\"25\" />"),
        "writer output should preserve run defaults and normalize explicit complex-script off values: {field_paragraph}"
    );
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    assert!(
        jsx.contains("fieldParagraph={{")
            && jsx.contains(r#""lineSpacing":28"#)
            && jsx.contains(r#""bidi":true"#)
            && jsx.contains(r#""font":"Arial""#)
            && jsx.contains(r#""size":12"#)
            && jsx.contains(r#""bold":false"#)
            && jsx.contains(r#""italic":true"#)
            && jsx.contains(r#""color":"336699""#)
            && jsx.contains(r#""characterSpacing":1.25"#),
        "reverse should restore field paragraph properties without a manifest: {jsx}"
    );
    assert!(
        jsx.contains(r#""suppressLineNumbers":true"#)
            && jsx.contains(r#""suppressLineNumbers":false"#),
        "reverse should restore both line-number suppression values from external XML: {jsx}"
    );
    assert!(
        jsx.contains("ordinary paragraph")
            && jsx.contains("bold={false}")
            && jsx.contains("italic={false}"),
        "ordinary paragraph explicit run-default off values should survive external reverse: {jsx}"
    );
    assert!(
        jsx.contains(r#""csTheme":"majorBidi""#) && jsx.contains(r#""hint":"eastAsia""#),
        "external reverse should restore independent complex-script theme font and hint slots: {jsx}"
    );
    for indentation in [
        r#""indentLeftChars":-100"#,
        r#""indentRightChars":200"#,
        r#""firstLineChars":640"#,
        r#""indentStartChars":120"#,
        r#""indentEndChars":80"#,
        r#""hangingChars":20"#,
    ] {
        assert!(jsx.contains(indentation), "missing {indentation}: {jsx}");
    }
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();

    assert_eq!(
        first_index_field_paragraph_properties(&compiled),
        first_index_field_paragraph_properties(&rebuilt)
    );
}

#[test]
fn cli_should_roundtrip_external_threaded_comments_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("comments.tsx");
    let compiled = directory.path().join("comments.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    let rebuilt = directory.path().join("rebuilt.docx");
    let rebuilt_external = directory.path().join("rebuilt-external.docx");
    let rebuilt_jsx = directory.path().join("rebuilt.jsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, Comment } from "docx-jsx";
export default <Document><Section><Paragraph><Comment text="Automatic">automatic selection</Comment><Comment id={10} text="Root" author="Ada">root selection</Comment><Comment id={20} parentCommentId={10} done={true} text="Reply" author="Lin">reply selection</Comment></Paragraph></Section></Document>;"#,
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(
        fs::read(&compiled).expect("DOCX should exist"),
    ))
    .expect("compiled DOCX should be ZIP");
    let mut comment_replies = String::new();
    archive
        .by_name("word/commentsExtended.xml")
        .expect("commentsExtended part should be present for reply relationships")
        .read_to_string(&mut comment_replies)
        .expect("commentsExtended should be readable");
    assert!(
        comment_replies.contains("w15:paraIdParent=") && comment_replies.contains("w15:done=\"1\""),
        "the writer should emit the reply relationship: {comment_replies}"
    );

    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    assert!(
        jsx.contains("id={1}")
            && jsx.contains("id={10}")
            && jsx.contains("id={20}")
            && jsx.contains("parentCommentId={10}")
            && jsx.contains("done={true}"),
        "external reverse should preserve comment ids, reply parent, and resolved state: {jsx}"
    );

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    fs::write(
        &rebuilt_external,
        strip_ir_manifest(&fs::read(&rebuilt).expect("rebuilt DOCX should exist")),
    )
    .expect("rebuilt external DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&rebuilt_external)
        .args(["-o"])
        .arg(&rebuilt_jsx)
        .assert()
        .success();
    let rebuilt_jsx = fs::read_to_string(rebuilt_jsx).expect("rebuilt JSX should exist");
    assert!(
        rebuilt_jsx.contains("id={1}")
            && rebuilt_jsx.contains("id={10}")
            && rebuilt_jsx.contains("id={20}")
            && rebuilt_jsx.contains("parentCommentId={10}")
            && rebuilt_jsx.contains("done={true}"),
        "reply relationships and resolved state should survive recompilation: {rebuilt_jsx}"
    );
}

#[test]
fn cli_should_roundtrip_external_paragraph_mark_revisions_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("paragraph-revisions.tsx");
    let compiled = directory.path().join("paragraph-revisions.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    let rebuilt = directory.path().join("rebuilt.docx");
    let rebuilt_external = directory.path().join("rebuilt-external.docx");
    let rebuilt_jsx = directory.path().join("rebuilt.jsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph } from "docx-jsx";
export default <Document><Section><Paragraph inserted={{author:"Ada",date:"2026-08-14T00:00:00Z"}}>inserted paragraph</Paragraph><Paragraph deleted={{author:"Lin",date:"2026-08-15T00:00:00Z"}}>deleted paragraph</Paragraph></Section></Document>;"#,
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    assert!(
        jsx.contains(r#"inserted={{"author":"Ada","date":"2026-08-14T00:00:00Z"}}"#)
            && jsx.contains(r#"deleted={{"author":"Lin","date":"2026-08-15T00:00:00Z"}}"#),
        "paragraph mark revisions should be restored from external DOCX: {jsx}"
    );
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    fs::write(
        &rebuilt_external,
        strip_ir_manifest(&fs::read(&rebuilt).expect("rebuilt DOCX should exist")),
    )
    .expect("rebuilt external DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&rebuilt_external)
        .args(["-o"])
        .arg(&rebuilt_jsx)
        .assert()
        .success();
    let rebuilt_jsx = fs::read_to_string(rebuilt_jsx).expect("second reverse JSX should exist");
    assert!(
        rebuilt_jsx.contains(r#"inserted={{"author":"Ada","date":"2026-08-14T00:00:00Z"}}"#)
            && rebuilt_jsx.contains(r#"deleted={{"author":"Lin","date":"2026-08-15T00:00:00Z"}}"#),
        "paragraph mark revisions should survive recompilation: {rebuilt_jsx}"
    );
}

#[test]
fn cli_should_roundtrip_contextual_spacing_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("contextual-spacing.tsx");
    let compiled = directory.path().join("contextual-spacing.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    let rebuilt = directory.path().join("rebuilt.docx");
    let rebuilt_external = directory.path().join("rebuilt-external.docx");
    let rebuilt_jsx = directory.path().join("rebuilt.jsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, Heading, List, ListItem } from "docx-jsx";
export default <Document styles={[{id:"Body",name:"Body",type:"paragraph",paragraph:{contextualSpacing:false}}]}><Section><Paragraph contextualSpacing>on</Paragraph><Paragraph contextualSpacing={false}>off</Paragraph><Heading level={2} contextualSpacing>heading</Heading><List type="bullet"><ListItem contextualSpacing>item</ListItem></List></Section></Document>;"#,
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    assert!(
        jsx.contains("contextualSpacing={true}")
            && jsx.contains("contextualSpacing={false}")
            && jsx.contains(r#""contextualSpacing":false"#),
        "CLI reverse should preserve paragraph, heading, list-item, and style values: {jsx}"
    );
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    fs::write(
        &rebuilt_external,
        strip_ir_manifest(&fs::read(&rebuilt).expect("rebuilt DOCX should exist")),
    )
    .expect("rebuilt manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&rebuilt_external)
        .args(["-o"])
        .arg(&rebuilt_jsx)
        .assert()
        .success();
    let rebuilt_jsx = fs::read_to_string(rebuilt_jsx).expect("rebuilt reverse JSX should exist");
    assert!(
        rebuilt_jsx.contains("contextualSpacing={true}")
            && rebuilt_jsx.contains("contextualSpacing={false}")
            && rebuilt_jsx.contains(r#""contextualSpacing":false"#),
        "contextual spacing should survive a full CLI reverse/recompile cycle: {rebuilt_jsx}"
    );
}

#[test]
fn cli_should_roundtrip_run_vertical_alignment_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("run-alignment.tsx");
    let compiled = directory.path().join("run-alignment.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    let rebuilt = directory.path().join("rebuilt.docx");
    let rebuilt_external = directory.path().join("rebuilt-external.docx");
    let rebuilt_jsx = directory.path().join("rebuilt.jsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, Run } from "docx-jsx";
export default <Document><Section><Paragraph><Run verticalAlign="baseline">baseline</Run><Run verticalAlign="superscript">superscript</Run><Run verticalAlign="subscript">subscript</Run></Paragraph></Section></Document>;"#,
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(
        fs::read(&compiled).expect("DOCX should exist"),
    ))
    .expect("compiled DOCX should be ZIP");
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part should exist")
        .read_to_string(&mut document_xml)
        .expect("document XML should read");
    for alignment in ["baseline", "superscript", "subscript"] {
        assert!(
            document_xml.contains(&format!("<w:vertAlign w:val=\"{alignment}\"")),
            "writer should emit the {alignment} OOXML value: {document_xml}"
        );
    }
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    for alignment in ["baseline", "superscript", "subscript"] {
        assert!(
            jsx.contains(&format!("verticalAlign=\"{alignment}\"")),
            "external reverse should preserve {alignment}: {jsx}"
        );
    }

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    fs::write(
        &rebuilt_external,
        strip_ir_manifest(&fs::read(&rebuilt).expect("rebuilt DOCX should exist")),
    )
    .expect("rebuilt manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&rebuilt_external)
        .args(["-o"])
        .arg(&rebuilt_jsx)
        .assert()
        .success();
    let rebuilt_jsx = fs::read_to_string(&rebuilt_jsx).expect("second reverse JSX should exist");
    assert!(
        rebuilt_jsx.contains("verticalAlign")
            && rebuilt_jsx.contains("baseline")
            && rebuilt_jsx.contains("superscript")
            && rebuilt_jsx.contains("subscript"),
        "all explicit vertical alignment values should survive recompilation: {rebuilt_jsx}"
    );
}

#[test]
fn cli_should_roundtrip_run_underline_values_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("run-underline.tsx");
    let compiled = directory.path().join("run-underline.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    let tokens = [
        "single",
        "words",
        "double",
        "thick",
        "dotted",
        "dottedHeavy",
        "dash",
        "dashedHeavy",
        "dashLong",
        "dashLongHeavy",
        "dotDash",
        "dashDotHeavy",
        "dotDotDash",
        "dashDotDotHeavy",
        "wave",
        "wavyHeavy",
        "wavyDouble",
        "none",
    ];
    let runs = tokens.iter().fold(String::new(), |mut runs, token| {
        runs.push_str("<Run underline=\"");
        runs.push_str(token);
        runs.push_str("\">");
        runs.push_str(token);
        runs.push_str("</Run>");
        runs
    });
    fs::write(
        &input,
        format!(
            "import {{ Document, Section, Paragraph, Run }} from \"docx-jsx\";\nexport default <Document><Section><Paragraph><Run underline>legacy-on</Run><Run underline={{false}}>legacy-off</Run>{runs}</Paragraph></Section></Document>;"
        ),
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(
        fs::read(&compiled).expect("DOCX should exist"),
    ))
    .expect("compiled DOCX should be ZIP");
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part should exist")
        .read_to_string(&mut document_xml)
        .expect("document XML should read");
    for token in tokens {
        assert!(
            document_xml.contains(&format!("<w:u w:val=\"{token}\"")),
            "writer should emit underline token {token}: {document_xml}"
        );
    }
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    for token in tokens {
        assert!(
            jsx.contains(&format!("underline=\"{token}\"")),
            "external reverse should preserve underline token {token}: {jsx}"
        );
    }
    assert!(jsx.contains("underline=\"single\""));
    assert!(jsx.contains("underline=\"none\""));
}

#[test]
fn validate_should_explain_invalid_run_underline_style() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("invalid-underline.tsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, Run } from "docx-jsx";
export default <Document><Section><Paragraph><Run underline="dotted-heavy">text</Run></Paragraph></Section></Document>;"#,
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["validate"])
        .arg(&input)
        .assert()
        .failure()
        .stderr(predicate::str::contains("underline"))
        .stderr(predicate::str::contains("dottedHeavy"))
        .stderr(predicate::str::contains("suggestion"));
}

#[test]
fn cli_should_roundtrip_complex_script_run_properties_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("complex-script.tsx");
    let compiled = directory.path().join("complex-script.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    let rebuilt = directory.path().join("rebuilt.docx");
    let rebuilt_external = directory.path().join("rebuilt-external.docx");
    let reversed_again = directory.path().join("reversed-again.jsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, Run, List, ListItem } from "docx-jsx";
export default <Document styles={[{id:"ComplexScriptStyle",name:"Complex Script Style",type:"character",run:{size:10,sizeCs:14,bold:true,boldCs:false,italic:false,italicCs:true}}]}><Section><Paragraph><Run style="Emphasis" size={12} sizeCs={18} bold={false} boldCs italic italicCs={false}>mixed script</Run></Paragraph><Paragraph><Run style="ComplexScriptStyle">styled script</Run></Paragraph><List type="ordered" levels={[{format:"decimal",text:"%1.",size:11,sizeCs:16,bold:false,boldCs:true,italic:true,italicCs:false}]}><ListItem>numbered script</ListItem></List></Section></Document>;"#,
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    assert_complex_script_compiled_xml(&compiled);
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    for expected in [
        "size={12}",
        "sizeCs={18}",
        "bold={false}",
        "boldCs",
        "italic",
        "italicCs={false}",
        "style=\"Emphasis\"",
        "ComplexScriptStyle",
        "numbered script",
        "styled script",
        r#""sizeCs":14"#,
        r#""boldCs":false"#,
        r#""italicCs":true"#,
        r#""sizeCs":16"#,
        r#""boldCs":true"#,
        r#""italicCs":false"#,
    ] {
        assert!(
            jsx.contains(expected),
            "external reverse should preserve {expected}: {jsx}"
        );
    }
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    fs::write(
        &rebuilt_external,
        strip_ir_manifest(&fs::read(&rebuilt).expect("rebuilt DOCX should exist")),
    )
    .expect("rebuilt manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&rebuilt_external)
        .args(["-o"])
        .arg(&reversed_again)
        .assert()
        .success();
    let jsx_again = fs::read_to_string(&reversed_again).expect("second reverse should exist");
    for expected in [
        "sizeCs={18}",
        r#""sizeCs":14"#,
        r#""boldCs":false"#,
        r#""italicCs":true"#,
        r#""sizeCs":16"#,
        r#""boldCs":true"#,
        r#""italicCs":false"#,
    ] {
        assert!(
            jsx_again.contains(expected),
            "reverse/recompile should preserve {expected}: {jsx_again}"
        );
    }
}

fn assert_complex_script_compiled_xml(compiled: &std::path::Path) {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(
        fs::read(compiled).expect("DOCX should exist"),
    ))
    .expect("compiled DOCX should be ZIP");
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part should exist")
        .read_to_string(&mut document_xml)
        .expect("document XML should read");
    for expected in [
        r#"<w:rStyle w:val="Emphasis" />"#,
        r#"<w:b w:val="false" />"#,
        r"<w:bCs />",
        r"<w:i />",
        r#"<w:iCs w:val="0" />"#,
        r#"<w:sz w:val="24" />"#,
        r#"<w:szCs w:val="36" />"#,
    ] {
        assert!(
            document_xml.contains(expected),
            "compiled OOXML should contain {expected}: {document_xml}"
        );
    }
    let mut styles_xml = String::new();
    archive
        .by_name("word/styles.xml")
        .expect("styles part should exist")
        .read_to_string(&mut styles_xml)
        .expect("styles XML should read");
    assert!(styles_xml.contains(r#"<w:bCs w:val="0" />"#));
    assert!(styles_xml.contains(r"<w:iCs />"));
    let mut numbering_xml = String::new();
    archive
        .by_name("word/numbering.xml")
        .expect("numbering part should exist")
        .read_to_string(&mut numbering_xml)
        .expect("numbering XML should read");
    assert!(numbering_xml.contains(r"<w:bCs />"));
    assert!(numbering_xml.contains(r#"<w:iCs w:val="0" />"#));
}

#[test]
fn cli_should_roundtrip_run_caps_on_and_off_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("run-caps.tsx");
    let compiled = directory.path().join("run-caps.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    fs::write(
        &input,
        r#"import { AllCaps, Document, Section, Paragraph, Run } from "docx-jsx";
export default <Document><Section><Paragraph><Run caps={true}>caps on</Run><Run caps={false}>caps off</Run><AllCaps>semantic caps</AllCaps></Paragraph></Section></Document>;"#,
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(
        fs::read(&compiled).expect("DOCX should exist"),
    ))
    .expect("compiled DOCX should be ZIP");
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part should exist")
        .read_to_string(&mut document_xml)
        .expect("document XML should read");
    assert!(
        document_xml.contains("<w:caps w:val=\"true\"")
            && document_xml.contains("<w:caps w:val=\"false\""),
        "writer should preserve explicit caps on/off values: {document_xml}"
    );
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    assert!(
        jsx.contains("caps={true}") && jsx.contains("caps={false}"),
        "external reverse should preserve explicit caps on/off: {jsx}"
    );
}

#[test]
fn cli_should_roundtrip_run_hidden_properties_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("run-hidden.tsx");
    let compiled = directory.path().join("run-hidden.docx");
    let external = directory.path().join("external.docx");
    let reversed = directory.path().join("reversed.jsx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph, Run, HiddenText, SpecialHiddenText } from "docx-jsx";
export default <Document><Section><Paragraph><Run hidden={true} specVanish={true}>direct</Run><HiddenText>hidden wrapper</HiddenText><SpecialHiddenText>special hidden wrapper</SpecialHiddenText></Paragraph></Section></Document>;"#,
    )
    .expect("fixture should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("DOCX should exist")),
    )
    .expect("manifest-free DOCX should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    assert!(
        jsx.contains("hidden={true}") && jsx.contains("specVanish={true}"),
        "external reverse should preserve hidden and special-hidden run properties: {jsx}"
    );
}

#[test]
fn cli_should_reverse_style_line_number_suppression_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("style.tsx");
    let docx = directory.path().join("style.docx");
    let external = directory.path().join("external.docx");
    let jsx = directory.path().join("reverse.jsx");
    let rebuilt = directory.path().join("rebuilt.docx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph } from "docx-jsx";
export default <Document styles={[{id:"NoLines",name:"No Lines",type:"paragraph",paragraph:{suppressLineNumbers:true}},{id:"KeepLines",name:"Keep Lines",type:"paragraph",paragraph:{suppressLineNumbers:false}}]}><Section><Paragraph style="NoLines">excluded</Paragraph><Paragraph style="KeepLines">included</Paragraph></Section></Document>;"#,
    )
    .expect("input should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&docx)
        .assert()
        .success();
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&docx).expect("DOCX should exist")),
    )
    .expect("external DOCX should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&jsx)
        .assert()
        .success();
    let source = fs::read_to_string(&jsx).expect("reverse JSX should exist");
    assert!(source.contains("\"suppressLineNumbers\":true"), "{source}");
    assert!(source.contains("\"suppressLineNumbers\":false"), "{source}");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&jsx)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(
        fs::read(&rebuilt).expect("rebuilt DOCX should exist"),
    ))
    .expect("rebuilt DOCX should be ZIP");
    let mut styles = String::new();
    archive
        .by_name("word/styles.xml")
        .expect("styles part should exist")
        .read_to_string(&mut styles)
        .expect("styles XML should read");
    assert!(styles.contains("<w:suppressLineNumbers />"), "{styles}");
    assert!(
        styles.contains("<w:suppressLineNumbers w:val=\"0\" />"),
        "{styles}"
    );
}

#[test]
fn cli_should_reverse_custom_section_columns_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("columns.tsx");
    let docx = directory.path().join("columns.docx");
    let external = directory.path().join("external.docx");
    let jsx = directory.path().join("reverse.jsx");
    let rebuilt = directory.path().join("rebuilt.docx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph } from "docx-jsx";
export default <Document><Section columns={{count:2,spacing:18,separator:true}}><Paragraph>equal</Paragraph></Section><Section columns={{widths:[{width:144,spacing:24},{width:180}]}}><Paragraph>unequal</Paragraph></Section></Document>;"#,
    )
    .expect("input should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&docx)
        .assert()
        .success();
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&docx).expect("DOCX should exist")),
    )
    .expect("external DOCX should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&jsx)
        .assert()
        .success();
    let source = fs::read_to_string(&jsx).expect("reverse JSX should exist");
    assert!(source.contains("\"count\":2"), "{source}");
    assert!(source.contains("\"separator\":true"), "{source}");
    assert!(
        source.contains("\"widths\":[{\"width\":144,\"spacing\":24},{\"width\":180}]"),
        "{source}"
    );

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&jsx)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(
        fs::read(&rebuilt).expect("rebuilt DOCX should exist"),
    ))
    .expect("rebuilt DOCX should be ZIP");
    let mut document = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part should exist")
        .read_to_string(&mut document)
        .expect("document XML should read");
    assert!(
        document.contains(r#"w:space="360" w:num="2" w:sep="1""#),
        "{document}"
    );
    assert!(document.contains(r#"w:equalWidth="0""#), "{document}");
}

#[test]
fn cli_should_reverse_and_recompile_core_properties_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("metadata.tsx");
    let compiled = directory.path().join("metadata.docx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph } from "docx-jsx";
export default <Document creator="Ada & Co" lastModifiedBy="Lin" description="Weekly <report>" language="zh-CN" revision={7} subject="Economy" title="Status" documentProtection={{ edit: "readOnly", formatting: false, enforcement: true, hash: "YWJjZA==", salt: "AQIDBA==" }}><Section><Paragraph>body</Paragraph></Section></Document>;"#,
    )
    .expect("source should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    let external = directory.path().join("external.docx");
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("compiled DOCX should read")),
    )
    .expect("external DOCX should write");
    let reversed = directory.path().join("reversed.jsx");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let rebuilt = directory.path().join("rebuilt.docx");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should read");
    assert!(jsx.contains("creator=\"Ada &amp; Co\""), "{jsx}");
    let source = fs::read(&rebuilt).expect("rebuilt DOCX should read");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(source))
        .expect("rebuilt DOCX should remain a ZIP archive");
    let mut core = String::new();
    archive
        .by_name("docProps/core.xml")
        .expect("core properties should exist")
        .read_to_string(&mut core)
        .expect("core properties should be UTF-8");
    assert!(core.contains("Ada &amp; Co") && core.contains("Weekly &lt;report&gt;"));
    let mut settings = String::new();
    archive
        .by_name("word/settings.xml")
        .expect("settings part should exist")
        .read_to_string(&mut settings)
        .expect("settings should be UTF-8");
    assert!(
        settings.contains(r#"w:edit="readOnly""#)
            && settings.contains(r#"w:hash="YWJjZA==" w:salt="AQIDBA==""#),
        "{settings}"
    );
}

#[test]
fn cli_should_roundtrip_document_zoom_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("zoom.tsx");
    let compiled = directory.path().join("zoom.docx");
    fs::write(
        &input,
        r#"import { Document, Section, Paragraph } from "docx-jsx";
export default <Document zoom={125}><Section><Paragraph>body</Paragraph></Section></Document>;"#,
    )
    .expect("source should write");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&input)
        .args(["-o"])
        .arg(&compiled)
        .assert()
        .success();
    let external = directory.path().join("external.docx");
    fs::write(
        &external,
        strip_ir_manifest(&fs::read(&compiled).expect("compiled DOCX should read")),
    )
    .expect("external DOCX should write");
    let reversed = directory.path().join("reversed.jsx");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let jsx = fs::read_to_string(&reversed).expect("reverse JSX should read");
    assert!(jsx.contains("zoom={125}"), "{jsx}");
    let rebuilt = directory.path().join("rebuilt.docx");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let bytes = fs::read(rebuilt).expect("rebuilt DOCX should read");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes))
        .expect("rebuilt DOCX should remain a ZIP archive");
    let mut settings = String::new();
    archive
        .by_name("word/settings.xml")
        .expect("settings should exist")
        .read_to_string(&mut settings)
        .expect("settings should be UTF-8");
    assert!(
        settings.contains(r#"<w:zoom w:percent="125""#),
        "{settings}"
    );
}

#[test]
fn cli_should_roundtrip_document_revision_tracking_without_manifest() {
    let directory = tempdir().expect("tempdir should work");
    for enabled in [true, false] {
        let input = directory.path().join(format!("tracking-{enabled}.tsx"));
        let compiled = directory.path().join(format!("tracking-{enabled}.docx"));
        let moves = !enabled;
        let track_revisions = if enabled {
            "trackRevisions"
        } else {
            "trackRevisions={false}"
        };
        let do_not_track_moves = if moves {
            "doNotTrackMoves"
        } else {
            "doNotTrackMoves={false}"
        };
        let do_not_track_formatting = if enabled {
            "doNotTrackFormatting"
        } else {
            "doNotTrackFormatting={false}"
        };
        fs::write(
            &input,
            format!(
                r#"import {{ Document, Section, Paragraph }} from "docx-jsx";
export default <Document {track_revisions} {do_not_track_moves} {do_not_track_formatting}><Section><Paragraph>body</Paragraph></Section></Document>;"#
            ),
        )
        .expect("source should write");
        Command::cargo_bin("docx-jsx")
            .expect("binary should build")
            .arg(&input)
            .args(["-o"])
            .arg(&compiled)
            .assert()
            .success();
        let external = directory.path().join(format!("external-{enabled}.docx"));
        fs::write(
            &external,
            strip_ir_manifest(&fs::read(&compiled).expect("compiled DOCX should read")),
        )
        .expect("external DOCX should write");
        let reversed = directory.path().join(format!("reversed-{enabled}.jsx"));
        Command::cargo_bin("docx-jsx")
            .expect("binary should build")
            .args(["reverse"])
            .arg(&external)
            .args(["-o"])
            .arg(&reversed)
            .assert()
            .success();
        let jsx = fs::read_to_string(&reversed).expect("reverse JSX should read");
        for (name, value) in [
            ("trackRevisions", enabled),
            ("doNotTrackMoves", moves),
            ("doNotTrackFormatting", enabled),
        ] {
            let expected = if value {
                name.to_owned()
            } else {
                format!("{name}={{false}}")
            };
            assert!(jsx.contains(&expected), "missing {expected} from {jsx}");
        }
        let rebuilt = directory.path().join(format!("rebuilt-{enabled}.docx"));
        Command::cargo_bin("docx-jsx")
            .expect("binary should build")
            .arg(&reversed)
            .args(["-o"])
            .arg(&rebuilt)
            .assert()
            .success();
        let bytes = fs::read(rebuilt).expect("rebuilt DOCX should read");
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes))
            .expect("rebuilt DOCX should remain a ZIP archive");
        let mut settings = String::new();
        archive
            .by_name("word/settings.xml")
            .expect("settings should exist")
            .read_to_string(&mut settings)
            .expect("settings should be UTF-8");
        if enabled {
            assert!(
                settings.contains("<w:trackRevisions />")
                    && settings.contains(r#"<w:doNotTrackMoves w:val="0" />"#)
                    && settings.contains("<w:doNotTrackFormatting />"),
                "{settings}"
            );
        } else {
            assert!(
                settings.contains(r#"<w:trackRevisions w:val="0" />"#)
                    && settings.contains("<w:doNotTrackMoves />")
                    && settings.contains(r#"<w:doNotTrackFormatting w:val="0" />"#),
                "{settings}"
            );
        }
    }
}

#[test]
fn cli_should_reverse_and_rebuild_external_simple_position_image() {
    let directory = tempdir().expect("tempdir should work");
    let mut png = std::io::Cursor::new(Vec::new());
    image::DynamicImage::new_rgba8(1, 1)
        .write_to(&mut png, image::ImageFormat::Png)
        .expect("PNG should encode");
    let mut picture = docx_rs::Pic::new_with_dimensions(png.into_inner(), 1, 1).floating();
    picture.simple_pos = true;
    picture.simple_pos_x = 25_400;
    picture.simple_pos_y = -12_700;
    let mut package = std::io::Cursor::new(Vec::new());
    docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(picture)))
        .build()
        .pack(&mut package)
        .expect("external DOCX should pack");
    let package = package.into_inner();
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(package))
        .expect("external DOCX should be a ZIP archive");
    let mut rewritten = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("DOCX part should read");
        if entry.is_dir() {
            rewritten
                .raw_copy_file(entry)
                .expect("directory should copy");
            continue;
        }
        let name = entry.name().to_owned();
        let options = entry.options();
        let mut bytes = Vec::new();
        entry
            .read_to_end(&mut bytes)
            .expect("DOCX part should read");
        if name == "word/document.xml" {
            let xml = String::from_utf8(bytes).expect("document XML should be UTF-8");
            bytes = xml
                .replacen("simplePos=\"0\"", "simplePos=\"1\"", 1)
                .replacen("allowOverlap=\"1\"", "allowOverlap=\"0\"", 1)
                .into_bytes();
        }
        rewritten
            .start_file(name, options)
            .expect("DOCX part should write");
        rewritten.write_all(&bytes).expect("DOCX part should write");
    }
    let external_docx = directory.path().join("external.docx");
    fs::write(
        &external_docx,
        rewritten
            .finish()
            .expect("DOCX should finalize")
            .into_inner(),
    )
    .expect("external DOCX should write");
    let reversed = directory.path().join("reversed.jsx");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external_docx)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    assert!(
        source.contains("simplePosition={{ x: 25400, y: -12700 }}"),
        "{source}"
    );
    let rebuilt = directory.path().join("rebuilt.docx");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let bytes = fs::read(&rebuilt).expect("rebuilt DOCX should exist");
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes))
        .expect("rebuilt DOCX should be valid ZIP");
    let mut document = String::new();
    archive
        .by_name("word/document.xml")
        .expect("document part should exist")
        .read_to_string(&mut document)
        .expect("document XML should read");
    assert!(
        document.contains("simplePos=\"1\"")
            && document.contains("allowOverlap=\"0\"")
            && document.contains("<wp:simplePos x=\"25400\" y=\"-12700\" />"),
        "{document}"
    );
}

#[test]
fn cli_should_reverse_and_rebuild_numbering_level_overrides() {
    let directory = tempdir().expect("tempdir should work");
    let replacement = docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("lowerLetter"),
        docx_rs::LevelText::new("%1)"),
        docx_rs::LevelJc::new("left"),
    );
    let mut abstract_numbering = docx_rs::AbstractNumbering::new(9)
        .add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("%1."),
            docx_rs::LevelJc::new("left"),
        ))
        .add_level(docx_rs::Level::new(
            1,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("lowerLetter"),
            docx_rs::LevelText::new("%2."),
            docx_rs::LevelJc::new("left"),
        ));
    abstract_numbering.multi_level_type = Some("hybridMultilevel".to_owned());
    let mut package = std::io::Cursor::new(Vec::new());
    docx_rs::Docx::new()
        .add_abstract_numbering(abstract_numbering)
        .add_numbering(
            docx_rs::Numbering::new(9, 9)
                .add_override(docx_rs::LevelOverride::new(0).level(replacement))
                .add_override(docx_rs::LevelOverride::new(1).start(4)),
        )
        .add_paragraph(
            docx_rs::Paragraph::new()
                .numbering(docx_rs::NumberingId::new(9), docx_rs::IndentLevel::new(0))
                .add_run(docx_rs::Run::new().add_text("item")),
        )
        .build()
        .pack(&mut package)
        .expect("external DOCX should pack");
    let external_docx = directory.path().join("external.docx");
    fs::write(&external_docx, package.into_inner()).expect("external DOCX should write");

    let reversed = directory.path().join("reversed.jsx");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["reverse"])
        .arg(&external_docx)
        .args(["-o"])
        .arg(&reversed)
        .assert()
        .success();
    let source = fs::read_to_string(&reversed).expect("reverse JSX should exist");
    assert!(
        source.contains(r#"overrides={[{"level":0,"definition":{"format":"lowerLetter""#)
            && source.contains(r#"{"level":1,"start":4}"#)
            && source.contains(r#"multiLevelType="hybridMultilevel""#),
        "{source}"
    );

    let rebuilt = directory.path().join("rebuilt.docx");
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .arg(&reversed)
        .args(["-o"])
        .arg(&rebuilt)
        .assert()
        .success();
    let bytes = fs::read(&rebuilt).expect("rebuilt DOCX should exist");
    let mut archive =
        zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("rebuilt DOCX should be a ZIP");
    let mut numbering = String::new();
    archive
        .by_name("word/numbering.xml")
        .expect("numbering part should exist")
        .read_to_string(&mut numbering)
        .expect("numbering XML should read");
    assert!(
        numbering.contains(r#"<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0">"#)
            && numbering.contains(r#"<w:lvlOverride w:ilvl="1"><w:startOverride w:val="4" />"#)
            && numbering.contains(r#"<w:multiLevelType w:val="hybridMultilevel" />"#)
            && numbering.contains(r#"w:val="lowerLetter""#),
        "{numbering}"
    );
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
fn cli_should_infer_pdf_output_from_extension() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("report.jsx");
    let output = directory.path().join("report.pdf");
    fs::write(
        &input,
        r#"import { Document, Section } from "docx-jsx"; export default <Document><Section /></Document>;"#,
    )
    .expect("input should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .env_remove("DOCX_JSX_SOFFICE")
        .env_remove("DOCX_JSX_WORD")
        .env_remove("DOCX_JSX_WPS")
        .args([
            input.to_str().expect("UTF-8 path"),
            "-o",
            output.to_str().expect("UTF-8 path"),
            "--soffice",
            "/no/such/soffice",
        ])
        .assert()
        .failure()
        .stderr(predicate::str::contains("explanation:"))
        .stderr(predicate::str::contains("DOCX_JSX_SOFFICE"))
        .stderr(predicate::str::contains("LibreOffice"));
    assert!(!output.exists(), "failed PDF compile must not write output");
    assert!(
        !directory.path().join("report.docx").exists(),
        "PDF compile must not leave a sibling DOCX"
    );
}

#[test]
fn cli_should_reject_unknown_compile_output_extension() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("report.jsx");
    let output = directory.path().join("report.txt");
    fs::write(
        &input,
        r#"import { Document, Section } from "docx-jsx"; export default <Document><Section /></Document>;"#,
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
        .failure()
        .stderr(predicate::str::contains(".docx"))
        .stderr(predicate::str::contains(".pdf"))
        .stderr(predicate::str::contains("suffix"));
    assert!(!output.exists());
}

#[test]
fn cli_should_reject_pdf_flags_when_writing_docx() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("report.jsx");
    let output = directory.path().join("report.docx");
    fs::write(
        &input,
        r#"import { Document, Section } from "docx-jsx"; export default <Document><Section /></Document>;"#,
    )
    .expect("input should write");

    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args([
            input.to_str().expect("UTF-8 path"),
            "-o",
            output.to_str().expect("UTF-8 path"),
            "--engine",
            "word",
        ])
        .assert()
        .failure()
        .stderr(predicate::str::contains("--engine"))
        .stderr(predicate::str::contains(".pdf"));
    assert!(!output.exists());
}

#[test]
fn pdf_should_explain_how_to_install_and_configure_soffice() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("source.jsx");
    let docx = directory.path().join("source.docx");
    fs::write(
        &input,
        r#"import { Document, Section } from "docx-jsx"; export default <Document><Section /></Document>;"#,
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
        .env_remove("DOCX_JSX_SOFFICE")
        .args([
            "pdf",
            docx.to_str().expect("UTF-8 path"),
            "--soffice",
            "/no/such/soffice",
        ])
        .assert()
        .failure()
        .stderr(predicate::str::contains("explanation:"))
        .stderr(predicate::str::contains("LibreOffice"))
        .stderr(predicate::str::contains("DOCX_JSX_SOFFICE"))
        .stderr(predicate::str::contains("--soffice"))
        .stderr(predicate::str::contains("libreoffice.org"))
        .stderr(predicate::str::contains("WPS"))
        .stderr(predicate::str::contains("DOCX_JSX_WORD"));
    assert!(!docx.with_extension("pdf").exists());
}

#[test]
fn pdf_should_reject_missing_word_engine() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("source.jsx");
    let docx = directory.path().join("source.docx");
    fs::write(
        &input,
        r#"import { Document, Section } from "docx-jsx"; export default <Document><Section /></Document>;"#,
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
        .env_remove("DOCX_JSX_WORD")
        .env_remove("DOCX_JSX_SOFFICE")
        .args([
            "pdf",
            docx.to_str().expect("UTF-8 path"),
            "--engine",
            "word",
        ])
        .assert()
        .failure()
        .stderr(predicate::str::contains("word"))
        .stderr(predicate::str::contains("DOCX_JSX_WORD"));
}

#[test]
fn pdf_should_reject_non_docx_input() {
    Command::cargo_bin("docx-jsx")
        .expect("binary should build")
        .args(["pdf", "notes.txt"])
        .assert()
        .failure()
        .stderr(predicate::str::contains(".docx"));
}

#[test]
#[allow(clippy::too_many_lines)]
fn cli_should_compile_advanced_document_components() {
    let directory = tempdir().expect("tempdir should work");
    let input = directory.path().join("advanced.tsx");
    let output = directory.path().join("advanced.docx");
    let reversed = directory.path().join("advanced.reversed.jsx");
    let recompiled = directory.path().join("advanced.recompiled.docx");
    image::DynamicImage::new_rgba8(1, 1)
        .save(directory.path().join("pixel.png"))
        .expect("image fixture should write");
    fs::write(
        &input,
        r##"import { Document, Section, Header, Footer, Heading, Caption, Index, IndexEntry, Paragraph, Hyperlink, PageNumber, TotalPages, List, ListItem, Run, Image, Bookmark, Table, TableRow, TableCell, TableOfContents, TableOfFigures, TableOfEntries, TocItem, TocBefore, TocAfter, TocEntry, Comment, Footnote, Tab, TabStop, CarriageReturn, NonBreakingSpace, SoftHyphen, NonBreakingHyphen, Symbol, Bold, Italic, Underline, StrikeThrough, Superscript, Subscript, AllCaps, HiddenText, SpecialHiddenText, DoubleStrike, SpacedText, ScaledText, FitText, BorderedText, ShadedText, Inserted, Deleted, DeletedInstruction, MovedFrom, MovedTo, PageReference, PositionalTab, ContentControl, Field, DateField, TimeField, FileNameField, AuthorField, TitleField, SubjectField, SequenceField, ReferenceField, MergeField, DocumentPropertyField, FormulaField } from "docx-jsx";
export default <Document defaultCharacterSpacing={0.5} defaultLineSpacing={{before: 2, after: 6, line: 14, beforeLines: 100, afterLines: 200, lineRule: "atLeast"}} createdAt="2026-08-14T00:00:00Z" updatedAt="2026-08-15T00:00:00Z" customProperties={{Project: "Apollo"}} documentId="01234567-89AB-CDEF-0123-456789ABCDEF" defaultTabStop={36} documentVariables={{Customer: "Ada"}} evenAndOddHeaders adjustLineHeightInTable characterSpacingControl="compressPunctuation" webExtensions={[{id: "7f33b723-fb58-4524-8733-dbedc4b7c095", referenceId: "office-addin", version: "1.0.0.0", store: "developer", storeType: "Registry", properties: {mode: "review"}}, {id: "11111111-2222-3333-4444-555555555555", referenceId: "second", version: "2.0", store: "OMEX", storeType: "Marketplace"}]} customXmlItems={[{id: "06AC5857-5C65-A94A-BCEC-37356A209BC3", xml: "<customer><name>Ada</name></customer>"}, {id: "11111111-AAAA-BBBB-CCCC-222222222222", xml: "<order id=\"42\"/>"}]} styles={[{id: "BodyBase", name: "Body Base", type: "paragraph", paragraph: {spacingAfter: 6}}, {id: "Body", name: "Body", type: "paragraph", basedOn: "BodyBase", next: "Body", link: "BodyChar", run: {font: "Noto Sans CJK SC", size: 12}}, {id: "BodyChar", name: "Body Char", type: "character", link: "Body", run: {italic: true}}, {id: "ReportTitle", name: "Report Title", type: "paragraph", basedOn: "Normal", quickFormat: true, run: {font: "Noto Sans CJK SC", size: 18, color: "336699", bold: true, textBorder: {style: "double", size: 1, color: "336699", space: 2}}, paragraph: {align: "center", spacingAfter: 12, spacingBeforeLines: 50, spacingAfterLines: 75, lineRule: "exact", outlineLevel: 1}}, {id: "ReportTable", name: "Report Table", type: "table", table: {align: "center", margins: {top: 1, right: 2, bottom: 3, left: 4}, border: {style: "double", size: 1, color: "336699"}}, cell: {verticalAlign: "center", shading: "FFF2CC", border: {style: "dotted", size: 0.5, color: "993366"}}}]}><Section titlePage textDirection="tbRl" documentGrid={{type: "linesAndChars", linePitch: 18, charSpace: -10}} pageNumbering={{start: 3, chapterStyle: "1"}}>
  <Header type="first"><Paragraph>Report header <Image src="pixel.png" width={12} height={12} /></Paragraph></Header>
  <Footer type="even"><Paragraph>Page <PageNumber /> of <TotalPages /></Paragraph></Footer>
  <TableOfContents startLevel={1} endLevel={2} alias="Contents"><TocBefore><Paragraph>Contents introduction</Paragraph></TocBefore><TocItem text="Introduction" bookmark="intro" /><TocAfter><Paragraph>Contents note</Paragraph></TocAfter></TableOfContents>
  <TableOfFigures label="Figure" alias="Figures"><TocItem text="Architecture" bookmark="intro" level={1} /></TableOfFigures>
  <TableOfEntries identifier="manual" alias="Manual entries"><TocItem text="Rust" bookmark="intro" pageRef="4" /></TableOfEntries>
  <Paragraph style="Body"><Hyperlink href="https://example.com" history><Run style="BodyChar" underline color="2E74B5" themeColor="accent1" themeShade="BF" bold={false} italic={false} strike={false} doubleStrike>Example</Run></Hyperlink></Paragraph>
  <Paragraph><Image src="pixel.png" width={48} height={32} rotate={45} floating allowOverlap positionH="right" positionV={18.5} relativeFromH="page" relativeFromV="paragraph" distanceTop={2} distanceBottom={3} distanceLeft={4} distanceRight={5} relativeHeight={251658240} /></Paragraph>
  <Bookmark name="intro"><Heading level={1} font="Noto Sans CJK SC" size={16}>Introduction</Heading><Paragraph style="Quote" snapToGrid={false} widowControl font="Noto Sans CJK SC" size={12} bold italic={false} color="1a2B3c" characterSpacing={0.5}>Bookmarked text</Paragraph></Bookmark>
  <Paragraph><Hyperlink anchor="intro">Jump to introduction</Hyperlink></Paragraph>
  <Paragraph><TabStop position={72} align="right" leader="dot" /><TocEntry text="Manual entry" level={2} identifier="manual" />Introduction is on page <PageReference bookmark="intro" placeholder="?" relativePosition />.<Tab /><ContentControl alias="Customer" xpath="/root/customer" storeItemId="{06AC5857-5C65-A94A-BCEC-37356A209BC3}">Ada</ContentControl><CarriageReturn /><MovedFrom author="Ada">old</MovedFrom><MovedTo author="Ada"><Run bold>new</Run></MovedTo><PositionalTab align="right" relativeTo="margin" leader="dot" /></Paragraph>
  <Paragraph><Comment text="Please verify" author="Ada">Reviewed text</Comment><Tab /><Symbol font="Wingdings" char="F0A7" /><Footnote>Footnote content</Footnote><Deleted><Run><DeletedInstruction instruction={'HYPERLINK "https://example.com"'} /></Run></Deleted></Paragraph>
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
  <Paragraph paragraphId="A1B2C3D4" spacingBeforeLines={125} spacingAfterLines={250} lineRule="auto" bidi textAlign="baseline" adjustRightIndent={1} shading="DDEEFF" outlineLevel={3} frame={{wrap: "around", horizontalAnchor: "page", x: 12, yAlign: "top", width: 144, height: 72}} border={{top: {style: "double", size: 1, color: "336699", space: 2}, between: false, bar: {style: "babyRattle"}}} inserted={{author: "Ada", date: "2026-08-14T00:00:00Z"}} propertyChange={{author: "Lin", previous: {align: "right", spacingAfter: 6, bidi: false}}}>Revised paragraph formatting</Paragraph>
  <List type="ordered" start={3}><ListItem>Third</ListItem><ListItem level={1}>Nested</ListItem></List>
  <Table style="ReportTable" indent={12} margins={{top: 2, right: 3, bottom: 4, left: 5}} border={{top: {style: "double", size: 1, color: "336699"}, insideHorizontal: false}} position={{leftFromText: 7.1, rightFromText: 7.1, verticalAnchor: "text", horizontalAnchor: "margin", xAlign: "right", y: 25.5}}><TableRow inserted={{author: "Ada", date: "2026-08-14T00:00:00Z"}}><TableCell verticalMerge="restart" textDirection="tbRl" margins={{top: 1, right: 2, bottom: 3, left: 4}} border={{left: false, topLeftToBottomRight: {style: "dotted", size: 0.5, color: "993366"}}}><Paragraph>Cell</Paragraph><TableOfContents startLevel={1} endLevel={2} /><ContentControl alias="CellValue">value</ContentControl></TableCell></TableRow><TableRow deleted={{author: "Linus"}}><TableCell><Paragraph>Removed row</Paragraph></TableCell></TableRow></Table>
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
    assert!(
        read_part(&header_name).contains("Report header")
            && read_part(&header_name).contains("a:blip")
    );
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
    assert!(document.contains("w:framePr") && document.contains("w:pPrChange"));
    assert!(document.contains("w:textAlignment") && document.contains("w:adjustRightInd"));
    assert!(document.contains("w14:paraId=\"A1B2C3D4\"") && document.contains("w:pBdr"));
    assert!(document.contains("w:delText") && document.contains("Old text"));
    assert!(
        document.contains(r#"<wp:anchor distT="25400" distB="38100" distL="50800" distR="63500""#)
            && document.contains(r#"allowOverlap="1""#)
            && document.contains(r#"<a:xfrm rot="2700000">"#)
    );
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
    assert!(read_part("word/webextensions/webextension1.xml").contains("office-addin"));
    assert!(read_part("word/webextensions/webextension2.xml").contains("second"));
    assert!(read_part("customXml/item1.xml").contains("<name>Ada</name>"));
    assert!(read_part("customXml/item2.xml").contains("id=\"42\""));

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
