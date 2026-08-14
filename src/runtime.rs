use std::borrow::Cow;
use std::path::{Path, PathBuf};
use std::rc::Rc;

use deno_ast::{
    DecoratorsTranspileOption, EmitOptions, ImportsNotUsedAsValues, JsxAutomaticOptions,
    JsxRuntime, MediaType, ParseParams, SourceMapOption, TranspileModuleOptions, TranspileOptions,
};
use deno_core::error::ModuleLoaderError;
use deno_core::{
    JsRuntime, ModuleLoadOptions, ModuleLoadReferrer, ModuleLoadResponse, ModuleLoader,
    ModuleResolveResponse, ModuleSource, ModuleSourceCode, ModuleSpecifier, ModuleType,
    PollEventLoopOptions, ResolutionKind, RuntimeOptions, resolve_import, v8,
};
use deno_error::JsErrorBox;
use serde_json::Value;

use crate::error::{Error, Result};
use crate::ir::IrEnvelope;

const JSX_RUNTIME_SOURCE: &str = include_str!("../runtime/jsx-runtime.js");
const MAIN_MODULE_SOURCE: &str = include_str!("../runtime/main.js");
const EXTENSIONS: [&str; 5] = ["ts", "tsx", "js", "jsx", "json"];

#[derive(Debug, Default)]
struct DocxModuleLoader;

impl ModuleLoader for DocxModuleLoader {
    fn resolve(
        &self,
        specifier: &str,
        referrer: &str,
        _kind: ResolutionKind,
    ) -> ModuleResolveResponse {
        match specifier {
            "docx-jsx" => return parse_builtin("docx-jsx:main"),
            "docx-jsx:entry" => return parse_builtin("docx-jsx:entry"),
            "docx-jsx/jsx-runtime" | "docx-jsx/jsx-dev-runtime" => {
                return parse_builtin("docx-jsx:jsx-runtime");
            }
            _ => {}
        }
        if is_bare_specifier(specifier) {
            return Err(JsErrorBox::type_error(format!(
                "Bare module `{specifier}` is not supported; only docx-jsx and local files are allowed"
            )));
        }
        let resolved = resolve_import(specifier, referrer).map_err(JsErrorBox::from_err)?;
        resolve_local_extension(&resolved)
    }

    fn load(
        &self,
        module_specifier: &ModuleSpecifier,
        _maybe_referrer: Option<&ModuleLoadReferrer>,
        _options: ModuleLoadOptions,
    ) -> ModuleLoadResponse {
        let result = load_module(module_specifier);
        ModuleLoadResponse::Sync(result)
    }

    fn get_source_map(&self, _specifier: &str) -> Option<Cow<'_, [u8]>> {
        None
    }
}

fn parse_builtin(specifier: &str) -> ModuleResolveResponse {
    ModuleSpecifier::parse(specifier).map_err(JsErrorBox::from_err)
}

fn is_bare_specifier(specifier: &str) -> bool {
    !specifier.starts_with('.')
        && !specifier.starts_with('/')
        && !specifier.starts_with("file:")
        && !specifier.contains(':')
}

fn resolve_local_extension(specifier: &ModuleSpecifier) -> ModuleResolveResponse {
    if specifier.scheme() != "file" {
        return Err(JsErrorBox::type_error(format!(
            "Unsupported module scheme `{}`",
            specifier.scheme()
        )));
    }
    let path = specifier
        .to_file_path()
        .map_err(|()| JsErrorBox::type_error("Invalid file module URL"))?;
    let resolved = resolve_candidate(&path)
        .ok_or_else(|| JsErrorBox::type_error(format!("Module not found: {}", path.display())))?;
    ModuleSpecifier::from_file_path(&resolved).map_err(|()| {
        JsErrorBox::type_error(format!("Invalid module path: {}", resolved.display()))
    })
}

fn resolve_candidate(path: &Path) -> Option<PathBuf> {
    if path.is_file() {
        return Some(path.to_path_buf());
    }
    if path.extension().is_none() {
        for extension in EXTENSIONS {
            let candidate = path.with_extension(extension);
            if candidate.is_file() {
                return Some(candidate);
            }
        }
    }
    if path.is_dir() {
        for extension in EXTENSIONS {
            let candidate = path.join(format!("index.{extension}"));
            if candidate.is_file() {
                return Some(candidate);
            }
        }
    }
    None
}

fn load_module(
    specifier: &ModuleSpecifier,
) -> std::result::Result<ModuleSource, ModuleLoaderError> {
    let source = match specifier.as_str() {
        "docx-jsx:main" => MAIN_MODULE_SOURCE.to_owned(),
        "docx-jsx:jsx-runtime" => JSX_RUNTIME_SOURCE.to_owned(),
        _ => load_file_module(specifier)?,
    };
    Ok(ModuleSource::new(
        ModuleType::JavaScript,
        ModuleSourceCode::String(source.into()),
        specifier,
        None,
    ))
}

fn load_file_module(specifier: &ModuleSpecifier) -> std::result::Result<String, ModuleLoaderError> {
    let path = specifier
        .to_file_path()
        .map_err(|()| JsErrorBox::type_error("Only local file modules are supported"))?;
    let source = std::fs::read_to_string(&path).map_err(JsErrorBox::from_err)?;
    if path.extension().and_then(|value| value.to_str()) == Some("json") {
        serde_json::from_str::<Value>(&source).map_err(JsErrorBox::from_err)?;
        return Ok(format!("export default {source};"));
    }
    let media_type = MediaType::from_path(&path);
    if matches!(media_type, MediaType::JavaScript | MediaType::Mjs) {
        return Ok(source);
    }
    transpile(specifier, source, media_type).map_err(|error| JsErrorBox::generic(error.to_string()))
}

fn transpile(specifier: &ModuleSpecifier, source: String, media_type: MediaType) -> Result<String> {
    let parsed = deno_ast::parse_module(ParseParams {
        specifier: specifier.clone(),
        text: source.into(),
        media_type,
        capture_tokens: false,
        scope_analysis: false,
        maybe_syntax: None,
    })
    .map_err(|error| Error::Transpile(error.to_string()))?;
    let emitted = parsed
        .transpile(
            &TranspileOptions {
                imports_not_used_as_values: ImportsNotUsedAsValues::Remove,
                decorators: DecoratorsTranspileOption::Ecma,
                jsx: Some(JsxRuntime::Automatic(JsxAutomaticOptions {
                    development: false,
                    import_source: Some("docx-jsx".to_owned()),
                })),
                ..Default::default()
            },
            &TranspileModuleOptions::default(),
            &EmitOptions {
                source_map: SourceMapOption::Inline,
                inline_sources: true,
                ..Default::default()
            },
        )
        .map_err(|error| Error::Transpile(error.to_string()))?
        .into_source();
    Ok(emitted.text)
}

/// Evaluates an entry JSX/TSX module and returns normalized IR v1.
///
/// # Errors
///
/// Returns an error when the module graph cannot be loaded or transpiled,
/// JavaScript evaluation fails, or the default export cannot be decoded as IR.
pub async fn evaluate_entry(entry: &Path, data: Option<&Value>) -> Result<IrEnvelope> {
    let canonical_entry = entry
        .canonicalize()
        .map_err(|error| Error::Input(format!("cannot open {}: {error}", entry.display())))?;
    let entry_url = ModuleSpecifier::from_file_path(&canonical_entry)
        .map_err(|()| Error::Input(format!("invalid entry path: {}", canonical_entry.display())))?;
    let wrapper_url = ModuleSpecifier::parse("docx-jsx:entry")
        .map_err(|error| Error::Runtime(format!("cannot create wrapper module URL: {error}")))?;
    let data_source = data.map_or_else(|| "undefined".to_owned(), Value::to_string);
    let entry_source = serde_json::to_string(entry_url.as_str())
        .map_err(|error| Error::Runtime(error.to_string()))?;
    let wrapper = format!(
        "import entry from {entry_source};\nimport {{ finalize }} from \"docx-jsx/jsx-runtime\";\nexport default await finalize(entry, {data_source});"
    );
    let mut runtime = JsRuntime::new(RuntimeOptions {
        module_loader: Some(Rc::new(DocxModuleLoader)),
        ..Default::default()
    });
    runtime
        .execute_script(
            "docx-jsx:bootstrap",
            r#"const __docxFormat = values => values.map(value => {
  if (typeof value === "string") return value;
  try { return JSON.stringify(value); } catch { return String(value); }
}).join(" ");
globalThis.console ??= {
  log(...values) { Deno.core.print(__docxFormat(values) + "\n", true); },
  info(...values) { Deno.core.print(__docxFormat(values) + "\n", true); },
  warn(...values) { Deno.core.print(__docxFormat(values) + "\n", true); },
  error(...values) { Deno.core.print(__docxFormat(values) + "\n", true); }
};"#,
        )
        .map_err(|error| Error::Runtime(error.to_string()))?;
    let module_id = runtime
        .load_main_es_module_from_code(&wrapper_url, wrapper)
        .await
        .map_err(|error| Error::Module(error.to_string()))?;
    let evaluation = runtime.mod_evaluate(module_id);
    runtime
        .run_event_loop(PollEventLoopOptions::default())
        .await
        .map_err(|error| Error::Runtime(error.to_string()))?;
    evaluation
        .await
        .map_err(|error| Error::Runtime(error.to_string()))?;
    let namespace = runtime
        .get_module_namespace(module_id)
        .map_err(|error| Error::Runtime(error.to_string()))?;
    let value = {
        deno_core::scope!(scope, runtime);
        let namespace = v8::Local::<v8::Object>::new(scope, namespace);
        let key = v8::String::new(scope, "default")
            .ok_or_else(|| Error::Runtime("cannot allocate export key".to_owned()))?;
        let export = namespace
            .get(scope, key.into())
            .ok_or_else(|| Error::Runtime("entry wrapper has no default export".to_owned()))?;
        deno_core::serde_v8::from_v8::<Value>(scope, export)
            .map_err(|error| Error::Runtime(format!("cannot decode JSX result: {error}")))?
    };
    let ir: IrEnvelope = serde_json::from_value(value)
        .map_err(|error| Error::Runtime(format!("invalid IR from JSX runtime: {error}")))?;
    ir.validate()?;
    Ok(ir)
}

#[cfg(test)]
mod tests {
    use std::fs;

    use tempfile::tempdir;

    use super::*;

    #[tokio::test(flavor = "current_thread")]
    async fn evaluate_should_execute_tsx_components_and_data() {
        let directory = tempdir().expect("tempdir should work");
        let entry = directory.path().join("report.tsx");
        fs::write(
            &entry,
            r#"import { Document, Section, Paragraph, Run } from "docx-jsx";
type Data = { names: string[] };
const Item = ({ name }: { name: string }) => <Paragraph><Run bold>{name}</Run></Paragraph>;
export default (data: Data) => <Document><Section>{data.names.map(name => <Item name={name} />)}</Section></Document>;"#,
        )
        .expect("fixture should write");
        let data = serde_json::json!({"names": ["Ada", "Linus"]});
        let ir = evaluate_entry(&entry, Some(&data))
            .await
            .expect("evaluation should work");
        assert_eq!(ir.document.children.len(), 1);
    }

    #[tokio::test(flavor = "current_thread")]
    async fn evaluate_should_load_extensionless_local_module() {
        let directory = tempdir().expect("tempdir should work");
        fs::write(
            directory.path().join("component.tsx"),
            r#"import { Paragraph } from "docx-jsx"; export const Greeting = () => <Paragraph>Hello</Paragraph>;"#,
        )
        .expect("component should write");
        let entry = directory.path().join("report.tsx");
        fs::write(
            &entry,
            r#"import { Document, Section } from "docx-jsx"; import { Greeting } from "./component"; export default <Document><Section><Greeting /></Section></Document>;"#,
        )
        .expect("entry should write");
        assert!(evaluate_entry(&entry, None).await.is_ok());
    }
}
