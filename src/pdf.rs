//! Optional DOCX to PDF conversion through a local office application.

use std::env;
use std::ffi::OsStr;
use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use crate::{Error, Result};

const ENV_SOFFICE: &str = "DOCX_JSX_SOFFICE";
const ENV_WORD: &str = "DOCX_JSX_WORD";
const ENV_WPS: &str = "DOCX_JSX_WPS";
const ENV_ENGINE: &str = "DOCX_JSX_PDF_ENGINE";
const CONVERT_TIMEOUT: Duration = Duration::from_mins(3);

/// Which local office application should produce the PDF.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PdfEngine {
    /// Prefer Microsoft Word, then WPS, then `LibreOffice` (order depends on OS).
    #[default]
    Auto,
    /// `LibreOffice` `soffice`.
    LibreOffice,
    /// Microsoft Word automation.
    Word,
    /// WPS Writer automation.
    Wps,
}

/// Options for [`convert_docx_to_pdf_with`].
#[derive(Debug, Clone, Copy)]
pub struct PdfOptions<'a> {
    /// Selected engine. `Auto` still honors an explicit `LibreOffice` path.
    pub engine: PdfEngine,
    /// Explicit `soffice` path; when set, `LibreOffice` is used.
    pub soffice: Option<&'a Path>,
    /// Replace an existing PDF.
    pub force: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BackendKind {
    LibreOffice,
    Word,
    Wps,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Backend {
    kind: BackendKind,
    path: PathBuf,
}

/// Resolves a `LibreOffice` `soffice` binary.
///
/// # Errors
///
/// Returns [`Error::Pdf`] when no usable binary is found.
pub fn resolve_soffice(explicit: Option<&Path>) -> Result<PathBuf> {
    resolve_soffice_from(
        explicit,
        env::var_os(ENV_SOFFICE).map(PathBuf::from).as_deref(),
    )
}

/// Converts in-memory DOCX bytes to `output` without leaving a sibling archive.
///
/// # Errors
///
/// Returns [`Error::Pdf`] when no office app is available or conversion fails,
/// and [`Error::Output`] when the destination cannot be written.
pub fn convert_docx_bytes_to_pdf(
    bytes: &[u8],
    output: &Path,
    options: PdfOptions<'_>,
) -> Result<()> {
    let work = unique_work_dir()?;
    let input = work.join("source.docx");
    let result = (|| {
        fs::write(&input, bytes).map_err(|source| Error::Output {
            path: input.clone(),
            source,
        })?;
        convert_docx_to_pdf_with(&input, output, options)
    })();
    let _ = fs::remove_dir_all(&work);
    result
}

/// Converts `input` to `output` using automatic engine detection.
///
/// # Errors
///
/// Returns [`Error::Pdf`] when no office app is available or conversion fails.
pub fn convert_docx_to_pdf(
    input: &Path,
    output: &Path,
    soffice: Option<&Path>,
    force: bool,
) -> Result<()> {
    convert_docx_to_pdf_with(
        input,
        output,
        PdfOptions {
            engine: PdfEngine::Auto,
            soffice,
            force,
        },
    )
}

/// Converts `input` to `output` using the selected office engine.
///
/// # Errors
///
/// Returns [`Error::Pdf`] when the engine is missing or conversion fails, and
/// [`Error::Output`] when the destination cannot be written.
pub fn convert_docx_to_pdf_with(
    input: &Path,
    output: &Path,
    options: PdfOptions<'_>,
) -> Result<()> {
    if input.extension().and_then(OsStr::to_str) != Some("docx") {
        return Err(Error::Input(
            "pdf input must have a .docx extension".to_owned(),
        ));
    }
    if !input.is_file() {
        return Err(Error::Resource {
            path: input.to_path_buf(),
            source: std::io::Error::new(std::io::ErrorKind::NotFound, "DOCX file not found"),
        });
    }
    if output.exists() && !options.force {
        return Err(Error::Output {
            path: output.to_path_buf(),
            source: std::io::Error::new(
                std::io::ErrorKind::AlreadyExists,
                "output exists; pass --force to replace it",
            ),
        });
    }
    let engine = resolve_engine(options.engine)?;
    let backend = resolve_backend(engine, options.soffice)?;
    let work = unique_work_dir()?;
    let produced = work.join("export.pdf");
    let converted = convert_with_backend(&backend, input, &produced);
    let result = converted.and_then(|()| {
        if !produced.is_file() {
            return Err(Error::Pdf(format!(
                "{} finished without writing a PDF",
                backend_name(backend.kind)
            )));
        }
        if let Some(parent) = output
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
        {
            fs::create_dir_all(parent).map_err(|source| Error::Output {
                path: parent.to_path_buf(),
                source,
            })?;
        }
        if options.force && output.exists() {
            fs::remove_file(output).map_err(|source| Error::Output {
                path: output.to_path_buf(),
                source,
            })?;
        }
        fs::copy(&produced, output).map_err(|source| Error::Output {
            path: output.to_path_buf(),
            source,
        })?;
        Ok(())
    });
    let _ = fs::remove_dir_all(&work);
    result
}

fn resolve_engine(requested: PdfEngine) -> Result<PdfEngine> {
    if requested != PdfEngine::Auto {
        return Ok(requested);
    }
    let Ok(value) = env::var(ENV_ENGINE) else {
        return Ok(PdfEngine::Auto);
    };
    parse_engine(&value)
}

fn parse_engine(value: &str) -> Result<PdfEngine> {
    match value.trim().to_ascii_lowercase().as_str() {
        "auto" => Ok(PdfEngine::Auto),
        "libreoffice" | "soffice" | "lo" => Ok(PdfEngine::LibreOffice),
        "word" | "office" | "winword" => Ok(PdfEngine::Word),
        "wps" => Ok(PdfEngine::Wps),
        other => Err(Error::Pdf(format!(
            "unknown PDF engine `{other}`; use auto, libreoffice, word, or wps"
        ))),
    }
}

fn resolve_backend(engine: PdfEngine, soffice: Option<&Path>) -> Result<Backend> {
    if let Some(path) = soffice {
        return Ok(Backend {
            kind: BackendKind::LibreOffice,
            path: require_file(path, " `--soffice`")?,
        });
    }
    let word = env_file(ENV_WORD).or_else(discover_word);
    let wps = env_file(ENV_WPS).or_else(discover_wps);
    let libre = resolve_soffice(None).ok();
    let found = |kind: BackendKind| -> Option<Backend> {
        let path = match kind {
            BackendKind::LibreOffice => libre.clone(),
            BackendKind::Word => word.clone(),
            BackendKind::Wps => wps.clone(),
        }?;
        Some(Backend { kind, path })
    };
    let wanted = match engine {
        PdfEngine::Auto => auto_backend_order(),
        PdfEngine::LibreOffice => vec![BackendKind::LibreOffice],
        PdfEngine::Word => vec![BackendKind::Word],
        PdfEngine::Wps => vec![BackendKind::Wps],
    };
    wanted.into_iter().find_map(found).ok_or_else(|| {
        Error::Pdf(format!(
            "no PDF converter found for engine `{}`",
            engine_name(engine)
        ))
    })
}

fn auto_backend_order() -> Vec<BackendKind> {
    if cfg!(windows) {
        vec![
            BackendKind::Word,
            BackendKind::Wps,
            BackendKind::LibreOffice,
        ]
    } else if cfg!(target_os = "macos") {
        vec![
            BackendKind::Word,
            BackendKind::LibreOffice,
            BackendKind::Wps,
        ]
    } else {
        vec![
            BackendKind::LibreOffice,
            BackendKind::Wps,
            BackendKind::Word,
        ]
    }
}

fn convert_with_backend(backend: &Backend, input: &Path, output: &Path) -> Result<()> {
    match backend.kind {
        BackendKind::LibreOffice => convert_with_soffice(&backend.path, input, output),
        BackendKind::Word => convert_with_word(&backend.path, input, output),
        BackendKind::Wps => convert_with_wps(&backend.path, input, output),
    }
}

fn convert_with_soffice(soffice: &Path, input: &Path, output: &Path) -> Result<()> {
    let work = output
        .parent()
        .map_or_else(unique_work_dir, |parent| Ok(parent.to_path_buf()))?;
    let profile = work.join("lo-profile");
    let outdir = work.join("lo-out");
    fs::create_dir_all(&profile).map_err(|source| Error::Output {
        path: profile.clone(),
        source,
    })?;
    fs::create_dir_all(&outdir).map_err(|source| Error::Output {
        path: outdir.clone(),
        source,
    })?;
    run_command(
        Command::new(soffice)
            .arg("--headless")
            .arg("--nologo")
            .arg("--nofirststartwizard")
            .arg("--norestore")
            .arg("--nolockcheck")
            .arg(format!("-env:UserInstallation={}", file_uri(&profile)))
            .arg("--convert-to")
            .arg("pdf:writer_pdf_Export")
            .arg("--outdir")
            .arg(&outdir)
            .arg(input),
        "LibreOffice",
        soffice,
    )?;
    let Some(stem) = input.file_stem().and_then(OsStr::to_str) else {
        return Err(Error::Input(
            "pdf input filename is not valid UTF-8".to_owned(),
        ));
    };
    let produced = outdir.join(format!("{stem}.pdf"));
    if !produced.is_file() {
        return Err(Error::Pdf(format!(
            "LibreOffice finished without writing `{}`",
            produced.display()
        )));
    }
    if produced != output {
        fs::copy(&produced, output).map_err(|source| Error::Output {
            path: output.to_path_buf(),
            source,
        })?;
    }
    Ok(())
}

fn convert_with_word(word: &Path, input: &Path, output: &Path) -> Result<()> {
    if cfg!(windows) {
        return convert_with_com(&["Word.Application"], input, output, "Microsoft Word", word);
    }
    if cfg!(target_os = "macos") {
        return convert_with_osascript(input, output, "Microsoft Word");
    }
    Err(Error::Pdf(format!(
        "Microsoft Word at `{}` cannot export PDF on this OS; use LibreOffice or WPS on Windows",
        word.display()
    )))
}

fn convert_with_wps(wps: &Path, input: &Path, output: &Path) -> Result<()> {
    if cfg!(windows) {
        return convert_with_com(
            &["KWPS.Application", "wps.Application", "Wps.Application"],
            input,
            output,
            "WPS Writer",
            wps,
        );
    }
    Err(Error::Pdf(format!(
        "WPS at `{}` has no supported headless PDF export on this OS; install LibreOffice or use WPS/Word on Windows",
        wps.display()
    )))
}

fn convert_with_com(
    prog_ids: &[&str],
    input: &Path,
    output: &Path,
    label: &str,
    binary: &Path,
) -> Result<()> {
    let input = abs(input)?;
    let output = abs_dest(output)?;
    let ids = prog_ids
        .iter()
        .map(|id| format!("'{id}'"))
        .collect::<Vec<_>>()
        .join(",");
    let script = format!(
        r"
$ErrorActionPreference = 'Stop'
$app = $null
foreach ($id in @({ids})) {{
  try {{ $app = New-Object -ComObject $id; break }} catch {{}}
}}
if ($null -eq $app) {{ throw '{label} COM object is not registered' }}
try {{ $app.Visible = $false }} catch {{}}
try {{ $app.DisplayAlerts = 0 }} catch {{}}
$doc = $null
try {{
  $doc = $app.Documents.Open('{input}', $false, $true)
  try {{
    $doc.ExportAsFixedFormat('{output}', 17)
  }} catch {{
    $doc.SaveAs([ref]'{output}', [ref]17)
  }}
}} finally {{
  if ($null -ne $doc) {{ $doc.Close($false) | Out-Null }}
  $app.Quit() | Out-Null
}}
",
        ids = ids,
        label = label,
        input = ps_single(&input),
        output = ps_single(&output),
    );
    run_command(
        Command::new("powershell.exe")
            .arg("-NoProfile")
            .arg("-NonInteractive")
            .arg("-ExecutionPolicy")
            .arg("Bypass")
            .arg("-Command")
            .arg(script),
        label,
        binary,
    )
}

fn convert_with_osascript(input: &Path, output: &Path, app: &str) -> Result<()> {
    let input = abs(input)?;
    let output = abs_dest(output)?;
    let script = format!(
        r#"tell application "{app}"
  set theDoc to open POSIX file "{input}"
  save as theDoc file name POSIX file "{output}" file format format PDF
  close theDoc saving no
end tell"#,
        app = app,
        input = osa_escape(&input),
        output = osa_escape(&output),
    );
    run_command(
        Command::new("osascript").arg("-e").arg(script),
        app,
        Path::new(app),
    )
}

fn run_command(command: &mut Command, label: &str, binary: &Path) -> Result<()> {
    let mut child = command
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|source| {
            Error::Pdf(format!(
                "cannot start {label} at `{}`: {source}",
                binary.display()
            ))
        })?;
    let status = match wait_with_timeout(&mut child, CONVERT_TIMEOUT) {
        Ok(status) => status,
        Err(error) => {
            let _ = child.kill();
            let _ = child.wait();
            return Err(error);
        }
    };
    if status.success() {
        return Ok(());
    }
    let mut stderr = String::new();
    if let Some(mut pipe) = child.stderr.take() {
        let _ = pipe.read_to_string(&mut stderr);
    }
    let detail = stderr.trim();
    Err(Error::Pdf(format!(
        "{label} at `{}` failed: {}",
        binary.display(),
        if detail.is_empty() {
            "the application exited with an error"
        } else {
            detail
        }
    )))
}

fn wait_with_timeout(
    child: &mut std::process::Child,
    timeout: Duration,
) -> Result<std::process::ExitStatus> {
    let started = Instant::now();
    loop {
        match child.try_wait() {
            Ok(Some(status)) => return Ok(status),
            Ok(None) if started.elapsed() >= timeout => {
                return Err(Error::Pdf(format!(
                    "PDF conversion exceeded {} seconds",
                    timeout.as_secs()
                )));
            }
            Ok(None) => thread::sleep(Duration::from_millis(50)),
            Err(source) => {
                return Err(Error::Pdf(format!(
                    "cannot wait for PDF conversion: {source}"
                )));
            }
        }
    }
}

fn resolve_soffice_from(explicit: Option<&Path>, env_path: Option<&Path>) -> Result<PathBuf> {
    if let Some(path) = explicit {
        return require_file(path, " `--soffice`");
    }
    if let Some(path) = env_path.filter(|path| !path.as_os_str().is_empty()) {
        return require_file(path, " `DOCX_JSX_SOFFICE`");
    }
    for name in [
        "soffice",
        "soffice.exe",
        "soffice.bin",
        "libreoffice",
        "libreoffice.exe",
    ] {
        if let Some(found) = search_path(name) {
            return Ok(found);
        }
    }
    well_known_soffice_paths()
        .into_iter()
        .find(|candidate| candidate.is_file())
        .ok_or_else(|| {
            Error::Pdf(
                "LibreOffice `soffice` was not found on PATH or in common install locations"
                    .to_owned(),
            )
        })
}

fn discover_word() -> Option<PathBuf> {
    if cfg!(not(any(windows, target_os = "macos"))) {
        return None;
    }
    for name in ["WINWORD.EXE", "winword.exe"] {
        if let Some(found) = search_path(name) {
            return Some(found);
        }
    }
    well_known_word_paths()
        .into_iter()
        .find(|candidate| candidate.is_file())
}

fn discover_wps() -> Option<PathBuf> {
    if !cfg!(windows) {
        return None;
    }
    for name in ["wps.exe", "wps"] {
        if let Some(found) = search_path(name) {
            return Some(found);
        }
    }
    well_known_wps_paths()
        .into_iter()
        .find(|candidate| candidate.is_file())
}

fn well_known_soffice_paths() -> Vec<PathBuf> {
    let mut paths = vec![
        PathBuf::from("/usr/bin/soffice"),
        PathBuf::from("/usr/bin/libreoffice"),
        PathBuf::from("/usr/lib/libreoffice/program/soffice"),
        PathBuf::from("/snap/bin/libreoffice"),
        PathBuf::from("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ];
    for root in program_files_roots() {
        paths.push(root.join("LibreOffice/program/soffice.exe"));
    }
    paths
}

fn well_known_word_paths() -> Vec<PathBuf> {
    let mut paths = vec![PathBuf::from(
        "/Applications/Microsoft Word.app/Contents/MacOS/Microsoft Word",
    )];
    for root in program_files_roots() {
        for office in ["Office16", "Office15", "Office14"] {
            paths.push(
                root.join("Microsoft Office")
                    .join(office)
                    .join("WINWORD.EXE"),
            );
            paths.push(
                root.join("Microsoft Office/root")
                    .join(office)
                    .join("WINWORD.EXE"),
            );
        }
    }
    paths
}

fn well_known_wps_paths() -> Vec<PathBuf> {
    let mut paths = vec![
        PathBuf::from("/usr/bin/wps"),
        PathBuf::from("/opt/kingsoft/wps-office/office6/wps"),
        PathBuf::from("/Applications/wpsoffice.app/Contents/MacOS/wpsoffice"),
    ];
    let mut roots = program_files_roots();
    if let Some(local) = env::var_os("LOCALAPPDATA") {
        roots.push(PathBuf::from(local));
    }
    for root in roots {
        paths.extend(wps_candidates_under(&root));
    }
    paths
}

fn wps_candidates_under(root: &Path) -> Vec<PathBuf> {
    let mut paths = Vec::new();
    for base in [
        root.join("Kingsoft/WPS Office"),
        root.join("WPS Office"),
        root.join("Kingsoft/WPS Office/ksolaunch.exe")
            .parent()
            .map_or_else(|| root.join("Kingsoft/WPS Office"), Path::to_path_buf),
    ] {
        let direct = base.join("office6/wps.exe");
        if direct.is_file() {
            paths.push(direct);
        }
        if let Ok(entries) = fs::read_dir(&base) {
            for entry in entries.flatten() {
                let candidate = entry.path().join("office6/wps.exe");
                if candidate.is_file() {
                    paths.push(candidate);
                }
            }
        }
    }
    paths
}

fn program_files_roots() -> Vec<PathBuf> {
    ["PROGRAMFILES", "PROGRAMFILES(X86)", "ProgramW6432"]
        .into_iter()
        .filter_map(|key| env::var_os(key).map(PathBuf::from))
        .collect()
}

fn env_file(key: &str) -> Option<PathBuf> {
    env::var_os(key)
        .map(PathBuf::from)
        .filter(|path| !path.as_os_str().is_empty() && path.is_file())
}

fn require_file(path: &Path, source: &str) -> Result<PathBuf> {
    if path.is_file() {
        return Ok(path.to_path_buf());
    }
    Err(Error::Pdf(format!(
        "PDF converter{source} path `{}` is not a file",
        path.display()
    )))
}

fn search_path(name: &str) -> Option<PathBuf> {
    env::var_os("PATH").and_then(|paths| {
        env::split_paths(&paths).find_map(|dir| {
            let candidate = dir.join(name);
            candidate.is_file().then_some(candidate)
        })
    })
}

fn unique_work_dir() -> Result<PathBuf> {
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_or(0, |duration| duration.as_nanos());
    let dir = env::temp_dir().join(format!("docx-jsx-pdf-{}-{nanos}", std::process::id()));
    fs::create_dir_all(&dir).map_err(|source| Error::Output {
        path: dir.clone(),
        source,
    })?;
    Ok(dir)
}

fn file_uri(path: &Path) -> String {
    let resolved = path.canonicalize().unwrap_or_else(|_| path.to_path_buf());
    let mut path = resolved.to_string_lossy().replace('\\', "/");
    if let Some(stripped) = path.strip_prefix("//?/") {
        path = stripped.to_owned();
    }
    if !path.starts_with('/') {
        path.insert(0, '/');
    }
    format!("file://{path}")
}

fn abs(path: &Path) -> Result<PathBuf> {
    path.canonicalize().map_err(|source| Error::Resource {
        path: path.to_path_buf(),
        source,
    })
}

fn abs_dest(path: &Path) -> Result<PathBuf> {
    if let Some(parent) = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent).map_err(|source| Error::Output {
            path: parent.to_path_buf(),
            source,
        })?;
        let name = path
            .file_name()
            .ok_or_else(|| Error::Input("pdf output filename is not valid UTF-8".to_owned()))?;
        return Ok(parent
            .canonicalize()
            .unwrap_or_else(|_| parent.to_path_buf())
            .join(name));
    }
    env::current_dir()
        .map(|cwd| cwd.join(path))
        .map_err(|source| Error::Output {
            path: path.to_path_buf(),
            source,
        })
}

fn ps_single(path: &Path) -> String {
    path.to_string_lossy().replace('\'', "''")
}

fn osa_escape(path: &Path) -> String {
    path.to_string_lossy()
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
}

fn engine_name(engine: PdfEngine) -> &'static str {
    match engine {
        PdfEngine::Auto => "auto",
        PdfEngine::LibreOffice => "libreoffice",
        PdfEngine::Word => "word",
        PdfEngine::Wps => "wps",
    }
}

fn backend_name(kind: BackendKind) -> &'static str {
    match kind {
        BackendKind::LibreOffice => "LibreOffice",
        BackendKind::Word => "Microsoft Word",
        BackendKind::Wps => "WPS Writer",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::io::Write;
    use tempfile::tempdir;

    #[test]
    fn resolve_should_reject_missing_explicit_soffice() {
        let error = resolve_soffice(Some(Path::new("/no/such/soffice")))
            .expect_err("missing explicit soffice");
        let text = error.to_string();
        assert!(
            text.contains("soffice") && text.contains("/no/such/soffice"),
            "{text}"
        );
    }

    #[test]
    fn resolve_should_prefer_explicit_file() {
        let directory = tempdir().expect("tempdir");
        let soffice = directory.path().join("soffice");
        fs::write(&soffice, b"fake").expect("write fake soffice");
        let resolved = resolve_soffice(Some(&soffice)).expect("explicit file should resolve");
        assert_eq!(resolved, soffice);
    }

    #[test]
    fn resolve_from_should_prefer_explicit_over_env() {
        let directory = tempdir().expect("tempdir");
        let explicit = directory.path().join("explicit");
        let env_path = directory.path().join("from-env");
        fs::write(&explicit, b"ok").expect("write");
        fs::write(&env_path, b"ok").expect("write");
        let resolved =
            resolve_soffice_from(Some(&explicit), Some(&env_path)).expect("explicit should win");
        assert_eq!(resolved, explicit);
    }

    #[test]
    fn parse_engine_should_accept_aliases() {
        assert_eq!(parse_engine("WORD").expect("word"), PdfEngine::Word);
        assert_eq!(parse_engine("wps").expect("wps"), PdfEngine::Wps);
        assert_eq!(parse_engine("soffice").expect("lo"), PdfEngine::LibreOffice);
    }

    #[test]
    fn resolve_backend_should_use_explicit_soffice_even_for_word_engine() {
        let directory = tempdir().expect("tempdir");
        let soffice = directory.path().join("soffice");
        fs::write(&soffice, b"lo").expect("write");
        let backend =
            resolve_backend(PdfEngine::Word, Some(&soffice)).expect("explicit soffice wins");
        assert_eq!(backend.kind, BackendKind::LibreOffice);
        assert_eq!(backend.path, soffice);
    }

    #[test]
    fn resolve_backend_should_error_when_word_missing() {
        let error = resolve_backend(PdfEngine::Word, None).expect_err("no word");
        assert!(
            error.to_string().contains("word") || error.to_string().contains("Word"),
            "{error}"
        );
    }

    #[test]
    fn convert_bytes_should_report_missing_converter_before_writing_output() {
        let directory = tempdir().expect("tempdir");
        let output = directory.path().join("out.pdf");
        let error = convert_docx_bytes_to_pdf(
            b"PK",
            &output,
            PdfOptions {
                engine: PdfEngine::LibreOffice,
                soffice: Some(Path::new("/no/such/soffice")),
                force: true,
            },
        )
        .expect_err("missing soffice");
        assert!(!output.exists(), "must not invent a PDF");
        let diagnostic = error.render_diagnostic();
        assert!(
            diagnostic.contains("DOCX_JSX_SOFFICE") && diagnostic.contains("--engine"),
            "{diagnostic}"
        );
    }

    #[test]
    fn convert_should_reject_non_docx_input() {
        let error = convert_docx_to_pdf(
            Path::new("notes.txt"),
            Path::new("notes.pdf"),
            Some(Path::new("/no/such/soffice")),
            true,
        )
        .expect_err("non-docx must fail");
        assert!(error.to_string().contains(".docx"), "{error}");
    }

    #[test]
    fn convert_should_report_missing_converter_before_writing_output() {
        let directory = tempdir().expect("tempdir");
        let input = directory.path().join("in.docx");
        let output = directory.path().join("out.pdf");
        let mut file = fs::File::create(&input).expect("docx");
        file.write_all(b"PK").expect("bytes");
        drop(file);
        let error = convert_docx_to_pdf(&input, &output, Some(Path::new("/no/such/soffice")), true)
            .expect_err("missing soffice");
        assert!(!output.exists(), "must not invent a PDF");
        let diagnostic = error.render_diagnostic();
        assert!(
            diagnostic.contains("DOCX_JSX_SOFFICE")
                && diagnostic.contains("--engine")
                && diagnostic.contains("WPS"),
            "{diagnostic}"
        );
    }

    #[test]
    fn well_known_paths_should_include_word_and_wps_locations() {
        let word = well_known_word_paths().iter().any(|path| {
            path.ends_with("WINWORD.EXE") || path.to_string_lossy().contains("Microsoft Word")
        });
        let wps = well_known_wps_paths()
            .iter()
            .any(|path| path.ends_with("wps") || path.ends_with("wps.exe"));
        assert!(word && wps);
    }
}
