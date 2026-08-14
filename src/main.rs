use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use clap::error::ErrorKind;
use clap::{Args, Parser, Subcommand, ValueEnum};
use docx_jsx::{Error, compile_document, evaluate_entry, reverse_document};
use serde_json::Value;

#[derive(Debug, Parser)]
#[command(
    name = "docx-jsx",
    version,
    about = "Compile executable JSX/TSX to DOCX"
)]
struct Cli {
    /// Entry .jsx or .tsx module.
    input: Option<PathBuf>,
    /// Output DOCX path; defaults to INPUT with a .docx extension.
    #[arg(short, long)]
    output: Option<PathBuf>,
    /// JSON file passed to a default-exported root function.
    #[arg(long)]
    data: Option<PathBuf>,
    /// Allow replacing an existing output file.
    #[arg(long)]
    force: bool,
    #[command(subcommand)]
    command: Option<Command>,
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Validate JSX/TSX without producing a DOCX file.
    Validate(ValidateArgs),
    /// Convert a DOCX document to recompilable JSX.
    Reverse(ReverseArgs),
    /// Print the component specification for agents and tooling.
    Spec(SpecArgs),
}

#[derive(Debug, Args)]
struct ValidateArgs {
    /// Input .jsx or .tsx module.
    input: PathBuf,
    /// JSON file passed to a default-exported root function.
    #[arg(long)]
    data: Option<PathBuf>,
}

#[derive(Debug, Args)]
struct SpecArgs {
    /// Output representation.
    #[arg(long, value_enum, default_value_t = SpecFormat::Markdown)]
    format: SpecFormat,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
enum SpecFormat {
    Markdown,
    JsonSchema,
}

enum RunOutput {
    Path(PathBuf),
    Text(String),
}

#[derive(Debug, Args)]
struct ReverseArgs {
    /// Input .docx file.
    input: PathBuf,
    /// Output JSX path; defaults to INPUT with a .jsx extension.
    #[arg(short, long)]
    output: Option<PathBuf>,
    /// Allow replacing an existing output file.
    #[arg(long)]
    force: bool,
}

#[tokio::main(flavor = "current_thread")]
async fn main() -> ExitCode {
    let cli = match Cli::try_parse() {
        Ok(cli) => cli,
        Err(error)
            if matches!(
                error.kind(),
                ErrorKind::DisplayHelp | ErrorKind::DisplayVersion
            ) =>
        {
            let _ = error.print();
            return ExitCode::SUCCESS;
        }
        Err(error) => {
            eprintln!(
                "error: invalid command line\nexplanation: The supplied arguments do not match the CLI contract.\nsuggestion: Correct the option shown below; run `docx-jsx --help` or `docx-jsx spec --help` for accepted arguments.\nspec: run `docx-jsx spec` for the complete component contract\n\n{error}"
            );
            return ExitCode::from(2);
        }
    };
    match run(cli).await {
        Ok(RunOutput::Path(output)) => {
            println!("{}", output.display());
            ExitCode::SUCCESS
        }
        Ok(RunOutput::Text(output)) => {
            print!("{output}");
            ExitCode::SUCCESS
        }
        Err(error) => {
            eprintln!("{}", error.render_diagnostic());
            ExitCode::FAILURE
        }
    }
}

async fn run(cli: Cli) -> docx_jsx::Result<RunOutput> {
    match cli.command {
        Some(Command::Validate(args)) => return run_validate(args).await.map(RunOutput::Text),
        Some(Command::Reverse(args)) => return run_reverse(args).map(RunOutput::Path),
        Some(Command::Spec(args)) => {
            let output = match args.format {
                SpecFormat::Markdown => include_str!("../docs/spec.md"),
                SpecFormat::JsonSchema => include_str!("../spec/ir-v1.schema.json"),
            };
            return Ok(RunOutput::Text(output.to_owned()));
        }
        None => {}
    }
    let input = cli
        .input
        .ok_or_else(|| Error::Input("missing JSX/TSX entry".to_owned()))?;
    validate_input_extension(&input)?;
    let output = cli.output.unwrap_or_else(|| input.with_extension("docx"));
    if output.exists() && !cli.force {
        return Err(Error::Output {
            path: output,
            source: std::io::Error::new(
                std::io::ErrorKind::AlreadyExists,
                "output exists; pass --force to replace it",
            ),
        });
    }
    let data = cli.data.as_deref().map(read_data).transpose()?;
    let ir = evaluate_entry(&input, data.as_ref()).await?;
    let entry_dir = input
        .canonicalize()
        .map_err(|error| Error::Input(format!("cannot open {}: {error}", input.display())))?
        .parent()
        .ok_or_else(|| Error::Input("entry has no parent directory".to_owned()))?
        .to_path_buf();
    let bytes = compile_document(&ir, &entry_dir)?;
    write_atomic(&output, &bytes, cli.force)?;
    Ok(RunOutput::Path(output))
}

async fn run_validate(args: ValidateArgs) -> docx_jsx::Result<String> {
    validate_input_extension(&args.input)?;
    let data = args.data.as_deref().map(read_data).transpose()?;
    evaluate_entry(&args.input, data.as_ref()).await?;
    Ok(format!("valid: {}\n", args.input.display()))
}

fn run_reverse(args: ReverseArgs) -> docx_jsx::Result<PathBuf> {
    if args.input.extension().and_then(|value| value.to_str()) != Some("docx") {
        return Err(Error::Input(
            "reverse input must have a .docx extension".to_owned(),
        ));
    }
    let output = args
        .output
        .unwrap_or_else(|| args.input.with_extension("jsx"));
    if output.exists() && !args.force {
        return Err(output_exists(output));
    }
    let bytes = fs::read(&args.input).map_err(|source| Error::Resource {
        path: args.input,
        source,
    })?;
    let jsx = reverse_document(&bytes)?;
    write_atomic(&output, jsx.as_bytes(), args.force)?;
    Ok(output)
}

fn output_exists(path: PathBuf) -> Error {
    Error::Output {
        path,
        source: std::io::Error::new(
            std::io::ErrorKind::AlreadyExists,
            "output exists; pass --force to replace it",
        ),
    }
}

fn validate_input_extension(path: &Path) -> docx_jsx::Result<()> {
    match path.extension().and_then(|value| value.to_str()) {
        Some("jsx" | "tsx") => Ok(()),
        _ => Err(Error::Input(
            "entry must have a .jsx or .tsx extension".to_owned(),
        )),
    }
}

fn read_data(path: &Path) -> docx_jsx::Result<Value> {
    let source = fs::read_to_string(path).map_err(|source| Error::Resource {
        path: path.to_path_buf(),
        source,
    })?;
    serde_json::from_str(&source)
        .map_err(|error| Error::Input(format!("invalid data JSON {}: {error}", path.display())))
}

fn write_atomic(path: &Path, bytes: &[u8], force: bool) -> docx_jsx::Result<()> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    fs::create_dir_all(parent).map_err(|source| Error::Output {
        path: parent.to_path_buf(),
        source,
    })?;
    let file_name = path
        .file_name()
        .and_then(|value| value.to_str())
        .ok_or_else(|| Error::Input("output filename is not valid UTF-8".to_owned()))?;
    let temporary = parent.join(format!(".{file_name}.{}.tmp", std::process::id()));
    let result = (|| {
        let mut file = OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)
            .map_err(|source| Error::Output {
                path: temporary.clone(),
                source,
            })?;
        file.write_all(bytes).map_err(|source| Error::Output {
            path: temporary.clone(),
            source,
        })?;
        file.sync_all().map_err(|source| Error::Output {
            path: temporary.clone(),
            source,
        })?;
        if force && path.exists() {
            fs::remove_file(path).map_err(|source| Error::Output {
                path: path.to_path_buf(),
                source,
            })?;
        }
        fs::rename(&temporary, path).map_err(|source| Error::Output {
            path: path.to_path_buf(),
            source,
        })
    })();
    if result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    result
}
