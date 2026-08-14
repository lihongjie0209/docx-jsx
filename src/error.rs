use std::path::PathBuf;

/// Error type for all compiler phases.
#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("input: {0}")]
    Input(String),
    #[error("transpile: {0}")]
    Transpile(String),
    #[error("module: {0}")]
    Module(String),
    #[error("runtime: {0}")]
    Runtime(String),
    #[error("validation at {path}: {message}")]
    Validation { path: String, message: String },
    #[error("resource {path}: {source}")]
    Resource {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    #[error("compile: {0}")]
    Compile(String),
    #[error("reverse: {0}")]
    Reverse(String),
    #[error("output {path}: {source}")]
    Output {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
}

impl Error {
    /// Formats a diagnostic with a cause explanation and an actionable fix.
    #[must_use]
    pub fn render_diagnostic(&self) -> String {
        format!(
            "error: {self}\nexplanation: {}\nsuggestion: {}\nspec: run `docx-jsx spec` for the complete component contract",
            self.explanation(),
            self.suggestion()
        )
    }

    fn explanation(&self) -> &'static str {
        match self {
            Self::Input(_) => "The command input or one of its arguments is invalid.",
            Self::Transpile(_) => {
                "The JSX/TSX source could not be converted to executable JavaScript."
            }
            Self::Module(_) => "A JavaScript module could not be resolved or evaluated.",
            Self::Runtime(_) => {
                "The JSX module failed while producing the document component tree."
            }
            Self::Validation { .. } => {
                "The produced component tree violates the docx-jsx v1 contract."
            }
            Self::Resource { .. } => "A referenced local file could not be read.",
            Self::Compile(_) => "The validated component tree could not be encoded as DOCX.",
            Self::Reverse(_) => "The DOCX package could not be converted to supported JSX.",
            Self::Output { .. } => "The requested output could not be written safely.",
        }
    }

    fn suggestion(&self) -> &'static str {
        match self {
            Self::Input(message) if message.contains("missing JSX/TSX") => {
                "Pass an entry file, for example: `docx-jsx report.tsx -o report.docx`."
            }
            Self::Input(message) if message.contains("extension") => {
                "Use a .jsx or .tsx entry; for DOCX input use `docx-jsx reverse INPUT.docx`."
            }
            Self::Transpile(_) => {
                "Fix the reported JSX/TypeScript syntax at the indicated source location."
            }
            Self::Module(_) => {
                "Check import paths and use only local ESM modules or the `docx-jsx` runtime import."
            }
            Self::Runtime(_) => {
                "Inspect the reported JavaScript stack and ensure the default export returns Document JSX."
            }
            Self::Validation { message, .. } if message.contains("unknown property") => {
                "Remove or rename the property using the component property list from `docx-jsx spec`."
            }
            Self::Validation { message, .. } if message.contains("cannot contain") => {
                "Move the child into one of the parent component's allowed containers shown by `docx-jsx spec`."
            }
            Self::Validation { message, .. } if message.contains("style inheritance cycle") => {
                "Remove or redirect one `basedOn` reference so the inheritance chain terminates."
            }
            Self::Validation { message, .. } if message.contains("`style` requires a") => {
                "Reference a declared id with the matching style type, or change the style definition's `type`."
            }
            Self::Validation { message, .. }
                if message.contains("style")
                    && (message.contains("same type")
                        || message.contains("must link back")
                        || message.contains("paragraph and character style")
                        || message.contains("`next`")) =>
            {
                "Correct the reported `basedOn`, `next`, or `link` relationship using the style rules from `docx-jsx spec`."
            }
            Self::Validation { message, .. } if message.contains("requires") => {
                "Add the required property or child named by the error, then compile again."
            }
            Self::Validation { message, .. } if message.contains("mutually exclusive") => {
                "Keep only one of the conflicting properties named by the error."
            }
            Self::Validation { .. } => {
                "Correct the value at the reported component path using `docx-jsx spec`."
            }
            Self::Resource { .. } => {
                "Check that the path exists, is readable, and is relative to the entry module when appropriate."
            }
            Self::Compile(_) => {
                "Check referenced images and numeric dimensions; rerun after correcting the reported resource."
            }
            Self::Reverse(_) => {
                "Use a valid .docx package; unsupported external OOXML must be simplified before retrying."
            }
            Self::Output { source, .. } if source.kind() == std::io::ErrorKind::AlreadyExists => {
                "Pass `--force` to replace the existing output, or choose a different `--output` path."
            }
            Self::Output { .. } => {
                "Check the parent directory permissions and available disk space, then retry."
            }
            Self::Input(_) => "Correct the input named by the error and run the command again.",
        }
    }
}

/// Result alias used throughout the crate.
pub type Result<T> = std::result::Result<T, Error>;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn style_cycle_diagnostic_should_explain_how_to_break_inheritance() {
        let diagnostic = Error::Validation {
            path: "Document/styles".to_owned(),
            message: "style inheritance cycle includes `Body`".to_owned(),
        }
        .render_diagnostic();

        assert!(
            diagnostic.contains("Remove or redirect one `basedOn`"),
            "{diagnostic}"
        );
    }

    #[test]
    fn style_reference_diagnostic_should_recommend_matching_style_type() {
        let diagnostic = Error::Validation {
            path: "Document/Section[0]/Paragraph[0]".to_owned(),
            message: "Paragraph `style` requires a paragraph style".to_owned(),
        }
        .render_diagnostic();

        assert!(diagnostic.contains("matching style type"), "{diagnostic}");
    }
}
