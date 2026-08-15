//! Executable JSX/TSX to DOCX compilation.

pub mod compiler;
pub mod error;
pub mod ir;
pub mod pdf;
pub mod reverse;
pub mod runtime;

pub use compiler::compile_document;
pub use error::{Error, Result};
pub use ir::IrEnvelope;
pub use pdf::{
    PdfEngine, PdfOptions, convert_docx_bytes_to_pdf, convert_docx_to_pdf,
    convert_docx_to_pdf_with, resolve_soffice,
};
pub use reverse::{reverse_document, reverse_package};
pub use runtime::evaluate_entry;
