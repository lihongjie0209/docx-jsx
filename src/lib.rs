//! Executable JSX/TSX to DOCX compilation.

pub mod compiler;
pub mod error;
pub mod ir;
pub mod reverse;
pub mod runtime;

pub use compiler::compile_document;
pub use error::{Error, Result};
pub use ir::IrEnvelope;
pub use reverse::reverse_document;
pub use runtime::evaluate_entry;
