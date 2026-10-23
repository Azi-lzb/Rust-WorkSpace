//! Shared primitives for creating new Excel workbooks.
//!
//! The workspace uses `rust_xlsxwriter` only for new workbooks. Editing an
//! existing workbook belongs in `excel-editor`.

pub use rust_xlsxwriter::{
    Color, ExcelDateTime, Format, FormatAlign, Note, Workbook, Worksheet, XlsxError,
};
