//! Shared primitives for creating new Excel workbooks.
//!
//! The workspace uses `rust_xlsxwriter` only for new workbooks. Editing an
//! existing workbook belongs in `excel-editor`.

pub use rust_xlsxwriter::{
    Color, ExcelDateTime, Format, FormatAlign, FormatUnderline, Note, Workbook, Worksheet,
    XlsxError,
};

#[cfg(test)]
mod tests {
    use super::{Format, FormatUnderline};

    #[test]
    fn reexported_underline_type_builds_hyperlink_format() {
        let _format = Format::new()
            .set_hyperlink()
            .set_font_size(12)
            .set_underline(FormatUnderline::None);
    }
}
