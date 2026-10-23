# Shared Rust Workspace

This workspace is the Rust-first migration target for Open-Data-Audit and py_statistics_tools. The business implementations are isolated in audit-core and pytools-core. They share only generic Excel primitives under crates/.

## Layout

rust-workspace/
├── Cargo.toml
├── AGENTS.md
├── crates/
│   ├── excel-reader/
│   ├── excel-writer/
│   ├── excel-editor/
│   ├── ooxml-patch/
│   └── probe/
├── audit-core/
├── pytools-core/
└── sidecar/lo-convert/

The reader uses pure Rust calamine with OOXML supplements where required. New workbooks use rust_xlsxwriter. Existing-workbook edits use an ExcelEditPlan/ExcelModifier boundary and route each complete operation batch only to a backend validated for its format, object set, preservation contract, and platform. OOXML Patch, umya-spreadsheet, and Excelize remain conditional candidates; there is no workspace-wide default editor. Go is considered only if real-template comparisons show Rust and OOXML cannot meet a specific preservation requirement.

The sidecar is an independent executable. The main programs invoke it only when users explicitly request formula recalculation or same-category conversion. It is not statically linked or imported by either business crate. PDF export and cross-category conversion are not supported.

See AGENTS.md for the full migration, parity, packaging, and target-platform rules.
