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

The conversion/recalculation sidecar is an independent executable and is invoked only for explicitly requested operations. A limited `recalculate` candidate exists; a shared ExcelCalculator interface is a future capability separate from reading, writing, and editing. Future Windows ExcelModifier COM backends (Microsoft Excel and WPS) are optional, capability-gated modification workers; they do not add COM readers or writers and are not Linux dependencies. PDF export and cross-category conversion are not supported.

See AGENTS.md for the full migration, parity, packaging, and target-platform rules.
