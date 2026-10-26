# Open-Data-Audit migration project

This directory implements Open-Data-Audit inside the shared rust-workspace. The original source at C:\Users\AZI\Desktop\VibeCode\3.open-data-audit is read-only.

## Rule applicability

- The parent workspace `AGENTS.md` and `../rules/13_Rust_Excel读写修改与引擎路由规范.md` are binding for all implementation and Excel-engine decisions.
- Files under `rules/` were carried over from the former Python product. Preserve their business semantics, validation rules, user-facing behavior, and OOXML safety details when still applicable; treat Python, openpyxl, COM, UNO, Office, and old shell/path instructions as descriptions of the source implementation unless explicitly restated by the workspace rules.
- Production business logic, reading, and new-workbook output belong in Rust, with Go only under the workspace's conditional modifier policy. Python baselines/test helpers are not product dependencies. Runtime exceptions are limited to the shared isolated `openpyxl_worker` for capability-gated XLSX edits and a future Windows-only pywin32 / `win32com.client` Office COM worker for separately validated edits/calculation; neither may import Audit business logic or use system Python/PATH. COM is not a Linux dependency.
- Current workspace rule map: rules/02 is platform boundary; rules/03 is OOXML safety plus historical implementation notes; rules/04 and rules/06/10 hold business and configuration semantics; rules/07 supplements Rust tests and parity evidence; rules/08 supplies compatible change-recording requirements; rules/11's legacy product component/version scheme does not apply unless separately adopted; rules/13 is an Audit-specific Excel operation map subordinate to the workspace rule 13.

## Scope and workflow

- Keep Audit business behavior in audit-core. Shared spreadsheet primitives belong under crates/. Never depend on pytools-core.
- For each business feature, check and record only the reader, writer, and editor capabilities that its actual data flow needs. Do not route an unvalidated Excel capability as production-ready. If a needed shared capability is blocked, retain the minimal reproduction and evidence for main, mark the affected feature blocked, and continue other independent migration work; completion of every workspace probe is not a prerequisite for starting unrelated features.
- Source Python, config-real data, and real business workbooks are read-only. Use synthetic or explicitly sanitized fixtures and write probe outputs to temporary directories.
- OpenPyXL is only an optional comparison baseline when the Rust probe cannot inspect a required property.
- Formula recalculation is off by default. The current sidecar has a limited `recalculate` candidate; a future shared calculation interface belongs to an independent ExcelCalculator. Explicit normal/full recalculation may select an implemented, available, validated LibreOffice Calc or Windows Excel/WPS COM backend. Current formula coverage remains capability-specific; the main path preserves formula text and remains usable when an engine is absent.
- Same-category format conversion may use the sidecar when explicitly requested. Initial conversions are .et to .xlsx, .xls to .xlsx for operations such as sheet rename, and .doc to .docx. PDF export and cross-category conversion are out of scope.
- Printing writes page setup properties only; it does not render or export PDF.
- OOXML adapters use namespace-aware parsing, maintain package relationships/content types, and reject unverified parts.
- Do not commit, push, merge, rebase, or publish without an explicit user request.

## Frontend migration boundary

- The unchanged source frontend is bundled at `frontend/web/index.html`; static assets are under `frontend/assets/`. Do not edit or regenerate the original HTML as part of backend migration.
- `src/web_bridge.rs` injects the `window.bridge.api` transport into the HTTP response at runtime and serves only on loopback. Add APIs there only when the corresponding Rust business function exists; unsupported calls must return an explicit error, not a success-shaped placeholder.
- The current bridge connects the existing S2 Rust service only. Do not present other S1/S3/S4/S5 UI workflows as migrated until their handlers and business behavior pass the project parity gates.
- Python comparison files under `probe/` and `tests/compare/` remain baselines/helpers. The shared openpyxl worker is a narrow ExcelModifier exception; a future COM worker is limited to explicit, capability-gated Windows modification/calculation. Audit does not gain Python business, reader, or writer paths. Do not add an Audit capability record until Audit-specific real-template and target-platform evidence exists.

## Validation

- Keep sanitized/synthetic fixtures and probe configs local to audit-core. Never write probe output over an input fixture.
- Do not claim compatibility for missing fixtures, unavailable APIs, unmeasured performance, or untested target systems.
- A Rust or Go backend becomes a default only after the relevant parity report meets the root AGENTS.md decision threshold.
- Record Audit-only changes under `audit-core/GPT修改记录/` according to rules/08; workspace-wide governance changes use the root `GPT修改记录/`. Document-only changes do not change product versions.
