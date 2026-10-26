# Open-Data-Audit migration project

This directory implements Open-Data-Audit inside the shared rust-workspace. The original source at C:\Users\AZI\Desktop\VibeCode\3.open-data-audit is read-only.

## Rule applicability

- The parent workspace `AGENTS.md` and `../rules/13_Rust_Excel读写修改与引擎路由规范.md` are binding for all implementation and Excel-engine decisions.
- Files under `rules/` were carried over from the former Python product. Preserve their business semantics, validation rules, user-facing behavior, and OOXML safety details when still applicable; treat Python, openpyxl, COM, UNO, Office, and old shell/path instructions as descriptions of the source implementation unless explicitly restated by the workspace rules.
- Production business logic, reading, and new-workbook output belong in Rust, with Go only under the workspace's conditional modifier policy. Python baselines/test helpers are not product dependencies. Runtime exceptions are limited to the shared isolated `openpyxl_worker` for capability-gated XLSX edits and a future Windows-only pywin32 / `win32com.client` Office COM worker for separately validated edits/calculation; neither may import Audit business logic or use system Python/PATH. COM is not a Linux dependency.
- Current workspace rule map: rules/02 is platform boundary; rules/03 is OOXML safety plus historical implementation notes; rules/04 and rules/06/10 hold business and configuration semantics; rules/07 supplements Rust tests and parity evidence; rules/08 supplies compatible change-recording requirements; rules/11's legacy product component/version scheme does not apply unless separately adopted; rules/13 is an Audit-specific Excel operation map subordinate to the workspace rule 13.

## Scope and workflow

- Keep Audit business behavior in audit-core. Shared spreadsheet primitives belong under crates/. Never depend on pytools-core.
- Reuse existing reader, writer, and editor evidence. Check the engine route needed by the current feature; an unrelated public capability gap must not stop independent Audit work. If a missing capability threatens data correctness, gate that operation, preserve the evidence, and continue other work.
- Source Python, config-real data, and real business workbooks are read-only. Use synthetic or explicitly sanitized fixtures and write probe outputs to temporary directories.
- OpenPyXL is only an optional comparison baseline when the Rust probe cannot inspect a required property.
- Formula recalculation is off by default. The current sidecar has a limited `recalculate` candidate; a future shared calculation interface belongs to an independent ExcelCalculator. Explicit normal/full recalculation may select an implemented, available, validated LibreOffice Calc or Windows Excel/WPS COM backend. Current formula coverage remains capability-specific; the main path preserves formula text and remains usable when an engine is absent.
- Same-category format conversion may use the sidecar when explicitly requested. Initial conversions are .et to .xlsx, .xls to .xlsx for operations such as sheet rename, and .doc to .docx. PDF export and cross-category conversion are out of scope.
- Printing writes page setup properties only; it does not render or export PDF.
- OOXML adapters use namespace-aware parsing, maintain package relationships/content types, and reject unverified parts.
- Do not commit, push, merge, rebase, or publish without an explicit user request.

## Migration phases and minimum gates

1. **Migration:** Audit and PyTools work in their own worktrees and branches in parallel. Complete each confirmed workflow across Rust business logic, the real user entry point, frontend connection where applicable, configuration, progress, error handling, and file operations. Do not leave UI controls as success-shaped placeholders. Prefer source-compatible behavior over unrelated redesign or optimization.
2. **Windows acceptance:** after the agreed workflows have runnable implementations, validate them through the original user operations, real samples, Python/Rust comparisons, integration checks, regression checks, and defect fixes. Implementation or unit-test completion alone is not acceptance.
3. **Linux acceptance:** after critical Windows business acceptance, use the stable Windows baseline for centralized Linux, glibc 2.28, offline deb, and target-system verification. Do not repeat full WSL builds for every feature. Run targeted Linux checks earlier only for a concrete Linux-specific risk; keep all existing delivery constraints.
4. **Delivery:** fix remaining acceptance defects and prepare release materials. Evaluate optional performance or architecture changes after correctness and compatibility are established.

Each feature still requires a relevant compile/build check, scoped unit tests, a smallest useful entry-point smoke check, confirmation that the real entry is connected, input/output protection, and an inventory update that records missing validation. Address known data-loss, overwrite, formula, and workbook-fidelity risks as soon as they are found; do not defer them to final acceptance.

Use these migration states in `MIGRATION_INVENTORY.md`: **未开始**, **迁移中**, **已实现，待 Windows 真实验收**, **Windows 真实验收通过**, **Linux 目标环境验收通过**, and **受阻**. Preserve existing evidence and localized results; a status only advances when its full phase criteria are met. A shared-capability blocker applies to the affected operation, not unrelated features.

## Frontend migration boundary

- The original Python project frontend remains read-only. The workspace-owned frontend copy is under `frontend/web/index.html`, with static assets under `frontend/assets/`; adapt it incrementally when needed to expose migrated Rust workflows, without bulk-regenerating it from Python source.
- `src/web_bridge.rs` injects the `window.bridge.api` transport into the HTTP response at runtime and serves only on loopback. Add APIs there only when the corresponding Rust business function exists; unsupported calls must return an explicit error, not a success-shaped placeholder.
- The current bridge connects the existing S2 Rust service only. Do not present other S1/S3/S4/S5 UI workflows as migrated until their handlers and business behavior pass the project parity gates.
- Python comparison files under `probe/` and `tests/compare/` remain baselines/helpers. The shared openpyxl worker is a narrow ExcelModifier exception; a future COM worker is limited to explicit, capability-gated Windows modification/calculation. Audit does not gain Python business, reader, or writer paths. Do not add an Audit capability record until Audit-specific real-template and target-platform evidence exists.

## Validation

- Keep sanitized/synthetic fixtures and probe configs local to audit-core. Never write probe output over an input fixture.
- Do not claim compatibility for missing fixtures, unavailable APIs, unmeasured performance, or untested target systems.
- A Rust or Go backend becomes a default only after the relevant parity report meets the root AGENTS.md decision threshold.
- Record Audit-only changes under `audit-core/GPT修改记录/` according to rules/08; workspace-wide governance changes use the root `GPT修改记录/`. Document-only changes do not change product versions.
