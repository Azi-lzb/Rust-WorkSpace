# pytools-core migration rules

This directory implements the `py_statistics_tools` migration inside the shared workspace. The original project at `C:\Users\AZI\Desktop\VibeCode\4.pytools-main` is read-only.

## Applicable workspace rules

- Follow the parent `AGENTS.md` and workspace Excel contract in `../rules/13_Rust_Excel读写修改与引擎路由规范.md`.
- Use Excel primitives and cross-product reusable capabilities through shared libraries under `crates/`. Generic third-party dependencies must be declared through workspace root Cargo.toml. Do not depend on `audit-core` or duplicate shared Excel engine implementations here.
- Product business logic, reading, and new-workbook output are Rust. Go may be used only under the workspace's conditional Excel modifier policy. Python baselines/helpers under `probe/` or tests are not product dependencies. The only Python runtime exception is the isolated `openpyxl_worker` under `crates/excel-editor`, invoked through a capability-gated ExcelModifier plan for tested XLSX edits; it must not import pytools business code or use system Python/PATH.
- Keep business behavior and configuration semantics in `pytools-core`; keep Excel primitives in shared crates. Use `ExcelEditPlan`/`ExcelModifier` for existing-workbook changes rather than coupling business rules to a concrete engine.
- `src/institution_normalize.rs` has a Rust CLI and `build_institution_normalize_plan`, which emits `ExcelEditPlan` cell updates and has synthetic Python-baseline parity. A separate ignored integration test exercises the CLI's `--apply` path through `ExcelModifier` using a fixture-scoped temporary capability; the synthetic fixture selected the isolated `openpyxl_worker` and passed semantic/file-preservation checks. This is test evidence only: there is no production Passed capability, no authorized real sample, and no Windows-native or glibc 2.28/deb validation. The older `run_normalize_institution` candidate still calls umya directly; this remains a documented gap, not an approved production pattern or default backend. Do not copy that pattern into new business modules.
- Keep fixtures, probe configuration, comparison baselines, and reports under this directory. Do not mix them with Audit evidence.
- Treat `MIGRATION_PLAN.md` as sequencing guidance and `MIGRATION_INVENTORY.md` as the migration status record. A listed candidate or parity on a synthetic fixture does not mean the whole menu feature is migrated or accepted.
- Formula recalculation and same-category conversion are the only permitted LO/COM sidecar use cases, and require explicit invocation. Do not add PDF rendering/export or other Office dependencies.
