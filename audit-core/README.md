# Open-Data-Audit core

This package is the Audit project inside the shared Rust workspace. Spreadsheet
primitives come from `../crates/`; project-specific probe operations, config,
rules, fixtures, and acceptance records stay here. No legacy Python business
entry point is included.

The original Open-Data-Audit source snapshot was commit
`cac9e1068367bdc46f27685e043aff0cfcd2294c` from `main`. Public placeholder
configuration is in `config/`; sanitized local templates belong in the ignored
`config-real/` directory.

## Probes

From the workspace root:

```text
cargo run -p audit-core --bin umya-probe -- --input <workbook.xlsx> --output-dir <tmp-dir> --report <report.json>
cargo run -p audit-core --bin rust-fastpath-probe -- read --input <workbook.xlsx> --output <dump.json>
cargo run -p audit-core --bin rust-fastpath-probe -- write-smoke --output <new-workbook.xlsx>
```

The Python scripts under `tests/compare/` use OpenPyXL only as a comparison
baseline. They are not runtime dependencies. `tests/TEST_GROUPS.md` is retained
as a clearly marked legacy Python regression map; its old `core/tests` commands
do not apply to this Rust workspace. Markdown findings are in `GPT修改记录/`;
per-run JSON output goes to `probe/reports/`.

## Rust S2 CLI candidate

The file-based S2 audit flow can be invoked without Python:

```text
open-data-audit s2 --current-dir <current-period-dir> --previous-dir <previous-period-dir> --config <s2-config.xlsx> --output-dir <results-dir> [--external <external-workbook.xlsx>]
```

Progress events are emitted as JSON lines on stderr; successful runs print a
JSON summary and result-workbook path on stdout. The CLI and its report output
match the read-only Python service on three synthetic scenarios. Real business
templates, Windows-native execution, target Linux/deb, and production routing
remain unverified; see `MIGRATION_INVENTORY.md` and
`probe/reports/s2_service_parity.json`.

The current reports cover a synthetic fixture only. Real sanitized workbook and
Kylin/UOS installation acceptance are still required before business migration
or deb packaging.

## Rust S1 template recommendation slice

The `open-data-audit template-suggest` command reads `.xlsx` worksheet names
through the shared namespace-aware OOXML metadata adapter and applies the
original S1 worksheet-set recommendation rules without writing to the template
or input directories:

```text
open-data-audit template-suggest --template-dir <templates-dir> --input-dir <source-dir> [--recursive]
```

This is the template-discovery slice only. It does not validate S1 rules,
recalculate formulas, write audit results, or update history. Four synthetic
cases match the Python baseline; a read-only smoke comparison also matched the
source `config-real` templates and the 2026-07-31 S1 source tree. One WSL2
measurement on that directory tree completed in a 1.37-second median versus
1.76 seconds for the Python baseline; this is a single-sample observation, not
a general performance claim. See
`probe/reports/template_recommendation_parity.json` and
`probe/reports/template_recommendation_real_smoke.json`.

## Rust S1 conditional-format issue extraction slice

`open-data-audit conditional-audit` evaluates supported static conditional
format rules and converts triggers into Audit `Issue` records in Rust. Supply
the S1 metadata already resolved by a caller; optional scan and structure
areas use `Sheet!A1:B20` notation:

```text
open-data-audit conditional-audit --workbook <source.xlsx> --period <period> --batch-id <id> --audit-time <timestamp> [--source-file <original-path>] [--org-code <code>] [--org-name <name>] [--scan-area Sheet1!A1:Z200] [--structure-area Sheet1!A1:Z1] [--developer-view]
```

The single-workbook command prints issue records as JSON on stdout. For a
selected multi-workbook S1 batch, pass a manifest with per-file identity:

```text
open-data-audit conditional-audit-batch --manifest <batch.json> --period <period> --batch-id <id> --audit-time <timestamp> [--scan-area Sheet1!A1:Z200] [--structure-area Sheet1!A1:Z1] [--developer-view]
```

The manifest has a `files` array; each entry supplies `workbook` and may supply
`source_file`, `org_code`, and `org_name`. Relative paths resolve from the
manifest directory. Workbooks are evaluated in parallel by file; returned
summaries and issues retain manifest order and per-file metadata. Any file
failure fails the whole batch without returning a partial issue summary. Both
commands print actual engine and stage timings in ordinary logs;
`--developer-view` also emits per-file JSON progress events on stderr. Formula
cells use stored caches; no recalculation or workbook write occurs.
`--structure-area` lets the caller resolve `check_field`. This is only
the conditional-format issue-generation stage: S1 config/template selection,
ordinary business rules, output workbook writeback, and history are not yet
connected. The synthetic evaluator, issue mapping, and CLI tests pass; one
read-only real-workbook comparison matched all 348 Python issues field by field
(see `probe/reports/s1_conditional_audit_real.json`). Both runs used the same
temporary ASCII basename to work around WSL Unicode argument handling, so
Issue identity with the original workbook filename remains unverified. Full
S1-F01 remains unverified.

## Rust web UI bridge (incremental)

The original Audit frontend is copied without edits to `frontend/web/index.html`;
the original icon files are copied to `frontend/assets/`. The Rust server injects
the `window.bridge.api` transport into the HTTP response and binds only to
`127.0.0.1`:

```text
cargo run -p audit-core --bin open-data-audit -- serve --port 8750 --config <s2-config.xlsx>
```

The server defaults to `OPEN_DATA_AUDIT_S2_CONFIG`, then searches the workspace
for a local real config and the public placeholder. Pass `--config` when running
outside the workspace or when the intended S2 config is elsewhere.

Currently connected APIs are `get_state`, `browse_dir`,
`check_period_audit_config`, `choose_period_compare`, `clear_path`,
`toggle_period_output_pin`, `start_period_compare_v2`, and a limited `update`
subset for appearance and interaction settings. The page opens on S2. Period
directory selection now refreshes the workbook/sheet pair preview using the
Rust Calamine sheet-name reader and Audit filename identity rules; the original
HTML remains unchanged. Other frontend APIs return HTTP 501 with the
unimplemented API name. Cancellation, settings persistence, S1/S3/S4/S5
workflows, and Linux/deb packaging are not implemented or accepted yet. The
local HTTP bridge is an incremental UI entry point, not evidence that all Audit
features have migrated.
