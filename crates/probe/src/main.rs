use anyhow::{bail, Context, Result};
use clap::{Parser, Subcommand};
use serde::Serialize;
use serde_json::Value;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    io::{self, Write},
    path::{Path, PathBuf},
};

#[derive(Debug, Parser)]
#[command(
    name = "probe",
    about = "Rust-first workbook dump and parity comparator"
)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Read a workbook with the shared Rust reader and emit its normalized JSON dump.
    Dump {
        input: PathBuf,
        #[arg(short, long)]
        output: Option<PathBuf>,
    },
    /// Compare a Rust dump with a baseline dump using the same JSON schema.
    Compare {
        rust_dump: PathBuf,
        baseline_dump: PathBuf,
        #[arg(short, long)]
        output: Option<PathBuf>,
    },
}

#[derive(Debug, Serialize)]
struct Metric {
    item: String,
    status: String,
    matched: usize,
    total: usize,
    pass_rate: Option<f64>,
    missing_in_rust: Vec<String>,
    extra_in_rust: Vec<String>,
    changed: Vec<String>,
    note: Option<String>,
}

#[derive(Debug, Serialize)]
struct ComparisonReport {
    schema_version: u32,
    rust_dump: String,
    baseline_dump: String,
    overall_status: String,
    metrics: Vec<Metric>,
}

fn read_json(path: &Path) -> Result<Value> {
    let bytes = fs::read(path).with_context(|| format!("read {}", path.display()))?;
    serde_json::from_slice(&bytes).with_context(|| format!("parse JSON {}", path.display()))
}

fn emit_json(value: &impl Serialize, output: Option<&Path>) -> Result<()> {
    let bytes = serde_json::to_vec_pretty(value)?;
    match output {
        Some(path) => {
            if let Some(parent) = path.parent().filter(|p| !p.as_os_str().is_empty()) {
                fs::create_dir_all(parent)
                    .with_context(|| format!("create output directory {}", parent.display()))?;
            }
            fs::write(path, &bytes).with_context(|| format!("write {}", path.display()))?;
            println!("wrote {}", path.display());
        }
        None => {
            io::stdout().write_all(&bytes)?;
            io::stdout().write_all(b"\n")?;
        }
    }
    Ok(())
}

fn sheets(root: &Value) -> Result<&Vec<Value>> {
    root.get("sheets")
        .and_then(Value::as_array)
        .context("dump must contain a sheets array")
}

fn sheet_names(root: &Value) -> Result<BTreeMap<String, Value>> {
    let mut result = BTreeMap::new();
    for (index, sheet) in sheets(root)?.iter().enumerate() {
        let name = sheet
            .get("name")
            .and_then(Value::as_str)
            .with_context(|| format!("sheets[{index}] must contain a string name"))?;
        result.insert(index.to_string(), Value::String(name.to_owned()));
    }
    Ok(result)
}

fn collection(root: &Value, field: &str) -> Result<BTreeMap<String, Value>> {
    let mut result = BTreeMap::new();
    for (index, sheet) in sheets(root)?.iter().enumerate() {
        let name = sheet
            .get("name")
            .and_then(Value::as_str)
            .with_context(|| format!("sheets[{index}] must contain a string name"))?;

        match field {
            "merges" | "conditional_formatting" | "data_validations" => {
                let ranges = sheet
                    .get(field)
                    .and_then(Value::as_array)
                    .with_context(|| format!("sheet {name} must contain a {field} array"))?;
                for (merge_index, range) in ranges.iter().enumerate() {
                    result.insert(format!("{name}!merge[{merge_index}]"), range.clone());
                }
            }
            _ => {
                let entries = sheet
                    .get(field)
                    .and_then(Value::as_object)
                    .with_context(|| format!("sheet {name} must contain a {field} object"))?;
                for (coordinate, value) in entries {
                    result.insert(format!("{name}!{coordinate}"), value.clone());
                }
            }
        }
    }
    Ok(result)
}

fn has_unsupported_status(root: &Value, field: &str) -> Option<String> {
    let sheet_list = root.get("sheets")?.as_array()?;
    for sheet in sheet_list {
        if let Some(status) = sheet.get(field).and_then(Value::as_str) {
            if status.starts_with("unsupported")
                || status.starts_with("unverified")
                || status.starts_with("read_error")
            {
                return Some(status.to_owned());
            }
        } else {
            return Some(format!("missing_status:{field}"));
        }
    }
    None
}

fn compare_maps(
    item: &str,
    rust: &BTreeMap<String, Value>,
    baseline: &BTreeMap<String, Value>,
    unsupported: Option<String>,
) -> Metric {
    let keys: BTreeSet<_> = rust.keys().chain(baseline.keys()).cloned().collect();
    let mut matched = 0;
    let mut missing = Vec::new();
    let mut extra = Vec::new();
    let mut changed = Vec::new();

    for key in &keys {
        match (rust.get(key), baseline.get(key)) {
            (Some(left), Some(right)) if left == right => matched += 1,
            (None, Some(_)) => missing.push(key.clone()),
            (Some(_), None) => extra.push(key.clone()),
            (Some(_), Some(_)) => changed.push(key.clone()),
            (None, None) => {}
        }
    }

    let total = keys.len();
    let status = if let Some(reason) = unsupported.as_ref() {
        format!("未验证（{reason}）")
    } else if total == 0 {
        "不适用".to_owned()
    } else if !changed.is_empty() || !extra.is_empty() {
        "损坏".to_owned()
    } else if !missing.is_empty() {
        "丢失".to_owned()
    } else {
        "保留".to_owned()
    };

    Metric {
        item: item.to_owned(),
        status,
        matched,
        total,
        pass_rate: (total > 0).then_some(matched as f64 / total as f64),
        missing_in_rust: missing.into_iter().take(100).collect(),
        extra_in_rust: extra.into_iter().take(100).collect(),
        changed: changed.into_iter().take(100).collect(),
        note: unsupported,
    }
}

fn compare(rust_path: &Path, baseline_path: &Path) -> Result<ComparisonReport> {
    let rust = read_json(rust_path)?;
    let baseline = read_json(baseline_path)?;
    let mut metrics = Vec::new();

    metrics.push(compare_maps(
        "sheet_names",
        &sheet_names(&rust)?,
        &sheet_names(&baseline)?,
        None,
    ));

    for (field, item, status_field) in [
        ("values", "values", None),
        ("formulas", "formula_text", Some("formula_status")),
        ("merges", "merged_ranges", Some("merge_status")),
        ("comments", "comments", Some("comments_status")),
        (
            "conditional_formatting",
            "conditional_formatting",
            Some("conditional_formatting_status"),
        ),
        (
            "data_validations",
            "data_validations",
            Some("data_validations_status"),
        ),
    ] {
        let unsupported = status_field.and_then(|name| has_unsupported_status(&rust, name));
        metrics.push(compare_maps(
            item,
            &collection(&rust, field)?,
            &collection(&baseline, field)?,
            unsupported,
        ));
    }

    let overall_status = if metrics.iter().any(|m| m.status.starts_with("未验证")) {
        "未验证"
    } else if metrics.iter().any(|m| m.status == "损坏") {
        "损坏"
    } else if metrics.iter().any(|m| m.status == "丢失") {
        "丢失"
    } else if metrics.iter().any(|m| m.status == "保留") {
        "保留"
    } else {
        "不适用"
    }
    .to_owned();

    Ok(ComparisonReport {
        schema_version: 1,
        rust_dump: rust_path.display().to_string(),
        baseline_dump: baseline_path.display().to_string(),
        overall_status,
        metrics,
    })
}

fn run() -> Result<()> {
    let cli = Cli::parse();
    match cli.command {
        Command::Dump { input, output } => {
            let dump = excel_reader::read_workbook(&input)
                .with_context(|| format!("probe {}", input.display()))?;
            emit_json(&dump, output.as_deref())
        }
        Command::Compare {
            rust_dump,
            baseline_dump,
            output,
        } => {
            let report = compare(&rust_dump, &baseline_dump)?;
            let rejected = report.overall_status != "保留" && report.overall_status != "不适用";
            emit_json(&report, output.as_deref())?;
            if rejected {
                bail!(
                    "parity report status is {}; inspect the report before selecting a default backend",
                    report.overall_status
                );
            }
            Ok(())
        }
    }
}

fn main() {
    if let Err(error) = run() {
        eprintln!("{error:#}");
        std::process::exit(2);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn one_entry(key: &str, value: Value) -> BTreeMap<String, Value> {
        BTreeMap::from([(key.to_owned(), value)])
    }

    #[test]
    fn equal_metadata_is_reported_as_preserved() {
        let values = one_entry("Sheet1!A1", json!({"type":"integer","value":7}));
        let metric = compare_maps("values", &values, &values, None);

        assert_eq!(metric.status, "保留");
        assert_eq!(metric.matched, 1);
        assert_eq!(metric.total, 1);
        assert_eq!(metric.pass_rate, Some(1.0));
    }

    #[test]
    fn missing_metadata_is_reported_as_lost() {
        let rust = BTreeMap::new();
        let baseline = one_entry("Sheet1!A1", json!("expected"));
        let metric = compare_maps("comments", &rust, &baseline, None);

        assert_eq!(metric.status, "丢失");
        assert_eq!(metric.missing_in_rust, vec!["Sheet1!A1"]);
    }

    #[test]
    fn changed_metadata_is_reported_as_damaged() {
        let rust = one_entry("Sheet1!A1", json!("new"));
        let baseline = one_entry("Sheet1!A1", json!("old"));
        let metric = compare_maps("formulas", &rust, &baseline, None);

        assert_eq!(metric.status, "损坏");
        assert_eq!(metric.changed, vec!["Sheet1!A1"]);
    }

    #[test]
    fn empty_metadata_is_not_applicable() {
        let empty = BTreeMap::new();
        let metric = compare_maps("merges", &empty, &empty, None);

        assert_eq!(metric.status, "不适用");
        assert_eq!(metric.pass_rate, None);
    }

    #[test]
    fn unsupported_reader_status_takes_precedence_over_empty_data() {
        let dump = json!({
            "sheets": [{
                "name": "Sheet1",
                "merge_status": "unsupported_api: no merge reader"
            }]
        });
        let status = has_unsupported_status(&dump, "merge_status");
        let empty = BTreeMap::new();
        let metric = compare_maps("merged_ranges", &empty, &empty, status);

        assert!(metric.status.starts_with("未验证"));
    }
}
