//! Pure discovery helpers migrated from `base_audit.discovery`.
//!
//! The module includes filename/text period-token parsing and source workbook
//! discovery with period selection, migrated from `base_audit.discovery`.

use anyhow::{Context, Result};
use serde::Serialize;
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
};

const SOURCE_SUFFIX: &str = "xlsx";
const SKIP_DIRECTORY_MARKER: &str = "_skip";
const GENERATED_OUTPUT_FOLDER: &str = "执行结果";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScanDepth {
    /// Scan only direct entries under the input directory.
    RootOnly,
    /// Recurse at most this many directory levels below the input directory.
    MaxDirectoryDepth(usize),
    /// Recurse through all descendant directories.
    Full,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct PeriodDetection {
    pub period: String,
    pub source: String,
    pub conflict: bool,
    pub details: String,
}

/// Infer the institution identity from an input workbook's filename using the
/// source Audit service convention: the first two non-empty underscore parts,
/// or the stem for both fields when fewer than two parts are present.
pub fn organisation_from_name(path: &Path) -> (String, String) {
    let stem = path
        .file_stem()
        .map(|value| value.to_string_lossy().into_owned())
        .unwrap_or_default();
    let parts = stem
        .split('_')
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>();
    if parts.len() >= 2 {
        (parts[0].to_owned(), parts[1].to_owned())
    } else {
        (stem.clone(), stem)
    }
}

/// Enumerate source xlsx paths with the same filters and recursion depth as the
/// original `base_audit.discovery.source_workbooks` helper.
pub fn source_workbooks(input_dir: &Path, depth: ScanDepth) -> Result<Vec<PathBuf>> {
    if !input_dir.is_dir() {
        return Ok(Vec::new());
    }

    let mut paths = Vec::new();
    collect_paths(input_dir, 0, depth, &mut paths)?;
    paths.retain(|path| {
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case(SOURCE_SUFFIX))
            && !path
                .file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("~$"))
            && !path
                .file_stem()
                .and_then(|name| name.to_str())
                .is_some_and(|stem| stem.contains("_审核版"))
            && !relative_parent_parts(input_dir, path)
                .iter()
                .any(|part| is_skipped_directory(part))
    });
    paths.sort();
    Ok(paths)
}

/// Detect one period from source workbook names, then fall back through the
/// input directory and its ancestors. Multiple periods are reported as a
/// conflict instead of selecting one arbitrarily.
pub fn detect_period(input_dir: &Path, depth: ScanDepth) -> Result<PeriodDetection> {
    let files = source_workbooks(input_dir, depth)?;
    Ok(detect_period_from_workbooks(input_dir, &files))
}

/// Detect a period using only the already-selected source workbooks, then fall
/// back through the input directory and its ancestors.
pub fn detect_period_from_workbooks(input_dir: &Path, files: &[PathBuf]) -> PeriodDetection {
    let filename_periods = files
        .iter()
        .filter_map(|path| path.file_stem().and_then(|stem| stem.to_str()))
        .flat_map(periods_in_text)
        .collect::<BTreeSet<_>>();

    if filename_periods.len() == 1 {
        let period = filename_periods.iter().next().unwrap().clone();
        return PeriodDetection {
            period: period.clone(),
            source: "源文件名".to_owned(),
            conflict: false,
            details: format!("从源文件名识别为 {period}"),
        };
    }
    if filename_periods.len() > 1 {
        return period_conflict("源文件名", &filename_periods);
    }

    for folder in input_dir.ancestors() {
        let folder_name = folder
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or_default();
        let periods = periods_in_text(folder_name);
        if periods.len() == 1 {
            let period = periods.iter().next().unwrap().clone();
            return PeriodDetection {
                period: period.clone(),
                source: "文件夹名称".to_owned(),
                conflict: false,
                details: format!("从文件夹名称识别为 {period}"),
            };
        }
        if periods.len() > 1 {
            return period_conflict("文件夹名称", &periods);
        }
    }

    PeriodDetection {
        period: String::new(),
        source: String::new(),
        conflict: false,
        details: "未识别到数据期，请手动填写".to_owned(),
    }
}

fn collect_paths(
    directory: &Path,
    child_depth: usize,
    max_depth: ScanDepth,
    paths: &mut Vec<PathBuf>,
) -> Result<()> {
    let entries = fs::read_dir(directory)
        .with_context(|| format!("read input directory {}", directory.display()))?
        .collect::<std::io::Result<Vec<_>>>()?;
    for entry in entries {
        let path = entry.path();
        let file_type = entry.file_type()?;
        paths.push(path.clone());
        let should_recurse = match max_depth {
            ScanDepth::RootOnly => false,
            ScanDepth::MaxDirectoryDepth(max) => child_depth < max,
            ScanDepth::Full => true,
        };
        if file_type.is_dir() && should_recurse {
            collect_paths(&path, child_depth + 1, max_depth, paths)?;
        }
    }
    Ok(())
}

fn relative_parent_parts(input_dir: &Path, path: &Path) -> Vec<String> {
    path.strip_prefix(input_dir)
        .ok()
        .and_then(Path::parent)
        .into_iter()
        .flat_map(Path::components)
        .map(|component| component.as_os_str().to_string_lossy().into_owned())
        .collect()
}

fn is_skipped_directory(name: &str) -> bool {
    let folded = name.to_lowercase();
    folded == GENERATED_OUTPUT_FOLDER || folded.contains(SKIP_DIRECTORY_MARKER)
}

fn period_conflict(source: &str, periods: &BTreeSet<String>) -> PeriodDetection {
    let values = periods.iter().cloned().collect::<Vec<_>>().join("、");
    let details = if source == "源文件名" {
        format!("源文件中识别出多个数据期：{values}")
    } else {
        format!("文件夹名称中识别出多个数据期：{values}")
    };
    PeriodDetection {
        period: String::new(),
        source: source.to_owned(),
        conflict: true,
        details,
    }
}

/// Find `YYYY-MM` periods in text using the source project's accepted formats.
///
/// Recognized forms include separated periods (`2026-6`, `2026年06月`,
/// `2026/06/30`) and compact periods (`202606`, `20260630`). Like the source
/// regular expressions, the parser validates the month and the optional day
/// token's shape, but does not validate calendar dates (for example, February
/// 30 is accepted). The returned set is sorted and deduplicated.
pub fn periods_in_text(value: &str) -> BTreeSet<String> {
    let chars: Vec<char> = value.chars().collect();
    let mut periods = BTreeSet::new();

    for start in 0..chars.len() {
        if start > 0 && decimal_digit(chars[start - 1]).is_some() {
            continue;
        }

        if let Some((year, month_end, month)) = separated_prefix(&chars, start) {
            // The source allows the optional `月` or date suffix to be omitted.
            // Because those suffixes begin with non-digits, the regex's final
            // negative digit boundary is equivalent to checking this position.
            if chars
                .get(month_end)
                .is_none_or(|character| decimal_digit(*character).is_none())
            {
                periods.insert(format!("{year}-{month:02}"));
            }
        }

        if let Some((year, month_end, month)) = compact_prefix(&chars, start) {
            let end = match chars.get(month_end).copied() {
                Some(next) if decimal_digit(next).is_some() => {
                    let Some(day_end) = compact_day_end(&chars, month_end) else {
                        continue;
                    };
                    day_end
                }
                _ => month_end,
            };

            if chars
                .get(end)
                .is_none_or(|character| decimal_digit(*character).is_none())
            {
                periods.insert(format!("{year}-{month:02}"));
            }
        }
    }

    periods
}

fn separated_prefix(chars: &[char], start: usize) -> Option<(String, usize, u32)> {
    let (year, after_year) = year_at(chars, start)?;
    if !matches!(chars.get(after_year), Some('-' | '.' | '_' | '年' | '/')) {
        return None;
    }

    let month_start = after_year + 1;
    let (month, month_end) = separated_month(chars, month_start)?;
    Some((year, month_end, month))
}

fn compact_prefix(chars: &[char], start: usize) -> Option<(String, usize, u32)> {
    let (year, after_year) = year_at(chars, start)?;
    let month_first = ascii_digit(*chars.get(after_year)?)?;
    let month_second = ascii_digit(*chars.get(after_year + 1)?)?;
    let month = two_digit_month(month_first, month_second)?;
    Some((year, after_year + 2, month))
}

fn year_at(chars: &[char], start: usize) -> Option<(String, usize)> {
    if chars.get(start) != Some(&'2') || chars.get(start + 1) != Some(&'0') {
        return None;
    }
    decimal_digit(*chars.get(start + 2)?)?;
    decimal_digit(*chars.get(start + 3)?)?;
    let year: String = chars[start..start + 4].iter().collect();
    Some((year, start + 4))
}

fn separated_month(chars: &[char], start: usize) -> Option<(u32, usize)> {
    let first = ascii_digit(*chars.get(start)?)?;
    if first == 0 {
        let second = ascii_digit(*chars.get(start + 1)?)?;
        return (1..=9).contains(&second).then_some((second, start + 2));
    }
    if first == 1 {
        if let Some(second @ 0..=2) = chars.get(start + 1).and_then(|c| ascii_digit(*c)) {
            return Some((10 + second, start + 2));
        }
    }
    (1..=9).contains(&first).then_some((first, start + 1))
}

fn two_digit_month(first: u32, second: u32) -> Option<u32> {
    match (first, second) {
        (0, 1..=9) => Some(second),
        (1, 0..=2) => Some(10 + second),
        _ => None,
    }
}

fn compact_day_end(chars: &[char], start: usize) -> Option<usize> {
    let first = ascii_digit(*chars.get(start)?)?;
    let second_character = *chars.get(start + 1)?;
    let valid = match first {
        0 => ascii_digit(second_character).is_some_and(|digit| (1..=9).contains(&digit)),
        1 | 2 => decimal_digit(second_character).is_some(),
        3 => ascii_digit(second_character).is_some_and(|digit| (0..=1).contains(&digit)),
        _ => false,
    };
    valid.then_some(start + 2)
}

fn ascii_digit(character: char) -> Option<u32> {
    character.to_digit(10).filter(|_| character.is_ascii())
}

// Python's Unicode `\d` matches General_Category=Nd. The starts below are
// the Nd blocks in Unicode 15.0; this keeps boundary checks and decimal-value
// extraction aligned with the Python regex for the relevant Unicode repertoire.
const DECIMAL_ZEROES_UNICODE_15: &[u32] = &[
    0x0030, 0x0660, 0x06F0, 0x07C0, 0x0966, 0x09E6, 0x0A66, 0x0AE6, 0x0B66, 0x0BE6, 0x0C66, 0x0CE6,
    0x0D66, 0x0DE6, 0x0E50, 0x0ED0, 0x0F20, 0x1040, 0x1090, 0x17E0, 0x1810, 0x1946, 0x19D0, 0x1A80,
    0x1A90, 0x1B50, 0x1BB0, 0x1C40, 0x1C50, 0xA620, 0xA8D0, 0xA900, 0xA9D0, 0xA9F0, 0xAA50, 0xABF0,
    0xFF10, 0x104A0, 0x10D30, 0x11066, 0x110F0, 0x11136, 0x111D0, 0x112F0, 0x11450, 0x114D0,
    0x11650, 0x116C0, 0x11730, 0x118E0, 0x11950, 0x11C50, 0x11D50, 0x11DA0, 0x11F50, 0x16A60,
    0x16AC0, 0x16B50, 0x1D7CE, 0x1D7D8, 0x1D7E2, 0x1D7EC, 0x1D7F6, 0x1E140, 0x1E2F0, 0x1E4F0,
    0x1E950, 0x1FBF0,
];

pub(crate) fn decimal_digit(character: char) -> Option<u32> {
    let codepoint = character as u32;
    DECIMAL_ZEROES_UNICODE_15
        .iter()
        .find_map(|zero| (codepoint >= *zero && codepoint < *zero + 10).then(|| codepoint - *zero))
}

#[cfg(test)]
mod tests {
    use super::{
        detect_period, organisation_from_name, periods_in_text, source_workbooks, ScanDepth,
    };
    use serde::Deserialize;
    use serde_json::{json, Value};
    use std::{
        fs,
        path::{Path, PathBuf},
        time::{SystemTime, UNIX_EPOCH},
    };

    #[derive(Deserialize)]
    struct PythonBaseline {
        cases: Vec<PythonCase>,
    }

    #[derive(Deserialize)]
    struct PythonCase {
        input: String,
        periods: Vec<String>,
    }

    #[test]
    fn organisation_from_name_matches_source_service_split_rule() {
        assert_eq!(
            organisation_from_name(Path::new("9144_甲银行_20260731.xlsx")),
            ("9144".to_owned(), "甲银行".to_owned())
        );
        assert_eq!(
            organisation_from_name(Path::new("__机构甲___20260731.xlsx")),
            ("机构甲".to_owned(), "20260731".to_owned())
        );
        assert_eq!(
            organisation_from_name(Path::new("未分隔文件.xlsx")),
            ("未分隔文件".to_owned(), "未分隔文件".to_owned())
        );
    }

    #[test]
    fn periods_in_text_matches_python_baseline_vectors() {
        let baseline: PythonBaseline = serde_json::from_str(include_str!(
            "../probe/reports/period_parser_python_baseline.json"
        ))
        .expect("valid Python baseline report");

        assert_eq!(baseline.cases.len(), 49, "baseline vector count changed");
        for (index, case) in baseline.cases.iter().enumerate() {
            let actual: Vec<_> = periods_in_text(&case.input).into_iter().collect();
            assert_eq!(
                actual,
                case.periods,
                "Python parity case {}: {:?}",
                index + 1,
                case.input
            );
        }
    }

    #[test]
    fn source_workbooks_and_period_detection_match_python_baseline() {
        let cases: Value = serde_json::from_str(include_str!(
            "../tests/fixtures/period_detection_cases.json"
        ))
        .expect("valid period detection cases");
        let baseline: Value = serde_json::from_str(include_str!(
            "../probe/reports/period_detection_python_baseline.json"
        ))
        .expect("valid Python period detection baseline");
        let cases = cases["cases"].as_array().unwrap();
        let expected = baseline["cases"].as_array().unwrap();
        assert_eq!(cases.len(), expected.len());

        for (case, expected) in cases.iter().zip(expected) {
            let id = case["id"].as_str().unwrap();
            assert_eq!(id, expected["id"].as_str().unwrap());
            let root = materialize_case(case);
            let depth = parse_depth(&case["recursive"]);

            let actual_files = source_workbooks(&root, depth)
                .unwrap()
                .iter()
                .map(|path| relative_path(&root, path))
                .collect::<Vec<_>>();
            let actual_detection = detect_period(&root, depth).unwrap();
            let actual = json!({
                "source_workbooks": actual_files,
                "result": actual_detection,
            });
            assert_eq!(
                actual["source_workbooks"], expected["source_workbooks"],
                "file list case {id}"
            );
            assert_eq!(actual["result"], expected["result"], "period case {id}");

            fs::remove_dir_all(root.parent().unwrap()).unwrap();
        }
    }

    fn parse_depth(value: &Value) -> ScanDepth {
        match value {
            Value::Bool(false) => ScanDepth::RootOnly,
            Value::Bool(true) => ScanDepth::Full,
            Value::Number(number) => match number.as_i64().unwrap() {
                depth if depth < 0 => ScanDepth::Full,
                depth => ScanDepth::MaxDirectoryDepth(depth as usize),
            },
            other => panic!("unsupported recursive depth: {other}"),
        }
    }

    fn materialize_case(case: &Value) -> PathBuf {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let base = std::env::temp_dir().join(format!(
            "audit-period-detection-{}-{unique}",
            std::process::id()
        ));
        let mut parent = base.clone();
        for name in case["parents"].as_array().unwrap() {
            parent.push(name.as_str().unwrap());
            fs::create_dir_all(&parent).unwrap();
        }
        let root = parent.join(case["root_name"].as_str().unwrap());
        fs::create_dir_all(&root).unwrap();

        for directory in case["directories"].as_array().unwrap() {
            fs::create_dir_all(root.join(directory.as_str().unwrap())).unwrap();
        }
        for file in case["files"].as_array().unwrap() {
            let path = root.join(file.as_str().unwrap());
            fs::create_dir_all(path.parent().unwrap()).unwrap();
            fs::write(path, b"").unwrap();
        }
        root
    }

    fn relative_path(root: &Path, path: &Path) -> String {
        path.strip_prefix(root)
            .unwrap()
            .to_string_lossy()
            .replace('\\', "/")
    }
}
