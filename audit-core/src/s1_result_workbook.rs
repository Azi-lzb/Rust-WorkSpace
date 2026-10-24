//! S1-F01 result-workbook output migrated from the native Python handler.

use crate::models::Issue;
use anyhow::{ensure, Context, Result};
use excel_reader::read_workbook;
use excel_writer::{Color, Format, FormatUnderline, Workbook};
use ooxml_patch::{patch_external_hyperlinks, ExternalHyperlinkPatch};
use serde::Serialize;
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    time::{Instant, SystemTime, UNIX_EPOCH},
};

pub const S1_RESULT_SHEET: &str = "本期审核结果";
pub const S1_RESULT_HEADERS: [&str; 13] = [
    "工作簿名",
    "工作表名",
    "定位单元格",
    "错误类型",
    "校验类型",
    "校验指标",
    "描述",
    "当前值",
    "对比值",
    "差值",
    "规则编号",
    "历史校验说明",
    "审核意见",
];
const COLUMN_WIDTHS: [u32; 13] = [38, 24, 14, 16, 20, 28, 42, 18, 18, 18, 38, 34, 24];
const ENGINE: &str = "rust_xlsxwriter 0.101.0";

#[derive(Clone, Debug, Eq, PartialEq)]
struct ResultLink {
    target_uri: String,
    location: String,
    display: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S1ResultWorkbookSummary {
    pub path: PathBuf,
    pub sheet_name: String,
    pub issue_count: usize,
    pub hyperlink_count: usize,
    pub engine: &'static str,
    pub elapsed_ms: u128,
}

/// Create the S1-F01 current-result workbook using the Python output contract.
///
/// This creates a new XLSX. It does not edit source workbooks or recalculate
/// formulas. Existing output paths are rejected to avoid overwriting user data.
pub fn write_s1_result_workbook(
    output_path: impl AsRef<Path>,
    sheet_name: Option<&str>,
    issues: &[Issue],
) -> Result<S1ResultWorkbookSummary> {
    let started = Instant::now();
    let output_path = output_path.as_ref();
    ensure!(
        output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S1 结果工作簿输出路径必须使用 .xlsx"
    );
    ensure!(
        !output_path.exists(),
        "拒绝覆盖已有 S1 结果工作簿：{}",
        output_path.display()
    );
    ensure!(
        issues.len() <= 1_048_575,
        "S1 结果工作簿超过 Excel 的最大数据行数"
    );
    let parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    fs::create_dir_all(parent)
        .with_context(|| format!("创建 S1 结果工作簿目录 {}", parent.display()))?;

    let sheet_name = safe_sheet_name(sheet_name.unwrap_or(S1_RESULT_SHEET));
    let rows = issues.iter().map(issue_row).collect::<Vec<_>>();
    let links = issues
        .iter()
        .map(|issue| result_link(issue))
        .collect::<Result<Vec<_>>>()?;
    let hyperlink_count = links.iter().filter(|link| link.is_some()).count();
    let writer_temp_path = temporary_path(output_path)?;

    let mut patched_path = None;
    let write_result = (|| -> Result<()> {
        {
            let mut workbook = Workbook::new();
            let worksheet = workbook.add_worksheet();
            worksheet
                .set_name(&sheet_name)
                .context("设置 S1 结果工作表名称")?;

            let header_format = Format::new()
                .set_bold()
                .set_background_color(Color::RGB(0xD9_EA_D3));
            let text_format = Format::new().set_num_format("@");
            let hyperlink_format = Format::new()
                .set_hyperlink()
                .set_font_size(12)
                .set_underline(FormatUnderline::None);
            for (column, header) in S1_RESULT_HEADERS.iter().enumerate() {
                worksheet
                    .write_string_with_format(0, column as u16, *header, &header_format)
                    .with_context(|| format!("写入 S1 结果表头第 {} 列", column + 1))?;
            }

            for (row_index, row) in rows.iter().enumerate() {
                let row_number = row_index as u32 + 1;
                for (column, value) in row.iter().enumerate() {
                    if let Some(link) = links[row_index].as_ref().filter(|_| column == 2) {
                        worksheet
                            .write_url_with_options(
                                row_number,
                                column as u16,
                                link.target_uri.as_str(),
                                link.display.as_str(),
                                "",
                                Some(&hyperlink_format),
                            )
                            .with_context(|| {
                                format!("写入 S1 审核副本超链接第 {} 行", row_index + 1)
                            })?;
                    } else {
                        worksheet
                            .write_string_with_format(
                                row_number,
                                column as u16,
                                value,
                                &text_format,
                            )
                            .with_context(|| {
                                format!("写入 S1 结果第 {} 行第 {} 列", row_index + 1, column + 1)
                            })?;
                    }
                }
            }

            worksheet
                .set_freeze_panes(1, 0)
                .context("冻结 S1 结果表头")?;
            worksheet
                .autofilter(0, 0, rows.len() as u32, S1_RESULT_HEADERS.len() as u16 - 1)
                .context("设置 S1 结果表筛选")?;
            for (column, width) in COLUMN_WIDTHS.iter().enumerate() {
                worksheet
                    .set_column_width_pixels(column as u16, width * 7)
                    .with_context(|| format!("设置 S1 结果表第 {} 列宽度", column + 1))?;
            }
            workbook
                .save(&writer_temp_path)
                .context("保存 S1 结果工作簿临时文件")?;
        }

        validate_result_workbook(&writer_temp_path, &sheet_name, &rows)?;
        let candidate_path = if hyperlink_count == 0 {
            writer_temp_path.as_path()
        } else {
            let output = temporary_path(output_path)?;
            patched_path = Some(output.clone());
            let mut patches = Vec::with_capacity(hyperlink_count);
            for (row_index, link) in links.iter().enumerate() {
                let Some(link) = link else {
                    continue;
                };
                let mut patch = ExternalHyperlinkPatch::new(
                    &sheet_name,
                    cell_address(row_index + 1, 2),
                    &link.target_uri,
                )?;
                patch.location = Some(link.location.clone());
                patch.display = Some(link.display.clone());
                patches.push(patch);
            }
            let reports = patch_external_hyperlinks(&writer_temp_path, &output, &patches)
                .context("修复 S1 结果工作簿审核副本链接")?;
            ensure!(
                reports.len() == hyperlink_count,
                "S1 结果工作簿超链接补丁数量不匹配：预期 {}，实际 {}",
                hyperlink_count,
                reports.len()
            );
            validate_result_workbook(&output, &sheet_name, &rows)?;
            patched_path.as_deref().context("S1 超链接候选路径丢失")?
        };
        fs::rename(candidate_path, output_path).with_context(|| {
            format!(
                "提交 S1 结果工作簿 {} → {}",
                candidate_path.display(),
                output_path.display()
            )
        })?;
        Ok(())
    })();
    if write_result.is_err() {
        let _ = fs::remove_file(&writer_temp_path);
        if let Some(path) = patched_path.as_deref() {
            let _ = fs::remove_file(path);
        }
    } else if patched_path.is_some() {
        let _ = fs::remove_file(&writer_temp_path);
    }
    write_result?;

    let path = fs::canonicalize(output_path)
        .with_context(|| format!("解析 S1 结果工作簿路径 {}", output_path.display()))?;
    Ok(S1ResultWorkbookSummary {
        path,
        sheet_name,
        issue_count: issues.len(),
        hyperlink_count,
        engine: ENGINE,
        elapsed_ms: started.elapsed().as_millis(),
    })
}

fn issue_row(issue: &Issue) -> Vec<String> {
    let source_name = Path::new(&issue.source_file)
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_default();
    vec![
        source_name,
        issue.sheet_name.clone(),
        issue.target_cell.clone(),
        issue.severity.clone(),
        issue.check_type.clone(),
        issue.check_field.clone(),
        if issue.detail.is_empty() {
            issue.message.clone()
        } else {
            issue.detail.clone()
        },
        python_plain_value(&issue.target_value),
        python_plain_value(&issue.comparison_value),
        python_plain_value(&issue.difference_value),
        issue.issue_id.clone(),
        issue.institution_feedback.clone(),
        issue.auditor_opinion.clone(),
    ]
}

fn python_plain_value(value: &Value) -> String {
    match value {
        Value::Null => String::new(),
        Value::String(value) => value.clone(),
        Value::Bool(true) => "True".to_owned(),
        Value::Bool(false) => "False".to_owned(),
        Value::Number(value) => value.to_string(),
        Value::Array(_) | Value::Object(_) => value.to_string(),
    }
}

fn result_link(issue: &Issue) -> Result<Option<ResultLink>> {
    if issue.audit_file.is_empty() || issue.sheet_name.is_empty() || issue.target_cell.is_empty() {
        return Ok(None);
    }
    let Ok(path) = fs::canonicalize(&issue.audit_file) else {
        return Ok(None);
    };
    if !path.is_file() {
        return Ok(None);
    }
    let location = format!(
        "'{}'!{}",
        issue.sheet_name.replace('\'', "''"),
        issue.target_cell
    );
    Ok(Some(ResultLink {
        target_uri: absolute_file_uri(&path)?,
        location,
        display: issue.target_cell.clone(),
    }))
}

fn safe_sheet_name(value: &str) -> String {
    let safe = value
        .chars()
        .map(|character| {
            if "\\[]:*?/".contains(character) {
                '_'
            } else {
                character
            }
        })
        .take(31)
        .collect::<String>();
    if safe.is_empty() {
        S1_RESULT_SHEET.to_owned()
    } else {
        safe
    }
}

fn absolute_file_uri(path: &Path) -> Result<String> {
    let normalized = path.to_string_lossy().replace('\\', "/");
    absolute_file_uri_from_path(&normalized)
}

fn absolute_file_uri_from_path(path: &str) -> Result<String> {
    let path = if path.starts_with("//?/UNC/") {
        format!("//{}", &path[8..])
    } else if path.starts_with("//?/") {
        path[4..].to_owned()
    } else {
        path.to_owned()
    };

    if path.starts_with("//") {
        let (authority, share_path) = path[2..]
            .split_once('/')
            .context("UNC 审核副本路径缺少共享目录")?;
        ensure!(
            !authority.is_empty()
                && authority
                    .bytes()
                    .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'-')),
            "UNC 审核副本路径的服务器名称无效"
        );
        Ok(format!(
            "file://{authority}/{}",
            percent_encode_uri_path(share_path)
        ))
    } else if path.as_bytes().get(1) == Some(&b':') {
        ensure!(
            path.as_bytes().first().is_some_and(u8::is_ascii_alphabetic)
                && path.as_bytes().get(2) == Some(&b'/'),
            "Windows 审核副本路径必须是绝对盘符路径"
        );
        Ok(format!("file:///{}", percent_encode_uri_path(&path)))
    } else {
        ensure!(path.starts_with('/'), "审核副本路径必须是绝对文件系统路径");
        Ok(format!("file://{}", percent_encode_uri_path(&path)))
    }
}

fn percent_encode_uri_path(path: &str) -> String {
    let mut result = String::with_capacity(path.len());
    for byte in path.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'.' | b'_' | b'~' | b'/' | b':') {
            result.push(byte as char);
        } else {
            result.push_str(&format!("%{byte:02X}"));
        }
    }
    result
}

fn temporary_path(output_path: &Path) -> Result<PathBuf> {
    let parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let stem = output_path
        .file_stem()
        .context("S1 结果工作簿输出路径缺少文件名")?
        .to_string_lossy();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("系统时间早于 Unix epoch")?
        .as_nanos();
    Ok(parent.join(format!(
        ".{stem}.rust-audit-{}-{nonce}.tmp.xlsx",
        std::process::id()
    )))
}

fn validate_result_workbook(path: &Path, sheet_name: &str, rows: &[Vec<String>]) -> Result<()> {
    let workbook = read_workbook(path).context("重新打开 S1 结果工作簿")?;
    ensure!(
        workbook.sheets.len() == 1 && workbook.sheets[0].name == sheet_name,
        "S1 结果工作簿的工作表名称或数量不匹配"
    );
    let sheet = &workbook.sheets[0];
    ensure!(sheet.formulas.is_empty(), "S1 结果工作簿不应包含公式");
    ensure!(sheet.merges.is_empty(), "S1 结果工作簿不应包含合并单元格");

    let mut expected_values = Vec::new();
    expected_values.extend(
        S1_RESULT_HEADERS
            .iter()
            .enumerate()
            .map(|(column, value)| (cell_address(0, column), (*value).to_owned())),
    );
    for (row_index, row) in rows.iter().enumerate() {
        expected_values.extend(row.iter().enumerate().filter_map(|(column, value)| {
            (!value.is_empty()).then(|| (cell_address(row_index + 1, column), value.clone()))
        }));
    }
    ensure!(
        sheet.values.len() == expected_values.len(),
        "S1 结果工作簿单元格数量不匹配：预期 {}，实际 {}",
        expected_values.len(),
        sheet.values.len()
    );
    for (address, expected) in expected_values {
        let actual = sheet
            .values
            .get(&address)
            .with_context(|| format!("S1 结果工作簿缺少单元格 {address}"))?;
        ensure!(
            actual.get("type").and_then(Value::as_str) == Some("string")
                && actual.get("value").and_then(Value::as_str) == Some(expected.as_str()),
            "S1 结果工作簿单元格 {address} 内容不匹配"
        );
    }
    Ok(())
}

fn cell_address(row: usize, column: usize) -> String {
    let mut number = column + 1;
    let mut letters = Vec::new();
    while number > 0 {
        let remainder = (number - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        number = (number - 1) / 26;
    }
    letters.reverse();
    format!("{}{}", letters.into_iter().collect::<String>(), row + 1)
}

#[cfg(test)]
mod tests {
    use super::{absolute_file_uri_from_path, percent_encode_uri_path};

    #[test]
    fn encodes_absolute_posix_file_paths_as_uris() {
        assert_eq!(
            absolute_file_uri_from_path("/mnt/c/Audit/检查 #1.xlsx").unwrap(),
            "file:///mnt/c/Audit/%E6%A3%80%E6%9F%A5%20%231.xlsx"
        );
    }

    #[test]
    fn normalizes_windows_drive_and_extended_file_paths() {
        assert_eq!(
            absolute_file_uri_from_path("C:/Audit/检查.xlsx").unwrap(),
            "file:///C:/Audit/%E6%A3%80%E6%9F%A5.xlsx"
        );
        assert_eq!(
            absolute_file_uri_from_path("//?/C:/Audit/a%?.xlsx").unwrap(),
            "file:///C:/Audit/a%25%3F.xlsx"
        );
    }

    #[test]
    fn normalizes_unc_paths_as_file_uris() {
        assert_eq!(
            absolute_file_uri_from_path("//audit-server/shared/审计副本.xlsx").unwrap(),
            "file://audit-server/shared/%E5%AE%A1%E8%AE%A1%E5%89%AF%E6%9C%AC.xlsx"
        );
    }

    #[test]
    fn only_uri_path_bytes_are_percent_encoded() {
        assert_eq!(
            percent_encode_uri_path("C:/a b#c.xlsx"),
            "C:/a%20b%23c.xlsx"
        );
    }
}
