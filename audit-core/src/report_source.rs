//! Audit-specific workbook and worksheet identity derived from source names.

use std::path::Path;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ReportSourceIdentity {
    pub institution_id: String,
    pub institution_name: String,
    pub form_key: String,
    pub form_code: String,
    pub data_date: String,
    pub recognized: bool,
}

/// Resolve the source identity using only the file stem and worksheet name.
///
/// This mirrors the source project's `base_audit.report_source` protocol,
/// including the distinct `form_key` used for pairing and `form_code` used for
/// business import/configuration.
pub fn report_source_identity(path: &Path, sheet_name: &str) -> ReportSourceIdentity {
    let stem = path
        .file_stem()
        .map(|value| value.to_string_lossy().into_owned())
        .unwrap_or_default();
    let parts: Vec<String> = stem.split('#').map(|part| part.trim().to_owned()).collect();
    let sheet_name = sheet_name.trim();

    if parts.len() >= 4 && parts[0].eq_ignore_ascii_case("banks") {
        return ReportSourceIdentity {
            institution_id: sheet_name.to_owned(),
            institution_name: String::new(),
            form_key: extract_form_code(&parts[3]).unwrap_or_else(|| parts[3].clone()),
            form_code: extract_form_code(&parts[3]).unwrap_or_default(),
            data_date: parts[1].clone(),
            recognized: true,
        };
    }

    if parts.len() >= 5 {
        let reports = parts[0].eq_ignore_ascii_case("reports");
        let institution_id = if reports {
            parts[1].clone()
        } else {
            parts[0].clone()
        };
        let data_date = if reports {
            parts[2].clone()
        } else {
            parts[1].clone()
        };
        let fallback_form = if reports { "" } else { parts[3].as_str() };
        let sheet_form = extract_form_code(sheet_name);
        let default_sheet = ["sheet", "sheet1", "sheet2", "sheet3"]
            .iter()
            .any(|name| sheet_name.eq_ignore_ascii_case(name));
        let form_key = sheet_form.clone().unwrap_or_else(|| {
            if default_sheet && !fallback_form.is_empty() {
                fallback_form.to_owned()
            } else {
                sheet_name.to_owned()
            }
        });
        return ReportSourceIdentity {
            institution_id,
            institution_name: parts[4..].join("#"),
            form_key,
            form_code: sheet_form
                .or_else(|| extract_form_code(fallback_form))
                .unwrap_or_default(),
            data_date,
            recognized: true,
        };
    }

    ReportSourceIdentity {
        institution_id: String::new(),
        institution_name: stem,
        form_key: sheet_name.to_owned(),
        form_code: extract_form_code(sheet_name).unwrap_or_default(),
        data_date: String::new(),
        recognized: false,
    }
}

fn extract_form_code(value: &str) -> Option<String> {
    let characters: Vec<char> = value.chars().collect();
    if characters.len() < 5 {
        return None;
    }
    for start in 0..=characters.len() - 5 {
        let end = start + 5;
        if characters[start] != '2'
            || characters[start + 1] != '0'
            || !characters[start + 2..end]
                .iter()
                .all(|character| character.is_ascii_digit())
        {
            continue;
        }
        let left_is_digit = start > 0 && characters[start - 1].is_ascii_digit();
        let right_is_digit = end < characters.len() && characters[end].is_ascii_digit();
        if !left_is_digit && !right_is_digit {
            return Some(characters[start..end].iter().collect());
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::report_source_identity;
    use std::path::Path;

    #[test]
    fn identity_matches_standard_reports_banks_and_unrecognized_names() {
        let standard =
            report_source_identity(Path::new("9144#2026-06-30#x#20202#甲银行.xlsx"), "Sheet1");
        assert_eq!(standard.institution_id, "9144");
        assert_eq!(standard.institution_name, "甲银行");
        assert_eq!(standard.form_key, "20202");
        assert_eq!(standard.form_code, "20202");
        assert_eq!(standard.data_date, "2026-06-30");

        let reports = report_source_identity(
            Path::new("reports#9144#2026-06-30#01#甲银行.xlsx"),
            "Sheet1",
        );
        assert_eq!(reports.institution_id, "9144");
        assert_eq!(reports.institution_name, "甲银行");
        assert_eq!(reports.form_key, "Sheet1");

        let banks = report_source_identity(Path::new("banks#2026-06-30#01#20202.xlsx"), "9144");
        assert_eq!(banks.institution_id, "9144");
        assert_eq!(banks.institution_name, "");
        assert_eq!(banks.form_key, "20202");
        assert_eq!(banks.form_code, "20202");

        let unknown = report_source_identity(Path::new("legacy.xlsx"), "20203 明细");
        assert_eq!(unknown.institution_name, "legacy");
        assert_eq!(unknown.form_key, "20203 明细");
        assert!(!unknown.recognized);
    }
}
