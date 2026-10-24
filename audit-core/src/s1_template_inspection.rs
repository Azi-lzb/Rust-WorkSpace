use anyhow::{ensure, Context, Result};
use excel_reader::{
    read_defined_names, read_workbook, DefinedNameRecord, DefinedNameScope, WorkbookDump,
};
use serde::Serialize;
use std::{collections::BTreeMap, path::Path, time::Instant};

const RANGE_FAMILIES: [&str; 3] = ["校验区域", "表结构区域", "条件格式区域"];
const INTERNAL_NAME_PREFIXES: [&str; 3] = ["_filterdatabase", "print_area", "print_titles"];
const ERROR_TOKENS: [&str; 9] = [
    "#REF!", "#VALUE!", "#NAME?", "#N/A", "#DIV/0!", "#NUM!", "#NULL!", "#CALC!", "#SPILL!",
];
pub const INSPECTION_ENGINE: &str = "calamine+namespace_aware_ooxml";

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct S1TemplateInspectionReport {
    pub template: String,
    pub formula_count: usize,
    pub errors: Vec<InspectionError>,
    pub named_range_groups: Vec<NamedRangeGroup>,
    pub internal_names: Vec<NamedRangeItem>,
    pub other_names: Vec<NamedRangeItem>,
}

#[derive(Clone, Debug, Serialize)]
pub struct InspectionError {
    pub kind: &'static str,
    pub location: String,
    pub token: String,
    pub message: String,
}

#[derive(Clone, Debug, Serialize)]
pub struct NamedRangeGroup {
    pub family: &'static str,
    pub count: usize,
    pub items: Vec<NamedRangeItem>,
}

#[derive(Clone, Debug, Serialize)]
pub struct NamedRangeItem {
    pub name: String,
    pub range: String,
    pub scope: String,
}

pub fn inspect_s1_template(
    template_path: impl AsRef<Path>,
) -> Result<(S1TemplateInspectionReport, u128)> {
    let started = Instant::now();
    let template_path = template_path.as_ref();
    ensure!(
        template_path.is_file(),
        "模板文件不存在：{}",
        template_path.display()
    );
    ensure!(
        template_path
            .extension()
            .and_then(|ext| ext.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx") || ext.eq_ignore_ascii_case("xlsm")),
        "S1 模板公式检查目前支持 .xlsx 和 .xlsm"
    );
    let workbook = read_workbook(template_path).context("读取 S1 模板公式文本")?;
    let names = read_defined_names(template_path).context("读取 S1 模板命名区域")?;
    Ok((
        build_report(template_path, &workbook, &names),
        started.elapsed().as_millis(),
    ))
}

fn build_report(
    template_path: &Path,
    workbook: &WorkbookDump,
    defined_names: &[DefinedNameRecord],
) -> S1TemplateInspectionReport {
    let sheet_order = workbook
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| (sheet.name.as_str(), index))
        .collect::<BTreeMap<_, _>>();
    let mut ordered_names = defined_names.iter().enumerate().collect::<Vec<_>>();
    ordered_names.sort_by_key(|(index, name)| match &name.scope {
        DefinedNameScope::Workbook => (0usize, 0usize, *index),
        DefinedNameScope::Worksheet {
            local_sheet_id,
            sheet_name,
            ..
        } => (
            1,
            sheet_name
                .as_deref()
                .and_then(|name| sheet_order.get(name).copied())
                .unwrap_or(*local_sheet_id),
            *index,
        ),
        DefinedNameScope::UnresolvedWorksheet { local_sheet_id } => (2, *local_sheet_id, *index),
        DefinedNameScope::InvalidWorksheetScope { .. } => (3, 0, *index),
    });

    let mut groups = RANGE_FAMILIES
        .iter()
        .map(|family| NamedRangeGroup {
            family,
            count: 0,
            items: Vec::new(),
        })
        .collect::<Vec<_>>();
    let mut internal_names = Vec::new();
    let mut other_names = Vec::new();
    let mut errors = Vec::new();

    for (_, name) in ordered_names {
        // The Python/openpyxl source workflow does not expose reserved _xlnm.* definitions.
        if name.name.to_lowercase().starts_with("_xlnm") {
            continue;
        }
        if name.name.is_empty() {
            continue;
        }
        if let DefinedNameScope::Worksheet {
            sheet_kind: Some(kind),
            ..
        } = &name.scope
        {
            if !kind.eq_ignore_ascii_case("worksheet") {
                continue;
            }
        }
        let item = NamedRangeItem {
            name: name.name.clone(),
            range: name.raw_reference.clone(),
            scope: scope_label(&name.scope),
        };
        let class = classify_name(&name.name);
        match class {
            NameClass::Internal => internal_names.push(item),
            NameClass::Other => other_names.push(item),
            NameClass::Family(index) => groups[index].items.push(item),
        }
        if !matches!(class, NameClass::Internal) {
            if let Some(token) = first_error_token(&name.raw_reference) {
                errors.push(InspectionError {
                    kind: "命名区域",
                    location: name.name.clone(),
                    message: format!("名称定义引用了错误值：{token}"),
                    token,
                });
            }
        }
    }

    let mut formula_count = 0;
    for sheet in &workbook.sheets {
        for (address, formula) in &sheet.formulas {
            formula_count += 1;
            if let Some(token) = first_error_token(formula) {
                errors.push(InspectionError {
                    kind: "公式",
                    location: format!("{}!{address}", sheet.name),
                    message: format!("公式文本含错误引用：{token}"),
                    token,
                });
            }
        }
    }
    for group in &mut groups {
        group.count = group.items.len();
    }
    S1TemplateInspectionReport {
        template: template_path.to_string_lossy().into_owned(),
        formula_count,
        errors,
        named_range_groups: groups,
        internal_names,
        other_names,
    }
}

#[derive(Clone, Copy)]
enum NameClass {
    Family(usize),
    Internal,
    Other,
}

fn classify_name(name: &str) -> NameClass {
    let lowered = name.to_lowercase();
    if lowered.starts_with("_xlnm")
        || INTERNAL_NAME_PREFIXES
            .iter()
            .any(|prefix| lowered == *prefix || lowered.starts_with(prefix))
    {
        return NameClass::Internal;
    }
    if let Some(index) = RANGE_FAMILIES
        .iter()
        .position(|family| name == *family || name.starts_with(family))
    {
        return NameClass::Family(index);
    }
    if name.contains("校验区域") {
        return NameClass::Family(0);
    }
    NameClass::Other
}

fn scope_label(scope: &DefinedNameScope) -> String {
    match scope {
        DefinedNameScope::Workbook => "工作簿".to_owned(),
        DefinedNameScope::Worksheet {
            sheet_name: Some(name),
            ..
        } => name.clone(),
        DefinedNameScope::Worksheet { local_sheet_id, .. }
        | DefinedNameScope::UnresolvedWorksheet { local_sheet_id } => {
            format!("未知工作表（localSheetId={local_sheet_id}）")
        }
        DefinedNameScope::InvalidWorksheetScope { raw_local_sheet_id } => {
            format!("无效工作表作用域（localSheetId={raw_local_sheet_id}）")
        }
    }
}

fn first_error_token(expression: &str) -> Option<String> {
    let bytes = expression.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            if let Some(close_offset) = bytes[index + 1..].iter().position(|byte| *byte == b'"') {
                index += close_offset + 2;
            } else {
                index += 1;
            }
            continue;
        }
        if bytes[index] == b'#' {
            for token in ERROR_TOKENS {
                let end = index + token.len();
                if end <= bytes.len()
                    && bytes[index..end]
                        .iter()
                        .zip(token.as_bytes())
                        .all(|(left, right)| left.to_ascii_uppercase() == *right)
                {
                    return Some(token.to_owned());
                }
            }
        }
        index += 1;
    }
    None
}

#[cfg(test)]
mod tests {
    use super::{classify_name, first_error_token, inspect_s1_template, NameClass, RANGE_FAMILIES};
    use excel_writer::Workbook;
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    fn temp_path() -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        std::env::temp_dir().join(format!(
            "audit_s1_inspection_{}_{}.xlsx",
            std::process::id(),
            nonce
        ))
    }

    #[test]
    fn error_scanner_ignores_quoted_strings_and_finds_first_token() {
        assert_eq!(
            first_error_token("=IF(A1=\"#N/A\",1,#ref!)"),
            Some("#REF!".to_owned())
        );
        assert_eq!(first_error_token("=\"say \"\"#N/A\"\"\"&1"), None);
        assert_eq!(first_error_token("=SUM(A1:A3)"), None);
        assert_eq!(
            first_error_token("=\"unterminated #N/A"),
            Some("#N/A".to_owned())
        );
    }

    #[test]
    fn name_classifier_keeps_business_families_and_internal_names_distinct() {
        for (index, family) in RANGE_FAMILIES.iter().enumerate() {
            assert!(matches!(classify_name(family), NameClass::Family(actual) if actual == index));
            assert!(
                matches!(classify_name(&format!("{family}_附加")), NameClass::Family(actual) if actual == index)
            );
        }
        assert!(matches!(
            classify_name("本地校验区域1"),
            NameClass::Family(0)
        ));
        assert!(matches!(
            classify_name("_xlnm.Print_Area"),
            NameClass::Internal
        ));
        assert!(matches!(
            classify_name("_FilterDatabase"),
            NameClass::Internal
        ));
        assert!(matches!(classify_name("业务自定义名称"), NameClass::Other));
    }

    #[test]
    fn inspection_report_matches_python_contract_for_synthetic_workbook() {
        let path = temp_path();
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").unwrap();
        sheet.write_formula(0, 0, "=1+#REF!").unwrap();
        sheet
            .write_formula(1, 0, "=IF(A1=\"#N/A 文本\",1,0)")
            .unwrap();
        for (name, reference) in [
            ("校验区域_1", "='检查表'!$A$1:$A$2"),
            ("表结构区域_附加", "='检查表'!$A$1:$A$2"),
            ("条件格式区域额外", "='检查表'!$A$1:$A$2"),
            ("历史本地校验区域1", "='检查表'!$A$1"),
            ("Print_Area", "='检查表'!$A$1:$A$2"),
            ("自定义名称", "='检查表'!$A$1"),
        ] {
            workbook.define_name(name, reference).unwrap();
        }
        workbook.save(&path).unwrap();
        let (report, _) = inspect_s1_template(&path).unwrap();
        fs::remove_file(&path).unwrap();

        assert_eq!(report.template, path.to_string_lossy());
        assert_eq!(report.formula_count, 2);
        assert_eq!(report.errors.len(), 1);
        assert_eq!(report.errors[0].kind, "公式");
        assert_eq!(report.errors[0].location, "检查表!A1");
        assert_eq!(report.errors[0].token, "#REF!");
        assert_eq!(report.named_range_groups.len(), 3);
        assert_eq!(report.named_range_groups[0].family, "校验区域");
        assert_eq!(report.named_range_groups[0].count, 2);
        assert_eq!(report.named_range_groups[1].count, 1);
        assert_eq!(report.named_range_groups[2].count, 1);
        assert_eq!(report.internal_names[0].name, "Print_Area");
        assert_eq!(report.other_names[0].name, "自定义名称");
        assert_eq!(report.named_range_groups[0].items[0].scope, "工作簿");
    }
}
