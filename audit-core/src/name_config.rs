//! Pure value helpers from `base_audit.name_config`.
//!
//! This module does not load or modify the Excel configuration workbook. It
//! only represents already-read feature mappings and applies their scalar
//! matching rules.

use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct FeatureMapping {
    pub name: String,
    pub feature_type: String,
    pub range_names: Vec<String>,
    pub workbook_limited: bool,
    pub workbook_keyword: String,
    pub remark: String,
    #[serde(default)]
    pub output: String,
    #[serde(default)]
    pub input_note: String,
}

/// Split a delimited named-range cell using the source project's separators.
pub fn split_names(value: &str) -> Vec<String> {
    value
        .split(|character| matches!(character, '、' | '，' | ',' | '；' | ';' | '\n'))
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .map(str::to_owned)
        .collect()
}

/// Interpret the source project's affirmative workbook-limit values.
pub fn is_limited(value: &str) -> bool {
    matches!(
        value.to_lowercase().as_str(),
        "是" | "y" | "yes" | "true" | "1"
    )
}

/// Return mappings whose feature type matches exactly.
pub fn features_of_type<'a>(
    features: &'a [FeatureMapping],
    feature_type: &str,
) -> Vec<&'a FeatureMapping> {
    features
        .iter()
        .filter(|feature| feature.feature_type == feature_type)
        .collect()
}

/// Match an exact named range or an underscore-suffixed variant.
pub fn matches_named_range(mapping: &FeatureMapping, candidate: &str) -> bool {
    let simple = candidate
        .rsplit('!')
        .next()
        .unwrap_or(candidate)
        .trim()
        .trim_matches('\'');
    let folded = simple.to_lowercase();

    mapping.range_names.iter().any(|name| {
        let name = name.to_lowercase();
        folded == name || folded.starts_with(&format!("{name}_"))
    })
}
