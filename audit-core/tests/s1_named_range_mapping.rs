use audit_core::{
    models::CopyRange,
    name_config::{resolve_named_ranges_for_mappings, FeatureMapping, NamedRangeArea},
};

#[test]
fn resolves_configured_workbook_and_sheet_scoped_names_in_order() {
    let mappings = [FeatureMapping {
        name: "条件格式结果提取".to_owned(),
        feature_type: "汇总_条件格式结果提取".to_owned(),
        range_names: vec!["条件格式区域".to_owned()],
        ..FeatureMapping::default()
    }];
    let areas = [
        NamedRangeArea {
            name: "其他区域".to_owned(),
            sheet_name: "报表一".to_owned(),
            address: "A1:A3".to_owned(),
        },
        NamedRangeArea {
            name: "条件格式区域_模板".to_owned(),
            sheet_name: "报表二".to_owned(),
            address: "$B$2:$D$9".to_owned(),
        },
        NamedRangeArea {
            name: "条件格式区域".to_owned(),
            sheet_name: "报表一".to_owned(),
            address: "C5:C8".to_owned(),
        },
        NamedRangeArea {
            name: "条件格式区域_重复".to_owned(),
            sheet_name: "报表二".to_owned(),
            address: "b2:d9".to_owned(),
        },
        NamedRangeArea {
            name: "条件格式区域".to_owned(),
            sheet_name: "模板独有表".to_owned(),
            address: "A1:A5".to_owned(),
        },
        NamedRangeArea {
            name: "条件格式区域额外".to_owned(),
            sheet_name: "报表一".to_owned(),
            address: "A20:A25".to_owned(),
        },
        NamedRangeArea {
            name: "条件格式区域_空".to_owned(),
            sheet_name: "报表一".to_owned(),
            address: "  $  ".to_owned(),
        },
    ];

    let resolved = resolve_named_ranges_for_mappings(
        &mappings,
        &areas,
        &["报表一".to_owned(), "报表二".to_owned()],
    );

    assert_eq!(
        resolved,
        [
            CopyRange {
                sheet_name: "报表二".to_owned(),
                address: "B2:D9".to_owned(),
            },
            CopyRange {
                sheet_name: "报表一".to_owned(),
                address: "C5:C8".to_owned(),
            },
        ]
    );
}
