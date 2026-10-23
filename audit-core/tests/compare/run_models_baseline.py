#!/usr/bin/env python3
"""Export deterministic JSON vectors from the read-only Open-Data-Audit models."""

from __future__ import annotations

import argparse
import dataclasses
import importlib.util
import json
from pathlib import Path
import sys


def load_models(source_root: Path):
    module_path = source_root / "core" / "src" / "base_audit" / "models.py"
    spec = importlib.util.spec_from_file_location("audit_models_baseline", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load source model module: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def jsonable(value):
    if dataclasses.is_dataclass(value):
        return {field.name: jsonable(getattr(value, field.name)) for field in dataclasses.fields(value)}
    if isinstance(value, Path):
        return value.as_posix()
    if isinstance(value, tuple):
        return [jsonable(item) for item in value]
    if isinstance(value, list):
        return [jsonable(item) for item in value]
    if isinstance(value, dict):
        return {key: jsonable(item) for key, item in value.items()}
    return value


def build_report(models):
    rule = models.AuditRule(
        rule_id="R-001",
        enabled=True,
        report_code="REP-A",
        sheet_name="汇总",
        formula_cell="C8",
        target_cell="D8",
        severity="warning",
        message="结果不一致",
        copy_range="A1:D12",
        result_mode="nonempty",
        value_from_result=True,
    )
    rule_defaults = models.AuditRule(
        rule_id="R-default",
        enabled=False,
        report_code="REP-B",
        sheet_name="明细",
        formula_cell="A2",
        target_cell="B2",
        severity="info",
        message="默认值检查",
    )
    copy_range = models.CopyRange(sheet_name="汇总", address="A1:D12")
    template = models.TemplateDefinition(
        rules=[rule],
        copy_ranges=[copy_range],
        structured=True,
        structure_ranges=[models.CopyRange("结构", "A1:B3")],
        extraction_ranges=[models.CopyRange("数据", "A5:H20")],
    )
    preflight_item = models.PreflightItem(
        category="template",
        level="warning",
        status="mismatch",
        sheet_name="汇总",
        location="C8",
        message="标签未匹配",
    )
    structure_check = models.StructureCheck(
        sheet_name="汇总",
        cell_address="A1",
        expected_value="机构代码",
        actual_value="机构编号",
        matched=False,
    )
    source_match = models.SourceMatch(
        matched=False,
        score=0.5,
        matched_labels=1,
        checked_labels=2,
        missing_sheets=("说明",),
        details="模板结构不完整",
        structure_checks=(structure_check,),
    )
    preflight = models.PreflightRunResult(
        report_path=Path("output/preflight.xlsx"),
        total_files=3,
        matched_files=2,
        template_warnings=1,
        source_matches=((Path("input/source-01.xlsx"), source_match),),
        batch_id="batch-001",
        log_path=Path("logs/preflight.log"),
    )
    issue = models.Issue(
        issue_id="issue-001",
        period="2026-09",
        batch_id="batch-001",
        audit_time="2026-10-01 09:30:00",
        triggered=True,
        status="new",
        first_seen_period="2026-09",
        previous_seen_period="2026-08",
        consecutive_count=2,
        org_code="0001",
        org_name="样本机构",
        report_code="REP-A",
        sheet_name="汇总",
        rule_id="R-001",
        severity="warning",
        formula_cell="C8",
        target_cell="D8",
        target_value=100.25,
        formula_result=None,
        message="结果不一致",
        source_file="input/source-01.xlsx",
        audit_file="output/source-01_audit.xlsx",
        check_field="余额",
        check_type="difference",
        comparison_value="100.25",
        reference_value="99.00",
        difference_value="1.25",
        detail="样本差额",
        institution_feedback="待反馈",
        auditor_opinion="待复核",
    )
    file_result = models.FileAuditResult(
        source_path=Path("input/source-01.xlsx"),
        audit_path=Path("output/source-01_audit.xlsx"),
        org_code="0001",
        org_name="样本机构",
        issues=[issue],
        error="",
    )
    skipped_file = models.FileAuditResult(
        source_path=Path("input/source-02.xlsx"),
        audit_path=None,
        org_code="0002",
        org_name="待匹配机构",
        error="报送文件与模板不匹配：缺少工作表",
    )
    audit = models.AuditRunResult(
        batch_id="batch-001",
        period="2026-09",
        output_dir=Path("output"),
        current_issues=[issue],
        resolved_issues=[],
        files=[file_result, skipped_file],
        summary_path=Path("output/summary.xlsx"),
        history_path=Path("output/history.xlsx"),
        preflight_path=Path("output/preflight.xlsx"),
        performance_lines=("读取 0.10s", "规则 0.20s"),
        log_path=Path("logs/audit.log"),
        copies={"source-copy": Path("output/copies")},
    )

    rule_default_values = models.AuditRule(
        rule_id="R-default",
        enabled=False,
        report_code="REP-B",
        sheet_name="明细",
        formula_cell="A2",
        target_cell="B2",
        severity="info",
        message="默认值检查",
    )
    template_defaults = models.TemplateDefinition(rules=[], copy_ranges=[], structured=False)
    source_match_defaults = models.SourceMatch(
        matched=True, score=1.0, matched_labels=1, checked_labels=1
    )
    preflight_defaults = models.PreflightRunResult(
        report_path=None, total_files=0, matched_files=0, template_warnings=0
    )
    issue_defaults = models.Issue(
        issue_id="issue-default",
        period="2026-09",
        batch_id="batch-default",
        audit_time="",
        triggered=False,
        status="new",
        first_seen_period="",
        previous_seen_period="",
        consecutive_count=0,
        org_code="",
        org_name="",
        report_code="",
        sheet_name="",
        rule_id="",
        severity="",
        formula_cell="",
        target_cell="",
        target_value=None,
        formula_result=None,
        message="",
        source_file="",
        audit_file="",
    )
    file_defaults = models.FileAuditResult(
        source_path=Path("input/minimal.xlsx"),
        audit_path=None,
        org_code="",
        org_name="",
    )
    audit_defaults = models.AuditRunResult(
        batch_id="batch-default",
        period="2026-09",
        output_dir=Path("output"),
        current_issues=[],
        resolved_issues=[],
        files=[],
    )

    def default_record(value):
        args = {
            field.name: jsonable(getattr(value, field.name))
            for field in dataclasses.fields(value)
            if field.default is dataclasses.MISSING
            and field.default_factory is dataclasses.MISSING
        }
        return {"args": args, "value": jsonable(value)}

    return {
        "audit_rule": jsonable(rule),
        "audit_rule_defaults": jsonable(rule_defaults),
        "copy_range": jsonable(copy_range),
        "template_definition": jsonable(template),
        "preflight_item": jsonable(preflight_item),
        "structure_check": jsonable(structure_check),
        "source_match": jsonable(source_match),
        "preflight_run_result": {
            "value": jsonable(preflight),
            "failed_files": preflight.failed_files,
            "summary_text": preflight.summary_text(),
        },
        "issue": jsonable(issue),
        "file_audit_result": jsonable(file_result),
        "audit_run_result": {
            "value": jsonable(audit),
            "successful_files": audit.successful_files,
            "failed_files": audit.failed_files,
            "summary_text": audit.summary_text(),
        },
        "constructor_defaults": {
            "audit_rule": default_record(rule_default_values),
            "template_definition": default_record(template_defaults),
            "source_match": default_record(source_match_defaults),
            "preflight_run_result": default_record(preflight_defaults),
            "issue": default_record(issue_defaults),
            "file_audit_result": default_record(file_defaults),
            "audit_run_result": default_record(audit_defaults),
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    report = build_report(load_models(args.source_root))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {len(report)} model parity records to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
