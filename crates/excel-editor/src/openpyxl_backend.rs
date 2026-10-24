//! Isolated Python/openpyxl adapter for batch edits to existing XLSX files.
//!
//! The adapter is opt-in and evidence gated. It launches one bundled Python
//! process per workbook plan; the Python runtime and worker are never linked
//! into the Rust process.

use crate::{
    BackendCapability, BackendFailure, EditBackend, EditOperation, EditOperationKind,
    ExcelEditPlan, OutputValidator, PreservationRequirement, RangeStyleMode, WorkbookFormat,
};
use anyhow::{bail, Context, Result};
use excel_reader::{read_workbook, SheetDump};
use serde::Deserialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, File, OpenOptions},
    io::Read,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    thread,
    time::{Duration, Instant},
};

static SCRATCH_SEQUENCE: AtomicU64 = AtomicU64::new(0);
const DEFAULT_TIMEOUT: Duration = Duration::from_secs(120);

#[derive(Clone, Debug)]
pub struct OpenpyxlWorkerConfig {
    pub python_executable: PathBuf,
    pub worker_script: PathBuf,
    pub timeout: Duration,
}

impl OpenpyxlWorkerConfig {
    pub fn new(
        python_executable: impl Into<PathBuf>,
        worker_script: impl Into<PathBuf>,
    ) -> Result<Self> {
        let config = Self {
            python_executable: python_executable.into(),
            worker_script: worker_script.into(),
            timeout: DEFAULT_TIMEOUT,
        };
        config.validate_paths()?;
        Ok(config)
    }

    pub fn with_timeout(mut self, timeout: Duration) -> Result<Self> {
        anyhow::ensure!(
            !timeout.is_zero(),
            "openpyxl worker timeout must be positive"
        );
        self.timeout = timeout;
        Ok(self)
    }

    /// Resolve the isolated runtime layout installed with a product binary.
    /// For `/usr/bin/pytools`, this returns paths under
    /// `/usr/libexec/pytools/openpyxl/`; no PATH lookup is performed.
    pub fn for_product_binary(binary: impl AsRef<Path>, product: &str) -> Option<Self> {
        if product.is_empty()
            || !product
                .bytes()
                .all(|byte| byte.is_ascii_lowercase() || byte.is_ascii_digit() || byte == b'-')
        {
            return None;
        }
        let binary = binary.as_ref();
        let installation_root = binary.parent()?.parent()?;
        let base = installation_root
            .join("libexec")
            .join(product)
            .join("openpyxl");
        Some(Self {
            python_executable: base.join("runtime/bin/python3"),
            worker_script: base.join("worker/openpyxl_worker.py"),
            timeout: DEFAULT_TIMEOUT,
        })
    }

    fn validate_paths(&self) -> Result<()> {
        anyhow::ensure!(
            self.python_executable.is_absolute(),
            "Python executable path must be absolute"
        );
        anyhow::ensure!(
            self.worker_script.is_absolute(),
            "openpyxl worker path must be absolute"
        );
        anyhow::ensure!(
            self.timeout > Duration::ZERO,
            "openpyxl worker timeout must be positive"
        );
        Ok(())
    }

    fn available(&self) -> bool {
        if self.validate_paths().is_err()
            || !self.python_executable.is_file()
            || !self.worker_script.is_file()
        {
            return false;
        }
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            return fs::metadata(&self.python_executable)
                .map(|metadata| metadata.permissions().mode() & 0o111 != 0)
                .unwrap_or(false);
        }
        #[cfg(not(unix))]
        true
    }
}

pub struct OpenpyxlXlsxBackend {
    capability: BackendCapability,
    config: Option<OpenpyxlWorkerConfig>,
    unavailable_reason: Option<String>,
}

impl OpenpyxlXlsxBackend {
    pub fn new(capability: BackendCapability, config: OpenpyxlWorkerConfig) -> Result<Self> {
        Self::build(capability, Some(config))
    }

    /// Register the candidate against the product's bundled runtime paths.
    /// Missing files mark this registration unsupported, allowing the router
    /// to select another independently qualified backend.
    pub fn for_product_binary(
        capability: BackendCapability,
        binary: impl AsRef<Path>,
        product: &str,
    ) -> Result<Self> {
        let config = OpenpyxlWorkerConfig::for_product_binary(binary, product);
        Self::build(capability, config)
    }

    fn build(
        mut capability: BackendCapability,
        config: Option<OpenpyxlWorkerConfig>,
    ) -> Result<Self> {
        anyhow::ensure!(
            capability.engine == crate::EngineId::OpenpyxlWorker,
            "openpyxl backend requires an openpyxl_worker capability"
        );
        anyhow::ensure!(
            capability.format == WorkbookFormat::Xlsx,
            "openpyxl worker supports XLSX only"
        );

        capability
            .operations
            .retain(|operation| supported_operation(*operation));
        // The Rust validator checks these semantic contracts independently.
        // The worker also verifies requested edits, but that does not qualify
        // preservation requirements Rust cannot inspect here.
        capability.preserved.retain(|requirement| {
            matches!(
                requirement,
                PreservationRequirement::CellValuesAndTypes
                    | PreservationRequirement::FormulaText
                    | PreservationRequirement::CommentsAndPositions
                    | PreservationRequirement::Styles
                    | PreservationRequirement::MergedRanges
                    | PreservationRequirement::SheetIdentityAndOrder
                    | PreservationRequirement::ConditionalFormatting
                    | PreservationRequirement::DataValidations
                    | PreservationRequirement::DefinedNames
                    | PreservationRequirement::RowAndColumnDimensions
                    | PreservationRequirement::PageSettings
                    | PreservationRequirement::HeaderFooterText
                    | PreservationRequirement::Relationships
                    | PreservationRequirement::FreezePanes
                    | PreservationRequirement::AutoFilter
            )
        });

        let unavailable_reason = match &config {
            Some(config) if config.available() => None,
            Some(_) => Some(
                "bundled Python/openpyxl worker paths are missing or not executable".to_owned(),
            ),
            None => Some("bundled Python/openpyxl worker paths could not be resolved".to_owned()),
        };
        if unavailable_reason.is_some() {
            capability.status = crate::CapabilityStatus::Unsupported;
        }
        Ok(Self {
            capability,
            config,
            unavailable_reason,
        })
    }
}

impl EditBackend for OpenpyxlXlsxBackend {
    fn capability(&self) -> &BackendCapability {
        &self.capability
    }

    fn apply(&self, plan: &ExcelEditPlan) -> std::result::Result<(), BackendFailure> {
        let Some(config) = self.config.as_ref().filter(|config| config.available()) else {
            return Err(BackendFailure::unsupported_before_write(
                self.unavailable_reason
                    .as_deref()
                    .unwrap_or("openpyxl worker is unavailable"),
            ));
        };
        if plan.format != WorkbookFormat::Xlsx {
            return Err(BackendFailure::unsupported_before_write(
                "openpyxl worker accepts XLSX only",
            ));
        }
        if !plan.input.is_file() {
            return Err(BackendFailure::failed(format!(
                "input workbook is missing: {}",
                plan.input.display()
            )));
        }
        let observed = crate::inspect_xlsx_profile(&plan.input).map_err(|error| {
            BackendFailure::failed(format!(
                "inspect source workbook before openpyxl: {error:#}"
            ))
        })?;
        if !observed.inventory_complete {
            return Err(BackendFailure::unsupported_before_write(
                "source workbook contains unverified OOXML parts",
            ));
        }
        if observed.features_present != plan.workbook.features_present
            || observed.format != plan.workbook.format
        {
            return Err(BackendFailure::unsupported_before_write(
                "ExcelEditPlan workbook profile differs from the source package inventory",
            ));
        }
        if !plan
            .operations
            .iter()
            .all(|operation| supported_operation(operation.kind()))
        {
            return Err(BackendFailure::unsupported_before_write(
                "openpyxl worker does not support the complete operation batch",
            ));
        }
        for operation in &plan.operations {
            if let EditOperation::ImportRange {
                source_workbook, ..
            } = operation
            {
                let profile = crate::inspect_xlsx_profile(source_workbook).map_err(|error| {
                    BackendFailure::failed(format!("inspect ImportRange source: {error:#}"))
                })?;
                if !profile.inventory_complete
                    || !profile
                        .features_present
                        .is_subset(&self.capability.supported_features)
                {
                    return Err(BackendFailure::unsupported_before_write(
                        "ImportRange source contains workbook features outside the openpyxl capability evidence",
                    ));
                }
            }
            if let EditOperation::ImportRange {
                copy_row_heights,
                copy_column_widths,
                copy_row_hidden,
                copy_column_hidden,
                copy_outline_levels,
                copy_all_row_dimensions,
                copy_all_column_dimensions,
                ..
            } = operation
            {
                if *copy_row_heights
                    || *copy_column_widths
                    || *copy_row_hidden
                    || *copy_column_hidden
                    || *copy_outline_levels
                    || *copy_all_row_dimensions
                    || *copy_all_column_dimensions
                {
                    return Err(BackendFailure::unsupported_before_write(
                        "ImportRange dimension-copy flags are not enabled until Rust readback validation covers them",
                    ));
                }
            }
        }

        let parent = plan
            .output
            .parent()
            .filter(|path| !path.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let request_file = match ScratchFile::create(parent, "openpyxl-request") {
            Ok(file) => file,
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "create worker request file: {error:#}"
                )))
            }
        };
        let response_file = match ScratchFile::create(parent, "openpyxl-response") {
            Ok(file) => file,
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "create worker response file: {error:#}"
                )))
            }
        };
        let stdout_file = match ScratchFile::create(parent, "openpyxl-stdout") {
            Ok(file) => file,
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "create worker stdout file: {error:#}"
                )))
            }
        };
        let stderr_file = match ScratchFile::create(parent, "openpyxl-stderr") {
            Ok(file) => file,
            Err(error) => {
                return Err(BackendFailure::failed(format!(
                    "create worker stderr file: {error:#}"
                )))
            }
        };
        let request = serde_json::json!({
            "schema_version": 1,
            "request_id": plan.request_id,
            "format": "xlsx",
            "input": plan.input,
            "output": plan.output,
            "plan": plan,
        });
        let request_bytes = serde_json::to_vec(&request).map_err(|error| {
            BackendFailure::failed(format!("serialize openpyxl request: {error}"))
        })?;
        fs::write(request_file.path(), request_bytes)
            .map_err(|error| BackendFailure::failed(format!("write openpyxl request: {error}")))?;

        let stdout = File::create(stdout_file.path()).map_err(|error| {
            BackendFailure::failed(format!("open worker stdout capture: {error}"))
        })?;
        let stderr = File::create(stderr_file.path()).map_err(|error| {
            BackendFailure::failed(format!("open worker stderr capture: {error}"))
        })?;
        let mut child = Command::new(&config.python_executable)
            .arg("-I")
            .arg(&config.worker_script)
            .arg("--request")
            .arg(request_file.path())
            .arg("--response")
            .arg(response_file.path())
            .stdin(Stdio::null())
            .stdout(Stdio::from(stdout))
            .stderr(Stdio::from(stderr))
            .spawn()
            .map_err(|error| {
                BackendFailure::unsupported_before_write(format!(
                    "start bundled openpyxl worker: {error}"
                ))
            })?;
        let started = Instant::now();
        let status = loop {
            match child.try_wait() {
                Ok(Some(status)) => break status,
                Ok(None) if started.elapsed() < config.timeout => {
                    thread::sleep(Duration::from_millis(10))
                }
                Ok(None) => {
                    let _ = child.kill();
                    let _ = child.wait();
                    return Err(BackendFailure::failed(format!(
                        "openpyxl worker timed out after {} ms; candidate discarded and fallback stopped",
                        config.timeout.as_millis()
                    )));
                }
                Err(error) => {
                    let _ = child.kill();
                    let _ = child.wait();
                    return Err(BackendFailure::failed(format!(
                        "wait for openpyxl worker: {error}"
                    )));
                }
            }
        };
        let response_bytes = fs::read(response_file.path()).unwrap_or_default();
        let response: WorkerResponse =
            serde_json::from_slice(&response_bytes).map_err(|error| {
                let diagnostic = read_diagnostic(stderr_file.path());
                BackendFailure::failed(format!(
                    "invalid openpyxl worker response (exit {:?}): {error}; {diagnostic}",
                    status.code()
                ))
            })?;
        if response.schema_version != 1
            || response.request_id.as_deref() != Some(plan.request_id.as_str())
        {
            return Err(BackendFailure::failed(
                "openpyxl worker response version/request_id mismatch",
            ));
        }
        if response.status == "unsupported_before_write" {
            let error = response
                .error
                .map(|error| format!("{}: {}", error.kind, error.message))
                .unwrap_or_else(|| "worker rejected operation before writing".to_owned());
            return Err(BackendFailure::unsupported_before_write(error));
        }
        if response.status != "applied" || !status.success() {
            let error = response
                .error
                .map(|error| format!("{}: {}", error.kind, error.message))
                .unwrap_or_else(|| format!("worker exited with {:?}", status.code()));
            let diagnostic = read_diagnostic(stderr_file.path());
            return Err(BackendFailure::failed(format!(
                "openpyxl worker failed: {error}; {diagnostic}"
            )));
        }
        if response.operation_count != Some(plan.operations.len()) {
            return Err(BackendFailure::failed(
                "openpyxl worker applied-operation count mismatch",
            ));
        }
        if response.verified_operation_count != Some(plan.operations.len()) {
            return Err(BackendFailure::failed(
                "openpyxl worker verified-operation count mismatch",
            ));
        }
        if response.input_sha256_before.is_none()
            || response.input_sha256_before != response.input_sha256_after
            || response.output_sha256.as_deref().map_or(true, |hash| {
                hash.len() != 64 || !hash.bytes().all(|byte| byte.is_ascii_hexdigit())
            })
        {
            return Err(BackendFailure::failed(
                "openpyxl worker integrity diagnostics are incomplete or inconsistent",
            ));
        }
        Ok(())
    }

    fn validate_output(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        OpenpyxlXlsxValidator.validate(plan, candidate_output)
    }
}

fn supported_operation(operation: EditOperationKind) -> bool {
    matches!(
        operation,
        EditOperationKind::SetCellValue
            | EditOperationKind::SetFormulaText
            | EditOperationKind::SetComment
            | EditOperationKind::SetColumnWidth
            | EditOperationKind::SetRowHeight
            | EditOperationKind::RenameSheet
            | EditOperationKind::CreateSheet
            | EditOperationKind::CreateSheetAt
            | EditOperationKind::SetFreezePanes
            | EditOperationKind::SetAutoFilter
            | EditOperationKind::DeleteSheet
            | EditOperationKind::SetCellFillColor
            | EditOperationKind::CopyRange
            | EditOperationKind::ImportRange
            | EditOperationKind::SetWorkbookCalculationProperties
            | EditOperationKind::SetPageSettings
    )
}

pub struct OpenpyxlXlsxValidator;

impl OutputValidator for OpenpyxlXlsxValidator {
    fn validate(&self, plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
        let source =
            read_workbook(&plan.input).context("read source workbook for openpyxl validation")?;
        let output = read_workbook(candidate_output)
            .context("reopen openpyxl candidate with Calamine/OOXML reader")?;
        let profile = crate::inspect_xlsx_profile(candidate_output)?;
        anyhow::ensure!(
            profile.inventory_complete,
            "openpyxl output contains unknown OOXML parts"
        );
        anyhow::ensure!(
            !output.sheets.is_empty(),
            "openpyxl output contains no worksheets"
        );
        validate_sheet_names(&source.sheets, &output.sheets, &plan.operations)?;
        validate_requested_operations(&source.sheets, &output.sheets, &plan.operations)?;
        validate_calculation_properties(candidate_output, &plan.operations)?;
        if plan.operations.iter().any(needs_umya_readback)
            || plan.preserve.contains(&PreservationRequirement::Styles)
        {
            validate_umya_readback(plan, candidate_output)?;
        }
        validate_preserved_values(&source.sheets, &output.sheets, plan)?;
        validate_preserved_formulas(&source.sheets, &output.sheets, plan)?;
        validate_preserved_comments(&source.sheets, &output.sheets, plan)?;
        validate_preserved_merges(&source.sheets, &output.sheets, plan)?;
        validate_preserved_ooxml_metadata(&plan.input, candidate_output, plan)?;
        validate_requested_sheet_layout(plan, candidate_output)?;
        Ok(())
    }
}

#[derive(Clone, Debug, Eq, Ord, PartialEq, PartialOrd)]
struct CanonicalXmlElement {
    name: String,
    attributes: Vec<(String, String, String)>,
    text: Option<String>,
    children: Vec<Self>,
}

#[derive(Clone, Debug, Default, Eq, PartialEq)]
struct WorkbookPreservationSnapshot {
    defined_names: Vec<CanonicalXmlElement>,
    dimensions: Vec<(
        String,
        Vec<CanonicalXmlElement>,
        Vec<(String, Vec<(String, String)>)>,
    )>,
    page_settings: Vec<(String, Vec<CanonicalXmlElement>)>,
    header_footers: Vec<(String, Option<CanonicalXmlElement>)>,
    conditional_formatting: Vec<(String, Vec<CanonicalXmlElement>)>,
    data_validations: Vec<(String, Option<CanonicalXmlElement>)>,
    relationships: Vec<(String, Vec<CanonicalXmlElement>)>,
    freeze_panes: Vec<(String, Option<CanonicalXmlElement>)>,
    auto_filters: Vec<(String, Option<CanonicalXmlElement>)>,
}

fn validate_preserved_ooxml_metadata(
    input_path: &Path,
    output_path: &Path,
    plan: &ExcelEditPlan,
) -> Result<()> {
    let checked = [
        PreservationRequirement::DefinedNames,
        PreservationRequirement::RowAndColumnDimensions,
        PreservationRequirement::PageSettings,
        PreservationRequirement::HeaderFooterText,
        PreservationRequirement::Relationships,
        PreservationRequirement::ConditionalFormatting,
        PreservationRequirement::DataValidations,
        PreservationRequirement::FreezePanes,
        PreservationRequirement::AutoFilter,
    ];
    if !checked
        .iter()
        .any(|requirement| plan.preserve.contains(requirement))
    {
        return Ok(());
    }

    let source = workbook_preservation_snapshot(input_path).map_err(|error| {
        anyhow::anyhow!(
            "read source OOXML metadata for openpyxl preservation validation: {error:#}"
        )
    })?;
    let output = workbook_preservation_snapshot(output_path).map_err(|error| {
        anyhow::anyhow!(
            "read output OOXML metadata for openpyxl preservation validation: {error:#}"
        )
    })?;

    if plan
        .preserve
        .contains(&PreservationRequirement::DefinedNames)
    {
        anyhow::ensure!(
            source.defined_names == output.defined_names,
            "openpyxl changed workbook defined names"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::RowAndColumnDimensions)
    {
        anyhow::ensure!(
            source.dimensions == output.dimensions,
            "openpyxl changed worksheet row or column dimensions"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::PageSettings)
    {
        anyhow::ensure!(
            source.page_settings == output.page_settings,
            "openpyxl changed worksheet page settings"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::HeaderFooterText)
    {
        anyhow::ensure!(
            source.header_footers == output.header_footers,
            "openpyxl changed worksheet header/footer settings"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::Relationships)
    {
        anyhow::ensure!(
            source.relationships == output.relationships,
            "openpyxl changed OOXML relationships"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::ConditionalFormatting)
    {
        anyhow::ensure!(
            source.conditional_formatting == output.conditional_formatting,
            "openpyxl changed conditional formatting rules"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::DataValidations)
    {
        anyhow::ensure!(
            source.data_validations == output.data_validations,
            "openpyxl changed data validations"
        );
    }
    if plan
        .preserve
        .contains(&PreservationRequirement::FreezePanes)
    {
        anyhow::ensure!(
            preserved_sheet_metadata_matches(
                &source.freeze_panes,
                &output.freeze_panes,
                &plan.operations,
                EditOperationKind::SetFreezePanes,
            ),
            "openpyxl changed freeze-pane metadata on a preserved worksheet"
        );
    }
    if plan.preserve.contains(&PreservationRequirement::AutoFilter) {
        anyhow::ensure!(
            preserved_sheet_metadata_matches(
                &source.auto_filters,
                &output.auto_filters,
                &plan.operations,
                EditOperationKind::SetAutoFilter,
            ),
            "openpyxl changed AutoFilter metadata on a preserved worksheet"
        );
    }
    Ok(())
}

fn preserved_sheet_metadata_matches(
    source: &[(String, Option<CanonicalXmlElement>)],
    output: &[(String, Option<CanonicalXmlElement>)],
    operations: &[EditOperation],
    operation_kind: EditOperationKind,
) -> bool {
    let (renamed, deleted) = sheet_name_changes(operations);
    source.iter().all(|(source_name, metadata)| {
        if deleted.contains(source_name) {
            return true;
        }
        let output_name = renamed.get(source_name).unwrap_or(source_name);
        let modified_by_plan =
            operations
                .iter()
                .any(|operation| match (operation_kind, operation) {
                    (
                        EditOperationKind::SetFreezePanes,
                        EditOperation::SetFreezePanes { sheet, .. },
                    )
                    | (
                        EditOperationKind::SetAutoFilter,
                        EditOperation::SetAutoFilter { sheet, .. },
                    ) => sheet == source_name || sheet == output_name,
                    _ => false,
                });
        if modified_by_plan {
            return true;
        }
        output
            .iter()
            .find(|(name, _)| name == output_name)
            .is_some_and(|(_, actual)| actual == metadata)
    })
}

fn validate_requested_sheet_layout(plan: &ExcelEditPlan, output_path: &Path) -> Result<()> {
    if !plan.operations.iter().any(|operation| {
        matches!(
            operation,
            EditOperation::SetFreezePanes { .. } | EditOperation::SetAutoFilter { .. }
        )
    }) {
        return Ok(());
    }
    let output = workbook_preservation_snapshot(output_path)?;
    for operation in &plan.operations {
        match operation {
            EditOperation::SetFreezePanes { sheet, cell } => {
                let actual = output
                    .freeze_panes
                    .iter()
                    .find(|(name, _)| name == sheet)
                    .and_then(|(_, pane)| pane.as_ref())
                    .and_then(|pane| {
                        pane.attributes.iter().find(|(namespace, name, _)| {
                            namespace.is_empty() && name == "topLeftCell"
                        })
                    })
                    .map(|(_, _, value)| normalize_address(value));
                let expected = cell.as_deref().map(normalize_address);
                anyhow::ensure!(
                    actual == expected,
                    "freeze panes readback mismatch for worksheet {sheet:?}: expected {expected:?}, got {actual:?}"
                );
            }
            EditOperation::SetAutoFilter { sheet, range } => {
                let actual = output
                    .auto_filters
                    .iter()
                    .find(|(name, _)| name == sheet)
                    .and_then(|(_, filter)| filter.as_ref())
                    .and_then(|filter| {
                        filter
                            .attributes
                            .iter()
                            .find(|(namespace, name, _)| namespace.is_empty() && name == "ref")
                    })
                    .map(|(_, _, value)| normalize_range(value));
                let expected = range.as_deref().map(normalize_range);
                anyhow::ensure!(
                    actual == expected,
                    "AutoFilter readback mismatch for worksheet {sheet:?}: expected {expected:?}, got {actual:?}"
                );
            }
            _ => {}
        }
    }
    Ok(())
}

fn workbook_preservation_snapshot(path: &Path) -> Result<WorkbookPreservationSnapshot> {
    const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const DOCUMENT_REL_NS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

    let mut archive = zip::ZipArchive::new(File::open(path)?)?;
    let workbook_xml = read_zip_part(&mut archive, "xl/workbook.xml")?;
    let workbook_document = roxmltree::Document::parse(&workbook_xml)?;
    let workbook_relationships_xml = read_zip_part(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let workbook_relationships_document = roxmltree::Document::parse(&workbook_relationships_xml)?;
    let relationship_targets = workbook_relationships_document
        .root_element()
        .children()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .filter_map(|node| {
            let id = node.attribute("Id")?;
            let target = node.attribute("Target")?;
            let target_mode = node.attribute("TargetMode").unwrap_or_default();
            Some((id.to_owned(), (target.to_owned(), target_mode.to_owned())))
        })
        .collect::<BTreeMap<_, _>>();

    let mut sheet_parts = Vec::new();
    for node in workbook_document.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let name = node
            .attribute("name")
            .context("workbook sheet is missing its name")?;
        let relationship_id = node
            .attributes()
            .find(|attribute| {
                attribute.name() == "id" && attribute.namespace() == Some(DOCUMENT_REL_NS)
            })
            .map(|attribute| attribute.value())
            .context("workbook sheet is missing its relationship id")?;
        let (target, target_mode) = relationship_targets
            .get(relationship_id)
            .with_context(|| format!("worksheet relationship {relationship_id:?} is missing"))?;
        anyhow::ensure!(
            target_mode.is_empty(),
            "worksheet relationship {relationship_id:?} is external"
        );
        sheet_parts.push((name.to_owned(), resolve_opc_target("xl", target)?));
    }

    let mut snapshot = WorkbookPreservationSnapshot::default();
    if let Some(defined_names) = workbook_document.descendants().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(MAIN_NS)
            && node.tag_name().name() == "definedNames"
    }) {
        snapshot.defined_names = defined_names
            .children()
            .filter(|node| node.is_element())
            .map(canonical_xml_element)
            .collect();
    }

    for (sheet_name, part_name) in sheet_parts {
        let xml = read_zip_part(&mut archive, &part_name)
            .with_context(|| format!("read worksheet part {part_name}"))?;
        let document = roxmltree::Document::parse(&xml)
            .with_context(|| format!("parse worksheet part {part_name}"))?;
        let root = document.root_element();

        let column_dimensions = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "cols")
            .map(|node| {
                node.children()
                    .filter(|child| child.is_element() && child.tag_name().name() == "col")
                    .map(canonical_xml_element)
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        let row_dimensions = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "sheetData")
            .map(|node| {
                node.children()
                    .filter(|child| child.is_element() && child.tag_name().name() == "row")
                    .map(|row| {
                        let row_id = row.attribute("r").unwrap_or_default().to_owned();
                        let attributes = row
                            .attributes()
                            .filter(|attribute| {
                                matches!(
                                    attribute.name(),
                                    "r" | "ht"
                                        | "hidden"
                                        | "customHeight"
                                        | "outlineLevel"
                                        | "collapsed"
                                        | "s"
                                        | "customFormat"
                                        | "thickTop"
                                        | "thickBot"
                                )
                            })
                            .map(|attribute| {
                                (attribute.name().to_owned(), attribute.value().to_owned())
                            })
                            .collect::<BTreeMap<_, _>>()
                            .into_iter()
                            .collect();
                        (row_id, attributes)
                    })
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        snapshot
            .dimensions
            .push((sheet_name.clone(), column_dimensions, row_dimensions));

        let mut page_settings = Vec::new();
        for child in root.children().filter(|node| node.is_element()) {
            if matches!(
                child.tag_name().name(),
                "pageMargins" | "pageSetup" | "printOptions" | "rowBreaks" | "colBreaks"
            ) {
                page_settings.push(canonical_xml_element(child));
            } else if child.tag_name().name() == "sheetPr" {
                if let Some(setup_properties) = child
                    .children()
                    .find(|node| node.is_element() && node.tag_name().name() == "pageSetUpPr")
                {
                    page_settings.push(canonical_xml_element(setup_properties));
                }
            }
        }
        snapshot
            .page_settings
            .push((sheet_name.clone(), page_settings));

        let pane = root
            .descendants()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "pane"
            })
            .map(canonical_xml_element);
        snapshot.freeze_panes.push((sheet_name.clone(), pane));
        let auto_filter = root
            .children()
            .find(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(MAIN_NS)
                    && node.tag_name().name() == "autoFilter"
            })
            .map(canonical_xml_element);
        snapshot
            .auto_filters
            .push((sheet_name.clone(), auto_filter));

        let header_footer = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "headerFooter")
            .map(canonical_xml_element);
        snapshot
            .header_footers
            .push((sheet_name.clone(), header_footer));

        let conditional_formatting = root
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "conditionalFormatting")
            .map(canonical_xml_element)
            .collect();
        snapshot
            .conditional_formatting
            .push((sheet_name.clone(), conditional_formatting));

        let data_validations = root
            .children()
            .find(|node| node.is_element() && node.tag_name().name() == "dataValidations")
            .map(canonical_xml_element);
        snapshot
            .data_validations
            .push((sheet_name, data_validations));
    }

    let relationship_parts = archive
        .file_names()
        .filter(|name| name.ends_with(".rels"))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    for part_name in relationship_parts {
        let xml = read_zip_part(&mut archive, &part_name)?;
        let document = roxmltree::Document::parse(&xml)
            .with_context(|| format!("parse relationship part {part_name}"))?;
        anyhow::ensure!(
            document.root_element().tag_name().namespace() == Some(PACKAGE_REL_NS),
            "relationship part {part_name} has an unexpected root namespace"
        );
        let mut entries = document
            .root_element()
            .children()
            .filter(|node| {
                node.is_element()
                    && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                    && node.tag_name().name() == "Relationship"
            })
            .map(canonical_xml_element)
            .collect::<Vec<_>>();
        entries.sort();
        snapshot.relationships.push((part_name, entries));
    }
    snapshot
        .relationships
        .sort_by(|left, right| left.0.cmp(&right.0));
    Ok(snapshot)
}

fn read_zip_part(archive: &mut zip::ZipArchive<File>, part_name: &str) -> Result<String> {
    let mut part = archive.by_name(part_name)?;
    let mut xml = String::new();
    part.read_to_string(&mut xml)
        .with_context(|| format!("read OOXML part {part_name}"))?;
    Ok(xml)
}

fn canonical_xml_element(node: roxmltree::Node<'_, '_>) -> CanonicalXmlElement {
    let mut attributes = node
        .attributes()
        .map(|attribute| {
            (
                attribute.namespace().unwrap_or_default().to_owned(),
                attribute.name().to_owned(),
                attribute.value().to_owned(),
            )
        })
        .collect::<Vec<_>>();
    attributes.sort();
    let children = node
        .children()
        .filter(|child| child.is_element())
        .map(canonical_xml_element)
        .collect();
    let text = node
        .children()
        .filter(|child| child.is_text())
        .filter_map(|child| child.text())
        .collect::<String>();
    CanonicalXmlElement {
        name: format!(
            "{}:{}",
            node.tag_name().namespace().unwrap_or_default(),
            node.tag_name().name()
        ),
        attributes,
        text: (!text.trim().is_empty()).then_some(text),
        children,
    }
}

fn resolve_opc_target(base: &str, target: &str) -> Result<String> {
    anyhow::ensure!(
        !target.contains(':') && !target.contains('?') && !target.contains('#'),
        "unsupported worksheet relationship target {target:?}"
    );
    let mut segments = if target.starts_with('/') {
        Vec::new()
    } else {
        base.split('/').map(str::to_owned).collect()
    };
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                anyhow::ensure!(
                    segments.pop().is_some(),
                    "relationship target escapes package root"
                );
            }
            value => segments.push(value.to_owned()),
        }
    }
    anyhow::ensure!(!segments.is_empty(), "empty worksheet relationship target");
    Ok(segments.join("/"))
}

fn validate_sheet_names(
    source: &[SheetDump],
    output: &[SheetDump],
    operations: &[EditOperation],
) -> Result<()> {
    let mut expected: Vec<String> = source.iter().map(|sheet| sheet.name.clone()).collect();
    for operation in operations {
        match operation {
            EditOperation::RenameSheet { old_name, new_name } => {
                let name = expected
                    .iter_mut()
                    .find(|name| *name == old_name)
                    .with_context(|| {
                        format!("source sheet {old_name:?} missing from original workbook")
                    })?;
                *name = new_name.clone();
            }
            EditOperation::CreateSheet { name } => expected.push(name.clone()),
            EditOperation::CreateSheetAt { name, index } => {
                anyhow::ensure!(
                    *index <= expected.len(),
                    "worksheet insertion index {index} exceeds expected sheet count {}",
                    expected.len()
                );
                expected.insert(*index, name.clone());
            }
            EditOperation::DeleteSheet { name } => expected.retain(|candidate| candidate != name),
            _ => {}
        }
    }
    let actual: Vec<String> = output.iter().map(|sheet| sheet.name.clone()).collect();
    anyhow::ensure!(
        actual == expected,
        "sheet identity/order mismatch: expected {expected:?}, got {actual:?}"
    );
    Ok(())
}

fn validate_requested_operations(
    source_sheets: &[SheetDump],
    sheets: &[SheetDump],
    operations: &[EditOperation],
) -> Result<()> {
    let by_name: BTreeMap<_, _> = sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect();
    for (operation_index, operation) in operations.iter().enumerate() {
        match operation {
            EditOperation::SetCellValue { sheet, cell, value } => {
                let sheet = by_name
                    .get(sheet.as_str())
                    .with_context(|| format!("output sheet {sheet:?} missing"))?;
                let actual = sheet.values.get(&normalize_address(cell));
                validate_value(value, actual).with_context(|| format!("verify {cell}"))?;
            }
            EditOperation::SetFormulaText {
                sheet,
                cell,
                formula,
            } => {
                let actual = by_name
                    .get(sheet.as_str())
                    .and_then(|sheet| sheet.formulas.get(&normalize_address(cell)))
                    .with_context(|| {
                        format!("formula {sheet}!{cell} missing after openpyxl edit")
                    })?;
                anyhow::ensure!(
                    actual.trim_start_matches('=').trim() == formula.trim_start_matches('=').trim(),
                    "formula mismatch at {sheet}!{cell}: {actual:?}"
                );
            }
            EditOperation::SetComment {
                sheet, cell, text, ..
            } => {
                let actual = by_name
                    .get(sheet.as_str())
                    .and_then(|sheet| sheet.comments.get(&normalize_address(cell)));
                match (text, actual) {
                    (None, None) => {}
                    (Some(expected), Some(actual)) => anyhow::ensure!(
                        expected == actual,
                        "comment text mismatch at {sheet}!{cell}"
                    ),
                    _ => bail!("comment presence mismatch at {sheet}!{cell}"),
                }
            }
            EditOperation::ReplaceHeaderText { .. } => {
                bail!("ReplaceHeaderText is intentionally routed to the OOXML Patch backend")
            }
            EditOperation::CopyRange {
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_comments,
                ..
            } => validate_range_copy(
                source_sheets,
                sheets,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                *copy_values,
                *copy_formulas,
                *copy_comments,
                &operations[..operation_index],
            )?,
            EditOperation::ImportRange {
                source_workbook,
                source_sheet,
                source_range,
                target_sheet,
                target_start,
                copy_values,
                copy_formulas,
                copy_comments,
                ..
            } => {
                let imported = read_workbook(source_workbook).with_context(|| {
                    format!(
                        "read ImportRange source {} for output validation",
                        source_workbook.display()
                    )
                })?;
                validate_range_copy(
                    &imported.sheets,
                    sheets,
                    source_sheet,
                    source_range,
                    target_sheet,
                    target_start,
                    *copy_values,
                    *copy_formulas,
                    *copy_comments,
                    &[],
                )?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn validate_calculation_properties(
    candidate_output: &Path,
    operations: &[EditOperation],
) -> Result<()> {
    let Some((full_calc_on_load, force_full_calc)) = operations.iter().find_map(|operation| {
        if let EditOperation::SetWorkbookCalculationProperties {
            full_calc_on_load,
            force_full_calc,
        } = operation
        {
            Some((*full_calc_on_load, *force_full_calc))
        } else {
            None
        }
    }) else {
        return Ok(());
    };
    let actual = ooxml_patch::read_workbook_calculation_properties(candidate_output)?;
    if let Some(expected) = full_calc_on_load {
        anyhow::ensure!(
            actual.full_calc_on_load == Some(expected),
            "fullCalcOnLoad readback mismatch: expected {expected}, got {:?}",
            actual.full_calc_on_load
        );
    }
    if let Some(expected) = force_full_calc {
        anyhow::ensure!(
            actual.force_full_calc == Some(expected),
            "forceFullCalc readback mismatch: expected {expected}, got {:?}",
            actual.force_full_calc
        );
    }
    Ok(())
}

fn needs_umya_readback(operation: &EditOperation) -> bool {
    matches!(
        operation,
        EditOperation::SetColumnWidth { .. }
            | EditOperation::SetRowHeight { .. }
            | EditOperation::SetCellFillColor { .. }
            | EditOperation::SetPageSettings { .. }
            | EditOperation::SetCellValue {
                value: crate::CellValue::DateTime { .. },
                ..
            }
            | EditOperation::CopyRange {
                copy_styles: true,
                ..
            }
            | EditOperation::ImportRange {
                style_mode: RangeStyleMode::Basic | RangeStyleMode::Full,
                ..
            }
    )
}

fn validate_umya_readback(plan: &ExcelEditPlan, candidate_output: &Path) -> Result<()> {
    let source = if plan.preserve.contains(&PreservationRequirement::Styles) {
        Some(crate::open_xlsx(&plan.input).context("open source workbook for style readback")?)
    } else {
        None
    };
    let output = crate::open_xlsx(candidate_output)
        .context("open openpyxl output with umya for independent Rust readback")?;

    if let Some(source) = &source {
        validate_unmodified_styles(plan, source, &output)?;
    }

    let expected_styles = expected_import_styles(plan)?;
    for ((sheet_name, address), expected) in expected_styles {
        let actual = style_at(&output, &sheet_name, &address)?;
        anyhow::ensure!(
            styles_semantically_equal(&actual, &expected),
            "Rust style readback mismatch at {sheet_name}!{address}: expected {expected:?}, got {actual:?}"
        );
    }
    validate_copy_style_results(&output, &plan.operations)?;

    let mut row_heights = BTreeMap::new();
    let mut column_widths = BTreeMap::new();
    for operation in &plan.operations {
        match operation {
            EditOperation::SetRowHeight {
                sheet,
                row_index,
                height,
            } => {
                row_heights.insert((sheet.as_str(), *row_index), *height);
            }
            EditOperation::SetColumnWidth {
                sheet,
                column_index,
                width,
            } => {
                column_widths.insert((sheet.as_str(), *column_index), *width);
            }
            _ => {}
        }
    }
    for ((sheet_name, row_index), expected) in row_heights {
        let row = output
            .sheet_by_name(sheet_name)
            .with_context(|| format!("row-height output sheet {sheet_name:?} missing"))?
            .row_dimension(row_index)
            .with_context(|| format!("row {row_index} missing after openpyxl edit"))?;
        anyhow::ensure!(
            (row.height() - expected).abs() <= 1e-10,
            "Rust row-height readback mismatch at {sheet_name}!{row_index}"
        );
    }
    for ((sheet_name, column_index), expected) in column_widths {
        let column = output
            .sheet_by_name(sheet_name)
            .with_context(|| format!("column-width output sheet {sheet_name:?} missing"))?
            .column_dimension_by_number(column_index)
            .with_context(|| format!("column {column_index} missing after openpyxl edit"))?;
        anyhow::ensure!(
            (column.width() - expected).abs() <= 1e-10,
            "Rust column-width readback mismatch at {sheet_name} column {column_index}"
        );
    }

    for operation in &plan.operations {
        match operation {
            EditOperation::SetCellFillColor { sheet, cell, argb } => {
                let expected = normalize_argb_color(argb)?;
                let actual = output
                    .sheet_by_name(sheet)
                    .with_context(|| format!("fill output sheet {sheet:?} missing"))?
                    .get_cell(cell.as_str())
                    .and_then(|cell| cell.style().fill())
                    .and_then(|fill| fill.pattern_fill())
                    .and_then(|fill| fill.foreground_color())
                    .map(|color| color.argb_str().to_ascii_uppercase());
                anyhow::ensure!(
                    actual.as_deref() == Some(expected.as_str()),
                    "Rust fill-color readback mismatch at {sheet}!{cell}"
                );
            }
            EditOperation::SetCellValue {
                sheet,
                cell,
                value: crate::CellValue::DateTime { number_format, .. },
            } => {
                let actual = output
                    .sheet_by_name(sheet)
                    .with_context(|| format!("date-style output sheet {sheet:?} missing"))?
                    .get_cell(cell.as_str())
                    .and_then(|cell| {
                        cell.style()
                            .number_format()
                            .map(|format| format.format_code().to_owned())
                    });
                anyhow::ensure!(
                    actual.as_deref() == Some(number_format.as_str()),
                    "Rust date number-format readback mismatch at {sheet}!{cell}"
                );
            }
            EditOperation::SetPageSettings { sheet, settings } => {
                let worksheet = output
                    .sheet_by_name(sheet)
                    .with_context(|| format!("page-settings output sheet {sheet:?} missing"))?;
                let page = worksheet.page_setup();
                if let Some(orientation) = settings.orientation {
                    let matches = match orientation {
                        crate::PageOrientation::Portrait => matches!(
                            page.orientation(),
                            umya_spreadsheet::OrientationValues::Portrait
                        ),
                        crate::PageOrientation::Landscape => matches!(
                            page.orientation(),
                            umya_spreadsheet::OrientationValues::Landscape
                        ),
                    };
                    anyhow::ensure!(
                        matches,
                        "Rust page orientation readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.fit_to_page {
                    anyhow::ensure!(
                        page.fit_to_page() == value,
                        "Rust fit-to-page readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.fit_to_width {
                    anyhow::ensure!(
                        page.fit_to_width() == u32::from(value),
                        "Rust fit-to-width readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.fit_to_height {
                    anyhow::ensure!(
                        page.fit_to_height() == u32::from(value),
                        "Rust fit-to-height readback mismatch for {sheet:?}"
                    );
                }
                if settings.clear_scale {
                    anyhow::ensure!(
                        page.scale() == 0,
                        "Rust scale-clear readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.paper_size {
                    anyhow::ensure!(
                        page.paper_size() == u32::from(value),
                        "Rust paper-size readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.horizontal_centered {
                    anyhow::ensure!(
                        worksheet.print_options().horizontal_centered() == value,
                        "Rust horizontal-centering readback mismatch for {sheet:?}"
                    );
                }
                if let Some(value) = settings.vertical_centered {
                    anyhow::ensure!(
                        worksheet.print_options().vertical_centered() == value,
                        "Rust vertical-centering readback mismatch for {sheet:?}"
                    );
                }
                if let Some(expected) = &settings.footer {
                    let actual = worksheet.header_footer().odd_footer().value();
                    anyhow::ensure!(
                        actual == expected || actual == format!("&C{expected}"),
                        "Rust footer readback mismatch for {sheet:?}: expected {expected:?}, got {actual:?}"
                    );
                }
                if let Some(expected) = &settings.print_area {
                    let actual = read_local_print_area(candidate_output, sheet)?
                        .map(|value| canonical_print_area(&value));
                    anyhow::ensure!(
                        actual.as_deref() == Some(canonical_print_area(expected).as_str()),
                        "Rust print-area readback mismatch for {sheet:?}: expected {:?}, got {actual:?}",
                        canonical_print_area(expected)
                    );
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn validate_unmodified_styles(
    plan: &ExcelEditPlan,
    source: &umya_spreadsheet::Workbook,
    output: &umya_spreadsheet::Workbook,
) -> Result<()> {
    let deleted_sheets: BTreeSet<&str> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::DeleteSheet { name } => Some(name.as_str()),
            _ => None,
        })
        .collect();
    for source_sheet in source.sheet_collection() {
        let source_name = source_sheet.name();
        if deleted_sheets.contains(source_name) {
            continue;
        }
        let output_name = final_sheet_name(plan, source_name);
        let output_sheet = output
            .sheet_by_name(&output_name)
            .with_context(|| format!("style-preservation output sheet {output_name:?} missing"))?;
        for source_cell in source_sheet.cells() {
            let address = source_cell.coordinate().to_string();
            if style_is_modified_by_plan(plan, source_name, &output_name, &address)? {
                continue;
            }
            let actual = output_sheet
                .cell(address.as_str())
                .map(|cell| cell.style().clone())
                .unwrap_or_default();
            // openpyxl normalizes non-anchor cells in merged ranges when it
            // loads/saves a workbook. Validate only border sides on the
            // merged range perimeter here; non-border style comes from the
            // anchor, which is checked independently above.
            if let Some(equal) =
                merged_member_effective_borders_equal(source_sheet, output_sheet, &address)
            {
                anyhow::ensure!(
                    equal,
                    "Rust unmodified merged-cell border readback mismatch at {output_name}!{address}"
                );
                continue;
            }
            anyhow::ensure!(
                styles_semantically_equal(&actual, source_cell.style()),
                "Rust unmodified-style readback mismatch at {output_name}!{address}"
            );
        }
    }
    Ok(())
}

fn merged_member_effective_borders_equal(
    source_sheet: &umya_spreadsheet::Worksheet,
    output_sheet: &umya_spreadsheet::Worksheet,
    address: &str,
) -> Option<bool> {
    let Ok((row, column)) = parse_cell(address) else {
        return None;
    };
    let merged = source_sheet.merge_cells().iter().find_map(|merged| {
        let Ok((start_row, start_column, end_row, end_column)) = parse_range(&merged.range())
        else {
            return None;
        };
        (row >= start_row
            && row <= end_row
            && column >= start_column
            && column <= end_column
            && (row != start_row || column != start_column))
            .then_some((start_row, start_column, end_row, end_column))
    })?;

    let anchor_address = format!("{}{}", column_letters(merged.1), merged.0);
    let bottom_right_address = format!("{}{}", column_letters(merged.3), merged.2);
    let source_anchor = source_sheet
        .cell(anchor_address.as_str())
        .map(|cell| cell.style().clone())
        .unwrap_or_default();
    let output_anchor = output_sheet
        .cell(anchor_address.as_str())
        .map(|cell| cell.style().clone())
        .unwrap_or_default();
    let source_bottom_right = source_sheet
        .cell(bottom_right_address.as_str())
        .map(|cell| cell.style().clone())
        .unwrap_or_default();
    let output_bottom_right = output_sheet
        .cell(bottom_right_address.as_str())
        .map(|cell| cell.style().clone())
        .unwrap_or_default();
    let source_anchor = source_anchor.borders().cloned().unwrap_or_default();
    let output_anchor = output_anchor.borders().cloned().unwrap_or_default();
    let source_bottom_right = source_bottom_right.borders().cloned().unwrap_or_default();
    let output_bottom_right = output_bottom_right.borders().cloned().unwrap_or_default();
    let source_bottom = merged_border_side(source_anchor.bottom(), source_bottom_right.bottom());
    let output_bottom = merged_border_side(output_anchor.bottom(), output_bottom_right.bottom());
    let source_right = merged_border_side(source_anchor.right(), source_bottom_right.right());
    let output_right = merged_border_side(output_anchor.right(), output_bottom_right.right());

    // openpyxl builds merged-cell border semantics from the top-left anchor,
    // then fills missing bottom/right edges from the bottom-right cell. An
    // existing anchor side takes precedence over the bottom-right side.
    Some(
        (row != merged.0 || source_anchor.top() == output_anchor.top())
            && (column != merged.1 || source_anchor.left() == output_anchor.left())
            && (row != merged.2 || source_bottom == output_bottom)
            && (column != merged.3 || source_right == output_right),
    )
}

fn merged_border_side<'a>(
    anchor: &'a umya_spreadsheet::Border,
    bottom_right: &'a umya_spreadsheet::Border,
) -> &'a umya_spreadsheet::Border {
    if anchor == &umya_spreadsheet::Border::default() {
        bottom_right
    } else {
        anchor
    }
}

fn final_sheet_name(plan: &ExcelEditPlan, source_name: &str) -> String {
    let mut name = source_name.to_owned();
    for operation in &plan.operations {
        if let EditOperation::RenameSheet { old_name, new_name } = operation {
            if old_name == &name {
                name.clone_from(new_name);
            }
        }
    }
    name
}

fn style_is_modified_by_plan(
    plan: &ExcelEditPlan,
    source_sheet_name: &str,
    output_sheet_name: &str,
    address: &str,
) -> Result<bool> {
    let sheet_matches = |sheet: &str| sheet == source_sheet_name || sheet == output_sheet_name;
    let address = normalize_address(address);
    for operation in &plan.operations {
        match operation {
            EditOperation::SetCellFillColor { sheet, cell, .. }
                if sheet_matches(sheet) && normalize_address(cell) == address =>
            {
                return Ok(true);
            }
            EditOperation::SetCellValue {
                sheet,
                cell,
                value: crate::CellValue::DateTime { .. },
            } if sheet_matches(sheet) && normalize_address(cell) == address => return Ok(true),
            EditOperation::CopyRange {
                target_sheet,
                source_range,
                target_start,
                copy_styles: true,
                ..
            } if sheet_matches(target_sheet)
                && target_range_contains(source_range, target_start, &address)? =>
            {
                return Ok(true);
            }
            EditOperation::ImportRange {
                target_sheet,
                source_range,
                target_start,
                style_mode: RangeStyleMode::Basic | RangeStyleMode::Full,
                ..
            } if sheet_matches(target_sheet)
                && target_range_contains(source_range, target_start, &address)? =>
            {
                return Ok(true);
            }
            _ => {}
        }
    }
    Ok(false)
}

fn target_range_contains(source_range: &str, target_start: &str, address: &str) -> Result<bool> {
    let (source_min_row, source_min_column, source_max_row, source_max_column) =
        parse_range(source_range)?;
    let (target_row, target_column) = parse_cell(target_start)?;
    let (row, column) = parse_cell(address)?;
    let target_max_row = target_row + source_max_row - source_min_row;
    let target_max_column = target_column + source_max_column - source_min_column;
    Ok((target_row..=target_max_row).contains(&row)
        && (target_column..=target_max_column).contains(&column))
}

fn read_local_print_area(path: &Path, sheet_name: &str) -> Result<Option<String>> {
    const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let mut archive = zip::ZipArchive::new(File::open(path)?)?;
    let mut workbook_xml = String::new();
    archive
        .by_name("xl/workbook.xml")?
        .read_to_string(&mut workbook_xml)?;
    let document = roxmltree::Document::parse(&workbook_xml)?;
    let sheets = document
        .descendants()
        .filter(|node| {
            node.tag_name().namespace() == Some(MAIN_NS) && node.tag_name().name() == "sheet"
        })
        .collect::<Vec<_>>();
    let sheet_index = sheets
        .iter()
        .position(|node| node.attribute("name") == Some(sheet_name))
        .context("find print-area worksheet in workbook.xml")?;
    let sheet_index = sheet_index.to_string();

    Ok(document
        .descendants()
        .find(|node| {
            node.tag_name().namespace() == Some(MAIN_NS)
                && node.tag_name().name() == "definedName"
                && node
                    .attribute("name")
                    .is_some_and(|name| name.eq_ignore_ascii_case("_xlnm.Print_Area"))
                && node.attribute("localSheetId") == Some(sheet_index.as_str())
        })
        .and_then(|node| node.text())
        .map(str::to_owned))
}

fn expected_import_styles(
    plan: &ExcelEditPlan,
) -> Result<BTreeMap<(String, String), umya_spreadsheet::Style>> {
    let mut styles = BTreeMap::new();
    for operation in &plan.operations {
        let EditOperation::ImportRange {
            source_workbook: import_path,
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            style_mode,
            ..
        } = operation
        else {
            continue;
        };
        if *style_mode == RangeStyleMode::None {
            continue;
        }
        let imported = crate::open_xlsx(import_path)
            .with_context(|| format!("open ImportRange styles from {}", import_path.display()))?;
        let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
        let (target_row, target_col) = parse_cell(target_start)?;
        for row in min_row..=max_row {
            for column in min_col..=max_col {
                let source_cell = format!("{}{}", column_letters(column), row);
                let target_cell = format!(
                    "{}{}",
                    column_letters(target_col + column - min_col),
                    target_row + row - min_row
                );
                let style = style_at(&imported, source_sheet, &source_cell)?;
                styles.insert(
                    (target_sheet.clone(), target_cell),
                    project_import_style(&style, *style_mode),
                );
            }
        }
    }
    Ok(styles)
}

fn validate_copy_style_results(
    output: &umya_spreadsheet::Workbook,
    operations: &[EditOperation],
) -> Result<()> {
    for operation in operations {
        let EditOperation::CopyRange {
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            copy_styles: true,
            ..
        } = operation
        else {
            continue;
        };
        let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
        let (target_row, target_col) = parse_cell(target_start)?;
        for row in min_row..=max_row {
            for column in min_col..=max_col {
                let source_cell = format!("{}{}", column_letters(column), row);
                let target_cell = format!(
                    "{}{}",
                    column_letters(target_col + column - min_col),
                    target_row + row - min_row
                );
                let source_style = style_at(output, source_sheet, &source_cell)?;
                let target_style = style_at(output, target_sheet, &target_cell)?;
                anyhow::ensure!(
                    source_style == target_style,
                    "Rust CopyRange style readback mismatch at {target_sheet}!{target_cell}"
                );
            }
        }
    }
    Ok(())
}

fn style_at(
    workbook: &umya_spreadsheet::Workbook,
    sheet: &str,
    address: &str,
) -> Result<umya_spreadsheet::Style> {
    let worksheet = workbook
        .sheet_by_name(sheet)
        .with_context(|| format!("style worksheet {sheet:?} missing"))?;
    Ok(worksheet
        .get_cell(address)
        .map(|cell| cell.style().clone())
        .unwrap_or_default())
}

fn project_import_style(
    source: &umya_spreadsheet::Style,
    mode: RangeStyleMode,
) -> umya_spreadsheet::Style {
    let mut target = umya_spreadsheet::Style::default();
    if let Some(value) = source.font() {
        target.set_font(value.clone());
    }
    if let Some(value) = source.fill() {
        target.set_fill(value.clone());
    }
    if let Some(value) = source.alignment() {
        target.set_alignment(value.clone());
    }
    if let Some(value) = source.number_format() {
        target.set_number_format(value.clone());
    }
    if mode == RangeStyleMode::Full {
        if let Some(value) = source.borders() {
            target.set_borders(value.clone());
        }
        if let Some(value) = source.protection() {
            target.set_protection(value.clone());
        }
    }
    target
}

fn styles_semantically_equal(
    actual: &umya_spreadsheet::Style,
    expected: &umya_spreadsheet::Style,
) -> bool {
    fn equal_defaulted<T>(actual: Option<&T>, expected: Option<&T>) -> bool
    where
        T: Clone + Default + PartialEq,
    {
        actual.cloned().unwrap_or_default() == expected.cloned().unwrap_or_default()
    }

    fn fill_is_visually_empty(fill: &umya_spreadsheet::Fill) -> bool {
        fill.gradient_fill().is_none()
            && fill.pattern_fill().is_none_or(|pattern| {
                pattern.pattern_type() == &umya_spreadsheet::PatternValues::None
                    && pattern.foreground_color().is_none()
                    && pattern.background_color().is_none()
            })
    }

    fn equal_fill(
        actual: Option<&umya_spreadsheet::Fill>,
        expected: Option<&umya_spreadsheet::Fill>,
    ) -> bool {
        let actual = actual.cloned().unwrap_or_default();
        let expected = expected.cloned().unwrap_or_default();

        if actual.gradient_fill() != expected.gradient_fill() {
            return false;
        }

        match (actual.pattern_fill(), expected.pattern_fill()) {
            (Some(actual), Some(expected)) => {
                actual.pattern_type() == expected.pattern_type()
                    && actual.foreground_color() == expected.foreground_color()
                    && actual.background_color() == expected.background_color()
            }
            (None, None) => true,
            (None, Some(_)) => fill_is_visually_empty(&expected),
            (Some(_), None) => fill_is_visually_empty(&actual),
        }
    }

    equal_defaulted(actual.font(), expected.font())
        && equal_fill(actual.fill(), expected.fill())
        && equal_defaulted(actual.borders(), expected.borders())
        && equal_defaulted(actual.alignment(), expected.alignment())
        && equal_defaulted(actual.protection(), expected.protection())
        && actual
            .number_format()
            .map(|format| format.format_code())
            .unwrap_or("General")
            == expected
                .number_format()
                .map(|format| format.format_code())
                .unwrap_or("General")
}

fn normalize_argb_color(value: &str) -> Result<String> {
    let mut value = value.trim().trim_start_matches('#').to_ascii_uppercase();
    if value.len() == 6 {
        value.insert_str(0, "FF");
    }
    anyhow::ensure!(
        value.len() == 8 && value.bytes().all(|byte| byte.is_ascii_hexdigit()),
        "fill color must be RGB or ARGB hexadecimal"
    );
    Ok(value)
}

fn canonical_print_area(value: &str) -> String {
    value
        .rsplit('!')
        .next()
        .unwrap_or(value)
        .replace('$', "")
        .replace('\'', "")
        .to_ascii_uppercase()
}

fn validate_range_copy(
    source_sheets: &[SheetDump],
    output_sheets: &[SheetDump],
    source_sheet_name: &str,
    source_range: &str,
    target_sheet_name: &str,
    target_start: &str,
    copy_values: bool,
    copy_formulas: bool,
    copy_comments: bool,
    preceding_operations: &[EditOperation],
) -> Result<()> {
    let source = source_sheets
        .iter()
        .find(|sheet| sheet.name == source_sheet_name)
        .with_context(|| format!("range source sheet {source_sheet_name:?} missing"))?;
    let target = output_sheets
        .iter()
        .find(|sheet| sheet.name == target_sheet_name)
        .with_context(|| format!("range target sheet {target_sheet_name:?} missing"))?;
    let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
    let (target_row, target_col) = parse_cell(target_start)?;
    for row in min_row..=max_row {
        for column in min_col..=max_col {
            let source_address = format!("{}{}", column_letters(column), row);
            let target_address = format!(
                "{}{}",
                column_letters(target_col + column - min_col),
                target_row + row - min_row
            );
            let mut planned_value = None;
            let mut planned_formula = None;
            let mut planned_comment = None;
            for operation in preceding_operations {
                match operation {
                    EditOperation::SetCellValue { sheet, cell, value }
                        if sheet == source_sheet_name
                            && normalize_address(cell) == source_address =>
                    {
                        planned_value = Some(value.clone());
                        planned_formula = Some(None);
                    }
                    EditOperation::SetFormulaText {
                        sheet,
                        cell,
                        formula,
                    } if sheet == source_sheet_name
                        && normalize_address(cell) == source_address =>
                    {
                        planned_value = Some(crate::CellValue::Blank);
                        planned_formula = Some(Some(formula.clone()));
                    }
                    EditOperation::SetComment {
                        sheet, cell, text, ..
                    } if sheet == source_sheet_name
                        && normalize_address(cell) == source_address =>
                    {
                        planned_comment = Some(text.clone());
                    }
                    _ => {}
                }
            }
            let source_formula = match &planned_formula {
                Some(formula) => formula.as_ref(),
                None => source.formulas.get(&source_address),
            };
            if copy_formulas && source_formula.is_some() {
                anyhow::ensure!(
                    target.formulas.contains_key(&target_address),
                    "copied formula is missing at {target_sheet_name}!{target_address}"
                );
            } else if copy_values {
                let actual = target.values.get(&target_address);
                if let Some(expected) = planned_value {
                    validate_planned_copy_value(&expected, actual).with_context(|| {
                        format!("copied cell value/type mismatch at {target_sheet_name}!{target_address}")
                    })?;
                } else {
                    let expected = source.values.get(&source_address);
                    anyhow::ensure!(
                        match (expected, actual) {
                            (None, None) => true,
                            (Some(expected), Some(actual)) =>
                                json_value_equal(expected, Some(actual)),
                            _ => false,
                        },
                        "copied cell value/type mismatch at {target_sheet_name}!{target_address}"
                    );
                }
            }
            if copy_comments {
                let expected = match &planned_comment {
                    Some(comment) => comment.as_deref(),
                    None => source.comments.get(&source_address).map(String::as_str),
                };
                let actual = target.comments.get(&target_address);
                anyhow::ensure!(
                    expected == actual.map(String::as_str),
                    "copied comment mismatch at {target_sheet_name}!{target_address}"
                );
            }
        }
    }
    Ok(())
}

fn validate_planned_copy_value(
    expected: &crate::CellValue,
    actual: Option<&serde_json::Value>,
) -> Result<()> {
    use crate::CellValue;
    match expected {
        CellValue::Blank => anyhow::ensure!(actual.is_none(), "expected blank copied cell"),
        CellValue::Text(value) => anyhow::ensure!(
            actual
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_str)
                == Some(value),
            "copied text value mismatch"
        ),
        CellValue::Boolean(value) => anyhow::ensure!(
            actual
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_bool)
                == Some(*value),
            "copied boolean value mismatch"
        ),
        CellValue::Error(value) => anyhow::ensure!(
            actual
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_str)
                == Some(value),
            "copied error value mismatch"
        ),
        CellValue::Decimal(value) | CellValue::DateTime { serial: value, .. } => {
            let expected = value.parse::<f64>().context("parse copied numeric value")?;
            let actual = actual
                .and_then(|value| value.get("value"))
                .and_then(serde_json::Value::as_f64)
                .context("copied numeric value is missing")?;
            let scale = expected.abs().max(actual.abs()).max(1.0);
            anyhow::ensure!(
                (expected - actual).abs() <= scale * 1e-15,
                "copied numeric value mismatch"
            );
        }
    }
    Ok(())
}

fn validate_value(expected: &crate::CellValue, actual: Option<&serde_json::Value>) -> Result<()> {
    use crate::CellValue;
    match expected {
        CellValue::Blank => anyhow::ensure!(actual.is_none(), "expected blank cell"),
        CellValue::Text(value) => anyhow::ensure!(
            actual
                .and_then(|v| v.get("value"))
                .and_then(serde_json::Value::as_str)
                == Some(value),
            "text value mismatch"
        ),
        CellValue::Boolean(value) => anyhow::ensure!(
            actual
                .and_then(|v| v.get("value"))
                .and_then(serde_json::Value::as_bool)
                == Some(*value),
            "boolean value mismatch"
        ),
        CellValue::Error(value) => anyhow::ensure!(
            actual
                .and_then(|v| v.get("value"))
                .and_then(serde_json::Value::as_str)
                == Some(value),
            "error value mismatch"
        ),
        CellValue::Decimal(value) | CellValue::DateTime { serial: value, .. } => {
            let expected = value
                .parse::<f64>()
                .context("parse expected numeric cell")?;
            let actual = actual
                .and_then(|v| v.get("value"))
                .and_then(serde_json::Value::as_f64)
                .with_context(|| "numeric cell missing or has another type")?;
            let scale = expected.abs().max(actual.abs()).max(1.0);
            anyhow::ensure!(
                (expected - actual).abs() <= scale * 1e-15,
                "numeric cell differs beyond 1e-15 relative tolerance"
            );
        }
    }
    Ok(())
}

fn validate_preserved_values(
    source: &[SheetDump],
    output: &[SheetDump],
    plan: &ExcelEditPlan,
) -> Result<()> {
    if !plan
        .preserve
        .contains(&PreservationRequirement::CellValuesAndTypes)
    {
        return Ok(());
    }
    let renamed: BTreeMap<String, String> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::RenameSheet { old_name, new_name } => {
                Some((old_name.clone(), new_name.clone()))
            }
            _ => None,
        })
        .collect();
    let deleted: BTreeSet<String> = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::DeleteSheet { name } => Some(name.clone()),
            _ => None,
        })
        .collect();
    for old_sheet in source {
        if deleted.contains(&old_sheet.name) {
            continue;
        }
        let output_name = renamed.get(&old_sheet.name).unwrap_or(&old_sheet.name);
        let Some(new_sheet) = output.iter().find(|sheet| &sheet.name == output_name) else {
            continue;
        };
        for (address, value) in &old_sheet.values {
            if cell_value_is_modified(&old_sheet.name, address, &plan.operations) {
                continue;
            }
            let actual = new_sheet.values.get(address);
            anyhow::ensure!(
                json_value_equal(value, actual),
                "unmodified cell value changed at {}!{}: expected {:?}, got {:?}",
                old_sheet.name,
                address,
                value,
                actual
            );
        }
    }
    Ok(())
}

fn validate_preserved_formulas(
    source: &[SheetDump],
    output: &[SheetDump],
    plan: &ExcelEditPlan,
) -> Result<()> {
    if !plan
        .preserve
        .contains(&PreservationRequirement::FormulaText)
    {
        return Ok(());
    }
    let (renamed, deleted) = sheet_name_changes(&plan.operations);
    for old_sheet in source {
        if deleted.contains(&old_sheet.name) {
            continue;
        }
        let output_name = renamed.get(&old_sheet.name).unwrap_or(&old_sheet.name);
        let Some(new_sheet) = output.iter().find(|sheet| &sheet.name == output_name) else {
            continue;
        };
        for (address, formula) in &old_sheet.formulas {
            if cell_formula_is_modified(&old_sheet.name, address, &plan.operations) {
                continue;
            }
            anyhow::ensure!(
                new_sheet.formulas.get(address) == Some(formula),
                "unmodified formula changed at {}!{}",
                old_sheet.name,
                address
            );
        }
    }
    Ok(())
}

fn validate_preserved_comments(
    source: &[SheetDump],
    output: &[SheetDump],
    plan: &ExcelEditPlan,
) -> Result<()> {
    if !plan
        .preserve
        .contains(&PreservationRequirement::CommentsAndPositions)
    {
        return Ok(());
    }
    let (renamed, deleted) = sheet_name_changes(&plan.operations);
    for old_sheet in source {
        if deleted.contains(&old_sheet.name) {
            continue;
        }
        let output_name = renamed.get(&old_sheet.name).unwrap_or(&old_sheet.name);
        let Some(new_sheet) = output.iter().find(|sheet| &sheet.name == output_name) else {
            continue;
        };
        for (address, text) in &old_sheet.comments {
            if cell_comment_is_modified(&old_sheet.name, address, &plan.operations) {
                continue;
            }
            anyhow::ensure!(
                new_sheet.comments.get(address) == Some(text),
                "unmodified comment moved or changed at {}!{}",
                old_sheet.name,
                address
            );
        }
    }
    Ok(())
}

fn validate_preserved_merges(
    source: &[SheetDump],
    output: &[SheetDump],
    plan: &ExcelEditPlan,
) -> Result<()> {
    if !plan
        .preserve
        .contains(&PreservationRequirement::MergedRanges)
    {
        return Ok(());
    }
    let (renamed, deleted) = sheet_name_changes(&plan.operations);
    let mut expected: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    for old_sheet in source {
        if deleted.contains(&old_sheet.name) {
            continue;
        }
        let output_name = renamed.get(&old_sheet.name).unwrap_or(&old_sheet.name);
        expected.insert(
            output_name.clone(),
            old_sheet
                .merges
                .iter()
                .map(|range| normalize_range(range))
                .collect(),
        );
    }
    for operation in &plan.operations {
        match operation {
            EditOperation::CreateSheet { name } => {
                expected.entry(name.clone()).or_default();
            }
            EditOperation::CreateSheetAt { name, .. } => {
                expected.entry(name.clone()).or_default();
            }
            _ => {}
        }
    }
    for operation in &plan.operations {
        let EditOperation::ImportRange {
            source_workbook,
            source_sheet,
            source_range,
            target_sheet,
            target_start,
            copy_merges: true,
            ..
        } = operation
        else {
            continue;
        };
        let imported = read_workbook(source_workbook).with_context(|| {
            format!(
                "read ImportRange source {} for merge validation",
                source_workbook.display()
            )
        })?;
        let imported_sheet = imported
            .sheets
            .iter()
            .find(|sheet| sheet.name == *source_sheet)
            .with_context(|| format!("ImportRange source sheet {source_sheet:?} missing"))?;
        let (min_row, min_col, max_row, max_col) = parse_range(source_range)?;
        let (target_row, target_col) = parse_cell(target_start)?;
        let target_merges = expected.entry(target_sheet.clone()).or_default();
        for merge in &imported_sheet.merges {
            let (first_row, first_col, last_row, last_col) = parse_range(merge)?;
            let intersects = first_row <= max_row
                && last_row >= min_row
                && first_col <= max_col
                && last_col >= min_col;
            if !intersects {
                continue;
            }
            anyhow::ensure!(
                first_row >= min_row
                    && last_row <= max_row
                    && first_col >= min_col
                    && last_col <= max_col,
                "ImportRange cuts through merged range {merge:?}"
            );
            let row_offset = target_row as i64 - min_row as i64;
            let column_offset = target_col as i64 - min_col as i64;
            let translated = format!(
                "{}{}:{}{}",
                column_letters((first_col as i64 + column_offset) as u32),
                (first_row as i64 + row_offset),
                column_letters((last_col as i64 + column_offset) as u32),
                (last_row as i64 + row_offset),
            );
            target_merges.insert(normalize_range(&translated));
        }
    }
    for sheet in output {
        let actual: BTreeSet<_> = sheet
            .merges
            .iter()
            .map(|range| normalize_range(range))
            .collect();
        let expected_ranges = expected.get(&sheet.name).cloned().unwrap_or_default();
        anyhow::ensure!(
            expected_ranges == actual,
            "merged ranges changed on sheet {:?}: expected {:?}, got {:?}",
            sheet.name,
            expected_ranges,
            actual
        );
    }
    Ok(())
}

fn sheet_name_changes(
    operations: &[EditOperation],
) -> (BTreeMap<String, String>, BTreeSet<String>) {
    let renamed = operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::RenameSheet { old_name, new_name } => {
                Some((old_name.clone(), new_name.clone()))
            }
            _ => None,
        })
        .collect();
    let deleted = operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::DeleteSheet { name } => Some(name.clone()),
            _ => None,
        })
        .collect();
    (renamed, deleted)
}

fn normalize_range(value: &str) -> String {
    value.replace('$', "").to_ascii_uppercase()
}

fn column_letters(mut column: u32) -> String {
    let mut result = String::new();
    while column > 0 {
        let remainder = ((column - 1) % 26) as u8;
        result.push((b'A' + remainder) as char);
        column = (column - 1) / 26;
    }
    result.chars().rev().collect()
}

fn cell_value_is_modified(sheet: &str, address: &str, operations: &[EditOperation]) -> bool {
    operations.iter().any(|operation| match operation {
        EditOperation::SetCellValue {
            sheet: target_sheet,
            cell,
            ..
        }
        | EditOperation::SetFormulaText {
            sheet: target_sheet,
            cell,
            ..
        } => target_sheet == sheet && normalize_address(cell) == address,
        EditOperation::CopyRange {
            target_sheet,
            target_start,
            source_range,
            copy_values,
            copy_formulas,
            ..
        }
        | EditOperation::ImportRange {
            target_sheet,
            target_start,
            source_range,
            copy_values,
            copy_formulas,
            ..
        } => {
            (*copy_values || *copy_formulas)
                && target_sheet == sheet
                && address_in_copied_range(address, target_start, source_range)
        }
        _ => false,
    })
}

fn cell_formula_is_modified(sheet: &str, address: &str, operations: &[EditOperation]) -> bool {
    operations.iter().any(|operation| match operation {
        EditOperation::SetCellValue {
            sheet: target_sheet,
            cell,
            ..
        }
        | EditOperation::SetFormulaText {
            sheet: target_sheet,
            cell,
            ..
        } => target_sheet == sheet && normalize_address(cell) == address,
        EditOperation::CopyRange {
            target_sheet,
            target_start,
            source_range,
            copy_values,
            copy_formulas,
            ..
        }
        | EditOperation::ImportRange {
            target_sheet,
            target_start,
            source_range,
            copy_values,
            copy_formulas,
            ..
        } => {
            (*copy_values || *copy_formulas)
                && target_sheet == sheet
                && address_in_copied_range(address, target_start, source_range)
        }
        _ => false,
    })
}

fn cell_comment_is_modified(sheet: &str, address: &str, operations: &[EditOperation]) -> bool {
    operations.iter().any(|operation| match operation {
        EditOperation::SetComment {
            sheet: target_sheet,
            cell,
            ..
        } => target_sheet == sheet && normalize_address(cell) == address,
        EditOperation::CopyRange {
            target_sheet,
            target_start,
            source_range,
            copy_comments,
            ..
        }
        | EditOperation::ImportRange {
            target_sheet,
            target_start,
            source_range,
            copy_comments,
            ..
        } => {
            *copy_comments
                && target_sheet == sheet
                && address_in_copied_range(address, target_start, source_range)
        }
        _ => false,
    })
}

fn address_in_copied_range(address: &str, target_start: &str, source_range: &str) -> bool {
    let Ok((row, column)) = parse_cell(address) else {
        return false;
    };
    let Ok((start_row, start_column)) = parse_cell(target_start) else {
        return false;
    };
    let Ok((source_start_row, source_start_column, source_end_row, source_end_column)) =
        parse_range(source_range)
    else {
        return false;
    };
    row >= start_row
        && row <= start_row + source_end_row - source_start_row
        && column >= start_column
        && column <= start_column + source_end_column - source_start_column
}

fn json_value_equal(expected: &serde_json::Value, actual: Option<&serde_json::Value>) -> bool {
    let Some(actual) = actual else {
        return false;
    };
    let expected_type = expected.get("type").and_then(serde_json::Value::as_str);
    let actual_type = actual.get("type").and_then(serde_json::Value::as_str);
    if expected_type != actual_type {
        return false;
    }
    let (Some(left), Some(right)) = (expected.get("value"), actual.get("value")) else {
        return false;
    };
    match (left.as_f64(), right.as_f64()) {
        (Some(left), Some(right)) => {
            (left - right).abs() <= left.abs().max(right.abs()).max(1.0) * 1e-15
        }
        _ => left == right,
    }
}

fn normalize_address(address: &str) -> String {
    address.replace('$', "").to_ascii_uppercase()
}

fn parse_cell(value: &str) -> Result<(u32, u32)> {
    let mut column = 0u32;
    let mut row = 0u32;
    let mut digits = false;
    for character in value.trim().trim_start_matches('$').chars() {
        if character == '$' {
            continue;
        }
        if character.is_ascii_alphabetic() && !digits {
            column = column
                .checked_mul(26)
                .and_then(|v| v.checked_add(character.to_ascii_uppercase() as u32 - 'A' as u32 + 1))
                .context("column overflow")?;
        } else if character.is_ascii_digit() {
            digits = true;
            row = row
                .checked_mul(10)
                .and_then(|v| v.checked_add(character as u32 - '0' as u32))
                .context("row overflow")?;
        } else {
            bail!("invalid A1 coordinate {value:?}");
        }
    }
    anyhow::ensure!(row > 0 && column > 0, "invalid A1 coordinate {value:?}");
    Ok((row, column))
}

fn parse_range(value: &str) -> Result<(u32, u32, u32, u32)> {
    let (start, end) = value.split_once(':').unwrap_or((value, value));
    let (start_row, start_column) = parse_cell(start)?;
    let (end_row, end_column) = parse_cell(end)?;
    Ok((start_row, start_column, end_row, end_column))
}

#[derive(Deserialize)]
struct WorkerResponse {
    schema_version: u32,
    request_id: Option<String>,
    status: String,
    operation_count: Option<usize>,
    verified_operation_count: Option<usize>,
    input_sha256_before: Option<String>,
    input_sha256_after: Option<String>,
    output_sha256: Option<String>,
    error: Option<WorkerResponseError>,
}

#[derive(Deserialize)]
struct WorkerResponseError {
    kind: String,
    message: String,
}

struct ScratchFile(PathBuf);

impl ScratchFile {
    fn create(parent: &Path, purpose: &str) -> Result<Self> {
        for _ in 0..128 {
            let sequence = SCRATCH_SEQUENCE.fetch_add(1, Ordering::Relaxed);
            let path = parent.join(format!(
                ".openpyxl-{}-{sequence}-{purpose}.tmp",
                std::process::id()
            ));
            match OpenOptions::new().write(true).create_new(true).open(&path) {
                Ok(_) => return Ok(Self(path)),
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => return Err(error).with_context(|| format!("reserve {purpose}")),
            }
        }
        bail!("could not reserve a unique {purpose} path")
    }
    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for ScratchFile {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn read_diagnostic(path: &Path) -> String {
    fs::read_to_string(path)
        .unwrap_or_default()
        .chars()
        .take(4096)
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn merged_style_readback_treats_only_non_anchor_cells_as_normalized() {
        let mut source = umya_spreadsheet::new_file();
        let source_sheet = source.sheet_by_name_mut("Sheet1").unwrap();
        source_sheet.add_merge_cells("A1:I1");
        source_sheet
            .cell_mut("A1")
            .style_mut()
            .borders_mut()
            .top_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        source_sheet
            .cell_mut("A1")
            .style_mut()
            .borders_mut()
            .bottom_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        source_sheet
            .cell_mut("I1")
            .style_mut()
            .borders_mut()
            .right_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        source_sheet
            .cell_mut("I1")
            .style_mut()
            .borders_mut()
            .bottom_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Double);
        source_sheet
            .cell_mut("G1")
            .style_mut()
            .borders_mut()
            .bottom_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Double);

        let mut output = umya_spreadsheet::new_file();
        let output_sheet = output.sheet_by_name_mut("Sheet1").unwrap();
        output_sheet.add_merge_cells("A1:I1");
        output_sheet
            .cell_mut("A1")
            .style_mut()
            .borders_mut()
            .top_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        output_sheet
            .cell_mut("A1")
            .style_mut()
            .borders_mut()
            .bottom_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);
        output_sheet
            .cell_mut("I1")
            .style_mut()
            .borders_mut()
            .right_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Thin);

        // G1 inherits top/bottom from the anchor. The source bottom-right
        // cell's conflicting raw bottom border is ignored because the anchor
        // already defines that side.
        assert_eq!(
            merged_member_effective_borders_equal(source_sheet, output_sheet, "G1"),
            Some(true)
        );
        // The anchor remains under the full style comparison.
        assert_eq!(
            merged_member_effective_borders_equal(source_sheet, output_sheet, "A1"),
            None
        );
        // The right edge must still match: the anchor leaves it empty, so
        // openpyxl fills it from the bottom-right cell.
        output_sheet
            .cell_mut("I1")
            .style_mut()
            .borders_mut()
            .right_mut()
            .set_style(umya_spreadsheet::BorderStyleValues::Double);
        assert_eq!(
            merged_member_effective_borders_equal(source_sheet, output_sheet, "I1"),
            Some(false)
        );
        assert_eq!(
            merged_member_effective_borders_equal(source_sheet, output_sheet, "J1"),
            None
        );
    }

    #[test]
    fn packaged_runtime_uses_an_absolute_internal_path_and_never_path_lookup() {
        let config =
            OpenpyxlWorkerConfig::for_product_binary("/usr/bin/pytools", "pytools").unwrap();
        assert_eq!(
            config.python_executable,
            PathBuf::from("/usr/libexec/pytools/openpyxl/runtime/bin/python3")
        );
        assert_eq!(
            config.worker_script,
            PathBuf::from("/usr/libexec/pytools/openpyxl/worker/openpyxl_worker.py")
        );
        assert!(
            OpenpyxlWorkerConfig::for_product_binary("/usr/bin/pytools", "../python").is_none()
        );
    }

    #[test]
    fn capability_is_narrowed_to_implemented_operations_and_semantic_parts() {
        let capability = BackendCapability {
            engine: crate::EngineId::OpenpyxlWorker,
            status: crate::CapabilityStatus::Unverified,
            format: WorkbookFormat::Xlsx,
            operations: [
                EditOperationKind::SetCellValue,
                EditOperationKind::DeleteRows,
            ]
            .into_iter()
            .collect(),
            supported_features: BTreeSet::new(),
            preserved: [
                PreservationRequirement::CellValuesAndTypes,
                PreservationRequirement::Styles,
                PreservationRequirement::PageSettings,
                PreservationRequirement::RowAndColumnDimensions,
                PreservationRequirement::Relationships,
                PreservationRequirement::DefinedNames,
                PreservationRequirement::HeaderFooterText,
                PreservationRequirement::ConditionalFormatting,
                PreservationRequirement::DataValidations,
                PreservationRequirement::UnmodifiedPackageParts,
            ]
            .into_iter()
            .collect(),
            platform: "test".into(),
            evidence: None,
        };
        let backend = OpenpyxlXlsxBackend::build(capability, None).unwrap();
        assert!(backend
            .capability()
            .operations
            .contains(&EditOperationKind::SetCellValue));
        assert!(!backend
            .capability()
            .operations
            .contains(&EditOperationKind::DeleteRows));
        assert!(!backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::UnmodifiedPackageParts));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::Styles));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::PageSettings));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::RowAndColumnDimensions));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::Relationships));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::DefinedNames));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::HeaderFooterText));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::ConditionalFormatting));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::DataValidations));
        assert!(backend
            .capability()
            .preserved
            .contains(&PreservationRequirement::CellValuesAndTypes));
        assert_eq!(
            backend.capability().status,
            crate::CapabilityStatus::Unsupported
        );
    }

    #[test]
    fn style_only_copy_does_not_exempt_values_formulas_or_comments_from_preservation() {
        let operation = EditOperation::CopyRange {
            source_sheet: "Data".into(),
            source_range: "A1:B2".into(),
            target_sheet: "Data".into(),
            target_start: "D4".into(),
            copy_values: false,
            copy_formulas: false,
            copy_styles: true,
            copy_comments: false,
        };
        let operations = [operation];
        assert!(!cell_value_is_modified("Data", "D4", &operations));
        assert!(!cell_formula_is_modified("Data", "D4", &operations));
        assert!(!cell_comment_is_modified("Data", "D4", &operations));
    }

    #[test]
    fn changing_a_value_does_not_exempt_existing_comment_from_preservation() {
        let operations = [EditOperation::SetCellValue {
            sheet: "Data".into(),
            cell: "A1".into(),
            value: crate::CellValue::Text("replacement".into()),
        }];
        assert!(cell_value_is_modified("Data", "A1", &operations));
        assert!(cell_formula_is_modified("Data", "A1", &operations));
        assert!(!cell_comment_is_modified("Data", "A1", &operations));
    }
}
