//! S5 workbook-structure snapshots (`SourceRepository` schema checks port).
//!
//! Snapshot payloads retain the Python contract: rule-path labels are HMACed
//! with a per-data-directory secret, logical changes require review, and
//! position-only changes may be accepted automatically after an intact run.

use crate::{
    s5_config::MetricSpec,
    s5_layout::NamedRegion,
    s5_scan::{ScanBook, ScanResult, ScanTable},
};
use serde_json::{json, Map, Value};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, OpenOptions},
    io::{self, Write},
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

#[cfg(unix)]
use std::{fs::File, io::Read};

const SECRET_NAME: &str = ".schema-key";

#[derive(Clone, Debug)]
pub struct S5StructureObservation {
    pub role: String,
    pub source_id: String,
    pub business: String,
    pub sheet: String,
    pub table_id: String,
    pub table: ScanTable,
    pub expected_row_paths: Vec<Vec<String>>,
    pub expected_column_paths: Vec<Vec<String>>,
    pub row_merge_shapes: Vec<String>,
    pub column_merge_shapes: Vec<String>,
    pub numeric_regions: Vec<NamedRegion>,
}

pub fn collect_observations(
    book: &ScanBook,
    role: &str,
    source_id: &str,
    business: &str,
    sheet: &str,
    metrics: &[MetricSpec],
    scan: &ScanResult,
    numeric_regions: &[NamedRegion],
) -> Vec<S5StructureObservation> {
    let expected_row_paths: Vec<Vec<String>> = metrics
        .iter()
        .filter(|metric| metric.business == business && metric.sheet == sheet)
        .map(|metric| metric.row_path.clone())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect();
    let expected_column_paths: Vec<Vec<String>> = metrics
        .iter()
        .filter(|metric| metric.business == business && metric.sheet == sheet)
        .map(|metric| metric.column_path.clone())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect();

    scan.tables
        .iter()
        .map(|table| {
            let search_region = table.column_region.as_ref().unwrap_or(&table.region);
            S5StructureObservation {
                role: role.to_owned(),
                source_id: source_id.to_owned(),
                business: business.to_owned(),
                sheet: sheet.to_owned(),
                table_id: format!("{}#1", table.region.name),
                table: table.clone(),
                expected_row_paths: expected_row_paths.clone(),
                expected_column_paths: expected_column_paths.clone(),
                row_merge_shapes: book.merged_shapes_in_region(sheet, &table.region),
                column_merge_shapes: book.merged_shapes_in_region(sheet, search_region),
                numeric_regions: numeric_regions.to_vec(),
            }
        })
        .collect()
}

struct PendingRecord {
    path: PathBuf,
    payload: Value,
    event: Value,
}

pub struct PendingSnapshots {
    directory: PathBuf,
    secret: Vec<u8>,
    records: Vec<PendingRecord>,
}

impl PendingSnapshots {
    /// Compare observations before output copying, matching the engine's
    /// pre-output snapshot comparison. No raw row or column labels are stored.
    pub fn compare(
        observations: &[S5StructureObservation],
        directory: &Path,
    ) -> Result<Self, String> {
        let mut pending = Self {
            directory: directory.to_path_buf(),
            secret: Vec::new(),
            records: Vec::new(),
        };
        if observations.is_empty() {
            return Ok(pending);
        }
        pending.secret = read_or_create_secret(directory)?;

        for observation in observations {
            pending
                .records
                .push(compare_one(observation, &pending.secret, directory)?);
        }
        pending.records.sort_by(|left, right| {
            left.event["snapshot_id"]
                .as_str()
                .cmp(&right.event["snapshot_id"].as_str())
        });
        Ok(pending)
    }

    /// Persist only after source/configuration integrity has been confirmed.
    /// A logical structure change keeps the accepted baseline for review.
    pub fn finish(mut self, inputs_unchanged: bool) -> Vec<Value> {
        if self.records.is_empty() {
            return Vec::new();
        }
        if !inputs_unchanged {
            for record in &mut self.records {
                if matches!(
                    record.event["status"].as_str(),
                    Some("BASELINE_CREATED" | "UNCHANGED" | "POSITION_CHANGED")
                ) {
                    record.event["status"] = json!("SNAPSHOT_NOT_SAVED");
                    record.event["detail"] =
                        json!("输入或配置在生成审核副本时发生变化，未更新结构快照");
                }
            }
            return self
                .records
                .into_iter()
                .map(|record| record.event)
                .collect();
        }

        let directory_result = fs::create_dir_all(&self.directory)
            .and_then(|()| write_secret_if_missing(&self.directory, &self.secret));
        if let Err(error) = directory_result {
            let error_name = python_error_name(&error);
            for record in &mut self.records {
                if record.event["status"].as_str() != Some("STRUCTURE_CHANGED") {
                    record.event["status"] = json!("SNAPSHOT_WRITE_FAILED");
                    record.event["detail"] = json!(format!("结构快照未写入：{error_name}"));
                }
            }
            return self
                .records
                .into_iter()
                .map(|record| record.event)
                .collect();
        }

        for record in &mut self.records {
            if record.event["status"].as_str() == Some("STRUCTURE_CHANGED") {
                continue;
            }
            if let Err(error) = write_snapshot_atomic(&record.path, &record.payload) {
                let error_name = python_error_name(&error);
                record.event["status"] = json!("SNAPSHOT_WRITE_FAILED");
                record.event["detail"] = json!(format!("结构快照未写入：{error_name}"));
            }
        }
        self.records
            .into_iter()
            .map(|record| record.event)
            .collect()
    }
}

fn compare_one(
    observation: &S5StructureObservation,
    secret: &[u8],
    directory: &Path,
) -> Result<PendingRecord, String> {
    let mut identity = BTreeMap::new();
    identity.insert("role", observation.role.as_str());
    identity.insert("source_id", observation.source_id.as_str());
    identity.insert("business", observation.business.as_str());
    identity.insert("sheet", observation.sheet.as_str());
    identity.insert("table_id", observation.table_id.as_str());
    let identity_bytes = serde_json::to_vec(&identity).map_err(|error| error.to_string())?;
    let snapshot_id = sha256_hex(&identity_bytes);

    let row_header_width = observation
        .expected_row_paths
        .iter()
        .map(Vec::len)
        .max()
        .unwrap_or(0);
    let mut row_presence = Map::new();
    let mut configured_rows = Map::new();
    for path in &observation.expected_row_paths {
        let token = path_token(secret, path)?;
        let hits = observation
            .table
            .row_hits
            .get(path)
            .cloned()
            .unwrap_or_default();
        let mut rows = BTreeSet::new();
        let mut unique_hits = BTreeSet::new();
        for (row, first_column, last_column) in hits {
            rows.insert(row);
            unique_hits.insert((row, first_column, last_column));
        }
        row_presence.insert(token.clone(), json!(unique_hits.len().min(2)));
        configured_rows.insert(token, json!(rows.into_iter().collect::<Vec<_>>()));
    }

    let mut column_presence = Map::new();
    let mut configured_columns = Map::new();
    for path in &observation.expected_column_paths {
        let token = path_token(secret, path)?;
        let columns: BTreeSet<usize> = observation
            .table
            .column_hits
            .get(path)
            .cloned()
            .unwrap_or_default()
            .into_iter()
            .collect();
        column_presence.insert(token.clone(), json!(columns.len().min(2)));
        configured_columns.insert(token, json!(columns.into_iter().collect::<Vec<_>>()));
    }

    let search_region = observation
        .table
        .column_region
        .as_ref()
        .unwrap_or(&observation.table.region);
    let mut header_merges = BTreeSet::new();
    header_merges.extend(observation.row_merge_shapes.iter().cloned());
    header_merges.extend(observation.column_merge_shapes.iter().cloned());
    let logical = json!({
        "row_header_width": row_header_width,
        "configured_row_presence": row_presence,
        "configured_column_presence": column_presence,
        "row_region_width": region_width(&observation.table.region),
        "column_region_width": region_width(search_region),
        "row_header_merge_shapes": observation.row_merge_shapes,
        "column_header_merge_shapes": observation.column_merge_shapes,
        "numeric_region_shapes": numeric_region_shapes(&observation.numeric_regions),
    });
    let row_header_anchor =
        first_row_header_anchor(&observation.expected_row_paths, &observation.table);
    let positions = json!({
        "row_header_anchor": row_header_anchor,
        "row_region": observation.table.region.reference,
        "column_region": observation.table.column_region.as_ref().map(|region| region.reference.as_str()).unwrap_or(""),
        "configured_rows": configured_rows,
        "configured_columns": configured_columns,
        "header_merges": header_merges,
        "numeric_regions": observation.numeric_regions.iter().map(|region| region.reference.as_str()).collect::<Vec<_>>(),
    });
    let configured_row_path_count = observation.expected_row_paths.len();
    let configured_column_path_count = observation.expected_column_paths.len();
    let matched_row_path_count = row_presence_count(&logical["configured_row_presence"]);
    let matched_column_path_count = row_presence_count(&logical["configured_column_presence"]);
    let row_region = observation.table.region.reference.as_str();
    let column_region = observation
        .table
        .column_region
        .as_ref()
        .map(|region| region.reference.as_str())
        .unwrap_or("");
    let summary = json!({
        "row_header_width": row_header_width,
        "configured_row_path_count": configured_row_path_count,
        "configured_column_path_count": configured_column_path_count,
        "matched_row_path_count": matched_row_path_count,
        "matched_column_path_count": matched_column_path_count,
        "row_region": row_region,
        "column_region": column_region,
    });
    let identity = json!({
        "role": observation.role,
        "source_id": observation.source_id,
        "business": observation.business,
        "sheet": observation.sheet,
        "table_id": observation.table_id,
    });
    let payload = json!({
        "format_version": 1,
        "snapshot_id": snapshot_id,
        "identity": identity,
        "logical": logical,
        "positions": positions,
        "summary": summary,
    });
    let path = directory.join(format!("{snapshot_id}.json"));
    let mut event = json!({
        "snapshot_id": snapshot_id,
        "role": observation.role,
        "source_id": observation.source_id,
        "business": observation.business,
        "sheet": observation.sheet,
        "table_id": observation.table_id,
        "row_header_width": row_header_width,
        "configured_row_path_count": configured_row_path_count,
        "configured_column_path_count": configured_column_path_count,
        "matched_row_path_count": matched_row_path_count,
        "matched_column_path_count": matched_column_path_count,
        "row_region": row_region,
        "column_region": column_region,
    });
    match fs::read_to_string(&path) {
        Ok(text) => match serde_json::from_str::<Value>(&text) {
            Ok(previous) if !previous.is_object() => {
                event["status"] = json!("SNAPSHOT_INVALID");
                event["detail"] = json!("历史结构快照无法读取：AttributeError");
            }
            Ok(previous) if previous.get("logical") != payload.get("logical") => {
                event["status"] = json!("STRUCTURE_CHANGED");
                event["detail"] = json!("配置路径或表头逻辑结构发生变化");
            }
            Ok(previous) if previous.get("positions") != payload.get("positions") => {
                event["status"] = json!("POSITION_CHANGED");
                event["detail"] = json!("表结构逻辑未变，单元格位置已自动重定位");
            }
            Ok(_) => event["status"] = json!("UNCHANGED"),
            Err(_) => {
                event["status"] = json!("SNAPSHOT_INVALID");
                event["detail"] = json!("历史结构快照无法读取：JSONDecodeError");
            }
        },
        Err(error) if error.kind() == io::ErrorKind::NotFound => {
            event["status"] = json!("BASELINE_CREATED");
        }
        Err(error) => {
            event["status"] = json!("SNAPSHOT_INVALID");
            event["detail"] = json!(format!(
                "历史结构快照无法读取：{}",
                python_error_name(&error)
            ));
        }
    }
    Ok(PendingRecord {
        path,
        payload,
        event,
    })
}

fn first_row_header_anchor(paths: &[Vec<String>], table: &ScanTable) -> String {
    for path in paths {
        if let Some((row, _, last_column)) = table.row_hits.get(path).and_then(|hits| hits.first())
        {
            return format!("{}{}", column_letters(*last_column), row);
        }
    }
    String::new()
}

fn region_width(region: &NamedRegion) -> usize {
    region.max_col.saturating_sub(region.min_col) + 1
}

fn numeric_region_shapes(regions: &[NamedRegion]) -> Vec<Vec<usize>> {
    let mut shapes: Vec<Vec<usize>> = regions
        .iter()
        .map(|region| {
            vec![
                region.max_row.saturating_sub(region.min_row) + 1,
                region.max_col.saturating_sub(region.min_col) + 1,
            ]
        })
        .collect();
    shapes.sort();
    shapes
}

fn row_presence_count(value: &Value) -> usize {
    value
        .as_object()
        .map(|presence| {
            presence
                .values()
                .filter(|count| count.as_u64().is_some_and(|count| count > 0))
                .count()
        })
        .unwrap_or_default()
}

fn path_token(secret: &[u8], path: &[String]) -> Result<String, String> {
    path.iter()
        .map(|part| {
            let characters: Vec<char> = part.chars().collect();
            let payload = serde_json::to_vec(&characters).map_err(|error| error.to_string())?;
            Ok(hmac_sha256(secret, &payload))
        })
        .collect::<Result<Vec<_>, String>>()
        .map(|tokens| tokens.join("|"))
}

fn hmac_sha256(secret: &[u8], payload: &[u8]) -> String {
    let mut key_block = [0u8; 64];
    if secret.len() > key_block.len() {
        key_block[..32].copy_from_slice(&Sha256::digest(secret));
    } else {
        key_block[..secret.len()].copy_from_slice(secret);
    }
    let mut inner_pad = [0x36u8; 64];
    let mut outer_pad = [0x5cu8; 64];
    for index in 0..64 {
        inner_pad[index] ^= key_block[index];
        outer_pad[index] ^= key_block[index];
    }
    let mut inner = Sha256::new();
    inner.update(inner_pad);
    inner.update(payload);
    let inner_digest = inner.finalize();
    let mut outer = Sha256::new();
    outer.update(outer_pad);
    outer.update(inner_digest);
    format!("{:x}", outer.finalize())
}

fn sha256_hex(payload: &[u8]) -> String {
    format!("{:x}", Sha256::digest(payload))
}

fn read_or_create_secret(directory: &Path) -> Result<Vec<u8>, String> {
    fs::create_dir_all(directory).map_err(|error| format!("创建结构快照目录失败：{error}"))?;
    let path = directory.join(SECRET_NAME);
    match fs::read(&path) {
        Ok(secret) => validate_secret(secret),
        Err(error) if error.kind() == io::ErrorKind::NotFound => {
            let mut secret = vec![0u8; 32];
            fill_random(&mut secret).map_err(|error| format!("生成结构快照密钥失败：{error}"))?;
            match write_secret_file(&path, &secret) {
                Ok(()) => Ok(secret),
                Err(error) if error.kind() == io::ErrorKind::AlreadyExists => validate_secret(
                    fs::read(&path)
                        .map_err(|read_error| format!("读取结构快照密钥失败：{read_error}"))?,
                ),
                Err(error) => Err(format!("写入结构快照密钥失败：{error}")),
            }
        }
        Err(error) => Err(format!("读取结构快照密钥失败：{error}")),
    }
}

fn validate_secret(secret: Vec<u8>) -> Result<Vec<u8>, String> {
    if secret.len() != 32 {
        return Err("结构快照密钥长度无效".to_owned());
    }
    Ok(secret)
}

fn write_secret_if_missing(directory: &Path, secret: &[u8]) -> io::Result<()> {
    let path = directory.join(SECRET_NAME);
    match write_secret_file(&path, secret) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == io::ErrorKind::AlreadyExists => Ok(()),
        Err(error) => Err(error),
    }
}

fn write_secret_file(path: &Path, secret: &[u8]) -> io::Result<()> {
    let mut options = OpenOptions::new();
    options.write(true).create_new(true);
    #[cfg(unix)]
    {
        use std::os::unix::fs::OpenOptionsExt;
        options.mode(0o600);
    }
    let mut file = options.open(path)?;
    file.write_all(secret)?;
    file.sync_all()
}

#[cfg(unix)]
fn fill_random(buffer: &mut [u8]) -> io::Result<()> {
    File::open("/dev/urandom")?.read_exact(buffer)
}

#[cfg(windows)]
fn fill_random(buffer: &mut [u8]) -> io::Result<()> {
    use std::ffi::c_void;
    #[link(name = "bcrypt")]
    extern "system" {
        fn BCryptGenRandom(
            algorithm: *mut c_void,
            output: *mut u8,
            output_len: u32,
            flags: u32,
        ) -> i32;
    }
    let length = u32::try_from(buffer.len())
        .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "random buffer too long"))?;
    // SAFETY: BCryptGenRandom writes exactly `length` bytes into the valid
    // mutable slice and uses the OS-preferred system RNG.
    let status = unsafe {
        BCryptGenRandom(
            std::ptr::null_mut(),
            buffer.as_mut_ptr(),
            length,
            0x0000_0002,
        )
    };
    if status < 0 {
        return Err(io::Error::from_raw_os_error(status));
    }
    Ok(())
}

#[cfg(not(any(unix, windows)))]
fn fill_random(_buffer: &mut [u8]) -> io::Result<()> {
    Err(io::Error::new(
        io::ErrorKind::Unsupported,
        "OS random source is not implemented for this target",
    ))
}

fn write_snapshot_atomic(path: &Path, payload: &Value) -> io::Result<()> {
    let serialized = serde_json::to_vec_pretty(payload)
        .map_err(|error| io::Error::new(io::ErrorKind::InvalidData, error))?;
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos())
        .unwrap_or_default();
    let file_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("snapshot.json");
    let temporary =
        path.with_file_name(format!("{file_name}.{}.{}.tmp", std::process::id(), nonce));
    let result = (|| {
        let mut file = OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)?;
        file.write_all(&serialized)?;
        file.sync_all()?;
        replace_file(&temporary, path)
    })();
    if result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    result
}

#[cfg(windows)]
fn replace_file(source: &Path, destination: &Path) -> io::Result<()> {
    use std::os::windows::ffi::OsStrExt;
    #[link(name = "kernel32")]
    extern "system" {
        fn MoveFileExW(existing: *const u16, replacement: *const u16, flags: u32) -> i32;
    }
    let source: Vec<u16> = source.as_os_str().encode_wide().chain(Some(0)).collect();
    let destination: Vec<u16> = destination
        .as_os_str()
        .encode_wide()
        .chain(Some(0))
        .collect();
    // SAFETY: both paths are NUL-terminated UTF-16 buffers kept alive for the
    // call. MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH provides the
    // atomic replace semantics used by pathlib.Path.replace on the baseline.
    let result = unsafe { MoveFileExW(source.as_ptr(), destination.as_ptr(), 0x1 | 0x8) };
    if result == 0 {
        return Err(io::Error::last_os_error());
    }
    Ok(())
}

#[cfg(not(windows))]
fn replace_file(source: &Path, destination: &Path) -> io::Result<()> {
    fs::rename(source, destination)
}

fn python_error_name(error: &io::Error) -> &'static str {
    match error.kind() {
        io::ErrorKind::PermissionDenied => "PermissionError",
        io::ErrorKind::InvalidData => "UnicodeDecodeError",
        io::ErrorKind::NotFound => "FileNotFoundError",
        _ => "OSError",
    }
}

fn column_letters(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = ((column - 1) % 26) as u8;
        letters.insert(0, (b'A' + digit) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().collect()
}

#[cfg(test)]
mod tests {
    use super::{hmac_sha256, PendingSnapshots, S5StructureObservation};
    use crate::{s5_layout::NamedRegion, s5_scan::ScanTable};
    use excel_writer::{Color, Format, Workbook};
    use serde_json::Value;
    use std::{collections::BTreeMap, fs, time::SystemTime};

    fn observation(row_label: &str, row: usize, column: usize) -> S5StructureObservation {
        let region = NamedRegion {
            name: "表结构区域_1".to_owned(),
            sheet: "贷款表".to_owned(),
            reference: format!("贷款表!A1:D{}", row + 3),
            min_row: 1,
            min_col: 1,
            max_row: row + 3,
            max_col: 4,
        };
        S5StructureObservation {
            role: "current".to_owned(),
            source_id: "CURRENT".to_owned(),
            business: "单位贷款".to_owned(),
            sheet: "贷款表".to_owned(),
            table_id: "表结构区域_1#1".to_owned(),
            table: ScanTable {
                region,
                column_region: None,
                row_hits: BTreeMap::from([(vec![row_label.to_owned()], vec![(row, 1, 1)])]),
                column_hits: BTreeMap::from([(vec!["本期".to_owned()], vec![column])]),
            },
            expected_row_paths: vec![vec![row_label.to_owned()]],
            expected_column_paths: vec![vec!["本期".to_owned()]],
            row_merge_shapes: vec!["R0C0:R0C1".to_owned()],
            column_merge_shapes: vec![],
            numeric_regions: vec![NamedRegion {
                name: "数值区域_1".to_owned(),
                sheet: "贷款表".to_owned(),
                reference: format!("贷款表!B{}:C{}", row, row + 2),
                min_row: row,
                min_col: 2,
                max_row: row + 2,
                max_col: 3,
            }],
        }
    }

    fn temporary_directory(label: &str) -> std::path::PathBuf {
        let nonce = SystemTime::now()
            .duration_since(SystemTime::UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "audit_s5_schema_{label}_{}_{}",
            std::process::id(),
            nonce
        ))
    }

    fn status(events: &[Value]) -> &str {
        events[0]["status"].as_str().expect("event status")
    }

    #[test]
    fn snapshots_accept_positions_but_hold_logical_changes_for_review() {
        let directory = temporary_directory("lifecycle");
        let baseline = observation("贷款余额", 5, 2);
        let initial =
            PendingSnapshots::compare(&[baseline.clone()], &directory).expect("compare baseline");
        let initial_events = initial.finish(true);
        assert_eq!(status(&initial_events), "BASELINE_CREATED");
        let snapshot_path = directory.join(format!(
            "{}.json",
            initial_events[0]["snapshot_id"].as_str().unwrap()
        ));
        let snapshot_before = fs::read(&snapshot_path).expect("read baseline snapshot");
        let text = String::from_utf8(snapshot_before.clone()).expect("snapshot utf8");
        assert!(!text.contains("贷款余额"));
        assert!(!text.contains("本期"));
        assert_eq!(fs::read(directory.join(".schema-key")).unwrap().len(), 32);

        let moved = observation("贷款余额", 6, 3);
        let position = PendingSnapshots::compare(&[moved.clone()], &directory)
            .expect("compare position change");
        let position_events = position.finish(true);
        assert_eq!(status(&position_events), "POSITION_CHANGED");
        let accepted_position = fs::read(&snapshot_path).expect("read accepted position snapshot");
        assert_ne!(accepted_position, snapshot_before);

        let logical = observation("经营贷款余额", 6, 3);
        let changed =
            PendingSnapshots::compare(&[logical], &directory).expect("compare logical change");
        let changed_events = changed.finish(true);
        assert_eq!(status(&changed_events), "STRUCTURE_CHANGED");
        assert_eq!(fs::read(&snapshot_path).unwrap(), accepted_position);
        fs::remove_dir_all(directory).expect("remove schema snapshot test data");
    }

    #[test]
    fn hmac_sha256_matches_the_rfc_4231_known_answer() {
        assert_eq!(
            hmac_sha256(&[0x0b; 20], b"Hi There"),
            "b0344c61d8db38535ca8afceaf0bf12b881dc200c9833da726e9376c2e32cff7"
        );
    }

    #[test]
    fn changed_inputs_prevent_baseline_snapshot_persistence() {
        let directory = temporary_directory("integrity");
        let pending = PendingSnapshots::compare(&[observation("贷款余额", 5, 2)], &directory)
            .expect("compare baseline");
        let events = pending.finish(false);
        assert_eq!(status(&events), "SNAPSHOT_NOT_SAVED");
        assert_eq!(fs::read_dir(&directory).unwrap().count(), 1);
        fs::remove_dir_all(directory).expect("remove schema snapshot test data");
    }

    #[test]
    fn non_object_snapshot_matches_python_attribute_error_handling() {
        let directory = temporary_directory("invalid_json_root");
        let first = PendingSnapshots::compare(&[observation("贷款余额", 5, 2)], &directory)
            .expect("compare baseline");
        let event = first.finish(true).remove(0);
        let snapshot_path =
            directory.join(format!("{}.json", event["snapshot_id"].as_str().unwrap()));
        fs::write(&snapshot_path, "[]").expect("write valid JSON with an invalid root type");

        let pending = PendingSnapshots::compare(&[observation("贷款余额", 5, 2)], &directory)
            .expect("compare malformed snapshot");
        let event = pending.finish(true).remove(0);
        assert_eq!(event["status"], "SNAPSHOT_INVALID");
        assert_eq!(event["detail"], "历史结构快照无法读取：AttributeError");
        fs::remove_dir_all(directory).expect("remove schema snapshot test data");
    }

    #[test]
    fn fallback_sheet_extent_includes_styled_blank_cells_like_openpyxl() {
        let directory = temporary_directory("styled_extent");
        fs::create_dir_all(&directory).expect("create extent test directory");
        let workbook_path = directory.join("styled-extent.xlsx");
        let mut workbook = Workbook::new();
        let styled_blank = Format::new().set_background_color(Color::Yellow);
        let sheet = workbook
            .add_worksheet()
            .set_name("数据表")
            .expect("name test worksheet");
        sheet
            .write_string(0, 0, "表头")
            .expect("write populated cell");
        sheet
            .write_blank(9, 22, &styled_blank)
            .expect("write styled blank cell");
        workbook.save(&workbook_path).expect("save test workbook");

        let book = crate::s5_scan::ScanBook::open(&workbook_path)
            .expect("read worksheet and OOXML extents");
        assert_eq!(book.sheet_max_row("数据表"), 10);
        assert_eq!(book.sheet_max_col("数据表"), 23);
        fs::remove_dir_all(directory).expect("remove extent test directory");
    }
}
