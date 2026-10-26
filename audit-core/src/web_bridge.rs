//! Local HTTP bridge for the bundled, unchanged Audit web frontend.
//!
//! This bridge exposes migrated Audit workflows through the existing frontend.

use crate::{
    period_pairs::{list_period_pairs, report_source_warning},
    s1_region_summary::{s1_region_summary_source_workbooks, summarize_s1_named_regions_selected},
    s1_template::recommend_template_for_workbooks,
    s1_template_inspection::{INSPECTION_ENGINE, inspect_s1_template},
    s2_config_check::check_s2_config,
    s2_service::{S2ProgressEvent, S2RunRequest, run_s2_audit, s2_engine_for_stage},
    s3_current_value_check::{
        input_reader_for_paths, python_regex_scope_matcher, run_current_value_check_files,
    },
    s3_execution_compare::{S3ExpressionSchema, run_execution_compare_files},
    s3_explanation_export::export_selected_explanations,
};
use anyhow::{Context, Result};
use axum::{
    Json, Router,
    body::Bytes,
    extract::{Path as RoutePath, State},
    http::{HeaderMap, HeaderValue, StatusCode, header},
    response::{Html, IntoResponse, Response},
    routing::{get, post},
};
use serde_json::{Value, json};
#[cfg(any(windows, target_os = "macos", target_os = "linux"))]
use std::process::Command;
use std::{
    fs,
    net::{IpAddr, Ipv4Addr, SocketAddr},
    path::{Path, PathBuf},
    sync::{Arc, Mutex, MutexGuard},
    thread,
    time::Instant,
};
use tokio::{net::TcpListener, runtime::Builder};

const ORIGINAL_HTML: &str = include_str!("../frontend/web/index.html");
const APP_ICON_PNG: &[u8] = include_bytes!("../frontend/assets/app-icon.png");
const APP_ICON_ICO: &[u8] = include_bytes!("../frontend/assets/app-icon.ico");

const BRIDGE_JS: &str = r#"
<script>
(function () {
  function callApi(method, args) {
    return fetch('/api/' + method, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(args || [])
    }).then(function (response) {
      return response.json().then(function (data) {
        if (!response.ok || (data && data.error)) {
          throw new Error((data && data.error) || ('HTTP ' + response.status));
        }
        return data.result;
      });
    });
  }
  var api = new Proxy({}, {
    get: function (target, name) {
      if (typeof name !== 'string') { return undefined; }
      return function () {
        return callApi(name, Array.prototype.slice.call(arguments));
      };
    }
  });
  window.bridge = { api: api };
  window.addEventListener('load', function () {
    window.dispatchEvent(new Event('bridgeready'));
    if (typeof openCollectWorkbench === 'function') { openCollectWorkbench(); }
  });
})();
</script>
"#;

#[derive(Clone)]
struct AppState {
    ui: Arc<Mutex<UiState>>,
}

struct UiState {
    value: Value,
    started_at: Option<Instant>,
}

impl AppState {
    fn new(config_path: PathBuf) -> Self {
        let config_path = config_path.to_string_lossy().into_owned();
        let central_common_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("3.0大集中通用配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let central_cross_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("3.2大集中本期数值核对_配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let central_comparison_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("3.1大集中执行比较_配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let history_config = Path::new(&config_path)
            .parent()
            .map(|parent| parent.join("逐笔统计系统_配置.xlsx"))
            .filter(|path| path.is_file())
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let config_report = check_s2_config(&config_path);
        let mut state = json!({
            "appInfo": {"title": "Open-Data-Audit (Rust)"},
            "busy": false,
            "cancelRequested": false,
            "status": "Rust 迁移版已启动；当前接通报表采集 S2 功能",
            "log": [],
            "filePickerMode": "浏览器内置",
            "filePickerSystemNativeAvailable": false,
            "uiTheme": "日间",
            "uiAccent": "默认蓝",
            "confirmBeforeRun": true,
            "showRunDetailLogs": false,
            "showPerformanceDiagnostics": false,
            "exportRunLogs": false,
            "recursiveDepth": -1,
            "developerAccess": {},
            "developerAccessMessage": "",
            "distributionUnlocked": false,
            "distributionProfile": "s1s2",
            "configIssues": {},
            "sourceFiles": [],
            "selectedFiles": [],
            "historyConfig": history_config,
            "extraFiles": [],
            "mixedTemplates": [],
            "explanationFiles": [],
            "s5": null,
            "periodSummary": {},
            "periodProgress": {
                "status": "idle",
                "stage": "",
                "label": "",
                "percent": 0,
                "current": 0,
                "total": 0,
                "message": "",
                "elapsedSeconds": 0
            },
            "centralCrossOutputAuto": true,
            "centralOutputAuto": true,
            "centralFormOutputAuto": true,
            "centralOutput": "",
            "centralFormOutput": "",
            "centralCompareFile": "",
            "centralExplainFile": "",
            "centralCommonConfig": central_common_config,
            "engines": [{"value": "自动", "label": "Rust 迁移功能"}]
        });
        let s2_state = json!({
            "input": "",
            "templateDir": "",
            "template": "",
            "templateManual": false,
            "external": "",
            "output": "",
            "outputPinned": false,
            "pcCurDir": "",
            "pcPreDir": "",
            "pcCentral": "",
            "pcConfig": config_path,
            "pcOutput": "",
            "pcOutputAuto": true,
            "periodPairs": [],
            "periodSourceWarning": "",
            "periodConfigCheck": config_report
        });
        let s3_state = json!({
            "centralCrossCur": "",
            "centralCrossConfig": central_cross_config,
            "centralCrossLastResult": ""
        });
        let s3_execution_state = json!({
            "centralCur": "",
            "centralPre": "",
            "centralComparisonConfig": central_comparison_config,
            "centralExpressionSchema": "FIVE_SEGMENT_V1"
        });
        state
            .as_object_mut()
            .expect("UI state is a JSON object")
            .extend(
                s2_state
                    .as_object()
                    .expect("S2 state is a JSON object")
                    .clone(),
            );
        state
            .as_object_mut()
            .expect("UI state is a JSON object")
            .extend(
                s3_state
                    .as_object()
                    .expect("S3 state is a JSON object")
                    .clone(),
            );
        state
            .as_object_mut()
            .expect("UI state is a JSON object")
            .extend(
                s3_execution_state
                    .as_object()
                    .expect("S3 execution compare state is a JSON object")
                    .clone(),
            );
        Self {
            ui: Arc::new(Mutex::new(UiState {
                value: state,
                started_at: None,
            })),
        }
    }

    fn lock(&self) -> MutexGuard<'_, UiState> {
        self.ui
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    fn get_state(&self) -> Value {
        let ui = self.lock();
        let mut value = ui.value.clone();
        if let Some(started_at) = ui.started_at {
            if value["periodProgress"]["status"] == "running" {
                value["periodProgress"]["elapsedSeconds"] =
                    json!(started_at.elapsed().as_secs_f64());
            }
        }
        value
    }
}

/// Run the local-only frontend/API service. It never binds to a LAN address.
pub fn serve_web_ui(port: u16, config_path: PathBuf) -> Result<()> {
    let runtime = Builder::new_multi_thread()
        .enable_all()
        .build()
        .context("create Rust web runtime")?;
    runtime.block_on(async move {
        let address = SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port);
        let listener = TcpListener::bind(address)
            .await
            .with_context(|| format!("listen on http://{address}"))?;
        let address = listener
            .local_addr()
            .context("read the bound Audit UI address")?;
        let app = router(AppState::new(config_path));
        let url = format!("http://{address}");
        println!("Open-Data-Audit Rust web UI: {url}");
        if let Err(error) = open_default_browser(&url) {
            eprintln!("无法自动打开浏览器，请手动访问 {url}：{error}");
        }
        axum::serve(listener, app)
            .await
            .context("serve Audit web UI")
    })
}

fn open_default_browser(url: &str) -> std::io::Result<()> {
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;

        const CREATE_NO_WINDOW: u32 = 0x0800_0000;
        Command::new("cmd")
            .args(["/C", "start", "", url])
            .creation_flags(CREATE_NO_WINDOW)
            .spawn()?;
    }

    #[cfg(target_os = "macos")]
    {
        Command::new("open").arg(url).spawn()?;
    }

    #[cfg(target_os = "linux")]
    {
        Command::new("xdg-open").arg(url).spawn()?;
    }

    #[cfg(not(any(windows, target_os = "macos", target_os = "linux")))]
    {
        let _ = url;
        return Err(std::io::Error::new(
            std::io::ErrorKind::Unsupported,
            "automatic browser launch is unsupported on this platform",
        ));
    }

    Ok(())
}

fn router(state: AppState) -> Router {
    Router::new()
        .route("/", get(index))
        .route("/favicon.ico", get(favicon))
        .route("/assets/app-icon.png", get(app_icon_png))
        .route("/api/{method}", post(call_api))
        .with_state(state)
}

fn render_index() -> String {
    let mut page = ORIGINAL_HTML.to_owned();
    for (old, new) in [
        ("centralCrossCur:'csv'", "centralCrossCur:'central_data'"),
        (
            "选择本期数值核对 CSV（UTF-8/GBK）；XLS/XLSX 尚未接通",
            "选择本期数值核对文件（CSV/XLSX/XLS）",
        ),
        (
            "选择本期数值核对 CSV（UTF-8/GBK）",
            "选择本期数值核对文件（CSV/XLSX/XLS）",
        ),
        (
            "选择单个本期数值核对 CSV（UTF-8/GBK）",
            "选择单个本期数值核对文件（CSV/XLSX/XLS）",
        ),
        (
            "当前 Rust Web 入口暂不支持 XLS/XLSX。",
            "Rust 入口支持 CSV、XLSX 和 XLS。",
        ),
        (
            "选择本期数值核对 CSV",
            "选择本期数值核对文件（CSV/XLSX/XLS）",
        ),
    ] {
        page = page.replace(old, new);
    }
    page.replacen("</head>", &format!("{BRIDGE_JS}</head>"), 1)
}

async fn index() -> Response {
    if !ORIGINAL_HTML.contains("</head>") {
        return (
            StatusCode::INTERNAL_SERVER_ERROR,
            "Bundled Audit frontend is missing its head element.",
        )
            .into_response();
    }
    Html(render_index()).into_response()
}

async fn app_icon_png() -> Response {
    (
        [(header::CONTENT_TYPE, HeaderValue::from_static("image/png"))],
        APP_ICON_PNG,
    )
        .into_response()
}

async fn favicon() -> Response {
    (
        [(
            header::CONTENT_TYPE,
            HeaderValue::from_static("image/x-icon"),
        )],
        APP_ICON_ICO,
    )
        .into_response()
}

async fn call_api(
    State(state): State<AppState>,
    RoutePath(method): RoutePath<String>,
    headers: HeaderMap,
    body: Bytes,
) -> Response {
    if let Err(error) = validate_origin(&headers) {
        return error.into_response();
    }
    let args = match serde_json::from_slice::<Value>(&body) {
        Ok(Value::Array(args)) => args,
        Ok(_) => {
            return ApiError::new(StatusCode::BAD_REQUEST, "请求体必须是 JSON 数组")
                .into_response();
        }
        Err(error) => {
            return ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("请求体不是有效的 JSON：{error}"),
            )
            .into_response();
        }
    };

    match dispatch_api(&state, &method, &args) {
        Ok(result) => Json(json!({"result": result})).into_response(),
        Err(error) => error.into_response(),
    }
}

fn validate_origin(headers: &HeaderMap) -> std::result::Result<(), ApiError> {
    let Some(origin) = headers.get(header::ORIGIN) else {
        return Ok(());
    };
    let Some(host) = headers
        .get(header::HOST)
        .and_then(|value| value.to_str().ok())
    else {
        return Err(ApiError::new(StatusCode::FORBIDDEN, "本机页面来源校验失败"));
    };
    let host_name = host
        .split(':')
        .next()
        .unwrap_or_default()
        .trim_matches(['[', ']']);
    if !matches!(host_name, "127.0.0.1" | "localhost") {
        return Err(ApiError::new(
            StatusCode::FORBIDDEN,
            "仅允许本机页面调用接口",
        ));
    }
    let origin = origin.to_str().unwrap_or_default();
    if origin != format!("http://{host}") {
        return Err(ApiError::new(
            StatusCode::FORBIDDEN,
            "仅允许同源本机页面调用接口",
        ));
    }
    Ok(())
}

fn dispatch_api(
    state: &AppState,
    method: &str,
    args: &[Value],
) -> std::result::Result<Value, ApiError> {
    match method {
        "get_state" => Ok(state.get_state()),
        "check_period_audit_config" => check_period_audit_config(state),
        "choose_period_compare" => choose_period_compare(state, args),
        "choose_central" => choose_central(state, args),
        "choose_folder" => choose_folder(state, args),
        "choose_file" => choose_file(state, args),
        "choose_history_config" => choose_history_config(state, args),
        "clear_path" => clear_path(state, args),
        "recognize" => recognize_s1_sources(state),
        "set_selected_files" => set_selected_files(state, args),
        "inspect_template_formulas" => inspect_template_formulas(state),
        "start_dag_flow" => start_dag_flow(state, args),
        "start_central_comparison" => start_central_comparison(state),
        "start_central_explanation_export" => start_central_explanation_export(state),
        "start_central_cross_period" => start_central_cross_period(state),
        "toggle_central_output_pin" => toggle_central_output_pin(state, args),
        "toggle_period_output_pin" => toggle_period_output_pin(state),
        "browse_dir" => browse_dir(state, args),
        "start_period_compare_v2" => start_period_compare_v2(state),
        "update" => update_state(state, args),
        _ => Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版尚未实现接口：{method}"),
        )),
    }
}

fn choose_folder(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let field = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "choose_folder 缺少路径类型"))?;
    if !matches!(field, "input" | "templateDir" | "output") {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版暂不支持选择目录：{field}"),
        ));
    }
    let selected = arg_string(args, 1).unwrap_or_default().trim();
    if selected.is_empty() {
        return Ok(state.lock().value[field].clone());
    }
    let path = PathBuf::from(unquote_path(selected));
    if !path.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("所选目录不存在：{}", path.display()),
        ));
    }
    let path = fs::canonicalize(&path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析所选目录失败：{error}"),
        )
    })?;
    Ok(json!(path.to_string_lossy()))
}

fn choose_history_config(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let selected = arg_string(args, 0).unwrap_or_default().trim();
    if selected.is_empty() {
        let current = state.lock().value["historyConfig"]
            .as_str()
            .unwrap_or_default()
            .to_owned();
        if current.is_empty() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                "请先选择逐笔统计系统配置工作簿",
            ));
        }
        return Ok(json!(current));
    }
    let path = PathBuf::from(unquote_path(selected));
    if !path.is_file()
        || !path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "逐笔统计系统配置必须是存在的 .xlsx 文件：{}",
                path.display()
            ),
        ));
    }
    let path = fs::canonicalize(&path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析逐笔统计系统配置失败：{error}"),
        )
    })?;
    let value = path.to_string_lossy().into_owned();
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改逐笔统计系统配置",
        ));
    }
    ui.value["historyConfig"] = json!(value);
    push_log(&mut ui.value, "已选择逐笔统计系统历史审核配置".to_owned());
    Ok(json!(value))
}

fn recognize_s1_sources(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (input_dir, template_dir, manual_template, busy) = {
        let ui = state.lock();
        (
            ui.value["input"].as_str().unwrap_or_default().to_owned(),
            ui.value["templateDir"]
                .as_str()
                .unwrap_or_default()
                .to_owned(),
            ui.value["templateManual"] == true,
            ui.value["busy"] == true,
        )
    };
    if busy {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核正在运行，暂不能识别模板",
        ));
    }
    let input_dir = PathBuf::from(input_dir);
    if !input_dir.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "请先选择有效的源数据目录",
        ));
    }
    let files = s1_region_summary_source_workbooks(&input_dir, true)
        .map_err(|error| ApiError::internal(format!("扫描 S1 源数据目录失败：{error:#}")))?;
    let template_dir = PathBuf::from(template_dir);
    let can_match_templates = template_dir.is_dir();
    let mut source_rows = Vec::with_capacity(files.len());
    let mut matched_templates = std::collections::BTreeSet::new();
    for path in &files {
        let matched_template = if can_match_templates {
            let suggestion =
                recommend_template_for_workbooks(&template_dir, std::slice::from_ref(path))
                    .map_err(|error| ApiError::internal(format!("识别 S1 模板失败：{error:#}")))?;
            if suggestion.matched {
                suggestion.template_path
            } else {
                None
            }
        } else {
            None
        };
        let template_label = matched_template
            .as_ref()
            .and_then(|path| path.file_name())
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| "未识别".to_owned());
        if let Some(path) = &matched_template {
            matched_templates.insert(path.clone());
        }
        source_rows.push(json!({
            "name": path.file_name().unwrap_or_default().to_string_lossy(),
            "path": path,
            "template": template_label
        }));
    }
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核已开始，不能更新待处理文件列表",
        ));
    }
    ui.value["sourceFiles"] = json!(source_rows);
    ui.value["selectedFiles"] = json!(files);
    ui.value["mixedTemplates"] = json!(
        matched_templates
            .iter()
            .filter_map(|path| path.file_name())
            .map(|name| name.to_string_lossy().into_owned())
            .collect::<Vec<_>>()
    );
    if !manual_template && matched_templates.len() == 1 {
        if let Some(template) = matched_templates.iter().next() {
            ui.value["template"] = json!(template.to_string_lossy());
        }
    }
    ui.value["status"] = json!(format!("已扫描 {} 个 S1-F02 源工作簿", files.len()));
    push_log(
        &mut ui.value,
        format!("S1 源文件扫描完成：{} 个可汇总工作簿", files.len()),
    );
    Ok(ui.value.clone())
}

fn set_selected_files(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let Some(paths) = args.first().and_then(Value::as_array) else {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "set_selected_files 需要路径数组",
        ));
    };
    let selected = paths
        .iter()
        .map(|value| {
            value
                .as_str()
                .map(str::to_owned)
                .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "已选源文件路径必须是字符串"))
        })
        .collect::<std::result::Result<Vec<_>, _>>()?;
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改待处理文件",
        ));
    }
    let available = ui.value["sourceFiles"]
        .as_array()
        .into_iter()
        .flatten()
        .filter_map(|file| file["path"].as_str())
        .collect::<std::collections::HashSet<_>>();
    if let Some(path) = selected
        .iter()
        .find(|path| !available.contains(path.as_str()))
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("待处理文件不在当前扫描清单中：{path}"),
        ));
    }
    ui.value["selectedFiles"] = json!(selected);
    Ok(json!(true))
}

fn choose_file(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let field = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "choose_file 缺少路径类型"))?;
    if field != "template" {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版暂不支持选择文件：{field}"),
        ));
    }
    let selected = arg_string(args, 1).unwrap_or_default().trim();
    if selected.is_empty() {
        return Ok(json!(""));
    }
    let path = PathBuf::from(unquote_path(selected));
    if !path.is_file() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("所选模板文件不存在：{}", path.display()),
        ));
    }
    if !path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
        })
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "模板公式检查目前只支持 .xlsx 和 .xlsm",
        ));
    }

    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核正在运行，暂不能更改路径",
        ));
    }
    ui.value["template"] = json!(path.to_string_lossy());
    ui.value["templateManual"] = json!(true);
    push_log(&mut ui.value, "S1 模板：已选择待检查模板".to_owned());
    Ok(ui.value["template"].clone())
}

fn start_dag_flow(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let workflow = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "start_dag_flow 缺少流程名称"))?;
    if workflow != "dag:汇总校验结果说明" {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版尚未接通工作流：{workflow}"),
        ));
    }
    let values = args.get(1).and_then(Value::as_object).ok_or_else(|| {
        ApiError::new(StatusCode::BAD_REQUEST, "启动 S1-F02 需要当前 GUI 路径设置")
    })?;
    let input_dir = values
        .get("input")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择 S1-F02 源数据目录"))?;
    if !input_dir.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("S1-F02 源数据目录不存在：{}", input_dir.display()),
        ));
    }
    let input_dir = fs::canonicalize(&input_dir).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析 S1-F02 源目录失败：{error}"),
        )
    })?;
    let source_files = s1_region_summary_source_workbooks(&input_dir, true)
        .map_err(|error| ApiError::internal(format!("扫描 S1-F02 源文件失败：{error:#}")))?;
    if source_files.is_empty() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "源数据目录及其子目录中没有可汇总的 .xlsx 文件",
        ));
    }
    let requested_selection = args
        .get(2)
        .and_then(Value::as_array)
        .map(|values| {
            values
                .iter()
                .map(|value| {
                    value.as_str().map(PathBuf::from).ok_or_else(|| {
                        ApiError::new(StatusCode::BAD_REQUEST, "已选源文件路径必须是字符串")
                    })
                })
                .collect::<std::result::Result<Vec<_>, _>>()
        })
        .transpose()?
        .unwrap_or_default();
    let selected_files = if requested_selection.is_empty() {
        source_files
    } else {
        let eligible = source_files
            .iter()
            .map(|path| fs::canonicalize(path))
            .collect::<std::io::Result<std::collections::HashSet<_>>>()
            .map_err(|error| ApiError::internal(format!("解析 S1-F02 源文件路径失败：{error}")))?;
        let mut selected = Vec::new();
        for path in requested_selection {
            let canonical = fs::canonicalize(&path).map_err(|error| {
                ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("已选 S1-F02 源文件不可用：{}：{error}", path.display()),
                )
            })?;
            if !eligible.contains(&canonical) {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!(
                        "已选文件不在当前可汇总源目录或扫描范围内：{}",
                        path.display()
                    ),
                ));
            }
            selected.push(canonical);
        }
        selected.sort();
        selected.dedup();
        selected
    };
    if selected_files.is_empty() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "没有选中可汇总的 S1-F02 源文件",
        ));
    }

    let template_path = values
        .get("template")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from);
    let template_path = match template_path {
        Some(path) => path,
        None => {
            let template_dir = values
                .get("templateDir")
                .and_then(Value::as_str)
                .filter(|path| !path.trim().is_empty())
                .map(PathBuf::from)
                .filter(|path| path.is_dir())
                .ok_or_else(|| {
                    ApiError::new(
                        StatusCode::BAD_REQUEST,
                        "请先选择匹配模板，或选择模板目录后自动识别",
                    )
                })?;
            let mut matched = std::collections::BTreeSet::new();
            let mut unmatched = Vec::new();
            for source in &selected_files {
                let suggestion =
                    recommend_template_for_workbooks(&template_dir, std::slice::from_ref(source))
                        .map_err(|error| {
                        ApiError::internal(format!("自动识别 S1-F02 模板失败：{error:#}"))
                    })?;
                if suggestion.matched {
                    if let Some(path) = suggestion.template_path {
                        matched.insert(path);
                    } else {
                        unmatched.push(source.display().to_string());
                    }
                } else {
                    unmatched.push(source.display().to_string());
                }
            }
            if !unmatched.is_empty() || matched.len() != 1 {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!(
                        "自动识别未得到唯一模板；未匹配文件 {} 个、匹配模板 {} 个。请按模板筛选文件或手动选择模板。",
                        unmatched.len(),
                        matched.len()
                    ),
                ));
            }
            matched.into_iter().next().expect("one matched template")
        }
    };
    if !template_path.is_file()
        || !template_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "S1-F02 模板必须是存在的 .xlsx 文件：{}",
                template_path.display()
            ),
        ));
    }
    let template_path = fs::canonicalize(&template_path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析 S1-F02 模板失败：{error}"),
        )
    })?;

    let history_workbook = state.lock().value["historyConfig"]
        .as_str()
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| {
            ApiError::new(
                StatusCode::BAD_REQUEST,
                "请先选择逐笔统计系统配置，以关联历史审核结果",
            )
        })?;
    if !history_workbook.is_file()
        || !history_workbook
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "逐笔统计系统历史配置必须是存在的 .xlsx 文件：{}",
                history_workbook.display()
            ),
        ));
    }
    let history_workbook = fs::canonicalize(&history_workbook).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析历史审核配置失败：{error}"),
        )
    })?;

    let output_directory = values
        .get("output")
        .and_then(Value::as_str)
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .unwrap_or_else(|| input_dir.join("执行结果"));
    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("S1-F02 输出位置不是目录：{}", output_directory.display()),
        ));
    }
    fs::create_dir_all(&output_directory)
        .map_err(|error| ApiError::internal(format!("创建 S1-F02 输出目录失败：{error}")))?;
    let output_directory = fs::canonicalize(&output_directory)
        .map_err(|error| ApiError::internal(format!("解析 S1-F02 输出目录失败：{error}")))?;
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S").to_string();
    let mut output_workbook = output_directory.join(format!("校验结果与报送说明汇总_{stamp}.xlsx"));
    if output_workbook.exists() {
        output_workbook = output_directory.join(format!(
            "校验结果与报送说明汇总_{stamp}_{:03}.xlsx",
            chrono::Local::now().timestamp_subsec_millis()
        ));
    }
    if output_workbook.exists() {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "S1-F02 时间戳输出文件已存在，未覆盖；请稍后重试",
        ));
    }

    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行 S1-F02",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在汇总校验结果说明，请勿关闭窗口……");
        ui.value["template"] = json!(template_path.to_string_lossy());
        ui.value["auditSummaryOutput"] = json!(output_workbook.to_string_lossy());
        push_log(
            &mut ui.value,
            format!("S1-F02 开始汇总：已选 {} 个源工作簿", selected_files.len()),
        );
        push_detail_log(
            &mut ui.value,
            format!(
                "S1-F02 输入：{}；模板：{}；历史配置：{}；reader=calamine+namespace_aware_ooxml；writer=rust_xlsxwriter；输出：{}",
                selected_files.len(),
                template_path.display(),
                history_workbook.display(),
                output_workbook.display()
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s1-region-summary".to_owned())
        .spawn(move || {
            run_s1_region_summary_worker(
                worker_state,
                template_path,
                input_dir,
                output_workbook,
                history_workbook,
                selected_files,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("S1-F02 汇总启动失败");
        push_log(&mut ui.value, format!("启动 S1-F02 后台任务失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S1-F02 后台任务失败：{error}"
        )));
    }
    Ok(json!(true))
}

fn run_s1_region_summary_worker(
    state: AppState,
    template_path: PathBuf,
    input_dir: PathBuf,
    output_workbook: PathBuf,
    history_workbook: PathBuf,
    selected_files: Vec<PathBuf>,
) {
    let started = Instant::now();
    let feature_names = ["任意行汇总".to_owned()];
    let result = summarize_s1_named_regions_selected(
        &template_path,
        &input_dir,
        &output_workbook,
        Some(&history_workbook),
        Some(&history_workbook),
        &feature_names,
        true,
        &selected_files,
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(summary) => {
            let path = summary.output_path.to_string_lossy().into_owned();
            ui.value["auditSummaryOutput"] = json!(path);
            ui.value["status"] = json!(format!("S1-F02 汇总完成：{} 行", summary.data_rows));
            push_log(
                &mut ui.value,
                format!(
                    "S1-F02 汇总完成：{} 个源工作簿，结构通过 {} 个、部分继续 {} 个、跳过 {} 个；{} 张输出表、{} 行；文件 {}",
                    summary.source_workbooks,
                    summary.structure_preflight_passed_workbooks,
                    summary.structure_preflight_partial_workbooks,
                    summary.structure_preflight_skipped_workbooks,
                    summary.output_sheets,
                    summary.data_rows,
                    summary.output_path.display()
                ),
            );
            for detail in &summary.structure_preflight_skipped_details {
                push_log(
                    &mut ui.value,
                    format!("S1-F02 结构预检跳过文件：{detail}"),
                );
            }
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F02 完成：reader={}；writer={}；结构预检={} 个，通过={}，部分继续={}，跳过={}；历史匹配={} 行；扫描={} ms；模板读取={} ms；结构预检={} ms；源读取={} ms；汇总={} ms；历史富化={} ms；写入={} ms；读回校验={} ms；任务总计={} ms",
                    summary.reader,
                    summary.writer,
                    summary.structure_preflight_checked_workbooks,
                    summary.structure_preflight_passed_workbooks,
                    summary.structure_preflight_partial_workbooks,
                    summary.structure_preflight_skipped_workbooks,
                    summary.history_rows_matched,
                    summary.discovery_time_ms,
                    summary.template_read_time_ms,
                    summary.structure_preflight_time_ms,
                    summary.source_read_time_ms,
                    summary.calculation_time_ms,
                    summary.history_elapsed_ms,
                    summary.write_time_ms,
                    summary.validation_time_ms,
                    started.elapsed().as_millis()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("S1-F02 汇总失败");
            push_log(
                &mut ui.value,
                format!("S1-F02 汇总失败，未报告成功：{error:#}"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S1-F02 失败：reader=calamine+namespace_aware_ooxml；writer=rust_xlsxwriter；elapsed_ms={}；error={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn choose_central(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let field = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "choose_central 缺少路径类型"))?;
    let is_file = match field {
        "centralCur"
        | "centralPre"
        | "centralCompareFile"
        | "centralExplainFile"
        | "centralCommonConfig"
        | "centralComparisonConfig"
        | "centralCrossCur"
        | "centralCrossConfig" => true,
        "centralOutput" | "centralFormOutput" => false,
        _ => {
            return Err(ApiError::new(
                StatusCode::NOT_IMPLEMENTED,
                format!("Rust 迁移版暂不支持选择大集中路径：{field}"),
            ));
        }
    };
    let selected = arg_string(args, 1).unwrap_or_default().trim();
    if selected.is_empty() {
        return Ok(state.lock().value[field].clone());
    }
    let path = PathBuf::from(unquote_path(selected));
    if is_file {
        if !path.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("所选大集中工作簿不存在：{}", path.display()),
            ));
        }
        let extension = path
            .extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or_default();
        let supported = match field {
            "centralCur" | "centralPre" => {
                extension.eq_ignore_ascii_case("csv")
                    || extension.eq_ignore_ascii_case("xlsx")
                    || extension.eq_ignore_ascii_case("xls")
            }
            "centralCompareFile" => {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }
            "centralExplainFile" | "centralCommonConfig" | "centralComparisonConfig" => {
                extension.eq_ignore_ascii_case("xlsx")
            }
            "centralCrossCur" => {
                extension.eq_ignore_ascii_case("csv")
                    || extension.eq_ignore_ascii_case("xlsx")
                    || extension.eq_ignore_ascii_case("xls")
            }
            "centralCrossConfig" => extension.eq_ignore_ascii_case("xlsx"),
            _ => false,
        };
        if !supported {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                match field {
                    "centralCur" | "centralPre" => "大集中本期/上期数据目前只支持 CSV、XLSX、XLS",
                    "centralCompareFile" => "S3-F02 比较结果目前只支持 .xlsx 和 .xlsm",
                    "centralExplainFile" => "说明工作簿目前只支持 .xlsx",
                    "centralCommonConfig" => "大集中通用配置目前只支持 .xlsx",
                    "centralCrossCur" => "S3-F04 只支持 CSV、XLSX、XLS 输入",
                    "centralCrossConfig" => "本期数值核对配置目前只支持 .xlsx",
                    _ => "大集中执行比较配置目前只支持 .xlsx",
                },
            ));
        }
    } else if !path.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("所选大集中输出目录不存在：{}", path.display()),
        ));
    }
    let path = fs::canonicalize(&path).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("解析大集中路径失败：{error}"),
        )
    })?;

    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改大集中路径",
        ));
    }
    let value = path.to_string_lossy().into_owned();
    ui.value[field] = json!(value);
    match field {
        "centralCur" => {
            if ui.value["centralOutputAuto"] == true {
                ui.value["centralOutput"] = json!(
                    path.parent()
                        .unwrap_or_else(|| Path::new("."))
                        .join("执行结果")
                        .to_string_lossy()
                );
            }
        }
        "centralCompareFile" => {
            let parent = path
                .parent()
                .map(|parent| parent.to_string_lossy().into_owned())
                .unwrap_or_default();
            if ui.value["centralOutputAuto"] == true {
                ui.value["centralOutput"] = json!(parent);
            }
            if ui.value["centralFormOutputAuto"] == true {
                ui.value["centralFormOutput"] = json!(parent);
            }
        }
        "centralOutput" => ui.value["centralOutputAuto"] = json!(false),
        "centralFormOutput" => ui.value["centralFormOutputAuto"] = json!(false),
        _ => {}
    }
    push_log(&mut ui.value, format!("S3 路径已选择：{field}"));
    Ok(ui.value[field].clone())
}

fn start_central_comparison(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (current_file, previous_file, common_config, feature_config, output_directory, schema) = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行大集中比较",
            ));
        }
        let current_file = ui.value["centralCur"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择本期大集中数据"))?;
        let previous_file = ui.value["centralPre"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择上期大集中数据"))?;
        for (label, path) in [("本期数据", &current_file), ("上期数据", &previous_file)] {
            if !path.is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}文件不存在：{}", path.display()),
                ));
            }
            let supported = path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("csv")
                        || extension.eq_ignore_ascii_case("xlsx")
                        || extension.eq_ignore_ascii_case("xls")
                });
            if !supported {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}只支持 CSV、XLSX、XLS：{}", path.display()),
                ));
            }
        }

        let common_config = ui.value["centralCommonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择大集中通用配置（3.0）")
            })?;
        let feature_config = ui.value["centralComparisonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择大集中执行比较配置（3.1）")
            })?;
        for (label, path) in [
            ("大集中通用配置（3.0）", &common_config),
            ("大集中执行比较配置（3.1）", &feature_config),
        ] {
            if !path.is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}不存在：{}", path.display()),
                ));
            }
            if !path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
            {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}目前只支持 .xlsx"),
                ));
            }
        }
        let schema = match ui.value["centralExpressionSchema"].as_str() {
            Some("FIVE_SEGMENT_V1") | None | Some("") => S3ExpressionSchema::FiveSegment,
            Some("LEGACY_8") => S3ExpressionSchema::LegacyEightSegment,
            Some(value) => {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("不支持的大集中表达式语法：{value}"),
                ));
            }
        };
        let output_directory = ui.value["centralOutput"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .or_else(|| current_file.parent().map(|parent| parent.join("执行结果")))
            .unwrap_or_else(|| PathBuf::from("执行结果"));
        (
            current_file,
            previous_file,
            common_config,
            feature_config,
            output_directory,
            schema,
        )
    };

    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("大集中比较输出位置不是目录：{}", output_directory.display()),
        ));
    }
    fs::create_dir_all(&output_directory).map_err(|error| {
        ApiError::internal(format!(
            "创建大集中比较输出目录失败：{}：{error}",
            output_directory.display()
        ))
    })?;
    let output_directory = fs::canonicalize(&output_directory)
        .map_err(|error| ApiError::internal(format!("解析大集中比较输出目录失败：{error}")))?;
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S").to_string();
    let mut output_workbook = output_directory.join(format!("比较结果_{stamp}.xlsx"));
    if output_workbook.exists() {
        let suffix = chrono::Local::now().timestamp_subsec_millis();
        output_workbook = output_directory.join(format!("比较结果_{stamp}_{suffix:03}.xlsx"));
        if output_workbook.exists() {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "大集中比较时间戳输出仍有同名文件，请稍后重试",
            ));
        }
    }

    let input_paths = [current_file.clone(), previous_file.clone()];
    let input_reader = input_reader_for_paths(&input_paths);
    let schema_name = match schema {
        S3ExpressionSchema::FiveSegment => "five-segment",
        S3ExpressionSchema::LegacyEightSegment => "legacy-eight-segment",
    };
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行大集中比较",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在执行大集中比较，请勿关闭窗口……");
        push_log(&mut ui.value, "S3-F01 开始大集中执行比较".to_owned());
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F01 开始：input_reader={input_reader}，config_reader=Calamine，writer=rust_xlsxwriter，expression=Rust local ({schema_name})；本期 {}；上期 {}；3.0配置 {}；3.1配置 {}；输出 {}",
                current_file.display(),
                previous_file.display(),
                common_config.display(),
                feature_config.display(),
                output_workbook.display()
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s3-execution-compare".to_owned())
        .spawn(move || {
            run_s3_execution_compare_worker(
                worker_state,
                current_file,
                previous_file,
                common_config,
                feature_config,
                output_workbook,
                schema,
                input_reader,
                schema_name,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("大集中执行比较失败");
        push_log(&mut ui.value, format!("启动 S3-F01 比较线程失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S3-F01 比较线程失败：{error}"
        )));
    }
    Ok(json!(true))
}

#[allow(clippy::too_many_arguments)]
fn run_s3_execution_compare_worker(
    state: AppState,
    current_file: PathBuf,
    previous_file: PathBuf,
    common_config: PathBuf,
    feature_config: PathBuf,
    output_workbook: PathBuf,
    schema: S3ExpressionSchema,
    input_reader: String,
    schema_name: &'static str,
) {
    let started = Instant::now();
    let current_paths = [current_file];
    let previous_paths = [previous_file];
    let progress_state = state.clone();
    let result = run_execution_compare_files(
        &current_paths,
        &previous_paths,
        &common_config,
        &feature_config,
        &output_workbook,
        schema,
        |step| {
            let mut ui = progress_state.lock();
            ui.value["status"] = json!(format!("S3-F01：{step}"));
            push_log(&mut ui.value, format!("S3-F01：{step}"));
            push_detail_log(
                &mut ui.value,
                format!("S3-F01 阶段：{step}；expression={schema_name}"),
            );
        },
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(report) => {
            let output = report.output_path.to_string_lossy().into_owned();
            ui.value["centralCompareFile"] = json!(output);
            if ui.value["centralFormOutputAuto"] == true {
                if let Some(parent) = report.output_path.parent() {
                    ui.value["centralFormOutput"] = json!(parent.to_string_lossy());
                }
            }
            ui.value["status"] = json!(format!("执行比较完成：{} 行", report.rows));
            push_log(
                &mut ui.value,
                format!(
                    "S3-F01 执行比较完成：{} 行，结果文件 {}",
                    report.rows,
                    report.output_path.display()
                ),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F01 完成：reader={}/{}, writer={}，本期/上期记录={}/{}，结果行={}，动作规则触发={}/{}/{}，表达式规则触发={}/{}，表达式错误={}，配置读取={} ms，输入读取={} ms，规则计算={} ms，工作簿写入={} ms，引擎总计={} ms，任务总计={} ms",
                    report.config_reader,
                    report.input_reader,
                    report.writer,
                    report.current_records,
                    report.previous_records,
                    report.rows,
                    report.triggered_action_rules,
                    report.applicable_action_rules,
                    report.enabled_action_rules,
                    report.triggered_expression_rules,
                    report.enabled_expression_rules,
                    report.expression_errors,
                    report.config_read_time_ms,
                    report.input_read_time_ms,
                    report.calculation_time_ms,
                    report.workbook_write_time_ms,
                    report.total_time_ms,
                    started.elapsed().as_millis()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("大集中执行比较失败");
            push_log(
                &mut ui.value,
                format!("S3-F01 执行比较失败，未报告成功结果：{error:#}"),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F01 失败：reader={input_reader}，config_reader=Calamine，writer=rust_xlsxwriter，expression=Rust local ({schema_name})，任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn start_central_explanation_export(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (comparison_file, config_workbook, output_directory) = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再导出指标说明",
            ));
        }
        let comparison_file = ui.value["centralCompareFile"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择比较结果.xlsx"))?;
        if !comparison_file.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("比较结果文件不存在：{}", comparison_file.display()),
            ));
        }
        let config_workbook = ui.value["centralCommonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from);
        let config_workbook = config_workbook
            .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择大集中通用配置.xlsx"))?;
        if !config_workbook.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("大集中通用配置不存在：{}", config_workbook.display()),
            ));
        }
        let output_directory = ui.value["centralOutput"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .or_else(|| {
                ui.value["centralFormOutput"]
                    .as_str()
                    .filter(|path| !path.trim().is_empty())
            })
            .map(PathBuf::from)
            .or_else(|| comparison_file.parent().map(Path::to_path_buf))
            .map(|path| {
                if ui.value["centralOutput"]
                    .as_str()
                    .is_some_and(|value| !value.is_empty())
                    || ui.value["centralFormOutput"]
                        .as_str()
                        .is_some_and(|value| !value.is_empty())
                {
                    path
                } else {
                    path.join("执行结果")
                }
            })
            .unwrap_or_else(|| PathBuf::from("执行结果"));
        (comparison_file, config_workbook, output_directory)
    };

    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("指标说明输出位置不是目录：{}", output_directory.display()),
        ));
    }
    fs::create_dir_all(&output_directory).map_err(|error| {
        ApiError::internal(format!(
            "创建指标说明输出目录失败：{}：{error}",
            output_directory.display()
        ))
    })?;
    let output_workbook = output_directory.join(format!(
        "指标说明_{}.xlsx",
        chrono::Local::now().format("%Y%m%d_%H%M%S")
    ));
    if output_workbook.exists() {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            format!(
                "说明文件输出路径已存在，未覆盖：{}",
                output_workbook.display()
            ),
        ));
    }

    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再导出指标说明",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在导出选定指标说明……");
        push_log(&mut ui.value, "S3-F02 开始导出选定指标说明".to_owned());
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F02 输入比较结果 {}；读取引擎 Calamine；输出引擎 rust_xlsxwriter；大集中通用配置 {}；输出 {}",
                comparison_file.display(),
                config_workbook.display(),
                output_workbook.display()
            ),
        );
    }

    let worker_state = state.clone();
    let thread_output = output_workbook.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s3-explanation-export".to_owned())
        .spawn(move || {
            run_s3_explanation_export_worker(
                worker_state,
                comparison_file,
                Some(config_workbook),
                thread_output,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("指标说明导出失败");
        push_log(&mut ui.value, format!("启动 S3-F02 导出线程失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S3-F02 导出线程失败：{error}"
        )));
    }
    Ok(json!(true))
}

fn run_s3_explanation_export_worker(
    state: AppState,
    comparison_file: PathBuf,
    config_workbook: Option<PathBuf>,
    output_workbook: PathBuf,
) {
    let started = Instant::now();
    let result = export_selected_explanations(
        &comparison_file,
        &output_workbook,
        config_workbook.as_deref(),
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(report) => {
            let output = report.output_workbook.to_string_lossy().into_owned();
            ui.value["centralExplainFile"] = json!(output);
            ui.value["status"] = json!(format!("指标说明导出完成：{} 条", report.selected_rows));
            push_log(
                &mut ui.value,
                format!(
                    "S3-F02 指标说明导出完成：{} 条，文件 {}",
                    report.selected_rows,
                    report.output_workbook.display()
                ),
            );
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F02 完成：reader={}，writer={}，选中行={}，比较读取={} ms，配置读取={} ms，工作簿写入={} ms，引擎总计={} ms，任务总计={} ms",
                    report.reader,
                    report.writer,
                    report.selected_rows,
                    report.comparison_read_time_ms,
                    report.config_read_time_ms,
                    report.workbook_write_time_ms,
                    report.total_time_ms,
                    started.elapsed().as_millis()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("指标说明导出失败");
            push_log(&mut ui.value, format!("S3-F02 指标说明导出失败：{error:#}"));
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F02 失败：reader=Calamine，writer=rust_xlsxwriter，任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn start_central_cross_period(state: &AppState) -> std::result::Result<Value, ApiError> {
    let (current_file, config_workbooks, output_directory, write_flow_logs) = {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行本期数值核对",
            ));
        }
        let current_file = ui.value["centralCrossCur"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择本期数值核对数据文件")
            })?;
        if !current_file.is_file() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("本期数值核对文件不存在：{}", current_file.display()),
            ));
        }
        if !current_file
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("csv")
                    || extension.eq_ignore_ascii_case("xlsx")
                    || extension.eq_ignore_ascii_case("xls")
            })
        {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                "S3-F04 只支持 CSV、XLSX、XLS 输入",
            ));
        }

        let common_config = ui.value["centralCommonConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择大集中通用配置（3.0）")
            })?;
        let cross_config = ui.value["centralCrossConfig"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .ok_or_else(|| {
                ApiError::new(StatusCode::BAD_REQUEST, "请先选择本期数值核对配置（3.2）")
            })?;
        for (label, path) in [
            ("大集中通用配置", &common_config),
            ("本期数值核对配置", &cross_config),
        ] {
            if !path.is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}不存在：{}", path.display()),
                ));
            }
            if !path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
            {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    format!("{label}目前只支持 .xlsx"),
                ));
            }
        }

        let output_directory = ui.value["centralOutput"]
            .as_str()
            .filter(|path| !path.trim().is_empty())
            .map(PathBuf::from)
            .or_else(|| current_file.parent().map(|parent| parent.join("执行结果")))
            .unwrap_or_else(|| PathBuf::from("执行结果"));
        (
            current_file,
            vec![common_config, cross_config],
            output_directory,
            ui.value["exportRunLogs"] == true,
        )
    };

    if output_directory.exists() && !output_directory.is_dir() {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "本期数值核对输出位置不是目录：{}",
                output_directory.display()
            ),
        ));
    }
    fs::create_dir_all(&output_directory).map_err(|error| {
        ApiError::internal(format!(
            "创建本期数值核对输出目录失败：{}：{error}",
            output_directory.display()
        ))
    })?;
    let output_directory = fs::canonicalize(&output_directory)
        .map_err(|error| ApiError::internal(format!("解析本期数值核对输出目录失败：{error}")))?;
    let stamp = chrono::Local::now().format("%Y%m%d_%H%M%S").to_string();
    let mut output_workbook = output_directory.join(format!("数值核对结果_{stamp}.xlsx"));
    let mut log_workbook = write_flow_logs
        .then(|| output_directory.join(format!("本期数值核对_运行日志_{stamp}.xlsx")));
    if output_workbook.exists() || log_workbook.as_ref().is_some_and(|path| path.exists()) {
        let suffix = chrono::Local::now().timestamp_subsec_millis();
        output_workbook = output_directory.join(format!("数值核对结果_{stamp}_{suffix:03}.xlsx"));
        log_workbook = write_flow_logs.then(|| {
            output_directory.join(format!("本期数值核对_运行日志_{stamp}_{suffix:03}.xlsx"))
        });
        if output_workbook.exists() || log_workbook.as_ref().is_some_and(|path| path.exists()) {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "本期数值核对时间戳输出仍有同名文件，请稍后重试",
            ));
        }
    }

    let input_reader = input_reader_for_paths(std::slice::from_ref(&current_file));
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "已有任务正在运行，请完成后再执行本期数值核对",
            ));
        }
        ui.value["busy"] = json!(true);
        ui.value["status"] = json!("正在执行本期数值核对，请勿关闭窗口……");
        push_log(&mut ui.value, "S3-F04 开始本期数值核对".to_owned());
        push_detail_log(
            &mut ui.value,
            format!(
                "S3-F04 开始：reader={}，writer=rust_xlsxwriter；当前数据 {}；通用配置 {}；本期核对配置 {}；结果输出 {}；运行日志={}",
                input_reader,
                current_file.display(),
                config_workbooks[0].display(),
                config_workbooks[1].display(),
                output_workbook.display(),
                log_workbook.is_some()
            ),
        );
    }

    let worker_state = state.clone();
    if let Err(error) = thread::Builder::new()
        .name("audit-s3-current-value-check".to_owned())
        .spawn(move || {
            run_s3_current_value_check_worker(
                worker_state,
                current_file,
                config_workbooks,
                output_workbook,
                log_workbook,
            )
        })
    {
        let mut ui = state.lock();
        ui.value["busy"] = json!(false);
        ui.value["status"] = json!("本期数值核对失败");
        push_log(&mut ui.value, format!("启动 S3-F04 任务线程失败：{error}"));
        return Err(ApiError::internal(format!(
            "启动 S3-F04 任务线程失败：{error}"
        )));
    }
    Ok(json!(true))
}

fn run_s3_current_value_check_worker(
    state: AppState,
    current_file: PathBuf,
    config_workbooks: Vec<PathBuf>,
    output_workbook: PathBuf,
    log_workbook: Option<PathBuf>,
) {
    let started = Instant::now();
    let input_reader = input_reader_for_paths(std::slice::from_ref(&current_file));
    let result = run_current_value_check_files(
        std::slice::from_ref(&current_file),
        &[],
        &config_workbooks,
        &output_workbook,
        log_workbook.as_deref(),
        python_regex_scope_matcher(),
    );
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    match result {
        Ok(report) => {
            let output = report.output_workbook.to_string_lossy().into_owned();
            ui.value["centralCrossLastResult"] = json!(output);
            ui.value["status"] = json!(format!("本期数值核对完成：{} 行", report.rows));
            push_log(
                &mut ui.value,
                format!(
                    "S3-F04 本期数值核对完成：{} 行，结果文件 {}",
                    report.rows,
                    report.output_workbook.display()
                ),
            );
            if let Some(log_path) = report.log_workbook.as_ref() {
                push_log(
                    &mut ui.value,
                    format!("S3-F04 运行日志文件：{}", log_path.display()),
                );
            }
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F04 完成：reader={}/{}, writer={}，规则={}/{}/{}，输入记录={}/{}，结果行={}，配置读取={} ms，输入读取={} ms，规则计算={} ms，工作簿写入={} ms，引擎总计={} ms，任务总计={} ms",
                    report.config_reader,
                    report.input_reader,
                    report.writer,
                    report.enabled_rules,
                    report.total_rules,
                    report.disabled_rules,
                    report.current_records,
                    report.previous_records,
                    report.rows,
                    report.config_read_time_ms,
                    report.input_read_time_ms,
                    report.calculation_time_ms,
                    report.workbook_write_time_ms,
                    report.total_time_ms,
                    started.elapsed().as_millis()
                ),
            );
        }
        Err(error) => {
            ui.value["status"] = json!("本期数值核对失败");
            push_log(&mut ui.value, format!("S3-F04 本期数值核对失败：{error:#}"));
            push_detail_log(
                &mut ui.value,
                format!(
                    "S3-F04 失败：reader={input_reader}，writer=rust_xlsxwriter，任务总计={} ms，错误={error:#}",
                    started.elapsed().as_millis()
                ),
            );
        }
    }
}

fn toggle_central_output_pin(
    state: &AppState,
    args: &[Value],
) -> std::result::Result<Value, ApiError> {
    let kind = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "缺少大集中输出目录类型"))?;
    let (auto_key, source_key) = match kind {
        "centralOutput" => ("centralOutputAuto", "centralCur"),
        "centralFormOutput" => ("centralFormOutputAuto", "centralCompareFile"),
        _ => {
            return Err(ApiError::new(
                StatusCode::NOT_IMPLEMENTED,
                format!("Rust 迁移版暂不支持固定该输出目录：{kind}"),
            ));
        }
    };
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，暂不能更改输出目录",
        ));
    }
    let auto = ui.value[auto_key] != true;
    ui.value[auto_key] = json!(auto);
    if auto {
        if let Some(source) = ui.value[source_key]
            .as_str()
            .filter(|path| !path.is_empty())
        {
            if let Some(parent) = Path::new(source).parent() {
                ui.value[kind] = json!(parent.join("执行结果").to_string_lossy());
            }
        }
    }
    Ok(json!(!auto))
}

fn inspect_template_formulas(state: &AppState) -> std::result::Result<Value, ApiError> {
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "任务正在运行，请完成后再检查模板公式",
        ));
    }
    let template = ui.value["template"]
        .as_str()
        .filter(|path| !path.trim().is_empty())
        .map(PathBuf::from)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "请先选择或自动识别匹配模板"))?;
    let started = Instant::now();
    let (report, read_elapsed_ms) = inspect_s1_template(&template).map_err(|error| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("模板公式检查失败：{error:#}"),
        )
    })?;
    let formula_count = report.formula_count;
    let error_count = report.errors.len();
    let ranges = report
        .named_range_groups
        .iter()
        .filter(|group| group.count > 0)
        .map(|group| format!("{} {}个", group.family, group.count))
        .collect::<Vec<_>>()
        .join("、");
    let other_count = report.other_names.len();
    let internal_count = report.internal_names.len();
    let value = serde_json::to_value(report)
        .map_err(|error| ApiError::internal(format!("序列化模板检查结果失败：{error}")))?;
    push_log(
        &mut ui.value,
        format!(
            "检查模板：共 {formula_count} 个公式；明确错误 {error_count} 处；命名区域 {}；其他 {other_count} 个；内部名称 {internal_count} 个",
            if ranges.is_empty() { "无" } else { &ranges }
        ),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S1-F01 模板检查步骤：读取公式文本与命名区域；引擎 {INSPECTION_ENGINE}；读取耗时 {read_elapsed_ms} ms；总耗时 {} ms",
            started.elapsed().as_millis()
        ),
    );
    Ok(value)
}

fn clear_path(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    if let Some(
        field @ ("centralCompareFile"
        | "centralExplainFile"
        | "centralCur"
        | "centralPre"
        | "centralCommonConfig"
        | "centralComparisonConfig"
        | "centralCrossCur"
        | "centralCrossConfig"
        | "centralOutput"
        | "centralFormOutput"),
    ) = arg_string(args, 0)
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "任务正在运行，暂不能清除大集中路径",
            ));
        }
        ui.value[field] = json!("");
        push_log(&mut ui.value, format!("S3 路径已清除：{field}"));
        return Ok(json!(true));
    }
    if arg_string(args, 0) == Some("template") {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "审核正在运行，暂不能清除路径",
            ));
        }
        ui.value["template"] = json!("");
        ui.value["templateManual"] = json!(false);
        push_log(&mut ui.value, "S1 模板：已清除模板选择".to_owned());
        return Ok(json!(true));
    }
    if let Some(field @ ("input" | "templateDir" | "external" | "output" | "historyConfig")) =
        arg_string(args, 0)
    {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "任务正在运行，暂不能清除路径",
            ));
        }
        ui.value[field] = json!("");
        if field == "historyConfig" {
            push_log(&mut ui.value, "逐笔统计系统历史配置路径已清除".to_owned());
        }
        return Ok(json!(true));
    }
    clear_period_path(state, args)
}
fn check_period_audit_config(state: &AppState) -> std::result::Result<Value, ApiError> {
    let config_path = state.lock().value["pcConfig"]
        .as_str()
        .unwrap_or_default()
        .to_owned();
    let started = Instant::now();
    let report = check_s2_config(config_path);
    let passed = report.passed;
    let blocking_errors = report.blocking_errors;
    let warnings = report.warnings;
    let sheet_count = report.sheet_stats.len();
    let indicator_count = report.indicator_count;
    let strategy_count = report.strategy_count;
    let elapsed_ms = started.elapsed().as_millis();
    let value = serde_json::to_value(report)
        .map_err(|error| ApiError::internal(format!("序列化配置检查结果失败：{error}")))?;
    let mut ui = state.lock();
    ui.value["periodConfigCheck"] = value.clone();
    push_log(
        &mut ui.value,
        format!(
            "S2 配置检查{}：{} 项阻断错误、{} 项警告；耗时 {} ms",
            if passed { "通过" } else { "未通过" },
            blocking_errors,
            warnings,
            elapsed_ms
        ),
    );
    push_detail_log(
        &mut ui.value,
        format!(
            "S2-F01 配置检查步骤：reader=calamine；工作表={sheet_count}；指标={indicator_count}；环比策略={strategy_count}；blocking_errors={blocking_errors}；warnings={warnings}；elapsed_ms={elapsed_ms}"
        ),
    );
    Ok(value)
}

fn choose_period_compare(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let kind = arg_string(args, 0).ok_or_else(|| {
        ApiError::new(
            StatusCode::BAD_REQUEST,
            "choose_period_compare 缺少路径类型",
        )
    })?;
    if !matches!(
        kind,
        "pcCurDir" | "pcPreDir" | "pcCentral" | "pcConfig" | "pcOutput"
    ) {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!("不支持的报表采集路径类型：{kind}"),
        ));
    }
    let selected = arg_string(args, 1).unwrap_or_default().trim().to_owned();
    if selected.is_empty() {
        return Ok(state.lock().value[kind].clone());
    }
    {
        let ui = state.lock();
        if ui.value["busy"] == true {
            return Err(ApiError::new(
                StatusCode::CONFLICT,
                "审核正在运行，暂不能更改路径",
            ));
        }
    }
    let selected_path = PathBuf::from(unquote_path(&selected));
    let is_file = matches!(kind, "pcCentral" | "pcConfig");
    if (is_file && !selected_path.is_file()) || (!is_file && !selected_path.is_dir()) {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            format!(
                "所选{}不存在或类型不正确：{}",
                if is_file { "文件" } else { "目录" },
                selected_path.display()
            ),
        ));
    }
    let absolute = selected_path
        .canonicalize()
        .unwrap_or(selected_path)
        .to_string_lossy()
        .into_owned();
    let mut ui = state.lock();
    ui.value[kind] = json!(absolute);
    if kind == "pcCurDir" && ui.value["pcOutputAuto"] == true {
        ui.value["pcOutput"] = json!(Path::new(&absolute).join("执行结果").to_string_lossy());
    } else if kind == "pcOutput" {
        ui.value["pcOutputAuto"] = json!(false);
    }
    push_log(
        &mut ui.value,
        format!("报表采集：{}已选择", period_path_label(kind)),
    );
    let chosen = ui.value[kind].clone();
    drop(ui);
    refresh_period_pairs(state);
    if kind == "pcConfig" {
        let _ = check_period_audit_config(state)?;
    }
    Ok(chosen)
}

fn toggle_period_output_pin(state: &AppState) -> std::result::Result<Value, ApiError> {
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Ok(json!(ui.value["pcOutputAuto"] != true));
    }
    let auto = ui.value["pcOutputAuto"] != true;
    ui.value["pcOutputAuto"] = json!(auto);
    if auto {
        if let Some(current) = ui.value["pcCurDir"].as_str() {
            if !current.is_empty() {
                ui.value["pcOutput"] = json!(Path::new(current).join("执行结果").to_string_lossy());
            }
        }
    }
    push_log(
        &mut ui.value,
        if auto {
            "报表采集：输出目录跟随本期目录".to_owned()
        } else {
            "报表采集：输出目录已固定".to_owned()
        },
    );
    Ok(json!(!auto))
}

fn clear_period_path(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let kind = arg_string(args, 0)
        .ok_or_else(|| ApiError::new(StatusCode::BAD_REQUEST, "clear_path 缺少路径类型"))?;
    if !matches!(
        kind,
        "pcCurDir" | "pcPreDir" | "pcCentral" | "pcConfig" | "pcOutput"
    ) {
        return Err(ApiError::new(
            StatusCode::NOT_IMPLEMENTED,
            format!("Rust 迁移版暂不支持清除路径：{kind}"),
        ));
    }
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核正在运行，暂不能清除路径",
        ));
    }
    ui.value[kind] = json!("");
    if kind == "pcCurDir" && ui.value["pcOutputAuto"] == true {
        ui.value["pcOutput"] = json!("");
    }
    if kind == "pcOutput" {
        ui.value["pcOutputAuto"] = json!(false);
    }
    push_log(
        &mut ui.value,
        format!("报表采集：已清除{}", period_path_label(kind)),
    );
    drop(ui);
    refresh_period_pairs(state);
    if kind == "pcConfig" {
        let _ = check_period_audit_config(state)?;
    }
    Ok(json!(true))
}

fn browse_dir(_state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let requested = arg_string(args, 0).unwrap_or_default();
    let mode = arg_string(args, 1).unwrap_or_default();
    let root = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    let cleaned = unquote_path(requested.trim());
    let mut target = if cleaned.is_empty() {
        root.clone()
    } else {
        let candidate = PathBuf::from(cleaned);
        if candidate.is_absolute() {
            candidate
        } else {
            root.join(candidate)
        }
    };
    target = match target.canonicalize() {
        Ok(path) => path,
        Err(_) => {
            return Ok(browse_error(
                &root,
                &root,
                "路径不存在或不是可访问的目录",
                "",
            ));
        }
    };
    let extensions = browse_extensions(mode);
    let mut selected_file = String::new();
    if target.is_file() {
        if is_hidden_name(&target) || !extension_allowed(&target, extensions) {
            return Ok(browse_error(
                &target.parent().unwrap_or(&root),
                &root,
                "该文件不符合当前选择类型",
                "",
            ));
        }
        selected_file = target.to_string_lossy().into_owned();
        target = target.parent().unwrap_or(&root).to_path_buf();
    } else if !target.is_dir() {
        return Ok(browse_error(
            &root,
            &root,
            "路径不存在或不是可访问的目录",
            "",
        ));
    }

    let mut directories = Vec::new();
    let mut files = Vec::new();
    let entries = match fs::read_dir(&target) {
        Ok(entries) => entries,
        Err(error) => {
            return Ok(browse_error(
                &target,
                &root,
                &error.to_string(),
                &selected_file,
            ));
        }
    };
    let mut entries = entries.filter_map(|entry| entry.ok()).collect::<Vec<_>>();
    entries.sort_by_key(|entry| entry.file_name().to_string_lossy().to_lowercase());
    for entry in entries {
        let name = entry.file_name().to_string_lossy().into_owned();
        if is_hidden_name(Path::new(&name)) || name == "__pycache__" {
            continue;
        }
        let Ok(file_type) = entry.file_type() else {
            continue;
        };
        let path = entry.path();
        if file_type.is_dir() {
            directories.push(json!({"name": name, "path": path.to_string_lossy()}));
        } else if file_type.is_file() && extension_allowed(&path, extensions) {
            let size = entry.metadata().map(|metadata| metadata.len()).unwrap_or(0);
            let size_text = if size >= 1024 {
                format!("{} KB", (size + 512) / 1024)
            } else {
                format!("{size} B")
            };
            files.push(json!({"name": name, "path": path.to_string_lossy(), "size": size_text}));
        }
    }
    let parent = target
        .parent()
        .filter(|path| *path != target)
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_default();
    Ok(json!({
        "path": target.to_string_lossy(),
        "parent": parent,
        "dirs": directories,
        "files": files,
        "root": root.to_string_lossy(),
        "selectedFile": selected_file,
        "error": ""
    }))
}

fn start_period_compare_v2(state: &AppState) -> std::result::Result<Value, ApiError> {
    let request = {
        let mut ui = state.lock();
        if ui.value["busy"] == true {
            return Ok(json!(false));
        }
        let current = ui.value["pcCurDir"].as_str().unwrap_or_default().to_owned();
        let previous = ui.value["pcPreDir"].as_str().unwrap_or_default().to_owned();
        if current.is_empty() || previous.is_empty() {
            push_log(
                &mut ui.value,
                "跨期比较失败：请先选择本期与上期目录".to_owned(),
            );
            ui.value["status"] = json!("跨期比较失败：请先选择本期与上期目录");
            return Ok(json!(false));
        }
        let current_path = PathBuf::from(&current);
        let previous_path = PathBuf::from(&previous);
        if !current_path.is_dir() || !previous_path.is_dir() {
            return Err(ApiError::new(
                StatusCode::BAD_REQUEST,
                "本期或上期目录已不存在",
            ));
        }
        let config = ui.value["pcConfig"].as_str().unwrap_or_default().to_owned();
        if config.is_empty() || !Path::new(&config).is_file() {
            ui.value["periodConfigCheck"] = serde_json::to_value(check_s2_config(&config))
                .unwrap_or_else(|_| json!({"passed": false, "message": "配置文件不可读取"}));
            ui.value["status"] = json!("跨期比较失败：报表采集系统配置不存在");
            push_log(
                &mut ui.value,
                "跨期比较失败：请先选择有效的报表采集系统配置".to_owned(),
            );
            return Ok(json!(false));
        }
        let config_report = check_s2_config(&config);
        ui.value["periodConfigCheck"] = serde_json::to_value(&config_report)
            .map_err(|error| ApiError::internal(format!("序列化配置检查结果失败：{error}")))?;
        if !config_report.passed {
            ui.value["status"] = json!("跨期比较失败：报表采集系统配置存在阻断错误");
            push_log(&mut ui.value, config_report.message.clone());
            ui.value["periodProgress"] = json!({
                "status": "failed", "stage": "config", "label": "检查审核配置",
                "percent": 0, "current": 0, "total": 0,
                "message": config_report.message, "elapsedSeconds": 0
            });
            return Ok(json!(false));
        }
        let external = ui.value["pcCentral"]
            .as_str()
            .filter(|path| !path.is_empty())
            .map(str::to_owned);
        if let Some(external) = external.as_deref() {
            if !Path::new(external).is_file() {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    "大集中数据文件已不存在",
                ));
            }
        }
        let output = ui.value["pcOutput"]
            .as_str()
            .filter(|path| !path.is_empty())
            .map(PathBuf::from)
            .unwrap_or_else(|| current_path.join("执行结果"));
        ui.value["busy"] = json!(true);
        ui.value["cancelRequested"] = json!(false);
        ui.value["periodSummary"] = json!({});
        ui.value["status"] = json!("正在检查审核配置……");
        ui.value["periodProgress"] = json!({
            "status": "running", "stage": "config", "label": "检查审核配置",
            "percent": 0, "current": 0, "total": 0,
            "message": "准备执行 Rust S2 审核", "elapsedSeconds": 0
        });
        ui.started_at = Some(Instant::now());
        let request = S2RunRequest {
            current_directory: current_path,
            previous_directory: previous_path,
            config_path: PathBuf::from(config),
            output_directory: output,
            external_workbook: external.map(PathBuf::from),
        };
        push_log(
            &mut ui.value,
            format!(
                "Rust S2 开始：本期 {}；上期 {}；输出 {}",
                request.current_directory.display(),
                request.previous_directory.display(),
                request.output_directory.display()
            ),
        );
        request
    };

    let worker_state = state.clone();
    thread::Builder::new()
        .name("audit-s2-service".to_owned())
        .spawn(move || run_s2_worker(worker_state, request))
        .map_err(|error| {
            let mut ui = state.lock();
            ui.value["busy"] = json!(false);
            ui.value["status"] = json!("无法启动 Rust S2 审核线程");
            ApiError::internal(format!("启动审核线程失败：{error}"))
        })?;
    Ok(json!(true))
}

fn run_s2_worker(state: AppState, request: S2RunRequest) {
    let worker_started = Instant::now();
    let mut active_stage: Option<(String, String, Instant)> = None;
    let mut on_progress = |event: S2ProgressEvent| {
        let stage_changed = active_stage
            .as_ref()
            .map(|(stage, _, _)| stage != &event.stage)
            .unwrap_or(true);
        let mut ui = state.lock();
        if stage_changed {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_detail_log(
                    &mut ui.value,
                    format!(
                        "S2 阶段完成：stage={stage} engine={engine} elapsed_ms={}",
                        started.elapsed().as_millis()
                    ),
                );
            }
            let engine = s2_engine_for_stage(&event.stage).unwrap_or("未映射");
            push_detail_log(
                &mut ui.value,
                format!(
                    "S2 阶段开始：stage={} label={} engine={} progress={}/{}",
                    event.stage, event.label, engine, event.current, event.total
                ),
            );
            active_stage = Some((event.stage.clone(), engine.to_owned(), Instant::now()));
        }
        let stage_elapsed = active_stage
            .as_ref()
            .map(|(_, _, started)| started.elapsed().as_secs_f64())
            .unwrap_or_default();
        ui.value["status"] = json!(format!("{}……", event.label));
        ui.value["periodProgress"] = progress_value(
            &event,
            "running",
            worker_started.elapsed().as_secs_f64(),
            stage_elapsed,
        );
    };
    let result = run_s2_audit(&request, Some(&mut on_progress));
    drop(on_progress);
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    ui.started_at = None;
    match result {
        Ok(result) => {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_detail_log(
                    &mut ui.value,
                    format!(
                        "S2 阶段完成：stage={stage} engine={engine} elapsed_ms={}",
                        started.elapsed().as_millis()
                    ),
                );
            }
            let summary = result.analysis.summary;
            let report_path = result.report_path.to_string_lossy().into_owned();
            ui.value["periodSummary"] = json!({
                "result": report_path,
                "institutions": summary.institutions,
                "forms": summary.forms,
                "indicators": summary.indicators,
                "abnormal": summary.abnormal_findings,
                "totalFindings": summary.total_findings,
                "byType": summary.by_type,
                "bySeverity": summary.by_severity
            });
            ui.value["periodProgress"] = json!({
                "status": "completed", "stage": "completed", "label": "审核完成",
                "percent": 100, "current": 1, "total": 1,
                "message": format!("报告已生成：{report_path}"),
                "elapsedSeconds": worker_started.elapsed().as_secs_f64()
            });
            ui.value["status"] = json!("跨期比较完成：已生成审核结果");
            push_detail_log(
                &mut ui.value,
                format!(
                    "S2 任务完成：reader=Calamine writer=rust_xlsxwriter findings={} total_ms={}",
                    summary.total_findings,
                    worker_started.elapsed().as_millis()
                ),
            );
            push_log(&mut ui.value, format!("审核完成：{}", report_path));
        }
        Err(error) => {
            if let Some((stage, engine, started)) = active_stage.take() {
                push_detail_log(
                    &mut ui.value,
                    format!(
                        "S2 阶段失败：stage={stage} engine={engine} elapsed_ms={}",
                        started.elapsed().as_millis()
                    ),
                );
            }
            let message = format!("跨期比较失败：{error:#}");
            ui.value["periodProgress"] = json!({
                "status": "failed", "stage": "failed", "label": "审核失败",
                "percent": 0, "current": 0, "total": 0,
                "message": message,
                "elapsedSeconds": worker_started.elapsed().as_secs_f64()
            });
            ui.value["status"] = json!("跨期比较失败");
            push_detail_log(
                &mut ui.value,
                format!(
                    "S2 任务失败：total_ms={} error={error:#}",
                    worker_started.elapsed().as_millis()
                ),
            );
            push_log(&mut ui.value, message);
        }
    }
}

fn progress_value(
    event: &S2ProgressEvent,
    status: &str,
    elapsed: f64,
    stage_elapsed: f64,
) -> Value {
    let engine = s2_engine_for_stage(&event.stage);
    let label = engine.map_or_else(
        || event.label.clone(),
        |engine| {
            format!(
                "{} · {engine} · 本阶段已运行 {:.2} 秒",
                event.label, stage_elapsed
            )
        },
    );
    json!({
        "status": status,
        "stage": event.stage,
        "label": label,
        "engine": engine,
        "percent": event.percent.clamp(0.0, 99.0),
        "current": event.current,
        "total": event.total,
        "message": event.message,
        "elapsedSeconds": elapsed,
        "stageElapsedSeconds": stage_elapsed
    })
}

fn update_state(state: &AppState, args: &[Value]) -> std::result::Result<Value, ApiError> {
    let Some(updates) = args.first().and_then(Value::as_object) else {
        return Err(ApiError::new(
            StatusCode::BAD_REQUEST,
            "update 需要一个 JSON 对象",
        ));
    };
    let mut ui = state.lock();
    if ui.value["busy"] == true {
        return Err(ApiError::new(
            StatusCode::CONFLICT,
            "审核正在运行，暂不能更改设置",
        ));
    }
    for (key, value) in updates {
        match key.as_str() {
            "input" | "templateDir" | "template" | "external" | "output" | "historyConfig" => {
                let Some(path) = value.as_str() else {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        format!("{key} 必须是字符串"),
                    ));
                };
                ui.value[key] = json!(path.trim());
                if key == "template" && !path.trim().is_empty() {
                    ui.value["templateManual"] = json!(true);
                }
            }
            "uiTheme" | "uiAccent" => {
                if !value.is_string() {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        format!("{key} 必须是字符串"),
                    ));
                }
                ui.value[key] = value.clone();
            }
            "centralExpressionSchema"
                if matches!(value.as_str(), Some("FIVE_SEGMENT_V1" | "LEGACY_8")) =>
            {
                ui.value[key] = value.clone();
            }
            "centralExpressionSchema" => {
                return Err(ApiError::new(
                    StatusCode::BAD_REQUEST,
                    "centralExpressionSchema 必须是 FIVE_SEGMENT_V1 或 LEGACY_8",
                ));
            }
            "confirmBeforeRun"
            | "showRunDetailLogs"
            | "showPerformanceDiagnostics"
            | "exportRunLogs" => {
                if !value.is_boolean() {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        format!("{key} 必须是布尔值"),
                    ));
                }
                ui.value[key] = value.clone();
            }
            "filePickerMode" if value.as_str() == Some("浏览器内置") => {
                ui.value[key] = value.clone();
            }
            "filePickerMode" => {
                return Err(ApiError::new(
                    StatusCode::NOT_IMPLEMENTED,
                    "Rust Web 版当前只支持浏览器内置文件选择器",
                ));
            }
            _ => {
                return Err(ApiError::new(
                    StatusCode::NOT_IMPLEMENTED,
                    format!("Rust 迁移版暂不支持修改设置：{key}"),
                ));
            }
        }
    }
    Ok(ui.value.clone())
}

fn browse_extensions(mode: &str) -> &'static [&'static str] {
    match mode {
        "excel" => &[".xlsx", ".xlsm", ".xls"],
        "data" => &[".csv", ".xlsx"],
        "config" => &[".xlsx"],
        "excel_ooxml" => &[".xlsx", ".xlsm"],
        "central_data" => &[".csv", ".xlsx", ".xls"],
        "csv" => &[".csv"],
        "report_data" => &[".xlsx", ".xls"],
        _ => &[],
    }
}

fn refresh_period_pairs(state: &AppState) {
    let (current, previous, old_warning) = {
        let ui = state.lock();
        (
            ui.value["pcCurDir"]
                .as_str()
                .filter(|path| !path.is_empty())
                .map(PathBuf::from),
            ui.value["pcPreDir"]
                .as_str()
                .filter(|path| !path.is_empty())
                .map(PathBuf::from),
            ui.value["periodSourceWarning"]
                .as_str()
                .unwrap_or_default()
                .to_owned(),
        )
    };
    let warning = report_source_warning(current.as_deref(), previous.as_deref());
    let pairs = list_period_pairs(current.as_deref(), previous.as_deref());
    let mut ui = state.lock();
    ui.value["periodSourceWarning"] = json!(warning);
    match pairs {
        Ok(pairs) => {
            ui.value["periodPairs"] = serde_json::to_value(pairs).unwrap_or_else(|_| json!([]));
        }
        Err(error) => {
            ui.value["periodPairs"] = json!([]);
            push_log(&mut ui.value, format!("两期文件配对失败：{error:#}"));
        }
    }
    if !warning.is_empty() && warning != old_warning {
        push_log(&mut ui.value, warning);
    }
}

fn extension_allowed(path: &Path, extensions: &[&str]) -> bool {
    extensions.is_empty()
        || path
            .file_name()
            .and_then(|name| name.to_str())
            .map(|name| {
                let name = name.to_lowercase();
                extensions.iter().any(|extension| name.ends_with(extension))
            })
            .unwrap_or(false)
}

fn is_hidden_name(path: &Path) -> bool {
    path.file_name()
        .and_then(|name| name.to_str())
        .map(|name| name.starts_with('.') || name.starts_with("~$"))
        .unwrap_or(false)
}

fn browse_error(target: &Path, root: &Path, error: &str, selected_file: &str) -> Value {
    let parent = target
        .parent()
        .filter(|path| *path != target)
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_default();
    json!({
        "path": target.to_string_lossy(),
        "parent": parent,
        "dirs": [],
        "files": [],
        "root": root.to_string_lossy(),
        "selectedFile": selected_file,
        "error": error
    })
}

fn unquote_path(value: &str) -> String {
    let mut text = value.trim().to_owned();
    for _ in 0..3 {
        let pairs = [("\"", "\""), ("'", "'"), ("“", "”"), ("‘", "’")];
        let mut changed = false;
        for (open, close) in pairs {
            if text.starts_with(open)
                && text.ends_with(close)
                && text.len() >= open.len() + close.len()
            {
                text = text[open.len()..text.len() - close.len()].trim().to_owned();
                changed = true;
            }
        }
        if !changed {
            break;
        }
    }
    text
}

fn arg_string(args: &[Value], index: usize) -> Option<&str> {
    args.get(index).and_then(Value::as_str)
}

fn period_path_label(kind: &str) -> &'static str {
    match kind {
        "pcCurDir" => "本期目录",
        "pcPreDir" => "上期目录",
        "pcCentral" => "大集中数据",
        "pcConfig" => "报表采集系统配置",
        "pcOutput" => "输出目录",
        _ => "路径",
    }
}

fn push_detail_log(state: &mut Value, message: String) {
    let Some(log) = state.get_mut("log").and_then(Value::as_array_mut) else {
        return;
    };
    log.push(json!({"text": message, "detail": true}));
    if log.len() > 200 {
        let excess = log.len() - 200;
        log.drain(0..excess);
    }
}
fn push_log(state: &mut Value, message: String) {
    let Some(log) = state.get_mut("log").and_then(Value::as_array_mut) else {
        return;
    };
    log.push(json!({"text": message, "detail": false}));
    if log.len() > 200 {
        let excess = log.len() - 200;
        log.drain(0..excess);
    }
}

#[derive(Debug)]
struct ApiError {
    status: StatusCode,
    message: String,
}

impl ApiError {
    fn new(status: StatusCode, message: impl Into<String>) -> Self {
        Self {
            status,
            message: message.into(),
        }
    }

    fn internal(message: impl Into<String>) -> Self {
        Self::new(StatusCode::INTERNAL_SERVER_ERROR, message)
    }
}

impl IntoResponse for ApiError {
    fn into_response(self) -> Response {
        (self.status, Json(json!({"error": self.message}))).into_response()
    }
}

#[cfg(test)]
mod tests {
    use super::{AppState, browse_dir, dispatch_api, progress_value, render_index};
    use crate::central_csv::EXPECTED_HEADERS;
    use crate::s2_service::S2ProgressEvent;
    use excel_writer::Workbook;
    use serde_json::{Value, json};
    use std::{
        fs,
        path::{Path, PathBuf},
        time::{Duration, SystemTime, UNIX_EPOCH},
    };

    fn test_state() -> AppState {
        AppState::new(PathBuf::from("missing-test-config.xlsx"))
    }

    fn write_s1_region_source(path: &Path) {
        let mut workbook = Workbook::new();
        let data = workbook.add_worksheet();
        data.set_name("Data").expect("name S1 summary source sheet");
        data.write_string(1, 1, "财务")
            .expect("write S1 summary structure group header");
        data.write_string(2, 1, "单位")
            .expect("write S1 summary structure row header");
        data.write_string(2, 2, "金额")
            .expect("write S1 summary structure value header");
        data.write_string(4, 1, "甲")
            .expect("write S1 summary row label");
        data.write_number(4, 2, 10.0)
            .expect("write S1 summary row value");
        data.write_string(6, 1, "乙")
            .expect("write second S1 summary row label");
        data.write_number(6, 2, 20.0)
            .expect("write second S1 summary row value");
        data.write_string(0, 5, "测试分组")
            .expect("write S1 local metadata");
        let meta = workbook.add_worksheet();
        meta.set_name("Meta")
            .expect("name S1 summary metadata sheet");
        meta.write_string(0, 0, "测试批次")
            .expect("write S1 global metadata");
        workbook.save(path).expect("save S1 summary source");
    }

    fn write_s3_comparison(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet
            .set_name("比较结果")
            .expect("set comparison sheet name");
        let headers = [
            "机构类代码",
            "机构类名称",
            "地区代码",
            "地区名称",
            "指标代码",
            "指标名称",
            "数据值",
            "上期值",
            "输出说明",
            "说明内容",
            "增减额",
        ];
        for (column, header) in headers.iter().enumerate() {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write comparison header");
        }
        let row = [
            "12345",
            "示例机构",
            "4401",
            "示例地区",
            "I-01",
            "示例指标",
            "100",
            "90",
            "是",
            "需要说明",
            "10",
        ];
        for (column, value) in row.iter().enumerate() {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write selected comparison row");
        }
        workbook.save(path).expect("save comparison workbook");
    }

    fn write_s3_common_config(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet
            .set_name("机构地区参照")
            .expect("set organization reference sheet name");
        for (column, header) in ["机构类代码", "地区代码", "机构名称"].iter().enumerate()
        {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write organization reference header");
        }
        for (column, value) in ["12345", "4401", "配置机构名称"].iter().enumerate() {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write organization reference row");
        }
        workbook
            .save(path)
            .expect("save common configuration workbook");
    }

    fn write_s3_current_value_common_config(path: &Path) {
        let mut workbook = Workbook::new();
        let parameters = workbook.add_worksheet();
        parameters.set_name("运行参数").unwrap();
        for (column, header) in ["参数", "值"].iter().enumerate() {
            parameters.write_string(0, column as u16, *header).unwrap();
        }
        parameters.write_string(1, 0, "默认源数据单位").unwrap();
        parameters.write_string(1, 1, "元").unwrap();
        parameters.write_string(2, 0, "默认目标单位").unwrap();
        parameters.write_string(2, 1, "亿元").unwrap();

        let exceptions = workbook.add_worksheet();
        exceptions.set_name("单位换算例外").unwrap();
        exceptions.write_string(0, 0, "指标代码").unwrap();
        workbook.save(path).unwrap();
    }

    fn write_s3_current_value_cross_config(path: &Path) {
        let mut workbook = Workbook::new();
        let rules = workbook.add_worksheet();
        rules.set_name("跨期数值核对").unwrap();
        let headers = [
            "校验编码",
            "校验名称",
            "机构类代码",
            "地区代码",
            "前提条件",
            "left",
            "right",
            "校验公式",
            "校验提示",
            "取反标识",
            "无误是否提示",
            "是否逻辑校验",
            "禁用",
            "备注",
        ];
        for (column, header) in headers.iter().enumerate() {
            rules.write_string(0, column as u16, *header).unwrap();
        }
        let values = [
            "R1",
            "Web 集成测试规则",
            "^0001$",
            "1234567",
            "",
            "[I1,余额,人民币,月,2]",
            "0",
            "left > right",
            "请核实",
            "",
            "是",
            "",
            "",
            "synthetic web fixture",
        ];
        for (column, value) in values.iter().enumerate() {
            rules.write_string(1, column as u16, *value).unwrap();
        }
        workbook.save(path).unwrap();
    }

    fn write_s3_current_value_input_xlsx(path: &Path) {
        let mut workbook = Workbook::new();
        let cover = workbook.add_worksheet();
        cover.set_name("说明").unwrap();
        cover.write_string(0, 0, "非导数工作表").unwrap();

        let data = workbook.add_worksheet();
        data.set_name("本期导数").unwrap();
        for (column, header) in EXPECTED_HEADERS.iter().enumerate() {
            data.write_string(0, column as u16, *header).unwrap();
        }
        let values = [
            "人民币",
            "2026-08-31",
            "0001",
            "机构",
            "1234567",
            "地区",
            "1",
            "I1",
            "指标",
            "余额",
            "人民币",
            "月",
            "2",
        ];
        for (column, value) in values.iter().enumerate() {
            data.write_string(2, column as u16, *value).unwrap();
        }
        data.write_number(2, 13, 200_000_000.0).unwrap();
        workbook.save(path).unwrap();
    }

    fn write_s3_execution_configs(common_path: &Path, feature_path: &Path) {
        let mut common = Workbook::new();
        let parameters = common.add_worksheet();
        parameters.set_name("运行参数").unwrap();
        for (column, header) in ["参数", "值"].iter().enumerate() {
            parameters.write_string(0, column as u16, *header).unwrap();
        }
        parameters.write_string(1, 0, "默认源数据单位").unwrap();
        parameters.write_string(1, 1, "元").unwrap();
        parameters.write_string(2, 0, "默认目标单位").unwrap();
        parameters.write_string(2, 1, "亿元").unwrap();

        let exceptions = common.add_worksheet();
        exceptions.set_name("单位换算例外").unwrap();
        exceptions.write_string(0, 0, "指标代码").unwrap();
        exceptions.write_string(0, 1, "启用").unwrap();

        let alerts = common.add_worksheet();
        alerts.set_name("环比警戒").unwrap();
        for (column, header) in [
            "下限",
            "上限",
            "备注文字",
            "填充颜色",
            "是否整行填充",
            "金额变动阈值",
            "是否说明",
        ]
        .iter()
        .enumerate()
        {
            alerts.write_string(0, column as u16, *header).unwrap();
        }
        common.save(common_path).unwrap();

        let mut feature = Workbook::new();
        let indicators = feature.add_worksheet();
        indicators.set_name("指标参照").unwrap();
        for (column, header) in ["指标代码", "表单代码", "行序号", "指标别名"]
            .iter()
            .enumerate()
        {
            indicators.write_string(0, column as u16, *header).unwrap();
        }
        let actions = feature.add_worksheet();
        actions.set_name("规则动作").unwrap();
        for (column, header) in [
            "规则编号",
            "规则说明",
            "规则动作",
            "指标代码",
            "数据属性",
            "币种",
            "比较指标代码",
            "人民币阈值(亿元)",
            "美元阈值(亿美元)",
            "变幅阈值(%)",
            "来源分组",
            "适用表单",
            "适用频度",
            "适用批次",
            "适用场景",
            "详细说明",
            "启用",
            "停用原因",
            "备注",
        ]
        .iter()
        .enumerate()
        {
            actions.write_string(0, column as u16, *header).unwrap();
        }
        let expressions = feature.add_worksheet();
        expressions.set_name("表达式校验").unwrap();
        for (column, header) in [
            "规则编号",
            "规则说明",
            "校验表达式",
            "是否取反",
            "容差值(万元)",
            "启用",
            "停用原因",
        ]
        .iter()
        .enumerate()
        {
            expressions.write_string(0, column as u16, *header).unwrap();
        }
        feature.save(feature_path).unwrap();
    }

    fn write_s3_execution_csv(path: &Path, date: &str, amount: &str) {
        let row = format!(
            "业务一,{date},ORG1,示例机构,R001,示例地区,1,I001,指标一,余额,CNY,月,1,{amount}"
        );
        fs::write(path, format!("{}\n{row}\n", EXPECTED_HEADERS.join(","))).unwrap();
    }

    #[test]
    fn served_page_is_original_markup_plus_runtime_bridge() {
        let served = render_index();
        let served_markup = served.replace(super::BRIDGE_JS, "");
        assert!(super::ORIGINAL_HTML.contains("centralCrossCur:'csv'"));
        assert!(served_markup.contains("centralCrossCur:'central_data'"));
        assert!(!served_markup.contains("centralCrossCur:'csv'"));
        assert!(served_markup.contains("本期数值核对文件（CSV/XLSX/XLS）"));
        assert!(!served_markup.contains("XLS/XLSX 尚未接通"));
        assert!(served.contains("window.bridge = { api: api }"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_comparison()"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_explanation_export()"));
        assert!(super::ORIGINAL_HTML.contains("bridge.api.start_central_cross_period()"));
    }

    #[test]
    fn s3_current_value_check_runs_through_the_web_api() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_current_value_{}_{}",
            std::process::id(),
            suffix
        ));
        fs::create_dir_all(&root).expect("create S3-F04 temporary directory");
        let current = root.join("current.csv");
        let xlsx_current = root.join("current.xlsx");
        let common_config = root.join("3.0大集中通用配置.xlsx");
        let cross_config = root.join("3.2大集中本期数值核对_配置.xlsx");
        let output = root.join("output");
        fs::create_dir_all(&output).expect("create S3-F04 output directory");
        let csv = format!(
            "{}\n人民币,2026-08-31,0001,机构,1234567,地区,1,I1,指标,余额,人民币,月,2,200000000\n",
            EXPECTED_HEADERS.join(",")
        );
        fs::write(&current, &csv).expect("write synthetic current CSV");
        write_s3_current_value_input_xlsx(&xlsx_current);
        write_s3_current_value_common_config(&common_config);
        write_s3_current_value_cross_config(&cross_config);

        let state = test_state();
        let selected_xlsx = dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCrossCur"),
                json!(xlsx_current.to_string_lossy()),
            ],
        )
        .expect("XLSX input should be selectable for S3-F04");
        assert_eq!(
            selected_xlsx,
            json!(fs::canonicalize(&xlsx_current).unwrap().to_string_lossy())
        );

        for (field, path) in [
            ("centralCrossCur", &current),
            ("centralCommonConfig", &common_config),
            ("centralCrossConfig", &cross_config),
            ("centralOutput", &output),
        ] {
            dispatch_api(
                &state,
                "choose_central",
                &[json!(field), json!(path.to_string_lossy())],
            )
            .unwrap_or_else(|error| panic!("select {field}: {}", error.message));
        }
        dispatch_api(&state, "update", &[json!({"exportRunLogs": true})])
            .expect("enable run-log workbook for the integration check");
        assert_eq!(
            dispatch_api(&state, "start_central_cross_period", &[])
                .expect("start S3-F04 through its Web API"),
            json!(true)
        );

        for _ in 0..200 {
            if state.get_state()["busy"] != true {
                break;
            }
            std::thread::sleep(Duration::from_millis(25));
        }
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert_eq!(final_state["status"], "本期数值核对完成：1 行");
        let result_paths = fs::read_dir(&output)
            .unwrap()
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .filter(|path| {
                path.file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| name.starts_with("数值核对结果_"))
            })
            .collect::<Vec<_>>();
        assert_eq!(result_paths.len(), 1);
        let canonical_result = fs::canonicalize(&result_paths[0])
            .unwrap()
            .to_string_lossy()
            .into_owned();
        assert_eq!(
            final_state["centralCrossLastResult"].as_str(),
            Some(canonical_result.as_str())
        );
        let result_sheet = excel_reader::read_value_columns(&result_paths[0], 17)
            .unwrap()
            .remove(0);
        assert_eq!(result_sheet.max_row, 2);
        assert_eq!(
            result_sheet.values.get("O2"),
            Some(&excel_reader::CellValue::Text(
                "一边有，一边无;请核实".to_owned()
            ))
        );
        assert_eq!(
            fs::read_dir(&output)
                .unwrap()
                .filter_map(|entry| entry.ok().map(|entry| entry.path()))
                .filter(|path| {
                    path.file_name()
                        .and_then(|name| name.to_str())
                        .is_some_and(|name| name.starts_with("本期数值核对_运行日志_"))
                })
                .count(),
            1
        );
        let logs = final_state["log"].as_array().unwrap();
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"].as_str().is_some_and(|text| {
                    text.contains("reader=calamine/csv")
                        && text.contains("writer=rust_xlsxwriter")
                        && text.contains("工作簿写入=")
                })
        }));
        assert_eq!(fs::read(&current).unwrap(), csv.as_bytes());
        let xlsx_before = fs::read(&xlsx_current).unwrap();
        assert!(fs::read(&common_config).is_ok());
        assert!(fs::read(&cross_config).is_ok());

        dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCrossCur"),
                json!(xlsx_current.to_string_lossy()),
            ],
        )
        .expect("select the XLSX input for an end-to-end Web run");
        dispatch_api(&state, "start_central_cross_period", &[])
            .expect("start S3-F04 XLSX run through its Web API");
        for _ in 0..200 {
            if state.get_state()["busy"] != true {
                break;
            }
            std::thread::sleep(Duration::from_millis(25));
        }
        let xlsx_state = state.get_state();
        assert_eq!(xlsx_state["busy"], false);
        assert_eq!(xlsx_state["status"], "本期数值核对完成：1 行");
        assert!(xlsx_state["log"].as_array().unwrap().iter().any(|entry| {
            entry["detail"] == true
                && entry["text"].as_str().is_some_and(|text| {
                    text.contains("reader=calamine") && text.contains("结果行=1")
                })
        }));
        assert_eq!(fs::read(&xlsx_current).unwrap(), xlsx_before);
        fs::remove_dir_all(root).expect("remove S3-F04 temporary files");
    }

    #[test]
    fn s2_web_run_persists_stage_engine_and_elapsed_detail_logs() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s2_progress_{}_{}",
            std::process::id(),
            suffix
        ));
        let current = root.join("current");
        let previous = root.join("previous");
        let output = root.join("output");
        fs::create_dir_all(&current).expect("create S2 current directory");
        fs::create_dir_all(&previous).expect("create S2 previous directory");
        fs::create_dir_all(&output).expect("create S2 output directory");

        let fixture_root = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
        let source_report = fixture_root.join("s2_import/9144#2026-06-30#x#20202#甲银行.xlsx");
        fs::copy(
            &source_report,
            current.join("9144#2026-06-30#x#20202#甲银行.xlsx"),
        )
        .expect("copy current S2 fixture");
        fs::copy(
            &source_report,
            previous.join("9144#2026-03-31#x#20202#甲银行.xlsx"),
        )
        .expect("copy previous S2 fixture");
        let config = fixture_root.join("s2_config/check_valid.xlsx");
        let external = fixture_root.join("s2_external_import/plain_layout.xlsx");
        let state = AppState::new(config.clone());

        for (kind, path) in [
            ("pcCurDir", current.as_path()),
            ("pcPreDir", previous.as_path()),
            ("pcConfig", config.as_path()),
            ("pcCentral", external.as_path()),
            ("pcOutput", output.as_path()),
        ] {
            dispatch_api(
                &state,
                "choose_period_compare",
                &[json!(kind), json!(path.to_string_lossy())],
            )
            .expect("select S2 Web API input path");
        }
        assert_eq!(
            dispatch_api(&state, "start_period_compare_v2", &[]).expect("start S2 Web API run"),
            json!(true)
        );

        for _ in 0..400 {
            if state.get_state()["busy"] != true {
                break;
            }
            std::thread::sleep(Duration::from_millis(25));
        }
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert_eq!(final_state["periodProgress"]["status"], "completed");
        let report = PathBuf::from(final_state["periodSummary"]["result"].as_str().unwrap());
        assert!(report.is_file());

        let logs = final_state["log"].as_array().expect("S2 Web logs");
        for stage in [
            "config",
            "current",
            "previous",
            "validate",
            "period",
            "rules",
            "external_read",
            "external",
            "export",
            "save",
            "complete",
        ] {
            let engine = crate::s2_service::s2_engine_for_stage(stage).unwrap();
            assert!(
                logs.iter().any(|entry| {
                    entry["detail"] == true
                        && entry["text"].as_str().is_some_and(|text| {
                            text.contains(&format!("S2 阶段开始：stage={stage}"))
                                && text.contains(&format!("engine={engine}"))
                        })
                }),
                "missing start detail for S2 stage {stage}"
            );
            assert!(
                logs.iter().any(|entry| {
                    entry["detail"] == true
                        && entry["text"].as_str().is_some_and(|text| {
                            text.contains(&format!("S2 阶段完成：stage={stage}"))
                                && text.contains(&format!("engine={engine}"))
                                && text.contains("elapsed_ms=")
                        })
                }),
                "missing completion detail for S2 stage {stage}"
            );
        }
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("S2 任务完成：reader=Calamine"))
        }));
        fs::remove_dir_all(root).expect("remove temporary S2 Web API files");
    }

    #[test]
    fn s3_explanation_export_runs_through_the_existing_frontend_api() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_explanation_{}_{}",
            std::process::id(),
            suffix
        ));
        let comparison = root.join("comparison.xlsx");
        let common_config = root.join("3.0大集中通用配置.xlsx");
        let output = root.join("output");
        fs::create_dir_all(&output).expect("create S3 output directory");
        write_s3_comparison(&comparison);
        write_s3_common_config(&common_config);

        let state = test_state();
        let selected = dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCompareFile"),
                json!(comparison.to_string_lossy()),
            ],
        )
        .expect("select comparison file through the Web API");
        assert_eq!(
            selected,
            json!(fs::canonicalize(&comparison).unwrap().to_string_lossy())
        );
        dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCommonConfig"),
                json!(common_config.to_string_lossy()),
            ],
        )
        .expect("select common configuration through the Web API");
        dispatch_api(
            &state,
            "choose_central",
            &[json!("centralOutput"), json!(output.to_string_lossy())],
        )
        .expect("select output directory through the Web API");
        assert_eq!(
            dispatch_api(&state, "start_central_explanation_export", &[])
                .expect("start S3-F02 through the Web API"),
            json!(true)
        );

        for _ in 0..200 {
            if state.get_state()["busy"] != true {
                break;
            }
            std::thread::sleep(Duration::from_millis(25));
        }
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert_eq!(
            final_state["status"],
            "指标说明导出完成：1 条",
            "S3 Web API logs: {}",
            serde_json::to_string(&final_state["log"]).unwrap()
        );
        let output_path = PathBuf::from(final_state["centralExplainFile"].as_str().unwrap());
        assert!(output_path.is_file());
        let canonical_output = fs::canonicalize(&output).expect("canonical output directory");
        assert_eq!(output_path.parent(), Some(canonical_output.as_path()));
        assert!(
            output_path
                .file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("指标说明_")
        );
        let output_sheets = excel_reader::read_value_columns(&output_path, 8)
            .expect("read generated explanation workbook");
        let explanation = output_sheets
            .iter()
            .find(|sheet| sheet.name == "说明文件")
            .expect("explanation worksheet");
        assert_eq!(
            explanation.values.get("B3"),
            Some(&excel_reader::CellValue::Text(
                "配置机构名称（示例地区）".to_owned()
            )),
            "the Web entry must apply the selected 3.0 common organization mapping"
        );
        let logs = final_state["log"].as_array().unwrap();
        assert!(logs.iter().any(|entry| {
            entry["detail"] == false
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("导出完成：1 条"))
        }));
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"].as_str().is_some_and(|text| {
                    text.contains("reader=calamine")
                        && text.contains("writer=rust_xlsxwriter")
                        && text.contains("配置读取=")
                })
        }));
        fs::remove_dir_all(root).expect("remove temporary S3 Web API files");
    }

    #[test]
    fn s3_explanation_export_requires_the_python_compatible_common_config() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_missing_config_{}_{}",
            std::process::id(),
            suffix
        ));
        fs::create_dir_all(&root).expect("create S3 input directory");
        let comparison = root.join("comparison.xlsx");
        write_s3_comparison(&comparison);

        let state = test_state();
        dispatch_api(
            &state,
            "choose_central",
            &[
                json!("centralCompareFile"),
                json!(comparison.to_string_lossy()),
            ],
        )
        .expect("select comparison file through the Web API");
        let error = dispatch_api(&state, "start_central_explanation_export", &[])
            .expect_err("export must require the common config used by Python");

        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
        assert!(error.message.contains("大集中通用配置"));
        assert_eq!(state.get_state()["busy"], false);
        fs::remove_dir_all(root).expect("remove temporary S3 missing-config files");
    }

    #[test]
    fn state_reports_s2_paths_and_unmigrated_status() {
        let state = test_state();
        let value = state.get_state();
        assert_eq!(value["distributionProfile"], "s1s2");
        assert!(value["pcConfig"].is_string());
        assert_eq!(value["filePickerMode"], "浏览器内置");
        assert_eq!(value["busy"], false);
        assert_eq!(value["template"], "");
    }

    #[test]
    fn s2_progress_view_includes_stage_engine_and_elapsed_time() {
        let event = S2ProgressEvent {
            stage: "current".to_owned(),
            label: "读取本期报表".to_owned(),
            percent: 12.0,
            current: 1,
            total: 2,
            message: "读取工作簿".to_owned(),
        };
        let value = progress_value(&event, "running", 1.5, 0.25);

        assert_eq!(value["engine"], "Calamine");
        assert_eq!(value["elapsedSeconds"], 1.5);
        assert_eq!(value["stageElapsedSeconds"], 0.25);
        assert!(value["label"].as_str().unwrap().contains("Calamine"));
        assert!(value["label"].as_str().unwrap().contains("本阶段已运行"));
    }

    #[test]
    fn s1_template_inspection_is_reachable_through_the_existing_ui_api() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "audit_web_template_inspection_{}_{}.xlsx",
            std::process::id(),
            suffix
        ));
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("set sheet");
        sheet
            .write_formula(0, 0, "=1+#REF!")
            .expect("write formula");
        workbook
            .define_name("校验区域_1", "='检查表'!$A$1")
            .expect("define named range");
        workbook.save(&path).expect("save workbook");

        let state = test_state();
        let chosen = dispatch_api(
            &state,
            "choose_file",
            &[json!("template"), json!(path.to_string_lossy())],
        )
        .expect("choose template");
        assert_eq!(chosen, json!(path.to_string_lossy()));
        let report =
            dispatch_api(&state, "inspect_template_formulas", &[]).expect("inspect template");
        assert_eq!(report["template"], path.to_string_lossy().as_ref());
        assert_eq!(report["formulaCount"], 1);
        assert_eq!(report["errors"][0]["kind"], "公式");
        assert_eq!(report["errors"][0]["token"], "#REF!");
        assert_eq!(report["namedRangeGroups"][0]["count"], 1);
        let logs = state.get_state()["log"].as_array().unwrap().clone();
        assert!(logs.iter().any(|entry| entry["detail"] == false));
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("calamine+namespace_aware_ooxml"))
        }));
        fs::remove_file(path).expect("remove workbook");
    }

    #[test]
    fn s1_template_inspection_respects_busy_state() {
        let state = test_state();
        state.lock().value["busy"] = json!(true);
        let error = dispatch_api(&state, "inspect_template_formulas", &[]).unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::CONFLICT);
    }

    #[test]
    fn s2_config_check_logs_user_result_and_developer_reader_timing() {
        let config = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("s2_config")
            .join("check_valid.xlsx");
        let state = AppState::new(config);

        let report = dispatch_api(&state, "check_period_audit_config", &[])
            .expect("check S2 config through the Web API");
        assert_eq!(report["passed"], true);

        let logs = state.get_state()["log"].as_array().unwrap().clone();
        assert!(logs.iter().any(|entry| {
            entry["detail"] == false
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("S2 配置检查通过"))
        }));
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"].as_str().is_some_and(|text| {
                    text.contains("S2-F01 配置检查步骤")
                        && text.contains("reader=calamine")
                        && text.contains("指标=")
                        && text.contains("elapsed_ms=")
                })
        }));
    }

    #[test]
    fn path_browser_returns_the_frontend_contract() {
        let state = test_state();
        let value = browse_dir(&state, &[json!(""), json!("report_data")]).unwrap();
        assert!(value["path"].is_string());
        assert!(value["parent"].is_string());
        assert!(value["dirs"].is_array());
        assert!(value["files"].is_array());
        assert!(value["selectedFile"].is_string());
        assert!(value["error"].is_string());
    }

    #[test]
    fn s3_execution_compare_reports_missing_gui_inputs() {
        let error = dispatch_api(&test_state(), "start_central_comparison", &[]).unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
        assert!(error.message.contains("本期大集中数据"));
    }

    #[test]
    fn s3_execution_compare_runs_through_the_existing_gui_api() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_web_s3_execution_compare_{}_{}",
            std::process::id(),
            suffix
        ));
        fs::create_dir_all(&root).expect("create S3-F01 web test directory");
        let current = root.join("current.csv");
        let previous = root.join("previous.csv");
        let common = root.join("3.0大集中通用配置.xlsx");
        let feature = root.join("3.1大集中执行比较_配置.xlsx");
        write_s3_execution_csv(&current, "2026-08-31", "1000");
        write_s3_execution_csv(&previous, "2026-07-31", "900");
        write_s3_execution_configs(&common, &feature);
        let current_before = fs::read(&current).unwrap();
        let previous_before = fs::read(&previous).unwrap();
        let common_before = fs::read(&common).unwrap();
        let feature_before = fs::read(&feature).unwrap();

        let state = test_state();
        for (field, path) in [
            ("centralCur", &current),
            ("centralPre", &previous),
            ("centralCommonConfig", &common),
            ("centralComparisonConfig", &feature),
        ] {
            dispatch_api(
                &state,
                "choose_central",
                &[json!(field), json!(path.to_string_lossy())],
            )
            .unwrap_or_else(|error| panic!("select {field}: {}", error.message));
        }
        dispatch_api(
            &state,
            "update",
            &[json!({"centralExpressionSchema": "FIVE_SEGMENT_V1"})],
        )
        .expect("set the legacy GUI expression schema");
        assert_eq!(
            dispatch_api(&state, "start_central_comparison", &[]).unwrap(),
            json!(true)
        );

        let started = std::time::Instant::now();
        while state.get_state()["busy"] == true && started.elapsed() < Duration::from_secs(10) {
            std::thread::sleep(Duration::from_millis(10));
        }
        let value = state.get_state();
        assert_eq!(value["busy"], false, "background GUI task should finish");
        assert_eq!(value["status"], "执行比较完成：1 行");
        let output = PathBuf::from(
            value["centralCompareFile"]
                .as_str()
                .expect("GUI should publish comparison output path"),
        );
        assert!(
            output.is_file(),
            "GUI task should create the result workbook"
        );
        assert_eq!(
            value["centralFormOutput"],
            output.parent().unwrap().to_string_lossy().as_ref()
        );
        let logs = value["log"].as_array().expect("GUI log array");
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("writer=rust_xlsxwriter"))
        }));
        assert!(logs.iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("S3-F01 完成：reader=calamine/csv"))
        }));
        assert_eq!(fs::read(&current).unwrap(), current_before);
        assert_eq!(fs::read(&previous).unwrap(), previous_before);
        assert_eq!(fs::read(&common).unwrap(), common_before);
        assert_eq!(fs::read(&feature).unwrap(), feature_before);
        fs::remove_dir_all(&root).expect("remove synthetic S3-F01 web test directory");
    }

    #[test]
    fn s1_f2_runs_through_the_existing_gui_workflow_without_modifying_inputs() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-s1-f2-{suffix}"));
        let input_dir = root.join("input");
        let template_dir = root.join("templates");
        let config_dir = root.join("config");
        fs::create_dir_all(&input_dir).expect("create S1-F02 input directory");
        fs::create_dir_all(&template_dir).expect("create S1-F02 template directory");
        fs::create_dir_all(&config_dir).expect("create S1-F02 config directory");

        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/s1_region_summary_template.xlsx");
        let template = template_dir.join("summary_template.xlsx");
        let history = config_dir.join("逐笔统计系统_配置.xlsx");
        let source = input_dir.join("机构甲_20260731.xlsx");
        fs::copy(&fixture, &template).expect("copy template fixture");
        fs::copy(&fixture, &history).expect("copy history config fixture");
        write_s1_region_source(&source);
        let source_before = fs::read(&source).expect("snapshot S1 source");
        let history_before = fs::read(&history).expect("snapshot history config");

        let state = test_state();
        dispatch_api(
            &state,
            "choose_folder",
            &[json!("input"), json!(input_dir.to_string_lossy())],
        )
        .expect("choose S1-F02 input directory");
        dispatch_api(
            &state,
            "choose_folder",
            &[json!("templateDir"), json!(template_dir.to_string_lossy())],
        )
        .expect("choose S1-F02 template directory");
        dispatch_api(
            &state,
            "update",
            &[json!({
                "input": input_dir.to_string_lossy(),
                "templateDir": template_dir.to_string_lossy()
            })],
        )
        .expect("update selected S1-F02 paths");
        dispatch_api(&state, "recognize", &[]).expect("scan S1-F02 source workbooks");
        assert_eq!(
            state.get_state()["sourceFiles"].as_array().unwrap().len(),
            1
        );
        assert_eq!(
            state.get_state()["selectedFiles"].as_array().unwrap().len(),
            1
        );
        dispatch_api(
            &state,
            "choose_file",
            &[json!("template"), json!(template.to_string_lossy())],
        )
        .expect("select S1-F02 template");
        dispatch_api(
            &state,
            "choose_history_config",
            &[json!(history.to_string_lossy())],
        )
        .expect("select S1 history config");

        let selected = state.get_state()["selectedFiles"].clone();
        dispatch_api(
            &state,
            "start_dag_flow",
            &[
                json!("dag:汇总校验结果说明"),
                json!({
                    "input": input_dir.to_string_lossy(),
                    "templateDir": template_dir.to_string_lossy(),
                    "template": template.to_string_lossy(),
                    "output": ""
                }),
                selected,
            ],
        )
        .expect("start S1-F02 through the original GUI API");

        let started = std::time::Instant::now();
        while state.get_state()["busy"] == true && started.elapsed() < Duration::from_secs(10) {
            std::thread::sleep(Duration::from_millis(10));
        }
        let final_state = state.get_state();
        assert_eq!(final_state["busy"], false);
        assert!(
            final_state["status"]
                .as_str()
                .unwrap_or_default()
                .starts_with("S1-F02 汇总完成"),
            "{}",
            final_state["status"]
        );
        let output = PathBuf::from(
            final_state["auditSummaryOutput"]
                .as_str()
                .expect("S1-F02 output path"),
        );
        assert!(output.is_file());
        assert_eq!(
            output.parent().map(fs::canonicalize).transpose().unwrap(),
            Some(fs::canonicalize(input_dir.join("执行结果")).unwrap())
        );
        assert_eq!(
            super::s1_region_summary_source_workbooks(&input_dir, true)
                .expect("rediscover S1 sources after output")
                .len(),
            1,
            "generated execution output must not become a later source file"
        );
        assert!(
            !excel_reader::read_workbook(&output)
                .expect("read generated S1-F02 result")
                .sheets
                .is_empty()
        );
        assert_eq!(
            fs::read(&source).expect("read source after S1-F02"),
            source_before
        );
        assert_eq!(
            fs::read(&history).expect("read history after S1-F02"),
            history_before
        );
        assert!(final_state["log"].as_array().unwrap().iter().any(|entry| {
            entry["detail"] == true
                && entry["text"]
                    .as_str()
                    .is_some_and(|text| text.contains("reader=calamine+namespace_aware_ooxml"))
        }));
        fs::remove_dir_all(&root).expect("remove synthetic S1-F02 web test directory");
    }

    #[test]
    fn s1_f2_gui_workflow_rejects_missing_history_configuration() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-s1-f2-no-history-{suffix}"));
        let input_dir = root.join("input");
        fs::create_dir_all(&input_dir).expect("create S1-F02 input directory");
        let template = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/s1_region_summary_template.xlsx");
        write_s1_region_source(&input_dir.join("source.xlsx"));
        let state = test_state();
        let error = dispatch_api(
            &state,
            "start_dag_flow",
            &[
                json!("dag:汇总校验结果说明"),
                json!({
                    "input": input_dir.to_string_lossy(),
                    "template": template.to_string_lossy()
                }),
                json!([]),
            ],
        )
        .unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
        assert!(error.message.contains("逐笔统计系统配置"));
        assert_eq!(state.get_state()["busy"], false);
        fs::remove_dir_all(&root).expect("remove S1-F02 missing config test directory");
    }

    #[test]
    fn invalid_path_kind_is_rejected() {
        let error = dispatch_api(
            &test_state(),
            "choose_period_compare",
            &[
                Value::String("unknown".to_owned()),
                Value::String("/tmp".to_owned()),
            ],
        )
        .unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::BAD_REQUEST);
    }

    #[test]
    fn choosing_and_clearing_period_directories_refreshes_pair_preview() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-pairs-{suffix}"));
        let current = root.join("current");
        let previous = root.join("previous");
        fs::create_dir_all(&current).expect("create current directory");
        fs::create_dir_all(&previous).expect("create previous directory");
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/s2_import/9144#2026-06-30#x#20202#甲银行.xlsx");
        fs::copy(
            &fixture,
            current.join("9144#2026-06-30#x#20202#甲银行.xlsx"),
        )
        .expect("copy current workbook");
        fs::copy(
            &fixture,
            previous.join("9144#2026-03-31#x#20202#甲银行.xlsx"),
        )
        .expect("copy previous workbook");

        let state = test_state();
        dispatch_api(
            &state,
            "choose_period_compare",
            &[json!("pcCurDir"), json!(current.to_string_lossy())],
        )
        .expect("select current directory");
        let current_only = state.get_state();
        assert_eq!(current_only["periodPairs"].as_array().unwrap().len(), 2);
        assert!(
            current_only["periodPairs"]
                .as_array()
                .unwrap()
                .iter()
                .all(|pair| pair["side"] == "cur")
        );

        dispatch_api(
            &state,
            "choose_period_compare",
            &[json!("pcPreDir"), json!(previous.to_string_lossy())],
        )
        .expect("select previous directory");
        let matched = state.get_state();
        assert!(
            matched["periodPairs"]
                .as_array()
                .unwrap()
                .iter()
                .all(|pair| pair["matched"] == true)
        );
        assert_eq!(matched["periodPairs"][0]["curDate"], "2026-06-30");
        assert_eq!(matched["periodPairs"][0]["preDate"], "2026-03-31");

        dispatch_api(&state, "clear_path", &[json!("pcPreDir")]).expect("clear previous");
        let cleared = state.get_state();
        assert!(
            cleared["periodPairs"]
                .as_array()
                .unwrap()
                .iter()
                .all(|pair| pair["side"] == "cur")
        );
        let _ = fs::remove_dir_all(root);
    }
}
