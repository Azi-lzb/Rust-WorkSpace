//! Local HTTP bridge for the bundled, unchanged Audit web frontend.
//!
//! This first bridge exposes only the S2 paths already implemented in Rust.
//! Other frontend APIs fail explicitly until their business workflows migrate.

use crate::{
    period_pairs::{list_period_pairs, report_source_warning},
    s2_config_check::check_s2_config,
    s2_service::{run_s2_audit, S2ProgressEvent, S2RunRequest},
};
use anyhow::{Context, Result};
use axum::{
    body::Bytes,
    extract::{Path as RoutePath, State},
    http::{header, HeaderMap, HeaderValue, StatusCode},
    response::{Html, IntoResponse, Response},
    routing::{get, post},
    Json, Router,
};
use serde_json::{json, Value};
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
const APP_TITLE: &str = "Open-Data-Audit (Rust)";

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
          var message = (data && data.error) || ('HTTP ' + response.status);
          if (response.status === 501 && method === 'start_dag_flow') {
            if (typeof window.showToast === 'function') {
              window.showToast('该功能尚未接入 Rust 版本。');
            } else {
              window.alert('该功能尚未接入 Rust 版本。');
            }
          }
          throw new Error(message);
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
        let config_report = check_s2_config(&config_path);
        let mut state = json!({
            "appInfo": {"title": APP_TITLE},
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
            "engines": [{"value": "自动", "label": "Rust 迁移功能"}]
        });
        let s2_state = json!({
            "input": "",
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
        state
            .as_object_mut()
            .expect("UI state is a JSON object")
            .extend(
                s2_state
                    .as_object()
                    .expect("S2 state is a JSON object")
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
    serve_web_ui_inner(port, config_path, true)
}

/// Launch the existing Audit interface in the user's system browser, then serve its Rust bridge.
pub fn launch_web_ui(port: u16, config_path: PathBuf) -> Result<()> {
    serve_web_ui(port, config_path)
}

/// Run the local-only frontend/API service without opening a browser window.
pub fn serve_web_ui_headless(port: u16, config_path: PathBuf) -> Result<()> {
    serve_web_ui_inner(port, config_path, false)
}

fn serve_web_ui_inner(port: u16, config_path: PathBuf, open_browser: bool) -> Result<()> {
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
        if open_browser {
            if let Err(error) = open_default_browser(&url) {
                eprintln!("无法自动打开浏览器，请手动访问 {url}：{error}");
            }
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
    ORIGINAL_HTML
        .replace("__AUDIT_APP_TITLE__", APP_TITLE)
        .replace("__AUDIT_ICON_URL__", "/assets/app-icon.png")
        .replacen("</head>", &format!("{BRIDGE_JS}</head>"), 1)
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
            return ApiError::new(StatusCode::BAD_REQUEST, "请求体必须是 JSON 数组").into_response()
        }
        Err(error) => {
            return ApiError::new(
                StatusCode::BAD_REQUEST,
                format!("请求体不是有效的 JSON：{error}"),
            )
            .into_response()
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
        "clear_path" => clear_period_path(state, args),
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

fn check_period_audit_config(state: &AppState) -> std::result::Result<Value, ApiError> {
    let config_path = state.lock().value["pcConfig"]
        .as_str()
        .unwrap_or_default()
        .to_owned();
    let report = check_s2_config(config_path);
    let value = serde_json::to_value(report)
        .map_err(|error| ApiError::internal(format!("序列化配置检查结果失败：{error}")))?;
    state.lock().value["periodConfigCheck"] = value.clone();
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
            ))
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
    let mut on_progress = |event: S2ProgressEvent| {
        let mut ui = state.lock();
        ui.value["status"] = json!(format!("{}……", event.label));
        ui.value["periodProgress"] =
            progress_value(&event, "running", worker_started.elapsed().as_secs_f64());
    };
    let result = run_s2_audit(&request, Some(&mut on_progress));
    let mut ui = state.lock();
    ui.value["busy"] = json!(false);
    ui.started_at = None;
    match result {
        Ok(result) => {
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
            push_log(&mut ui.value, format!("审核完成：{}", report_path));
        }
        Err(error) => {
            let message = format!("跨期比较失败：{error:#}");
            ui.value["periodProgress"] = json!({
                "status": "failed", "stage": "failed", "label": "审核失败",
                "percent": 0, "current": 0, "total": 0,
                "message": message,
                "elapsedSeconds": worker_started.elapsed().as_secs_f64()
            });
            ui.value["status"] = json!("跨期比较失败");
            push_log(&mut ui.value, message);
        }
    }
}

fn progress_value(event: &S2ProgressEvent, status: &str, elapsed: f64) -> Value {
    json!({
        "status": status,
        "stage": event.stage,
        "label": event.label,
        "percent": event.percent.clamp(0.0, 99.0),
        "current": event.current,
        "total": event.total,
        "message": event.message,
        "elapsedSeconds": elapsed
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
            "uiTheme" | "uiAccent" => {
                if !value.is_string() {
                    return Err(ApiError::new(
                        StatusCode::BAD_REQUEST,
                        format!("{key} 必须是字符串"),
                    ));
                }
                ui.value[key] = value.clone();
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
            // The unchanged legacy frontend syncs its shared form before S2 actions.
            // Empty paths belong to unmigrated workflows and do not change S2 state.
            "input" | "templateDir" | "template" | "external" | "output"
                if value.as_str() == Some("") => {}
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
    use super::{browse_dir, dispatch_api, render_index, update_state, AppState};
    use serde_json::{json, Value};
    use std::{
        fs,
        path::{Path, PathBuf},
        thread,
        time::{Duration, Instant, SystemTime, UNIX_EPOCH},
    };

    fn test_state() -> AppState {
        AppState::new(PathBuf::from("missing-test-config.xlsx"))
    }

    #[test]
    fn served_page_is_original_markup_plus_runtime_bridge() {
        let served = render_index();
        let expected = super::ORIGINAL_HTML
            .replace("__AUDIT_APP_TITLE__", super::APP_TITLE)
            .replace("__AUDIT_ICON_URL__", "/assets/app-icon.png");
        assert_eq!(served.replace(super::BRIDGE_JS, ""), expected);
        assert!(served.contains("window.bridge = { api: api }"));
        assert!(served.contains("该功能尚未接入 Rust 版本。"));
        assert!(!served.contains("__AUDIT_APP_TITLE__"));
        assert!(!served.contains("__AUDIT_ICON_URL__"));
    }

    #[test]
    fn state_reports_s2_paths_and_unmigrated_status() {
        let state = test_state();
        let value = state.get_state();
        assert_eq!(value["distributionProfile"], "s1s2");
        assert!(value["pcConfig"].is_string());
        assert_eq!(value["filePickerMode"], "浏览器内置");
        assert_eq!(value["busy"], false);
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
    fn unsupported_frontend_methods_fail_explicitly() {
        let error = dispatch_api(&test_state(), "start_dag_flow", &[]).unwrap_err();
        assert_eq!(error.status, axum::http::StatusCode::NOT_IMPLEMENTED);
        assert!(error.message.contains("尚未实现接口"));
    }

    #[test]
    fn empty_legacy_workflow_paths_do_not_block_s2_settings_sync() {
        let state = test_state();
        let updated = update_state(
            &state,
            &[json!({
                "input": "",
                "templateDir": "",
                "template": "",
                "external": "",
                "output": "",
                "uiAccent": "blue"
            })],
        )
        .expect("empty legacy paths should not block the S2 frontend sync");

        assert_eq!(updated["uiAccent"], "blue");
        assert_eq!(updated["pcCurDir"], state.get_state()["pcCurDir"]);
    }

    #[test]
    fn nonempty_legacy_workflow_path_is_still_explicitly_unsupported() {
        let error = update_state(
            &test_state(),
            &[json!({"external": "C:/not-migrated.xlsx"})],
        )
        .unwrap_err();

        assert_eq!(error.status, axum::http::StatusCode::NOT_IMPLEMENTED);
        assert!(error.message.contains("暂不支持修改设置：external"));
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
        assert!(current_only["periodPairs"]
            .as_array()
            .unwrap()
            .iter()
            .all(|pair| pair["side"] == "cur"));

        dispatch_api(
            &state,
            "choose_period_compare",
            &[json!("pcPreDir"), json!(previous.to_string_lossy())],
        )
        .expect("select previous directory");
        let matched = state.get_state();
        assert!(matched["periodPairs"]
            .as_array()
            .unwrap()
            .iter()
            .all(|pair| pair["matched"] == true));
        assert_eq!(matched["periodPairs"][0]["curDate"], "2026-06-30");
        assert_eq!(matched["periodPairs"][0]["preDate"], "2026-03-31");

        dispatch_api(&state, "clear_path", &[json!("pcPreDir")]).expect("clear previous");
        let cleared = state.get_state();
        assert!(cleared["periodPairs"]
            .as_array()
            .unwrap()
            .iter()
            .all(|pair| pair["side"] == "cur"));
        let _ = fs::remove_dir_all(root);
    }

    #[test]
    fn s2_api_route_runs_rust_service_and_publishes_report() {
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!("audit-web-s2-run-{suffix}"));
        let current = root.join("current");
        let previous = root.join("previous");
        let output = root.join("output");
        fs::create_dir_all(&current).expect("create current directory");
        fs::create_dir_all(&previous).expect("create previous directory");
        fs::create_dir_all(&output).expect("create output directory");

        let fixtures = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
        let source_report = fixtures.join("s2_import/9144#2026-06-30#x#20202#甲银行.xlsx");
        let current_report = current.join("9144#2026-06-30#x#20202#甲银行.xlsx");
        let previous_report = previous.join("9144#2026-03-31#x#20202#甲银行.xlsx");
        fs::copy(&source_report, &current_report).expect("copy current report fixture");
        fs::copy(&source_report, &previous_report).expect("copy previous report fixture");
        let current_before = fs::read(&current_report).expect("read current input before run");
        let previous_before = fs::read(&previous_report).expect("read previous input before run");
        let config = fixtures.join("s2_config/check_valid.xlsx");

        let state = AppState::new(config.clone());
        for (kind, path) in [
            ("pcCurDir", current.as_path()),
            ("pcPreDir", previous.as_path()),
            ("pcOutput", output.as_path()),
            ("pcConfig", config.as_path()),
        ] {
            dispatch_api(
                &state,
                "choose_period_compare",
                &[json!(kind), json!(path.to_string_lossy())],
            )
            .expect("select S2 input or output path through GUI API");
        }

        assert_eq!(
            dispatch_api(&state, "start_period_compare_v2", &[])
                .expect("start Rust S2 service through GUI API"),
            json!(true)
        );

        let deadline = Instant::now() + Duration::from_secs(20);
        loop {
            if state.get_state()["busy"] == false {
                break;
            }
            assert!(Instant::now() < deadline, "S2 GUI API call did not finish");
            thread::sleep(Duration::from_millis(10));
        }

        let finished = state.get_state();
        assert_eq!(finished["periodProgress"]["status"], "completed");
        let report_path = PathBuf::from(
            finished["periodSummary"]["result"]
                .as_str()
                .expect("GUI state contains result report path"),
        );
        assert!(
            report_path.is_file(),
            "S2 report should exist at {report_path:?}"
        );
        assert_eq!(
            fs::read(&current_report).expect("read current input after run"),
            current_before
        );
        assert_eq!(
            fs::read(&previous_report).expect("read previous input after run"),
            previous_before
        );

        fs::remove_dir_all(&root).expect("remove temporary S2 GUI fixture");
    }
}
