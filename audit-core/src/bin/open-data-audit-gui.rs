use anyhow::Result;
use audit_core::web_bridge::{launch_web_ui, serve_web_ui_headless};
use clap::Parser;
use std::path::PathBuf;

#[derive(Debug, Parser)]
#[command(
    name = "open-data-audit-gui",
    version,
    about = "Open the migrated Audit interface in the system browser"
)]
struct Cli {
    /// Local HTTP port. The server only listens on 127.0.0.1.
    #[arg(long, default_value_t = 8750)]
    port: u16,

    /// Audit S2 configuration workbook.
    #[arg(long)]
    config: Option<PathBuf>,

    /// Keep the HTTP bridge available without opening a browser window (useful for headless checks).
    #[arg(long, default_value_t = false)]
    no_browser: bool,
}

fn main() {
    if let Err(error) = run() {
        eprintln!("open-data-audit-gui: {error:#}");
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let cli = Cli::parse();
    let config_path = cli.config.unwrap_or_else(default_s2_config_path);
    if cli.no_browser {
        serve_web_ui_headless(cli.port, config_path)
    } else {
        launch_web_ui(cli.port, config_path)
    }
}

fn default_s2_config_path() -> PathBuf {
    if let Some(path) = std::env::var_os("OPEN_DATA_AUDIT_S2_CONFIG") {
        return PathBuf::from(path);
    }

    let current_directory = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    let candidates = [
        current_directory.join("audit-core/config-real/2.报表采集系统_配置.xlsx"),
        current_directory.join("audit-core/config/默认配置/2.报表采集系统_配置.xlsx"),
        current_directory.join("config-real/2.报表采集系统_配置.xlsx"),
        current_directory.join("config/默认配置/2.报表采集系统_配置.xlsx"),
    ];
    candidates
        .into_iter()
        .find(|path| path.is_file())
        .unwrap_or_else(|| PathBuf::from("2.报表采集系统_配置.xlsx"))
}
