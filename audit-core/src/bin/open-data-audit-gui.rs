use anyhow::Result;
use audit_core::{s2_config::resolve_default_config_path, web_bridge::serve_web_ui};
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
}

fn main() {
    if let Err(error) = run() {
        eprintln!("open-data-audit-gui: {error:#}");
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let cli = Cli::parse();
    let config_path = match cli.config {
        Some(path) => path,
        None => resolve_default_config_path()?,
    };
    serve_web_ui(cli.port, config_path)
}
