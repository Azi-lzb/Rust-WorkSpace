use anyhow::Result;
use audit_core::{
    conditional_format::evaluate_xlsx_workbook,
    discovery::ScanDepth,
    s1_template::recommend_template,
    s2_service::{run_s2_audit, S2ProgressEvent, S2RunRequest},
    s2_summary::S2AuditSummary,
    web_bridge::serve_web_ui_with_fallback,
};
use clap::{Args, Parser, Subcommand};
use serde::Serialize;
use std::path::{Path, PathBuf};

#[derive(Debug, Parser)]
#[command(
    name = "open-data-audit",
    version,
    about = "Rust implementation of Open-Data-Audit workflows"
)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Run the S2 period-report audit and create its three-sheet result workbook.
    S2(S2Args),
    /// Recommend an S1 template from source and template workbook sheet names.
    TemplateSuggest(TemplateSuggestArgs),
    /// Evaluate the Rust-supported static conditional-format rules in an XLSX workbook.
    ConditionalEvaluate(ConditionalEvaluateArgs),
    /// Serve the bundled Audit frontend through the Rust API bridge on localhost.
    Serve(ServeArgs),
}

#[derive(Debug, Args)]
struct ConditionalEvaluateArgs {
    /// Input .xlsx/.xlsm workbook. Formula cells are read from stored caches; this command does not recalculate.
    #[arg(long)]
    workbook: PathBuf,

    /// Optional scan range in Sheet!A1:B20 form. Repeat for multiple ranges. Defaults to rule AppliesTo ranges.
    #[arg(long = "scan-area")]
    scan_areas: Vec<String>,
}

#[derive(Debug, Args)]
struct TemplateSuggestArgs {
    /// Directory containing .xlsx templates.
    #[arg(long)]
    template_dir: PathBuf,

    /// Directory containing source .xlsx workbooks.
    #[arg(long)]
    input_dir: PathBuf,

    /// Include files below the input directory.
    #[arg(long, default_value_t = false)]
    recursive: bool,
}

#[derive(Debug, Args)]
struct ServeArgs {
    /// Local HTTP port. The server always binds to 127.0.0.1.
    #[arg(long, default_value_t = 8750)]
    port: u16,

    /// S2 configuration workbook. Defaults to OPEN_DATA_AUDIT_S2_CONFIG or the workspace placeholder.
    #[arg(long)]
    config: Option<PathBuf>,

    /// Enabled only by the no-argument desktop entry.
    #[arg(skip)]
    fallback_if_port_busy: bool,
}

#[derive(Debug, Args)]
struct S2Args {
    /// Directory containing current-period report workbooks.
    #[arg(long)]
    current_dir: PathBuf,

    /// Directory containing previous-period report workbooks.
    #[arg(long)]
    previous_dir: PathBuf,

    /// S2 comparison configuration workbook.
    #[arg(long)]
    config: PathBuf,

    /// Directory in which to create the result workbook.
    #[arg(long)]
    output_dir: PathBuf,

    /// Optional external workbook used by the configured cross-check rules.
    #[arg(long)]
    external: Option<PathBuf>,
}

#[derive(Debug, Serialize)]
struct S2CliOutput {
    status: &'static str,
    report_path: PathBuf,
    summary: S2AuditSummary,
}

fn main() {
    if let Err(error) = run() {
        eprintln!("open-data-audit: {error:#}");
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let command = if std::env::args_os().len() == 1 {
        default_ui_command()
    } else {
        Cli::parse().command
    };
    match command {
        Command::Serve(arguments) => {
            let config_path = arguments.config.unwrap_or_else(default_s2_config_path);
            serve_web_ui_with_fallback(
                arguments.port,
                config_path,
                arguments.fallback_if_port_busy,
            )?;
        }
        Command::S2(arguments) => {
            let request = S2RunRequest {
                current_directory: arguments.current_dir,
                previous_directory: arguments.previous_dir,
                config_path: arguments.config,
                output_directory: arguments.output_dir,
                external_workbook: arguments.external,
            };
            let mut on_progress = |event: S2ProgressEvent| {
                let output = serde_json::json!({"type": "progress", "event": event});
                eprintln!(
                    "{}",
                    serde_json::to_string(&output).expect("progress event is serializable")
                );
            };
            let result = run_s2_audit(&request, Some(&mut on_progress))?;
            let output = S2CliOutput {
                status: "completed",
                report_path: result.report_path,
                summary: result.analysis.summary,
            };
            println!("{}", serde_json::to_string_pretty(&output)?);
        }
        Command::TemplateSuggest(arguments) => {
            let depth = if arguments.recursive {
                ScanDepth::Full
            } else {
                ScanDepth::RootOnly
            };
            let suggestion =
                recommend_template(&arguments.template_dir, &arguments.input_dir, depth)?;
            println!("{}", serde_json::to_string_pretty(&suggestion)?);
        }
        Command::ConditionalEvaluate(arguments) => {
            let report = evaluate_xlsx_workbook(&arguments.workbook, &arguments.scan_areas)?;
            println!("{}", serde_json::to_string_pretty(&report)?);
        }
    }
    Ok(())
}

fn default_ui_command() -> Command {
    Command::Serve(ServeArgs {
        port: 8750,
        config: None,
        fallback_if_port_busy: true,
    })
}

fn default_s2_config_path() -> PathBuf {
    if let Some(path) = std::env::var_os("OPEN_DATA_AUDIT_S2_CONFIG") {
        return PathBuf::from(path);
    }

    let current_directory = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    let executable = std::env::current_exe().ok();
    default_s2_config_path_from(&current_directory, executable.as_deref())
}

fn default_s2_config_path_from(current_directory: &Path, executable: Option<&Path>) -> PathBuf {
    let mut roots = vec![current_directory];
    if let Some(executable_directory) = executable.and_then(Path::parent) {
        roots.extend(executable_directory.ancestors());
    }

    for root in roots {
        for relative in [
            "audit-core/config-real/2.报表采集系统_配置.xlsx",
            "audit-core/config/默认配置/2.报表采集系统_配置.xlsx",
            "config-real/2.报表采集系统_配置.xlsx",
            "config/默认配置/2.报表采集系统_配置.xlsx",
        ] {
            let candidate = root.join(relative);
            if candidate.is_file() {
                return candidate;
            }
        }
    }

    executable
        .and_then(Path::parent)
        .map(|directory| directory.join("config/默认配置/2.报表采集系统_配置.xlsx"))
        .unwrap_or_else(|| current_directory.join("2.报表采集系统_配置.xlsx"))
}

#[cfg(test)]
mod tests {
    use super::{default_s2_config_path_from, default_ui_command, Cli, Command};
    use clap::Parser;
    use std::{
        fs,
        time::{SystemTime, UNIX_EPOCH},
    };

    #[test]
    fn no_arguments_select_the_existing_audit_ui() {
        assert!(matches!(
            default_ui_command(),
            Command::Serve(super::ServeArgs {
                port: 8750,
                config: None,
                fallback_if_port_busy: true,
            })
        ));
    }

    #[test]
    fn explicit_serve_command_keeps_the_requested_port_strict() {
        let parsed = Cli::try_parse_from(["open-data-audit", "serve", "--port", "8750"])
            .expect("parse serve command");
        let Command::Serve(arguments) = parsed.command else {
            panic!("expected serve command");
        };
        assert!(!arguments.fallback_if_port_busy);
    }

    #[test]
    fn default_config_searches_up_from_the_executable_directory() {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock after epoch")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit-default-config-{}-{nonce}",
            std::process::id()
        ));
        let expected = root.join("audit-core/config/默认配置/2.报表采集系统_配置.xlsx");
        fs::create_dir_all(expected.parent().expect("config parent")).expect("create config dir");
        fs::write(&expected, b"placeholder").expect("write placeholder");

        let executable = root.join("target/debug/open-data-audit.exe");
        let actual = default_s2_config_path_from(
            &std::env::temp_dir().join("unrelated-working-directory"),
            Some(&executable),
        );

        let _ = fs::remove_dir_all(&root);
        assert_eq!(actual, expected);
    }
}
