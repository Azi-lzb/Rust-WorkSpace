use anyhow::{Context, Result};
use clap::{Parser, Subcommand};
use excel_writer::{Color, Format, Workbook};
use serde_json::json;
use std::{fs, path::PathBuf};

#[derive(Parser)]
#[command(about = "Audit parity probe using the shared Excel reader and writer crates")]
struct Args {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    Read {
        #[arg(long)]
        input: PathBuf,
        #[arg(long)]
        output: PathBuf,
    },
    WriteSmoke {
        #[arg(long)]
        output: PathBuf,
    },
}

fn main() -> Result<()> {
    match Args::parse().command {
        Command::Read { input, output } => {
            let dump = excel_reader::read_workbook(&input)
                .with_context(|| format!("read {}", input.display()))?;
            if let Some(parent) = output.parent() {
                fs::create_dir_all(parent)?;
            }
            fs::write(&output, serde_json::to_vec_pretty(&dump)?)?;
            println!("{}", serde_json::to_string(&dump)?);
        }
        Command::WriteSmoke { output } => {
            let started = std::time::Instant::now();
            write_smoke(&output)?;
            println!(
                "{}",
                json!({"writer":"rust_xlsxwriter", "version":"0.101.0", "output":output.display().to_string(), "elapsed_ms":started.elapsed().as_millis()})
            );
        }
    }
    Ok(())
}

fn write_smoke(path: &PathBuf) -> Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    let mut workbook = Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("Synthetic")?;
        worksheet.write_number(0, 0, 1)?;
        worksheet.write_number(1, 0, 2)?;
        worksheet.write_number(2, 0, 3)?;
        worksheet.write_string(0, 1, "writer smoke")?;
        worksheet.write_formula(3, 3, "=SUM(A1:A3)")?;
        worksheet.write_string(4, 4, "named target")?;
        let format = Format::new()
            .set_bold()
            .set_background_color(Color::Yellow)
            .set_num_format("0.00");
        worksheet.write_number_with_format(5, 0, 12.5, &format)?;
        worksheet.merge_range(6, 0, 6, 1, "merged", &format)?;
        worksheet.set_row_height(2, 30.0)?;
        worksheet.set_column_width(8, 25.0)?;
        worksheet.set_print_area(0, 0, 9, 3)?;
    }
    workbook.define_name("SyntheticNamedCell", "Synthetic!$E$5")?;
    workbook
        .save(path)
        .with_context(|| format!("write {}", path.display()))?;
    Ok(())
}
