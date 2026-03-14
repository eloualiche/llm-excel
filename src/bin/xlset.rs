use xlcat::cell::{parse_assignment, parse_cell_ref, CellAssignment};
use xlcat::writer::write_cells;

use anyhow::Result;
use clap::Parser;
use std::io::{self, BufRead};
use std::path::PathBuf;
use std::process;

// ---------------------------------------------------------------------------
// CLI definition
// ---------------------------------------------------------------------------

#[derive(Parser, Debug)]
#[command(name = "xlset", about = "Write values into Excel cells")]
struct Cli {
    /// Path to .xlsx file
    file: PathBuf,

    /// Cell assignments, e.g. A1=42 B2=hello
    #[arg(trailing_var_arg = true, num_args = 0..)]
    assignments: Vec<String>,

    /// Target sheet by name or 0-based index (default: first sheet)
    #[arg(long, default_value = "")]
    sheet: String,

    /// Write to a different file instead of updating in-place
    #[arg(long)]
    output: Option<PathBuf>,

    /// Read assignments from a CSV file, or `-` for stdin
    #[arg(long)]
    from: Option<String>,
}

// ---------------------------------------------------------------------------
// ArgError — user-facing argument errors (exit code 2)
// ---------------------------------------------------------------------------

#[derive(Debug)]
struct ArgError(String);

impl std::fmt::Display for ArgError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl std::error::Error for ArgError {}

// ---------------------------------------------------------------------------
// CSV parsing
// ---------------------------------------------------------------------------

/// Read cell assignments from a CSV source (file path or `-` for stdin).
///
/// Format: `cell,value` per line.
/// - First row is skipped if its first field is not a valid cell reference (header detection).
/// - Values may use RFC 4180 quoting: `A1,"hello, world"`.
fn read_csv_assignments(source: &str) -> Result<Vec<CellAssignment>> {
    let lines: Vec<String> = if source == "-" {
        let stdin = io::stdin();
        stdin.lock().lines().collect::<std::io::Result<_>>()?
    } else {
        let file = std::fs::File::open(source)
            .map_err(|e| anyhow::anyhow!("cannot open --from file '{}': {}", source, e))?;
        io::BufReader::new(file)
            .lines()
            .collect::<std::io::Result<_>>()?
    };

    let mut assignments = Vec::new();
    let mut skip_first = false;
    let mut first_line = true;

    for (line_idx, line) in lines.iter().enumerate() {
        let line_num = line_idx + 1;
        let trimmed = line.trim();
        if trimmed.is_empty() {
            first_line = false;
            continue;
        }

        // Split on first comma not inside quotes
        let (cell_str, value_str) = split_csv_line(trimmed).ok_or_else(|| {
            ArgError(format!(
                "--from line {}: expected 'cell,value' but got '{}'",
                line_num, trimmed
            ))
        })?;

        let cell_str = cell_str.trim();
        let value_str = unquote_csv(value_str.trim());

        // Header detection: if the first row's cell field is not a valid cell ref, skip it
        if first_line {
            first_line = false;
            if parse_cell_ref(cell_str).is_err() {
                skip_first = true;
                continue;
            }
        }
        let _ = skip_first; // already consumed above

        let _cell = parse_cell_ref(cell_str).map_err(|e| {
            ArgError(format!("--from line {}: invalid cell reference: {}", line_num, e))
        })?;

        // Build a synthetic assignment string and parse value via infer logic.
        // Since we already have cell and raw value separately, construct CellAssignment directly.
        let assignment_str = format!("{}={}", cell_str, value_str);
        let assignment = parse_assignment(&assignment_str).map_err(|e| {
            ArgError(format!("--from line {}: {}", line_num, e))
        })?;

        assignments.push(assignment);
    }

    Ok(assignments)
}

/// Split a CSV line on the first comma that is not inside double quotes.
/// Returns `(left, right)` or `None` if no comma is found outside quotes.
fn split_csv_line(line: &str) -> Option<(&str, &str)> {
    let mut in_quotes = false;
    let mut escaped = false;

    for (i, ch) in line.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }
        match ch {
            '"' => in_quotes = !in_quotes,
            ',' if !in_quotes => {
                return Some((&line[..i], &line[i + 1..]));
            }
            _ => {}
        }
    }
    None
}

/// Remove RFC 4180 quoting from a CSV field value.
/// `"hello, world"` → `hello, world`
/// `"say ""hi"""` → `say "hi"`
fn unquote_csv(s: &str) -> String {
    if s.starts_with('"') && s.ends_with('"') && s.len() >= 2 {
        let inner = &s[1..s.len() - 1];
        inner.replace("\"\"", "\"")
    } else {
        s.to_string()
    }
}

// ---------------------------------------------------------------------------
// Orchestration
// ---------------------------------------------------------------------------

fn run(cli: &Cli) -> Result<()> {
    // 1. Validate input file exists
    if !cli.file.exists() {
        return Err(anyhow::anyhow!(
            "file not found: '{}'",
            cli.file.display()
        ));
    }

    // 2. Collect assignments from --from CSV if provided
    let mut assignments: Vec<CellAssignment> = Vec::new();

    if let Some(ref source) = cli.from {
        let csv_assignments = read_csv_assignments(source)?;
        assignments.extend(csv_assignments);
    }

    // 3. Collect assignments from positional args
    for arg in &cli.assignments {
        let a = parse_assignment(arg).map_err(|e| ArgError(e))?;
        assignments.push(a);
    }

    // 4. Require at least one assignment
    if assignments.is_empty() {
        return Err(ArgError(
            "no assignments provided — use A1=value syntax or --from <file>".into(),
        )
        .into());
    }

    // 5. Determine output path
    let output_path = cli.output.clone().unwrap_or_else(|| cli.file.clone());

    // 6. Call writer
    let (count, sheet_name) = write_cells(&cli.file, &output_path, &cli.sheet, &assignments)?;

    // 7. Print confirmation to stderr
    let file_name = output_path
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| output_path.display().to_string());

    eprintln!("xlset: updated {} cells in {} ({})", count, sheet_name, file_name);

    Ok(())
}

// ---------------------------------------------------------------------------
// main()
// ---------------------------------------------------------------------------

fn main() {
    let cli = Cli::parse();
    if let Err(err) = run(&cli) {
        if err.downcast_ref::<ArgError>().is_some() {
            eprintln!("xlset: {err}");
            process::exit(2);
        }
        eprintln!("xlset: {err}");
        process::exit(1);
    }
}
