# xlset Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an `xlset` binary to the xlcat repo that modifies cells in existing xlsx files using umya-spreadsheet, with shared cell-parsing code.

**Architecture:** Restructure the project from a single binary to a library crate with two binary entry points (`xlcat`, `xlset`). Shared code (cell address parsing, type inference) lives in `lib.rs`/`cell.rs`. xlset uses umya-spreadsheet for round-trip read-modify-write. xlcat continues using calamine + polars for reading.

**Tech Stack:** Rust, umya-spreadsheet (xlsx editing), clap (CLI), calamine + polars (xlcat reading, existing)

**Spec:** `docs/superpowers/specs/2026-03-13-xlset-design.md`

---

## Chunk 1: Restructure and Cell Module

### Task 1: Restructure to Library + Two Binaries

**Files:**
- Create: `src/lib.rs`
- Create: `src/bin/xlcat.rs` (moved from `src/main.rs`)
- Delete: `src/main.rs`
- Modify: `Cargo.toml`

The project currently has `src/main.rs` as the single binary entry point
with `mod formatter; mod metadata; mod reader;`. We need to convert to a
library crate so both `xlcat` and `xlset` can share code.

- [ ] **Step 1: Create `src/lib.rs`**

```rust
pub mod cell;
pub mod formatter;
pub mod metadata;
pub mod reader;
pub mod writer;
```

Note: `cell` and `writer` don't exist yet. Create stub files so the
crate compiles:

`src/cell.rs`:
```rust
// Cell address parsing and value type inference (implemented in Task 2)
```

`src/writer.rs`:
```rust
// umya-spreadsheet write logic (implemented in Task 3)
```

- [ ] **Step 2: Move `src/main.rs` to `src/bin/xlcat.rs`**

Copy `src/main.rs` to `src/bin/xlcat.rs`. Then change the module imports
at the top from:

```rust
mod formatter;
mod metadata;
mod reader;
```

to:

```rust
use xlcat::formatter;
use xlcat::metadata;
use xlcat::metadata::{FileInfo, SheetInfo};
use xlcat::reader;
```

Remove the `use metadata::{FileInfo, SheetInfo};` line (now part of the
use statement above). The `use polars::prelude::*;` stays — xlcat needs
it directly.

Then delete `src/main.rs`.

- [ ] **Step 3: Update `Cargo.toml`**

Add library and binary sections, plus umya-spreadsheet dependency:

```toml
[package]
name = "xlcat"
version = "0.2.0"
edition = "2024"

[lib]
name = "xlcat"
path = "src/lib.rs"

[[bin]]
name = "xlcat"
path = "src/bin/xlcat.rs"

[[bin]]
name = "xlset"
path = "src/bin/xlset.rs"

[dependencies]
calamine = "0.26"
polars = { version = "0.46", features = ["dtype-date", "dtype-datetime", "dtype-duration", "csv"] }
clap = { version = "4", features = ["derive"] }
anyhow = "1"
umya-spreadsheet = "2"

[profile.release]
strip = true
lto = true
codegen-units = 1
panic = "abort"
opt-level = "z"

[dev-dependencies]
rust_xlsxwriter = "0.82"
assert_cmd = "2"
predicates = "3"
tempfile = "3"
```

- [ ] **Step 4: Create stub `src/bin/xlset.rs`**

```rust
fn main() {
    eprintln!("xlset: not yet implemented");
    std::process::exit(1);
}
```

- [ ] **Step 5: Verify both binaries compile and tests pass**

Run: `cargo test`
Run: `cargo build --bin xlcat`
Run: `cargo build --bin xlset`
Expected: all 49 existing tests pass, both binaries compile.

If tests fail because they can't find the `xlcat` binary, check that
`assert_cmd::Command::cargo_bin("xlcat")` still resolves correctly with
the new `[[bin]]` layout.

- [ ] **Step 6: Commit**

```bash
git add src/ Cargo.toml
git commit -m "refactor: restructure to lib + two binaries (xlcat, xlset)"
```

---

### Task 2: Cell Module — A1 Parser and Value Type Inference

**Files:**
- Create: `src/cell.rs` (replace stub)

This is shared code used by xlset (and potentially xlcat later). No
external crate dependencies — pure Rust.

- [ ] **Step 1: Implement cell.rs with tests**

Replace the stub `src/cell.rs` with:

```rust
use std::fmt;

/// A parsed cell reference (e.g., "A1" → col=0, row=0).
#[derive(Debug, Clone, PartialEq)]
pub struct CellRef {
    pub col: u32,   // 0-based
    pub row: u32,   // 0-based
    pub label: String, // original string, uppercased (e.g., "A1")
}

/// A typed value to write into a cell.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    String(String),
    Integer(i64),
    Float(f64),
    Bool(bool),
    Date { year: i32, month: u32, day: u32 },
    Empty,
}

/// A full cell assignment: reference + value.
#[derive(Debug, Clone)]
pub struct CellAssignment {
    pub cell: CellRef,
    pub value: CellValue,
}

/// Parse a cell reference like "A1", "Z99", "AA1", "XFD1048576".
/// Case-insensitive. Returns error if invalid.
pub fn parse_cell_ref(s: &str) -> Result<CellRef, String> {
    let s = s.trim();
    if s.is_empty() {
        return Err("Empty cell reference".into());
    }

    let upper = s.to_uppercase();
    let bytes = upper.as_bytes();

    // Split into letter part and digit part
    let letter_end = bytes.iter().position(|b| b.is_ascii_digit())
        .ok_or_else(|| format!("Invalid cell reference: {s} (no row number)"))?;
    if letter_end == 0 {
        return Err(format!("Invalid cell reference: {s} (no column letter)"));
    }

    let col_str = &upper[..letter_end];
    let row_str = &upper[letter_end..];

    // Parse column: A=0, B=1, ..., Z=25, AA=26, ...
    let col = parse_col(col_str)
        .ok_or_else(|| format!("Invalid column: {col_str}"))?;
    if col > 16383 {
        return Err(format!("Column out of range: {col_str} (max XFD)"));
    }

    // Parse row: 1-based in input, 0-based internally
    let row_1based: u32 = row_str.parse()
        .map_err(|_| format!("Invalid row number: {row_str}"))?;
    if row_1based == 0 || row_1based > 1_048_576 {
        return Err(format!("Row out of range: {row_1based} (must be 1-1048576)"));
    }

    Ok(CellRef {
        col,
        row: row_1based - 1,
        label: upper,
    })
}

fn parse_col(s: &str) -> Option<u32> {
    let mut result: u32 = 0;
    for &b in s.as_bytes() {
        if !b.is_ascii_uppercase() {
            return None;
        }
        result = result.checked_mul(26)?.checked_add((b - b'A') as u32 + 1)?;
    }
    Some(result - 1) // convert to 0-based
}

/// Parse a cell assignment string: "A1=42", "B2:str=hello", etc.
/// Format: <cellref>[:<type_tag>]=<value>
pub fn parse_assignment(s: &str) -> Result<CellAssignment, String> {
    let eq_pos = s.find('=')
        .ok_or_else(|| format!("Invalid assignment (no '='): {s}"))?;

    let lhs = &s[..eq_pos];
    let rhs = &s[eq_pos + 1..];

    // Check for type tag
    let (cell_str, type_tag) = if let Some(colon_pos) = lhs.find(':') {
        (&lhs[..colon_pos], Some(&lhs[colon_pos + 1..]))
    } else {
        (lhs, None)
    };

    let cell = parse_cell_ref(cell_str)?;

    let value = if let Some(tag) = type_tag {
        parse_value_with_tag(rhs, tag)?
    } else {
        infer_value(rhs)
    };

    Ok(CellAssignment { cell, value })
}

fn parse_value_with_tag(s: &str, tag: &str) -> Result<CellValue, String> {
    match tag.to_lowercase().as_str() {
        "str" => Ok(CellValue::String(s.to_string())),
        "num" => {
            if let Ok(i) = s.parse::<i64>() {
                Ok(CellValue::Integer(i))
            } else if let Ok(f) = s.parse::<f64>() {
                Ok(CellValue::Float(f))
            } else {
                Err(format!("Cannot parse as number: {s}"))
            }
        }
        "bool" => {
            match s.to_lowercase().as_str() {
                "true" | "1" | "yes" => Ok(CellValue::Bool(true)),
                "false" | "0" | "no" => Ok(CellValue::Bool(false)),
                _ => Err(format!("Cannot parse as boolean: {s}")),
            }
        }
        "date" => parse_date(s),
        _ => Err(format!("Unknown type tag: {tag}. Valid tags: str, num, bool, date")),
    }
}

/// Auto-infer value type from string content.
pub fn infer_value(s: &str) -> CellValue {
    if s.is_empty() {
        return CellValue::Empty;
    }
    // Boolean
    match s.to_lowercase().as_str() {
        "true" => return CellValue::Bool(true),
        "false" => return CellValue::Bool(false),
        _ => {}
    }
    // Integer (no decimal point)
    if let Ok(i) = s.parse::<i64>() {
        return CellValue::Integer(i);
    }
    // Float
    if let Ok(f) = s.parse::<f64>() {
        return CellValue::Float(f);
    }
    // Date (YYYY-MM-DD)
    if let Ok(cv) = parse_date(s) {
        return cv;
    }
    // String fallback
    CellValue::String(s.to_string())
}

fn parse_date(s: &str) -> Result<CellValue, String> {
    let parts: Vec<&str> = s.split('-').collect();
    if parts.len() != 3 {
        return Err(format!("Invalid date format: {s} (expected YYYY-MM-DD)"));
    }
    let year: i32 = parts[0].parse().map_err(|_| format!("Invalid year: {}", parts[0]))?;
    let month: u32 = parts[1].parse().map_err(|_| format!("Invalid month: {}", parts[1]))?;
    let day: u32 = parts[2].parse().map_err(|_| format!("Invalid day: {}", parts[2]))?;
    if month < 1 || month > 12 || day < 1 || day > 31 {
        return Err(format!("Invalid date: {s}"));
    }
    Ok(CellValue::Date { year, month, day })
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{}", self.label)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref_simple() {
        let r = parse_cell_ref("A1").unwrap();
        assert_eq!(r.col, 0);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn test_parse_cell_ref_z26() {
        let r = parse_cell_ref("Z1").unwrap();
        assert_eq!(r.col, 25);
    }

    #[test]
    fn test_parse_cell_ref_aa() {
        let r = parse_cell_ref("AA1").unwrap();
        assert_eq!(r.col, 26);
    }

    #[test]
    fn test_parse_cell_ref_case_insensitive() {
        let r = parse_cell_ref("a1").unwrap();
        assert_eq!(r.col, 0);
        assert_eq!(r.row, 0);
        assert_eq!(r.label, "A1");
    }

    #[test]
    fn test_parse_cell_ref_row_offset() {
        let r = parse_cell_ref("B10").unwrap();
        assert_eq!(r.col, 1);
        assert_eq!(r.row, 9); // 0-based
    }

    #[test]
    fn test_parse_cell_ref_invalid() {
        assert!(parse_cell_ref("").is_err());
        assert!(parse_cell_ref("123").is_err());
        assert!(parse_cell_ref("A0").is_err()); // row 0 invalid
        assert!(parse_cell_ref("A").is_err()); // no row
    }

    #[test]
    fn test_infer_value_integer() {
        assert_eq!(infer_value("42"), CellValue::Integer(42));
        assert_eq!(infer_value("-5"), CellValue::Integer(-5));
    }

    #[test]
    fn test_infer_value_float() {
        assert_eq!(infer_value("3.14"), CellValue::Float(3.14));
    }

    #[test]
    fn test_infer_value_bool() {
        assert_eq!(infer_value("true"), CellValue::Bool(true));
        assert_eq!(infer_value("false"), CellValue::Bool(false));
        assert_eq!(infer_value("TRUE"), CellValue::Bool(true));
    }

    #[test]
    fn test_infer_value_date() {
        assert_eq!(
            infer_value("2024-01-15"),
            CellValue::Date { year: 2024, month: 1, day: 15 }
        );
    }

    #[test]
    fn test_infer_value_string() {
        assert_eq!(infer_value("hello"), CellValue::String("hello".into()));
    }

    #[test]
    fn test_infer_leading_zero_becomes_integer() {
        // Leading zeros get lost — this is why :str type tags exist
        assert_eq!(infer_value("07401"), CellValue::Integer(7401));
    }

    #[test]
    fn test_infer_value_empty() {
        assert_eq!(infer_value(""), CellValue::Empty);
    }

    #[test]
    fn test_parse_assignment_basic() {
        let a = parse_assignment("A1=42").unwrap();
        assert_eq!(a.cell.col, 0);
        assert_eq!(a.cell.row, 0);
        assert_eq!(a.value, CellValue::Integer(42));
    }

    #[test]
    fn test_parse_assignment_with_tag() {
        let a = parse_assignment("A1:str=07401").unwrap();
        assert_eq!(a.value, CellValue::String("07401".into()));
    }

    #[test]
    fn test_parse_assignment_no_equals() {
        assert!(parse_assignment("A1").is_err());
    }

    #[test]
    fn test_parse_assignment_empty_value() {
        let a = parse_assignment("A1=").unwrap();
        assert_eq!(a.value, CellValue::Empty);
    }

    #[test]
    fn test_parse_assignment_string_with_spaces() {
        let a = parse_assignment("A1=hello world").unwrap();
        assert_eq!(a.value, CellValue::String("hello world".into()));
    }
}
```

- [ ] **Step 2: Run tests**

Run: `cargo test cell::tests`
Expected: all tests pass.

- [ ] **Step 3: Commit**

```bash
git add src/cell.rs
git commit -m "feat: add cell module — A1 parser and value type inference"
```

---

## Chunk 2: Writer and xlset Binary

### Task 3: Writer Module

**Files:**
- Create: `src/writer.rs` (replace stub)

Uses umya-spreadsheet to open an existing xlsx, modify cells, and save.

- [ ] **Step 1: Implement writer.rs**

Replace the stub `src/writer.rs` with:

```rust
use anyhow::{Context, Result};
use std::path::Path;
use umya_spreadsheet::*;

use crate::cell::{CellAssignment, CellValue};

/// Open an xlsx file, apply cell assignments to the given sheet, and save.
/// Open an xlsx, apply assignments, save. Returns (count, sheet_name).
pub fn write_cells(
    input_path: &Path,
    output_path: &Path,
    sheet_selector: &str,
    assignments: &[CellAssignment],
) -> Result<(usize, String)> {
    // Validate extension
    let ext = input_path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase());
    match ext.as_deref() {
        Some("xlsx") | Some("xlsm") => {}
        Some("xls") => anyhow::bail!("xlset only supports .xlsx files, not .xls"),
        Some(other) => anyhow::bail!("Expected .xlsx file, got: .{other}"),
        None => anyhow::bail!("Expected .xlsx file, got: no extension"),
    }

    let mut book = reader::xlsx::read(input_path)
        .with_context(|| format!("Cannot open workbook: {}", input_path.display()))?;

    let sheet_names = get_sheet_names(&book);
    let resolved_name = resolve_sheet_name(&sheet_names, sheet_selector)?;
    let sheet = resolve_sheet(&mut book, sheet_selector)?;

    let mut count = 0;
    for assignment in assignments {
        apply_assignment(sheet, assignment)?;
        count += 1;
    }

    writer::xlsx::write(&book, output_path)
        .with_context(|| format!("Cannot write to: {}", output_path.display()))?;

    Ok((count, resolved_name))
}

fn resolve_sheet_name(sheet_names: &[String], selector: &str) -> Result<String> {
    if selector.is_empty() {
        return Ok(sheet_names.first().cloned().unwrap_or_else(|| "Sheet1".into()));
    }
    if let Some(name) = sheet_names.iter().find(|n| n.as_str() == selector) {
        return Ok(name.clone());
    }
    if let Ok(idx) = selector.parse::<usize>() {
        if idx < sheet_names.len() {
            return Ok(sheet_names[idx].clone());
        }
    }
    Ok(selector.to_string())
}

fn get_sheet_names(book: &Spreadsheet) -> Vec<String> {
    let mut names = Vec::new();
    for i in 0..book.get_sheet_count() {
        if let Some(sheet) = book.get_sheet(&i) {
            names.push(sheet.get_name().to_string());
        }
    }
    names
}

fn resolve_sheet<'a>(
    book: &'a mut Spreadsheet,
    selector: &str,
) -> Result<&'a mut Worksheet> {
    let sheet_names = get_sheet_names(book);

    // Empty selector → first sheet
    if selector.is_empty() {
        return book.get_sheet_mut(&0)
            .ok_or_else(|| anyhow::anyhow!("Workbook has no sheets"));
    }

    // Try name match first
    if let Some(idx) = sheet_names.iter().position(|n| n == selector) {
        return book.get_sheet_mut(&idx)
            .ok_or_else(|| anyhow::anyhow!("Sheet not found: {selector}"));
    }

    // Try 0-based index
    if let Ok(idx) = selector.parse::<usize>() {
        if idx < sheet_names.len() {
            return book.get_sheet_mut(&idx)
                .ok_or_else(|| anyhow::anyhow!("Sheet index {idx} out of range"));
        }
        let available = sheet_names.join(", ");
        anyhow::bail!("Sheet index {idx} out of range. Available sheets: {available}");
    }

    let available = sheet_names.join(", ");
    anyhow::bail!("Sheet not found: {selector}. Available sheets: {available}");
}

fn apply_assignment(sheet: &mut Worksheet, assignment: &CellAssignment) -> Result<()> {
    // Use string-based cell reference (e.g., "A1") per spec.
    // This avoids 0-based/1-based conversion and col/row ordering issues.
    let cell = sheet.get_cell_mut(&assignment.cell.label);

    match &assignment.value {
        CellValue::String(s) => {
            // Use set_value_string to prevent auto-conversion of
            // numeric-looking strings (e.g., "07401" → 7401).
            cell.set_value_string(s);
        }
        CellValue::Integer(i) => {
            cell.set_value_number(*i as f64);
        }
        CellValue::Float(f) => {
            cell.set_value_number(*f);
        }
        CellValue::Bool(b) => {
            cell.set_value_bool(*b);
        }
        CellValue::Date { year, month, day } => {
            let serial = date_to_serial(*year, *month, *day);
            cell.set_value_number(serial);
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code("yyyy-mm-dd");
        }
        CellValue::Empty => {
            cell.set_value_string("");
        }
    }

    Ok(())
}

/// Convert a date to Excel serial number (days since 1899-12-30).
fn date_to_serial(year: i32, month: u32, day: u32) -> f64 {
    // Excel serial date: Jan 1, 1900 = 1
    // But Excel incorrectly treats 1900 as a leap year (Lotus 1-2-3 bug),
    // so dates after Feb 28, 1900 are off by 1.
    let y = year as i64;
    let m = month as i64;
    let d = day as i64;

    // Using the algorithm for Julian Day Number, then converting to Excel serial
    let a = (14 - m) / 12;
    let y2 = y + 4800 - a;
    let m2 = m + 12 * a - 3;

    let jdn = d + (153 * m2 + 2) / 5 + 365 * y2 + y2 / 4 - y2 / 100 + y2 / 400 - 32045;

    // Excel epoch: Dec 30, 1899 = JDN 2415018.5
    // But Excel serial 1 = Jan 1, 1900
    let excel_epoch_jdn: i64 = 2415020; // Jan 1, 1900 in JDN
    let serial = jdn - excel_epoch_jdn + 1;

    // Lotus 1-2-3 bug: Excel thinks Feb 29, 1900 exists.
    // For dates after Feb 28, 1900, add 1.
    if serial > 59 {
        (serial + 1) as f64
    } else {
        serial as f64
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_date_to_serial_known_dates() {
        // Jan 1, 1900 = 1
        assert_eq!(date_to_serial(1900, 1, 1), 1.0);
        // Jan 1, 2024 = 45292
        assert_eq!(date_to_serial(2024, 1, 1), 45292.0);
    }
}
```

**Important:** The umya-spreadsheet API may differ from what's shown.
Key things to verify:
- `reader::xlsx::read(path)` — reads xlsx file
- `writer::xlsx::write(&book, path)` — writes xlsx file
- `book.get_sheet_mut(&index)` — returns `Option<&mut Worksheet>`
- `sheet.get_cell_mut((col, row))` — may use `(col, row)` or
  `(&col, &row)` tuples. The column and row are 1-based u32.
- `cell.set_value()`, `cell.set_value_number()`, `cell.set_value_bool()`
- `cell.get_style_mut().get_number_format_mut().set_format_code()`

If any of these don't match, check `cargo doc -p umya-spreadsheet --open`
and adjust.

- [ ] **Step 2: Run tests**

Run: `cargo test writer::tests`
Expected: pass (the date serial test at minimum).

- [ ] **Step 3: Commit**

```bash
git add src/writer.rs
git commit -m "feat: add writer module — umya-spreadsheet cell editing"
```

---

### Task 4: xlset Binary

**Files:**
- Create: `src/bin/xlset.rs` (replace stub)

- [ ] **Step 1: Implement xlset.rs**

Replace the stub `src/bin/xlset.rs` with:

```rust
use anyhow::Result;
use clap::Parser;
use std::io::{self, BufRead};
use std::path::PathBuf;
use std::process;

use xlcat::cell::{self, CellAssignment};
use xlcat::writer;

#[derive(Parser, Debug)]
#[command(name = "xlset", about = "Modify cells in Excel (.xlsx) files")]
struct Cli {
    /// Path to .xlsx file
    file: PathBuf,

    /// Cell assignments: A1=42 B2:str=hello
    #[arg(trailing_var_arg = true)]
    assignments: Vec<String>,

    /// Target sheet by name or 0-based index (default: first sheet)
    #[arg(long, default_value = "")]
    sheet: String,

    /// Write to a different file instead of modifying in-place
    #[arg(long)]
    output: Option<PathBuf>,

    /// Read cell assignments from CSV file (or - for stdin)
    #[arg(long, value_name = "FILE")]
    from: Option<String>,
}

#[derive(Debug)]
struct ArgError(String);
impl std::fmt::Display for ArgError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", self.0)
    }
}
impl std::error::Error for ArgError {}

fn run(cli: &Cli) -> Result<()> {
    // Validate file exists
    if !cli.file.exists() {
        anyhow::bail!("File not found: {}", cli.file.display());
    }

    // Collect assignments from --from CSV
    let mut all_assignments: Vec<CellAssignment> = Vec::new();

    if let Some(ref from_path) = cli.from {
        let csv_assignments = read_csv_assignments(from_path)?;
        all_assignments.extend(csv_assignments);
    }

    // Collect assignments from positional args
    for arg in &cli.assignments {
        let assignment = cell::parse_assignment(arg)
            .map_err(|e| ArgError(e))?;
        all_assignments.push(assignment);
    }

    if all_assignments.is_empty() {
        return Err(ArgError(
            "No cell assignments provided. Use positional args or --from.".into()
        ).into());
    }

    // Determine output path
    let output_path = cli.output.as_ref().unwrap_or(&cli.file);

    // Write cells
    let (count, sheet_name) = writer::write_cells(
        &cli.file,
        output_path,
        &cli.sheet,
        &all_assignments,
    )?;

    let file_name = cli.file.file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| cli.file.display().to_string());

    eprintln!("xlset: updated {count} cells in {sheet_name} ({file_name})");
    Ok(())
}

fn read_csv_assignments(path: &str) -> Result<Vec<CellAssignment>> {
    let reader: Box<dyn BufRead> = if path == "-" {
        Box::new(io::stdin().lock())
    } else {
        let file = std::fs::File::open(path)
            .map_err(|e| anyhow::anyhow!("Cannot open CSV file {path}: {e}"))?;
        Box::new(io::BufReader::new(file))
    };

    let mut assignments = Vec::new();
    let mut line_num = 0;

    for line_result in reader.lines() {
        line_num += 1;
        let line = line_result?;
        let line = line.trim();
        if line.is_empty() {
            continue;
        }

        // Parse CSV line: cell,value (simple split on first comma)
        // For RFC 4180 quoting, handle quoted values
        let (cell_part, value_part) = parse_csv_line(line, line_num)?;

        // Skip header row: if first field doesn't parse as a cell ref
        if line_num == 1 {
            // Check if cell_part looks like a valid cell reference (possibly with type tag)
            let cell_str = if let Some(colon) = cell_part.find(':') {
                &cell_part[..colon]
            } else {
                &cell_part
            };
            if cell::parse_cell_ref(cell_str).is_err() {
                continue; // skip header
            }
        }

        let assignment_str = format!("{cell_part}={value_part}");
        let assignment = cell::parse_assignment(&assignment_str)
            .map_err(|e| anyhow::anyhow!("Error on line {line_num}: {e}"))?;
        assignments.push(assignment);
    }

    Ok(assignments)
}

/// Parse a CSV line into (cell, value), handling RFC 4180 quoting.
fn parse_csv_line(line: &str, line_num: usize) -> Result<(String, String)> {
    // Find the first comma not inside quotes
    let mut in_quotes = false;
    let mut comma_pos = None;

    for (i, ch) in line.char_indices() {
        match ch {
            '"' => in_quotes = !in_quotes,
            ',' if !in_quotes => {
                comma_pos = Some(i);
                break;
            }
            _ => {}
        }
    }

    let comma_pos = comma_pos
        .ok_or_else(|| anyhow::anyhow!("Error on line {line_num}: expected cell,value format"))?;

    let cell_part = line[..comma_pos].trim().to_string();
    let mut value_part = line[comma_pos + 1..].trim().to_string();

    // Unquote the value if it's quoted
    if value_part.starts_with('"') && value_part.ends_with('"') && value_part.len() >= 2 {
        value_part = value_part[1..value_part.len() - 1].replace("\"\"", "\"");
    }

    Ok((cell_part, value_part))
}

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
```

**Important notes on the CLI parsing:**
- `trailing_var_arg = true` on the `assignments` field tells clap to
  collect all remaining positional args. Verify this works — if not,
  try `#[arg(num_args = 0..)]` instead.
- The sheet name in the confirmation message should come from the actual
  sheet, not just the --sheet arg. The stub uses `cli.sheet` but ideally
  should resolve to the actual sheet name. Adjust if needed.

- [ ] **Step 2: Verify it compiles**

Run: `cargo build --bin xlset`
Expected: compiles.

- [ ] **Step 3: Commit**

```bash
git add src/bin/xlset.rs
git commit -m "feat: add xlset binary — Excel cell writer CLI"
```

---

## Chunk 3: Integration Tests and Skill

### Task 5: Integration Tests

**Files:**
- Create: `tests/test_xlset.rs`

- [ ] **Step 1: Create integration tests**

```rust
mod common;

use assert_cmd::Command;
use predicates::prelude::*;
use tempfile::TempDir;

fn xlset() -> Command {
    Command::cargo_bin("xlset").unwrap()
}

fn xlcat() -> Command {
    Command::cargo_bin("xlcat").unwrap()
}

#[test]
fn test_set_single_cell() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_simple(&path);

    // Set cell A2 to "Modified"
    xlset()
        .arg(path.to_str().unwrap())
        .arg("A2=Modified")
        .assert()
        .success()
        .stderr(predicate::str::contains("updated 1 cells"));

    // Verify with xlcat
    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("Modified"));
}

#[test]
fn test_set_multiple_cells() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_simple(&path);

    xlset()
        .arg(path.to_str().unwrap())
        .arg("A2=Changed")
        .arg("B2=999")
        .assert()
        .success()
        .stderr(predicate::str::contains("updated 2 cells"));

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("Changed"))
        .stdout(predicate::str::contains("999"));
}

#[test]
fn test_set_with_type_tag() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_simple(&path);

    xlset()
        .arg(path.to_str().unwrap())
        .arg("A2:str=07401")
        .assert()
        .success();

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("07401"));
}

#[test]
fn test_set_with_output_file() {
    let dir = TempDir::new().unwrap();
    let src = dir.path().join("source.xlsx");
    let dst = dir.path().join("output.xlsx");
    common::create_simple(&src);

    xlset()
        .arg(src.to_str().unwrap())
        .arg("--output")
        .arg(dst.to_str().unwrap())
        .arg("A2=New")
        .assert()
        .success();

    // Output file has the change
    xlcat()
        .arg(dst.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("New"));

    // Source file unchanged
    xlcat()
        .arg(src.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("Alice")); // original value
}

#[test]
fn test_set_with_sheet_selection() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_multi_sheet(&path);

    xlset()
        .arg(path.to_str().unwrap())
        .arg("--sheet")
        .arg("Expenses")
        .arg("A2=Modified")
        .assert()
        .success();

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--sheet")
        .arg("Expenses")
        .assert()
        .success()
        .stdout(predicate::str::contains("Modified"));
}

#[test]
fn test_set_from_csv() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    let csv = dir.path().join("updates.csv");
    common::create_simple(&path);

    std::fs::write(&csv, "cell,value\nA2,Updated\nB2,999\n").unwrap();

    xlset()
        .arg(path.to_str().unwrap())
        .arg("--from")
        .arg(csv.to_str().unwrap())
        .assert()
        .success()
        .stderr(predicate::str::contains("updated 2 cells"));

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("Updated"));
}

#[test]
fn test_set_from_csv_no_header() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    let csv = dir.path().join("updates.csv");
    common::create_simple(&path);

    // No header — first line starts with valid cell ref
    std::fs::write(&csv, "A2,Updated\nB2,999\n").unwrap();

    xlset()
        .arg(path.to_str().unwrap())
        .arg("--from")
        .arg(csv.to_str().unwrap())
        .assert()
        .success()
        .stderr(predicate::str::contains("updated 2 cells"));
}

#[test]
fn test_error_no_assignments() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_simple(&path);

    xlset()
        .arg(path.to_str().unwrap())
        .assert()
        .code(2)
        .stderr(predicate::str::contains("No cell assignments"));
}

#[test]
fn test_error_file_not_found() {
    xlset()
        .arg("/nonexistent.xlsx")
        .arg("A1=42")
        .assert()
        .code(1)
        .stderr(predicate::str::contains("not found"));
}

#[test]
fn test_error_bad_cell_ref() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_simple(&path);

    xlset()
        .arg(path.to_str().unwrap())
        .arg("ZZZZZ1=42")
        .assert()
        .failure();
}

#[test]
fn test_error_xls_not_supported() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xls");
    std::fs::write(&path, "fake").unwrap();

    xlset()
        .arg(path.to_str().unwrap())
        .arg("A1=42")
        .assert()
        .failure()
        .stderr(predicate::str::contains("only supports .xlsx"));
}
```

- [ ] **Step 2: Run tests**

Run: `cargo test test_xlset`
Expected: all pass. If tests fail, debug:
- Check umya-spreadsheet API matches (method names, argument types)
- Check that `create_simple` fixtures produce valid xlsx that umya can read
- Check cell coordinate order (umya uses (col, row) not (row, col))

- [ ] **Step 3: Commit**

```bash
git add tests/test_xlset.rs
git commit -m "test: add xlset integration tests"
```

---

### Task 6: Claude Code /xlset Skill

**Files:**
- Create: `~/.claude/skills/xlset/SKILL.md`

- [ ] **Step 1: Create the skill**

```markdown
---
name: xlset
description: Modify cells in Excel (.xlsx) files using xlset. Use when the user asks to edit, update, change, or write values to an Excel spreadsheet, or when you need to programmatically update cells in an xlsx file.
---

# xlset — Excel Cell Writer

Modify cells in existing .xlsx files. Preserves formatting, formulas,
and all structure it doesn't touch.

## Tool Location

\```
/Users/loulou/.local/bin/xlset
\```

## Quick Reference

\```bash
# Set a single cell
xlset file.xlsx A1=42

# Set multiple cells
xlset file.xlsx A1=42 B2="hello world" C3=true

# Force type with tag (e.g., preserve leading zero)
xlset file.xlsx A1:str=07401

# Target a specific sheet
xlset file.xlsx --sheet Revenue A1=42

# Write to a new file (don't modify original)
xlset file.xlsx --output new.xlsx A1=42

# Bulk update from CSV
xlset file.xlsx --from updates.csv

# Bulk from stdin
echo "A1,42" | xlset file.xlsx --from -
\```

## Type Inference

Values are auto-detected:
- `42` → integer, `3.14` → float
- `true`/`false` → boolean
- `2024-01-15` → date
- Everything else → string

Override with tags: `:str`, `:num`, `:bool`, `:date`

## CSV Format

\```csv
cell,value
A1,42
B2,hello
C3:str,07401
\```

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Runtime error |
| 2 | Invalid arguments |

## Workflow

1. Use `xlcat file.xlsx` first to see current content
2. Use `xlset` to modify cells
3. Use `xlcat` again to verify changes
```

- [ ] **Step 2: Build release binaries and install**

```bash
cargo build --release
cp target/release/xlcat ~/.local/bin/xlcat
cp target/release/xlset ~/.local/bin/xlset
```

- [ ] **Step 3: Smoke test**

```bash
xlset --help
# Create a test file, modify it, verify
```

- [ ] **Step 4: Commit code changes**

```bash
git add -A
git commit -m "feat: add xlset binary and Claude Code skill"
```

---

## Summary

| Task | Component | Key deliverable |
|------|-----------|----------------|
| 1 | Restructure | lib.rs + two binaries |
| 2 | Cell module | A1 parser, type inference, shared code |
| 3 | Writer module | umya-spreadsheet read-modify-write |
| 4 | xlset binary | CLI, CSV parsing, orchestration |
| 5 | Integration tests | 11 tests verifying xlset end-to-end |
| 6 | Skill + release | /xlset skill, installed binaries |
