# xlcat Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Rust CLI tool that reads xls/xlsx files and outputs structured, LLM-friendly text, plus a Claude Code `/xls` skill.

**Architecture:** Calamine reads Excel files (sheets, cells, dimensions). Cell data is bridged into Polars DataFrames for type inference and statistics. Output is formatted as markdown tables or CSV. Clap handles CLI parsing. The tool is a single binary with no runtime dependencies.

**Tech Stack:** Rust, calamine (Excel reading), polars (DataFrame/stats), clap (CLI), rust_xlsxwriter (dev-dep for test fixtures)

**Spec:** `docs/superpowers/specs/2026-03-13-xlcat-design.md`

---

## Chunk 1: Foundation

### Task 1: Project Scaffolding

**Files:**
- Create: `xlcat/Cargo.toml`
- Create: `xlcat/src/main.rs`

- [ ] **Step 1: Initialize Cargo project**

Run: `cargo init xlcat` from the project root.

- [ ] **Step 2: Set up Cargo.toml with dependencies**

```toml
[package]
name = "xlcat"
version = "0.1.0"
edition = "2024"

[dependencies]
calamine = "0.26"
polars = { version = "0.46", features = ["dtype-date", "dtype-datetime", "dtype-duration", "csv"] }
clap = { version = "4", features = ["derive"] }
anyhow = "1"

[dev-dependencies]
rust_xlsxwriter = "0.82"
assert_cmd = "2"
predicates = "3"
tempfile = "3"
```

Note: pin polars to 0.46 (latest with stable Rust 1.90 support — 0.53 may
require nightly). Verify with `cargo check` and adjust if needed.

- [ ] **Step 3: Define CLI args in main.rs**

```rust
use anyhow::Result;
use clap::Parser;
use std::path::PathBuf;

#[derive(Parser, Debug)]
#[command(name = "xlcat", about = "View Excel files in the terminal")]
struct Cli {
    /// Path to .xls or .xlsx file
    file: PathBuf,

    /// Show only column names and types
    #[arg(long)]
    schema: bool,

    /// Show summary statistics
    #[arg(long)]
    describe: bool,

    /// Show first N rows
    #[arg(long)]
    head: Option<usize>,

    /// Show last N rows
    #[arg(long)]
    tail: Option<usize>,

    /// Show all rows (overrides large-file gate)
    #[arg(long)]
    all: bool,

    /// Select sheet by name or 0-based index
    #[arg(long)]
    sheet: Option<String>,

    /// Large-file threshold (default: 1M). Accepts: 500K, 1M, 10M, 1G
    #[arg(long, default_value = "1M", value_parser = parse_size)]
    max_size: u64,

    /// Output as CSV instead of markdown
    #[arg(long)]
    csv: bool,
}

fn parse_size(s: &str) -> Result<u64, String> {
    let s = s.trim();
    let (num_part, multiplier) = if s.ends_with('G') || s.ends_with('g') {
        (&s[..s.len() - 1], 1_073_741_824u64)
    } else if s.ends_with("GB") || s.ends_with("gb") {
        (&s[..s.len() - 2], 1_073_741_824u64)
    } else if s.ends_with('M') || s.ends_with('m') {
        (&s[..s.len() - 1], 1_048_576u64)
    } else if s.ends_with("MB") || s.ends_with("mb") {
        (&s[..s.len() - 2], 1_048_576u64)
    } else if s.ends_with('K') || s.ends_with('k') {
        (&s[..s.len() - 1], 1_024u64)
    } else if s.ends_with("KB") || s.ends_with("kb") {
        (&s[..s.len() - 2], 1_024u64)
    } else {
        (s, 1u64)
    };
    let num: f64 = num_part.parse().map_err(|_| format!("Invalid size: {s}"))?;
    Ok((num * multiplier as f64) as u64)
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // Validate flag combinations
    let mode_count = cli.schema as u8 + cli.describe as u8;
    if mode_count > 1 {
        anyhow::bail!("--schema and --describe are mutually exclusive");
    }
    if (cli.schema || cli.describe) && (cli.head.is_some() || cli.tail.is_some() || cli.all) {
        anyhow::bail!("--schema and --describe cannot be combined with --head, --tail, or --all");
    }
    if (cli.schema || cli.describe) && cli.csv {
        anyhow::bail!("--csv can only be used in data mode (not with --schema or --describe)");
    }

    eprintln!("xlcat: not yet implemented");
    std::process::exit(1);
}
```

- [ ] **Step 4: Verify it compiles**

Run: `cd xlcat && cargo check`
Expected: compiles with no errors (warnings OK).

- [ ] **Step 5: Commit**

```bash
git add xlcat/
git commit -m "feat: scaffold xlcat project with CLI arg parsing"
```

---

### Task 2: Test Fixture Generator

**Files:**
- Create: `xlcat/tests/fixtures.rs` (helper module)
- Create: `xlcat/tests/common/mod.rs`

We generate test xlsx files programmatically so we don't commit binaries.

- [ ] **Step 1: Create test helper that generates fixture files**

Create `xlcat/tests/common/mod.rs`:

```rust
use rust_xlsxwriter::*;
use std::path::Path;

/// Single sheet, 5 rows of mixed types: string, float, int, bool, date
pub fn create_simple(path: &Path) {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet_with_name("Data").unwrap();

    // Headers
    ws.write_string(0, 0, "name").unwrap();
    ws.write_string(0, 1, "amount").unwrap();
    ws.write_string(0, 2, "count").unwrap();
    ws.write_string(0, 3, "active").unwrap();

    // Row 1
    ws.write_string(1, 0, "Alice").unwrap();
    ws.write_number(1, 1, 100.50).unwrap();
    ws.write_number(1, 2, 10.0).unwrap();
    ws.write_boolean(1, 3, true).unwrap();

    // Row 2
    ws.write_string(2, 0, "Bob").unwrap();
    ws.write_number(2, 1, 200.75).unwrap();
    ws.write_number(2, 2, 20.0).unwrap();
    ws.write_boolean(2, 3, false).unwrap();

    // Row 3
    ws.write_string(3, 0, "Charlie").unwrap();
    ws.write_number(3, 1, 300.00).unwrap();
    ws.write_number(3, 2, 30.0).unwrap();
    ws.write_boolean(3, 3, true).unwrap();

    // Row 4
    ws.write_string(4, 0, "Diana").unwrap();
    ws.write_number(4, 1, 400.25).unwrap();
    ws.write_number(4, 2, 40.0).unwrap();
    ws.write_boolean(4, 3, false).unwrap();

    // Row 5
    ws.write_string(5, 0, "Eve").unwrap();
    ws.write_number(5, 1, 500.00).unwrap();
    ws.write_number(5, 2, 50.0).unwrap();
    ws.write_boolean(5, 3, true).unwrap();

    wb.save(path).unwrap();
}

/// 3 sheets: Revenue (4 rows), Expenses (3 rows), Summary (2 rows)
pub fn create_multi_sheet(path: &Path) {
    let mut wb = Workbook::new();

    let ws1 = wb.add_worksheet_with_name("Revenue").unwrap();
    ws1.write_string(0, 0, "region").unwrap();
    ws1.write_string(0, 1, "amount").unwrap();
    for i in 1..=4u32 {
        ws1.write_string(i, 0, &format!("Region {i}")).unwrap();
        ws1.write_number(i, 1, i as f64 * 1000.0).unwrap();
    }

    let ws2 = wb.add_worksheet_with_name("Expenses").unwrap();
    ws2.write_string(0, 0, "category").unwrap();
    ws2.write_string(0, 1, "amount").unwrap();
    for i in 1..=3u32 {
        ws2.write_string(i, 0, &format!("Category {i}")).unwrap();
        ws2.write_number(i, 1, i as f64 * 500.0).unwrap();
    }

    let ws3 = wb.add_worksheet_with_name("Summary").unwrap();
    ws3.write_string(0, 0, "metric").unwrap();
    ws3.write_string(0, 1, "value").unwrap();
    ws3.write_string(1, 0, "Total Revenue").unwrap();
    ws3.write_number(1, 1, 10000.0).unwrap();
    ws3.write_string(2, 0, "Total Expenses").unwrap();
    ws3.write_number(2, 1, 3000.0).unwrap();

    wb.save(path).unwrap();
}

/// Single sheet with 80 rows (to test head/tail adaptive behavior)
pub fn create_many_rows(path: &Path) {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet_with_name("Data").unwrap();

    ws.write_string(0, 0, "id").unwrap();
    ws.write_string(0, 1, "value").unwrap();

    for i in 1..=80u32 {
        ws.write_number(i, 0, i as f64).unwrap();
        ws.write_number(i, 1, i as f64 * 1.5).unwrap();
    }

    wb.save(path).unwrap();
}

/// Single sheet with header row but no data rows
pub fn create_empty_data(path: &Path) {
    let mut wb = Workbook::new();
    let ws = wb.add_worksheet_with_name("Empty").unwrap();
    ws.write_string(0, 0, "col_a").unwrap();
    ws.write_string(0, 1, "col_b").unwrap();
    wb.save(path).unwrap();
}

/// Completely empty sheet
pub fn create_empty_sheet(path: &Path) {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Blank").unwrap();
    wb.save(path).unwrap();
}
```

- [ ] **Step 2: Verify the helper compiles**

Run: `cargo test --no-run`
Expected: compiles (no tests to run yet).

- [ ] **Step 3: Commit**

```bash
git add xlcat/tests/
git commit -m "test: add xlsx fixture generators for integration tests"
```

---

### Task 3: Metadata Module

**Files:**
- Create: `xlcat/src/metadata.rs`
- Modify: `xlcat/src/main.rs` (add `mod metadata;`)
- Create: `xlcat/tests/test_metadata.rs`

- [ ] **Step 1: Write failing tests for metadata**

Create `xlcat/tests/test_metadata.rs`:

```rust
mod common;

use std::path::PathBuf;
use tempfile::TempDir;

// We test the binary output since metadata is an internal module.
// For unit tests, we'll add tests inside the module itself.

#[test]
fn test_simple_file_metadata_header() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    // For now, just verify the file was created and is non-empty
    assert!(path.exists());
    assert!(std::fs::metadata(&path).unwrap().len() > 0);
}
```

Run: `cargo test test_simple_file_metadata_header`
Expected: PASS (this is a sanity check that fixtures work).

- [ ] **Step 2: Implement metadata module**

Create `xlcat/src/metadata.rs`:

```rust
use anyhow::{Context, Result};
use calamine::{open_workbook_auto, Reader};
use std::path::Path;

/// Info about a single sheet (without loading data).
#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub name: String,
    pub rows: usize, // data rows including header
    pub cols: usize,
}

/// Info about the whole workbook file.
#[derive(Debug)]
pub struct FileInfo {
    pub file_size: u64,
    pub sheets: Vec<SheetInfo>,
}

/// Read metadata: file size, sheet names, and dimensions.
/// This loads each sheet's range to get dimensions, but we
/// don't keep the data in memory.
pub fn read_file_info(path: &Path) -> Result<FileInfo> {
    let file_size = std::fs::metadata(path)
        .with_context(|| format!("Cannot read file: {}", path.display()))?
        .len();

    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase());

    match ext.as_deref() {
        Some("xlsx") | Some("xls") | Some("xlsb") | Some("xlsm") => {}
        Some(other) => anyhow::bail!("Expected .xls or .xlsx file, got: .{other}"),
        None => anyhow::bail!("Expected .xls or .xlsx file, got: no extension"),
    }

    let mut workbook = open_workbook_auto(path)
        .with_context(|| format!("Cannot open workbook: {}", path.display()))?;

    let sheet_names: Vec<String> = workbook.sheet_names().to_vec();
    let mut sheets = Vec::new();

    for name in &sheet_names {
        let range = workbook
            .worksheet_range(name)
            .with_context(|| format!("Cannot read sheet: {name}"))?;
        let (rows, cols) = range.get_size();
        sheets.push(SheetInfo {
            name: name.clone(),
            rows,
            cols,
        });
    }

    Ok(FileInfo { file_size, sheets })
}

/// Format file size for display: "245 KB", "1.2 MB", etc.
pub fn format_file_size(bytes: u64) -> String {
    if bytes < 1_024 {
        format!("{bytes} B")
    } else if bytes < 1_048_576 {
        format!("{:.0} KB", bytes as f64 / 1_024.0)
    } else if bytes < 1_073_741_824 {
        format!("{:.1} MB", bytes as f64 / 1_048_576.0)
    } else {
        format!("{:.1} GB", bytes as f64 / 1_073_741_824.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_file_size() {
        assert_eq!(format_file_size(500), "500 B");
        assert_eq!(format_file_size(2_048), "2 KB");
        assert_eq!(format_file_size(1_500_000), "1.4 MB");
    }
}
```

- [ ] **Step 3: Add `mod metadata;` to main.rs**

Add at top of `xlcat/src/main.rs`:
```rust
mod metadata;
```

- [ ] **Step 4: Run tests**

Run: `cargo test`
Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add xlcat/src/metadata.rs xlcat/src/main.rs xlcat/tests/
git commit -m "feat: add metadata module for file info and sheet dimensions"
```

---

## Chunk 2: Core Reading and Formatting

### Task 4: Reader Module (calamine → polars)

**Files:**
- Create: `xlcat/src/reader.rs`
- Modify: `xlcat/src/main.rs` (add `mod reader;`)

This is the most complex module. It bridges calamine's cell data into
Polars DataFrames with proper type inference.

- [ ] **Step 1: Write failing unit tests for the reader**

These tests go inside `reader.rs` as `#[cfg(test)]` module. Write the
implementation file with just stub functions and the tests first.

Create `xlcat/src/reader.rs`:

```rust
use anyhow::{Context, Result};
use calamine::{open_workbook_auto, Data, Reader};
use polars::prelude::*;
use std::path::Path;

/// Inferred column type from scanning calamine cells.
#[derive(Debug, Clone, Copy, PartialEq)]
enum InferredType {
    Int,
    Float,
    String,
    Bool,
    DateTime,
    Empty,
}

/// Read a single sheet into a Polars DataFrame.
/// First row is treated as headers.
pub fn read_sheet(path: &Path, sheet_name: &str) -> Result<DataFrame> {
    let mut workbook = open_workbook_auto(path)
        .with_context(|| format!("Cannot open workbook: {}", path.display()))?;

    let range = workbook
        .worksheet_range(sheet_name)
        .with_context(|| format!("Cannot read sheet: {sheet_name}"))?;

    range_to_dataframe(&range)
}

/// Convert a calamine Range into a Polars DataFrame.
/// First row is treated as column headers.
fn range_to_dataframe(range: &calamine::Range<Data>) -> Result<DataFrame> {
    let (total_rows, cols) = range.get_size();
    if total_rows == 0 || cols == 0 {
        return Ok(DataFrame::default());
    }

    let rows: Vec<&[Data]> = range.rows().collect();

    // First row = headers
    let headers: Vec<String> = rows[0]
        .iter()
        .enumerate()
        .map(|(i, cell)| match cell {
            Data::String(s) => s.clone(),
            _ => format!("column_{i}"),
        })
        .collect();

    if total_rows == 1 {
        // Header only, no data
        let series: Vec<Column> = headers
            .iter()
            .map(|name| {
                Series::new_empty(name.into(), &DataType::Null).into_column()
            })
            .collect();
        return DataFrame::new(series).map_err(Into::into);
    }

    let data_rows = &rows[1..];
    let mut columns: Vec<Column> = Vec::with_capacity(cols);

    for col_idx in 0..cols {
        let cells: Vec<&Data> = data_rows.iter().map(|row| {
            if col_idx < row.len() { &row[col_idx] } else { &Data::Empty }
        }).collect();

        let col_type = infer_column_type(&cells);
        let series = build_series(&headers[col_idx], &cells, col_type)?;
        columns.push(series.into_column());
    }

    DataFrame::new(columns).map_err(Into::into)
}

fn infer_column_type(cells: &[&Data]) -> InferredType {
    let mut has_int = false;
    let mut has_float = false;
    let mut has_string = false;
    let mut has_bool = false;
    let mut has_datetime = false;

    for cell in cells {
        match cell {
            Data::Int(_) => has_int = true,
            Data::Float(_) => has_float = true,
            Data::String(_) => has_string = true,
            Data::Bool(_) => has_bool = true,
            Data::DateTime(_) | Data::DateTimeIso(_) => has_datetime = true,
            Data::Empty | Data::Error(_) => {}
            Data::Duration(_) | Data::DurationIso(_) => has_float = true,
        }
    }

    if has_string {
        return InferredType::String; // String trumps all
    }
    if has_datetime && !has_int && !has_float && !has_bool {
        return InferredType::DateTime;
    }
    if has_bool && !has_int && !has_float {
        return InferredType::Bool;
    }
    if has_float {
        return InferredType::Float;
    }
    if has_int {
        return InferredType::Int;
    }
    InferredType::Empty
}

fn build_series(name: &str, cells: &[&Data], col_type: InferredType) -> Result<Series> {
    let pname = PlSmallStr::from(name);
    match col_type {
        InferredType::Int => {
            let values: Vec<Option<i64>> = cells
                .iter()
                .map(|c| match c {
                    Data::Int(v) => Some(*v),
                    Data::Float(v) => Some(*v as i64),
                    Data::Empty | Data::Error(_) => None,
                    _ => None,
                })
                .collect();
            Ok(Series::new(pname, &values))
        }
        InferredType::Float => {
            let values: Vec<Option<f64>> = cells
                .iter()
                .map(|c| match c {
                    Data::Float(v) => Some(*v),
                    Data::Int(v) => Some(*v as f64),
                    Data::Empty | Data::Error(_) => None,
                    _ => None,
                })
                .collect();
            Ok(Series::new(pname, &values))
        }
        InferredType::String => {
            let values: Vec<Option<String>> = cells
                .iter()
                .map(|c| match c {
                    Data::String(s) => Some(s.clone()),
                    Data::Int(v) => Some(v.to_string()),
                    Data::Float(v) => Some(v.to_string()),
                    Data::Bool(b) => Some(b.to_string()),
                    Data::Empty | Data::Error(_) => None,
                    _ => Some(format!("{c:?}")),
                })
                .collect();
            Ok(Series::new(pname, &values))
        }
        InferredType::Bool => {
            let values: Vec<Option<bool>> = cells
                .iter()
                .map(|c| match c {
                    Data::Bool(b) => Some(*b),
                    Data::Empty | Data::Error(_) => None,
                    _ => None,
                })
                .collect();
            Ok(Series::new(pname, &values))
        }
        InferredType::DateTime => {
            // Store as f64 (Excel serial dates) for now
            let values: Vec<Option<f64>> = cells
                .iter()
                .map(|c| match c {
                    Data::DateTime(v) => Some(*v),
                    Data::Empty | Data::Error(_) => None,
                    _ => None,
                })
                .collect();
            Ok(Series::new(pname, &values))
        }
        InferredType::Empty => {
            let values: Vec<Option<f64>> = vec![None; cells.len()];
            Ok(Series::new(pname, &values))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::Data;

    fn make_range(headers: Vec<&str>, rows: Vec<Vec<Data>>) -> calamine::Range<Data> {
        let ncols = headers.len();
        let nrows = rows.len() + 1; // +1 for header
        let mut range = calamine::Range::new((0, 0), ((nrows - 1) as u32, (ncols - 1) as u32));

        for (j, h) in headers.iter().enumerate() {
            range.set_value((0, j as u32), Data::String(h.to_string()));
        }
        for (i, row) in rows.iter().enumerate() {
            for (j, cell) in row.iter().enumerate() {
                range.set_value(((i + 1) as u32, j as u32), cell.clone());
            }
        }
        range
    }

    #[test]
    fn test_infer_int_column() {
        let cells = vec![
            &Data::Int(1),
            &Data::Int(2),
            &Data::Int(3),
        ];
        assert_eq!(infer_column_type(&cells), InferredType::Int);
    }

    #[test]
    fn test_infer_float_when_mixed_int_float() {
        let cells = vec![
            &Data::Int(1),
            &Data::Float(2.5),
        ];
        assert_eq!(infer_column_type(&cells), InferredType::Float);
    }

    #[test]
    fn test_infer_string_trumps_all() {
        let cells = vec![
            &Data::Int(1),
            &Data::String("hello".to_string()),
        ];
        assert_eq!(infer_column_type(&cells), InferredType::String);
    }

    #[test]
    fn test_range_to_dataframe_basic() {
        let range = make_range(
            vec!["name", "value"],
            vec![
                vec![Data::String("Alice".to_string()), Data::Float(100.0)],
                vec![Data::String("Bob".to_string()), Data::Float(200.0)],
            ],
        );
        let df = range_to_dataframe(&range).unwrap();
        assert_eq!(df.shape(), (2, 2));
        assert_eq!(df.get_column_names(), &["name", "value"]);
    }

    #[test]
    fn test_empty_range() {
        // calamine::Range::new((0,0),(0,0)) creates a 1x1 range, not empty.
        // Use an empty Range via Default or by creating one with inverted bounds.
        let range: calamine::Range<Data> = Default::default();
        let df = range_to_dataframe(&range).unwrap();
        assert_eq!(df.shape().0, 0);
    }
}
```

- [ ] **Step 2: Add `mod reader;` to main.rs**

- [ ] **Step 3: Run tests**

Run: `cargo test`
Expected: all tests pass. The `make_range` helper constructs ranges
in-memory without needing a file.

Note: `calamine::Range::new()` and `set_value()` are public API. If
they don't exist in the version we pin, adjust to construct ranges
differently or use file-based tests instead.

- [ ] **Step 4: Commit**

```bash
git add xlcat/src/reader.rs xlcat/src/main.rs
git commit -m "feat: add reader module — calamine to polars DataFrame bridge"
```

---

### Task 5: Formatter Module

**Files:**
- Create: `xlcat/src/formatter.rs`
- Modify: `xlcat/src/main.rs` (add `mod formatter;`)

- [ ] **Step 1: Implement formatter with inline tests**

Create `xlcat/src/formatter.rs`:

```rust
use crate::metadata::{FileInfo, SheetInfo};
use polars::prelude::*;
use std::fmt::Write;

/// Format the file metadata header.
pub fn format_header(file_name: &str, info: &FileInfo) -> String {
    let size = crate::metadata::format_file_size(info.file_size);
    let mut out = String::new();
    writeln!(out, "# File: {file_name} ({size})").unwrap();
    writeln!(out, "# Sheets: {}", info.sheets.len()).unwrap();
    out
}

/// Format a sheet's schema (column names and types).
pub fn format_schema(sheet: &SheetInfo, df: &DataFrame) -> String {
    let data_rows = if sheet.rows > 0 { sheet.rows - 1 } else { 0 };
    let mut out = String::new();
    writeln!(out, "## Sheet: {} ({} rows x {} cols)", sheet.name, data_rows, sheet.cols).unwrap();
    writeln!(out).unwrap();

    // Column type table
    writeln!(out, "| Column | Type |").unwrap();
    writeln!(out, "|--------|------|").unwrap();
    for col in df.get_columns() {
        writeln!(out, "| {} | {} |", col.name(), col.dtype()).unwrap();
    }
    out
}

/// Format the sheet listing for multi-sheet files (no data, just schemas).
pub fn format_sheet_listing(file_name: &str, info: &FileInfo, schemas: &[(&SheetInfo, DataFrame)]) -> String {
    let mut out = format_header(file_name, info);
    writeln!(out).unwrap();

    for (sheet, df) in schemas {
        out.push_str(&format_schema(sheet, df));
        writeln!(out).unwrap();
    }

    writeln!(out, "Use --sheet <name|index> to view data.").unwrap();
    out
}

/// Format a DataFrame as a markdown table.
pub fn format_data_table(df: &DataFrame) -> String {
    let mut out = String::new();
    let cols = df.get_columns();
    if cols.is_empty() {
        return out;
    }

    // Header row
    let header: Vec<String> = cols.iter().map(|c| c.name().to_string()).collect();
    writeln!(out, "| {} |", header.join(" | ")).unwrap();

    // Separator
    let seps: Vec<String> = header.iter().map(|h| "-".repeat(h.len().max(3))).collect();
    writeln!(out, "| {} |", seps.join(" | ")).unwrap();

    // Data rows
    let height = df.height();
    for i in 0..height {
        let row: Vec<String> = cols
            .iter()
            .map(|c| format_cell(c, i))
            .collect();
        writeln!(out, "| {} |", row.join(" | ")).unwrap();
    }

    out
}

/// Format data table showing head + tail with omission separator.
pub fn format_head_tail(
    df: &DataFrame,
    head_n: usize,
    tail_n: usize,
) -> String {
    let total = df.height();
    if total <= head_n + tail_n {
        return format_data_table(df);
    }

    let head_df = df.head(Some(head_n));
    let tail_df = df.tail(Some(tail_n));
    let omitted = total - head_n - tail_n;

    let mut out = String::new();
    let cols = df.get_columns();

    // Header row
    let header: Vec<String> = cols.iter().map(|c| c.name().to_string()).collect();
    writeln!(out, "| {} |", header.join(" | ")).unwrap();
    let seps: Vec<String> = header.iter().map(|h| "-".repeat(h.len().max(3))).collect();
    writeln!(out, "| {} |", seps.join(" | ")).unwrap();

    // Head rows
    for i in 0..head_df.height() {
        let row: Vec<String> = head_df.get_columns().iter().map(|c| format_cell(c, i)).collect();
        writeln!(out, "| {} |", row.join(" | ")).unwrap();
    }

    // Omission line
    writeln!(out, "... ({omitted} rows omitted) ...").unwrap();

    // Tail rows
    for i in 0..tail_df.height() {
        let row: Vec<String> = tail_df.get_columns().iter().map(|c| format_cell(c, i)).collect();
        writeln!(out, "| {} |", row.join(" | ")).unwrap();
    }

    out
}

/// Format a DataFrame as CSV.
pub fn format_csv(df: &DataFrame) -> String {
    let mut buf = Vec::new();
    CsvWriter::new(&mut buf).finish(&mut df.clone()).unwrap();
    String::from_utf8(buf).unwrap()
}

/// Format a single cell value for display.
fn format_cell(col: &Column, idx: usize) -> String {
    let val = col.get(idx);
    match val {
        Ok(AnyValue::Null) => String::new(),
        Ok(v) => v.to_string(),
        Err(_) => String::new(),
    }
}

/// Format the empty-sheet message.
pub fn format_empty_sheet(sheet: &SheetInfo) -> String {
    if sheet.rows == 0 && sheet.cols == 0 {
        format!("## Sheet: {} (empty)\n", sheet.name)
    } else {
        format!("## Sheet: {} (no data rows)\n", sheet.name)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::metadata::{FileInfo, SheetInfo};

    #[test]
    fn test_format_header() {
        let info = FileInfo {
            file_size: 250_000,
            sheets: vec![SheetInfo {
                name: "Sheet1".into(),
                rows: 100,
                cols: 5,
            }],
        };
        let out = format_header("test.xlsx", &info);
        assert!(out.contains("# File: test.xlsx (244 KB)"));
        assert!(out.contains("# Sheets: 1"));
    }

    #[test]
    fn test_format_data_table() {
        let s1 = Series::new("name".into(), &["Alice", "Bob"]);
        let s2 = Series::new("value".into(), &[100i64, 200]);
        let df = DataFrame::new(vec![s1.into_column(), s2.into_column()]).unwrap();
        let out = format_data_table(&df);
        assert!(out.contains("| name | value |"));
        assert!(out.contains("| Alice | 100 |"));
    }

    #[test]
    fn test_format_head_tail_small() {
        // When total rows <= head + tail, show all
        let s = Series::new("x".into(), &[1i64, 2, 3]);
        let df = DataFrame::new(vec![s.into_column()]).unwrap();
        let out = format_head_tail(&df, 25, 25);
        assert!(!out.contains("omitted"));
        assert!(out.contains("| 1 |"));
        assert!(out.contains("| 3 |"));
    }
}
```

- [ ] **Step 2: Add `mod formatter;` to main.rs**

- [ ] **Step 3: Run tests**

Run: `cargo test`
Expected: all tests pass.

- [ ] **Step 4: Commit**

```bash
git add xlcat/src/formatter.rs xlcat/src/main.rs
git commit -m "feat: add formatter module — markdown table and CSV output"
```

---

## Chunk 3: Wiring and Modes

### Task 6: Wire Up Main — Data and Schema Modes

**Files:**
- Modify: `xlcat/src/main.rs`
- Create: `xlcat/tests/test_integration.rs`

- [ ] **Step 1: Write integration tests**

Create `xlcat/tests/test_integration.rs`:

```rust
mod common;

use assert_cmd::Command;
use predicates::prelude::*;
use tempfile::TempDir;

fn xlcat() -> Command {
    Command::cargo_bin("xlcat").unwrap()
}

#[test]
fn test_simple_file_default() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("# File: simple.xlsx"))
        .stdout(predicate::str::contains("# Sheets: 1"))
        .stdout(predicate::str::contains("## Sheet: Data"))
        .stdout(predicate::str::contains("| name |"))
        .stdout(predicate::str::contains("| Alice |"));
}

#[test]
fn test_schema_mode() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--schema")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(stdout.contains("| Column | Type |"));
    assert!(stdout.contains("| name |"));
    // Schema mode should NOT contain data rows
    assert!(!stdout.contains("| Alice |"));
}

#[test]
fn test_multi_sheet_default_lists_schemas() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("# Sheets: 3"))
        .stdout(predicate::str::contains("## Sheet: Revenue"))
        .stdout(predicate::str::contains("## Sheet: Expenses"))
        .stdout(predicate::str::contains("## Sheet: Summary"))
        .stdout(predicate::str::contains("Use --sheet"));
}

#[test]
fn test_multi_sheet_select_by_name() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--sheet")
        .arg("Revenue")
        .assert()
        .success()
        .stdout(predicate::str::contains("| region |"))
        .stdout(predicate::str::contains("| Region 1 |"));
}

#[test]
fn test_multi_sheet_select_by_index() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--sheet")
        .arg("1")
        .assert()
        .success()
        .stdout(predicate::str::contains("## Sheet: Expenses"));
}

#[test]
fn test_head_tail_adaptive() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("rows omitted"));
}

#[test]
fn test_head_flag() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--head")
        .arg("3")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    // Should have header + 3 data rows, no omission
    assert!(!stdout.contains("omitted"));
}

#[test]
fn test_csv_mode() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--csv")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    // CSV mode: no markdown headers
    assert!(!stdout.contains("# File:"));
    assert!(stdout.contains("name,"));
}

#[test]
fn test_invalid_flag_combo() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--schema")
        .arg("--head")
        .arg("10")
        .assert()
        .code(2)
        .stderr(predicate::str::contains("cannot be combined"));
}

#[test]
fn test_file_not_found() {
    xlcat()
        .arg("/nonexistent/file.xlsx")
        .assert()
        .failure()
        .stderr(predicate::str::contains("Cannot"));
}

#[test]
fn test_empty_sheet() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("empty.xlsx");
    common::create_empty_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("empty"));
}

#[test]
fn test_all_without_sheet_on_multi() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--all")
        .assert()
        .failure()
        .stderr(predicate::str::contains("Multiple sheets"));
}
```

- [ ] **Step 2: Implement main.rs orchestration**

Replace the stub `main()` in `xlcat/src/main.rs` with full orchestration:

```rust
use anyhow::{Context, Result};
use clap::Parser;
use std::path::PathBuf;

mod formatter;
mod metadata;
mod reader;

#[derive(Parser, Debug)]
#[command(name = "xlcat", about = "View Excel files in the terminal")]
struct Cli {
    /// Path to .xls or .xlsx file
    file: PathBuf,

    /// Show only column names and types
    #[arg(long)]
    schema: bool,

    /// Show summary statistics
    #[arg(long)]
    describe: bool,

    /// Show first N rows
    #[arg(long)]
    head: Option<usize>,

    /// Show last N rows
    #[arg(long)]
    tail: Option<usize>,

    /// Show all rows (overrides large-file gate)
    #[arg(long)]
    all: bool,

    /// Select sheet by name or 0-based index
    #[arg(long)]
    sheet: Option<String>,

    /// Large-file threshold (default: 1M). Accepts: 500K, 1M, 10M, 1G
    #[arg(long, default_value = "1M", value_parser = parse_size)]
    max_size: u64,

    /// Output as CSV instead of markdown
    #[arg(long)]
    csv: bool,
}

fn parse_size(s: &str) -> Result<u64, String> {
    let s = s.trim();
    let (num_part, multiplier) = if s.ends_with('G') || s.ends_with('g') {
        (&s[..s.len() - 1], 1_073_741_824u64)
    } else if s.ends_with("GB") || s.ends_with("gb") {
        (&s[..s.len() - 2], 1_073_741_824u64)
    } else if s.ends_with('M') || s.ends_with('m') {
        (&s[..s.len() - 1], 1_048_576u64)
    } else if s.ends_with("MB") || s.ends_with("mb") {
        (&s[..s.len() - 2], 1_048_576u64)
    } else if s.ends_with('K') || s.ends_with('k') {
        (&s[..s.len() - 1], 1_024u64)
    } else if s.ends_with("KB") || s.ends_with("kb") {
        (&s[..s.len() - 2], 1_024u64)
    } else {
        (s, 1u64)
    };
    let num: f64 = num_part.parse().map_err(|_| format!("Invalid size: {s}"))?;
    Ok((num * multiplier as f64) as u64)
}

fn run(cli: &Cli) -> Result<()> {
    // Validate flag combinations (exit code 2 errors)
    let mode_count = cli.schema as u8 + cli.describe as u8;
    if mode_count > 1 {
        return Err(ArgError("--schema and --describe are mutually exclusive".into()).into());
    }
    if (cli.schema || cli.describe)
        && (cli.head.is_some() || cli.tail.is_some() || cli.all)
    {
        return Err(ArgError(
            "--schema and --describe cannot be combined with --head, --tail, or --all".into(),
        ).into());
    }
    if (cli.schema || cli.describe) && cli.csv {
        return Err(ArgError(
            "--csv can only be used in data mode (not with --schema or --describe)".into(),
        ).into());
    }

    let info = metadata::read_file_info(&cli.file)?;
    let file_name = cli.file.file_name().unwrap().to_string_lossy();

    // Resolve which sheet(s) to operate on
    let target_sheet = resolve_sheet(&cli, &info)?;

    match target_sheet {
        SheetTarget::Single(sheet_idx) => {
            let sheet = &info.sheets[sheet_idx];
            let df = reader::read_sheet(&cli.file, &sheet.name)?;

            if cli.csv {
                // Apply row selection before CSV output
                let selected = apply_row_selection(&cli, &info, &df);
                print!("{}", formatter::format_csv(&selected));
                return Ok(());
            }

            let mut out = formatter::format_header(&file_name, &info);

            if sheet.rows == 0 && sheet.cols == 0 {
                out.push_str(&formatter::format_empty_sheet(sheet));
            } else if df.height() == 0 {
                out.push_str(&formatter::format_schema(sheet, &df));
                out.push_str("\n(no data rows)\n");
            } else if cli.schema {
                out.push_str(&formatter::format_schema(sheet, &df));
            } else if cli.describe {
                out.push_str(&formatter::format_schema(sheet, &df));
                out.push('\n');
                out.push_str(&formatter::format_describe(&df));
            } else {
                out.push_str(&formatter::format_schema(sheet, &df));
                out.push('\n');
                out.push_str(&format_data_with_selection(&cli, &info, &df));
            }

            print!("{out}");
        }
        SheetTarget::ListAll => {
            // Multi-sheet: load all sheets
            let mut sheet_dfs = Vec::new();
            for sheet in &info.sheets {
                if sheet.rows == 0 && sheet.cols == 0 {
                    sheet_dfs.push((sheet, polars::prelude::DataFrame::default()));
                } else {
                    let df = reader::read_sheet(&cli.file, &sheet.name)?;
                    sheet_dfs.push((sheet, df));
                }
            }

            if cli.describe {
                // --describe on multi-sheet: describe each sheet
                let mut out = formatter::format_header(&file_name, &info);
                for (sheet, df) in &sheet_dfs {
                    out.push('\n');
                    out.push_str(&formatter::format_schema(sheet, df));
                    if df.height() > 0 {
                        out.push('\n');
                        out.push_str(&formatter::format_describe(df));
                    }
                }
                print!("{out}");
            } else {
                // Default: list schemas only
                let refs: Vec<(&metadata::SheetInfo, polars::prelude::DataFrame)> =
                    sheet_dfs.into_iter().collect();
                let out = formatter::format_sheet_listing(
                    &file_name,
                    &info,
                    &refs.iter().map(|(s, d)| (*s, d.clone())).collect::<Vec<_>>(),
                );
                print!("{out}");
            }
        }
    }

    Ok(())
}

enum SheetTarget {
    Single(usize),
    ListAll,
}

fn resolve_sheet(cli: &Cli, info: &metadata::FileInfo) -> Result<SheetTarget> {
    if let Some(ref sheet_arg) = cli.sheet {
        // Try as index first
        if let Ok(idx) = sheet_arg.parse::<usize>() {
            if idx < info.sheets.len() {
                return Ok(SheetTarget::Single(idx));
            }
            anyhow::bail!(
                "Sheet index {idx} out of range. Available sheets: {}",
                info.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>().join(", ")
            );
        }
        // Try as name
        if let Some(idx) = info.sheets.iter().position(|s| s.name == *sheet_arg) {
            return Ok(SheetTarget::Single(idx));
        }
        anyhow::bail!(
            "Sheet '{}' not found. Available sheets: {}",
            sheet_arg,
            info.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>().join(", ")
        );
    }

    if info.sheets.len() == 1 {
        Ok(SheetTarget::Single(0))
    } else {
        // Multi-sheet: check for flags that require a single sheet
        if cli.all || cli.head.is_some() || cli.tail.is_some() {
            return Err(ArgError(
                "Multiple sheets found. Use --sheet to select one, or --schema to see all.".into(),
            ).into());
        }
        if cli.csv {
            return Err(ArgError(
                "Multiple sheets found. Use --sheet to select one for CSV output.".into(),
            ).into());
        }
        Ok(SheetTarget::ListAll)
    }
}

/// Apply row selection and return the resulting DataFrame.
/// Used by --csv mode to respect --head/--tail/--all.
fn apply_row_selection(
    cli: &Cli,
    _info: &metadata::FileInfo,
    df: &polars::prelude::DataFrame,
) -> polars::prelude::DataFrame {
    if cli.all {
        return df.clone();
    }
    let total = df.height();
    let has_explicit_head = cli.head.is_some();
    let has_explicit_tail = cli.tail.is_some();

    if has_explicit_head && has_explicit_tail {
        let head_n = cli.head.unwrap();
        let tail_n = cli.tail.unwrap();
        if head_n + tail_n >= total {
            return df.clone();
        }
        let head_df = df.head(Some(head_n));
        let tail_df = df.tail(Some(tail_n));
        head_df.vstack(&tail_df).unwrap()
    } else if has_explicit_head {
        df.head(Some(cli.head.unwrap()))
    } else if has_explicit_tail {
        df.tail(Some(cli.tail.unwrap()))
    } else {
        df.clone()
    }
}

fn format_data_with_selection(
    cli: &Cli,
    info: &metadata::FileInfo,
    df: &polars::prelude::DataFrame,
) -> String {
    let total = df.height();

    // Explicit flags
    if cli.all {
        return formatter::format_data_table(df);
    }

    let has_explicit_head = cli.head.is_some();
    let has_explicit_tail = cli.tail.is_some();

    if has_explicit_head || has_explicit_tail {
        let head_n = cli.head.unwrap_or(0);
        let tail_n = cli.tail.unwrap_or(0);

        if head_n + tail_n >= total {
            return formatter::format_data_table(df);
        }

        if has_explicit_head && has_explicit_tail {
            return formatter::format_head_tail(df, head_n, tail_n);
        } else if has_explicit_head {
            let head_df = df.head(Some(head_n));
            return formatter::format_data_table(&head_df);
        } else {
            let tail_df = df.tail(Some(tail_n));
            return formatter::format_data_table(&tail_df);
        }
    }

    // Large-file gate (no explicit row flags)
    if info.file_size > cli.max_size {
        let head_df = df.head(Some(25));
        let mut out = formatter::format_data_table(&head_df);
        out.push_str(&format!(
            "\nShowing first 25 rows. Use --head N / --tail N / --all for more.\n"
        ));
        return out;
    }

    // Adaptive default
    if total <= 50 {
        formatter::format_data_table(df)
    } else {
        formatter::format_head_tail(df, 25, 25)
    }
}

/// Errors that represent invalid argument combinations (exit code 2).
#[derive(Debug)]
struct ArgError(String);
impl std::fmt::Display for ArgError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", self.0)
    }
}
impl std::error::Error for ArgError {}

fn main() {
    let cli = Cli::parse();
    if let Err(e) = run(&cli) {
        eprintln!("xlcat: {e}");
        // Exit code 2 for argument validation errors, 1 for runtime errors
        if e.downcast_ref::<ArgError>().is_some() {
            std::process::exit(2);
        }
        std::process::exit(1);
    }
}
```

Note: This references `formatter::format_describe` which doesn't exist yet.
Add a stub to `formatter.rs` for now:

```rust
pub fn format_describe(_df: &DataFrame) -> String {
    "(describe not yet implemented)\n".to_string()
}
```

Also adjust `format_sheet_listing` signature to take owned tuples:

```rust
pub fn format_sheet_listing(
    file_name: &str,
    info: &FileInfo,
    schemas: &[(&SheetInfo, DataFrame)],
) -> String {
```

- [ ] **Step 3: Run tests**

Run: `cargo test`
Expected: integration tests pass. Some may need adjustment based on
exact output formatting.

- [ ] **Step 4: Fix any test failures and iterate**

Common issues to watch for:
- File name in output: `format_header` receives just the filename, not path.
- Column separator widths in markdown tables.
- Exact match on `# Sheets: 1` vs `# Sheets: 1\n`.

- [ ] **Step 5: Commit**

```bash
git add xlcat/
git commit -m "feat: wire up main orchestration — data, schema, multi-sheet modes"
```

---

### Task 7: Describe Mode

**Files:**
- Modify: `xlcat/src/formatter.rs` (replace stub)
- Add test to: `xlcat/tests/test_integration.rs`

- [ ] **Step 1: Add integration test for describe**

Add to `xlcat/tests/test_integration.rs`:

```rust
#[test]
fn test_describe_mode() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--describe")
        .assert()
        .success()
        .stdout(predicate::str::contains("count"))
        .stdout(predicate::str::contains("mean"));
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cargo test test_describe_mode`
Expected: FAIL — output contains "(describe not yet implemented)".

- [ ] **Step 3: Implement format_describe**

Replace the stub in `xlcat/src/formatter.rs`:

```rust
/// Format summary statistics for each column.
pub fn format_describe(df: &DataFrame) -> String {
    use polars::prelude::*;

    let mut out = String::new();
    let cols = df.get_columns();
    if cols.is_empty() {
        return out;
    }

    // Build stats rows
    let stat_names = ["count", "null_count", "mean", "std", "min", "max", "median", "unique"];

    // Header
    let mut header = vec!["stat".to_string()];
    header.extend(cols.iter().map(|c| c.name().to_string()));
    writeln!(out, "| {} |", header.join(" | ")).unwrap();
    let seps: Vec<String> = header.iter().map(|h| "-".repeat(h.len().max(3))).collect();
    writeln!(out, "| {} |", seps.join(" | ")).unwrap();

    for stat in &stat_names {
        let mut row = vec![stat.to_string()];
        for col in cols {
            let val = compute_stat(col, stat);
            row.push(val);
        }
        writeln!(out, "| {} |", row.join(" | ")).unwrap();
    }

    out
}

fn compute_stat(col: &Column, stat: &str) -> String {
    let series = col.as_materialized_series();
    let len = series.len();

    match stat {
        "count" => len.to_string(),
        "null_count" => series.null_count().to_string(),
        "mean" => {
            if series.dtype().is_numeric() {
                series
                    .mean()
                    .map(|v| format!("{v:.4}"))
                    .unwrap_or_else(|| "-".into())
            } else {
                "-".into()
            }
        }
        "std" => {
            if series.dtype().is_numeric() {
                series
                    .std(1) // ddof=1
                    .map(|v| format!("{v:.4}"))
                    .unwrap_or_else(|| "-".into())
            } else {
                "-".into()
            }
        }
        "min" => {
            if series.dtype().is_numeric() {
                match series.min_reduce() {
                    Ok(v) => v.value().to_string(),
                    Err(_) => "-".into(),
                }
            } else {
                "-".into()
            }
        }
        "max" => {
            if series.dtype().is_numeric() {
                match series.max_reduce() {
                    Ok(v) => v.value().to_string(),
                    Err(_) => "-".into(),
                }
            } else {
                "-".into()
            }
        }
        "median" => {
            if series.dtype().is_numeric() {
                series
                    .median()
                    .map(|v| format!("{v:.4}"))
                    .unwrap_or_else(|| "-".into())
            } else {
                "-".into()
            }
        }
        "unique" => {
            match series.n_unique() {
                Ok(n) => n.to_string(),
                Err(_) => "-".into(),
            }
        }
        _ => "-".into(),
    }
}
```

Note: `series.mean()`, `series.std()`, `series.median()` return `Option<f64>`.
`series.min_reduce()`, `series.max_reduce()` return `Result<Scalar>`.
If the Polars version doesn't have these exact signatures, adjust. Check
with `cargo doc --open` to see available methods.

- [ ] **Step 4: Run tests**

Run: `cargo test`
Expected: all tests pass including `test_describe_mode`.

- [ ] **Step 5: Commit**

```bash
git add xlcat/src/formatter.rs xlcat/tests/
git commit -m "feat: implement describe mode — summary statistics per column"
```

---

## Chunk 4: Polish, Errors, and Skill

### Task 8: Large-File Gate and Edge Cases

**Files:**
- Modify: `xlcat/tests/test_integration.rs`

The large-file gate logic is already wired in Task 6. Here we add
targeted tests to verify edge cases.

- [ ] **Step 1: Add edge case integration tests**

Add to `xlcat/tests/test_integration.rs`:

```rust
#[test]
fn test_large_file_gate_triggers() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    // Set max-size very low so gate triggers
    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--max-size")
        .arg("1K")
        .assert()
        .success()
        .stdout(predicate::str::contains("Showing first 25 rows"));
}

#[test]
fn test_large_file_gate_overridden_by_head() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--max-size")
        .arg("1K")
        .arg("--head")
        .arg("5")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(!stdout.contains("Showing first 25 rows"));
}

#[test]
fn test_large_file_gate_overridden_by_all() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--max-size")
        .arg("1K")
        .arg("--all")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(!stdout.contains("Showing first 25 rows"));
}

#[test]
fn test_empty_data_headers_only() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("empty_data.xlsx");
    common::create_empty_data(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("no data rows"));
}

#[test]
fn test_head_and_tail_together() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--head")
        .arg("3")
        .arg("--tail")
        .arg("2")
        .assert()
        .success()
        .stdout(predicate::str::contains("omitted"));
}

#[test]
fn test_wrong_extension() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("data.csv");
    std::fs::write(&path, "a,b\n1,2\n").unwrap();

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .failure()
        .stderr(predicate::str::contains("Expected .xls or .xlsx"));
}

#[test]
fn test_sheet_not_found() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--sheet")
        .arg("Nonexistent")
        .assert()
        .failure()
        .stderr(predicate::str::contains("not found"))
        .stderr(predicate::str::contains("Data")); // lists available sheets
}

#[test]
fn test_exit_code_success() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .assert()
        .code(0);
}

#[test]
fn test_exit_code_runtime_error() {
    xlcat()
        .arg("/nonexistent.xlsx")
        .assert()
        .code(1);
}

#[test]
fn test_exit_code_invalid_args() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--schema")
        .arg("--describe")
        .assert()
        .code(2);
}

#[test]
fn test_tail_flag_alone() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--tail")
        .arg("3")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(!stdout.contains("omitted"));
    // Last 3 rows of 80-row data: ids 78, 79, 80
    assert!(stdout.contains("| 80"));
}

#[test]
fn test_csv_respects_head() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("many.xlsx");
    common::create_many_rows(&path);

    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--csv")
        .arg("--head")
        .arg("3")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    // Header + 3 data rows = 4 lines
    let lines: Vec<&str> = stdout.trim().lines().collect();
    assert_eq!(lines.len(), 4);
}

#[test]
fn test_head_tail_overlap_shows_all() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("simple.xlsx");
    common::create_simple(&path);

    // 5 rows, head 3 + tail 3 = 6 > 5, so show all without duplication
    let output = xlcat()
        .arg(path.to_str().unwrap())
        .arg("--head")
        .arg("3")
        .arg("--tail")
        .arg("3")
        .assert()
        .success();

    let stdout = String::from_utf8(output.get_output().stdout.clone()).unwrap();
    assert!(!stdout.contains("omitted"));
    assert!(stdout.contains("Alice"));
    assert!(stdout.contains("Eve"));
}

#[test]
fn test_describe_multi_sheet_no_sheet_flag() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--describe")
        .assert()
        .success()
        .stdout(predicate::str::contains("## Sheet: Revenue"))
        .stdout(predicate::str::contains("## Sheet: Expenses"))
        .stdout(predicate::str::contains("count"))
        .stdout(predicate::str::contains("mean"));
}

#[test]
fn test_csv_multi_sheet_without_sheet_is_error() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("multi.xlsx");
    common::create_multi_sheet(&path);

    xlcat()
        .arg(path.to_str().unwrap())
        .arg("--csv")
        .assert()
        .failure()
        .stderr(predicate::str::contains("Multiple sheets"));
}
```

- [ ] **Step 2: Run tests, fix any failures**

Run: `cargo test`
Expected: all pass.

`clap` returns exit code 2 for its own parse errors (e.g., `--unknown`).
Our validation errors in `run()` also return exit code 2 via `ArgError`.
Runtime errors (file not found, corrupt file, etc.) return exit code 1.

- [ ] **Step 3: Commit**

```bash
git add xlcat/tests/
git commit -m "test: add edge case tests — large file gate, empty sheets, errors"
```

---

### Task 9: Install Binary and Claude Code Skill

**Files:**
- Create: Claude Code skill file (location depends on user setup)

- [ ] **Step 1: Build release binary**

Run: `cd xlcat && cargo build --release`
Expected: binary at `xlcat/target/release/xlcat`.

- [ ] **Step 2: Install to PATH**

Run: `cp xlcat/target/release/xlcat /usr/local/bin/xlcat` (or user's preferred location).
Verify: `xlcat --help`

- [ ] **Step 3: Smoke test with a real file**

Find or create a real xlsx file and test:
```bash
xlcat some_real_file.xlsx
xlcat some_real_file.xlsx --schema
xlcat some_real_file.xlsx --describe
xlcat some_real_file.xlsx --head 5 --tail 5
xlcat some_real_file.xlsx --csv | head
```

- [ ] **Step 4: Create Claude Code /xls skill**

Determine where user's Claude Code skills live. Create the skill file:

```markdown
---
name: xls
description: View and analyze Excel (.xls/.xlsx) files using xlcat
---

Use `xlcat` to examine Excel files. Run commands via the Bash tool.

## Quick reference

| Command | Purpose |
|---------|---------|
| `xlcat <file>` | Overview: metadata + schema + first/last 25 rows |
| `xlcat <file> --schema` | Column names and types only |
| `xlcat <file> --describe` | Summary statistics per column |
| `xlcat <file> --sheet <name>` | View a specific sheet |
| `xlcat <file> --head N` | First N rows |
| `xlcat <file> --tail N` | Last N rows |
| `xlcat <file> --head N --tail M` | First N + last M rows |
| `xlcat <file> --all` | All rows (overrides size limit) |
| `xlcat <file> --csv` | Raw CSV output for piping |
| `xlcat <file> --max-size 5M` | Override large-file threshold |

## Workflow

1. Start with `xlcat <file>` to see the overview.
2. For multi-sheet files, pick a sheet with `--sheet`.
3. Use `--describe` for statistical analysis.
4. Use `--head`/`--tail` to zoom into specific regions.
5. Use `--csv` when you need to pipe data to other tools.
```

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat: add Claude Code /xls skill and build release binary"
```

---

## Summary

| Task | Component | Key deliverable |
|------|-----------|----------------|
| 1 | Scaffolding | Cargo project + CLI args |
| 2 | Test fixtures | xlsx generators for tests |
| 3 | Metadata | File info + sheet dimensions |
| 4 | Reader | calamine → polars DataFrame |
| 5 | Formatter | Markdown tables + CSV output |
| 6 | Main wiring | All modes orchestrated |
| 7 | Describe | Summary statistics |
| 8 | Edge cases | Large file gate, errors, tests |
| 9 | Ship | Release binary + /xls skill |
