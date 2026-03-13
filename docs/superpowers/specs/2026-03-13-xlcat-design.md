# xlcat — Excel CLI Viewer for LLMs

## Purpose

A Rust CLI tool that reads `.xls` and `.xlsx` files and outputs structured,
LLM-friendly text to stdout. Paired with a Claude Code `/xls` skill that
knows how to invoke it. The goal: let Claude Code view and analyze
spreadsheets without custom Python scripting.

## CLI Interface

```
xlcat <file>                          # adaptive default (see below)
xlcat <file> --schema                 # column names + types only
xlcat <file> --describe               # summary statistics per column
xlcat <file> --head 20                # first 20 rows
xlcat <file> --tail 10                # last 10 rows
xlcat <file> --head 10 --tail 5       # first 10 + last 5 rows
xlcat <file> --all                    # full dump
xlcat <file> --sheet "Revenue"        # select sheet by name or index
xlcat <file> --max-size 5M            # override large-file threshold (default 1MB)
xlcat <file> --csv                    # output as CSV instead of markdown table
```

## Modes

The tool operates in one of three mutually exclusive modes:

| Mode         | Triggered by   | Output                          |
|--------------|----------------|---------------------------------|
| **data**     | (default)      | Metadata + schema + data rows   |
| **schema**   | `--schema`     | Metadata + schema only          |
| **describe** | `--describe`   | Metadata + summary statistics   |

Flags `--head`, `--tail`, `--all`, and `--csv` apply only in **data** mode.
Combining `--schema` or `--describe` with row-selection flags is an error.

`--sheet` works in all modes — it selects which sheet to operate on.

## Default Behavior (Data Mode, No Flags)

Every invocation starts with a metadata header:

```
# File: report.xlsx (245 KB)
# Sheets: 3
```

Then, behavior is resolved in this order of precedence:

### 1. Sheet resolution
- **`--sheet` provided:** operate on that sheet only (treat as single-sheet).
- **Single sheet in file:** operate on that sheet.
- **Multiple sheets, no `--sheet`:** print each sheet's name, dimensions,
  and column schema. No data rows. User selects a sheet with `--sheet`.
  (This takes priority over all row-selection logic below.)

### 2. Large-file gate (file size on disk)
- The `--max-size` flag sets the threshold (default: 1 MB).
- **File exceeds threshold + no explicit row flags:** print schema + first
  25 rows only, with a note:
  `Showing first 25 rows. Use --head N / --tail N / --all for more.`
- **Explicit flags (`--head`, `--tail`, `--all`) override the gate.** If
  the user asks for `--head 500` on a large file, they get 500 rows.
- `--max-size` is intentionally based on file-on-disk size (cheap to check,
  no parsing needed). It is a rough proxy, not a precise row-count limit.

### 3. Row selection (single sheet resolved, within size gate)

| Flags              | Behavior                                       |
|--------------------|-------------------------------------------------|
| (none, ≤50 rows)   | All rows                                        |
| (none, >50 rows)   | First 25 + last 25                              |
| `--head N`         | First N rows                                    |
| `--tail N`         | Last N rows                                     |
| `--head N --tail M`| First N rows + last M rows                      |
| `--all`            | All rows (also overrides large-file gate)        |

When both `--head` and `--tail` are used together and the file has fewer
rows than N + M, all rows are shown without duplication.

### `--all` on multi-sheet files without `--sheet`
`--all` without `--sheet` on a multi-sheet file is an error:
`Multiple sheets found. Use --sheet to select one, or --schema to see all.`

## Output Formats

### Default: Markdown table

```
# File: report.xlsx (245 KB)
# Sheets: 1

## Sheet: Revenue (1,240 rows x 8 cols)

| Column   | Type    |
|----------|---------|
| date     | Date    |
| region   | String  |
| amount   | Float64 |
| quarter  | String  |

| date       | region | amount  | quarter |
|------------|--------|---------|---------|
| 2024-01-01 | East   | 1234.56 | Q1      |
| 2024-01-02 | West   | 987.00  | Q1      |
| ...        | ...    | ...     | ...     |
... (1,190 rows omitted) ...
| 2024-12-30 | East   | 1100.00 | Q4      |
| 2024-12-31 | West   | 1250.75 | Q4      |
```

### `--csv`

Raw CSV to stdout. No metadata header, no schema section. Suitable for
piping to other tools.

### `--schema`

Column names and inferred types only:

```
# File: report.xlsx (245 KB)
# Sheets: 1

## Sheet: Revenue (1,240 rows x 8 cols)

| Column   | Type    |
|----------|---------|
| date     | Date    |
| region   | String  |
| amount   | Float64 |
| quarter  | String  |
```

### `--describe`

Polars `describe()` output — count, null count, mean, min, max, median,
std for numeric columns; count, null count, unique for string columns.
Rendered as a markdown table.

`--describe` operates on the full sheet data (not a row subset). On a
multi-sheet file without `--sheet`, it describes each sheet sequentially.

### Multi-sheet default output example

```
# File: budget.xlsx (320 KB)
# Sheets: 3

## Sheet: Revenue (1,240 rows x 4 cols)

| Column   | Type    |
|----------|---------|
| date     | Date    |
| region   | String  |
| amount   | Float64 |
| quarter  | String  |

## Sheet: Expenses (890 rows x 5 cols)

| Column     | Type    |
|------------|---------|
| date       | Date    |
| department | String  |
| category   | String  |
| amount     | Float64 |
| approved   | Boolean |

## Sheet: Summary (12 rows x 3 cols)

| Column   | Type    |
|----------|---------|
| quarter  | String  |
| revenue  | Float64 |
| expenses | Float64 |

Use --sheet <name|index> to view data.
```

## Sheet Selection

`--sheet` accepts a sheet name (string) or a 0-based index (integer).

```
xlcat file.xlsx --sheet "Revenue"
xlcat file.xlsx --sheet 0
```

Without `--sheet`, the adaptive multi-sheet vs. single-sheet behavior applies.

## Technology

- **Language:** Rust
- **Excel reading:** Polars with calamine backend (`polars` crate with
  `xlsx` feature)
- **CLI parsing:** `clap` (derive API)
- **No runtime dependencies** — single compiled binary
- **Note:** calamine loads the entire sheet into memory. For very large
  Excel files, memory usage will be proportional to file content regardless
  of `--head`/`--max-size` flags. This is inherent to the Excel format.
- **Output is always UTF-8.**

## Project Structure

```
xlcat/
├── Cargo.toml
├── src/
│   ├── main.rs          # CLI arg parsing (clap), orchestration
│   ├── reader.rs        # Excel reading via Polars/calamine
│   ├── formatter.rs     # Markdown table and CSV output formatting
│   └── metadata.rs      # File size checks, sheet listing
```

## Claude Code Skill

A `/xls` skill that:

1. Accepts a file path argument.
2. Invokes `xlcat` with appropriate flags.
3. For large/multi-sheet files, starts with `--schema` to orient, then
   drills into specific sheets or row ranges as needed.
4. Guides the LLM to use `--describe` for statistical overview when
   analyzing data.

## Type Inference

Column types come from Polars' inference when constructing the DataFrame
from calamine data. Polars infers types from the full column by default.
Types are displayed using Polars' type names: `String`, `Float64`, `Int64`,
`Boolean`, `Date`, `Datetime`, `Null`, etc.

For the multi-sheet listing (schema-only, no data loaded), sheet metadata
(name, dimensions) is read via calamine directly to avoid loading all
sheets into memory. Column types require reading the sheet into a DataFrame,
so the schema listing does load each sheet.

## Empty Sheets

- A sheet with 0 rows and 0 columns: print sheet name and `(empty)`.
- A sheet with a header row but 0 data rows: print the schema table,
  then `(no data rows)`.

## Exit Codes

| Code | Meaning                                      |
|------|----------------------------------------------|
| 0    | Success                                      |
| 1    | Runtime error (file not found, corrupt, etc.) |
| 2    | Invalid arguments (bad flag combination)      |

All error messages go to stderr. Stdout contains only data output.
No partial output on error — the tool fails before writing to stdout
when possible.

## Error Handling

- File not found: clear error message with path.
- Unsupported format: "Expected .xls or .xlsx file, got: <ext>"
- Corrupt/unreadable file: surface the underlying Polars/calamine error.
- Sheet not found: list available sheets in the error message.
- Invalid flag combination (e.g., `--schema --head 10`): error with usage hint.

## Future Possibilities (Not in Scope)

- SQL query mode (`--sql "SELECT ..."`)
- Filter expressions (`--filter "amount > 1000"`)
- Multiple file comparison
- ODS / Google Sheets support
