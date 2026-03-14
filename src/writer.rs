use std::path::Path;

use anyhow::{bail, Context, Result};

use crate::cell::{CellAssignment, CellValue};

/// Write cell assignments to an Excel workbook, preserving existing content and formatting.
///
/// Opens the workbook at `input_path`, resolves the target sheet, applies all assignments,
/// and saves the result to `output_path`.
///
/// Returns `(count_of_cells_updated, resolved_sheet_name)`.
pub fn write_cells(
    input_path: &Path,
    output_path: &Path,
    sheet_selector: &str,
    assignments: &[CellAssignment],
) -> Result<(usize, String)> {
    // Validate file extension
    let ext = input_path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("");
    match ext.to_ascii_lowercase().as_str() {
        "xlsx" | "xlsm" => {}
        "xls" => bail!(
            "legacy .xls format is not supported — please convert to .xlsx first"
        ),
        other => bail!("unsupported file extension '.{}' — expected .xlsx or .xlsm", other),
    }

    // Open workbook
    let mut book = umya_spreadsheet::reader::xlsx::read(input_path)
        .with_context(|| format!("failed to open workbook '{}'", input_path.display()))?;

    // Resolve sheet
    let sheet_count = book.get_sheet_count();
    let sheet_index = resolve_sheet_index(&book, sheet_selector, sheet_count)?;

    let sheet = book
        .get_sheet_mut(&sheet_index)
        .with_context(|| format!("failed to access sheet at index {}", sheet_index))?;

    let sheet_name = sheet.get_name().to_string();

    // Apply assignments
    for assignment in assignments {
        apply_assignment(sheet, assignment);
    }

    // Save
    umya_spreadsheet::writer::xlsx::write(&book, output_path)
        .with_context(|| format!("failed to write workbook to '{}'", output_path.display()))?;

    Ok((assignments.len(), sheet_name))
}

/// Resolve a sheet selector string to a 0-based sheet index.
///
/// - Empty string → first sheet (index 0)
/// - Try matching by name first
/// - Then try parsing as a 0-based numeric index
/// - On failure, list available sheet names in the error
fn resolve_sheet_index(
    book: &umya_spreadsheet::Spreadsheet,
    selector: &str,
    sheet_count: usize,
) -> Result<usize> {
    if selector.is_empty() {
        return Ok(0);
    }

    // Try name match
    let sheets = book.get_sheet_collection_no_check();
    for (i, ws) in sheets.iter().enumerate() {
        if ws.get_name() == selector {
            return Ok(i);
        }
    }

    // Try 0-based index
    if let Ok(idx) = selector.parse::<usize>() {
        if idx < sheet_count {
            return Ok(idx);
        }
    }

    // Build error with available names
    let names: Vec<&str> = sheets.iter().map(|ws| ws.get_name()).collect();
    bail!(
        "sheet '{}' not found — available sheets: [{}]",
        selector,
        names.join(", ")
    );
}

/// Apply a single cell assignment to a worksheet.
fn apply_assignment(
    sheet: &mut umya_spreadsheet::Worksheet,
    assignment: &CellAssignment,
) {
    let cell = sheet.get_cell_mut(assignment.cell.label.as_str());

    match &assignment.value {
        CellValue::String(s) => {
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
}

/// Convert a (year, month, day) date to an Excel serial date number.
///
/// Excel serial dates count days since January 0, 1900 (i.e., Jan 1, 1900 = 1).
/// This function accounts for the Lotus 1-2-3 bug: Excel erroneously treats 1900
/// as a leap year, so dates after Feb 28, 1900 are incremented by 1 to match
/// Excel's numbering.
fn date_to_serial(year: i32, month: u32, day: u32) -> f64 {
    // Days in each month (non-leap)
    let days_in_month = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

    let mut total_days: i64 = 0;

    // Count full years from 1900 to year-1
    for y in 1900..year {
        total_days += if is_leap_year(y) { 366 } else { 365 };
    }

    // Count full months in the target year
    for m in 1..month {
        total_days += days_in_month[m as usize] as i64;
        if m == 2 && is_leap_year(year) {
            total_days += 1;
        }
    }

    // Add days
    total_days += day as i64;

    // Excel serial: Jan 1, 1900 = 1 (not 0)
    // Lotus 1-2-3 bug: Excel thinks Feb 29, 1900 exists.
    // Dates on or after Mar 1, 1900 (serial >= 61) need +1 to compensate.
    // Feb 29, 1900 itself would be serial 60 in Excel's world (the phantom day).
    if total_days >= 60 {
        total_days += 1;
    }

    total_days as f64
}

/// Check if a year is a real leap year.
fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_date_to_serial_known_dates() {
        // Jan 1, 1900 = serial 1
        assert_eq!(date_to_serial(1900, 1, 1), 1.0);
        // Jan 1, 2024 = serial 45292
        assert_eq!(date_to_serial(2024, 1, 1), 45292.0);
    }

    #[test]
    fn test_date_to_serial_epoch_boundary() {
        // Feb 28, 1900 = serial 59 (last real date before the phantom leap day)
        assert_eq!(date_to_serial(1900, 2, 28), 59.0);
        // Mar 1, 1900 = serial 61 (after the phantom Feb 29)
        assert_eq!(date_to_serial(1900, 3, 1), 61.0);
    }

    #[test]
    fn test_date_to_serial_common_dates() {
        // Dec 31, 1899 is not representable (before epoch) — but Jan 2, 1900 = 2
        assert_eq!(date_to_serial(1900, 1, 2), 2.0);
        // Excel: 2000-01-01 = 36526
        assert_eq!(date_to_serial(2000, 1, 1), 36526.0);
    }

    #[test]
    fn test_is_leap_year() {
        assert!(!is_leap_year(1900)); // not a real leap year
        assert!(is_leap_year(2000));  // divisible by 400
        assert!(is_leap_year(2024));  // divisible by 4, not by 100
        assert!(!is_leap_year(1999)); // odd year
    }
}
