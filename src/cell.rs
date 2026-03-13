use std::fmt;

// ── Data types ────────────────────────────────────────────────────────────────

/// A parsed Excel cell reference (e.g. "B10").
/// `col` and `row` are 0-based internally.
#[derive(Debug, PartialEq)]
pub struct CellRef {
    pub col: u32,
    pub row: u32,
    /// Canonical upper-case label, e.g. "B10"
    pub label: String,
}

/// A typed cell value.
#[derive(Debug, PartialEq)]
pub enum CellValue {
    String(String),
    Integer(i64),
    Float(f64),
    Bool(bool),
    Date { year: i32, month: u32, day: u32 },
    Empty,
}

/// A complete cell assignment: which cell gets which value.
#[derive(Debug, PartialEq)]
pub struct CellAssignment {
    pub cell: CellRef,
    pub value: CellValue,
}

// ── Display ───────────────────────────────────────────────────────────────────

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.label)
    }
}

// ── Column helpers ────────────────────────────────────────────────────────────

/// Convert a 0-based column index to its Excel column letters (e.g. 0→"A", 25→"Z", 26→"AA").
fn col_to_letters(mut col: u32) -> String {
    let mut letters = Vec::new();
    loop {
        letters.push((b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    letters.iter().rev().collect()
}

/// Parse the alphabetic column prefix of an A1-style reference.
/// Returns `(0-based column index, remaining string slice)` or an error.
fn parse_col_part(s: &str) -> Result<(u32, &str), String> {
    let upper = s.to_ascii_uppercase();
    let alpha_len = upper.chars().take_while(|c| c.is_ascii_alphabetic()).count();
    if alpha_len == 0 {
        return Err(format!("no column letters found in '{}'", s));
    }
    let col_str = &upper[..alpha_len];
    let rest = &s[alpha_len..];

    // Convert letters to 0-based index (Excel "bijective base-26")
    let mut col: u32 = 0;
    for ch in col_str.chars() {
        col = col * 26 + (ch as u32 - 'A' as u32 + 1);
    }
    col -= 1; // convert to 0-based

    // Max column is XFD (0-based index 16383)
    if col > 16383 {
        return Err(format!("column '{}' exceeds maximum XFD", col_str));
    }

    Ok((col, rest))
}

// ── Public parsing API ────────────────────────────────────────────────────────

/// Parse an A1-style cell reference string into a [`CellRef`].
///
/// - Column letters are case-insensitive.
/// - Row numbers are 1-based in the input, stored 0-based.
/// - Maximum column is XFD (index 16383); maximum row is 1 048 576.
pub fn parse_cell_ref(s: &str) -> Result<CellRef, String> {
    let s = s.trim();
    if s.is_empty() {
        return Err("cell reference is empty".to_string());
    }

    let (col, rest) = parse_col_part(s)?;

    if rest.is_empty() {
        return Err(format!("no row number found in '{}'", s));
    }

    let row_1based: u32 = rest
        .parse()
        .map_err(|_| format!("invalid row number '{}' in '{}'", rest, s))?;

    if row_1based == 0 {
        return Err(format!("row number must be >= 1, got 0 in '{}'", s));
    }
    if row_1based > 1_048_576 {
        return Err(format!(
            "row {} exceeds maximum 1048576 in '{}'",
            row_1based, s
        ));
    }

    let row = row_1based - 1; // convert to 0-based
    let label = format!("{}{}", col_to_letters(col), row_1based);

    Ok(CellRef { col, row, label })
}

/// Infer a [`CellValue`] from a raw string, applying automatic type detection.
///
/// Detection order:
/// 1. Empty string → [`CellValue::Empty`]
/// 2. `"true"` / `"false"` (case-insensitive) → [`CellValue::Bool`]
/// 3. Valid `i64` → [`CellValue::Integer`]
/// 4. Valid `f64` → [`CellValue::Float`]
/// 5. `YYYY-MM-DD` → [`CellValue::Date`]
/// 6. Everything else → [`CellValue::String`]
pub fn infer_value(s: &str) -> CellValue {
    if s.is_empty() {
        return CellValue::Empty;
    }

    // Bool
    match s.to_ascii_lowercase().as_str() {
        "true" => return CellValue::Bool(true),
        "false" => return CellValue::Bool(false),
        _ => {}
    }

    // Integer (must not contain a '.' to avoid "1.0" being parsed as integer)
    if !s.contains('.') {
        if let Ok(i) = s.parse::<i64>() {
            return CellValue::Integer(i);
        }
    }

    // Float
    if let Ok(f) = s.parse::<f64>() {
        return CellValue::Float(f);
    }

    // Date: YYYY-MM-DD
    if let Some(date) = try_parse_date(s) {
        return date;
    }

    CellValue::String(s.to_string())
}

fn try_parse_date(s: &str) -> Option<CellValue> {
    // Strict format: exactly YYYY-MM-DD
    let parts: Vec<&str> = s.splitn(3, '-').collect();
    if parts.len() != 3 {
        return None;
    }
    // Lengths: 4-2-2
    if parts[0].len() != 4 || parts[1].len() != 2 || parts[2].len() != 2 {
        return None;
    }
    // All must be ASCII digits
    if !parts.iter().all(|p| p.chars().all(|c| c.is_ascii_digit())) {
        return None;
    }
    let year: i32 = parts[0].parse().ok()?;
    let month: u32 = parts[1].parse().ok()?;
    let day: u32 = parts[2].parse().ok()?;
    if month < 1 || month > 12 || day < 1 || day > 31 {
        return None;
    }
    Some(CellValue::Date { year, month, day })
}

/// Force a value to a specific type based on a tag.
///
/// Supported tags: `str`, `num`, `bool`, `date`.
fn coerce_value(raw: &str, tag: &str) -> Result<CellValue, String> {
    match tag {
        "str" => Ok(CellValue::String(raw.to_string())),
        "num" => {
            // Try integer first, then float
            if !raw.contains('.') {
                if let Ok(i) = raw.parse::<i64>() {
                    return Ok(CellValue::Integer(i));
                }
            }
            raw.parse::<f64>()
                .map(CellValue::Float)
                .map_err(|_| format!("cannot coerce '{}' to num", raw))
        }
        "bool" => match raw.to_ascii_lowercase().as_str() {
            "true" | "1" | "yes" => Ok(CellValue::Bool(true)),
            "false" | "0" | "no" => Ok(CellValue::Bool(false)),
            _ => Err(format!("cannot coerce '{}' to bool", raw)),
        },
        "date" => try_parse_date(raw)
            .ok_or_else(|| format!("cannot coerce '{}' to date (expected YYYY-MM-DD)", raw)),
        other => Err(format!("unknown type tag ':{}'", other)),
    }
}

/// Parse an assignment string such as `"A1=42"` or `"B2:str=07401"`.
///
/// Format: `<cell_ref>[:<tag>]=<value>`
/// - `<tag>` is optional; if absent, the value type is inferred automatically.
/// - The split is on the **first** `=` only, so values may contain `=`.
pub fn parse_assignment(s: &str) -> Result<CellAssignment, String> {
    let eq_pos = s
        .find('=')
        .ok_or_else(|| format!("no '=' found in assignment '{}'", s))?;

    let lhs = &s[..eq_pos];
    let raw_value = &s[eq_pos + 1..];

    // Check for optional :tag in LHS
    let (cell_str, tag_opt) = if let Some(colon_pos) = lhs.rfind(':') {
        let tag = &lhs[colon_pos + 1..];
        let cell = &lhs[..colon_pos];
        (cell, Some(tag))
    } else {
        (lhs, None)
    };

    let cell = parse_cell_ref(cell_str)?;

    let value = match tag_opt {
        Some(tag) => coerce_value(raw_value, tag)?,
        None => infer_value(raw_value),
    };

    Ok(CellAssignment { cell, value })
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    // ── parse_cell_ref ────────────────────────────────────────────────────────

    #[test]
    fn test_parse_a1() {
        let r = parse_cell_ref("A1").unwrap();
        assert_eq!(r.col, 0);
        assert_eq!(r.row, 0);
        assert_eq!(r.label, "A1");
    }

    #[test]
    fn test_parse_z1_col25() {
        let r = parse_cell_ref("Z1").unwrap();
        assert_eq!(r.col, 25);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn test_parse_aa1_col26() {
        let r = parse_cell_ref("AA1").unwrap();
        assert_eq!(r.col, 26);
        assert_eq!(r.row, 0);
    }

    #[test]
    fn test_parse_case_insensitive() {
        let lower = parse_cell_ref("b5").unwrap();
        let upper = parse_cell_ref("B5").unwrap();
        assert_eq!(lower, upper);
        assert_eq!(lower.label, "B5");
    }

    #[test]
    fn test_parse_b10_row9() {
        let r = parse_cell_ref("B10").unwrap();
        assert_eq!(r.col, 1);
        assert_eq!(r.row, 9);
    }

    #[test]
    fn test_parse_invalid_no_row() {
        assert!(parse_cell_ref("A").is_err());
    }

    #[test]
    fn test_parse_invalid_no_col() {
        assert!(parse_cell_ref("123").is_err());
    }

    #[test]
    fn test_parse_invalid_empty() {
        assert!(parse_cell_ref("").is_err());
    }

    #[test]
    fn test_parse_invalid_row_zero() {
        assert!(parse_cell_ref("A0").is_err());
    }

    #[test]
    fn test_parse_invalid_row_too_large() {
        assert!(parse_cell_ref("A1048577").is_err());
    }

    #[test]
    fn test_parse_max_row() {
        let r = parse_cell_ref("A1048576").unwrap();
        assert_eq!(r.row, 1_048_575);
    }

    // ── infer_value ───────────────────────────────────────────────────────────

    #[test]
    fn test_infer_integer() {
        assert_eq!(infer_value("42"), CellValue::Integer(42));
    }

    #[test]
    fn test_infer_negative_integer() {
        assert_eq!(infer_value("-7"), CellValue::Integer(-7));
    }

    #[test]
    fn test_infer_float() {
        assert_eq!(infer_value("3.14"), CellValue::Float(3.14));
    }

    #[test]
    fn test_infer_bool_true() {
        assert_eq!(infer_value("true"), CellValue::Bool(true));
        assert_eq!(infer_value("TRUE"), CellValue::Bool(true));
    }

    #[test]
    fn test_infer_bool_false() {
        assert_eq!(infer_value("false"), CellValue::Bool(false));
        assert_eq!(infer_value("False"), CellValue::Bool(false));
    }

    #[test]
    fn test_infer_date() {
        assert_eq!(
            infer_value("2024-03-15"),
            CellValue::Date { year: 2024, month: 3, day: 15 }
        );
    }

    #[test]
    fn test_infer_string() {
        assert_eq!(
            infer_value("hello world"),
            CellValue::String("hello world".to_string())
        );
    }

    #[test]
    fn test_infer_leading_zero_becomes_integer() {
        // "07401" has no dot → parsed as i64 if it parses; but leading zeros parse fine as i64
        // The spec says "leading-zero-becomes-integer" — 07401 → Integer(7401)
        assert_eq!(infer_value("07401"), CellValue::Integer(7401));
    }

    #[test]
    fn test_infer_empty() {
        assert_eq!(infer_value(""), CellValue::Empty);
    }

    // ── parse_assignment ──────────────────────────────────────────────────────

    #[test]
    fn test_assignment_basic() {
        let a = parse_assignment("A1=42").unwrap();
        assert_eq!(a.cell, parse_cell_ref("A1").unwrap());
        assert_eq!(a.value, CellValue::Integer(42));
    }

    #[test]
    fn test_assignment_with_str_tag() {
        let a = parse_assignment("B2:str=07401").unwrap();
        assert_eq!(a.cell, parse_cell_ref("B2").unwrap());
        assert_eq!(a.value, CellValue::String("07401".to_string()));
    }

    #[test]
    fn test_assignment_no_equals_error() {
        assert!(parse_assignment("A1").is_err());
    }

    #[test]
    fn test_assignment_empty_value() {
        let a = parse_assignment("C3=").unwrap();
        assert_eq!(a.value, CellValue::Empty);
    }

    #[test]
    fn test_assignment_string_with_spaces() {
        let a = parse_assignment("D4=hello world").unwrap();
        assert_eq!(a.value, CellValue::String("hello world".to_string()));
    }

    #[test]
    fn test_assignment_value_contains_equals() {
        // Split on first '=' only — value may contain '='
        let a = parse_assignment("E5=a=b").unwrap();
        assert_eq!(a.value, CellValue::String("a=b".to_string()));
    }

    #[test]
    fn test_assignment_num_tag() {
        let a = parse_assignment("A1:num=3.14").unwrap();
        assert_eq!(a.value, CellValue::Float(3.14));
    }

    #[test]
    fn test_assignment_bool_tag() {
        let a = parse_assignment("A1:bool=true").unwrap();
        assert_eq!(a.value, CellValue::Bool(true));
    }

    #[test]
    fn test_assignment_date_tag() {
        let a = parse_assignment("A1:date=2025-01-01").unwrap();
        assert_eq!(a.value, CellValue::Date { year: 2025, month: 1, day: 1 });
    }

    // ── Display ───────────────────────────────────────────────────────────────

    #[test]
    fn test_display() {
        let r = parse_cell_ref("C7").unwrap();
        assert_eq!(format!("{}", r), "C7");
    }
}
