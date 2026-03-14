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

// ---------------------------------------------------------------------------
// Happy path — single and multiple cell assignments
// ---------------------------------------------------------------------------

#[test]
fn test_set_single_cell() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_simple(&path);

    xlset()
        .arg(path.to_str().unwrap())
        .arg("A2=Modified")
        .assert()
        .success()
        .stderr(predicate::str::contains("updated 1 cells"));

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

// ---------------------------------------------------------------------------
// Type tag coercion
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Output file (in-place vs separate destination)
// ---------------------------------------------------------------------------

#[test]
fn test_set_with_output_file() {
    let dir = TempDir::new().unwrap();
    let source = dir.path().join("source.xlsx");
    let output = dir.path().join("output.xlsx");
    common::create_simple(&source);

    xlset()
        .arg(source.to_str().unwrap())
        .arg("--output")
        .arg(output.to_str().unwrap())
        .arg("A2=New")
        .assert()
        .success();

    // Output file must contain the new value
    xlcat()
        .arg(output.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("New"));

    // Source file must still have the original value
    xlcat()
        .arg(source.to_str().unwrap())
        .assert()
        .success()
        .stdout(predicate::str::contains("Alice"));
}

// ---------------------------------------------------------------------------
// Sheet selection
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// --from CSV
// ---------------------------------------------------------------------------

#[test]
fn test_set_from_csv() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    let csv_path = dir.path().join("updates.csv");
    common::create_simple(&path);

    std::fs::write(&csv_path, "cell,value\nA2,Updated\nB2,999\n").unwrap();

    xlset()
        .arg(path.to_str().unwrap())
        .arg("--from")
        .arg(csv_path.to_str().unwrap())
        .assert()
        .success()
        .stderr(predicate::str::contains("updated 2 cells"));
}

#[test]
fn test_set_from_csv_no_header() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    let csv_path = dir.path().join("updates.csv");
    common::create_simple(&path);

    // No header row — first line is a valid cell reference
    std::fs::write(&csv_path, "A2,Updated\nB2,999\n").unwrap();

    xlset()
        .arg(path.to_str().unwrap())
        .arg("--from")
        .arg(csv_path.to_str().unwrap())
        .assert()
        .success()
        .stderr(predicate::str::contains("updated 2 cells"));
}

// ---------------------------------------------------------------------------
// Error cases
// ---------------------------------------------------------------------------

#[test]
fn test_error_no_assignments() {
    let dir = TempDir::new().unwrap();
    let path = dir.path().join("test.xlsx");
    common::create_simple(&path);

    xlset()
        .arg(path.to_str().unwrap())
        .assert()
        .code(2)
        .stderr(predicate::str::contains("no assignments"));
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

    // ZZZZZ exceeds the maximum column XFD (index 16383)
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
    // Write a fake file with .xls extension
    std::fs::write(&path, b"PK fake content").unwrap();

    xlset()
        .arg(path.to_str().unwrap())
        .arg("A1=42")
        .assert()
        .failure()
        .stderr(predicate::str::contains("only supports .xlsx").or(predicate::str::contains(
            "legacy .xls format is not supported",
        )));
}
