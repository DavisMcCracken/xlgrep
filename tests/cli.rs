//! The contract scripts and agents rely on: --json fields and exit codes.
//! tests/data/contacts.xlsx, sheet "Contacts": Name | Email | Phone, row 4 = Jane Smith.

use std::path::Path;
use std::process::{Command, Output};

fn xlgrep(args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_xlgrep")).args(args).output().unwrap()
}

const BOOK: &str = "tests/data/contacts.xlsx";

#[test]
fn json_hits_and_rows() {
    let out = xlgrep(&["jane", "-i", "--json", "--row", BOOK]);
    assert_eq!(out.status.code(), Some(0));
    let file = Path::new("tests").join("data").join("contacts.xlsx");
    let file = file.to_str().unwrap().replace('\\', r"\\"); // JSON-escaped on Windows
    assert_eq!(
        std::str::from_utf8(&out.stdout).unwrap(),
        format!(
            concat!(
                r#"{{"file":"{}","sheet":"Contacts","cell":"A4","value":"Jane Smith","#,
                r#""row_values":{{"A":"Jane Smith","B":"jsmith@contoso.com","C":"555-0178"}}}}"#,
                "\n"
            ),
            file
        )
    );
}

#[test]
fn exit_codes_tell_no_match_from_error() {
    let none = xlgrep(&["zzz", BOOK]);
    assert_eq!(none.status.code(), Some(1));
    assert!(none.stdout.is_empty() && none.stderr.is_empty());

    // a corrupt workbook next to a good one: matches still print, exit 2, stderr names it
    let dir = std::env::temp_dir().join(format!("xlgrep-test-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::copy(BOOK, dir.join("good.xlsx")).unwrap();
    std::fs::write(dir.join("broken.xlsx"), "not a zip").unwrap();
    let out = xlgrep(&["smith", "-i", "-l", dir.to_str().unwrap()]);
    std::fs::remove_dir_all(&dir).unwrap();

    assert_eq!(out.status.code(), Some(2));
    let good = dir.join("good.xlsx");
    assert_eq!(String::from_utf8_lossy(&out.stdout).trim(), good.to_str().unwrap());
    assert!(String::from_utf8_lossy(&out.stderr).contains("broken.xlsx"));
}

#[test]
fn folder_csvs_are_counted_not_searched() {
    // an agent asked "which files mention X" must learn the CSVs weren't looked at
    let dir = std::env::temp_dir().join(format!("xlgrep-csv-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::copy(BOOK, dir.join("good.xlsx")).unwrap();
    std::fs::write(dir.join("export.CSV"), "name\nJane Smith\n").unwrap();
    let out = xlgrep(&["smith", "-i", "-l", dir.to_str().unwrap()]);
    std::fs::remove_dir_all(&dir).unwrap();

    assert_eq!(out.status.code(), Some(0)); // a note, not an error
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(stderr.contains("not searched: 1 .csv/.tsv file"), "{stderr}");
}

#[test]
fn a_far_off_cell_is_streamed_not_allocated() {
    // far-cell.xlsx = contacts.xlsx plus "far corner" at XFD1048576, which once made calamine
    // allocate the whole sheet (512 GiB) and abort
    let out = xlgrep(&["corner", "--json", "--row", "tests/data/far-cell.xlsx"]);
    assert_eq!(out.status.code(), Some(0));
    assert!(String::from_utf8_lossy(&out.stdout).contains(r#""cell":"XFD1048576""#));
}

#[test]
fn a_mistyped_sheet_is_an_error_not_a_miss() {
    let out = xlgrep(&["jane", "--sheet", "Contcts", BOOK]);
    assert_eq!(out.status.code(), Some(2));
    assert!(String::from_utf8_lossy(&out.stderr).contains("\"Contcts\""));
    assert_eq!(xlgrep(&["jane", "-i", "--sheet", "contacts", BOOK]).status.code(), Some(0));
}

#[test]
fn templates_are_searched() {
    let dir = std::env::temp_dir().join(format!("xlgrep-templates-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::copy(BOOK, dir.join("a.xltx")).unwrap();
    std::fs::copy(BOOK, dir.join("b.xltm")).unwrap();
    let out = xlgrep(&["jane", "-i", "-c", dir.to_str().unwrap()]);
    std::fs::remove_dir_all(&dir).unwrap();
    assert_eq!(out.status.code(), Some(0));
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert!(stdout.contains("a.xltx: 1") && stdout.contains("b.xltm: 1"), "{stdout}");
}

#[test]
fn a_file_reached_twice_is_searched_once() {
    // the folder, plus the same file through `..` and, on Windows, in other case
    let mut args = vec!["jane", "-i", "-l", "tests/data", "tests/../tests/data/contacts.xlsx"];
    if cfg!(windows) {
        args.push("TESTS/DATA/CONTACTS.XLSX");
    }
    let out = xlgrep(&args);
    assert_eq!(out.status.code(), Some(0));
    let stdout = String::from_utf8_lossy(&out.stdout).to_lowercase();
    assert_eq!(stdout.matches("contacts.xlsx").count(), 1, "{stdout}");
}

#[test]
fn header_labels_values_with_row_1() {
    // --header takes its row as `--header=N`, so `jane` here is the pattern, not a row
    let out = xlgrep(&["--header", "jane", "-i", "--json", "--row", BOOK]);
    assert_eq!(out.status.code(), Some(0));
    let json = String::from_utf8_lossy(&out.stdout);
    assert!(
        json.contains(concat!(
            r#""value":"Jane Smith","header":"Name","#,
            r#""row_values":{"A":"Jane Smith","B":"jsmith@contoso.com","C":"555-0178"},"#,
            r#""row_headers":{"A":"Name","B":"Email","C":"Phone"}}"#
        )),
        "{json}"
    );
    // a hit in the header row itself has no header
    let out = xlgrep(&["Email", "--header", "--json", BOOK]);
    assert!(!String::from_utf8_lossy(&out.stdout).contains("\"header\""));
    // piped text labels the value
    let out = xlgrep(&["jsmith", "--header", BOOK]);
    assert!(String::from_utf8_lossy(&out.stdout).contains("Email: jsmith@contoso.com"));
}
