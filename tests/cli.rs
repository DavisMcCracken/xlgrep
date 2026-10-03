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
