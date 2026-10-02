//! Rust port of xlgrep for comparison. Same reader (calamine) as the Python version.

use std::io::{self, BufWriter, IsTerminal, Write};
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use calamine::{Data, DataType, Reader, open_workbook_auto};
use clap::Parser;
use regex::RegexBuilder;

const EXTS: [&str; 5] = ["xlsx", "xlsm", "xlsb", "xls", "ods"];
// Windows marks OneDrive online-only files with these; reading the content triggers a download
const FILE_ATTRIBUTE_OFFLINE: u32 = 0x1000;
const FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS: u32 = 0x400000;

/// Search spreadsheet cell values (read-only). Prints: file | sheet | cell | value.
#[derive(Parser)]
#[command(name = "xlgrep-rs")]
struct Args {
    /// regex (or literal text with -F)
    pattern: String,
    /// files or folders
    #[arg(default_value = ".")]
    paths: Vec<PathBuf>,
    /// case-insensitive match
    #[arg(short, long)]
    ignore_case: bool,
    /// literal text, not regex
    #[arg(short = 'F', long)]
    fixed_strings: bool,
    /// also search OneDrive online-only files (downloads them; skipped by default)
    #[arg(long)]
    download: bool,
    /// only print match count per file
    #[arg(short, long, conflicts_with = "files_with_matches")]
    count: bool,
    /// only print matching files
    #[arg(short = 'l', long)]
    files_with_matches: bool,
}

#[cfg(windows)]
fn is_cloud_only(meta: &std::fs::Metadata) -> bool {
    use std::os::windows::fs::MetadataExt;
    meta.file_attributes() & (FILE_ATTRIBUTE_OFFLINE | FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS) != 0
}

#[cfg(not(windows))]
fn is_cloud_only(_: &std::fs::Metadata) -> bool {
    false
}

fn is_spreadsheet(p: &Path) -> bool {
    let ext = p.extension().and_then(|e| e.to_str()).unwrap_or("").to_ascii_lowercase();
    // ~$ = Excel lock files
    EXTS.contains(&ext.as_str())
        && !p.file_name().and_then(|n| n.to_str()).is_some_and(|n| n.starts_with("~$"))
}

/// Collect (path, cloud_only) like os.walk: a folder's files (sorted), then its subfolders.
/// Rust's std handles paths past MAX_PATH itself.
fn walk(dir: &Path, out: &mut Vec<(PathBuf, bool)>) {
    let Ok(rd) = std::fs::read_dir(dir) else { return };
    let mut entries: Vec<_> = rd.filter_map(Result::ok).collect();
    entries.sort_by_key(|e| e.file_name());
    let mut subdirs = Vec::new();
    for e in entries {
        // DirEntry metadata comes from the directory listing: no file open, no download
        let Ok(meta) = e.metadata() else { continue };
        if meta.is_dir() {
            subdirs.push(e.path());
        } else if is_spreadsheet(&e.path()) {
            out.push((e.path(), is_cloud_only(&meta)));
        }
    }
    for d in subdirs {
        walk(&d, out);
    }
}

fn col_letter(mut i: usize) -> String {
    let mut s = Vec::new();
    i += 1;
    while i > 0 {
        let r = (i - 1) % 26;
        s.push(b'A' + r as u8);
        i = (i - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).unwrap()
}

fn display(p: &Path, cwd: &Path) -> String {
    let abs = std::path::absolute(p).unwrap_or_else(|_| p.to_path_buf());
    abs.strip_prefix(cwd).unwrap_or(p).display().to_string()
}

fn cell_text(cell: &Data) -> String {
    match cell {
        Data::DateTime(_) | Data::DateTimeIso(_) => {
            cell.as_datetime().map_or_else(|| cell.to_string(), |d| d.to_string())
        }
        _ => cell.to_string(),
    }
}

fn main() -> ExitCode {
    let args = Args::parse();
    let pat = if args.fixed_strings { regex::escape(&args.pattern) } else { args.pattern.clone() };
    let rx = match RegexBuilder::new(&pat).case_insensitive(args.ignore_case).build() {
        Ok(rx) => rx,
        Err(e) => {
            eprintln!("xlgrep-rs: invalid regex: {e} (use -F to search literal text)");
            return ExitCode::from(2);
        }
    };

    let mut files = Vec::new();
    for p in &args.paths {
        match std::fs::metadata(p) {
            Err(_) => eprintln!("xlgrep-rs: {}: no such file or directory", p.display()),
            Ok(m) if m.is_dir() => walk(p, &mut files),
            Ok(m) => {
                if is_spreadsheet(p) {
                    files.push((p.clone(), is_cloud_only(&m)));
                }
            }
        }
    }

    let cwd = std::env::current_dir().unwrap_or_default();
    let tty = io::stdout().is_terminal();
    let mut out = BufWriter::new(io::stdout().lock());
    let (mut hits, mut files_hit, mut cloud_skipped) = (0usize, 0usize, 0usize);

    for (path, cloud) in files {
        if cloud && !args.download {
            cloud_skipped += 1;
            continue;
        }
        let name = display(&path, &cwd);
        let mut wb = match open_workbook_auto(&path) {
            Ok(wb) => wb,
            Err(e) => {
                eprintln!("xlgrep-rs: skipped {name}: {e}");
                continue;
            }
        };
        let mut n = 0;
        'sheets: for sheet in wb.sheet_names() {
            let range = match wb.worksheet_range(&sheet) {
                Ok(r) => r,
                Err(e) => {
                    eprintln!("xlgrep-rs: skipped {name} | {sheet}: {e}");
                    continue;
                }
            };
            let (r0, c0) = range.start().unwrap_or((0, 0));
            for (r, c, cell) in range.used_cells() {
                let text = cell_text(cell);
                if text.is_empty() || !rx.is_match(&text) {
                    continue;
                }
                n += 1;
                if args.files_with_matches {
                    break 'sheets;
                }
                if !args.count {
                    let cell_ref = format!("{}{}", col_letter(c0 as usize + c), r0 as usize + r + 1);
                    let _ = writeln!(out, "{name} | {sheet} | {cell_ref} | {text}");
                }
            }
        }
        if n > 0 {
            hits += n;
            files_hit += 1;
            if args.files_with_matches {
                let _ = writeln!(out, "{name}");
            } else if args.count {
                let _ = writeln!(out, "{name}: {n}");
            }
        }
    }
    let _ = out.flush();

    if cloud_skipped > 0 {
        let s = if cloud_skipped == 1 { "" } else { "s" };
        eprintln!(
            "xlgrep-rs: skipped {cloud_skipped} OneDrive online-only file{s} (use --download to include)"
        );
    }
    if tty && !args.files_with_matches {
        eprintln!("{hits} matches in {files_hit} files");
    }
    ExitCode::from(if hits > 0 { 0 } else { 1 })
}
