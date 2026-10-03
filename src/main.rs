//! xlgrep: search spreadsheet cell values across folders. Reader: calamine.

use std::borrow::Cow;
use std::fmt::Write as _;
use std::fs::File;
use std::io::{self, BufReader, BufWriter, IsTerminal, Write};
use std::path::{Path, PathBuf};
use std::process::ExitCode;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Condvar, Mutex, mpsc};

use calamine::{
    Cell, CellType, Data, DataRef, DataType, Range, Reader, Sheets, XlsxError, open_workbook_auto,
};
use clap::Parser;
use regex::Regex;

const EXTS: [&str; 5] = ["xlsx", "xlsm", "xlsb", "xls", "ods"];
const ROW_SEP: &str = " │ ";

const EXAMPLES: &str = "\
Examples:
  xlgrep smith -i                 current folder, recursive, case-insensitive
  xlgrep \"O'Brien\" C:\\dir -F      literal text
  xlgrep smith --row              whole matching row (one line per row)
  xlgrep VLOOKUP --formulas -l    workbooks whose formulas use VLOOKUP
  xlgrep 1234 -x --json -m 5      exact cell match, JSON Lines, max 5 hits per file
  xlgrep smith --json | head -50  first 50 hits overall (the search stops early)

Output: grouped by file with Sheet!A1 refs in a terminal; `file | sheet | cell | value`
when piped. That text is for reading: control characters print escaped (\\n, \\u{1b}) and
long values are cut in a terminal. Use --json for anything that parses output.
Exit code: 0 match, 1 no match, 2 error (bad arguments, or a file/folder couldn't be
searched; matches found elsewhere are still printed). Skipped OneDrive files aren't errors.";

/// Search spreadsheet cell values across folders (read-only).
#[derive(Parser)]
#[command(name = "xlgrep", version, after_help = EXAMPLES)]
#[allow(clippy::struct_excessive_bools, clippy::doc_markdown)] // doc comments are --help text
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
    /// match whole words only
    #[arg(short, long)]
    word_regexp: bool,
    /// match the whole cell value
    #[arg(short = 'x', long)]
    cell_regexp: bool,
    /// print each matching row once, in full (with --formulas: the row's formula cells)
    #[arg(long)]
    row: bool,
    /// search formula text (e.g. =SUM(A1:A9)) instead of cell values
    #[arg(long)]
    formulas: bool,
    /// only search sheets with this name (case-insensitive; repeatable)
    #[arg(long, value_name = "NAME")]
    sheet: Vec<String>,
    /// stop reading a file after NUM hits
    #[arg(short, long, value_name = "NUM")]
    max_count: Option<usize>,
    /// only print hit count per file (hits = cells, or rows with --row)
    #[arg(short, long, conflicts_with = "files_with_matches")]
    count: bool,
    /// only print matching files
    #[arg(short = 'l', long)]
    files_with_matches: bool,
    /// JSON Lines: {"file","sheet","cell","value"} (+"row_values" with --row)
    #[arg(long, conflicts_with_all = ["count", "files_with_matches"])]
    json: bool,
    /// also search OneDrive online-only files (downloads them; skipped by default)
    #[arg(long)]
    download: bool,
}

#[cfg(windows)]
fn is_cloud_only(meta: &std::fs::Metadata) -> bool {
    use std::os::windows::fs::MetadataExt;
    // OneDrive marks online-only files with these; reading the content triggers a download
    const FILE_ATTRIBUTE_OFFLINE: u32 = 0x1000;
    const FILE_ATTRIBUTE_RECALL_ON_OPEN: u32 = 0x0004_0000;
    const FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS: u32 = 0x0040_0000;
    let cloud = FILE_ATTRIBUTE_OFFLINE
        | FILE_ATTRIBUTE_RECALL_ON_OPEN
        | FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS;
    meta.file_attributes() & cloud != 0
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

/// Collect (path, `cloud_only`) like `os.walk`: a folder's files (sorted), then its subfolders.
/// Rust's std handles paths past `MAX_PATH` itself. Unreadable folders count in `errors`.
fn walk(dir: &Path, out: &mut Vec<(PathBuf, bool)>, errors: &mut usize) {
    let rd = match std::fs::read_dir(dir) {
        Ok(rd) => rd,
        Err(e) => {
            eprintln!("xlgrep: skipped {}: {e}", dir.display());
            *errors += 1;
            return;
        }
    };
    let mut entries = Vec::new();
    for e in rd {
        match e {
            Ok(e) => entries.push(e),
            Err(e) => {
                eprintln!("xlgrep: skipped an entry in {}: {e}", dir.display());
                *errors += 1;
            }
        }
    }
    entries.sort_by_key(std::fs::DirEntry::file_name);
    let mut subdirs = Vec::new();
    for e in entries {
        // DirEntry metadata comes from the directory listing: no file open, no download
        let meta = match e.metadata() {
            Ok(m) => m,
            Err(err) => {
                eprintln!("xlgrep: skipped {}: {err}", e.path().display());
                *errors += 1;
                continue;
            }
        };
        if meta.is_dir() {
            subdirs.push(e.path());
        } else if is_spreadsheet(&e.path()) {
            out.push((e.path(), is_cloud_only(&meta)));
        }
    }
    for d in subdirs {
        walk(&d, out, errors);
    }
}

/// 0-based column index -> Excel letters (0 -> A, 26 -> AA).
fn col_letter(mut i: usize) -> String {
    let mut s = Vec::new();
    i += 1;
    while i > 0 {
        s.push(b'A' + u8::try_from((i - 1) % 26).unwrap());
        i = (i - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).unwrap()
}

/// Sheet name as Excel writes it in a reference: quoted unless purely alphanumeric/underscore.
fn sheet_ref(name: &str) -> Cow<'_, str> {
    let plain = name.chars().all(|c| c.is_alphanumeric() || c == '_')
        && !name.starts_with(|c: char| c.is_ascii_digit());
    if plain { Cow::Borrowed(name) } else { Cow::Owned(format!("'{}'", name.replace('\'', "''"))) }
}

fn display(p: &Path, cwd: &Path) -> String {
    let abs = std::path::absolute(p).unwrap_or_else(|_| p.to_path_buf());
    abs.strip_prefix(cwd).unwrap_or(p).display().to_string()
}

/// How a control character prints in human output: `\n` `\r` `\t`, else `\u{1b}`-style.
fn ctl_escape(c: char) -> String {
    match c {
        '\n' => "\\n".into(),
        '\r' => "\\r".into(),
        '\t' => "\\t".into(),
        c => format!("\\u{{{:x}}}", u32::from(c)),
    }
}

/// `s` with control characters escaped, so names can't move the cursor or inject escapes.
fn visible(s: &str) -> Cow<'_, str> {
    if !s.contains(char::is_control) {
        return Cow::Borrowed(s);
    }
    Cow::Owned(s.chars().map(|c| if c.is_control() { ctl_escape(c) } else { c.into() }).collect())
}

/// Text that gets searched: cell values, or formula text with `--formulas`.
trait CellText: CellType {
    /// Borrows when possible; otherwise formats into `buf`.
    fn text<'a>(&'a self, buf: &'a mut String) -> &'a str;
}

impl CellText for Data {
    fn text<'a>(&'a self, buf: &'a mut String) -> &'a str {
        match self {
            Data::String(s) => s,
            v => value_text(v, buf),
        }
    }
}

/// What xlsx and xlsb cells stream as: shared strings stay borrowed, nothing is copied.
impl CellText for DataRef<'_> {
    fn text<'a>(&'a self, buf: &'a mut String) -> &'a str {
        match self {
            DataRef::SharedString(s) => s,
            DataRef::String(s) => s,
            v => value_text(&v.clone().into(), buf),
        }
    }
}

/// A non-string value as Excel shows it, formatted into `buf`.
fn value_text<'b>(v: &Data, buf: &'b mut String) -> &'b str {
    buf.clear();
    let _ = match v {
        Data::Empty => return "",
        Data::Bool(b) => return if *b { "TRUE" } else { "FALSE" }, // as Excel shows them
        // [h]:mm:ss cells: hours keep counting past 24, as Excel shows them
        Data::DateTime(dt) if dt.is_duration() => match dt.as_duration() {
            Some(d) => {
                let (sign, s) = if d.num_seconds() < 0 {
                    ("-", -d.num_seconds())
                } else {
                    ("", d.num_seconds())
                };
                write!(buf, "{sign}{}:{:02}:{:02}", s / 3600, s / 60 % 60, s % 60)
            }
            None => write!(buf, "{v}"),
        },
        Data::DateTime(_) | Data::DateTimeIso(_) => match v.as_datetime() {
            Some(d) if d.time() == Default::default() => write!(buf, "{}", d.date()), // midnight
            Some(d) => write!(buf, "{d}"),
            None => write!(buf, "{v}"),
        },
        _ => write!(buf, "{v}"),
    };
    buf
}

impl CellText for String {
    fn text<'a>(&'a self, buf: &'a mut String) -> &'a str {
        // ODS stores formulas as `of:=SUM([.A1])`; xlsx/xlsb/xls without the `=`
        let f = self.strip_prefix("of:").unwrap_or(self);
        if f.is_empty() || f.starts_with('=') {
            return f;
        }
        buf.clear();
        buf.push('=');
        buf.push_str(f);
        buf
    }
}

/// Turn a read failure into a plain-English reason.
fn why_unreadable(e: &calamine::Error) -> String {
    let mut src: Option<&dyn std::error::Error> = Some(e);
    while let Some(s) = src {
        if let Some(io) = s.downcast_ref::<io::Error>() {
            if io.kind() == io::ErrorKind::PermissionDenied {
                return "access denied".into();
            }
            let msg = io.to_string();
            if msg.to_lowercase().contains("cloud") {
                return "OneDrive online-only file and it couldn't be downloaded (offline?)".into();
            }
            return msg;
        }
        src = s.source();
    }
    let msg = e.to_string();
    if msg.to_lowercase().contains("password") {
        return "password-protected".into();
    }
    format!("not a readable workbook ({msg})")
}

struct Hit {
    sheet: usize,
    /// 0-based, absolute (A1 = 0, 0)
    row: usize,
    col: usize,
    text: String,
    /// non-empty cells of the row with `--row`: (0-based column, text)
    row_values: Vec<(usize, String)>,
}

impl Hit {
    fn cell(&self) -> String {
        format!("{}{}", col_letter(self.col), self.row + 1)
    }
}

/// One file's results, built on a worker thread and printed in order by the main thread.
#[derive(Default)]
struct Report {
    sheets: Vec<String>,
    hits: Vec<Hit>,
    n: usize,
    /// per-sheet read failures: (sheet, reason)
    sheet_errors: Vec<(usize, String)>,
}

fn search(path: &Path, rx: &Regex, args: &Args) -> Result<Report, String> {
    if args.max_count == Some(0) {
        return Ok(Report::default()); // grep -m 0: nothing to find, so don't even open it
    }
    let mut wb = open_workbook_auto(path).map_err(|e| why_unreadable(&e))?;
    let mut scan = Scan::new(rx, args);
    scan.rep.sheets = wb.sheet_names();
    for si in 0..scan.rep.sheets.len() {
        let name = scan.rep.sheets[si].clone();
        // skipped before parsing, so filtering by sheet also saves the read (xlsx, xlsb)
        if !args.sheet.is_empty()
            && !args.sheet.iter().any(|s| s.to_lowercase() == name.to_lowercase())
        {
            continue;
        }
        match read_sheet(&mut wb, &name, args.formulas, |r, c, text| scan.cell(si, r, c, text)) {
            Ok(true) => break,
            Ok(false) => {}
            Err(e) => scan.rep.sheet_errors.push((si, why_unreadable(&e))),
        }
    }
    scan.end_row();
    Ok(scan.rep)
}

/// Feeds a sheet's cells to `cell` as (row, column, text) in row order, until it returns true;
/// returns whether it did. xlsx and xlsb stream, so one far-off cell can't make a small file
/// allocate its whole grid (calamine#693).
// ponytail: xls and ods are still read whole (calamine parses them on open); an ods with a
// far-off cell can still exhaust memory. Cap by dimensions if that shows up.
fn read_sheet(
    wb: &mut Sheets<BufReader<File>>,
    name: &str,
    formulas: bool,
    mut cell: impl FnMut(usize, usize, &str) -> bool,
) -> Result<bool, calamine::Error> {
    match wb {
        Sheets::Xlsx(x) => {
            let mut rd = match x.worksheet_cells_reader(name) {
                Err(XlsxError::NotAWorksheet(_)) => return Ok(false), // a chart sheet
                rd => rd.map_err(calamine::Error::Xlsx)?,
            };
            if formulas {
                stream(|| rd.next_formula(), &mut cell)
            } else {
                stream(|| rd.next_cell(), &mut cell)
            }
            .map_err(calamine::Error::Xlsx)
        }
        Sheets::Xlsb(x) => {
            let mut rd = x.worksheet_cells_reader(name).map_err(calamine::Error::Xlsb)?;
            if formulas {
                stream(|| rd.next_formula(), &mut cell)
            } else {
                stream(|| rd.next_cell(), &mut cell)
            }
            .map_err(calamine::Error::Xlsb)
        }
        _ if formulas => Ok(feed_range(&wb.worksheet_formula(name)?, &mut cell)),
        _ => Ok(feed_range(&wb.worksheet_range(name)?, &mut cell)),
    }
}

fn stream<T: CellText, E>(
    mut next: impl FnMut() -> Result<Option<Cell<T>>, E>,
    cell: &mut impl FnMut(usize, usize, &str) -> bool,
) -> Result<bool, E> {
    let mut buf = String::new();
    while let Some(c) = next()? {
        let (r, col) = c.get_position();
        if cell(r as usize, col as usize, c.get_value().text(&mut buf)) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn feed_range<T: CellText>(
    range: &Range<T>,
    cell: &mut impl FnMut(usize, usize, &str) -> bool,
) -> bool {
    let (r0, c0) = range.start().unwrap_or((0, 0));
    let mut buf = String::new();
    range.used_cells().any(|(r, c, v)| cell(r0 as usize + r, c0 as usize + c, v.text(&mut buf)))
}

/// Builds one file's report from its cells, fed in row order (the order calamine yields them).
struct Scan<'a> {
    rx: &'a Regex,
    args: &'a Args,
    rep: Report,
    /// with `--row`: the row being read as (sheet, row), its non-empty cells, whether it
    /// matched yet, and its hit, which gets the row's cells once the row ends
    row: Option<(usize, usize)>,
    row_cells: Vec<(usize, String)>,
    row_matched: bool,
    pending: Option<Hit>,
}

impl<'a> Scan<'a> {
    fn new(rx: &'a Regex, args: &'a Args) -> Self {
        let rep = Report::default();
        Scan { rx, args, rep, row: None, row_cells: Vec::new(), row_matched: false, pending: None }
    }

    /// Feeds one cell (0-based, absolute). Returns true when the file is done: an `-l` hit,
    /// or `-m` reached (with `--row`, once the row it was reached in is complete).
    fn cell(&mut self, sheet: usize, row: usize, col: usize, text: &str) -> bool {
        let args = self.args;
        if args.row && self.row != Some((sheet, row)) {
            if self.end_row() {
                return true;
            }
            self.row = Some((sheet, row));
        }
        if text.is_empty() {
            return false;
        }
        let whole_row = args.row && !args.count;
        if whole_row {
            self.row_cells.push((col, text.to_owned()));
        }
        if (args.row && self.row_matched) || !self.rx.is_match(text) {
            return false;
        }
        self.row_matched = true;
        self.rep.n += 1;
        if args.files_with_matches {
            return true;
        }
        if !args.count {
            let hit = Hit { sheet, row, col, text: text.to_owned(), row_values: Vec::new() };
            if args.row {
                self.pending = Some(hit);
            } else {
                self.rep.hits.push(hit);
            }
        }
        !whole_row && self.limit_reached()
    }

    /// Completes the current row's hit. Returns true when `-m` has been reached.
    fn end_row(&mut self) -> bool {
        if let Some(mut hit) = self.pending.take() {
            hit.row_values = std::mem::take(&mut self.row_cells);
            self.rep.hits.push(hit);
        }
        self.row_cells.clear();
        self.row_matched = false;
        self.limit_reached()
    }

    fn limit_reached(&self) -> bool {
        self.args.max_count.is_some_and(|m| self.rep.n >= m)
    }
}

fn json_str(out: &mut String, s: &str) {
    out.push('"');
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => {
                let _ = write!(out, "\\u{:04x}", c as u32);
            }
            c => out.push(c),
        }
    }
    out.push('"');
}

struct Printer<'a> {
    args: &'a Args,
    rx: &'a Regex,
    color: bool,
    /// terminal: group hits under a file heading
    heading: bool,
    /// terminal width; long values are cut to fit
    width: Option<usize>,
    files_printed: usize,
}

impl Printer<'_> {
    fn paint(&self, s: &str, code: &str) -> String {
        if self.color { format!("\x1b[{code}m{s}\x1b[0m") } else { s.to_owned() }
    }

    /// Cells joined by `ROW_SEP`, matches highlighted, control characters shown dim and escaped
    /// (`\n`, `\u{1b}`), cut to `budget` chars with a trailing `…` when they don't fit.
    // ponytail: counts chars, not display columns; wide (CJK) text can still wrap
    fn render(&self, cells: &[&str], budget: Option<usize>) -> String {
        let sep = ROW_SEP.chars().count();
        let total: usize = cells.iter().map(|c| visible(c).chars().count()).sum::<usize>()
            + sep * cells.len().saturating_sub(1);
        // one char reserved for the `…`, only when it'll be needed
        let mut left = match budget {
            Some(b) if total > b => b.saturating_sub(1),
            _ => usize::MAX,
        };
        let mut out = String::new();
        for (i, cell) in cells.iter().enumerate() {
            if i > 0 {
                if sep > left {
                    out.push('…');
                    break;
                }
                out += &self.paint(ROW_SEP, "2");
                left -= sep;
            }
            if !self.push_cell(&mut out, cell, &mut left) {
                out.push('…');
                break;
            }
        }
        out
    }

    /// Appends one cell; false when it ran out of room.
    fn push_cell(&self, out: &mut String, cell: &str, left: &mut usize) -> bool {
        let mut rest = cell;
        while !rest.is_empty() {
            let end = rest.find(char::is_control).unwrap_or(rest.len());
            let text = &rest[..end];
            let n = text.chars().count();
            if n > *left {
                let cut: String = text.chars().take(*left).collect();
                *out += &self.highlight(&cut);
                return false;
            }
            *out += &self.highlight(text);
            *left -= n;
            let Some(c) = rest[end..].chars().next() else { break };
            let esc = ctl_escape(c);
            if esc.len() > *left {
                return false;
            }
            *out += &self.paint(&esc, "2");
            *left -= esc.len();
            rest = &rest[end + c.len_utf8()..];
        }
        true
    }

    fn highlight(&self, text: &str) -> String {
        if !self.color {
            return text.to_owned();
        }
        self.rx.replace_all(text, |m: &regex::Captures| self.paint(&m[0], "1;31")).into_owned()
    }

    fn cells<'h>(&self, h: &'h Hit) -> Vec<&'h str> {
        if self.args.row {
            h.row_values.iter().map(|(_, v)| v.as_str()).collect()
        } else {
            vec![&h.text]
        }
    }

    fn report(&mut self, out: &mut impl Write, name: &str, rep: &Report) -> io::Result<()> {
        self.files_printed += 1;
        let args = self.args;
        if args.json {
            let mut line = String::new();
            for h in &rep.hits {
                line.clear();
                line.push_str("{\"file\":");
                json_str(&mut line, name);
                line.push_str(",\"sheet\":");
                json_str(&mut line, &rep.sheets[h.sheet]);
                let _ = write!(line, ",\"cell\":\"{}\",\"value\":", h.cell());
                json_str(&mut line, &h.text);
                if args.row {
                    // keyed by column letter so values map to columns even with gaps
                    line.push_str(",\"row_values\":{");
                    for (i, (c, v)) in h.row_values.iter().enumerate() {
                        if i > 0 {
                            line.push(',');
                        }
                        let _ = write!(line, "\"{}\":", col_letter(*c));
                        json_str(&mut line, v);
                    }
                    line.push('}');
                }
                line.push('}');
                writeln!(out, "{line}")?;
            }
            return Ok(());
        }
        let name = visible(name);
        if args.files_with_matches {
            return writeln!(out, "{}", self.paint(&name, "35"));
        }
        if args.count {
            return writeln!(out, "{}: {}", self.paint(&name, "35"), rep.n);
        }
        if !self.heading {
            let sep = self.paint("|", "2");
            let name = self.paint(&name, "35");
            for h in &rep.hits {
                let sheet = self.paint(&visible(&rep.sheets[h.sheet]), "36");
                let cell = self.paint(&h.cell(), "32");
                writeln!(
                    out,
                    "{name} {sep} {sheet} {sep} {cell} {sep} {}",
                    self.render(&self.cells(h), None)
                )?;
            }
            return Ok(());
        }

        if self.files_printed > 1 {
            writeln!(out)?;
        }
        writeln!(out, "{}", self.paint(&name, "1;35"))?;
        let refs: Vec<(String, String)> = rep
            .hits
            .iter()
            .map(|h| (sheet_ref(&visible(&rep.sheets[h.sheet])).into_owned(), h.cell()))
            .collect();
        let w = refs.iter().map(|(s, c)| s.chars().count() + 1 + c.len()).max().unwrap_or(0);
        let budget = self.width.map(|tw| tw.saturating_sub(2 + w + 2).max(20));
        for (h, (sheet, cell)) in rep.hits.iter().zip(&refs) {
            let pad = w - (sheet.chars().count() + 1 + cell.len());
            writeln!(
                out,
                "  {}!{}{:pad$}  {}",
                self.paint(sheet, "36"),
                self.paint(cell, "32"),
                "",
                self.render(&self.cells(h), budget)
            )?;
        }
        Ok(())
    }
}

fn plural(n: usize, one: &str, many: &str) -> String {
    format!("{n} {}", if n == 1 { one } else { many })
}

/// Wraps the parsed pattern, not its text, so `-w`/`-x` can't be swallowed by a `(?x)` comment.
fn build_regex(args: &Args) -> Result<Regex, String> {
    use regex_syntax::hir::{Hir, Look};
    let pat = if args.fixed_strings { regex::escape(&args.pattern) } else { args.pattern.clone() };
    let mut hir = regex_syntax::ParserBuilder::new()
        .case_insensitive(args.ignore_case)
        .build()
        .parse(&pat)
        .map_err(|e| e.to_string())?;
    if args.word_regexp {
        // half boundaries (grep -w): `-wF '(555)'` must match even though `(` isn't a word char
        let (start, end) = (Look::WordStartHalfUnicode, Look::WordEndHalfUnicode);
        hir = Hir::concat(vec![Hir::look(start), hir, Hir::look(end)]);
    }
    if args.cell_regexp {
        hir = Hir::concat(vec![Hir::look(Look::Start), hir, Hir::look(Look::End)]);
    }
    Regex::new(&hir.to_string()).map_err(|e| e.to_string())
}

/// Spreadsheets to search, as (path, `cloud_only`), each once in first-seen order. Counts
/// unusable path arguments and unreadable folders in `errors`.
fn collect(paths: &[PathBuf], errors: &mut usize) -> Vec<(PathBuf, bool)> {
    let mut found = Vec::new();
    for p in paths {
        match std::fs::metadata(p) {
            Err(_) => {
                eprintln!("xlgrep: {}: no such file or directory", p.display());
                *errors += 1;
            }
            Ok(m) if m.is_dir() => walk(p, &mut found, errors),
            Ok(m) if is_spreadsheet(p) => found.push((p.clone(), is_cloud_only(&m))),
            Ok(_) => {
                let exts = EXTS.join(" ");
                eprintln!("xlgrep: skipped {}: not a spreadsheet ({exts})", p.display());
                *errors += 1;
            }
        }
    }
    // `. a.xlsx` or `a.xlsx a.xlsx` would otherwise search (and print) a file twice
    let mut seen = std::collections::HashSet::new();
    found.retain(|(p, _)| seen.insert(std::path::absolute(p).unwrap_or_else(|_| p.clone())));
    found
}

/// Run `work` on files `0..n` in parallel and hand each result to `emit` in file order; `emit`
/// returns false to stop early (stdout closed). Workers run at most `window` files ahead of
/// `emit`, so one slow file can't make every later report pile up in memory; the worker
/// holding the oldest unemitted file never waits. A panic in `work` becomes that file's `Err`.
fn search_all(
    n: usize,
    work: impl Fn(usize) -> Result<Report, String> + Sync,
    mut emit: impl FnMut(usize, Result<Report, String>) -> bool,
) {
    let next = AtomicUsize::new(0);
    let workers = std::thread::available_parallelism().map_or(1, std::num::NonZero::get);
    let window = workers * 4;
    // (files emitted, stop requested)
    let progress = (Mutex::new((0usize, false)), Condvar::new());
    let (tx, results) = mpsc::channel();
    std::thread::scope(|s| {
        for _ in 0..workers.min(n) {
            let tx = tx.clone();
            let (next, progress, work) = (&next, &progress, &work);
            s.spawn(move || {
                loop {
                    let i = next.fetch_add(1, Ordering::Relaxed);
                    if i >= n {
                        break;
                    }
                    let p = progress.0.lock().unwrap();
                    let p = progress
                        .1
                        .wait_while(p, |(done, stop)| !*stop && i >= done.saturating_add(window))
                        .unwrap();
                    if p.1 {
                        break; // stopped early: don't start another workbook
                    }
                    drop(p);
                    // calamine panics on some malformed files; a lost result would stall `emit`
                    // and with it every worker. The default hook has already printed the panic.
                    let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| work(i)))
                        .unwrap_or_else(|_| Err("the reader crashed (a calamine bug)".into()));
                    if tx.send((i, res)).is_err() {
                        break;
                    }
                }
            });
        }
        drop(tx);

        let mut pending = std::collections::BTreeMap::new();
        let mut want = 0;
        'recv: for (i, res) in results {
            pending.insert(i, res);
            while let Some(res) = pending.remove(&want) {
                want += 1;
                progress.0.lock().unwrap().0 = want;
                progress.1.notify_all();
                if !emit(want - 1, res) {
                    break 'recv;
                }
            }
        }
        // done or stopped: wake waiting workers so they exit instead of blocking the scope
        progress.0.lock().unwrap().1 = true;
        progress.1.notify_all();
    });
}

fn main() -> ExitCode {
    let args = Args::parse();
    let rx = match build_regex(&args) {
        Ok(rx) => rx,
        Err(e) => {
            eprintln!("xlgrep: invalid regex: {e} (use -F to search literal text)");
            return ExitCode::from(2);
        }
    };

    // anything that left the search incomplete: exit 2, like grep (cloud skips are deliberate)
    let mut errors = 0usize;
    let found = collect(&args.paths, &mut errors);
    let total = found.len();
    let files: Vec<PathBuf> =
        found.into_iter().filter(|(_, cloud)| args.download || !cloud).map(|(p, _)| p).collect();
    let cloud_skipped = total - files.len();

    let cwd = std::env::current_dir().unwrap_or_default();
    let tty = io::stdout().is_terminal();
    let mut printer = Printer {
        args: &args,
        rx: &rx,
        // conhost shows escapes as garbage until VT is switched on; mintty (Git Bash) is a pipe
        // that can't be switched but speaks ANSI anyway. NO_COLOR counts only when non-empty.
        color: tty
            && !anstyle_query::no_color()
            && (anstyle_query::windows::enable_ansi_colors().unwrap_or(true)
                || anstyle_query::term_supports_ansi_color()),
        heading: tty,
        width: terminal_size::terminal_size().filter(|_| tty).map(|(w, _)| usize::from(w.0)),
        files_printed: 0,
    };
    // stdout is line-buffered by std: right for a terminal, slow for pipes. Flushed after each
    // file, so `| head` gets results as they're found and closing it stops the search.
    let mut out: Box<dyn Write> = if tty {
        Box::new(io::stdout().lock())
    } else {
        Box::new(BufWriter::new(io::stdout().lock()))
    };
    let (mut hits, mut files_hit) = (0usize, 0usize);
    let mut write_err = None;

    let search_file = |i: usize| search(&files[i], &rx, &args);
    search_all(files.len(), search_file, |i, res| {
        let name = display(&files[i], &cwd);
        let rep = match res {
            Ok(rep) => rep,
            Err(why) => {
                eprintln!("xlgrep: skipped {name}: {why}");
                errors += 1;
                return true;
            }
        };
        errors += rep.sheet_errors.len();
        for (si, why) in &rep.sheet_errors {
            eprintln!("xlgrep: skipped {name} | {}: {why}", visible(&rep.sheets[*si]));
        }
        if rep.n > 0 {
            hits += rep.n;
            files_hit += 1;
            if let Err(e) = printer.report(&mut out, &name, &rep).and_then(|()| out.flush()) {
                write_err = Some(e);
                return false;
            }
        }
        true
    });
    // a closed pipe (`| head`) is a normal way to stop; anything else (disk full) lost output
    if let Some(e) = write_err.filter(|e| e.kind() != io::ErrorKind::BrokenPipe) {
        eprintln!("xlgrep: error writing output: {e}");
        errors += 1;
    }

    if cloud_skipped > 0 {
        eprintln!(
            "xlgrep: skipped {} (use --download to include)",
            plural(cloud_skipped, "OneDrive online-only file", "OneDrive online-only files")
        );
    }
    if tty && !args.files_with_matches {
        let summary =
            format!("{} in {}", plural(hits, "hit", "hits"), plural(files_hit, "file", "files"));
        eprintln!("{}", printer.paint(&summary, "2"));
    }
    ExitCode::from(if errors > 0 { 2 } else { u8::from(hits == 0) })
}

#[cfg(test)]
mod tests {
    use super::{
        Args, Printer, Report, Scan, build_regex, col_letter, json_str, search_all, sheet_ref,
    };
    use clap::Parser;
    use std::time::Duration;

    fn rx(argv: &[&str]) -> regex::Regex {
        build_regex(&Args::parse_from([&["xlgrep"], argv].concat())).unwrap()
    }

    #[test]
    fn word_and_cell_flags_wrap_the_pattern() {
        // half boundaries: punctuation at the edges still counts as a whole word
        assert!(rx(&["-wF", "(555)"]).is_match("(555) 123-4567"));
        assert!(rx(&["-wF", "C++"]).is_match("C++ dev"));
        assert!(!rx(&["-w", "de"]).is_match("C++ dev"));
        assert!(rx(&["-xw", "-F", "(555) 1"]).is_match("(555) 1"));
        assert!(!rx(&["-x", "smith"]).is_match("smiths"));
        // a (?x) comment must not swallow the -x anchors
        assert!(rx(&["-x", "(?x) C\\+\\+ \\s dev # role"]).is_match("C++ dev"));
        assert!(!rx(&["-x", "(?x) C\\+\\+ # role"]).is_match("C++ dev"));
    }

    #[test]
    fn render_cuts_only_what_overflows() {
        let args = Args::parse_from(["xlgrep", "x"]);
        let rx = regex::Regex::new("x").unwrap();
        let p = Printer {
            args: &args,
            rx: &rx,
            color: false,
            heading: true,
            width: None,
            files_printed: 0,
        };
        assert_eq!(p.render(&["abcde"], Some(5)), "abcde"); // exact fit: no ellipsis
        assert_eq!(p.render(&["abcdef"], Some(5)), "abcd…");
        assert_eq!(p.render(&["ab", "cdef"], Some(7)), "ab │ c…");
        assert_eq!(p.render(&["a\nb\u{1b}"], None), "a\\nb\\u{1b}");
    }

    #[test]
    fn a_panicking_file_is_skipped_not_hung() {
        // far more files than the worker window: a lost result used to stall every worker
        let n = 1000;
        let (tx, rx) = std::sync::mpsc::channel();
        std::thread::spawn(move || {
            let mut got = Vec::new();
            search_all(
                n,
                |i| if i == 1 { panic!("simulated calamine panic") } else { Ok(Report::default()) },
                |i, res| {
                    got.push((i, res.is_ok()));
                    true
                },
            );
            tx.send(got).unwrap();
        });
        let got = rx.recv_timeout(Duration::from_secs(30)).expect("search_all hung");
        let want: Vec<_> = (0..n).map(|i| (i, i != 1)).collect();
        assert_eq!(got, want);
    }

    #[test]
    fn row_hits_get_the_whole_row_and_m_ends_after_it() {
        let args = Args::parse_from(["xlgrep", "smith", "--row", "-m", "2"]);
        let rx = build_regex(&args).unwrap();
        let mut scan = Scan::new(&rx, &args);
        // in reading order: each row's hit comes before the rest of its row has been read
        let cells = [(0, 0, "a"), (0, 1, "smith"), (0, 2, "smith"), (1, 0, "smith"), (1, 3, "z")];
        assert!(!cells.iter().any(|&(r, c, t)| scan.cell(0, r, c, t)));
        assert!(scan.cell(0, 2, 0, "smith")); // -m 2 reached in row 1: stop once it's complete
        let s = |t: &str| t.to_owned();
        let got: Vec<_> = scan.rep.hits.iter().map(|h| (h.row, h.col, &h.row_values)).collect();
        let row0 = vec![(0, s("a")), (1, s("smith")), (2, s("smith"))];
        let row1 = vec![(0, s("smith")), (3, s("z"))];
        assert_eq!(got, [(0, 1, &row0), (1, 0, &row1)]);
    }

    #[test]
    fn col_letters_roll_over() {
        let got: Vec<_> = [0, 25, 26, 51, 52, 701, 702, 16383].map(col_letter).into();
        assert_eq!(got, ["A", "Z", "AA", "AZ", "BA", "ZZ", "AAA", "XFD"]);
    }

    #[test]
    fn sheet_refs_quote_like_excel() {
        assert_eq!(sheet_ref("Data"), "Data");
        assert_eq!(sheet_ref("Q3 Sales"), "'Q3 Sales'");
        assert_eq!(sheet_ref("2024"), "'2024'");
        assert_eq!(sheet_ref("Bob's"), "'Bob''s'");
    }

    #[test]
    fn json_strings_escape_per_rfc_8259() {
        let mut out = String::new();
        json_str(&mut out, "say \"hi\" C:\\x\n\r\t\u{1}\u{1f} é€😀\u{7f}");
        assert_eq!(out, r#""say \"hi\" C:\\x\n\r\t\u0001\u001f é€😀"#.to_owned() + "\u{7f}\"");
    }
}
