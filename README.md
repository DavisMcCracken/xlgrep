# xlgrep

Search spreadsheet cell values across folders of `.xlsx` / `.xlsm` / `.xlsb` / `.xls` / `.ods`
files. Read-only. Rust rewrite of the original [Python xlgrep](https://github.com/DavisMcCracken/xlgrep-py)
(same reader, calamine); files are read in parallel, output stays in file order.

```
xlgrep smith -i                  # current folder, recursive, case-insensitive
xlgrep "O'Brien" C:\dir -F       # literal text
xlgrep smith -w                  # whole word (-wF "C++" works too)
xlgrep smith -x                  # whole cell
xlgrep smith --row               # print each matching row in full
xlgrep smith --sheet Contacts    # only that sheet (others aren't even parsed)
xlgrep VLOOKUP --formulas -l     # workbooks whose formulas use VLOOKUP
xlgrep smith -c                  # hit count per file (rows with --row)
xlgrep smith -m 5                # stop each file after 5 hits
xlgrep smith --json --row        # JSON Lines for scripts and agents
xlgrep smith --json | head -50   # first 50 hits overall; the search stops early
```

In a terminal, hits are grouped under each file with Excel-style refs (`'Q3 Sales'!B7`, pastes
into the Name Box), colored (`NO_COLOR` disables) and cut to the window width. Piped, it's one
`file | sheet | cell | value` line per hit. Both are for reading: control characters print
escaped (`\n`, `\u{1b}`) so a cell can't break lines or inject terminal codes. Anything that
parses output should use `--json`: `{"file","sheet","cell","value"}` per hit, plus
`"row_values": {"A": …, "C": …}` keyed by column with `--row` (formula cells only with
`--formulas`). Values print as Excel shows them where calamine allows: `TRUE`, `2024-03-05`,
`36:05:07` durations; numbers are raw (`51200`, not `$51,200.00`).

Unreadable files and folders (password-protected, corrupt, access denied, OneDrive offline)
are skipped with a reason. OneDrive online-only files are skipped by default so a search
never mass-downloads your OneDrive; `--download` includes them. Exit code 0 if anything
matched, 1 if not, 2 on an error: bad arguments, or anything that couldn't be searched
(matches elsewhere still print). Skipped OneDrive files aren't errors.

```
cargo build --release
cargo test
```

On Windows without admin/Visual Studio, use the GNU toolchain: `scoop install rustup-gnu`.

Benchmarks vs the Python version (2026-10-02, hyperfine, before parallel reads): 6.5x faster on
a small folder (47 ms vs 303 ms, mostly Python startup), 3.4x on a 50k-row file, 2.2x walking
~12k files. Parallel reads: 400 files x 2000 rows went from 1.56 s to 0.25 s (warm cache).
