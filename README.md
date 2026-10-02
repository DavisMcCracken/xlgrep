# xlgrep-rs

Rust port of [xlgrep](https://github.com/DavisMcCracken/xlgrep), kept for comparison. Same
reader (calamine) and same flags: `-i`, `-F`, `-c`, `-l`, `--download`. Skips OneDrive
online-only files by default. No colors or plain-English error reasons (the Python version
has both).

```
cargo build --release
target/release/xlgrep-rs smith -i
```

On Windows without admin/Visual Studio, use the GNU toolchain: `scoop install rustup-gnu`.

Benchmarks vs the Python version (2026-10-02, hyperfine): 6.5x faster on a small folder
(47 ms vs 303 ms, mostly Python startup), 3.4x on a 50k-row file, 2.2x walking ~12k files.
