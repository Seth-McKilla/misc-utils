# pst2pdf (Node.js)

Convert an Outlook .pst archive into a single PDF containing all email conversations. Attachments are ignored by design.

## Requirements

- Linux/macOS with the `readpst` binary installed and available in PATH.
  - On Ubuntu/Debian: `sudo apt-get install readpst`
  - On macOS (Homebrew): `brew install libpst`
- Node.js 18+ (recommended) with pnpm or npm.

## Usage

From the repository root:

```
pnpm exec node pst2pdf -- --help
```

Or add a script (already added) and run:

```
pnpm run pst2pdf -- <path/to/archive.pst> -o out.pdf
```

CLI:

```
Usage: pst2pdf <input.pst> [options]

Options:
  -o, --output <file>     Output PDF path (default: <input>.pdf)
  -w, --workdir <dir>     Working directory for extracted files (default: temp)
  --keep-workdir          Do not delete working directory
  --max-emails <n>        Limit processed emails (for quick tests)
  -h, --help              Show this help

Requirements:
  - The 'readpst' binary must be installed and available on PATH.
```

## What it does

- Uses `readpst` to export emails in the PST to `.eml` files.
- Parses emails with `mailparser`.
- Groups messages into threads by a normalized subject (strips common Re:/Fwd: prefixes).
- Sorts messages by date within each thread.
- Renders a single PDF using `pdfkit` with headers (date/from/to/cc) and the text body.
- Attachments are ignored.

## Notes and caveats

- HTML bodies are converted to text with a simple tag strip; formatting is not preserved.
- Threading by subject is heuristic and may merge/split some conversations incorrectly.
- Very large PSTs can take time; use `--max-emails` to test quickly.
- The working directory is deleted by default; pass `--keep-workdir` to inspect extracted `.eml` files.

## Development

Install dependencies at the repo root:

```
pnpm install
```

Try the CLI help:

```
pnpm run pst2pdf -- --help
```

Run on a PST:

```
pnpm run pst2pdf -- /path/to/mail.pst -o mail.pdf
```

## License

MIT
