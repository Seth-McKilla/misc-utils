# pst2pdf (Node.js)

Convert an Outlook .pst archive into a single PDF containing all email conversations. Attachments are ignored by design.

## Requirements

- Node.js 18+ (recommended) with pnpm or npm.
- The `readpst` binary is required, but you don't need it installed system-wide:
  - Option A: Place the binary at `pst2pdf/bin/readpst` (make it executable)
  - Option B: Provide a custom path via `--readpst-bin /path/to/readpst`
  - Option C: Have `readpst` available on PATH (e.g., Ubuntu: `sudo apt-get install readpst`, macOS: `brew install libpst`)

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
Usage: pst2pdf [<input.pst>] [options]

Options:
  -o, --output <file>     Output PDF path (default: <input>.pdf)
  -i, --input-dir <dir>   Directory to read .pst files from (batch mode)
  -O, --output-dir <dir>  Directory to write generated PDFs to (batch mode)
  -R, --readpst-bin <p>   Path to readpst binary (default: ./bin/readpst or PATH)
  -w, --workdir <dir>     Working directory for extracted files (default: temp)
  --keep-workdir          Do not delete working directory
  --max-emails <n>        Limit processed emails (for quick tests)
  -h, --help              Show this help

Requirements:
  - Provide a 'readpst' binary either in pst2pdf/bin/readpst, via --readpst-bin, or on PATH.
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

### Batch mode with input/output folders

This tool creates two folders by default:

- `pst2pdf/input_pst` — put your `.pst` files here
- `pst2pdf/output_pdf` — generated PDFs will be placed here

To run in batch mode (process all `.pst` files in `input_pst` and write PDFs to `output_pdf`):

```

pnpm run pst2pdf

```

You can override the folders:

```

pnpm run pst2pdf -- --input-dir /path/to/pst_dir --output-dir /path/to/pdf_dir

### Using a local readpst without installing system-wide

Place the `readpst` binary into `pst2pdf/bin/readpst` and ensure it is executable:

```
chmod +x pst2pdf/bin/readpst
```

Alternatively, specify a custom path:

```
pnpm run pst2pdf -- --readpst-bin /opt/libpst/bin/readpst
```

```

```

## License

MIT
