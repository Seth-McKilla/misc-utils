#!/usr/bin/env node
/**
 * pst2pdf - Convert a .pst Outlook archive to a single PDF of email conversations.
 *
 * Implementation notes:
 * - Uses the external `readpst` binary to extract messages to .eml files.
 * - Parses EML with `mailparser` to build threads grouped by normalized subject.
 * - Orders emails in each thread by date and renders a single PDF using `pdfkit`.
 * - Attachments are ignored by design.
 */

const { spawn } = require("child_process");
const fs = require("fs");
const fsp = fs.promises;
const path = require("path");
const os = require("os");
const { simpleParser } = require("mailparser");
const PDFDocument = require("pdfkit");

function printHelp() {
  const msg = `
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
`;
  process.stdout.write(msg);
}

function normalizeSubject(subj) {
  if (!subj) return "(no subject)";
  // Strip common reply/forward prefixes and whitespace
  let s = subj.trim();
  s = s.replace(/^\s*(re|fw|fwd)\s*:\s*/i, "");
  return s || "(no subject)";
}

async function runReadPst(readpstBin, pstPath, outDir) {
  await fsp.mkdir(outDir, { recursive: true });
  // Use readpst to emit each message as an .eml file, flattening the folder structure.
  // Common flags: -e (eml), -m (split), -b (output basename safe), -o outputDir
  const args = ["-e", "-m", "-b", "-o", outDir, pstPath];
  return await new Promise((resolve, reject) => {
    const child = spawn(readpstBin, args, {
      stdio: ["ignore", "pipe", "pipe"],
    });
    const out = [],
      err = [];
    child.stdout.on("data", (d) => out.push(d));
    child.stderr.on("data", (d) => err.push(d));
    child.on("error", (errEvt) => {
      reject(
        new Error(`Failed to run readpst at '${readpstBin}': ${errEvt.message}`)
      );
    });
    child.on("close", (exit) => {
      if (exit !== 0) {
        reject(
          new Error(
            `readpst failed (code ${exit})\n${Buffer.concat(err).toString()}`
          )
        );
      } else {
        resolve(outDir);
      }
    });
  });
}

async function listFilesRecursively(dir, exts = [".eml"]) {
  const results = [];
  async function walk(p) {
    const entries = await fsp.readdir(p, { withFileTypes: true });
    for (const ent of entries) {
      const full = path.join(p, ent.name);
      if (ent.isDirectory()) await walk(full);
      else if (exts.includes(path.extname(ent.name).toLowerCase()))
        results.push(full);
    }
  }
  await walk(dir);
  return results;
}

async function parseEmails(emlFiles, maxEmails) {
  const emails = [];
  for (let i = 0; i < emlFiles.length; i++) {
    if (maxEmails && emails.length >= maxEmails) break;
    const file = emlFiles[i];
    try {
      const raw = await fsp.readFile(file);
      const mail = await simpleParser(raw);
      emails.push({
        file,
        subject: mail.subject || "(no subject)",
        normalizedSubject: normalizeSubject(mail.subject),
        date: mail.date ? new Date(mail.date) : new Date(0),
        from: mail.from ? mail.from.text : "",
        to: mail.to ? mail.to.text : "",
        cc: mail.cc ? mail.cc.text : "",
        text: mail.text || (mail.html ? stripHtml(mail.html) : ""),
      });
    } catch (e) {
      // Skip problematic email files but continue
      // eslint-disable-next-line no-console
      console.warn(`Failed to parse ${file}: ${e.message}`);
    }
  }
  return emails;
}

function stripHtml(html) {
  return String(html)
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function groupByThread(emails) {
  const map = new Map();
  for (const e of emails) {
    const key = e.normalizedSubject.toLowerCase();
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(e);
  }
  // Sort emails in each thread by date
  for (const arr of map.values()) arr.sort((a, b) => a.date - b.date);
  // Return array of {subject, emails}
  const threads = Array.from(map.entries()).map(([key, arr]) => ({
    subject: arr[0].normalizedSubject,
    emails: arr,
  }));
  // Sort threads by date of first message
  threads.sort((a, b) => a.emails[0].date - b.emails[0].date);
  return threads;
}

async function renderPdf(threads, outPdf) {
  await fsp.mkdir(path.dirname(outPdf), { recursive: true });
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ autoFirstPage: false, margin: 50 });
    const stream = fs.createWriteStream(outPdf);
    doc.pipe(stream);

    const titleStyle = { size: 18, color: "#000" };
    const headerStyle = { size: 10, color: "#333" };
    const bodyStyle = { size: 11, color: "#000" };

    const addPageIfNeeded = () => {
      if (doc.page === null) doc.addPage();
    };

    doc.info.Title = "PST Export";
    doc.info.Producer = "pst2pdf";

    for (const thread of threads) {
      addPageIfNeeded();
      doc.addPage();
      doc
        .fillColor("#000")
        .fontSize(titleStyle.size)
        .text(`Subject: ${thread.subject}`, { underline: true });
      doc.moveDown(0.5);
      for (const msg of thread.emails) {
        doc.fillColor("#000").fontSize(12).text("—", { continued: false });
        doc
          .fillColor(headerStyle.color)
          .fontSize(headerStyle.size)
          .text(`Date: ${msg.date.toISOString()}`)
          .text(`From: ${msg.from}`)
          .text(`To: ${msg.to}`);
        if (msg.cc) doc.text(`Cc: ${msg.cc}`);
        doc.moveDown(0.25);
        doc
          .fillColor(bodyStyle.color)
          .fontSize(bodyStyle.size)
          .text(msg.text || "[no body]", { align: "left" });
        doc.moveDown(0.75);
      }
    }

    doc.end();
    stream.on("finish", resolve);
    stream.on("error", reject);
  });
}

async function ensureDirs(...dirs) {
  for (const d of dirs) {
    if (!d) continue;
    await fsp.mkdir(d, { recursive: true });
  }
}

async function processSinglePst(
  readpstBin,
  inputPst,
  outPdf,
  workdir,
  keepWorkdir,
  maxEmails
) {
  if (!outPdf)
    outPdf = path.resolve(
      path.dirname(inputPst),
      path.basename(inputPst, path.extname(inputPst)) + ".pdf"
    );
  if (!workdir) workdir = await fsp.mkdtemp(path.join(os.tmpdir(), "pst2pdf-"));
  try {
    const extractDir = path.join(workdir, "eml");
    await runReadPst(readpstBin, inputPst, extractDir);
    const emlFiles = await listFilesRecursively(extractDir, [".eml"]);
    const emails = await parseEmails(emlFiles, maxEmails);
    const threads = groupByThread(emails);
    await renderPdf(threads, outPdf);
    console.log(`Wrote ${outPdf}`);
  } finally {
    if (!keepWorkdir) {
      try {
        await fsp.rm(workdir, { recursive: true, force: true });
      } catch (_) {}
    } else {
      console.log("Working directory kept at:", workdir);
    }
  }
}

async function main() {
  const args = process.argv.slice(2);
  if (args.includes("-h") || args.includes("--help")) {
    printHelp();
    process.exit(0);
  }

  let inputPst = null;
  let outPdf = null;
  let workdir = null;
  let keepWorkdir = false;
  let maxEmails = 0;
  let inputDir = null;
  let outputDir = null;
  let readpstCli = null;
  // Default folders inside this tool's directory
  const toolRoot = path.resolve(__dirname);
  const defaultInputDir = path.join(toolRoot, "input_pst");
  const defaultOutputDir = path.join(toolRoot, "output_pdf");
  for (let i = 1; i < args.length; i++) {
    const a = args[i];
    if ((a === "-o" || a === "--output") && args[i + 1]) {
      outPdf = args[++i];
    } else if ((a === "-i" || a === "--input-dir") && args[i + 1]) {
      inputDir = args[++i];
    } else if ((a === "-O" || a === "--output-dir") && args[i + 1]) {
      outputDir = args[++i];
    } else if ((a === "-R" || a === "--readpst-bin") && args[i + 1]) {
      readpstCli = args[++i];
    } else if ((a === "-w" || a === "--workdir") && args[i + 1]) {
      workdir = args[++i];
    } else if (a === "--keep-workdir") {
      keepWorkdir = true;
    } else if (a === "--max-emails" && args[i + 1]) {
      maxEmails = parseInt(args[++i], 10) || 0;
    }
  }
  // Ensure default input/output directories exist
  await ensureDirs(defaultInputDir, defaultOutputDir);

  // If a positional arg is present and not a flag, treat it as inputPst
  if (args[0] && !args[0].startsWith("-")) {
    inputPst = args[0];
  }

  // Resolve readpst binary prioritizing CLI flag, then local bin, then PATH
  const localBin = path.join(
    toolRoot,
    "bin",
    process.platform === "win32" ? "readpst.exe" : "readpst"
  );
  let readpstBin = "readpst";
  if (readpstCli) {
    readpstBin = path.resolve(readpstCli);
  } else if (fs.existsSync(localBin)) {
    readpstBin = localBin;
    try {
      await fsp.chmod(localBin, 0o755);
    } catch (_) {}
  }

  // Batch mode: no inputPst => read all .pst from inputDir (or default)
  if (!inputPst) {
    const srcDir = inputDir ? path.resolve(inputDir) : defaultInputDir;
    const dstDir = outputDir ? path.resolve(outputDir) : defaultOutputDir;
    await ensureDirs(srcDir, dstDir);
    const entries = fs.existsSync(srcDir) ? await fsp.readdir(srcDir) : [];
    const pstFiles = entries.filter((f) => f.toLowerCase().endsWith(".pst"));
    if (pstFiles.length === 0) {
      console.log(
        `No .pst files found in ${srcDir}. Place files there and rerun.`
      );
      printHelp();
      return;
    }
    for (const name of pstFiles) {
      const inPath = path.join(srcDir, name);
      const outPath = path.join(
        dstDir,
        path.basename(name, path.extname(name)) + ".pdf"
      );
      console.log(`Processing: ${inPath} -> ${outPath}`);
      try {
        await processSinglePst(
          readpstBin,
          inPath,
          outPath,
          workdir,
          keepWorkdir,
          maxEmails
        );
      } catch (e) {
        console.error(`Failed ${name}:`, e.message || e);
      }
    }
    return;
  }

  // Single file mode
  if (!fs.existsSync(inputPst)) {
    console.error("Input .pst not found.");
    process.exit(1);
  }
  if (!outPdf) {
    const dstDir = outputDir ? path.resolve(outputDir) : defaultOutputDir;
    await ensureDirs(dstDir);
    outPdf = path.join(
      dstDir,
      path.basename(inputPst, path.extname(inputPst)) + ".pdf"
    );
  }
  await processSinglePst(
    readpstBin,
    inputPst,
    outPdf,
    workdir,
    keepWorkdir,
    maxEmails
  );
}

if (require.main === module) {
  main();
}
