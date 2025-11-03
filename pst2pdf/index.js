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
Usage: pst2pdf <input.pst> [options]

Options:
	-o, --output <file>     Output PDF path (default: <input>.pdf)
	-w, --workdir <dir>     Working directory for extracted files (default: temp)
	--keep-workdir          Do not delete working directory
	--max-emails <n>        Limit processed emails (for quick tests)
	-h, --help              Show this help

Requirements:
	- The 'readpst' binary must be installed and available on PATH.
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

async function runReadPst(pstPath, outDir) {
  await fsp.mkdir(outDir, { recursive: true });
  // Use readpst to emit each message as an .eml file, flattening the folder structure.
  // Common flags: -e (eml), -m (split), -b (output basename safe), -o outputDir
  const args = ["-e", "-m", "-b", "-o", outDir, pstPath];
  const child = spawn("readpst", args, { stdio: ["ignore", "pipe", "pipe"] });
  const out = [],
    err = [];
  child.stdout.on("data", (d) => out.push(d));
  child.stderr.on("data", (d) => err.push(d));
  const exit = await new Promise((resolve) => child.on("close", resolve));
  if (exit !== 0) {
    throw new Error(
      `readpst failed (code ${exit})\n${Buffer.concat(err).toString()}`
    );
  }
  return outDir;
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

async function main() {
  const args = process.argv.slice(2);
  if (args.length === 0 || args.includes("-h") || args.includes("--help")) {
    printHelp();
    process.exit(0);
  }

  const inputPst = args[0];
  if (!inputPst || !fs.existsSync(inputPst)) {
    console.error("Input .pst not found.");
    process.exit(1);
  }

  let outPdf = null;
  let workdir = null;
  let keepWorkdir = false;
  let maxEmails = 0;
  for (let i = 1; i < args.length; i++) {
    const a = args[i];
    if ((a === "-o" || a === "--output") && args[i + 1]) {
      outPdf = args[++i];
    } else if ((a === "-w" || a === "--workdir") && args[i + 1]) {
      workdir = args[++i];
    } else if (a === "--keep-workdir") {
      keepWorkdir = true;
    } else if (a === "--max-emails" && args[i + 1]) {
      maxEmails = parseInt(args[++i], 10) || 0;
    }
  }
  if (!outPdf)
    outPdf = path.resolve(
      path.dirname(inputPst),
      path.basename(inputPst, path.extname(inputPst)) + ".pdf"
    );
  if (!workdir) workdir = await fsp.mkdtemp(path.join(os.tmpdir(), "pst2pdf-"));

  // Run pipeline
  try {
    const extractDir = path.join(workdir, "eml");
    await runReadPst(inputPst, extractDir);
    const emlFiles = await listFilesRecursively(extractDir, [".eml"]);
    const emails = await parseEmails(emlFiles, maxEmails);
    const threads = groupByThread(emails);
    await renderPdf(threads, outPdf);
    console.log(`Wrote ${outPdf}`);
  } catch (e) {
    console.error(e.message || e);
    process.exitCode = 1;
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

if (require.main === module) {
  main();
}
