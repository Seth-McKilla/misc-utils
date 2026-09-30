const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawn } = require("child_process");
const heicConvert = require("heic-convert");

// Open a folder in the system file explorer
function openFolder(dir) {
  const isWsl =
    process.platform === "linux" && os.release().toLowerCase().includes("microsoft");

  let command;
  if (process.platform === "win32" || isWsl) {
    command = "explorer.exe";
  } else if (process.platform === "darwin") {
    command = "open";
  } else {
    command = "xdg-open";
  }

  // Run from inside the folder and open "." so WSL paths don't need translating
  const child = spawn(command, ["."], { cwd: dir, detached: true, stdio: "ignore" });
  child.on("error", (err) => {
    console.error(`Could not open ${dir}:`, err.message);
  });
  child.unref();
}

(async () => {
  // Adjust these paths as needed
  const inputDir = path.join(__dirname, "input_heic");
  const outputDir = path.join(__dirname, "output_jpg");

  // Ensure output folder exists
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Read all files in inputDir
  let files;
  try {
    files = fs.readdirSync(inputDir);
  } catch (err) {
    console.error(`Error reading directory ${inputDir}:`, err);
    return;
  }

  // Process each file if it has a .heic extension
  for (const file of files) {
    const ext = path.extname(file).toLowerCase();
    if (ext === ".heic") {
      const baseName = path.basename(file, ext);
      const inputFilePath = path.join(inputDir, file);
      const outputFilePath = path.join(outputDir, `${baseName}.jpg`);

      try {
        // Read the HEIC file into a buffer
        const inputBuffer = fs.readFileSync(inputFilePath);

        // Convert the HEIC buffer to a JPEG buffer
        const outputBuffer = await heicConvert({
          buffer: inputBuffer,
          format: "JPEG",
          quality: 1, // quality = 1 => highest JPEG quality (0 to 1)
        });

        // Write out the new JPEG file
        fs.writeFileSync(outputFilePath, outputBuffer);

        console.log(`Successfully converted: ${file} → ${baseName}.jpg`);
      } catch (conversionError) {
        console.error(`Failed to convert ${file}:`, conversionError);
      }
    }
  }

  // Open the output folder so the converted files are easy to grab
  openFolder(outputDir);
})();
