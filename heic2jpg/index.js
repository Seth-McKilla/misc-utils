const fs = require("fs");
const path = require("path");
const heicConvert = require("heic-convert");

(async () => {
  // Adjust these paths as needed
  const inputDir = path.join(__dirname, "input_heic");
  const outputDir = path.join(__dirname, "output_jpg");

  // Ensure output folder exists
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Read all files in inputDir
  fs.readdir(inputDir, (err, files) => {
    if (err) {
      console.error(`Error reading directory ${inputDir}:`, err);
      return;
    }

    // Process each file if it has a .heic extension
    files.forEach(async (file) => {
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
    });
  });
})();
