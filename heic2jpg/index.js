const fs = require("fs");
const path = require("path");
const sharp = require("sharp");

// Change these paths as needed
const inputDir = path.join(__dirname, "input_heic");
const outputDir = path.join(__dirname, "output_jpg");

// Ensure the output directory exists
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

fs.readdir(inputDir, (err, files) => {
  if (err) {
    console.error("Error reading the directory:", err);
    return;
  }

  files.forEach(async (file) => {
    const ext = path.extname(file).toLowerCase();
    if (ext === ".heic") {
      const fileNameWithoutExt = path.basename(file, ext);
      const inputFilePath = path.join(inputDir, file);
      const outputFilePath = path.join(outputDir, `${fileNameWithoutExt}.jpg`);

      try {
        await sharp(inputFilePath)
          .jpeg({ quality: 90 }) // Adjust quality as needed
          .toFile(outputFilePath);

        console.log(
          `Successfully converted: ${file} → ${fileNameWithoutExt}.jpg`
        );
      } catch (conversionError) {
        console.error(`Failed to convert ${file}:`, conversionError);
      }
    }
  });
});
