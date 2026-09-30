const fs = require("fs");
const path = require("path");

// Folders to empty out (the .gitkeep in each one is left alone)
const dirs = [
  path.join(__dirname, "input_heic"),
  path.join(__dirname, "output_jpg"),
];

for (const dir of dirs) {
  if (!fs.existsSync(dir)) {
    continue;
  }

  let removed = 0;
  for (const entry of fs.readdirSync(dir)) {
    if (entry === ".gitkeep") {
      continue;
    }

    fs.rmSync(path.join(dir, entry), { recursive: true, force: true });
    removed++;
  }

  console.log(`Removed ${removed} item(s) from ${path.basename(dir)}`);
}
