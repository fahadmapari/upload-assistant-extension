const JavaScriptObfuscator = require("javascript-obfuscator");
const archiver = require("archiver");
const fs = require("fs");
const path = require("path");

const BUILD_DIR = path.join(__dirname, "dist");
const ZIP_PATH = path.join(__dirname, "extension.zip");

// Files/folders to include in the zip (relative to project root)
const STATIC_FILES = ["manifest.json", "popup.html"];
const STATIC_DIRS = ["icons"];

// JS files to obfuscate (relative to project root)
const JS_FILES = [
  "src/content.js",
  "src/background.js",
  "src/popup.js",
  "src/parisDocParser.js",
  "src/license.js",
];

// Obfuscator options — tweak as needed
const OBFUSCATOR_OPTIONS = {
  compact: true,
  controlFlowFlattening: true,
  controlFlowFlatteningThreshold: 0.5,
  deadCodeInjection: false,
  debugProtection: false,
  disableConsoleOutput: false,
  identifierNamesGenerator: "hexadecimal",
  renameGlobals: false,
  selfDefending: false,
  stringArray: true,
  stringArrayEncoding: ["base64"],
  stringArrayThreshold: 0.75,
  transformObjectKeys: false,
  unicodeEscapeSequence: false,
};

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function copyDir(src, dest) {
  ensureDir(dest);
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyDir(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

// 1. Clean and recreate dist dir
if (fs.existsSync(BUILD_DIR)) fs.rmSync(BUILD_DIR, { recursive: true });
ensureDir(BUILD_DIR);

// 2. Copy static files
for (const file of STATIC_FILES) {
  const dest = path.join(BUILD_DIR, file);
  ensureDir(path.dirname(dest));
  fs.copyFileSync(path.join(__dirname, file), dest);
  console.log(`Copied: ${file}`);
}

// 3. Copy static directories
for (const dir of STATIC_DIRS) {
  copyDir(path.join(__dirname, dir), path.join(BUILD_DIR, dir));
  console.log(`Copied dir: ${dir}/`);
}

// 4. Obfuscate JS files
for (const file of JS_FILES) {
  const src = path.join(__dirname, file);
  const dest = path.join(BUILD_DIR, file);
  ensureDir(path.dirname(dest));

  const code = fs.readFileSync(src, "utf8");
  const result = JavaScriptObfuscator.obfuscate(code, OBFUSCATOR_OPTIONS);
  fs.writeFileSync(dest, result.getObfuscatedCode(), "utf8");
  console.log(`Obfuscated: ${file}`);
}

// 5. Zip dist dir
if (fs.existsSync(ZIP_PATH)) fs.unlinkSync(ZIP_PATH);

const output = fs.createWriteStream(ZIP_PATH);
const archive = archiver("zip", { zlib: { level: 9 } });

output.on("close", () => {
  console.log(`\nDone! extension.zip created (${archive.pointer()} bytes)`);
});

archive.on("error", (err) => {
  throw err;
});

archive.pipe(output);
archive.directory(BUILD_DIR, false);
archive.finalize();
