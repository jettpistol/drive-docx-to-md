// Copies the browser builds of the third-party libraries into vendor/.
// Run after `npm install` whenever a dependency is upgraded.
import { copyFileSync, mkdirSync } from "node:fs";

const files = {
  "node_modules/mammoth/mammoth.browser.min.js": "vendor/mammoth.browser.min.js",
  "node_modules/turndown/lib/turndown.browser.umd.js": "vendor/turndown.js",
  "node_modules/exceljs/dist/exceljs.min.js": "vendor/exceljs.min.js",
  "node_modules/jszip/dist/jszip.min.js": "vendor/jszip.min.js",
};

mkdirSync("vendor", { recursive: true });
for (const [from, to] of Object.entries(files)) {
  copyFileSync(from, to);
  console.log(`${from} -> ${to}`);
}
