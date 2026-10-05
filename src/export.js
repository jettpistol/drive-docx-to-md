import {
  MIME,
  PathAllocator,
  classifyFile,
  csvToMarkdown,
  docxToMarkdown,
  extractDataUriImages,
  markdownFileName,
  sanitizeName,
  xlsxToMarkdown,
} from "./lib/converters.js";
import { DriveClient } from "./lib/drive.js";

/* global mammoth, TurndownService, ExcelJS, JSZip */

const CONCURRENCY = 4;

const FRIENDLY_ERRORS = {
  cannotDownloadFile: "the owner has disabled downloading for this file",
  cannotExportFile: "the owner has disabled downloading for this file",
  exportSizeLimitExceeded: "too large for Google to export (10 MB limit)",
};

const params = new URLSearchParams(location.search);
const folderId = params.get("folder");
const includeOther = params.get("other") === "1";

const $ = (id) => document.getElementById(id);
const counts = { converted: 0, copied: 0, skipped: 0, failed: 0 };

function setStatus(text, isError = false) {
  $("status").textContent = text;
  $("status").classList.toggle("error", isError);
}

function log(kind, path, detail = "") {
  counts[kind]++;
  const li = document.createElement("li");
  li.className = kind;
  li.innerHTML = '<span class="tag"></span><span class="path"></span><span class="detail"></span>';
  li.querySelector(".tag").textContent = kind;
  li.querySelector(".path").textContent = path;
  li.querySelector(".detail").textContent = detail;
  // Problems first, so they are easy to spot in long exports.
  if (kind === "failed" || kind === "skipped") $("log").prepend(li);
  else $("log").append(li);
  $("counts").textContent =
    `${counts.converted} converted, ${counts.copied} copied, ` +
    `${counts.skipped} skipped, ${counts.failed} failed`;
}

async function getToken({ staleToken }) {
  if (staleToken) await chrome.identity.removeCachedAuthToken({ token: staleToken });
  const { token } = await chrome.identity.getAuthToken({ interactive: true });
  return token;
}

// Walks the folder tree and returns every file with its destination directory.
async function collectFiles(drive, rootDir, allocator) {
  const files = [];
  const visited = new Set([folderId]);
  const queue = [{ id: folderId, dir: rootDir }];
  while (queue.length) {
    const { id, dir } = queue.shift();
    setStatus(`Listing ${dir} ... (${files.length} files found so far)`);
    for (const f of await drive.listChildren(id)) {
      let file = f;
      if (f.mimeType === MIME.shortcut && f.shortcutDetails) {
        file = { ...f, id: f.shortcutDetails.targetId, mimeType: f.shortcutDetails.targetMimeType };
      }
      if (file.mimeType === MIME.folder) {
        if (f !== file) {
          log("skipped", `${dir}/${f.name}`, "shortcut to a folder");
        } else if (!visited.has(file.id)) {
          visited.add(file.id);
          queue.push({ id: file.id, dir: allocator.allocate(dir, sanitizeName(file.name)) });
        }
        continue;
      }
      files.push({ ...file, dir });
    }
  }
  return files;
}

function addAssets(zip, mdPath, assets) {
  const assetDir = mdPath.replace(/\.md$/, "_files");
  for (const a of assets) zip.file(`${assetDir}/${a.name}`, a.base64, { base64: true });
}

// Converts or copies one file into the zip. Returns [kind, path, detail] for the log.
async function processFile(drive, zip, allocator, file) {
  const kind = classifyFile(file);
  const shownPath = `${file.dir}/${file.name}`;

  if (kind === "other") {
    if (!includeOther) return ["skipped", shownPath, "not a document or spreadsheet"];
    if (file.mimeType.startsWith("application/vnd.google-apps.")) {
      return ["skipped", shownPath, "Google file type that can't be downloaded as a file"];
    }
    const path = allocator.allocate(file.dir, sanitizeName(file.name));
    zip.file(path, await drive.download(file.id));
    return ["copied", path, ""];
  }

  if (kind === "markdown") {
    let name = sanitizeName(file.name);
    if (!/\.(md|markdown|mdown|mkd)$/i.test(name)) name += ".md";
    const path = allocator.allocate(file.dir, name);
    zip.file(path, await drive.download(file.id));
    return ["copied", path, "unchanged"];
  }

  const path = allocator.allocate(file.dir, markdownFileName(file.name, kind));
  const assetDirName = path.split("/").pop().replace(/\.md$/, "_files");
  const title = kind === "gdoc" || kind === "gsheet" ? file.name : file.name.replace(/\.[^.]+$/, "");
  let markdown;
  let detail = "";

  switch (kind) {
    case "gdoc": {
      const exported = await drive.exportFile(file.id, "text/markdown", "text");
      const { markdown: md, assets } = extractDataUriImages(exported, assetDirName);
      addAssets(zip, path, assets);
      markdown = md;
      detail = "from Google Doc";
      break;
    }
    case "docx": {
      const data = await drive.download(file.id);
      const result = await docxToMarkdown({ mammoth, TurndownService }, data, assetDirName);
      addAssets(zip, path, result.assets);
      markdown = result.markdown;
      detail = result.warnings.length ? `from .docx (${result.warnings.join("; ")})` : "from .docx";
      break;
    }
    case "gsheet": {
      const data = await drive.exportFile(file.id, MIME.xlsx);
      markdown = await xlsxToMarkdown(ExcelJS, data, title);
      detail = "from Google Sheet";
      break;
    }
    case "xlsx": {
      const data = await drive.download(file.id);
      markdown = await xlsxToMarkdown(ExcelJS, data, title);
      detail = "from .xlsx";
      break;
    }
    case "csv": {
      markdown = csvToMarkdown(await drive.download(file.id, "text"), title);
      detail = "from .csv";
      break;
    }
  }
  zip.file(path, markdown);
  return ["converted", path, detail];
}

async function runPool(items, worker) {
  let next = 0;
  let done = 0;
  const progress = $("progress");
  progress.hidden = false;
  progress.max = items.length;
  progress.value = 0;
  await Promise.all(
    Array.from({ length: Math.min(CONCURRENCY, items.length) }, async () => {
      while (next < items.length) {
        const item = items[next++];
        await worker(item);
        progress.value = ++done;
        setStatus(`Converting files ... ${done} of ${items.length}`);
      }
    }),
  );
}

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  return chrome.downloads.download({ url, filename, conflictAction: "uniquify" });
}

async function main() {
  if (!folderId) throw new Error("No folder ID given. Open this page from the extension button.");
  if (chrome.runtime.getManifest().oauth2?.client_id?.startsWith("YOUR_CLIENT_ID")) {
    throw new Error(
      "The extension has no OAuth client ID yet. Follow the setup steps in README.md " +
        "and put your client ID in manifest.json.",
    );
  }

  const drive = new DriveClient(getToken);
  setStatus("Signing in to Google ...");
  const rootName = await drive.getFolderName(folderId);
  const rootDir = sanitizeName(rootName);
  document.title = `Exporting ${rootName}`;
  $("title").textContent = rootName;

  const allocator = new PathAllocator();
  const files = await collectFiles(drive, rootDir, allocator);
  const zip = new JSZip();

  window.onbeforeunload = () => "The export is still running.";
  await runPool(files, async (file) => {
    try {
      log(...(await processFile(drive, zip, allocator, file)));
    } catch (e) {
      log("failed", `${file.dir}/${file.name}`, FRIENDLY_ERRORS[e.reason] || e.message || String(e));
    }
  });

  const written = counts.converted + counts.copied;
  if (!written) {
    window.onbeforeunload = null;
    $("progress").hidden = true;
    setStatus("Nothing to export: no documents, spreadsheets or Markdown files were found.");
    return;
  }

  setStatus("Building ZIP ...");
  const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" }, (meta) => {
    $("progress").max = 100;
    $("progress").value = meta.percent;
  });
  window.onbeforeunload = null;
  $("progress").hidden = true;

  const filename = `${rootDir}.zip`;
  await download(blob, filename);
  setStatus(
    `Done. ${filename} was saved to your Downloads folder` +
      (counts.failed ? ` (${counts.failed} file(s) failed, see below).` : "."),
  );
  $("actions").hidden = false;
  $("download-again").onclick = () => download(blob, filename);
}

main().catch((e) => {
  window.onbeforeunload = null;
  $("progress").hidden = true;
  setStatus(`Export failed: ${e.message || e}`, true);
  console.error(e);
});
