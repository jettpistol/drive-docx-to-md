// Pure conversion helpers. Third-party libraries are passed in as arguments so
// this module runs both in the extension page (globals from vendor/) and in
// Node tests (npm packages).

export const MIME = {
  folder: "application/vnd.google-apps.folder",
  shortcut: "application/vnd.google-apps.shortcut",
  gdoc: "application/vnd.google-apps.document",
  gsheet: "application/vnd.google-apps.spreadsheet",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  markdown: "text/markdown",
};

const MARKDOWN_EXT = /\.(md|markdown|mdown|mkd)$/i;

// Returns the folder ID for a Google Drive folder URL, "root" for My Drive,
// or null when the URL is not a Drive folder.
export function parseDriveFolderUrl(url) {
  let u;
  try {
    u = new URL(url);
  } catch {
    return null;
  }
  if (u.hostname !== "drive.google.com") return null;
  const m = u.pathname.match(/\/folders\/([A-Za-z0-9_-]+)/);
  if (m) return m[1];
  if (/^\/drive(\/u\/\d+)?\/(my-drive|root)\/?$/.test(u.pathname)) return "root";
  if (/^\/(open|drive\/folderview)$/.test(u.pathname) && u.searchParams.get("id")) {
    return u.searchParams.get("id");
  }
  return null;
}

// Decides how a Drive file is handled.
export function classifyFile({ name, mimeType }) {
  if (mimeType === MIME.folder) return "folder";
  if (mimeType === MIME.shortcut) return "shortcut";
  if (mimeType === MIME.gdoc) return "gdoc";
  if (mimeType === MIME.gsheet) return "gsheet";
  if (MARKDOWN_EXT.test(name) || mimeType === MIME.markdown || mimeType === "text/x-markdown") {
    return "markdown";
  }
  if (mimeType === MIME.docx || /\.docx$/i.test(name)) return "docx";
  if (mimeType === MIME.xlsx || mimeType === MIME.xlsm || /\.xlsx$|\.xlsm$/i.test(name)) return "xlsx";
  if (mimeType === "text/csv" || /\.csv$/i.test(name)) return "csv";
  return "other";
}

// Makes a Drive name safe to use as a path segment on Windows, macOS and Linux.
export function sanitizeName(name) {
  let s = String(name)
    .replace(/[\u0000-\u001f<>:"/\\|?*]/g, "_")
    .replace(/[. ]+$/, "")
    .trim();
  if (!s || s === "." || s === "..") s = "untitled";
  if (/^(con|prn|aux|nul|com\d|lpt\d)(\.|$)/i.test(s)) s = "_" + s;
  return s.slice(0, 200);
}

// Output file name for a converted file: "Report.docx" -> "Report.md".
export function markdownFileName(name, kind) {
  const base = kind === "gdoc" || kind === "gsheet" ? name : name.replace(/\.[^.]+$/, "") || name;
  return sanitizeName(base) + ".md";
}

// Drive allows several files with the same name in one folder; a zip does not.
// Hands out unique (case-insensitive) paths by appending " (2)", " (3)", ...
export class PathAllocator {
  constructor() {
    this.used = new Set();
  }

  allocate(dir, fileName) {
    const dot = fileName.lastIndexOf(".");
    const stem = dot > 0 ? fileName.slice(0, dot) : fileName;
    const ext = dot > 0 ? fileName.slice(dot) : "";
    for (let n = 1; ; n++) {
      const candidate = n === 1 ? fileName : `${stem} (${n})${ext}`;
      const path = dir ? `${dir}/${candidate}` : candidate;
      if (!this.used.has(path.toLowerCase())) {
        this.used.add(path.toLowerCase());
        return path;
      }
    }
  }
}

const IMAGE_EXT = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/jpg": "jpg",
  "image/gif": "gif",
  "image/webp": "webp",
  "image/svg+xml": "svg",
  "image/bmp": "bmp",
  "image/tiff": "tiff",
  "image/x-emf": "emf",
  "image/x-wmf": "wmf",
};

function imageExt(contentType) {
  return IMAGE_EXT[String(contentType).toLowerCase()] || "bin";
}

function encodePath(path) {
  return path.split("/").map(encodeURIComponent).join("/");
}

// Google Docs' Markdown export embeds images as base64 data URIs, which makes
// files huge and unreadable. Moves them into `assetDir` and links them instead.
// Returns { markdown, assets: [{ name, base64 }] } with names relative to assetDir.
export function extractDataUriImages(markdown, assetDir) {
  const assets = [];
  const re = /data:(image\/[\w.+-]+);base64,([A-Za-z0-9+/=\s]+?)(?=[>)\s"'])/g;
  const out = markdown.replace(re, (_, contentType, data) => {
    const name = `image-${assets.length + 1}.${imageExt(contentType)}`;
    assets.push({ name, base64: data.replace(/\s+/g, "") });
    return encodePath(`${assetDir}/${name}`);
  });
  return { markdown: out, assets };
}

// ---------------------------------------------------------------------------
// Spreadsheets

function escapeCell(text) {
  return String(text)
    .replace(/\r\n?/g, "\n")
    .trim()
    .replace(/\\/g, "\\\\")
    .replace(/\|/g, "\\|")
    .replace(/\n/g, "<br>");
}

// Renders rows (arrays of strings) as a GFM table. The first non-empty row is
// the header. Trailing empty rows/columns are trimmed.
export function rowsToMarkdownTable(rows) {
  rows = rows.map((r) => r.map((c) => (c == null ? "" : String(c))));
  while (rows.length && rows[rows.length - 1].every((c) => c.trim() === "")) rows.pop();
  while (rows.length && rows[0].every((c) => c.trim() === "")) rows.shift();
  if (!rows.length) return "";

  let width = 0;
  for (const r of rows) {
    for (let i = r.length - 1; i >= 0; i--) {
      if (r[i].trim() !== "") {
        width = Math.max(width, i + 1);
        break;
      }
    }
  }
  const line = (r) => {
    const cells = [];
    for (let i = 0; i < width; i++) cells.push(escapeCell(r[i] ?? ""));
    return `| ${cells.join(" | ")} |`;
  };
  const [header, ...body] = rows;
  return [line(header), `| ${Array(width).fill("---").join(" | ")} |`, ...body.map(line)].join("\n");
}

function pad(n) {
  return String(n).padStart(2, "0");
}

function formatDate(d) {
  if (isNaN(d)) return "";
  const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const h = d.getUTCHours(), m = d.getUTCMinutes(), s = d.getUTCSeconds();
  if (!h && !m && !s) return date;
  return `${date} ${pad(h)}:${pad(m)}${s ? ":" + pad(s) : ""}`;
}

function formatNumber(n, numFmt) {
  if (!Number.isFinite(n)) return String(n);
  if (numFmt && /%/.test(numFmt.replace(/"[^"]*"/g, ""))) {
    const decimals = (numFmt.match(/\.(0+)/) || ["", ""])[1].length;
    return (n * 100).toFixed(decimals) + "%";
  }
  // Drop floating-point noise such as 0.30000000000000004.
  return String(Number(n.toPrecision(15)));
}

// Converts an ExcelJS cell value into display text.
export function cellText(value, numFmt) {
  if (value == null) return "";
  if (value instanceof Date) return formatDate(value);
  if (typeof value === "number") return formatNumber(value, numFmt);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return value;
  if (typeof value === "object") {
    if ("error" in value) return String(value.error);
    if (Array.isArray(value.richText)) return value.richText.map((t) => t.text).join("");
    if ("hyperlink" in value) {
      const text = cellText(value.text, numFmt) || value.hyperlink;
      return value.hyperlink ? `[${text}](${value.hyperlink})` : text;
    }
    if ("formula" in value || "sharedFormula" in value) return cellText(value.result, numFmt);
  }
  return String(value);
}

// Reads an ExcelJS worksheet into rows of display strings.
export function worksheetToRows(ws) {
  const rows = [];
  const cols = ws.columnCount;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const cells = [];
    for (let c = 1; c <= cols; c++) {
      const cell = row.getCell(c);
      // Only the top-left cell of a merged range carries the value.
      cells.push(cell.isMerged && cell.master !== cell ? "" : cellText(cell.value, cell.numFmt));
    }
    rows.push(cells);
  }
  return rows;
}

// Converts an .xlsx file (ArrayBuffer) to Markdown: one heading + table per sheet.
export async function xlsxToMarkdown(ExcelJS, data, title) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);
  const parts = [`# ${title}`];
  wb.eachSheet((ws) => {
    const hidden = ws.state && ws.state !== "visible" ? " (hidden)" : "";
    const table = rowsToMarkdownTable(worksheetToRows(ws));
    parts.push(`## ${ws.name}${hidden}`, table || "_(empty sheet)_");
  });
  return parts.join("\n\n") + "\n";
}

// Minimal RFC 4180 CSV parser.
export function parseCsv(text) {
  text = text.replace(/^﻿/, "");
  const rows = [];
  let row = [], field = "", quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (ch === '"') {
        quoted = false;
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === ",") {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += ch;
    }
  }
  if (field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

export function csvToMarkdown(text, title) {
  const table = rowsToMarkdownTable(parseCsv(text));
  return `# ${title}\n\n${table || "_(empty file)_"}\n`;
}

// ---------------------------------------------------------------------------
// Word documents

// Turndown rules for tables. Unlike turndown-plugin-gfm, the first row is always
// used as the header, because Word tables rarely mark header rows.
function addTableRules(turndown) {
  const isFirstRow = (tr) => {
    let table = tr.parentNode;
    while (table && table.nodeName !== "TABLE") table = table.parentNode;
    return table && table.querySelector("tr") === tr;
  };
  turndown.addRule("tableCell", {
    filter: ["th", "td"],
    replacement: (content) => ` ${escapeCellMarkdown(content)} |`,
  });
  turndown.addRule("tableRow", {
    filter: "tr",
    replacement: (content, node) => {
      let out = `\n|${content}`;
      if (isFirstRow(node)) {
        const n = node.querySelectorAll("th, td").length;
        out += `\n|${" --- |".repeat(n)}`;
      }
      return out;
    },
  });
  turndown.addRule("tableSection", {
    filter: ["thead", "tbody", "tfoot"],
    replacement: (content) => content,
  });
  turndown.addRule("table", {
    filter: "table",
    replacement: (content) => `\n\n${content.replace(/^\n+/, "")}\n\n`,
  });
  turndown.addRule("strikethrough", {
    filter: ["del", "s", "strike"],
    replacement: (content) => `~~${content}~~`,
  });
}

function escapeCellMarkdown(content) {
  return content
    .trim()
    .replace(/\n+/g, "<br>")
    .replace(/\|/g, "\\|");
}

export function createTurndown(TurndownService) {
  const turndown = new TurndownService({
    headingStyle: "atx",
    codeBlockStyle: "fenced",
    bulletListMarker: "-",
    emDelimiter: "*",
  });
  addTableRules(turndown);
  return turndown;
}

// Converts a .docx file (ArrayBuffer) to Markdown. Embedded images are written
// to `assetDir` (relative to the Markdown file).
// Returns { markdown, assets: [{ name, base64 }], warnings: string[] }.
export async function docxToMarkdown({ mammoth, TurndownService }, data, assetDir) {
  const assets = [];
  const convertImage = mammoth.images.imgElement(async (image) => {
    const name = `image-${assets.length + 1}.${imageExt(image.contentType)}`;
    assets.push({ name, base64: await image.read("base64") });
    return { src: encodePath(`${assetDir}/${name}`) };
  });
  // Node's mammoth build wants a Buffer; the browser build an ArrayBuffer.
  const input = typeof Buffer !== "undefined" ? { buffer: Buffer.from(data) } : { arrayBuffer: data };
  const result = await mammoth.convertToHtml(input, { convertImage });
  const markdown = createTurndown(TurndownService).turndown(result.value || "");
  const warnings = result.messages.filter((m) => m.type === "error").map((m) => m.message);
  return { markdown: markdown.trim() + "\n", assets, warnings };
}
