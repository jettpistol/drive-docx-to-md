import { test } from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import mammoth from "mammoth";
import TurndownService from "turndown";
import {
  PathAllocator,
  classifyFile,
  csvToMarkdown,
  docxToMarkdown,
  extractDataUriImages,
  markdownFileName,
  parseCsv,
  parseDriveFolderUrl,
  rowsToMarkdownTable,
  sanitizeName,
  xlsxToMarkdown,
} from "../src/lib/converters.js";

test("parseDriveFolderUrl", () => {
  assert.equal(parseDriveFolderUrl("https://drive.google.com/drive/folders/1AbC-_x"), "1AbC-_x");
  assert.equal(parseDriveFolderUrl("https://drive.google.com/drive/u/1/folders/1AbC?usp=sharing"), "1AbC");
  assert.equal(parseDriveFolderUrl("https://drive.google.com/drive/my-drive"), "root");
  assert.equal(parseDriveFolderUrl("https://drive.google.com/drive/u/0/my-drive"), "root");
  assert.equal(parseDriveFolderUrl("https://drive.google.com/open?id=XYZ"), "XYZ");
  assert.equal(parseDriveFolderUrl("https://drive.google.com/drive/recent"), null);
  assert.equal(parseDriveFolderUrl("https://docs.google.com/document/d/abc/edit"), null);
  assert.equal(parseDriveFolderUrl("not a url"), null);
});

test("classifyFile", () => {
  const c = (name, mimeType) => classifyFile({ name, mimeType });
  assert.equal(c("Doc", "application/vnd.google-apps.document"), "gdoc");
  assert.equal(c("Sheet", "application/vnd.google-apps.spreadsheet"), "gsheet");
  assert.equal(c("a.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"), "docx");
  assert.equal(c("a.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), "xlsx");
  assert.equal(c("notes.md", "application/octet-stream"), "markdown");
  assert.equal(c("README.markdown", "text/plain"), "markdown");
  assert.equal(c("data.csv", "text/csv"), "csv");
  assert.equal(c("scan.pdf", "application/pdf"), "other");
});

test("file names", () => {
  assert.equal(sanitizeName('a/b:c*?"<>|'), "a_b_c______");
  assert.equal(sanitizeName("trailing. "), "trailing");
  assert.equal(sanitizeName(""), "untitled");
  assert.equal(sanitizeName("CON"), "_CON");
  assert.equal(markdownFileName("Report.docx", "docx"), "Report.md");
  assert.equal(markdownFileName("Q3 v1.2", "gdoc"), "Q3 v1.2.md");

  const alloc = new PathAllocator();
  assert.equal(alloc.allocate("Root", "Report.md"), "Root/Report.md");
  assert.equal(alloc.allocate("Root", "report.md"), "Root/report (2).md");
  assert.equal(alloc.allocate("Root", "Report.md"), "Root/Report (3).md");
  assert.equal(alloc.allocate("Root/Sub", "Report.md"), "Root/Sub/Report.md");
});

test("extractDataUriImages moves base64 images out of Google Docs markdown", () => {
  const md = "Hello\n\n![][image1]\n\n[image1]: <data:image/png;base64,iVBORw0KGgo=>\n";
  const { markdown, assets } = extractDataUriImages(md, "My Doc_files");
  assert.equal(markdown, "Hello\n\n![][image1]\n\n[image1]: <My%20Doc_files/image-1.png>\n");
  assert.deepEqual(assets, [{ name: "image-1.png", base64: "iVBORw0KGgo=" }]);
});

test("rowsToMarkdownTable", () => {
  const md = rowsToMarkdownTable([
    ["", "", ""],
    ["Name", "Note", ""],
    ["a|b", "line1\nline2", ""],
    ["", "", ""],
  ]);
  assert.equal(md, "| Name | Note |\n| --- | --- |\n| a\\|b | line1<br>line2 |");
  assert.equal(rowsToMarkdownTable([["", ""]]), "");
});

test("parseCsv / csvToMarkdown", () => {
  assert.deepEqual(parseCsv('a,"b ""q""",c\r\n1,"x,y",\n'), [["a", 'b "q"', "c"], ["1", "x,y", ""]]);
  assert.equal(csvToMarkdown("h1,h2\n1,2\n", "Data"), "# Data\n\n| h1 | h2 |\n| --- | --- |\n| 1 | 2 |\n");
});

test("xlsxToMarkdown renders one table per sheet with formatted values", async () => {
  const wb = new ExcelJS.Workbook();
  const s1 = wb.addWorksheet("Inventory");
  s1.addRow(["SKU", "Qty", "Price", "Date", "Share", "Total"]);
  s1.addRow(["A-1", 3, 0.1 + 0.2, new Date(Date.UTC(2024, 0, 15)), 0.256, { formula: "B2*C2", result: 0.9 }]);
  s1.addRow([{ richText: [{ text: "Bold" }, { text: " part" }] }, true, null, null, null, { text: "link", hyperlink: "https://x.test" }]);
  s1.getCell("E2").numFmt = "0.0%";
  wb.addWorksheet("Empty");
  const s3 = wb.addWorksheet("Merged");
  s3.addRow(["Header", ""]);
  s3.addRow(["x", "y"]);
  s3.mergeCells("A1:B1");
  const buf = await wb.xlsx.writeBuffer();

  const md = await xlsxToMarkdown(ExcelJS, buf, "Stock");
  assert.equal(
    md,
    [
      "# Stock",
      "## Inventory",
      "| SKU | Qty | Price | Date | Share | Total |\n" +
        "| --- | --- | --- | --- | --- | --- |\n" +
        "| A-1 | 3 | 0.3 | 2024-01-15 | 25.6% | 0.9 |\n" +
        "| Bold part | TRUE |  |  |  | [link](https://x.test) |",
      "## Empty",
      "_(empty sheet)_",
      "## Merged",
      "| Header |  |\n| --- | --- |\n| x | y |",
    ].join("\n\n") + "\n",
  );
});

// Builds a minimal .docx with a heading, formatted text, a list-free table and an image.
async function makeDocx() {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.png"/>
</Relationships>`,
  );
  zip.file("word/media/pic.png", "iVBORw0KGgo=", { base64: true });
  const p = (text, style = "") =>
    `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const cell = (text) => `<w:tc>${p(text)}</w:tc>`;
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body>
${p("Title", "Heading1")}
<w:p><w:r><w:t xml:space="preserve">Some </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> text.</w:t></w:r></w:p>
<w:tbl>
<w:tr>${cell("SKU")}${cell("Qty")}</w:tr>
<w:tr>${cell("A|1")}${cell("5")}</w:tr>
</w:tbl>
<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1" descr="logo"/>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
<pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill>
</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
</w:body>
</w:document>`,
  );
  return zip.generateAsync({ type: "nodebuffer" });
}

test("docxToMarkdown converts headings, formatting, tables and images", async () => {
  const { markdown, assets } = await docxToMarkdown({ mammoth, TurndownService }, await makeDocx(), "Spec_files");
  assert.equal(
    markdown,
    "# Title\n\nSome **bold** text.\n\n| SKU | Qty |\n| --- | --- |\n| A\\|1 | 5 |\n\n![logo](Spec_files/image-1.png)\n",
  );
  assert.deepEqual(assets, [{ name: "image-1.png", base64: "iVBORw0KGgo=" }]);
});
