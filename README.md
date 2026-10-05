# Drive Folder to Markdown

A Chrome extension that downloads everything in a Google Drive folder as Markdown.
Open a folder in Google Drive, click the extension button, and you get one ZIP with
the folder's full subfolder structure, every document converted to `.md`.

It is based on the Apps Script from
[vladimirzb/Google-Drive-Folder-Export-Google-Docs-Converter](https://github.com/vladimirzb/Google-Drive-Folder-Export-Google-Docs-Converter).
That script converts only native Google Docs and copies everything else. This extension
also converts Word and Excel files, and it runs in your browser instead of writing the
result to another Drive folder.

## What gets converted

| In Drive | In the ZIP | How |
| --- | --- | --- |
| Google Doc | `Name.md` | Google's own Markdown export (`text/markdown`) |
| `.docx` | `Name.md` | [mammoth](https://github.com/mwilliamson/mammoth.js) (docx → HTML) + [Turndown](https://github.com/mixmark-io/turndown) (HTML → Markdown) |
| Google Sheet | `Name.md` | exported as .xlsx, then converted like an .xlsx |
| `.xlsx` / `.xlsm` | `Name.md` | [ExcelJS](https://github.com/exceljs/exceljs): a `##` heading and a Markdown table per sheet |
| `.csv` | `Name.md` | one Markdown table |
| `.md` / `.markdown` | unchanged | downloaded byte for byte |
| anything else (PDF, images, ...) | skipped, or copied unchanged if you tick the option in the popup | |

Details:

- **Images** in Docs and .docx files are saved next to the Markdown file in a
  `Name_files/` folder and linked with relative paths. They are not embedded as base64.
- **Spreadsheets:** the first non-empty row becomes the table header. Cells show
  values, not formulas (the last value calculated for each formula). Dates are written
  as `YYYY-MM-DD`, percentages keep their `%` format, merged cells show their value
  once, and hidden sheets are labelled `(hidden)`. Charts, images and cell colours
  are dropped.
- **Word tables:** the first row is always used as the header row.
- **Names:** characters that aren't allowed in file names are replaced with `_`. When
  two files in a folder end up with the same name (Drive allows that), the second
  becomes `Name (2).md`.
- **Shortcuts** to files are followed. Shortcuts to folders are skipped, so the same
  folder isn't exported twice and shortcut loops can't happen.
- **Not converted:** legacy `.doc`/`.xls`, Google Slides, Forms and Drawings. They are
  listed as skipped on the results page.

The results page lists every file as converted, copied, skipped or failed. Skipped
and failed files appear at the top.

## Setup (once)

The extension reads Drive through the official Drive API, so it needs an OAuth client
ID from a Google Cloud project. This takes about five minutes.

1. **Load the extension.** Open `chrome://extensions`, turn on **Developer mode**,
   click **Load unpacked** and select this repository's folder. Copy the extension
   **ID** shown on its card (32 letters).
2. **Create a Google Cloud project** at <https://console.cloud.google.com/>, or reuse
   one you already have.
3. **Enable the Drive API:** *APIs & Services → Library → Google Drive API → Enable*.
4. **Configure the OAuth consent screen** (*APIs & Services → OAuth consent screen*,
   called *Google Auth Platform → Branding/Audience* in newer consoles):
   - If everyone who will use the extension is in your Google Workspace organization,
     choose **Internal**. Then Google doesn't need to review the app.
   - Otherwise choose **External**, leave it in *Testing*, and add each user's email
     under **Test users**.
   - Add the scope `https://www.googleapis.com/auth/drive.readonly`.
5. **Create the client ID:** *APIs & Services → Credentials → Create credentials →
   OAuth client ID*. For the application type, choose **Chrome extension** and paste
   the extension ID from step 1.
6. Put the client ID in `manifest.json`:
   ```json
   "oauth2": {
     "client_id": "1234567890-abc123.apps.googleusercontent.com",
     ...
   ```
7. Click the reload icon on the extension's card in `chrome://extensions`.

The first time you export, Chrome asks you to choose a Google account and allow
read-only Drive access.

### Keeping the same extension ID on other computers

An unpacked extension's ID depends on the folder it was loaded from, so a teammate who
loads it from another path gets a different ID, and the client ID won't work for them.
To keep the ID the same everywhere, add a `"key"` to `manifest.json`:

1. Pack the extension once with **Pack extension** on `chrome://extensions`. This
   creates a `.pem` private key. Keep the key private.
2. Get the public key:
   `openssl rsa -in drive-docx-to-md.pem -pubout -outform DER | openssl base64 -A`
3. Add `"key": "<that base64 string>"` to `manifest.json`, reload, and use the new ID
   for the OAuth client.

Publishing the extension privately to your organization in the Chrome Web Store also
gives it a fixed ID.

## Usage

1. Open a folder in Google Drive. Any folder you can view works, including
   `drive.google.com/drive/my-drive` and folders in shared drives.
2. Click the extension button, then **Export as Markdown**.
3. A tab opens and shows the progress. When the export finishes, `<Folder name>.zip` is
   saved to your Downloads folder. Keep the tab open until it says *Done*.

## Permissions

- `identity`: signs in to Google with the read-only Drive scope. The extension never
  changes or creates anything in Drive.
- `activeTab`: reads the URL of the current tab when you click the button, to find
  the folder ID.
- `downloads`: saves the ZIP.
- `storage`: remembers the "include other files" checkbox.
- `https://www.googleapis.com/*`: calls the Drive API.

All conversion happens locally in the browser. File contents are sent only between
your browser and Google.

## Development

```sh
npm install      # libraries used by the tests and by the vendor script
npm test         # unit tests for the converters (Node 20+)
npm run vendor   # after upgrading a dependency: copies browser builds into vendor/
```

- `src/popup.*`: the toolbar popup. It finds the folder ID in the tab URL.
- `src/export.*`: the export tab. It walks the folder tree, converts the files and builds
  the ZIP.
- `src/lib/drive.js`: the Drive API v3 client, with retries for rate limits and
  expired tokens.
- `src/lib/converters.js`: the pure conversion functions (docx/xlsx/csv → Markdown,
  image extraction, file names).
- `vendor/`: the browser builds of mammoth, Turndown, ExcelJS and JSZip. They are
  committed so you can load the extension without a build step.
