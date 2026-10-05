import { parseDriveFolderUrl } from "./lib/converters.js";

const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
const folderId = tab?.url ? parseDriveFolderUrl(tab.url) : null;

if (!folderId) {
  document.getElementById("not-folder").hidden = false;
} else {
  document.getElementById("folder").hidden = false;
  const title = (tab.title || "").replace(/ - Google Drive$/, "");
  document.getElementById("folder-name").textContent = title || folderId;

  const includeOther = document.getElementById("include-other");
  const { includeOther: saved = false } = await chrome.storage.sync.get("includeOther");
  includeOther.checked = saved;
  includeOther.addEventListener("change", () => {
    chrome.storage.sync.set({ includeOther: includeOther.checked });
  });

  document.getElementById("export").addEventListener("click", async () => {
    const params = new URLSearchParams({ folder: folderId, other: includeOther.checked ? "1" : "0" });
    await chrome.tabs.create({ url: chrome.runtime.getURL(`src/export.html?${params}`) });
    window.close();
  });
}
