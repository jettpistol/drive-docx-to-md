// Small Google Drive API v3 client with retries for rate limits and expired tokens.

const API = "https://www.googleapis.com/drive/v3";
const LIST_FIELDS =
  "nextPageToken,files(id,name,mimeType,size,shortcutDetails(targetId,targetMimeType))";

export class DriveError extends Error {
  constructor(status, message, reason) {
    super(message);
    this.status = status;
    this.reason = reason;
  }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export class DriveClient {
  // getToken({ staleToken }) must resolve to an OAuth access token. When
  // staleToken is set, that token was rejected and must not be returned again.
  constructor(getToken) {
    this.getToken = getToken;
    this.token = null;
  }

  async request(path, params = {}, as = "json") {
    const url = new URL(API + path);
    for (const [k, v] of Object.entries({ supportsAllDrives: true, ...params })) {
      url.searchParams.set(k, v);
    }

    let refreshed = false;
    for (let attempt = 0; ; attempt++) {
      if (!this.token) this.token = await this.getToken({});
      const res = await fetch(url, { headers: { Authorization: `Bearer ${this.token}` } });
      if (res.ok) {
        if (as === "json") return res.json();
        if (as === "text") return res.text();
        return res.arrayBuffer();
      }

      let message = `${res.status} ${res.statusText}`;
      let reason;
      try {
        const body = await res.json();
        message = body.error?.message || message;
        reason = body.error?.errors?.[0]?.reason;
      } catch {
        // Non-JSON error body.
      }

      if (res.status === 401 && !refreshed) {
        refreshed = true;
        this.token = await this.getToken({ staleToken: this.token });
        continue;
      }
      const rateLimited =
        res.status === 429 ||
        (res.status === 403 && /rateLimitExceeded|userRateLimitExceeded/.test(reason || ""));
      if ((rateLimited || res.status >= 500) && attempt < 6) {
        await sleep(2 ** attempt * 1000 + Math.random() * 1000);
        continue;
      }
      throw new DriveError(res.status, message, reason);
    }
  }

  // Name of a folder, a shared drive root, or "root" (My Drive).
  async getFolderName(id) {
    try {
      const f = await this.request(`/files/${encodeURIComponent(id)}`, { fields: "id,name" });
      return f.name;
    } catch (e) {
      if (e.status !== 404) throw e;
      // Shared drive IDs are also valid folder IDs but may not resolve as files.
      const d = await this.request(`/drives/${encodeURIComponent(id)}`, { fields: "name" });
      return d.name;
    }
  }

  async listChildren(folderId) {
    const files = [];
    let pageToken;
    do {
      const page = await this.request("/files", {
        q: `'${folderId.replace(/'/g, "\\'")}' in parents and trashed = false`,
        fields: LIST_FIELDS,
        pageSize: 1000,
        orderBy: "folder,name",
        includeItemsFromAllDrives: true,
        ...(pageToken ? { pageToken } : {}),
      });
      files.push(...page.files);
      pageToken = page.nextPageToken;
    } while (pageToken);
    return files;
  }

  // Exports a Google Workspace file (Docs, Sheets, ...) to the given MIME type.
  exportFile(id, mimeType, as = "arrayBuffer") {
    return this.request(`/files/${encodeURIComponent(id)}/export`, { mimeType }, as);
  }

  // Downloads a regular (non-Google) file's bytes.
  download(id, as = "arrayBuffer") {
    return this.request(`/files/${encodeURIComponent(id)}`, { alt: "media" }, as);
  }
}
