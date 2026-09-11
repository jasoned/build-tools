(async function exportCanvasCourseFilesToJson() {
  const existing = document.getElementById("canvas-file-export-panel");
  if (existing) existing.remove();

  const courseId =
    (window.location.pathname.match(/\/courses\/(\d+)/) || [])[1] ||
    (window.ENV && window.ENV.COURSE_ID);

  if (!courseId) {
    alert("Could not find a Canvas course ID. Run this from inside a Canvas course.");
    return;
  }

  const canvasOrigin = window.location.origin;

  function createPanel() {
    const panel = document.createElement("div");
    panel.id = "canvas-file-export-panel";
    panel.style.cssText = `
      position: fixed;
      top: 80px;
      right: 30px;
      width: 420px;
      max-height: 70vh;
      overflow: auto;
      background: #fff;
      color: #222;
      border: 1px solid #ccc;
      border-radius: 10px;
      box-shadow: 0 8px 30px rgba(0,0,0,.25);
      z-index: 999999;
      font-family: Arial, sans-serif;
      padding: 16px;
    `;

    panel.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; gap:12px;">
        <h2 style="font-size:18px; margin:0;">Course File JSON Export</h2>
        <button id="cfe-close" style="border:0; background:#eee; border-radius:6px; padding:4px 8px; cursor:pointer;">Close</button>
      </div>

      <p style="font-size:13px; line-height:1.4; margin:12px 0;">
        Course ID: <strong>${courseId}</strong><br>
        This reads course files and folders, then downloads a JSON file.
      </p>

      <button id="cfe-start" style="
        width:100%;
        padding:10px;
        border:0;
        border-radius:6px;
        background:#2d3b45;
        color:#fff;
        cursor:pointer;
        font-weight:bold;
      ">Export files to JSON</button>

      <div id="cfe-progress-wrap" style="margin-top:14px; display:none;">
        <div style="height:10px; background:#eee; border-radius:10px; overflow:hidden;">
          <div id="cfe-progress" style="height:10px; width:0%; background:#2d3b45;"></div>
        </div>
      </div>

      <pre id="cfe-log" style="
        margin-top:12px;
        padding:10px;
        background:#f7f7f7;
        border:1px solid #ddd;
        border-radius:6px;
        white-space:pre-wrap;
        font-size:12px;
        line-height:1.4;
        max-height:320px;
        overflow:auto;
      ">Ready.</pre>
    `;

    document.body.appendChild(panel);

    document.getElementById("cfe-close").addEventListener("click", () => panel.remove());

    return {
      panel,
      startButton: document.getElementById("cfe-start"),
      logBox: document.getElementById("cfe-log"),
      progressWrap: document.getElementById("cfe-progress-wrap"),
      progressBar: document.getElementById("cfe-progress")
    };
  }

  const ui = createPanel();

  function log(message) {
    ui.logBox.textContent += `\n${message}`;
    ui.logBox.scrollTop = ui.logBox.scrollHeight;
  }

  function setProgress(percent) {
    ui.progressWrap.style.display = "block";
    ui.progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
  }

  function getNextLink(linkHeader) {
    if (!linkHeader) return null;

    const links = linkHeader.split(",");
    for (const link of links) {
      const sections = link.split(";").map(part => part.trim());
      const urlPart = sections[0];
      const relPart = sections.find(part => part === 'rel="next"' || part === "rel=next");

      if (relPart && urlPart.startsWith("<") && urlPart.endsWith(">")) {
        return urlPart.slice(1, -1);
      }
    }

    return null;
  }

  async function canvasApiRequest(pathOrUrl) {
    const url = pathOrUrl.startsWith("http")
      ? pathOrUrl
      : `${canvasOrigin}${pathOrUrl}`;

    const response = await fetch(url, {
      method: "GET",
      credentials: "same-origin",
      headers: {
        "Accept": "application/json"
      }
    });

    const responseText = await response.text();

    let data = null;
    if (responseText) {
      try {
        data = JSON.parse(responseText);
      } catch {
        data = responseText;
      }
    }

    if (!response.ok) {
      const detail = typeof data === "string" ? data : JSON.stringify(data, null, 2);
      throw new Error(`Canvas API error ${response.status}\n${detail}`);
    }

    return {
      data,
      response
    };
  }

  async function getAllPaginated(path) {
    let url = path;
    const allItems = [];
    let pageCount = 0;

    while (url) {
      pageCount++;
      const { data, response } = await canvasApiRequest(url);

      if (!Array.isArray(data)) {
        throw new Error(`Expected an array from ${url}, but Canvas returned something else.`);
      }

      allItems.push(...data);
      url = getNextLink(response.headers.get("Link"));

      log(`Fetched page ${pageCount}. Running total: ${allItems.length}`);
    }

    return allItems;
  }

  function safeFolderPath(folder) {
    if (!folder) return null;

    return (
      folder.full_name ||
      folder.name ||
      null
    );
  }

  function buildCanvasFileUrl(fileId) {
    return `${canvasOrigin}/courses/${courseId}/files/${fileId}?wrap=1`;
  }

  function buildApiFileUrl(fileId) {
    return `${canvasOrigin}/api/v1/files/${fileId}`;
  }

  function downloadJson(data, filename) {
    const blob = new Blob([JSON.stringify(data, null, 2)], {
      type: "application/json"
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();

    link.remove();
    URL.revokeObjectURL(url);
  }

  async function runExport() {
    ui.startButton.disabled = true;
    ui.startButton.textContent = "Exporting...";
    ui.logBox.textContent = "Starting export...";
    setProgress(5);

    try {
      log("Fetching folders...");
      const folders = await getAllPaginated(`/api/v1/courses/${courseId}/folders?per_page=100`);
      setProgress(35);

      const folderMap = new Map();
      folders.forEach(folder => {
        folderMap.set(Number(folder.id), folder);
      });

      log(`Folders found: ${folders.length}`);
      log("Fetching files...");
      const files = await getAllPaginated(`/api/v1/courses/${courseId}/files?per_page=100`);
      setProgress(75);

      log(`Files found: ${files.length}`);
      log("Building JSON...");

      const exportedFolders = folders.map(folder => ({
        id: folder.id,
        name: folder.name || null,
        full_name: folder.full_name || null,
        parent_folder_id: folder.parent_folder_id || null,
        context_type: folder.context_type || null,
        context_id: folder.context_id || null,
        files_count: folder.files_count ?? null,
        folders_count: folder.folders_count ?? null,
        position: folder.position ?? null,
        locked: folder.locked ?? null,
        hidden: folder.hidden ?? null,
        for_submissions: folder.for_submissions ?? null,
        created_at: folder.created_at || null,
        updated_at: folder.updated_at || null
      }));

      const exportedFiles = files.map(file => {
        const folderId = file.folder_id ? Number(file.folder_id) : null;
        const folder = folderId ? folderMap.get(folderId) : null;

        return {
          id: file.id,
          uuid: file.uuid || null,

          display_name: file.display_name || null,
          filename: file.filename || null,

          folder_id: folderId,
          folder_name: folder ? folder.name || null : null,
          folder_full_name: safeFolderPath(folder),

          canvas_url: buildCanvasFileUrl(file.id),
          download_url: file.url || null,
          api_url: buildApiFileUrl(file.id),
          preview_url: file.preview_url || null,
          thumbnail_url: file.thumbnail_url || null,

          size_bytes: file.size ?? null,
          content_type: file["content-type"] || file.content_type || null,
          mime_class: file.mime_class || null,

          created_at: file.created_at || null,
          updated_at: file.updated_at || null,
          modified_at: file.modified_at || null,

          locked: file.locked ?? null,
          hidden: file.hidden ?? null,
          unlock_at: file.unlock_at || null,
          lock_at: file.lock_at || null,

          locked_for_user: file.locked_for_user ?? null,
          hidden_for_user: file.hidden_for_user ?? null,

          media_entry_id: file.media_entry_id || null
        };
      });

      exportedFiles.sort((a, b) => {
        const folderA = a.folder_full_name || "";
        const folderB = b.folder_full_name || "";
        const nameA = a.display_name || a.filename || "";
        const nameB = b.display_name || b.filename || "";

        return folderA.localeCompare(folderB) || nameA.localeCompare(nameB);
      });

      const exportData = {
        exported_at: new Date().toISOString(),
        canvas_domain: canvasOrigin,
        course_id: Number(courseId),
        course_url: `${canvasOrigin}/courses/${courseId}`,

        counts: {
          folders: exportedFolders.length,
          files: exportedFiles.length
        },

        folders: exportedFolders,
        files: exportedFiles
      };

      const dateStamp = new Date().toISOString().replace(/[:.]/g, "-");
      const filename = `canvas-course-${courseId}-files-${dateStamp}.json`;

      downloadJson(exportData, filename);

      setProgress(100);
      log("");
      log("Finished.");
      log(`Downloaded: ${filename}`);
      log(`Folders exported: ${exportedFolders.length}`);
      log(`Files exported: ${exportedFiles.length}`);

      ui.startButton.textContent = "Export complete";
    } catch (error) {
      console.error(error);
      log("");
      log("Export failed.");
      log(error.message || String(error));

      ui.startButton.disabled = false;
      ui.startButton.textContent = "Try again";
      setProgress(0);
    }
  }

  ui.startButton.addEventListener("click", runExport);
})();
