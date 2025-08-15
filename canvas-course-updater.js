(() => {
  // =========================
  // Canvas Course Updater UI
  // =========================

  // --- Util: CSRF token (as requested) ---
  function getCsrfToken() {
    const csrfRegex = new RegExp('^_csrf_token=(.*)$');
    let csrf;
    const cookies = document.cookie.split(';');
    for (let i = 0; i < cookies.length; i++) {
      const cookie = cookies[i].trim();
      const match = csrfRegex.exec(cookie);
      if (match) {
        csrf = decodeURIComponent(match[1]);
        break;
      }
    }
    return csrf;
  }

  // --- Inject UI ---
  const panel = document.createElement('div');
  panel.id = 'canvas-course-updater-panel';
  panel.innerHTML = `
    <style>
      #canvas-course-updater-panel {
        all: initial;
        font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, "Apple Color Emoji", "Segoe UI Emoji";
        position: fixed;
        z-index: 999999999;
        right: 20px;
        bottom: 20px;
        width: 420px;
        max-height: 80vh;
        background: #fff;
        border: 1px solid rgba(0,0,0,.08);
        box-shadow: 0 10px 30px rgba(0,0,0,.15);
        border-radius: 12px;
        display: grid;
        grid-template-rows: auto auto 1fr auto;
        overflow: hidden;
      }
      #ccu-hdr {
        padding: 12px 14px;
        background: linear-gradient(180deg, #f8fafc, #f3f4f6);
        border-bottom: 1px solid #e5e7eb;
        font-weight: 600;
        color: #111827;
        display: flex;
        align-items: center;
        gap: 8px;
      }
      #ccu-controls {
        padding: 10px 14px;
        display: grid;
        grid-template-columns: 1fr auto;
        gap: 8px;
        align-items: center;
        border-bottom: 1px solid #f1f5f9;
      }
      #ccu-controls label {
        font-size: 12px;
        color: #334155;
        display: block;
        margin-bottom: 6px;
      }
      #ccu-logo-wrap {
        grid-column: 1 / -1;
      }
      #ccu-logo {
        width: 100%;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 8px;
        background: #fff;
      }
      #ccu-start {
        grid-column: 2 / 3;
        justify-self: end;
        padding: 8px 12px;
        background: #0ea5e9;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        cursor: pointer;
      }
      #ccu-start[disabled] {
        opacity: .6;
        cursor: not-allowed;
      }
      #ccu-log {
        padding: 10px 14px;
        background: #fafafa;
        font-size: 12px;
        color: #111827;
        overflow: auto;
        white-space: pre-wrap;
        border-top: 1px solid #f1f5f9;
      }
      .ccu-line { margin: 0 0 4px 0; }
      .ccu-ok { color: #065f46; }
      .ccu-warn { color: #92400e; }
      .ccu-err { color: #991b1b; }
      #ccu-foot {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        padding: 10px 14px;
        border-top: 1px solid #e5e7eb;
      }
      #ccu-spinner {
        display: none;
        width: 18px;
        height: 18px;
        border: 3px solid #e5e7eb;
        border-top-color: #0ea5e9;
        border-radius: 50%;
        animation: ccu-spin 1s linear infinite;
      }
      @keyframes ccu-spin { to { transform: rotate(360deg); } }
      #ccu-close {
        background: transparent;
        border: none;
        color: #6b7280;
        font-weight: 600;
        cursor: pointer;
      }
      #ccu-toast {
        position: fixed;
        right: 20px;
        bottom: 110px;
        z-index: 999999999;
        background: #16a34a;
        color: white;
        padding: 10px 12px;
        border-radius: 8px;
        font-weight: 600;
        display: none;
        box-shadow: 0 10px 25px rgba(0,0,0,.15);
      }
      #ccu-meta {
        font-size: 11px;
        color: #64748b;
      }
      #ccu-mini {
        font-size: 11px;
        color: #94a3b8;
      }
    </style>
    <div id="ccu-hdr">📄 Canvas Course Updater</div>
    <div id="ccu-controls">
      <div id="ccu-logo-wrap">
        <label for="ccu-logo">Select School Logo</label>
        <select id="ccu-logo">
          <option value="">No logo copy</option>
          <option value="2943463">CSFS (2943463)</option>
          <option value="2943466">CSPP (2943466)</option>
          <option value="2943464">CSOE (2943464)</option>
          <option value="2943465">CSML (2943465)</option>
          <option value="2943467">SNHS (2943467)</option>
        </select>
      </div>
      <div id="ccu-mini">Updates existing pages only</div>
      <button id="ccu-start">Start Processing</button>
    </div>
    <div id="ccu-log" aria-live="polite"></div>
    <div id="ccu-foot">
      <div style="display:flex;align-items:center;gap:8px;">
        <div id="ccu-spinner" aria-hidden="true"></div>
        <div id="ccu-meta"></div>
      </div>
      <button id="ccu-close" title="Close">✕</button>
    </div>
  `;
  document.body.appendChild(panel);
  const el = {
    log: panel.querySelector('#ccu-log'),
    start: panel.querySelector('#ccu-start'),
    spinner: panel.querySelector('#ccu-spinner'),
    meta: panel.querySelector('#ccu-meta'),
    close: panel.querySelector('#ccu-close'),
    logo: panel.querySelector('#ccu-logo'),
  };

  const toast = document.createElement('div');
  toast.id = 'ccu-toast';
  toast.textContent = 'All done!';
  document.body.appendChild(toast);

  function showToast(msg, ok=true) {
    toast.textContent = msg || 'Done';
    toast.style.background = ok ? '#16a34a' : '#b91c1c';
    toast.style.display = 'block';
    setTimeout(() => toast.style.display = 'none', 3000);
  }

  function log(msg, cls='') {
    const p = document.createElement('div');
    p.className = `ccu-line ${cls}`;
    p.textContent = msg;
    el.log.appendChild(p);
    el.log.scrollTop = el.log.scrollHeight;
    // also console
    console[(cls === 'ccu-err') ? 'error' : (cls === 'ccu-warn' ? 'warn' : 'log')](msg);
  }

  function setBusy(b) {
    el.start.disabled = b;
    el.spinner.style.display = b ? 'inline-block' : 'none';
  }

  el.close.onclick = () => panel.remove();

  // =========
  // Context
  // =========
  const ORIGIN = location.origin; // e.g., https://alliant.instructure.com
  function getCourseIdFromPath() {
    // supports /courses/:id/... anywhere on a Canvas course page
    const parts = location.pathname.split('/').filter(Boolean);
    const i = parts.findIndex(x => x === 'courses');
    if (i >= 0 && parts[i+1]) return parts[i+1].split('?')[0];
    return null;
  }

  const COURSE_ID = getCourseIdFromPath();
  const CSRF = getCsrfToken();

  el.meta.textContent = COURSE_ID ? `Course ${COURSE_ID} @ ${ORIGIN}` : 'Course not detected';
  if (!COURSE_ID) {
    log('Could not detect course ID from URL. Please visit a page under /courses/:id and retry.', 'ccu-err');
  }

  // ==========================
  // Canvas API Helper Methods
  // ==========================
  const COMMON_HEADERS = {
    'Accept': 'application/json',
    'X-CSRF-Token': CSRF || ''
  };

  // generic Canvas JSON fetch
  async function canvas(path, opts={}) {
    const url = path.startsWith('http') ? path : `${ORIGIN}${path}`;
    const res = await fetch(url, {
      method: opts.method || 'GET',
      headers: {
        ...COMMON_HEADERS,
        ...(opts.headers || {}),
        ...(opts.body && !opts.headers?.['Content-Type'] ? {'Content-Type': 'application/json'} : {})
      },
      body: opts.body ? (typeof opts.body === 'string' ? opts.body : JSON.stringify(opts.body)) : null,
      credentials: 'same-origin'
    });
    if (!res.ok) {
      const text = await res.text().catch(()=> '');
      throw new Error(`HTTP ${res.status} ${res.statusText} - ${url}\n${text}`);
    }
    const ct = res.headers.get('Content-Type') || '';
    if (ct.includes('application/json')) {
      return { data: await res.json(), res };
    }
    return { data: await res.text(), res };
  }

  // pagination helper: follows Link headers
  async function canvasAll(path) {
    let out = [];
    let next = `${path}${path.includes('?') ? '&' : '?'}per_page=100`;
    while (next) {
      const { data, res } = await canvas(next, { method: 'GET' });
      out = out.concat(Array.isArray(data) ? data : []);
      const link = res.headers.get('Link') || res.headers.get('link') || '';
      const m = /<([^>]+)>;\s*rel="next"/.exec(link);
      next = m ? m[1] : null;
    }
    return out;
  }

  async function getPageByUrl(courseId, url) {
    // include[]=body to ensure the page body returns reliably
    const path = `/api/v1/courses/${courseId}/pages/${encodeURIComponent(url)}?include[]=body`;
    const { data } = await canvas(path);
    return data;
  }

  async function updatePageBody(courseId, url, newBody) {
    const path = `/api/v1/courses/${courseId}/pages/${encodeURIComponent(url)}`;
    const payload = { wiki_page: { body: newBody, notify_of_update: false } };
    const { data } = await canvas(path, { method: 'PUT', body: payload });
    return data;
  }

  async function findImagesFolderId(courseId) {
    const folders = await canvasAll(`/api/v1/courses/${courseId}/folders`);
    // pick the folder named exactly "images" or with full_name ending in /images
    const hit = folders.find(f => (f.name && f.name.toLowerCase() === 'images') || (f.full_name && /\/images$/i.test(f.full_name)));
    return hit ? hit.id : null;
  }

  async function ensureImagesFolderId(courseId) {
    const existing = await findImagesFolderId(courseId);
    if (existing) return existing;
    const { data } = await canvas(`/api/v1/courses/${courseId}/folders`, {
      method: 'POST',
      body: { name: 'images', parent_folder_path: '/' }
    });
    return data?.id || null;
  }

  async function ensureFileReadable(fileId) {
    const path = `/api/v1/files/${fileId}`;
    const { data } = await canvas(path);
    return data;
  }

  // Find files in a course that exactly match a given name
  async function listFilesByExactName(courseId, name) {
    const files = await canvasAll(`/api/v1/courses/${courseId}/files?search_term=${encodeURIComponent(name)}`);
    const lname = (name || '').toLowerCase();
    return files.filter(f => {
      const dn = (f.display_name || '').toLowerCase();
      const fn = (f.filename || '').toLowerCase();
      return dn === lname || fn === lname;
    });
  }

  // Delete any files in the (optional) target folder that match a given name
  async function deleteFilesByName(courseId, folderId, name) {
    const matches = await listFilesByExactName(courseId, name);
    const targets = matches.filter(f => !folderId || f.folder_id === folderId);
    for (const f of targets) {
      await canvas(`/api/v1/files/${f.id}`, { method: 'DELETE' });
    }
    return targets.length;
  }

  async function copyFileToFolder(fileId, folderId) {
    // Fetch source meta to get a stable name and use it for conflict handling
    const { data: srcMeta } = await canvas(`/api/v1/files/${fileId}`);
    const desiredName = srcMeta.display_name || srcMeta.filename || `file-${fileId}`;

    async function tryFolderCopy() {
      const path = `/api/v1/folders/${folderId}/copy_file`;
      const { data } = await canvas(path, { method: 'POST', body: { source_file_id: fileId, name: desiredName } });
      return data;
    }
    async function tryFileCopy() {
      const path2 = `/api/v1/files/${fileId}/copy`;
      const { data } = await canvas(path2, { method: 'POST', body: { parent_folder_id: folderId, name: desiredName } });
      return data;
    }

    // First attempt: folder copy; on failure fall back to file copy. If there's a 409 (name conflict),
    // delete existing duplicates in the destination and retry once to honor overwrite intent.
    try {
      return await tryFolderCopy();
    } catch (e) {
      try {
        return await tryFileCopy();
      } catch (e2) {
        if (String(e2.message || '').includes('409') && desiredName) {
          try {
            await deleteFilesByName(COURSE_ID, folderId, desiredName);
            await new Promise(r => setTimeout(r, 300));
            return await tryFileCopy();
          } catch (e3) {
            throw e3;
          }
        }
        throw e2;
      }
    }
  }

  // --- Resolve file context robustly using folder_id fallback ---
  async function getFileAndContext(fileId) {
    const { data: file } = await canvas(`/api/v1/files/${fileId}`);
    let ctxType = file.context_type;
    let ctxId = file.context_id;

    if (!ctxType || !ctxId) {
      if (!file.folder_id) throw new Error(`File ${fileId} has no folder_id`);
      const { data: folder } = await canvas(`/api/v1/folders/${file.folder_id}`);
      ctxType = folder.context_type;
      ctxId = folder.context_id;
    }
    return { file, ctxType, ctxId };
  }

  // --- Content Migration helpers ---
  async function startSelectiveFileMigration(destCourseId, sourceCourseId, fileIds) {
    // Create a course copy migration with selective import, then poll until completion
    const { data: mig } = await canvas(
      `/api/v1/courses/${destCourseId}/content_migrations`,
      {
        method: 'POST',
        body: {
          migration_type: 'course_copy_importer',
          settings: { source_course_id: sourceCourseId },
          selective_import: true
        }
      }
    );

    // select only the file(s) requested
    await canvas(
      `/api/v1/courses/${destCourseId}/content_migrations/${mig.id}/selective_import`,
      { method: 'POST', body: { 'copy[files][]': fileIds } }
    );

    // poll workflow_state until completed/failed
    let state = mig.workflow_state;
    while (state && state !== 'completed' && state !== 'failed') {
      await new Promise(r => setTimeout(r, 1200));
      const { data: m2 } = await canvas(`/api/v1/courses/${destCourseId}/content_migrations/${mig.id}`);
      state = m2.workflow_state;
    }
    if (state === 'failed') throw new Error('content migration failed');
    return true;
  }

  async function migrateLogoByFileId(brandFileId, destCourseId, destFolderId) {
    // Look up the source file and resolve context via folder if needed
    const { file: srcMeta, ctxType, ctxId } = await getFileAndContext(brandFileId);
    if (!srcMeta || !srcMeta.id) throw new Error('source file not found');

    const desiredName = srcMeta.display_name || srcMeta.filename || `file-${brandFileId}`;

    // Overwrite behavior: delete any existing file with the same name in the destination images folder (if provided)
    if (desiredName) {
      await deleteFilesByName(destCourseId, destFolderId, desiredName);
      await new Promise(r => setTimeout(r, 400));
    }

    // If the source lives in a course, prefer selective content migration
    if (ctxType === 'Course' && ctxId) {
      await startSelectiveFileMigration(destCourseId, ctxId, [brandFileId]);

      // Find the migrated file in the destination by name
      const files = await canvasAll(`/api/v1/courses/${destCourseId}/files?search_term=${encodeURIComponent(desiredName)}`);
      const newest = files.sort((a,b) => new Date(b.created_at) - new Date(a.created_at))[0];
      if (!newest) throw new Error(`migrated file "${desiredName}" not located in destination`);

      // Move into images folder if needed
      if (destFolderId && newest.folder_id !== destFolderId) {
        await canvas(`/api/v1/files/${newest.id}`, { method: 'PUT', body: { parent_folder_id: destFolderId } });
      }
      return newest;
    }

    // If not a Course context (Account, User, Group), fall back to direct copy
    return await copyFileToFolder(brandFileId, destFolderId);
  }

  // ------------
  // Rewriters mirroring the PHP patterns
  // ------------
  const R = {
    losrc: /<p>\s*At the end of this module,\s*you will be able to:\s*<\/p>\s*<ul>[\s\S]*?<\/ul>/i,
    srcBlock: /<div[^>]+id=["']all-course-content["'][^>]*>[\s\S]*?<\/div>/i,
    wktitle: /<h1[^>]+class=["']allt-mod-title["'][^>]*>[\s\S]*?<\/h1>/i,
    cdes: /<div[^>]+id=["']all-course-des["'][^>]*>[\s\S]*?<\/div>/i,
    cloDiv: /<div[^>]+id=["']all-course-objective["'][^>]*>[\s\S]*?<\/div>/i,
    cloH2UL: /<h2>\s*Course Outcomes\s*<\/h2>[\s\S]*?<\/ul>/i,
    firstUL: /<ul>[\s\S]*?<\/ul>/i,
    matsStrict: /<div class=["']all-txtbook["']>[\s\S]*?http:\/\/www\.<\/a><\/p>/i,
    matsFallbackGroup: /<div class=["']all-txtbook["']>[\s\S]*?<\/div>(?:\s*<div class=["']all-txtbook["']>[\s\S]*?<\/div>)*/i
  };

  // Turn <div class="LO"><p>..</p><p>..</p></div> into <ul><li>..</li><li>..</li></ul>
  function convertLODivToUL(html) {
    if (!html) return html;
    let v = String(html);
    v = v.replace(/<p>/gi, '<li>').replace(/<\/p>/gi, '</li>');
    // trim accidental wrapping <div> if present
    v = v.replace(/^<div[^>]*>/i, '').replace(/<\/div>\s*$/i, '');
    return `<ul>${v}</ul>`;
  }

  function replaceWeeklyTitle(body, title) {
    if (!R.wktitle.test(body)) return body;
    return body.replace(R.wktitle, `<h1 class="allt-mod-title">${title}</h1>`);
  }

  function replaceLOList(body, loUL) {
    // Prefer replacing the exact list (including the "At the end..." header) like PHP
    if (R.losrc.test(body)) {
      return body.replace(R.losrc, loUL);
    }
    // Fallback: replace the content inside #all-course-objective
    if (R.cloDiv.test(body)) {
      return body.replace(R.cloDiv, `<div id="all-course-objective">${loUL}</div>`);
    }
    return body;
  }

  function replaceActivitiesOrAssignments(body, contentHTML) {
    if (R.srcBlock.test(body)) {
      return body.replace(R.srcBlock, `<div id="all-course-content">${contentHTML}</div>`);
    }
    return body;
  }

  function replaceCourseDescription(body, cdesHTML) {
    if (R.cdes.test(body)) {
      return body.replace(R.cdes, cdesHTML);
    }
    return body;
  }

  function replaceCLOsOnStartHere(body, cloHTML) {
    // try to replace inside all-course-objective block
    if (R.cloDiv.test(body)) {
      return body.replace(R.cloDiv, `<div id="all-course-objective">${cloHTML}</div>`);
    }
    // else try the Course Outcomes section until closing </ul>, like the PHP did
    if (R.cloH2UL.test(body)) {
      return body.replace(R.cloH2UL, `<div id="all-course-objective">${cloHTML}</div>`);
    }
    return body;
  }

  function replacePLOList(body, ploHTML) {
    // Replace first UL with the entire <div class="plo"> contents as the PHP did
    if (R.firstUL.test(body)) {
      return body.replace(R.firstUL, ploHTML);
    }
    return body;
  }

  function replaceMaterials(body, materialHTML) {
    // Try strict PHP-like region
    if (R.matsStrict.test(body)) {
      return body.replace(R.matsStrict, materialHTML);
    }
    // Fallback: replace one or multiple textbook cards block(s)
    if (R.matsFallbackGroup.test(body)) {
      return body.replace(R.matsFallbackGroup, materialHTML);
    }
    // Last resort: if page has all-course-content, drop it there
    if (R.srcBlock.test(body)) {
      return body.replace(R.srcBlock, `<div id="all-course-content">${materialHTML}</div>`);
    }
    return body;
  }

  // ===========
  // Parser for the source page (all-pages)
  // ===========
  function parseSourceHTML(sourceHTML) {
    const dp = new DOMParser();
    const doc = dp.parseFromString(sourceHTML, 'text/html');

    // Titles: prefer <h1 class="wk-topic">Week N: Title</h1>
    const topics = Array.from(doc.querySelectorAll('#wk-topics-list li')).map(li => li.textContent.trim());

    const wkHeadings = Array.from(doc.querySelectorAll('h1.wk-topic')).map(h => h.textContent.trim());
    function headingToTopic(t) {
      if (!t) return '';
      const parts = t.split(':');
      return parts.length > 1 ? parts.slice(1).join(':').trim() : t.trim();
    }
    const derivedTopics = topics.length ? topics : wkHeadings.map(headingToTopic);

    const topicForWeek = (n) => derivedTopics[n-1] || '';

    // One-time blocks
    const cdes = doc.querySelector('div.cdes')?.innerHTML?.trim() || '';
    const plo  = doc.querySelector('div.plo')?.innerHTML?.trim() || '';
    const clo  = doc.querySelector('div.clo')?.innerHTML?.trim() || '';
    const material = doc.querySelector('div.material')?.innerHTML?.trim() || '';

    // Weeked blocks come in repeating trios: LO, Activities, Assignments
    const loDivs = Array.from(doc.querySelectorAll('div.LO'));
    const acts   = Array.from(doc.querySelectorAll('div.activities'));
    const asgns  = Array.from(doc.querySelectorAll('div.assignments'));
    const count = Math.max(loDivs.length, acts.length, asgns.length);

    const weeks = [];
    for (let i = 0; i < count; i++) {
      const num = i + 1;
      weeks.push({
        number: num,
        topic: topicForWeek(num),
        loUL: convertLODivToUL(loDivs[i]?.innerHTML || ''),
        activitiesHTML: acts[i]?.innerHTML?.trim() || '',
        assignmentsHTML: asgns[i]?.innerHTML?.trim() || ''
      });
    }

    return { cdes, plo, clo, material, weeks };
  }

  // ===========
  // Main Runner
  // ===========
  async function run() {
    if (!COURSE_ID) return;

    setBusy(true);
    log('Starting...');

    // 1) LOGO FIRST
    try {
      const logoId = el.logo.value;
      if (logoId) {
        log(`Preparing to migrate logo file #${logoId} into this course...`);
        log('Overwrite mode: if a file with the same name exists in /images, it will be deleted first.');
        const folderId = await ensureImagesFolderId(COURSE_ID);
        if (!folderId) {
          log('Images folder not found or could not be created; skipping logo migration.', 'ccu-warn');
        } else {
          try {
            const f = await migrateLogoByFileId(logoId, COURSE_ID, folderId);
            log(`Logo migrated as file #${f.id}: ${f.display_name || f.filename || '(unnamed)'}`, 'ccu-ok');
          } catch (e1) {
            log(`Migration failed (${e1.message}). Trying direct copy as fallback...`, 'ccu-warn');
            try {
              const f2 = await copyFileToFolder(logoId, folderId);
              log(`Logo copied as file #${f2.id}: ${f2.display_name || f2.filename || '(unnamed)'}`, 'ccu-ok');
            } catch (e2) {
              log(`Logo copy failed after migration fallback: ${e2.message}`, 'ccu-err');
            }
          }
        }
      } else {
        log('No logo selected; skipping logo copy.');
      }
    } catch (e) {
      log(`Logo migration/copy error: ${e.message}`, 'ccu-err');
    }

    // 2) Fetch source content from all-pages
    log('Fetching source content from page: all-pages ...');
    let src;
    try {
      const page = await getPageByUrl(COURSE_ID, 'all-pages');
      src = page?.body || '';
      if (!src) throw new Error('Could not read body from all-pages.');
      log(`Source content loaded (${src.length.toLocaleString()} chars).`, 'ccu-ok');
    } catch (e) {
      log(`Failed to fetch source content: ${e.message}`, 'ccu-err');
      setBusy(false);
      showToast('Failed', false);
      return;
    }

    const parsed = parseSourceHTML(src);

    // 3) Start Here: description + CLOs
    try {
      log('Updating start-here (Course Description + CLOs)...');
      let startPage = await getPageByUrl(COURSE_ID, 'start-here');
      if (!startPage?.body) throw new Error('start-here not found or has no body');
      let body = startPage.body;

      if (parsed.cdes) {
        const nb = replaceCourseDescription(body, parsed.cdes);
        if (nb !== body) {
          body = nb;
          log('- Course Description updated.', 'ccu-ok');
        } else {
          log('- Course Description pattern not found; no change.', 'ccu-warn');
        }
      } else {
        log('- No <div class="cdes"> found in source; skipping Course Description.', 'ccu-warn');
      }

      if (parsed.clo) {
        const nb = replaceCLOsOnStartHere(body, convertLODivToUL(parsed.clo));
        if (nb !== body) {
          body = nb;
          log('- Course Learning Outcomes updated.', 'ccu-ok');
        } else {
          log('- CLOs pattern not found; no change.', 'ccu-warn');
        }
      } else {
        log('- No <div class="clo"> found in source; skipping CLOs.', 'ccu-warn');
      }

      if (body !== startPage.body) {
        await updatePageBody(COURSE_ID, 'start-here', body);
        log('start-here saved.', 'ccu-ok');
      } else {
        log('start-here: nothing to update.');
      }
    } catch (e) {
      log(`start-here update failed: ${e.message}`, 'ccu-err');
    }

    // 4) Program Learning Outcomes page
    try {
      if (parsed.plo) {
        log('Updating program-learning-outcomes ...');
        const page = await getPageByUrl(COURSE_ID, 'program-learning-outcomes');
        if (!page?.body) throw new Error('program-learning-outcomes not found or has no body');
        const nb = replacePLOList(page.body, parsed.plo);
        if (nb !== page.body) {
          await updatePageBody(COURSE_ID, 'program-learning-outcomes', nb);
          log('program-learning-outcomes saved.', 'ccu-ok');
        } else {
          log('program-learning-outcomes: target UL not found; no change.', 'ccu-warn');
        }
      } else {
        log('No <div class="plo"> found in source; skipping program-learning-outcomes.', 'ccu-warn');
      }
    } catch (e) {
      log(`PLO update failed: ${e.message}`, 'ccu-err');
    }

    // 5) Required Course Materials page
    try {
      if (parsed.material) {
        log('Updating required-course-materials ...');
        const page = await getPageByUrl(COURSE_ID, 'required-course-materials');
        if (!page?.body) throw new Error('required-course-materials not found or has no body');
        const nbTry1 = replaceMaterials(page.body, parsed.material);
        if (nbTry1 !== page.body) {
          await updatePageBody(COURSE_ID, 'required-course-materials', nbTry1);
          log('required-course-materials saved.', 'ccu-ok');
        } else {
          log('required-course-materials: pattern not found; no change.', 'ccu-warn');
        }
      } else {
        log('No <div class="material"> found in source; skipping required-course-materials.', 'ccu-warn');
      }
    } catch (e) {
      log(`Materials update failed: ${e.message}`, 'ccu-err');
    }

    // 6) Module titles (from Weekly Topics)
    try {
      const topics = parsed.weeks.map(w => w.topic || '').filter(Boolean);
      if (topics.length) {
        log('Updating module titles from Weekly Topics...');
        const modules = await canvasAll('/api/v1/courses/' + COURSE_ID + '/modules');
        const NUM_WORDS = ['one','two','three','four','five','six','seven','eight','nine','ten','eleven','twelve','thirteen','fourteen','fifteen'];
        function getWeekNumberFromName(name) {
          name = (name || '').toLowerCase();
          const pos = name.indexOf('week ');
          if (pos === -1) return null;
          const after = name.slice(pos + 5).trim();
          // parse leading digits
          let digits = '';
          for (let i=0;i<after.length;i++) {
            const c = after[i];
            if (c >= '0' && c <= '9') digits += c; else break;
          }
          if (digits) return parseInt(digits, 10);
          // parse word form
          for (let i=0;i<NUM_WORDS.length;i++) {
            if (after.startsWith(NUM_WORDS[i])) return i+1;
          }
          return null;
        }
        function makeModuleName(oldName, weekNum, topic) {
          const colonIdx = oldName.indexOf(':');
          if (colonIdx !== -1) {
            return oldName.slice(0, colonIdx + 1) + ' ' + topic;
          }
          const lower = oldName.toLowerCase();
          const wkIdx = lower.indexOf('week ');
          if (wkIdx !== -1) {
            const before = oldName.slice(0, wkIdx);
            return before + 'Week ' + weekNum + ': ' + topic;
          }
          if (oldName.includes('??')) return oldName.replace('??', topic);
          return oldName + ' - ' + topic;
        }
        for (const mod of modules) {
          const weekNum = getWeekNumberFromName(mod.name || '');
          if (!weekNum || !topics[weekNum - 1]) continue;
          const topic = topics[weekNum - 1];
          const newName = makeModuleName(mod.name || '', weekNum, topic);
          if (newName && newName !== mod.name) {
            await canvas('/api/v1/courses/' + COURSE_ID + '/modules/' + mod.id, { method: 'PUT', body: { module: { name: newName } } });
            log('Module renamed: "' + (mod.name || '') + '" -> "' + newName + '"', 'ccu-ok');
          }
        }
      } else {
        log('No weekly topics found to update module titles.', 'ccu-warn');
      }
    } catch (e) {
      log('Module titles update failed: ' + e.message, 'ccu-err');
    }

    // 7) Weekly pages (Learning Objectives, Activities & Resources, Assignments)
    for (const wk of parsed.weeks) {
      const b = wk.number; // week number
      log(`Week ${b}: updating pages...`);

      // 7a) Learning Objectives page
      try {
        const slug = `week-${b}-learning-objectives`;
        const p = await getPageByUrl(COURSE_ID, slug);
        if (!p?.body) {
          log(`- ${slug} not found; skipping.`, 'ccu-warn');
        } else {
          let body = p.body;
          if (wk.topic) {
            const nb = replaceWeeklyTitle(body, wk.topic);
            if (nb !== body) {
              body = nb;
              log('- Title set: ' + wk.topic, 'ccu-ok');
            }
          }
          if (wk.loUL) {
            const nb2 = replaceLOList(body, wk.loUL);
            if (nb2 !== body) {
              body = nb2;
              log('- Learning Objectives replaced.', 'ccu-ok');
            } else {
              log('- LO pattern not found on page; no change.', 'ccu-warn');
            }
          } else {
            log('- No LO HTML in source for this week; skipping.', 'ccu-warn');
          }
          if (body !== p.body) {
            await updatePageBody(COURSE_ID, slug, body);
            log(`- ${slug} saved.`, 'ccu-ok');
          }
        }
      } catch (e) {
        log(`- LO update error: ${e.message}`, 'ccu-err');
      }

      // 7b) Activities & Resources page
      try {
        const slug = `week-${b}-activities-and-resources`;
        const p = await getPageByUrl(COURSE_ID, slug);
        if (!p?.body) {
          log(`- ${slug} not found; skipping.`, 'ccu-warn');
        } else if (wk.activitiesHTML) {
          const nb = replaceActivitiesOrAssignments(p.body, wk.activitiesHTML);
          if (nb !== p.body) {
            await updatePageBody(COURSE_ID, slug, nb);
            log(`- ${slug} saved.`, 'ccu-ok');
          } else {
            log(`- ${slug}: target block not found; no change.`, 'ccu-warn');
          }
        } else {
          log(`- No activities HTML in source; skipping ${slug}.`, 'ccu-warn');
        }
      } catch (e) {
        log(`- Activities update error: ${e.message}`, 'ccu-err');
      }

      // 7c) Assignments page
      try {
        const slug = `week-${b}-assignments`;
        const p = await getPageByUrl(COURSE_ID, slug);
        if (!p?.body) {
          log(`- ${slug} not found; skipping.`, 'ccu-warn');
        } else if (wk.assignmentsHTML) {
          const nb = replaceActivitiesOrAssignments(p.body, wk.assignmentsHTML);
          if (nb !== p.body) {
            await updatePageBody(COURSE_ID, slug, nb);
            log(`- ${slug} saved.`, 'ccu-ok');
          } else {
            log(`- ${slug}: target block not found; no change.`, 'ccu-warn');
          }
        } else {
          log(`- No assignments HTML in source; skipping ${slug}.`, 'ccu-warn');
        }
      } catch (e) {
        log(`- Assignments update error: ${e.message}`, 'ccu-err');
      }
    }

    setBusy(false);
    log('All operations completed.');
    showToast('All done!');
  }

  el.start.onclick = () => {
    run().catch(e => {
      setBusy(false);
      log(`Unexpected error: ${e.message}`, 'ccu-err');
      showToast('Failed', false);
    });
  };
})();
