// === Canvas DOCX → RCE Toolbelt =============================================
// PASTE THIS IN THE BROWSER CONSOLE WHILE EDITING A CANVAS PAGE
// It adds a floating UI above the RCE to upload/convert/process DOCX content,
// run your post-processing tools on current content, preview styles,
// and save back via API. Now supports drag & drop of .docx files.
//
// Version: 1.4.1 (removed Convert/Process buttons; auto-convert on file select; fixed stray tail)
//
// -----------------------------------------------------------------------------
// 0) Helpers: CSRF + URL context
function getCsrfToken() {
  const csrfRegex = new RegExp('^_csrf_token=(.*)$');
  let csrf;
  const cookies = document.cookie.split(';');
  for (let i = 0; i < cookies.length; i++) {
    const cookie = cookies[i].trim();
    const match = csrfRegex.exec(cookie);
    if (match) { csrf = decodeURIComponent(match[1]); break; }
  }
  return csrf;
}

const CanvasCtx = (() => {
  const mCourse = location.pathname.match(/\/courses\/(\d+)/);
  const mPage   = location.pathname.match(/\/pages\/([^/]+)/);
  return {
    host: location.host,
    protocol: location.protocol,
    courseId: mCourse ? mCourse[1] : null,
    pageSlug: mPage ? mPage[1].replace(/\/edit$/, '') : null,
    apiBase: `${location.origin}/api/v1`
  };
})();

if (window.__CanvasDocxToolsLoaded) {
  console.log('Canvas DOCX Toolbelt already loaded.');
} else {
  window.__CanvasDocxToolsLoaded = true;

  // ---------------------------------------------------------------------------
  // 1) Inject Mammoth if missing
  function ensureMammoth() {
    return new Promise((resolve, reject) => {
      if (window.mammoth) return resolve();
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js';
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('Failed to load Mammoth.js'));
      document.head.appendChild(s);
    });
  }

  // ---------------------------------------------------------------------------
  // 2) UI Shell (top toolbar + modal + spinner + toast)
  const root = document.createElement('div');
  root.id = 'canvas-docx-tools';
  const css = `
#canvas-docx-tools { position: sticky; top: 0; z-index: 9999; margin-bottom: .5rem; }
#canvas-docx-tools .cdt-wrap { background: #f7fafc; border: 1px solid #cbd5e0; border-radius: 8px; padding: 10px; display:flex; flex-wrap:wrap; gap:8px; align-items:center; }
#canvas-docx-tools label { font-size: 12px; color:#2d3748; margin-right:4px; }
#canvas-docx-tools input[type="file"] { font-size: 12px; }
#canvas-docx-tools button.cdt-btn { background:#225f92; color:#fff; border:0; border-radius:8px; padding:8px 10px; cursor:pointer; font-size:12px; }
#canvas-docx-tools button.cdt-btn.secondary { background:#6c757d; }
#canvas-docx-tools button.cdt-btn.success { background:#28a745; }
#canvas-docx-tools button.cdt-btn.warn { background:#ffc107; color:#212529; }
#canvas-docx-tools button.cdt-btn.danger { background:#dc3545; }
#canvas-docx-tools .cdt-spacer { flex: 1 1 auto; }
#canvas-docx-tools .cdt-note { font-size:12px; color:#4a5568; }

#cdt-spinner { position:fixed; inset:0; background:rgba(0,0,0,.35); display:none; align-items:center; justify-content:center; z-index:10000; }
#cdt-spinner .box { background:#fff; padding:18px 22px; border-radius:10px; box-shadow:0 10px 30px rgba(0,0,0,0.25); font: 14px/1.4 system-ui, -apple-system, Segoe UI, Roboto; }
#cdt-toast { position:fixed; right:18px; bottom:18px; background:#225f92; color:#fff; padding:10px 14px; border-radius:8px; display:none; z-index:10001; font-size:13px; }

#cdt-modal { position:fixed; inset:0; background:rgba(0,0,0,.45); display:none; align-items:center; justify-content:center; z-index:10000;}
#cdt-modal .inner { background:#fff; max-width:640px; width:92vw; border-radius:10px; padding:18px; }
#cdt-modal h3 { margin:0 0 8px; font:600 18px/1.2 system-ui, -apple-system, Segoe UI, Roboto; }
#cdt-modal .body { max-height:60vh; overflow:auto; border:1px solid #e2e8f0; border-radius:8px; padding:10px; background:#f8fafc; }
#cdt-modal .row { margin-top:10px; display:flex; gap:8px; justify-content:flex-end; }
#cdt-modal .btn { background:#225f92; color:#fff; border:0; border-radius:8px; padding:8px 10px; cursor:pointer; font-size:12px; }
#cdt-modal .btn.gray { background:#6c757d; }
#cdt-modal table { border-collapse:collapse; width:100%; font-size:12px; }
#cdt-modal th, #cdt-modal td { border:1px solid #cfd4da; padding:6px; text-align:left; }
`;
  const style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);

  root.innerHTML = `
    <div class="cdt-wrap" title="You can drop a .docx file on this bar">
      <label>DOCX:</label>
      <input id="cdt-file" type="file" accept=".docx" />
      <span class="cdt-note cdt-drop-tip">choose or drop a .docx (auto-converts)</span>
      <button class="cdt-btn" type="button" id="cdt-report">Section Report</button>
      <button class="cdt-btn warn" type="button" id="cdt-cleanup">Clean Up Formatting</button>
      <button class="cdt-btn" type="button" id="cdt-undo" style="display:none">Undo Cleanup</button>
      <span class="cdt-spacer"></span>
      <button class="cdt-btn secondary" type="button" id="cdt-toggle-styles">Preview Styles: On</button>
      <button class="cdt-btn secondary" type="button" id="cdt-style-legend">Styles Legend</button>
      <button class="cdt-btn success" type="button" id="cdt-save">Save to Page</button>
      <span class="cdt-note">Page: ${CanvasCtx.courseId ?? 'unknown course'} / ${CanvasCtx.pageSlug ?? 'unknown page'}</span>
    </div>
  `;
  const mountTarget = document.querySelector('.edit-content') || document.body;
  mountTarget.prepend(root);

  const spinner = document.createElement('div');
  spinner.id = 'cdt-spinner';
  spinner.innerHTML = `<div class="box">Working…</div>`;
  document.body.appendChild(spinner);

  const toast = document.createElement('div');
  toast.id = 'cdt-toast';
  document.body.appendChild(toast);
  function showToast(msg, ms=1700){ toast.textContent = msg; toast.style.display='block'; setTimeout(()=>toast.style.display='none', ms); }

  const modal = document.createElement('div');
  modal.id = 'cdt-modal';
  modal.innerHTML = `
    <div class="inner">
      <h3 id="cdt-modal-title"></h3>
      <div class="body" id="cdt-modal-body"></div>
      <div class="row">
        <button class="btn gray" type="button" id="cdt-modal-close">Close</button>
      </div>
    </div>`;
  document.body.appendChild(modal);
  document.getElementById('cdt-modal-close').onclick = () => modal.style.display = 'none';

  function withSpinner(fn){ return async (...args) => { spinner.style.display='flex'; try { return await fn(...args); } finally { spinner.style.display='none'; } }; }

  // Extra style for drag highlight
  const cdtExtraStyle = document.createElement('style');
  cdtExtraStyle.textContent = '#canvas-docx-tools .cdt-drop-tip{opacity:.75;font-style:italic} #canvas-docx-tools.drop-active .cdt-wrap{outline:2px dashed #225f92;background:#edf2f7}';
  document.head.appendChild(cdtExtraStyle);

  // ---------------------------------------------------------------------------
  // 3) RCE bridge
  function getEditor() { return (window.tinymce && tinymce.get('wiki_page_body')) || null; }
  function getRCEHtml() {
    const ed = getEditor();
    if (ed) return ed.getContent({format:'html'});
    const ta = document.getElementById('wiki_page_body');
    return ta ? ta.value : '';
  }
  function setRCEHtml(html) {
    const ed = getEditor();
    if (ed) ed.setContent(html);
    const ta = document.getElementById('wiki_page_body');
    if (ta) ta.value = html;
    // Re-apply preview styles inside iframe after content changes
    setPreviewStyles(STATE.previewOn);
  }

  // ---------------------------------------------------------------------------
  // 4) Processing pipeline (ported from index.html)
  const APP_CONFIG = {
    styleMap: [
      "p[style-name='LO Heading'] => h1:fresh",
      "p[style-name='Weekly Objective'] => h2:fresh",
      "p[style-name='Assignment Title'] => h3:fresh",
      "p[style-name='Weekly Topic Heading'] => h1.wk-topic:fresh",
      "p[style-name='Course Description'] => h1.cdes:fresh",
      "p[style-name='Program Learning Outcomes'] => h1.plo:fresh",
      "p[style-name='Course Learning Outcomes'] => h1.clo:fresh",
      "p[style-name='Required Course Materials'] => h1.material:fresh"
    ],
    topLevelSections: [
      { name: 'Course Description', className: 'cdes' },
      { name: 'Program Learning Outcomes', className: 'plo' },
      { name: 'Course Learning Outcomes', className: 'clo' },
      { name: 'Required Course Materials', className: 'material' }
    ],
    weeklySections: [
      { name: 'Learning Objectives', className: 'LO' },
      { name: 'Activities and Resources', className: 'activities' },
      { name: 'Assignments', className: 'assignments' },
    ]
  };
  APP_CONFIG.allowedWrapperSelectors = [
    '#wk-topics-list',
    ...APP_CONFIG.topLevelSections.map(s => `.${s.className}`),
    ...APP_CONFIG.weeklySections.map(s => `.${s.className}`)
  ].join(', ');

  const textOnly = (el) => (el.textContent || '').trim().replace(/\s+/g, ' ');

  function stashNestedTables(doc) {
    const store = new Map(); let counter = 0;
    doc.querySelectorAll('table table').forEach(tbl => {
      const id = `__TABLE_PH__${++counter}__`;
      store.set(id, tbl.cloneNode(true));
      const ph = doc.createElement('span');
      ph.setAttribute('data-table-ph', id);
      tbl.replaceWith(ph);
    });
    return store;
  }
  function restorePlaceholders(doc, store) {
    if (!store) return;
    doc.querySelectorAll('span[data-table-ph]').forEach(ph => {
      const id = ph.getAttribute('data-table-ph');
      const original = store.get(id);
      if (original) ph.replaceWith(original);
    });
  }
  function moveChildrenPreservingBlocks(src, dst) {
    let currentP = null;
    const flushP = () => { if (currentP && currentP.childNodes.length) dst.appendChild(currentP); currentP = null; };
    const ensureP = () => (currentP ||= document.createElement('p'));
    function walk(node) {
      while (node.firstChild) {
        const child = node.firstChild; node.removeChild(child);
        if (child.nodeType === Node.TEXT_NODE) { if (child.textContent.trim()) ensureP().appendChild(document.createTextNode(child.textContent)); continue; }
        if (child.nodeType !== Node.ELEMENT_NODE) continue;
        if (child.matches('span[data-table-ph]')) { flushP(); dst.appendChild(child); continue; }
        if (child.tagName === 'UL' || child.tagName === 'OL') { flushP(); dst.appendChild(child); continue; }
        if (child.tagName === 'TABLE') { flushP(); dst.appendChild(child); continue; }
        if (['P','DIV','SECTION','ARTICLE'].includes(child.tagName)) { walk(child); flushP(); continue; }
        ensureP().appendChild(child);
      }
    }
    walk(src); flushP();
  }
  function processContentTables(doc, sectionHeadingText, wrapperClass, endWithHr = false) {
    const sectionHeadings = Array.from(doc.querySelectorAll('h1')).filter(h =>
      h.textContent.trim().toLowerCase().includes(sectionHeadingText.toLowerCase())
    );
    sectionHeadings.forEach(heading => {
      const mainWrapper = document.createElement('div'); mainWrapper.className = wrapperClass;
      const tablesToProcess = [];
      let currentNode = heading.nextElementSibling;
      while (currentNode && currentNode.tagName !== 'H1') { if (currentNode.tagName === 'TABLE') tablesToProcess.push(currentNode); currentNode = currentNode.nextElementSibling; }
      tablesToProcess.forEach((table, index) => {
        const itemDiv = document.createElement('div');
        const rows = Array.from(table.querySelectorAll('tr'));
        if (rows.length > 0) {
          const titleCell = rows[0].querySelector('td, th');
          if (titleCell && titleCell.textContent.trim()) { const h2 = document.createElement('h2'); h2.textContent = textOnly(titleCell); itemDiv.appendChild(h2); }
          for (let i = 1; i < rows.length; i++) {
            const contentCell = rows[i].querySelector('td, th');
            if (contentCell) moveChildrenPreservingBlocks(contentCell, itemDiv);
          }
        }
        if (itemDiv.hasChildNodes()) mainWrapper.appendChild(itemDiv);
        const isLast = index === tablesToProcess.length - 1;
        if (endWithHr || !isLast) mainWrapper.appendChild(document.createElement('hr'));
      });
      if (mainWrapper.hasChildNodes()) { heading.insertAdjacentElement('afterend', mainWrapper); tablesToProcess.forEach(t => t.remove()); }
    });
  }
  function wrapSectionContent(doc, headingText, wrapperClass) {
    const heading = Array.from(doc.querySelectorAll('h1')).find(h => h.textContent.trim().toLowerCase().includes(headingText.toLowerCase()));
    if (!heading) return;
    const wrapper = document.createElement('div'); wrapper.className = wrapperClass;
    let currentNode = heading.nextElementSibling; const move = [];
    while (currentNode && currentNode.tagName !== 'H1') { move.push(currentNode); currentNode = currentNode.nextElementSibling; }
    move.forEach(el => wrapper.appendChild(el));
    if (wrapper.hasChildNodes()) heading.insertAdjacentElement('afterend', wrapper);
  }
  function cleanupHtml(doc) {
    doc.querySelectorAll('p, a').forEach(el => { if (!el.textContent.trim() && !el.children.length) el.remove(); });
    return doc;
  }
  function createWeeklyTopicsList(doc) {
    const topicHeaders = doc.querySelectorAll('h1.wk-topic');
    if (topicHeaders.length === 0) return null;
    const listContainer = document.createElement('div'); listContainer.id = 'wk-topics-list';
    const title = document.createElement('h2'); title.textContent = 'Weekly Topics'; listContainer.appendChild(title);
    const ul = document.createElement('ul');
    topicHeaders.forEach(h1 => {
      const li = document.createElement('li');
      li.textContent = h1.textContent.trim().replace(/^Week\s+\d+:\s*/i, '');
      ul.appendChild(li);
    });
    listContainer.appendChild(ul); return listContainer;
  }
  function processHtml(rawHtml) {
    const parser = new DOMParser(); let doc = parser.parseFromString(rawHtml, 'text/html');
    const tableStore = stashNestedTables(doc);
    APP_CONFIG.topLevelSections.forEach(section => wrapSectionContent(doc, section.name, section.className));
    doc.querySelectorAll('h1.wk-topic').forEach(topic => {
      const nextH1 = topic.nextElementSibling;
      if (nextH1 && nextH1.tagName === 'H1' && nextH1.textContent.trim().toLowerCase().includes('learning objectives')) {
        const table = nextH1.nextElementSibling;
        if (table && table.tagName === 'TABLE') {
          const wrapper = document.createElement('div'); wrapper.className = 'LO';
          const seen = new Set();
          Array.from(table.querySelectorAll('li')).forEach(li => {
            const t = li.textContent.trim(); if (t && !seen.has(t)) { const p = document.createElement('p'); p.textContent = t; wrapper.appendChild(p); seen.add(t); }
          });
          if (wrapper.hasChildNodes()) { nextH1.insertAdjacentElement('afterend', wrapper); table.remove(); }
        }
      }
    });
    processContentTables(doc, 'Activities and Resources', 'activities', false);
    processContentTables(doc, 'Assignments', 'assignments', true);
    restorePlaceholders(doc, tableStore);
    doc = cleanupHtml(doc);
    const weeklyList = createWeeklyTopicsList(doc);
    if (weeklyList) doc.body.prepend(weeklyList);

    // Ensure table borders like your index.html tweak
    doc.querySelectorAll('table').forEach(tbl => tbl.setAttribute('border','1'));
    return doc.body.innerHTML;
  }

  function generateSectionPresenceReport(htmlContent) {
    const parser = new DOMParser(); const doc = parser.parseFromString(htmlContent, 'text/html');
    let issuesFound = false; const recommendations = []; const topLevelErrors = [];
    APP_CONFIG.topLevelSections.forEach(section => {
      if (doc.querySelectorAll(`.${section.className}`).length === 0) {
        issuesFound = true; const msg = `Top-Level section "${section.name}" is missing.`; topLevelErrors.push(msg); recommendations.push(msg);
      }
    });
    const wkTopicHeaders = doc.querySelectorAll('h1.wk-topic');
    const weeklyReportData = []; let anyUnwrapped = false;
    if (wkTopicHeaders.length === 0) { issuesFound = true; recommendations.push("No 'Weekly Topic' sections (h1.wk-topic) were found."); }
    wkTopicHeaders.forEach((wkTopic, index) => {
      const weekNumber = index + 1;
      const cleanText = wkTopic.textContent.trim().replace(/^Week\s+\d+:\s*/i, '');
      let els = []; let cur = wkTopic.nextElementSibling;
      while (cur && !cur.matches('h1.wk-topic')) { els.push(cur); cur = cur.nextElementSibling; }
      const frag = document.createElement('div'); els.forEach(el => frag.appendChild(el.cloneNode(true)));
      const counts = {}; APP_CONFIG.weeklySections.forEach(s => counts[s.className] = frag.querySelectorAll(`.${s.className}`).length);
      let unwrapped = false;
      Array.from(frag.children).forEach(child => {
        if (child.nodeType === Node.ELEMENT_NODE && !child.matches(APP_CONFIG.allowedWrapperSelectors) && !child.matches('h1, hr')) {
          if (child.textContent.trim() !== '') unwrapped = true;
        }
      });
      if (unwrapped) anyUnwrapped = true;
      weeklyReportData.push({weekNumber, cleanText, counts, unwrapped});
      APP_CONFIG.weeklySections.forEach(s => {
        const c = counts[s.className];
        if (c === 0) { issuesFound = true; recommendations.push(`Week ${weekNumber} is missing a ${s.name} section.`); }
        if (c > 1)  { issuesFound = true; recommendations.push(`Week ${weekNumber} has too many ${s.name} sections (found ${c}, expected 1).`); }
      });
      if (unwrapped) { issuesFound = true; recommendations.push(`Week ${weekNumber} has unwrapped content.`); }
    });

    let html = '<h4>Document Structure Report</h4>';
    if (issuesFound) {
      html += `<p style="color:#dc3545;font-weight:700">Issues detected. Please review recommendations.</p>`;
      topLevelErrors.forEach(e => html += `<p style=\"color:#dc3545;padding-left:12px\">${e}</p>`);
    } else {
      html += `<p style="color:#28a745;font-weight:700">Document structure looks good!</p>`;
    }

    html += '<h5>Top-Level Sections</h5><table><thead><tr><th>Section</th><th>Present?</th><th>Count</th></tr></thead><tbody>';
    APP_CONFIG.topLevelSections.forEach(s => {
      const c = doc.querySelectorAll(`.${s.className}`).length;
      html += `<tr><td>${s.name}</td><td>${c>0?'Yes':'No'}</td><td>${c}</td></tr>`;
    });
    html += '</tbody></table>';

    html += '<h5 style="margin-top:10px">Weekly Sections</h5>';
    if (anyUnwrapped) html += `<p style="color:#dc3545">Unwrapped content found. This will be deleted by "Clean Up".</p>`;
    html += '<table><thead><tr><th>Week</th><th>LO</th><th>A&R</th><th>Assignments</th><th>Unwrapped</th></tr></thead><tbody>';
    weeklyReportData.forEach(d => {
      html += `<tr><td>Week ${d.weekNumber}: ${d.cleanText}</td><td>${d.counts.LO}</td><td>${d.counts.activities}</td><td>${d.counts.assignments}</td><td>${d.unwrapped?'1':'0'}</td></tr>`;
    });
    html += '</tbody></table>';
    if (issuesFound) {
      html += '<h5 style="margin-top:10px">Recommendations</h5><ul>';
      recommendations.forEach(r => html += `<li>${r}</li>`);
      html += '</ul>';
    }
    return html;
  }

  function removeExtraContent(htmlContent) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlContent, 'text/html');
    const body = doc.body;
    const nodesToRemove = new Set();
    let lastH1 = null;
    Array.from(body.children).forEach(node => {
      if (node.tagName === 'H1') { lastH1 = node; return; }
      if (!node.matches(APP_CONFIG.allowedWrapperSelectors)) {
        if (node.nodeType === Node.TEXT_NODE && !node.textContent.trim()) return;
        nodesToRemove.add(node); if (lastH1) nodesToRemove.add(lastH1);
      }
    });
    nodesToRemove.forEach(node => node.remove());
    return body.innerHTML;
  }

  // ---------------------------------------------------------------------------
  // 4.5) Section preview styles & legend
  const SECTION_STYLE_CSS = `/* Course section styles */
.cdes, .plo, .clo, .material, .LO, .activities, .assignments { padding: 12px 14px; border-left: 6px solid; background: #f9fbfd; margin: 12px 0; }
.cdes { border-color:#1d4ed8; background:#eff6ff; }
.plo { border-color:#0ea5e9; background:#ecfeff; }
.clo { border-color:#10b981; background:#ecfdf5; }
.material { border-color:#f59e0b; background:#fff7ed; }
.LO { border-color:#7c3aed; background:#f5f3ff; }
.activities { border-color:#059669; background:#ecfdf5; }
.assignments { border-color:#ef4444; background:#fef2f2; }
#wk-topics-list { border:1px dashed #94a3b8; padding:10px 12px; background:#f8fafc; }
#wk-topics-list h2 { margin: 0 0 6px; }
#wk-topics-list ul { margin:0; padding-left: 18px; }
#wk-topics-list li { margin:2px 0; }
h1.wk-topic { color:#1f2937; border-bottom:2px solid #e5e7eb; padding-bottom:4px; }
.cdes > h2, .plo > h2, .clo > h2, .material > h2, .LO > h2, .activities > h2, .assignments > h2 { margin-top: 0; }
`;
  const PREVIEW_STYLE_PARENT_ID = 'cdt-preview-styles';
  const PREVIEW_STYLE_IFRAME_ID = 'cdt-preview-styles-editor';
  const STATE = { previewOn: true, iframeDnDAttached: false };

  function setPreviewStyles(on) {
    // Parent page
    let s = document.getElementById(PREVIEW_STYLE_PARENT_ID);
    if (on && !s) { s = document.createElement('style'); s.id = PREVIEW_STYLE_PARENT_ID; s.textContent = SECTION_STYLE_CSS; document.head.appendChild(s); }
    else if (!on && s) { s.remove(); }

    // TinyMCE iframe
    const ed = getEditor();
    if (ed && ed.getDoc) {
      const edDoc = ed.getDoc();
      if (edDoc) {
        let s2 = edDoc.getElementById(PREVIEW_STYLE_IFRAME_ID);
        if (on && !s2) { s2 = edDoc.createElement('style'); s2.id = PREVIEW_STYLE_IFRAME_ID; s2.textContent = SECTION_STYLE_CSS; edDoc.head.appendChild(s2); }
        else if (!on && s2) { s2.remove(); }
      }
    }
  }


  function openStyleLegend() {
    const rows = [
      ['cdes','Course Description','#1d4ed8'],
      ['plo','Program Learning Outcomes','#0ea5e9'],
      ['clo','Course Learning Outcomes','#10b981'],
      ['material','Required Course Materials','#f59e0b'],
      ['LO','Learning Objectives','#7c3aed'],
      ['activities','Activities and Resources','#059669'],
      ['assignments','Assignments','#ef4444']
    ].map(([cls,label,color]) => `<tr><td><code>.${cls}</code></td><td>${label}</td><td><span style="display:inline-block;width:16px;height:16px;background:${color};border-radius:3px;border:1px solid #0002"></span></td></tr>`).join('');
    document.getElementById('cdt-modal-title').textContent = 'Styles Legend';
    document.getElementById('cdt-modal-body').innerHTML = `
      <p>The following wrappers are applied by the processor. The <em>Preview Styles</em> button shows them visually in the editor.</p>
      <table><thead><tr><th>Class</th><th>Meaning</th><th>Color</th></tr></thead><tbody>${rows}</tbody></table>
      <p><strong>Note:</strong> Canvas pages strip <code>&lt;style&gt;</code> tags on save; previews are editor-only unless your theme defines these classes.</p>`;
    modal.style.display = 'flex';
  }

  // ---------------------------------------------------------------------------
  // 5) Actions (incl. drag & drop)
  const fileInput          = document.getElementById('cdt-file');
  const btnReport          = document.getElementById('cdt-report');
  const btnCleanup         = document.getElementById('cdt-cleanup');
  const btnUndo            = document.getElementById('cdt-undo');
  const btnToggleStyles    = document.getElementById('cdt-toggle-styles');
  const btnStyleLegend     = document.getElementById('cdt-style-legend');
  const btnSave            = document.getElementById('cdt-save');

  let lastPreCleanupHtml = null;

  async function convertAndReplace(fileOpt) {
    if (!fileOpt && (!fileInput.files || !fileInput.files[0])) { alert('Choose a .docx file first, or drop one on the bar.'); return; }
    await ensureMammoth();
    const file = fileOpt || fileInput.files[0];
    const isDocxByName = /\.docx$/i.test(file.name || '');
    const isDocxByType = /officedocument\.wordprocessingml\.document/.test(file.type || '');
    if (!(isDocxByName || isDocxByType)) { alert('Please provide a .docx file.'); return; }
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer }, { styleMap: APP_CONFIG.styleMap });
    const processed = processHtml(result.value);
    setRCEHtml(processed);
    if (fileInput) fileInput.value = '';
    showToast('Converted & inserted into the RCE.');
  }
  const convertAndReplaceWithSpinner = withSpinner(convertAndReplace);

  const runReport = () => {
    const html = getRCEHtml();
    if (!html) { alert('No content in the editor.'); return; }
    document.getElementById('cdt-modal-title').textContent = 'Section Report';
    document.getElementById('cdt-modal-body').innerHTML = generateSectionPresenceReport(html);
    modal.style.display = 'flex';
  };

  const cleanupNow = () => {
    const html = getRCEHtml();
    if (!html) { alert('No content in the editor.'); return; }
    lastPreCleanupHtml = html;
    const cleaned = removeExtraContent(html);
    setRCEHtml(cleaned);
    btnCleanup.style.display = 'none';
    btnUndo.style.display = 'inline-block';
    showToast('Cleanup applied. You can Undo.');
  };

  const undoCleanup = () => {
    if (lastPreCleanupHtml) {
      setRCEHtml(lastPreCleanupHtml);
      lastPreCleanupHtml = null;
      btnUndo.style.display = 'none';
      btnCleanup.style.display = 'inline-block';
      showToast('Cleanup undone.');
    }
  };

  const saveToPage = withSpinner(async () => {
    const titleEl = document.querySelector('#wikipage-title-input');
    const title = titleEl ? titleEl.value : undefined;
    const html = getRCEHtml();
    if (!html) { alert('Editor is empty. Nothing to save.'); return; }

    // Prefer API save (PUT /api/v1/courses/:course_id/pages/:url)
    if (CanvasCtx.courseId && CanvasCtx.pageSlug) {
      const res = await fetch(`${CanvasCtx.apiBase}/courses/${CanvasCtx.courseId}/pages/${CanvasCtx.pageSlug}`, {
        method: 'PUT',
        headers: {
          'Content-Type':'application/json',
          'Accept': 'application/json',
          'X-CSRF-Token': getCsrfToken()
        },
        body: JSON.stringify({ wiki_page: { body: html, ...(title ? { title } : {}) } })
      });
      if (!res.ok) throw new Error(`Save failed: ${res.status}`);
      showToast('Saved to page via API.');
    } else {
      // Fallback: click Canvas save button if present
      const btn = [...document.querySelectorAll('button, input[type=submit]')].find(b => /save/i.test(b.textContent || b.value || ''));
      if (btn) { btn.click(); showToast('Triggered Canvas Save button.'); }
      else { alert('Could not determine page context to save.'); }
    }
  });

  // ---- Drag & Drop wiring ----
  function preventDefaults(e){ e.preventDefault(); e.stopPropagation(); }
  function highlight(){ root.classList.add('drop-active'); }
  function unhighlight(){ root.classList.remove('drop-active'); }
  function handleDrop(e){ preventDefaults(e); unhighlight();
    const dt = e.dataTransfer; if (!dt || !dt.files || !dt.files.length) return;
    const file = [...dt.files].find(f => /\.docx$/i.test(f.name));
    if (!file) { showToast('Drop a .docx file.'); return; }
    convertAndReplaceWithSpinner(file);
  }
  ;['dragenter','dragover'].forEach(ev => root.addEventListener(ev, (e)=>{preventDefaults(e); highlight();}));
  ;['dragleave','drop'].forEach(ev => root.addEventListener(ev, (e)=>{preventDefaults(e); if(ev==='dragleave') unhighlight();}));
  root.addEventListener('drop', handleDrop);

  // Also allow dropping inside the TinyMCE iframe body
  function attachIframeDnD(){
    if (STATE.iframeDnDAttached) return;
    const ed = getEditor(); if (!ed || !ed.getDoc) return;
    const edDoc = ed.getDoc(); if (!edDoc || !edDoc.body) return;
    const body = edDoc.body; if (body.__cdtDnD) { STATE.iframeDnDAttached = true; return; }
    body.__cdtDnD = true;
    ['dragenter','dragover'].forEach(ev => body.addEventListener(ev, (e)=>{ preventDefaults(e); highlight(); }));
    ['dragleave','drop'].forEach(ev => body.addEventListener(ev, (e)=>{ preventDefaults(e); if(ev==='dragleave') unhighlight(); }));
    body.addEventListener('drop', (e)=>{
      const dt = e.dataTransfer; if (!dt || !dt.files || !dt.files.length) return; unhighlight();
      const file = [...dt.files].find(f => /\.docx$/i.test(f.name));
      if (!file) { ed.windowManager && ed.windowManager.alert ? ed.windowManager.alert('Drop a .docx file.') : showToast('Drop a .docx file.'); return; }
      convertAndReplaceWithSpinner(file);
    });
    STATE.iframeDnDAttached = true;
  }
  const dndInterval = setInterval(()=>{
    try { attachIframeDnD(); if (STATE.iframeDnDAttached) clearInterval(dndInterval); } catch(_){}
  }, 1000);

  // ---------------------------------------------------------------------------
  // 6) Wire up buttons (+ style controls)
  // Auto-convert when a file is selected
  if (fileInput) {
    fileInput.addEventListener('change', (e) => {
      if (fileInput.files && fileInput.files[0]) {
        convertAndReplaceWithSpinner();
      }
    });
  
  }

  btnReport.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); runReport(); });
  btnCleanup.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation();
    document.getElementById('cdt-modal-title').textContent = 'Confirm Cleanup';
    document.getElementById('cdt-modal-body').innerHTML = `
      <p>This will remove any content (and its heading) that is not in a standard section like
      "Learning Objectives", "Activities and Resources", or "Assignments".</p>
      <p>You can undo immediately after.</p>
      <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:8px">
        <button class="btn gray" type="button" id="cdt-cancel-clean">Cancel</button>
        <button class="btn" type="button" id="cdt-confirm-clean">Yes, Clean Up</button>
      </div>`;
    modal.style.display = 'flex';
    const cancel = document.getElementById('cdt-cancel-clean');
    const ok     = document.getElementById('cdt-confirm-clean');
    cancel.onclick = () => { modal.style.display = 'none'; };
    ok.onclick = () => { modal.style.display = 'none'; cleanupNow(); };
  });
  btnUndo.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); undoCleanup(); });

  btnToggleStyles.addEventListener('click', (e)=>{ e.preventDefault(); e.stopPropagation();
    STATE.previewOn = !STATE.previewOn;
    setPreviewStyles(STATE.previewOn);
    btnToggleStyles.textContent = `Preview Styles: ${STATE.previewOn ? 'On' : 'Off'}`;
  });
  btnStyleLegend.addEventListener('click', (e)=>{ e.preventDefault(); e.stopPropagation(); openStyleLegend(); });
  btnSave.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); saveToPage(); });

  // Start with preview styles ON
  setPreviewStyles(true);

  console.log('Canvas DOCX → RCE Toolbelt loaded (v1.4.1).');
}
// === end =====================================================================
