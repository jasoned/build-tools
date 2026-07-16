// === Canvas DOCX → RCE Toolbelt =============================================
// PASTE THIS IN THE BROWSER CONSOLE WHILE EDITING A CANVAS PAGE
// It adds a floating UI above the RCE to upload/convert/process DOCX content,
// run your post-processing tools on current content, preview styles,
// and save back via API. Now supports drag & drop of .docx files and
// wrapping the current selection in a section class, plus Find/Replace.
//
// Version: 1.8.3 (source-aware false-bullet removal + list preservation)
//
// -----------------------------------------------------------------------------
// 0) Helpers: CSRF + URL context  [UNCHANGED FROM YOUR REQUEST]
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
  console.log(`Canvas DOCX Toolbelt already loaded (${window.__CanvasDocxToolsLoaded}). Reload the page before loading a newer version.`);
} else {
  const TOOLBELT_VERSION = '1.8.3';
  window.__CanvasDocxToolsLoaded = TOOLBELT_VERSION;
  const MAMMOTH_VERSION = '1.7.2';

  // ---------------------------------------------------------------------------
  // 1) Inject Mammoth if missing
  function ensureMammoth() {
    return new Promise((resolve, reject) => {
      if (window.mammoth && window.__CanvasDocxToolsMammothVersion === MAMMOTH_VERSION) return resolve();
      const s = document.createElement('script');
      s.src = `https://cdnjs.cloudflare.com/ajax/libs/mammoth/${MAMMOTH_VERSION}/mammoth.browser.min.js`;
      s.onload = () => {
        window.__CanvasDocxToolsMammothVersion = MAMMOTH_VERSION;
        resolve();
      };
      s.onerror = () => reject(new Error('Failed to load Mammoth.js'));
      document.head.appendChild(s);
    });
  }

  function ensureJSZip() {
    return new Promise((resolve, reject) => {
      if (window.JSZip) return resolve();
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('Failed to load JSZip'));
      document.head.appendChild(s);
    });
  }

  // ---------------------------------------------------------------------------
  // 2) UI Shell (top toolbar + modal + spinner + toast)
  let root = document.getElementById('canvas-docx-tools');
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
#cdt-modal .inner { background:#fff; max-width:760px; width:92vw; border-radius:10px; padding:18px; }
#cdt-modal h3 { margin:0 0 8px; font:600 18px/1.2 system-ui, -apple-system, Segoe UI, Roboto; }
#cdt-modal .body { max-height:60vh; overflow:auto; border:1px solid #e2e8f0; border-radius:8px; padding:10px; background:#f8fafc; }
#cdt-modal .row { margin-top:10px; display:flex; gap:8px; justify-content:flex-end; }
#cdt-modal .btn { background:#225f92; color:#fff; border:0; border-radius:8px; padding:8px 10px; cursor:pointer; font-size:12px; }
#cdt-modal .btn.gray { background:#6c757d; }
#cdt-modal table { border-collapse:collapse; width:100%; font-size:12px; }
#cdt-modal th, #cdt-modal td { border:1px solid #cfd4da; padding:6px; text-align:left; }

/* Find/Replace inputs */
#cdt-fr-form { display:grid; grid-template-columns: 1fr 1fr; gap:8px; align-items:end; margin-bottom:8px; }
#cdt-fr-form .full { grid-column: 1 / -1; }
#cdt-fr-form label { display:block; font-size:12px; color:#2d3748; margin-bottom:2px; }
#cdt-fr-form input[type="text"] { width:100%; padding:6px 8px; font-size:13px; border:1px solid #cbd5e0; border-radius:6px; }
#cdt-fr-form .opts { display:flex; gap:12px; align-items:center; font-size:12px; }
#cdt-fr-actions { display:flex; gap:8px; justify-content:flex-end; }
`;
  const style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);

  if (!root) {
    root = document.createElement('div');
    root.id = 'canvas-docx-tools';
    root.innerHTML = `
      <div class="cdt-wrap" title="You can drop a .docx file on this bar">
        <label>DOCX:</label>
        <input id="cdt-file" type="file" accept=".docx" />
        <span class="cdt-note cdt-drop-tip">choose or drop a .docx (auto-converts)</span>
        <button class="cdt-btn" type="button" id="cdt-report">Section Report</button>
        <button class="cdt-btn warn" type="button" id="cdt-cleanup">Clean Up Formatting</button>
        <button class="cdt-btn" type="button" id="cdt-undo" style="display:none">Undo Cleanup</button>

        <span class="cdt-spacer"></span>

        <label for="cdt-wrap-class">Wrap selection:</label>
        <select id="cdt-wrap-class" title="Select wrapper class for highlighted content"></select>
        <button class="cdt-btn" type="button" id="cdt-wrap">Apply</button>

        <button class="cdt-btn secondary" type="button" id="cdt-find-replace">Find/Replace</button>
        <button class="cdt-btn secondary" type="button" id="cdt-toggle-styles">Preview Styles: On</button>
        <button class="cdt-btn secondary" type="button" id="cdt-style-legend">Styles Legend</button>
        <button class="cdt-btn success" type="button" id="cdt-save">Save to Page</button>
        <span class="cdt-note">Toolbelt v${TOOLBELT_VERSION}</span>
        <span class="cdt-note">Page: ${CanvasCtx.courseId ?? 'unknown course'} / ${CanvasCtx.pageSlug ?? 'unknown page'}</span>
      </div>
    `;
    const mountTarget = document.querySelector('.edit-content') || document.body;
    mountTarget.prepend(root);
  }

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
  // 3) RCE bridge + wait
  function getEditor() {
    if (window.tinymce) {
      const ed = tinymce.get('wiki_page_body') || tinymce.activeEditor;
      if (ed && !ed.destroyed) return ed;
    }
    return null;
  }

  async function waitForEditor(ms=8000) {
    const start = Date.now();
    while (Date.now() - start < ms) {
      const ed = getEditor();
      if (ed && ed.getContent) return ed;
      await new Promise(r => setTimeout(r, 150));
    }
    return null;
  }

  function getRCEHtml() {
    const ed = getEditor();
    if (ed) return ed.getContent({format:'html'});
    const ta = document.getElementById('wiki_page_body') || document.querySelector('textarea');
    return ta ? ta.value : '';
  }
  function setRCEHtml(html) {
    const ed = getEditor();
    if (ed) ed.setContent(html);
    const ta = document.getElementById('wiki_page_body') || document.querySelector('textarea');
    if (ta) ta.value = html;
    setPreviewStyles(STATE.previewOn);
  }

  // Reapply preview CSS if TinyMCE reloads
  (function observeEditorReload(){
    const rootEl = document.getElementById('content') || document.body;
    if (!rootEl) return;
    const mo = new MutationObserver(() => setPreviewStyles(STATE.previewOn));
    mo.observe(rootEl, { childList: true, subtree: true });
  })();

  // ---------------------------------------------------------------------------
  // 4) Processing pipeline
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
    ],
    labelAliases: {
      cdes: ['course description'],
      plo: ['program learning outcome', 'program learning outcomes'],
      clo: ['course learning outcome', 'course learning outcomes'],
      material: ['required course material', 'required course materials'],
      LO: ['learning objective', 'learning objectives', 'weekly learning objective', 'weekly learning objectives'],
      activities: ['activities and resources', 'learning activities and resources'],
      assignments: ['assignment', 'assignments']
    }
  };
  APP_CONFIG.allowedWrapperSelectors = [
    '#wk-topics-list',
    '.misc',
    ...APP_CONFIG.topLevelSections.map(s => `.${s.className}`),
    ...APP_CONFIG.weeklySections.map(s => `.${s.className}`)
  ].join(', ');

  // Expose config to window for quick tweaks
  window.CanvasDocxTools = window.CanvasDocxTools || {};
  window.CanvasDocxTools.config = APP_CONFIG;

  const textOnly = (el) => (el.textContent || '').trim().replace(/\s+/g, ' ');
  const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, ch => ({
    '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'
  })[ch]);

  function normalizeStructureText(value) {
    return String(value ?? '')
      .normalize('NFKC')
      .replace(/\u00a0/g, ' ')
      .replace(/[–—]/g, '-')
      .replace(/&/g, ' and ')
      .trim()
      .toLowerCase()
      .replace(/[.:]+$/g, '')
      .replace(/\s+/g, ' ');
  }

  function normalizeListMatchText(value) {
    return String(value ?? '')
      .normalize('NFKC')
      .replace(/[\u00ad\u200b]/g, '')
      .replace(/\u00a0/g, ' ')
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201c\u201d]/g, '"')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  }

  function xmlDirectChildren(element, localName) {
    return Array.from(element?.childNodes || []).filter(node =>
      node.nodeType === 1 && (!localName || node.localName === localName)
    );
  }

  function xmlDirectChild(element, localName) {
    return xmlDirectChildren(element, localName)[0] || null;
  }

  function xmlDescendants(element, localName) {
    return Array.from(element?.getElementsByTagName('*') || []).filter(node => node.localName === localName);
  }

  function xmlAttribute(element, localName) {
    const attribute = Array.from(element?.attributes || []).find(attr => attr.localName === localName);
    return attribute ? attribute.value : null;
  }

  function parseWordXml(xml, filename) {
    const parsed = new DOMParser().parseFromString(xml, 'application/xml');
    if (parsed.getElementsByTagName('parsererror').length) throw new Error(`Unable to parse ${filename}`);
    return parsed;
  }

  function readWordIndent(properties) {
    const indent = xmlDirectChild(properties, 'ind');
    if (!indent) return null;
    const start = xmlAttribute(indent, 'start') ?? xmlAttribute(indent, 'left');
    return {
      start: start == null || start === '' ? null : Number(start),
      hanging: Number(xmlAttribute(indent, 'hanging')) || null
    };
  }

  function readWordNumberingProperties(properties) {
    const numPr = xmlDirectChild(properties, 'numPr');
    if (!numPr) return null;
    return {
      numId: xmlAttribute(xmlDirectChild(numPr, 'numId'), 'val'),
      level: xmlAttribute(xmlDirectChild(numPr, 'ilvl'), 'val')
    };
  }

  function readWordNumberingLevel(levelElement) {
    if (!levelElement) return {};
    const properties = xmlDirectChild(levelElement, 'pPr');
    const startValue = xmlAttribute(xmlDirectChild(levelElement, 'start'), 'val');
    return {
      format: xmlAttribute(xmlDirectChild(levelElement, 'numFmt'), 'val'),
      levelText: xmlAttribute(xmlDirectChild(levelElement, 'lvlText'), 'val'),
      start: startValue == null ? null : Number(startValue),
      indent: readWordIndent(properties)
    };
  }

  function parseWordNumbering(numberingXml) {
    const abstractNums = new Map();
    xmlDescendants(numberingXml, 'abstractNum').forEach(abstractElement => {
      const levels = new Map();
      xmlDirectChildren(abstractElement, 'lvl').forEach(levelElement => {
        levels.set(xmlAttribute(levelElement, 'ilvl') || '0', readWordNumberingLevel(levelElement));
      });
      abstractNums.set(xmlAttribute(abstractElement, 'abstractNumId'), levels);
    });

    const nums = new Map();
    xmlDescendants(numberingXml, 'num').forEach(numElement => {
      const overrides = new Map();
      xmlDirectChildren(numElement, 'lvlOverride').forEach(overrideElement => {
        const level = xmlAttribute(overrideElement, 'ilvl') || '0';
        const override = readWordNumberingLevel(xmlDirectChild(overrideElement, 'lvl'));
        const startOverride = xmlAttribute(xmlDirectChild(overrideElement, 'startOverride'), 'val');
        if (startOverride != null) override.start = Number(startOverride);
        overrides.set(level, override);
      });
      nums.set(xmlAttribute(numElement, 'numId'), {
        abstractNumId: xmlAttribute(xmlDirectChild(numElement, 'abstractNumId'), 'val'),
        overrides
      });
    });

    return (numId, level) => {
      const num = nums.get(String(numId));
      if (!num) return {};
      const base = abstractNums.get(num.abstractNumId)?.get(String(level)) || {};
      const override = num.overrides.get(String(level)) || {};
      const merged = { ...base };
      Object.entries(override).forEach(([key, value]) => {
        if (value != null) merged[key] = value;
      });
      return merged;
    };
  }

  function parseWordParagraphStyles(stylesXml) {
    if (!stylesXml) return () => ({});
    const styles = new Map();
    xmlDescendants(stylesXml, 'style').forEach(styleElement => {
      if (xmlAttribute(styleElement, 'type') !== 'paragraph') return;
      const properties = xmlDirectChild(styleElement, 'pPr');
      styles.set(xmlAttribute(styleElement, 'styleId'), {
        basedOn: xmlAttribute(xmlDirectChild(styleElement, 'basedOn'), 'val'),
        numbering: readWordNumberingProperties(properties),
        indent: readWordIndent(properties)
      });
    });
    const resolved = new Map();
    const resolve = (styleId, trail = new Set()) => {
      if (!styleId || trail.has(styleId)) return {};
      if (resolved.has(styleId)) return resolved.get(styleId);
      const style = styles.get(styleId);
      if (!style) return {};
      const nextTrail = new Set(trail); nextTrail.add(styleId);
      const base = resolve(style.basedOn, nextTrail);
      const value = {
        numbering: style.numbering || base.numbering || null,
        indent: style.indent || base.indent || null
      };
      resolved.set(styleId, value);
      return value;
    };
    return resolve;
  }

  async function extractDocxListMetadata(arrayBuffer) {
    await ensureJSZip();
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const documentEntry = zip.file('word/document.xml');
    const numberingEntry = zip.file('word/numbering.xml');
    if (!documentEntry || !numberingEntry) return [];
    const [documentText, numberingText, stylesText] = await Promise.all([
      documentEntry.async('string'),
      numberingEntry.async('string'),
      zip.file('word/styles.xml')?.async('string') || Promise.resolve(null)
    ]);
    const documentXml = parseWordXml(documentText, 'word/document.xml');
    const numberingXml = parseWordXml(numberingText, 'word/numbering.xml');
    const stylesXml = stylesText ? parseWordXml(stylesText, 'word/styles.xml') : null;
    const findNumberingLevel = parseWordNumbering(numberingXml);
    const findParagraphStyle = parseWordParagraphStyles(stylesXml);
    const result = [];

    xmlDescendants(documentXml, 'p').forEach((paragraph, sequence) => {
      const properties = xmlDirectChild(paragraph, 'pPr');
      const styleId = xmlAttribute(xmlDirectChild(properties, 'pStyle'), 'val');
      const style = findParagraphStyle(styleId);
      const directNumbering = readWordNumberingProperties(properties);
      const text = xmlDescendants(paragraph, 't').map(node => node.textContent || '').join('');
      const matchText = normalizeListMatchText(text);
      if (!matchText) return;
      if (directNumbering?.numId === '0') {
        result.push({ sequence, text, matchText, isList: false });
        return;
      }
      const numbering = directNumbering || style.numbering;
      const numId = numbering?.numId;
      if (numId == null) {
        // Keep a neutral placeholder so repeated paragraph text remains aligned
        // with the corresponding HTML nodes. Without it, an ordinary paragraph
        // can consume the later explicit numId=0 marker for a false list item.
        result.push({ sequence, text, matchText, isList: null });
        return;
      }
      const level = numbering.level ?? style.numbering?.level ?? '0';
      const definition = findNumberingLevel(numId, level);
      const directIndent = readWordIndent(properties);
      const left = directIndent?.start ?? definition.indent?.start ?? style.indent?.start ?? null;
      result.push({
        sequence,
        text,
        matchText,
        isList: true,
        numId: String(numId),
        level: Number(level) || 0,
        format: definition.format || 'decimal',
        levelText: definition.levelText || null,
        start: Number.isFinite(definition.start) ? definition.start : null,
        left: Number.isFinite(left) ? left : null
      });
    });
    return result;
  }

  function listItemOwnText(item) {
    const clone = item.cloneNode(true);
    clone.querySelectorAll('ol, ul').forEach(list => list.remove());
    return textOnly(clone);
  }

  function listIndent(metadata) {
    return Number.isFinite(metadata?.left) ? metadata.left : (Number(metadata?.level) || 0) * 720;
  }

  function listFormatKey(metadata) {
    return metadata?.format || 'decimal';
  }

  function applyWordListFormat(list, metadata) {
    if (!list || !metadata) return;
    const htmlTypes = { lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I' };
    const cssTypes = {
      lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha',
      lowerRoman: 'lower-roman', upperRoman: 'upper-roman',
      lowerGreek: 'lower-greek', decimal: 'decimal', bullet: 'disc'
    };
    if (list.tagName === 'OL') {
      const type = htmlTypes[metadata.format];
      if (type) list.setAttribute('type', type); else list.removeAttribute('type');
      const cssType = cssTypes[metadata.format];
      if (cssType && !type) list.style.listStyleType = cssType;
      else list.style.removeProperty('list-style-type');
      if (Number.isFinite(metadata.start) && metadata.start > 1) list.setAttribute('start', String(metadata.start));
      else list.removeAttribute('start');
    }
  }

  function createWordList(doc, metadata) {
    const list = doc.createElement(metadata?.format === 'bullet' ? 'ul' : 'ol');
    applyWordListFormat(list, metadata);
    return list;
  }

  function appendListItemContentAsBlocks(item, target) {
    const doc = item.ownerDocument;
    const blockTags = new Set(['P', 'DIV', 'SECTION', 'ARTICLE', 'UL', 'OL', 'TABLE']);
    let paragraph = null;
    const flushParagraph = () => {
      if (paragraph && (paragraph.textContent.trim() || paragraph.children.length)) target.appendChild(paragraph);
      paragraph = null;
    };
    const ensureParagraph = () => (paragraph ||= doc.createElement('p'));
    while (item.firstChild) {
      const child = item.firstChild;
      item.removeChild(child);
      if (child.nodeType === Node.ELEMENT_NODE && blockTags.has(child.tagName)) {
        flushParagraph();
        target.appendChild(child);
      } else if (child.nodeType === Node.TEXT_NODE && !child.textContent.trim() && !paragraph) {
        continue;
      } else {
        ensureParagraph().appendChild(child);
      }
    }
    flushParagraph();
  }

  function repairFalseWordListItems(doc, metadataByItem, syntheticItems) {
    Array.from(doc.querySelectorAll('ol, ul')).reverse().forEach(list => {
      const items = Array.from(list.children).filter(child => child.tagName === 'LI');
      const shouldUnwrap = item => metadataByItem.get(item)?.isList === false || syntheticItems.has(item);
      if (!items.some(shouldUnwrap)) return;

      const replacement = doc.createDocumentFragment();
      let listChunk = null;
      const flushList = () => {
        if (listChunk?.children.length) replacement.appendChild(listChunk);
        listChunk = null;
      };
      const ensureList = () => {
        if (listChunk) return listChunk;
        listChunk = doc.createElement(list.tagName.toLowerCase());
        Array.from(list.attributes).forEach(attribute => listChunk.setAttribute(attribute.name, attribute.value));
        return listChunk;
      };

      items.forEach(item => {
        if (!shouldUnwrap(item)) {
          ensureList().appendChild(item);
          return;
        }
        flushList();
        appendListItemContentAsBlocks(item, replacement);
        item.remove();
      });
      flushList();
      list.replaceWith(replacement);
    });
  }

  function mergeAdjacentWordLists(doc) {
    Array.from(doc.querySelectorAll('ol, ul')).forEach(list => {
      let next = list.nextElementSibling;
      while (next && next.tagName === list.tagName) {
        const currentType = list.getAttribute('type') || '';
        const nextType = next.getAttribute('type') || '';
        const currentStyle = list.getAttribute('style') || '';
        const nextStyle = next.getAttribute('style') || '';
        if (currentType !== nextType || currentStyle !== nextStyle) break;
        while (next.firstChild) list.appendChild(next.firstChild);
        const consumed = next;
        next = next.nextElementSibling;
        consumed.remove();
      }
    });
  }

  function repairWordList(list, metadataByItem) {
    const items = Array.from(list.children).filter(child => child.tagName === 'LI');
    const firstIndex = items.findIndex(item => metadataByItem.get(item)?.isList !== false && metadataByItem.has(item));
    if (firstIndex < 0) return;
    const firstItem = items[firstIndex];
    const firstMetadata = metadataByItem.get(firstItem);
    applyWordListFormat(list, firstMetadata);
    const stack = [{
      indent: listIndent(firstMetadata),
      list,
      format: listFormatKey(firstMetadata),
      numId: firstMetadata.numId,
      lastLi: null
    }];

    items.forEach(item => {
      const metadata = metadataByItem.get(item);
      if (!metadata || metadata.isList === false) {
        stack[stack.length - 1].lastLi = item;
        return;
      }
      const indent = listIndent(metadata);
      while (stack.length > 1 && indent < stack[stack.length - 1].indent) stack.pop();
      let frame = stack[stack.length - 1];

      if (indent > frame.indent && frame.lastLi) {
        const nestedList = createWordList(list.ownerDocument, metadata);
        frame.lastLi.appendChild(nestedList);
        frame = { indent, list: nestedList, format: listFormatKey(metadata), numId: metadata.numId, lastLi: null };
        stack.push(frame);
      } else if (indent === frame.indent && frame.lastLi &&
        (listFormatKey(metadata) !== frame.format || metadata.numId !== frame.numId)) {
        const siblingList = createWordList(list.ownerDocument, metadata);
        frame.list.insertAdjacentElement('afterend', siblingList);
        frame.list = siblingList;
        frame.format = listFormatKey(metadata);
        frame.numId = metadata.numId;
        frame.lastLi = null;
      }

      frame.list.appendChild(item);
      frame.lastLi = item;
    });
  }

  function restoreWordListStructure(rawHtml, listMetadata) {
    if (!Array.isArray(listMetadata) || !listMetadata.length) return rawHtml;
    const doc = new DOMParser().parseFromString(rawHtml, 'text/html');
    const queues = new Map();
    listMetadata.forEach(metadata => {
      const queue = queues.get(metadata.matchText) || [];
      queue.push(metadata); queues.set(metadata.matchText, queue);
    });
    const metadataByItem = new Map();
    const syntheticItems = new Set();
    doc.querySelectorAll('p, li').forEach(item => {
      const ownText = item.tagName === 'LI' ? listItemOwnText(item) : textOnly(item);
      if (item.tagName === 'LI' && !ownText && item.querySelector(':scope > ol, :scope > ul')) syntheticItems.add(item);
      const queue = queues.get(normalizeListMatchText(ownText));
      if (queue?.length) {
        const metadata = queue.shift();
        if (metadata.isList !== null) metadataByItem.set(item, metadata);
      }
    });
    repairFalseWordListItems(doc, metadataByItem, syntheticItems);
    mergeAdjacentWordLists(doc);
    Array.from(doc.querySelectorAll('ol, ul')).forEach(list => repairWordList(list, metadataByItem));
    return doc.body.innerHTML;
  }

  function parseWeekHeading(value) {
    const text = String(value ?? '').replace(/\u00a0/g, ' ').trim();
    const match = text.match(/^week\s*(\d+)\s*(?::|[-–—])?\s*(.+)$/i);
    if (!match) return null;
    return { number: Number(match[1]), title: match[2].trim() };
  }

  function aliasClassForText(value, sectionDefs) {
    let normalized = normalizeStructureText(value);
    normalized = normalized.replace(/\s*\([^)]*\)\s*$/, '').trim();
    for (const section of sectionDefs) {
      const aliases = APP_CONFIG.labelAliases[section.className] || [];
      if (aliases.includes(normalized)) return section.className;
    }
    return null;
  }

  function addElementFix(el, message) {
    if (!el || !message) return;
    let fixes = [];
    try { fixes = JSON.parse(el.getAttribute('data-cdt-fixes') || '[]'); } catch (_) {}
    if (!fixes.includes(message)) fixes.push(message);
    el.setAttribute('data-cdt-fixes', JSON.stringify(fixes));
  }

  function getElementFixes(rootEl) {
    const fixes = [];
    if (!rootEl) return fixes;
    const candidates = [rootEl, ...rootEl.querySelectorAll('[data-cdt-fixes]')];
    candidates.forEach(el => {
      try {
        const parsed = JSON.parse(el.getAttribute('data-cdt-fixes') || '[]');
        parsed.forEach(fix => { if (fix && !fixes.includes(fix)) fixes.push(fix); });
      } catch (_) {}
    });
    return fixes;
  }

  function retagElement(doc, el, tagName) {
    if (el.tagName.toLowerCase() === tagName.toLowerCase()) return el;
    const replacement = doc.createElement(tagName);
    Array.from(el.attributes).forEach(attr => replacement.setAttribute(attr.name, attr.value));
    while (el.firstChild) replacement.appendChild(el.firstChild);
    el.replaceWith(replacement);
    return replacement;
  }

  function canonicalizeStructureLabels(doc) {
    let currentWeek = null;
    Array.from(doc.body.children).forEach(original => {
      if (!original.matches('p, h1, h2, h3, h4, h5, h6')) return;
      const originalText = textOnly(original);
      const week = parseWeekHeading(originalText);
      if (week) {
        let heading = retagElement(doc, original, 'h1');
        const canonicalText = `Week ${week.number}: ${week.title}`;
        const neededFix = !original.matches('h1.wk-topic') || originalText !== canonicalText;
        heading.classList.add('wk-topic');
        heading.setAttribute('data-cdt-week-number', String(week.number));
        heading.setAttribute('data-cdt-week-title', week.title);
        if (heading.textContent.trim() !== canonicalText) heading.textContent = canonicalText;
        if (neededFix) addElementFix(heading, `Week ${week.number}: heading recognized from text instead of the expected Word style.`);
        currentWeek = week.number;
        return;
      }

      const defs = currentWeek ? APP_CONFIG.weeklySections : APP_CONFIG.topLevelSections;
      const className = aliasClassForText(originalText, defs);
      if (!className) return;
      const section = defs.find(item => item.className === className);
      let heading = retagElement(doc, original, 'h1');
      heading.setAttribute('data-cdt-section-label', className);
      const canonicalText = section.name;
      const labelChanged = normalizeStructureText(originalText) !== normalizeStructureText(canonicalText);
      const tagChanged = original.tagName !== 'H1';
      if (heading.textContent.trim() !== canonicalText) heading.textContent = canonicalText;
      if (currentWeek && (labelChanged || tagChanged)) {
        addElementFix(heading, `Week ${currentWeek}: recognized “${originalText}” as “${canonicalText}”.`);
      } else if (!currentWeek && (labelChanged || tagChanged)) {
        addElementFix(heading, `Recognized “${originalText}” as the ${canonicalText} heading.`);
      }
    });
  }

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
    const ownerDoc = dst.ownerDocument || src.ownerDocument || document;
    let currentP = null;
    const flushP = () => { if (currentP && currentP.childNodes.length) dst.appendChild(currentP); currentP = null; };
    const ensureP = () => (currentP ||= ownerDoc.createElement('p'));
    function walk(node) {
      while (node.firstChild) {
        const child = node.firstChild; node.removeChild(child);
        if (child.nodeType === Node.TEXT_NODE) {
          const value = child.textContent || '';
          if (value.trim()) ensureP().appendChild(ownerDoc.createTextNode(value));
          else if (currentP?.childNodes.length && node.firstChild) currentP.appendChild(ownerDoc.createTextNode(' '));
          continue;
        }
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

  function mergeAdjacentInlineFormatting(root) {
    const mergeableTags = new Set(['STRONG', 'B', 'EM', 'I', 'U']);
    const sameAttributes = (left, right) => {
      if (left.attributes.length !== right.attributes.length) return false;
      return Array.from(left.attributes).every(attribute => right.getAttribute(attribute.name) === attribute.value);
    };
    const missingDeadlineSpace = (leftText, rightText) =>
      (/\bof$/i.test(leftText) && /^Week\b/i.test(rightText)) ||
      (/\bWeek$/i.test(leftText) && /^\d+\b/.test(rightText));

    root.querySelectorAll('p, li, h1, h2, h3, h4, h5, h6, td, th').forEach(container => {
      let current = container.firstChild;
      while (current) {
        if (current.nodeType !== Node.ELEMENT_NODE || !mergeableTags.has(current.tagName)) {
          current = current.nextSibling;
          continue;
        }

        const separator = current.nextSibling;
        const hasWhitespaceSeparator = separator?.nodeType === Node.TEXT_NODE && !separator.textContent.trim();
        const candidate = hasWhitespaceSeparator ? separator.nextSibling : separator;
        if (candidate?.nodeType !== Node.ELEMENT_NODE || candidate.tagName !== current.tagName || !sameAttributes(current, candidate)) {
          current = current.nextSibling;
          continue;
        }

        const needsSpace = hasWhitespaceSeparator || missingDeadlineSpace(current.textContent, candidate.textContent);
        if (hasWhitespaceSeparator) separator.remove();
        if (needsSpace && !/\s$/.test(current.textContent || '') && !/^\s/.test(candidate.textContent || '')) {
          current.appendChild(current.ownerDocument.createTextNode(' '));
        }
        while (candidate.firstChild) current.appendChild(candidate.firstChild);
        candidate.remove();
      }
    });
  }

  function headingMatches(h, target) {
    const t = h.textContent.trim().toLowerCase().replace(/[:]+$/,'');
    const k = target.trim().toLowerCase();
    return t === k;
  }

  function wrapSectionContent(doc, headingText, wrapperClass) {
    const heading = Array.from(doc.body.children).find(el => el.matches(`h1[data-cdt-section-label="${wrapperClass}"]`) || (el.tagName === 'H1' && headingMatches(el, headingText)));
    if (!heading) return;
    const existing = heading.nextElementSibling;
    if (existing && existing.classList.contains(wrapperClass)) return;
    const wrapper = doc.createElement('div'); wrapper.className = wrapperClass;
    wrapper.setAttribute('data-cdt-section', wrapperClass);
    let currentNode = heading.nextElementSibling; const move = [];
    while (currentNode && currentNode.tagName !== 'H1') { move.push(currentNode); currentNode = currentNode.nextElementSibling; }
    move.forEach(el => wrapper.appendChild(el));
    if (wrapper.hasChildNodes()) heading.insertAdjacentElement('afterend', wrapper);
  }

  function cleanupHtml(doc) {
    mergeAdjacentInlineFormatting(doc);
    doc.querySelectorAll('p, a').forEach(el => { if (!el.textContent.trim() && !el.children.length) el.remove(); });
    return doc;
  }

  function getWeeklySegments(doc) {
    const headings = Array.from(doc.body.children).filter(el => el.matches('h1.wk-topic'));
    return headings.map(heading => {
      const nodes = [];
      let cur = heading.nextElementSibling;
      while (cur && !cur.matches('h1.wk-topic')) { nodes.push(cur); cur = cur.nextElementSibling; }
      const parsed = parseWeekHeading(heading.textContent) || {};
      return {
        heading,
        number: parsed.number || Number(heading.getAttribute('data-cdt-week-number')),
        title: parsed.title || heading.getAttribute('data-cdt-week-title') || textOnly(heading),
        nodes
      };
    });
  }

  function appendTableAsSectionItem(doc, table, wrapper, addTrailingRule) {
    const itemDiv = doc.createElement('div');
    const rows = Array.from(table.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tr'));
    if (rows.length > 0) {
      const titleCell = rows[0].querySelector(':scope > td, :scope > th');
      if (titleCell && titleCell.textContent.trim()) {
        const h2 = doc.createElement('h2'); h2.textContent = textOnly(titleCell); itemDiv.appendChild(h2);
      }
      for (let i = 1; i < rows.length; i++) {
        const contentCell = rows[i].querySelector(':scope > td, :scope > th');
        if (contentCell) moveChildrenPreservingBlocks(contentCell, itemDiv);
      }
    }
    if (itemDiv.hasChildNodes()) wrapper.appendChild(itemDiv);
    if (addTrailingRule && itemDiv.hasChildNodes()) wrapper.appendChild(doc.createElement('hr'));
    table.remove();
  }

  function collectObjectiveTexts(root, addObjective) {
    const listItems = [
      ...(root.matches && root.matches('li') ? [root] : []),
      ...Array.from(root.querySelectorAll ? root.querySelectorAll('li') : [])
    ].filter(li => !li.querySelector('li') && textOnly(li));
    if (listItems.length) {
      listItems.forEach(li => addObjective(textOnly(li)));
      return { count: listItems.length, usedPlainParagraph: false };
    }

    const paragraphs = [
      ...(root.matches && root.matches('p') ? [root] : []),
      ...Array.from(root.querySelectorAll ? root.querySelectorAll('p') : [])
    ].filter(p => textOnly(p));
    if (paragraphs.length) {
      paragraphs.forEach(p => addObjective(textOnly(p)));
      return { count: paragraphs.length, usedPlainParagraph: true };
    }

    const value = textOnly(root);
    if (value) {
      addObjective(value);
      return { count: 1, usedPlainParagraph: true };
    }
    return { count: 0, usedPlainParagraph: false };
  }

  function appendLearningObjectiveContent(doc, nodes, wrapper, weekNumber) {
    const seen = new Set();
    const objectives = [];
    const preserved = [];
    let usedPlainParagraph = false;
    const addObjective = value => {
      const clean = String(value || '').trim().replace(/\s+/g, ' ');
      const key = normalizeStructureText(clean);
      if (!key || seen.has(key)) return false;
      seen.add(key); objectives.push(clean); return true;
    };
    const collect = (root, trackPlainParagraph = true) => {
      const before = objectives.length;
      const result = collectObjectiveTexts(root, addObjective);
      if (trackPlainParagraph && result.usedPlainParagraph && objectives.length > before) usedPlainParagraph = true;
      return objectives.length - before;
    };

    // Rebuild existing LO wrappers too. This removes list+paragraph duplicates
    // left by earlier versions when the tool is run again on current HTML.
    collect(wrapper, false);
    wrapper.replaceChildren();

    nodes.forEach(node => {
      if (node.matches && node.matches('.LO')) {
        collect(node, false);
        node.remove();
        return;
      }
      if (node.tagName === 'TABLE') {
        const before = objectives.length;
        const rows = Array.from(node.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tr'));
        rows.forEach(row => {
          const firstCell = row.querySelector(':scope > td, :scope > th');
          if (firstCell) collect(firstCell);
        });
        if (objectives.length > before) node.remove();
        else preserved.push(node);
        return;
      }
      const before = objectives.length;
      collect(node);
      if (objectives.length > before) node.remove();
      else preserved.push(node);
    });

    objectives.forEach(value => {
      const p = doc.createElement('p'); p.textContent = value; wrapper.appendChild(p);
    });
    preserved.forEach(node => wrapper.appendChild(node));

    if (usedPlainParagraph) addElementFix(wrapper, `Week ${weekNumber}: learning objectives extracted from ordinary paragraphs rather than Word list formatting.`);
    if (objectives.length === 0 && wrapper.querySelector('table')) {
      wrapper.setAttribute('data-cdt-needs-review', 'true');
      addElementFix(wrapper, `Week ${weekNumber}: the learning-objective table was preserved because its content could not be simplified safely.`);
    } else {
      wrapper.removeAttribute('data-cdt-needs-review');
    }
  }

  function appendStandardSectionContent(doc, nodes, wrapper, className) {
    const tables = nodes.filter(node => node.tagName === 'TABLE');
    let tableIndex = 0;
    nodes.forEach(node => {
      if (node.matches && node.matches(`.${className}`)) {
        while (node.firstChild) wrapper.appendChild(node.firstChild);
        node.remove();
        return;
      }
      if (node.tagName !== 'TABLE') {
        wrapper.appendChild(node);
        return;
      }
      const isLast = tableIndex === tables.length - 1;
      appendTableAsSectionItem(doc, node, wrapper, className === 'assignments' || !isLast);
      tableIndex++;
    });
  }

  function processWeeklySections(doc) {
    getWeeklySegments(doc).forEach(segment => {
      const occurrences = [];
      let current = null;
      segment.nodes.forEach(node => {
        const className = node.getAttribute && node.getAttribute('data-cdt-section-label');
        if (className && APP_CONFIG.weeklySections.some(section => section.className === className)) {
          current = { className, heading: node, nodes: [] };
          occurrences.push(current);
        } else if (current) {
          current.nodes.push(node);
        }
      });

      APP_CONFIG.weeklySections.forEach(section => {
        const matches = occurrences.filter(item => item.className === section.className);
        if (!matches.length) return;
        const first = matches[0];
        const existing = first.nodes.find(node => node.matches && node.matches(`.${section.className}`));
        const wrapper = existing || doc.createElement('div');
        wrapper.classList.add(section.className);
        wrapper.setAttribute('data-cdt-section', section.className);
        if (!existing) first.heading.insertAdjacentElement('afterend', wrapper);

        const allNodes = [];
        matches.forEach(match => {
          match.nodes.forEach(node => { if (node !== wrapper) allNodes.push(node); });
          if (match !== first) match.heading.remove();
          getElementFixes(match.heading).forEach(fix => addElementFix(wrapper, fix));
        });
        if (matches.length > 1) addElementFix(wrapper, `Week ${segment.number}: merged ${matches.length} “${section.name}” labels into one section.`);

        if (section.className === 'LO') appendLearningObjectiveContent(doc, allNodes, wrapper, segment.number);
        else appendStandardSectionContent(doc, allNodes, wrapper, section.className);
      });
    });
  }

  function createWeeklyTopicsList(doc) {
    const topicHeaders = doc.querySelectorAll('h1.wk-topic');
    if (topicHeaders.length === 0) return null;

    const old = doc.getElementById('wk-topics-list');
    if (old) old.remove();

    const listContainer = doc.createElement('div'); listContainer.id = 'wk-topics-list';
    const title = doc.createElement('h2'); title.textContent = 'Weekly Topics'; listContainer.appendChild(title);
    const ul = doc.createElement('ul');
    topicHeaders.forEach(h1 => {
      const li = doc.createElement('li');
      li.textContent = h1.textContent.trim().replace(/^Week\s+\d+:\s*/i, '');
      ul.appendChild(li);
    });
    listContainer.appendChild(ul); return listContainer;
  }

  function processHtml(rawHtml) {
    const parser = new DOMParser(); let doc = parser.parseFromString(rawHtml, 'text/html');
    const tableStore = stashNestedTables(doc);

    doc = cleanupHtml(doc);
    canonicalizeStructureLabels(doc);

    APP_CONFIG.topLevelSections.forEach(section => wrapSectionContent(doc, section.name, section.className));
    processWeeklySections(doc);

    restorePlaceholders(doc, tableStore);
    doc = cleanupHtml(doc);

    const weeklyList = createWeeklyTopicsList(doc);
    if (weeklyList) doc.body.prepend(weeklyList);

    // Keep original table border attribute
    doc.querySelectorAll('table').forEach(tbl => tbl.setAttribute('border','1'));

    return doc.body.innerHTML;
  }

  function analyzeDocumentStructure(htmlContent) {
    const parser = new DOMParser(); const doc = parser.parseFromString(htmlContent, 'text/html');
    const recommendations = [];
    const topLevel = APP_CONFIG.topLevelSections.map(section => {
      const count = doc.querySelectorAll(`.${section.className}`).length;
      if (count === 0) recommendations.push(`Top-level section “${section.name}” is missing.`);
      if (count > 1) recommendations.push(`Top-level section “${section.name}” appears ${count} times.`);
      return { ...section, count };
    });
    const weeks = getWeeklySegments(doc).map(segment => {
      const frag = doc.createElement('div'); segment.nodes.forEach(el => frag.appendChild(el.cloneNode(true)));
      const counts = {}; APP_CONFIG.weeklySections.forEach(s => counts[s.className] = frag.querySelectorAll(`.${s.className}`).length);
      const loose = Array.from(frag.children).filter(child =>
        !child.matches(APP_CONFIG.allowedWrapperSelectors) && !child.matches('h1, hr') && child.textContent.trim() !== ''
      );
      const misc = Array.from(frag.querySelectorAll('.misc')).filter(el => el.textContent.trim() !== '');
      let emptySectionCount = 0;
      APP_CONFIG.weeklySections.forEach(s => {
        const count = counts[s.className];
        if (count === 0) recommendations.push(`Week ${segment.number} is missing the ${s.name} section.`);
        if (count > 1) recommendations.push(`Week ${segment.number} has ${count} ${s.name} sections; expected one.`);
        if (count === 1) {
          const wrapper = frag.querySelector(`.${s.className}`);
          if (wrapper && !wrapper.textContent.trim()) {
            emptySectionCount++;
            recommendations.push(`Week ${segment.number} has an empty ${s.name} section.`);
          }
        }
      });
      const reviewCount = loose.length + misc.length + emptySectionCount + frag.querySelectorAll('[data-cdt-needs-review="true"]').length;
      if (loose.length) recommendations.push(`Week ${segment.number} has ${loose.length} item${loose.length === 1 ? '' : 's'} outside a recognized section.`);
      if (misc.length) recommendations.push(`Week ${segment.number} has Other Content that still needs review.`);
      return { ...segment, counts, looseCount: loose.length, miscCount: misc.length, reviewCount, fixes: getElementFixes(frag) };
    });

    if (!weeks.length) recommendations.push('No weekly topic headings were found.');
    const seenNumbers = new Set();
    weeks.forEach(week => {
      if (seenNumbers.has(week.number)) recommendations.push(`Week ${week.number} appears more than once.`);
      seenNumbers.add(week.number);
    });
    if (weeks.length) {
      const maxWeek = Math.max(...weeks.map(week => week.number));
      for (let expected = 1; expected <= maxWeek; expected++) {
        if (!seenNumbers.has(expected)) recommendations.push(`Week ${expected} heading is missing.`);
      }
    }
    const fixes = getElementFixes(doc.body);
    return { doc, topLevel, weeks, recommendations: [...new Set(recommendations)], fixes, issuesFound: recommendations.length > 0 };
  }

  function generateSectionPresenceReport(htmlContent) {
    const analysis = analyzeDocumentStructure(htmlContent);
    const { topLevel, weeks, recommendations, fixes, issuesFound } = analysis;

    let html = '<h4>Document Structure Report</h4>';
    if (issuesFound) {
      html += '<p style="color:#dc3545;font-weight:700">Some content still needs review.</p>';
    } else if (fixes.length) {
      html += '<p style="color:#2563eb;font-weight:700">Ready. Formatting variations were corrected automatically.</p>';
    } else {
      html += `<p style="color:#28a745;font-weight:700">Document structure looks good!</p>`;
    }

    html += '<h5>Top-Level Sections</h5><table><thead><tr><th>Section</th><th>Present?</th><th>Count</th></tr></thead><tbody>';
    topLevel.forEach(s => {
      html += `<tr><td>${escapeHtml(s.name)}</td><td>${s.count>0?'Yes':'No'}</td><td>${s.count}</td></tr>`;
    });
    html += '</tbody></table>';

    html += '<h5 style="margin-top:10px">Weekly Sections</h5>';
    html += '<table><thead><tr><th>Week</th><th>LO</th><th>A&R</th><th>Assignments</th><th>Review</th><th>Status</th></tr></thead><tbody>';
    weeks.forEach(d => {
      const missing = APP_CONFIG.weeklySections.some(section => d.counts[section.className] !== 1);
      const status = missing || d.reviewCount ? 'Needs review' : (d.fixes.length ? 'Auto-corrected' : 'Ready');
      html += `<tr><td>Week ${d.number}: ${escapeHtml(d.title)}</td><td>${d.counts.LO}</td><td>${d.counts.activities}</td><td>${d.counts.assignments}</td><td>${d.reviewCount}</td><td>${status}</td></tr>`;
    });
    html += '</tbody></table>';
    if (fixes.length) {
      html += '<h5 style="margin-top:10px">Automatic Corrections</h5><ul>';
      fixes.forEach(fix => { html += `<li>${escapeHtml(fix)}</li>`; });
      html += '</ul>';
    }
    if (issuesFound) {
      html += '<h5 style="margin-top:10px">Recommendations</h5><ul>';
      recommendations.forEach(r => html += `<li>${escapeHtml(r)}</li>`);
      html += '</ul>';
    }
    return html;
  }

  // Gentle cleanup: keep unclassified content inside its original week.
  const CLEANUP_OPTIONS = { mode: 'wrap' }; // 'wrap' or 'delete'

  function removeExtraContent(htmlContent) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlContent, 'text/html');
    cleanupHtml(doc);
    getWeeklySegments(doc).forEach(segment => {
      const loose = segment.nodes.filter(node =>
        !node.matches(APP_CONFIG.allowedWrapperSelectors) && !node.matches('h1, hr') && node.textContent.trim() !== ''
      );
      if (!loose.length) return;
      if (CLEANUP_OPTIONS.mode === 'delete') { loose.forEach(node => node.remove()); return; }
      let bucket = segment.nodes.find(node => node.matches && node.matches('.misc'));
      if (!bucket) {
        bucket = doc.createElement('div'); bucket.className = 'misc';
        bucket.setAttribute('data-cdt-section', 'misc');
        const h2 = doc.createElement('h2'); h2.textContent = `Other Content — Week ${segment.number}`; bucket.appendChild(h2);
        loose[0].before(bucket);
      }
      loose.forEach(node => bucket.appendChild(node));
      addElementFix(bucket, `Week ${segment.number}: kept unclassified content with its original week in Other Content.`);
    });
    return doc.body.innerHTML;
  }

  function runStructureSelfTests() {
    const results = [];
    const check = (name, condition, detail = '') => {
      if (!condition) throw new Error(`Structure test failed: ${name}${detail ? ` - ${detail}` : ''}`);
      results.push({ test: name, status: 'passed' });
    };
    const fixture = `
      <h1>Course Description</h1><p>Description</p>
      <h1>Program Learning Outcomes</h1><p>PLO</p>
      <h1>Course Learning Outcomes</h1><p>CLO</p>
      <h1>Required Course Materials</h1><p>Materials</p>
      <p>Week 1: Test Week</p><p>&nbsp;</p>
      <p>Learning Objectives</p>
      <table><tr><td><p>1.1 A plain-paragraph objective.</p></td><td>CLO1</td></tr></table>
      <h1>Activities &amp; Resources</h1>
      <table><tr><td>Reading</td><td>Week 1</td></tr><tr><td><p>Read chapter 1.</p></td><td></td></tr></table>
      <table><tr><td>Lecture</td><td>Week 1</td></tr><tr><td><p>Review lecture 1.</p></td><td></td></tr></table>
      <p>Assignment</p>
      <table><tr><td>Discussion</td><td>1.1</td></tr><tr><td><p>Post a response.</p><p><strong>Day 7 of</strong> <strong>Week</strong> <strong>2</strong>.</p><p><strong>Day 7 of</strong><strong>Week</strong><strong>3</strong>.</p></td><td></td></tr></table>`;
    const processed = processHtml(fixture);
    const analysis = analyzeDocumentStructure(processed);
    check('wrong-style week heading is recognized', analysis.weeks.length === 1 && analysis.weeks[0].number === 1);
    check('singular Assignment is normalized', analysis.weeks[0].counts.assignments === 1);
    check('plain-paragraph objectives are extracted', analysis.weeks[0].counts.LO === 1 && /plain-paragraph objective/i.test(processed));
    check('multiple A&R tables remain one semantic section', analysis.weeks[0].counts.activities === 1);
    check('safe corrections are recorded', analysis.fixes.length >= 3);
    check('processing is idempotent', processHtml(processed) === processed);
    const processedDoc = new DOMParser().parseFromString(processed, 'text/html');
    const deadlines = Array.from(processedDoc.querySelectorAll('.assignments strong')).map(element => element.textContent.trim());
    check('bold deadline runs keep and repair spaces', deadlines.includes('Day 7 of Week 2') && deadlines.includes('Day 7 of Week 3') && !/ofWeek\d/i.test(processed));

    const listFixture = fixture.replace(
      '<table><tr><td><p>1.1 A plain-paragraph objective.</p></td><td>CLO1</td></tr></table>',
      '<table><tr><td><ol><li>First objective.</li><li>Second objective.</li></ol></td><td>CLO1</td></tr></table>'
    );
    const listProcessed = processHtml(listFixture);
    const listDoc = new DOMParser().parseFromString(listProcessed, 'text/html');
    const listWrapper = listDoc.querySelector('.LO');
    check('list objectives become separate paragraphs', !!listWrapper && listWrapper.querySelectorAll(':scope > p').length === 2 && !listWrapper.querySelector('ol, ul'));
    const duplicateList = listDoc.createElement('ol');
    duplicateList.innerHTML = '<li>First objective.</li><li>Second objective.</li>';
    listWrapper.prepend(duplicateList);
    const repairedDoc = new DOMParser().parseFromString(processHtml(listDoc.body.innerHTML), 'text/html');
    const repairedWrapper = repairedDoc.querySelector('.LO');
    check('existing list-plus-paragraph duplicates are repaired', !!repairedWrapper && repairedWrapper.querySelectorAll(':scope > p').length === 2 && !repairedWrapper.querySelector('ol, ul'));

    const nestedListHtml = restoreWordListStructure(
      '<ol><li>First numbered item</li><li>Parent numbered item:</li><li>First lettered item</li><li>Second lettered item</li></ol>',
      [
        { matchText: normalizeListMatchText('First numbered item'), numId: '10', level: 0, format: 'decimal', start: 1, left: 720 },
        { matchText: normalizeListMatchText('Parent numbered item:'), numId: '10', level: 0, format: 'decimal', start: 1, left: 720 },
        { matchText: normalizeListMatchText('First lettered item'), numId: '11', level: 0, format: 'lowerLetter', start: 1, left: 1080 },
        { matchText: normalizeListMatchText('Second lettered item'), numId: '11', level: 0, format: 'lowerLetter', start: 1, left: 1080 }
      ]
    );
    const nestedListDoc = new DOMParser().parseFromString(nestedListHtml, 'text/html');
    const nestedOuter = nestedListDoc.querySelector('body > ol');
    const nestedChildren = nestedOuter ? nestedOuter.querySelectorAll(':scope > li') : [];
    const nestedLetters = nestedChildren[1]?.querySelector(':scope > ol[type="a"]');
    check('Word number and letter list hierarchy is restored', nestedChildren.length === 2 && nestedLetters?.querySelectorAll(':scope > li').length === 2);

    const falseBulletHtml = restoreWordListStructure(
      '<ul><li>Intro instruction</li><li>Case directions:<ul><li><ul><li>Case narrative</li><li>Real question</li></ul></li></ul></li><li>Closing instruction</li></ul>',
      [
        { matchText: normalizeListMatchText('Intro instruction'), isList: false },
        { matchText: normalizeListMatchText('Case directions:'), isList: false },
        { matchText: normalizeListMatchText('Case narrative'), isList: false },
        { matchText: normalizeListMatchText('Real question'), isList: true, numId: '20', level: 0, format: 'bullet', left: 720 },
        { matchText: normalizeListMatchText('Closing instruction'), isList: false }
      ]
    );
    const falseBulletDoc = new DOMParser().parseFromString(falseBulletHtml, 'text/html');
    const falseBulletTags = Array.from(falseBulletDoc.body.children).map(element => element.tagName).join(',');
    check('explicit Word non-list paragraphs are unwrapped from false bullets', falseBulletTags === 'P,P,P,UL,P' && falseBulletDoc.querySelectorAll('li').length === 1 && /Real question/.test(falseBulletDoc.querySelector('li')?.textContent || ''));

    const missingDoc = new DOMParser().parseFromString(processed, 'text/html');
    missingDoc.querySelector('.assignments')?.remove();
    const missing = analyzeDocumentStructure(missingDoc.body.innerHTML);
    check('true missing sections remain review items', missing.recommendations.some(item => /missing the Assignments section/i.test(item)));

    const looseFixture = fixture.replace('<p>Learning Objectives</p>', '<p>Loose weekly note</p><p>Learning Objectives</p>');
    const looseProcessed = processHtml(looseFixture);
    check('loose weekly content is detected', analyzeDocumentStructure(looseProcessed).weeks[0].looseCount === 1);
    const cleaned = removeExtraContent(looseProcessed);
    const cleanedDoc = new DOMParser().parseFromString(cleaned, 'text/html');
    const localBucket = cleanedDoc.querySelector('h1.wk-topic ~ .misc');
    check('cleanup keeps loose content in the same week', !!localBucket && /Loose weekly note/.test(localBucket.textContent));

    console.table(results);
    return { passed: results.length, results };
  }

  Object.assign(window.CanvasDocxTools, {
    processHtml,
    extractDocxListMetadata,
    restoreWordListStructure,
    analyzeDocumentStructure,
    generateSectionPresenceReport,
    removeExtraContent,
    runStructureSelfTests
  });

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
.misc { border-color:#64748b; background:#f1f5f9; padding: 12px 14px; border-left: 6px solid; margin: 12px 0; }
#wk-topics-list { border:1px dashed #94a3b8; padding:10px 12px; background:#f8fafc; }
#wk-topics-list h2 { margin: 0 0 6px; }
#wk-topics-list ul { margin:0; padding-left: 18px; }
#wk-topics-list li { margin:2px 0; }
h1.wk-topic { color:#1f2937; border-bottom:2px solid #e5e7eb; padding-bottom:4px; }
.cdes > h2, .plo > h2, .clo > h2, .material > h2, .LO > h2, .activities > h2, .assignments > h2, .misc > h2 { margin-top: 0; }
`;
  const PREVIEW_STYLE_PARENT_ID = 'cdt-preview-styles';
  const PREVIEW_STYLE_IFRAME_ID = 'cdt-preview-styles-editor';
  const STATE = { previewOn: true, iframeDnDAttached: false };

  function setPreviewStyles(on) {
    let s = document.getElementById(PREVIEW_STYLE_PARENT_ID);
    if (on && !s) { s = document.createElement('style'); s.id = PREVIEW_STYLE_PARENT_ID; s.textContent = SECTION_STYLE_CSS; document.head.appendChild(s); }
    else if (!on && s) { s.remove(); }

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
      ['assignments','Assignments','#ef4444'],
      ['misc','Other Content','#64748b']
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
  const selWrapClass       = document.getElementById('cdt-wrap-class');
  const btnWrap            = document.getElementById('cdt-wrap');
  const btnFindReplace     = document.getElementById('cdt-find-replace');

  // Populate wrapper select
  const wrapOptions = [
    ...APP_CONFIG.topLevelSections.map(s => s.className),
    ...APP_CONFIG.weeklySections.map(s => s.className),
    'misc'
  ];
  selWrapClass.innerHTML = wrapOptions.map(cls => `<option value="${cls}">${cls}</option>`).join('');

  let lastPreCleanupHtml = null;

  async function convertAndReplace(fileOpt) {
    await waitForEditor();
    await ensureMammoth();
    try { await ensureJSZip(); }
    catch (error) { console.warn('JSZip could not be loaded; Word list repair will be skipped.', error); }

    const file = fileOpt || (fileInput && fileInput.files && fileInput.files[0]);
    if (!file) { alert('Choose a .docx file first, or drop one on the bar.'); return; }
    const isDocxByName = /\.docx$/i.test(file.name || '');
    const isDocxByType = /officedocument\.wordprocessingml\.document/.test(file.type || '');
    if (!(isDocxByName || isDocxByType)) { alert('Please provide a .docx file.'); return; }

    const arrayBuffer = await file.arrayBuffer();
    let listMetadata = [];
    try {
      listMetadata = await extractDocxListMetadata(arrayBuffer.slice(0));
    } catch (error) {
      console.warn('Word list metadata could not be read; continuing with Mammoth list conversion.', error);
    }
    const result = await mammoth.convertToHtml({ arrayBuffer }, { styleMap: APP_CONFIG.styleMap });

    if (Array.isArray(result.messages) && result.messages.length) {
      const warn = result.messages.map(m => `• ${m.message || m.value || String(m)}`).join('\n');
      console.warn('Mammoth messages:\n' + warn);
      showToast('Converted with warnings. Check console.');
    }

    const sourceAwareHtml = restoreWordListStructure(result.value, listMetadata);
    const processed = processHtml(sourceAwareHtml);
    setRCEHtml(processed);
    if (fileInput) fileInput.value = '';
    showToast('Converted and inserted into the RCE.');
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

    if (CanvasCtx.courseId && CanvasCtx.pageSlug) {
      const url = `${CanvasCtx.apiBase}/courses/${CanvasCtx.courseId}/pages/${CanvasCtx.pageSlug}`;
      const res = await fetch(url, {
        method: 'PUT',
        headers: {
          'Content-Type':'application/json',
          'Accept': 'application/json',
          'X-CSRF-Token': getCsrfToken()
        },
        body: JSON.stringify({ wiki_page: { body: html, ...(title ? { title } : {}) } })
      });
      if (!res.ok) {
        const t = await res.text().catch(()=> '');
        throw new Error(`Save failed ${res.status}. ${t.slice(0,240)}`);
      }
      showToast('Saved to page via API.');
    } else {
      const btn = [...document.querySelectorAll('button, input[type=submit]')].find(b => /save/i.test(b.textContent || b.value || ''));
      if (btn) { btn.click(); showToast('Triggered Canvas Save button.'); }
      else { alert('Could not determine page context to save.'); }
    }
  });

  // ---- Wrap current selection with class ----
  function wrapSelectionWithClass(cls) {
    const ed = getEditor();
    if (!ed) { alert('Editor not ready.'); return; }
    const html = ed.selection && ed.selection.getContent({ format: 'html' });
    if (!html || !html.trim()) {
      alert('Select some content in the editor first.');
      return;
    }
    const wrapped = `<div class="${cls}">${html}</div>`;
    ed.selection.setContent(wrapped, { format: 'html' });
    setPreviewStyles(STATE.previewOn);
    showToast(`Wrapped selection in .${cls}`);
  }

  // ---- Drag & Drop wiring ----
  function preventDefaults(e){ e.preventDefault(); e.stopPropagation(); }
  function highlight(){ root.classList.add('drop-active'); }
  function unhighlight(){ root.classList.remove('drop-active'); }

  function firstDocxFileFromDataTransfer(dt) {
    if (!dt) return null;
    if (dt.items && dt.items.length) {
      for (const it of dt.items) {
        if (it.kind === 'file') {
          const f = it.getAsFile();
          if (f && /\.docx$/i.test(f.name)) return f;
        }
      }
    }
    return [...(dt.files || [])].find(f => /\.docx$/i.test(f.name)) || null;
  }

  function handleDrop(e){
    preventDefaults(e); unhighlight();
    const file = firstDocxFileFromDataTransfer(e.dataTransfer);
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
      const file = firstDocxFileFromDataTransfer(e.dataTransfer);
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
      <p>This will place content that is not in a standard section like
      "Learning Objectives", "Activities and Resources", or "Assignments" into an "Other Content" bucket for its current week.</p>
      <p>Nothing will be moved to another week. Cleanup is idempotent and can be undone.</p>
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

  btnWrap.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); wrapSelectionWithClass(selWrapClass.value); });

  // ---- Find/Replace UI ----
  function openFindReplaceModal() {
    document.getElementById('cdt-modal-title').textContent = 'Find and Replace (HTML)';
    document.getElementById('cdt-modal-body').innerHTML = `
      <form id="cdt-fr-form">
        <div class="full">
          <label for="cdt-fr-find">Find</label>
          <input type="text" id="cdt-fr-find" placeholder="text or regex">
        </div>
        <div class="full">
          <label for="cdt-fr-replace">Replace with</label>
          <input type="text" id="cdt-fr-replace" placeholder="replacement">
        </div>
        <div class="full opts">
          <label><input type="checkbox" id="cdt-fr-case"> Case sensitive</label>
          <label><input type="checkbox" id="cdt-fr-word"> Whole word</label>
          <label><input type="checkbox" id="cdt-fr-regex"> Regex</label>
          <span id="cdt-fr-count" style="margin-left:auto;color:#4a5568"></span>
        </div>
        <div id="cdt-fr-actions" class="full">
          <button class="btn" type="button" id="cdt-fr-replace-first">Replace First</button>
          <button class="btn" type="button" id="cdt-fr-replace-all">Replace All</button>
        </div>
      </form>
    `;
    modal.style.display = 'flex';

    const $find = document.getElementById('cdt-fr-find');
    const $rep  = document.getElementById('cdt-fr-replace');
    const $case = document.getElementById('cdt-fr-case');
    const $word = document.getElementById('cdt-fr-word');
    const $regex= document.getElementById('cdt-fr-regex');
    const $count= document.getElementById('cdt-fr-count');

    function escapeRegex(s){ return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
    function buildRe() {
      let pattern = $find.value || '';
      if (!pattern) return null;
      if (!$regex.checked) {
        pattern = escapeRegex(pattern);
        if ($word.checked) pattern = `\\b${pattern}\\b`;
      } else {
        if ($word.checked) pattern = `\\b(?:${pattern})\\b`;
      }
      const flags = $case.checked ? 'g' : 'gi';
      try { return new RegExp(pattern, flags); } catch(e) { return null; }
    }
    function updateCount() {
      const html = getRCEHtml() || '';
      const re = buildRe();
      if (!re) { $count.textContent = ''; return; }
      const matches = html.match(re);
      $count.textContent = matches ? `${matches.length} match${matches.length===1?'':'es'}` : '0 matches';
    }
    $find.addEventListener('input', updateCount);
    $rep.addEventListener('input', updateCount);
    $case.addEventListener('change', updateCount);
    $word.addEventListener('change', updateCount);
    $regex.addEventListener('change', updateCount);
    setTimeout(()=>{ $find.focus(); updateCount(); }, 0);

    function replaceImpl(all=false){
      const html = getRCEHtml() || '';
      const re = buildRe();
      if (!re) { alert('Enter a valid find pattern.'); return; }
      const replacement = $rep.value ?? '';
      if (!re.test(html)) { showToast('No matches.'); return; }
      const newHtml = all ? html.replace(re, replacement) : html.replace(re, replacement);
      setRCEHtml(newHtml);
      const count = (html.match(re) || []).length;
      if (all) showToast(`Replaced ${count} occurrence${count===1?'':'s'}.`);
      else showToast(`Replaced 1 occurrence.`);
      updateCount();
    }

    document.getElementById('cdt-fr-replace-first').onclick = () => replaceImpl(false);
    document.getElementById('cdt-fr-replace-all').onclick   = () => replaceImpl(true);
  }

  btnFindReplace.addEventListener('click', (e)=>{ e.preventDefault(); e.stopPropagation(); openFindReplaceModal(); });

  // Keyboard shortcuts
  document.addEventListener('keydown', e => {
    if (!e.ctrlKey || !e.altKey) return;
    const k = e.key.toLowerCase();
    if (k === 's') { e.preventDefault(); btnSave.click(); }
    if (k === 'r') { e.preventDefault(); btnReport.click(); }
    if (k === 'p') { e.preventDefault(); btnToggleStyles.click(); }
    if (k === 'w') { e.preventDefault(); btnWrap.click(); }
    if (k === 'f') { e.preventDefault(); btnFindReplace.click(); }
  });

  // Start with preview styles ON
  setPreviewStyles(true);

  // Make sure editor exists before first use
  waitForEditor();

  console.log(`Canvas DOCX → RCE Toolbelt loaded (v${TOOLBELT_VERSION}).`);
}
// === end =====================================================================
