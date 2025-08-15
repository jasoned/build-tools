(async () => {
  /* ========= Canvas helpers ========= */
  function getCourseIdFromUrl() {
    const m = location.pathname.match(/\/courses\/(\d+)/);
    if (!m) throw new Error("Cannot find course ID in URL");
    return m[1];
  }
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
  const API = {
    base: `${location.origin}/api/v1`,
    async get(url) { const r = await fetch(url, { credentials: "same-origin" }); if (!r.ok) throw new Error(`GET ${url} failed: ${r.status}`); return r; },
    async getJSON(url) { const r = await API.get(url); return r.json(); },
    parseLinkHeader(link) { if (!link) return {}; const parts = link.split(','); const rels = {}; for (const p of parts) { const m = p.match(/<([^>]+)>;\s*rel="([^"]+)"/); if (m) rels[m[2]] = m[1]; } return rels; },
    async getAll(path) {
      let url = `${API.base}${path}${path.includes('?') ? '&' : '?'}per_page=100`;
      const out = [];
      while (url) { const r = await API.get(url); const data = await r.json(); out.push(...data); const links = API.parseLinkHeader(r.headers.get('Link')); url = links.next || null; }
      return out;
    },
    async put(path, body) {
      const url = `${API.base}${path}`;
      const r = await fetch(url, { method: "PUT", credentials: "same-origin", headers: { "Content-Type": "application/json", "X-CSRF-Token": getCsrfToken() }, body: JSON.stringify(body) });
      if (!r.ok) { const t = await r.text().catch(()=> ""); throw new Error(`PUT ${url} failed: ${r.status} ${t.slice(0,200)}`); }
      return r.json();
    }
  };

  /* ========= UI ========= */
  function injectStyles() {
    const css = `
#wkx-overlay{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:99999;display:flex;align-items:center;justify-content:center}
#wkx-modal{background:#fff;max-width:1100px;width:95%;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.2);overflow:hidden;font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial}
#wkx-header{padding:12px 16px;border-bottom:1px solid #eee;display:flex;justify-content:space-between;align-items:center}
#wkx-body{padding:14px 16px;max-height:72vh;overflow:auto}
#wkx-footer{padding:12px 16px;border-top:1px solid #eee;display:flex;gap:8px;justify-content:flex-end;align-items:center}
#wkx-close{background:none;border:none;font-size:20px;cursor:pointer}
#wkx-spinner{display:none;align-items:center;gap:8px}
#wkx-spinner.on{display:flex}
small.mono{font-family:ui-monospace, Menlo, Consolas, monospace;color:#666}
.wkx-topopts{display:flex;gap:16px;align-items:center;flex-wrap:wrap;margin:6px 0 10px}
.wkx-topopts input[type="text"]{padding:6px;border:1px solid #ddd;border-radius:8px}
.wkx-week{margin:10px 0 0;border:1px solid #eee;border-radius:10px}
.wkx-week > .head{padding:8px 10px;background:#fafafa;border-bottom:1px solid #eee;display:flex;justify-content:space-between;align-items:center}
.wkx-week .rows{padding:8px 10px}
.wkx-row{display:grid;grid-template-columns:26px 1.25fr 1.35fr 120px;gap:10px;align-items:flex-start;border-top:1px solid #f3f3f3;padding:8px 0}
.wkx-row:first-child{border-top:none}
.wkx-title{font-weight:600}
.badge{display:inline-flex;align-items:center;justify-content:center;border:1px solid transparent;border-radius:999px;padding:2px 8px;font-size:12px;min-width:52px;text-align:center;font-weight:700}
.badge.ok{background:#0f9d58;color:#fff;border-color:#0f9d58}
.badge.warn{background:#fbbc04;color:#111;border-color:#fbbc04}
.badge.err{background:#d93025;color:#fff;border-color:#d93025}
.wkx-note{font-size:12px;color:#555;margin-top:2px}
.wkx-matchbox{display:flex;flex-direction:column;gap:6px}
.wkx-override{display:none;border:1px solid #e6e6e6;border-radius:8px;padding:8px;background:#fcfcfc}
.wkx-override.on{display:block}
.wkx-override .row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.wkx-override select{max-width:460px}
    `;
    const s = document.createElement("style"); s.textContent = css; document.head.appendChild(s);
  }
  function overlay(html) {
    const wrap = document.createElement("div"); wrap.id = "wkx-overlay"; wrap.innerHTML = `
      <div id="wkx-modal">
        <div id="wkx-header">
          <div><strong>Weekly Assignment Summaries → Assignment Descriptions</strong></div>
          <button id="wkx-close" aria-label="Close">×</button>
        </div>
        <div id="wkx-body">${html}</div>
        <div id="wkx-footer">
          <div id="wkx-spinner">
            <svg width="18" height="18" viewBox="0 0 100 100"><circle cx="50" cy="50" r="35" fill="none" stroke="#555" stroke-width="10" stroke-dasharray="164.9 56.98"><animateTransform attributeName="transform" type="rotate" dur="1s" repeatCount="indefinite" values="0 50 50;360 50 50"/></circle></svg>
            <span id="wkx-progress">Working…</span>
          </div>
          <button class="Button Button--secondary" id="wkx-cancel">Cancel</button>
          <button class="Button Button--primary" id="wkx-run">Start</button>
        </div>
      </div>`;
    document.body.appendChild(wrap);
    wrap.querySelector('#wkx-close').onclick = () => wrap.remove();
    wrap.querySelector('#wkx-cancel').onclick = () => wrap.remove();
    return wrap;
  }
  function setSpin(on, text) { const sp = document.getElementById('wkx-spinner'); const t = document.getElementById('wkx-progress'); if (text) t.textContent = text; sp.classList.toggle('on', !!on); }
  function done(message, html) { const body = document.getElementById('wkx-body'); body.innerHTML = `<div class="ic-flash-success">✅ ${message}</div>${html||''}`; const run = document.getElementById('wkx-run'); if (run) run.remove(); }
  async function copyHtmlToClipboard(html, plain) {
    try { if (navigator.clipboard && window.ClipboardItem) { const data = { 'text/html': new Blob([html],{type:'text/html'}), 'text/plain': new Blob([plain],{type:'text/plain'}) }; await navigator.clipboard.write([new ClipboardItem(data)]); alert('Copied summary to clipboard.'); } else { await navigator.clipboard.writeText(plain); alert('Copied plain text summary to clipboard.'); } } catch { await navigator.clipboard.writeText(plain); alert('Copied plain text summary to clipboard.'); }
  }

  /* ========= String + matching ========= */
  const WRE = /\bweek\s*(\d+)\b/i;
  const TYPE_PREFIXES = [ 'discussion','forum','assignment','assessment','quiz','exam','test','activity','project','paper','essay','task','lab','homework','worksheet','presentation' ];
  const DASH_CLASS = '\\-‐‑‒–—−'; // common dash variants
  // Remove leading week + type prefixes robustly, then any stray leading punctuation
  function stripWeekPrefix(title, week){ return title.replace(new RegExp(`^\\s*week\\s*${week}\\s*[:：${DASH_CLASS}]?\\s*`,'i'),'').trim(); }
  function removeTypePrefix(t){
    let out = t || '';
    for (let k=0;k<3;k++){
      const re = new RegExp(`^\\s*(?:\\[\\s*(?:${TYPE_PREFIXES.join('|')})\\s*\\]|(?:${TYPE_PREFIXES.join('|')}))\\s*(?:[:：[${DASH_CLASS}]]\\s*)?`,'i');
      const before = out; out = out.replace(re,''); if (out === before) break;
    }
    // Clean any leftover leading punctuation/quotes
    out = out.replace(new RegExp(`^\\s*[:：;,\.\\-—–${DASH_CLASS}]+\\s*`,'i'), '');
    out = out.replace(/^["'“”‘’]+\s*/, '');
    // Drop a leading article if present (helps with titles like "An / The …")
    out = out.replace(/^(?:a|an|the)\s+/i, '');
    return out.trim();
  }
  function coreTitle(s, week){ let t = s || ''; if (typeof week === 'number') t = stripWeekPrefix(t, week); return removeTypePrefix(t); }
  function normCore(s){ return (s||'').toLowerCase().replace(/&amp;/g,' and ').replace(/&/g,' and ').replace(/[“”„"]/g,'"').replace(/[’‘']/g,"'").replace(/[^a-z0-9'":\- ]+/g,' ').replace(/\s+/g,' ').trim(); }
  function tokenKeyCore(s){ const t = normCore(s).split(' ').filter(Boolean); const uniq=[...new Set(t)].sort(); return uniq.join(' '); }
  function tokenStats(a,b){
    const A = new Set((a||'').split(' ').filter(Boolean));
    const B = new Set((b||'').split(' ').filter(Boolean));
    let inter = 0; A.forEach(x => { if (B.has(x)) inter++; });
    const containment = inter / Math.max(A.size || 1, B.size || 1);
    const jaccard = inter / (new Set([...A, ...B]).size || 1);
    return { containment, jaccard, aSize:A.size, bSize:B.size, inter };
  }
  function jaroWinkler(a,b){ a=a||''; b=b||''; if (a===b) return 1; const mtp=matches(a,b); const m=mtp[0]; if(!m) return 0; const j=((m/a.length)+(m/b.length)+((m-mtp[1])/m))/3; const p=.1; const l=Math.min(mtp[2],4); return j + l*p*(1-j); function matches(s1,s2){ const maxDist=Math.floor(Math.max(s1.length,s2.length)/2)-1; const s1M=new Array(s1.length); const s2M=new Array(s2.length); let matches=0, trans=0; for(let i=0;i<s1.length;i++){ const start=Math.max(0,i-maxDist); const end=Math.min(i+maxDist+1,s2.length); for(let j=start;j<end;j++){ if(s2M[j]) continue; if(s1[i]!==s2[j]) continue; s1M[i]=true; s2M[j]=true; matches++; break; } } if(!matches) return [0,0,0]; let k=0; for(let i=0;i<s1.length;i++){ if(!s1M[i]) continue; while(!s2M[k]) k++; if(s1[i]!==s2[k]) trans++; k++; } let prefix=0; for(let i=0;i<Math.min(4,s1.length,s2.length)&&s1[i]===s2[i];i++) prefix++; return [matches, Math.floor(trans/2), prefix]; } }

  /* ========= Parse summary page ========= */
  function parseSectionsWithRaw(bodyHtml, pageWeek){ const div=document.createElement('div'); div.innerHTML=bodyHtml||''; const heads=Array.from(div.querySelectorAll('.assnt-section h2, .assnt-section h3, h2, h3')); if(!heads.length) return []; const sections=[]; for (let i=0;i<heads.length;i++){ const h=heads[i]; const titleRaw=h.textContent.trim(); const wm=titleRaw.match(WRE); if (wm && parseInt(wm[1],10)!==pageWeek) continue; const frag=[]; for (let n=h.nextSibling;n;n=n.nextSibling){ if(n.nodeType===1 && /^(H2|H3)$/.test(n.tagName)) break; frag.push(n.cloneNode(true)); } const wrap=document.createElement('div'); frag.forEach(n=>wrap.appendChild(n)); const titleClean = stripWeekPrefix(titleRaw, pageWeek) || titleRaw; sections.push({ titleRaw, titleClean, html: wrap.innerHTML.trim() }); } return sections; }

  /* ========= Insert or update link line under a section ========= */
  function upsertLinkAfterSection(pageHtml, sectionHeadingRaw, assignment){ const wrap=document.createElement('div'); wrap.innerHTML=pageHtml||''; const heads=Array.from(wrap.querySelectorAll('.assnt-section h2, .assnt-section h3, h2, h3')); const h=heads.find(x=>x.textContent.trim()===sectionHeadingRaw.trim()); if(!h) return { html: pageHtml, changed:false, reason:'heading not found' }; const sectionDiv=h.closest('.assnt-section')||h.parentNode; let cursor=sectionDiv.nextSibling; let linkP=null, hrNode=null; while(cursor){ if(cursor.nodeType===1 && cursor.matches('.assnt-section')) break; if(cursor.nodeType===1 && cursor.matches('p.wkx-assignment-link')){ linkP=cursor; break; } if(!hrNode && cursor.nodeType===1 && cursor.matches('hr')) hrNode=cursor; cursor=cursor.nextSibling; } const phrase=document.getElementById('opt-link-phrase')?.value || 'to access it.'; const htmlLine=`<strong>Click</strong> <a href="${assignment.url}">${assignment.name}</a> ${phrase}`; if(!linkP){ linkP=document.createElement('p'); linkP.className='wkx-assignment-link'; linkP.setAttribute('data-assignment-id',assignment.id); linkP.innerHTML=htmlLine; if(hrNode){ sectionDiv.parentNode.insertBefore(linkP, hrNode); } else { sectionDiv.parentNode.insertBefore(linkP, sectionDiv.nextSibling); } return { html: wrap.innerHTML, changed:true }; } else { linkP.setAttribute('data-assignment-id',assignment.id); linkP.innerHTML=htmlLine; return { html: wrap.innerHTML, changed:true }; } }

  /* ========= Main flow ========= */
  injectStyles();
  const ui = overlay(`
    <div class="wkx-topopts">
      <label style="display:inline-flex;gap:6px;align-items:center;cursor:pointer"><input type="checkbox" id="opt-append"> Append instead of replace</label>
      <label style="display:inline-flex;gap:6px;align-items:center;cursor:pointer"><input type="checkbox" id="opt-skip-nonempty"> Skip if assignment already has content</label>
      <label style="display:inline-flex;gap:6px;align-items:center;cursor:pointer"><input type="checkbox" id="opt-unique" checked> Enforce one‑to‑one matching per week</label>
      <label style="display:inline-flex;gap:6px;align-items:center;cursor:pointer"><input type="checkbox" id="opt-insert-links" checked> Also add link lines on weekly pages</label>
      <label style="display:inline-flex;gap:6px;align-items:center;cursor:pointer">Link phrase: <input type="text" id="opt-link-phrase" value="to access it." size="26"></label>
      <label style="display:inline-flex;gap:6px;align-items:center;cursor:pointer"><input type="checkbox" id="opt-links-selected-only" checked> Add links only for selected rows</label>
      <small class="mono">Matching ignores type prefixes (Discussion/Quiz/Assignment), leading punctuation, and articles (a/an/the) before scoring.</small>
    </div>
    <div id="wkx-preview" style="margin-top:6px"></div>
  `);
  const byId = id => document.getElementById(id);
  const delay = ms => new Promise(res => setTimeout(res, ms));
  byId('opt-unique').addEventListener('change', () => loadPreview());

  let preview = null; const courseId = getCourseIdFromUrl(); setTimeout(loadPreview, 0);

  async function loadPreview(){
    setSpin(true,'Fetching pages…');
    const allPages = await API.getAll(`/courses/${courseId}/pages`);
    const weekly = allPages.map(p=>{ const m=p.title && p.title.match(/^week\s*(\d+)\s*:\s*assignments/i); return m ? { week:parseInt(m[1],10), url:p.url, title:p.title } : null; }).filter(Boolean).sort((a,b)=>a.week-b.week);
    if(!weekly.length){ setSpin(false); byId('wkx-preview').innerHTML = `<div class="ic-flash-error">No pages titled "Week X: Assignments" found.</div>`; return; }

    setSpin(true,'Fetching assignments…');
    const assignments = await API.getAll(`/courses/${courseId}/assignments`);
    const rePrefix = /^week\s*(\d+)\s*-\s*(.+)$/i;
    const allAssn = []; const byWeek = new Map();
    for (const a of assignments){
      const m = a.name && a.name.match(rePrefix);
      const item = { id:a.id, name:a.name, html_url:a.html_url || `${location.origin}/courses/${courseId}/assignments/${a.id}`, raw:a };
      allAssn.push(item);
      if (m){ const wk=parseInt(m[1],10); const rest=(m[2]||'').trim(); const restCore = coreTitle(rest); const rec = { ...item, week:wk, rest, restCore, key: tokenKeyCore(restCore), norm: normCore(restCore) }; if(!byWeek.has(wk)) byWeek.set(wk,[]); byWeek.get(wk).push(rec); }
    }

    const pageBodies = new Map(); const items = []; let idxCounter=0;

    for (const pg of weekly){
      setSpin(true,`Fetching Week ${pg.week} page body…`);
      const page = await API.getJSON(`${API.base}/courses/${courseId}/pages/${encodeURIComponent(pg.url)}?include[]=body`);
      const bodyHtml = page.body || ''; pageBodies.set(pg.url, bodyHtml);
      const sections = parseSectionsWithRaw(bodyHtml, pg.week);
      const candidates = byWeek.get(pg.week) || [];
      for (const sec of sections){
        const titleCore = coreTitle(sec.titleClean, pg.week);
        const secKey = tokenKeyCore(titleCore); const secNorm = normCore(titleCore);
        const scored = candidates.map(c=>{
          const jw = jaroWinkler(secNorm, c.norm);
          const ts = secKey === c.key ? 1 : 0;
          const { containment } = tokenStats(secNorm, c.norm);
          let bonus = 0;
          if (containment >= 0.90) bonus = 0.98;
          else if (containment >= 0.80) bonus = 0.94;
          else if (containment >= 0.70) bonus = 0.90;
          const score = Math.max(jw, ts?0.98:0, jw*0.85 + (ts?0.15:0), bonus);
          return { c, score };
        }).sort((a,b)=>b.score-a.score);
        const best = scored[0] || { c:null, score:0 };
        const alts = scored.slice(0,3).map(x=>({ id:x.c?.id, name:x.c?.name, url:x.c?.html_url, score:x.score }));
        items.push({ idx: idxCounter++, week: pg.week, pageUrl: pg.url, srcPageTitle: pg.title, sectionHeadingRaw: sec.titleRaw, sectionTitle: sec.titleClean, titleCore, html: sec.html, match: best.c ? { id:best.c.id, name:best.c.name, url:best.c.html_url, rest:best.c.rest, restCore:best.c.restCore } : null, score: best.score, alts });
      }
    }

    // Ambiguity + one-to-one per week
    function postProcess(items){
      const byW = new Map(); items.forEach(it=>{ if(!byW.has(it.week)) byW.set(it.week, []); byW.get(it.week).push(it); });
      for (const [wk, arr] of byW){ arr.sort((a,b)=>b.score-a.score); const used = new Set(); for (const it of arr){ const rival = it.alts[1]?.score || 0; it.ambiguous = (it.score - rival) < 0.02 && rival >= 0.84; if (it.match){ if (used.has(it.match.id)) { it.conflict = true; } else { used.add(it.match.id); } } } }
    }
    postProcess(items);

    preview = { weekly, items, allAssn, byWeek, pageBodies };

    // Build preview UI
    const grouped = new Map(); for (const it of items){ if(!grouped.has(it.week)) grouped.set(it.week, []); grouped.get(it.week).push(it); }
    const weeksHtml = [...grouped.entries()].sort((a,b)=>a[0]-b[0]).map(([wk, arr])=>{
      const rows = arr.map(it=>{
        let cls = it.score >= 0.94 ? 'ok' : it.score >= 0.85 ? 'warn' : 'err';
        let note = '';
        if (it.conflict){ cls = 'err'; note = '<div class="wkx-note">Conflict: another section already claimed this assignment. Choose a different one.</div>'; }
        else if (it.ambiguous){ cls = 'warn'; const alt = it.alts[1]; note = `<div class=\"wkx-note\">Ambiguous: also close match \"${alt?.name || '—'}\" at ${Math.round((alt?.score||0)*100)}%.</div>`; }
        const matchHtml = it.match ? `<a href="${it.match.url}" target="_blank" rel="noopener">${it.match.name}</a><div><small class="mono">match on \"${it.match.restCore || it.match.rest}\"</small></div>` : `<span class="err">No same-week match</span>`;
        const defaultChecked = it.match && !it.conflict && it.score >= 0.85 && !it.ambiguous;
        const checkedAttr = defaultChecked ? 'checked' : '';
        const weekOpts = (preview.byWeek.get(wk) || []).map(a=>`<option value=\"${a.id}\">${a.name}</option>`).join('');
        const allOpts = preview.allAssn.map(a=>`<option value=\"${a.id}\">${a.name}</option>`).join('');
        return `
          <div class=\"wkx-row\" data-idx=\"${it.idx}\">\n            <div><input type=\"checkbox\" class=\"wkx-check\" ${checkedAttr}></div>\n            <div>\n              <div class=\"wkx-title\">${it.sectionTitle || '(Untitled section)'}</div>\n              <div><small class=\"mono\">core: \"${it.titleCore}\"</small></div>\n              ${note}\n              <details style=\"margin-top:4px\"><summary>Preview content</summary>\n                <div style=\"border:1px solid #eee;border-radius:6px;padding:8px;margin-top:6px;max-height:220px;overflow:auto\">${it.html || '<em>(no content)</em>'}</div>\n              </details>\n            </div>\n            <div class=\"wkx-matchbox\">\n              <div>\n                ${matchHtml}\n                <div style=\"margin-top:6px\"><button type=\"button\" class=\"Button Button--small wkx-change\">Change</button></div>\n              </div>\n              <div class=\"wkx-override\">\n                <div class=\"row\">\n                  <label>Pick assignment:\n                    <select class=\"wkx-select\">\n                      <optgroup label=\"Week ${wk} assignments\">${weekOpts}</optgroup>\n                      <optgroup label=\"Top suggestions\">${it.alts.map(a=>a.id?`<option value=\"${a.id}\">${a.name} (${Math.round(a.score*100)}%)</option>`:'').join('')}</optgroup>\n                      <optgroup label=\"All assignments\">${allOpts}</optgroup>\n                    </select>\n                  </label>\n                  <button type=\"button\" class=\"Button Button--small wkx-apply\">Apply</button>\n                  <button type=\"button\" class=\"Button Button--small Button--secondary wkx-cancel\">Cancel</button>\n                </div>\n                <small class=\"mono\">Tip: week items first, then top suggestions.</small>\n              </div>\n            </div>\n            <div><span class=\"badge ${cls}\">${Math.round(it.score*100)}%</span></div>\n          </div>`;
      }).join('');
      return `
        <div class=\"wkx-week\" data-week=\"${wk}\">\n          <div class=\"head\">\n            <div><strong>Week ${wk}</strong> <small class=\"mono\">— ${arr.length} sections</small></div>\n            <div><label style=\"display:inline-flex;align-items:center;gap:6px;cursor:pointer\"><input type=\"checkbox\" class=\"wkx-week-all\" data-week=\"${wk}\" checked> Select all good matches</label></div>\n          </div>\n          <div class=\"rows\">${rows || "<div style='padding:6px'>No sections</div>"}</div>\n        </div>`;
    }).join('');

    setSpin(false);
    byId('wkx-preview').innerHTML = `
      <div style="margin-bottom:6px">
        <ol style="margin:6px 0 8px;padding-left:18px">
          <li>Find all pages titled \"Week X: Assignments\".</li>
          <li>Parse each page into per-assignment sections for that week only.</li>
          <li>Match each section to assignments named \"Week X - Title\" (ignoring type prefixes). Use Change to override.</li>
          <li>Start to write descriptions and insert link lines on weekly pages.</li>
        </ol>
        <small class="mono">Ambiguous matches (within 2 percentage points) are warned and unchecked; one‑to‑one conflicts are flagged.</small>
      </div>
      ${weeksHtml}`;

    // Week-level toggles
    document.querySelectorAll('.wkx-week-all').forEach(cb=>{
      cb.addEventListener('change', ()=>{
        const wk = parseInt(cb.getAttribute('data-week'),10);
        document.querySelectorAll(`.wkx-week[data-week="${wk}"] .wkx-row`).forEach(row=>{
          const idx = parseInt(row.getAttribute('data-idx'),10);
          const it = preview.items.find(x=>x.idx===idx);
          if(!it) return; const eligible = it.match && !it.conflict && !it.ambiguous && it.score>=0.85;
          if(eligible) row.querySelector('.wkx-check').checked = cb.checked;
        });
      });
    });

    // Override controls
    document.querySelectorAll('.wkx-row').forEach(rowEl=>{
      const changeBtn=rowEl.querySelector('.wkx-change'); const panel=rowEl.querySelector('.wkx-override'); const applyBtn=rowEl.querySelector('.wkx-apply'); const cancelBtn=rowEl.querySelector('.wkx-cancel'); const select=rowEl.querySelector('.wkx-select');
      const idx=parseInt(rowEl.getAttribute('data-idx'),10); const item=preview.items.find(x=>x.idx===idx);
      if(item.match){ const opt=[...select.options].find(o=>parseInt(o.value,10)===item.match.id); if(opt) select.value=String(item.match.id); }
      if(item.conflict||item.ambiguous){ panel.classList.add('on'); }
      changeBtn.addEventListener('click',()=>panel.classList.add('on'));
      cancelBtn.addEventListener('click',()=>panel.classList.remove('on'));
      applyBtn.addEventListener('click',()=>{
        const id = parseInt(select.value,10); const foundInWeek=(preview.byWeek.get(item.week)||[]).find(a=>a.id===id); const foundAny=preview.allAssn.find(a=>a.id===id);
        const chosen = foundInWeek || (foundAny ? { id:foundAny.id, name:foundAny.name, url:foundAny.html_url, rest: foundAny.name.replace(/^week\s*\d+\s*-\s*/i,'').trim() } : null);
        if(!chosen){ alert('Could not find that assignment.'); return; }
        const restCore = coreTitle(chosen.rest || chosen.name);
        item.match = { id:chosen.id, name:chosen.name, url:chosen.url, rest:chosen.rest||chosen.name, restCore };
        item.score = 1; item.conflict = false; item.ambiguous = false;
        const linkBox = rowEl.querySelector('.wkx-matchbox > div:first-child');
        linkBox.innerHTML = `<a href="${item.match.url}" target="_blank" rel="noopener">${item.match.name}</a><div><small class=\"mono\">forced match</small></div><div style=\"margin-top:6px\"><button type=\"button\" class=\"Button Button--small wkx-change\">Change</button></div>`;
        const badge=rowEl.querySelector('.badge'); badge.className='badge ok'; badge.textContent='100%';
        linkBox.querySelector('.wkx-change').addEventListener('click',()=>panel.classList.add('on'));
        panel.classList.remove('on'); rowEl.querySelector('.wkx-check').checked=true;
      });
    });
  }

  document.getElementById('wkx-run').onclick = runUpdates;

  async function runUpdates(){
    if(!preview){ alert('Preview not ready yet.'); return; }
    const courseId = getCourseIdFromUrl();
    const append = document.getElementById('opt-append').checked; const skipNonEmpty = document.getElementById('opt-skip-nonempty').checked; const addLinks = document.getElementById('opt-insert-links').checked; const linksSelectedOnly = document.getElementById('opt-links-selected-only').checked;
    const rows=[...document.querySelectorAll('.wkx-row')]; const selected = rows.map(r=>{ const idx=parseInt(r.getAttribute('data-idx'),10); const cb=r.querySelector('.wkx-check'); return cb&&cb.checked? preview.items.find(x=>x.idx===idx):null; }).filter(Boolean);
    if(!selected.length){ alert('Nothing selected.'); return; }
    const summary = selected.map(s=>`• Week ${s.week}: ${s.sectionTitle} → ${s.match ? s.match.name : 'NO MATCH'}`).join('\n'); if(!confirm(`Proceed to update ${selected.length} assignments?\n\n${summary}`)) return;

    // 1) Update assignment descriptions
    setSpin(true,'Updating assignments…'); const results=[]; let i=0;
    for (const s of selected){ if(!s.match){ results.push({ok:false,s,err:'No match'}); continue; } i++; try {
        document.getElementById('wkx-progress').textContent = `Updating ${i} of ${selected.length} — Week ${s.week}: ${s.match.name}`;
        let newDescription = s.html || '';
        if (append || skipNonEmpty){ const current = await API.getJSON(`${API.base}/courses/${courseId}/assignments/${s.match.id}`); const existing = current.description || ''; if (skipNonEmpty && existing.replace(/\s|&nbsp;|<br\s*\/?>(?=\s|$)/gi,'') !== ''){ results.push({ok:true,s,skipped:true,reason:'Already had content'}); continue; } if (append && existing) newDescription = `${existing}\n\n${newDescription}`; }
        await API.put(`/courses/${courseId}/assignments/${s.match.id}`, { assignment: { description: newDescription } }); results.push({ok:true,s}); await delay(150);
      } catch(e){ results.push({ok:false,s,err:e.message}); }
    }

    // 2) Insert link lines on weekly pages
    const pageOps=[]; if(addLinks){ setSpin(true,'Updating weekly pages with links…'); const targets=new Map(); const pool = linksSelectedOnly? selected : preview.items.filter(x=>x.match); for (const it of pool){ if(!it.match) continue; if(!targets.has(it.pageUrl)) targets.set(it.pageUrl, []); targets.get(it.pageUrl).push(it); }
      let j=0, pagesToEdit=[...targets.keys()];
      for (const pageUrl of pagesToEdit){ j++; document.getElementById('wkx-progress').textContent = `Updating page ${j} of ${pagesToEdit.length}…`; const originalHtml = preview.pageBodies.get(pageUrl) || ''; let workingHtml = originalHtml; let changed=false; for (const it of targets.get(pageUrl)){ const res = upsertLinkAfterSection(workingHtml, it.sectionHeadingRaw, { id:it.match.id, name:it.match.name, url:it.match.url }); if (res.changed){ workingHtml = res.html; changed = true; } }
        if (changed){ try { await API.put(`/courses/${courseId}/pages/${encodeURIComponent(pageUrl)}`, { wiki_page: { body: workingHtml } }); pageOps.push({pageUrl, ok:true}); await delay(150); } catch(e){ pageOps.push({pageUrl, ok:false, err:e.message}); } } else { pageOps.push({pageUrl, ok:true, skipped:true, reason:'No changes needed'}); }
      }
    }

    setSpin(false);
    const ok = results.filter(r=>r.ok && !r.skipped); const skipped = results.filter(r=>r.ok && r.skipped); const bad = results.filter(r=>!r.ok);
    const updatedPages = pageOps.filter(p=>p.ok && !p.skipped); const skippedPages = pageOps.filter(p=>p.ok && p.skipped); const failedPages = pageOps.filter(p=>!p.ok);

    const htmlList = `<h3 style=\"margin:10px 0 6px\">Assignments updated</h3><ul>${ok.map(r=>`<li><a href=\"${r.s.match.url}\" target=\"_blank\" rel=\"noopener\">${r.s.match.name}</a></li>`).join('')||'<li>None</li>'}</ul>`+
      `${skipped.length?`<h3 style=\"margin:10px 0 6px\">Assignments skipped</h3><ul>${skipped.map(r=>`<li>Week ${r.s.week}: ${r.s.match.name} - ${r.reason}</li>`).join('')}</ul>`:''}`+
      `${bad.length?`<h3 style=\"margin:10px 0 6px\">Assignment failures</h3><ul>${bad.map(r=>`<li>Week ${r.s.week}: ${r.s.sectionTitle} - ${r.err||'Unknown error'}</li>`).join('')}</ul>`:''}`+
      `${pageOps.length?`<h3 style=\"margin:14px 0 6px\">Weekly pages updated</h3><ul>${updatedPages.map(p=>`<li>${p.pageUrl}</li>`).join('')||'<li>None</li>'}</ul>`+
        `${skippedPages.length?`<h3 style=\"margin:10px 0 6px\">Weekly pages skipped</h3><ul>${skippedPages.map(p=>`<li>${p.pageUrl} - ${p.reason}</li>`).join('')}</ul>`:''}`+
        `${failedPages.length?`<h3 style=\"margin:10px 0 6px\">Weekly page failures</h3><ul>${failedPages.map(p=>`<li>${p.pageUrl} - ${p.err||'Unknown error'}</li>`).join('')}</ul>`:''}`:''}`;

    const plainUpdated = ok.map(r=>`${r.s.match.name} - ${r.s.match.url}`).join('\n') || 'None';
    const plainSkipped = skipped.map(r=>`${r.s.match.name} - skipped (${r.reason})`).join('\n');
    const plainFailed = bad.map(r=>`Week ${r.s.week}: ${r.s.sectionTitle} - ${r.err||'Unknown error'}`).join('\n');
    const plainPageUpdated = updatedPages.map(p=>p.pageUrl).join('\n') || 'None';
    const plainPageSkipped = skippedPages.map(p=>`${p.pageUrl} - ${p.reason}`).join('\n');
    const plainPageFailed = failedPages.map(p=>`${p.pageUrl} - ${p.err||'Unknown error'}`).join('\n');

    done(`Process complete. ${ok.length} updated, ${skipped.length} skipped, ${bad.length} failed. ${pageOps.length?`Pages: ${updatedPages.length} updated, ${skippedPages.length} skipped, ${failedPages.length} failed.`:''}`,
      `<div style=\"margin-top:10px\">${htmlList}
         <div style=\"margin-top:12px;display:flex;gap:8px\">
           <button class=\"Button Button--secondary\" id=\"copy-html\">Copy HTML summary</button>
           <button class=\"Button Button--secondary\" id=\"copy-text\">Copy plain text</button>
         </div>
       </div>`
    );
    const htmlSummary = `<div>${htmlList}</div>`; const plainSummary = `Assignments updated:\n${plainUpdated}\n\nAssignments skipped:\n${plainSkipped||'None'}\n\nAssignment failures:\n${plainFailed||'None'}\n\n${pageOps.length?`Weekly pages updated:\n${plainPageUpdated}\n\nWeekly pages skipped:\n${plainPageSkipped||'None'}\n\nWeekly page failures:\n${plainPageFailed||'None'}\n`:''}`;
    document.getElementById('copy-html').onclick = () => copyHtmlToClipboard(htmlSummary, plainSummary);
    document.getElementById('copy-text').onclick = () => navigator.clipboard.writeText(plainSummary).then(()=>alert('Copied.'));
  }
})();
