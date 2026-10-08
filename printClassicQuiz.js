(async function canvasQuizPrinterPro() {
  /* Canvas Classic Quiz Printer Pro v3.0
     Standalone bookmarklet script. Generates genuine DOCX and print/PDF.
     Read-only Canvas API, with no external library dependency.
  */
  "use strict";

  const pathMatch = location.pathname.match(/^\/courses\/(\d+)\/quizzes\/(\d+)\/?$/);
  if (!pathMatch) {
    alert("Open the main details page of a Canvas Classic Quiz first.");
    return;
  }
  const courseId = pathMatch[1];
  const quizId = pathMatch[2];
  const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
  const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

  function esc(v) {
    return String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;")
      .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, "");
  }
  function clean(html) {
    const div = document.createElement("div");
    div.innerHTML = html || "";
    div.querySelectorAll("script,style,link,iframe,object,embed,form,button").forEach(n => n.remove());
    div.querySelectorAll("*").forEach(n => {
      for (const a of [...n.attributes]) {
        if (/^on/i.test(a.name) || (["href","src"].includes(a.name.toLowerCase()) &&
          /^\s*javascript:/i.test(a.value))) n.removeAttribute(a.name);
      }
    });
    return div.innerHTML;
  }
  function plain(html) {
    const div = document.createElement("div");
    div.innerHTML = clean(html);
    return (div.textContent || "").trim();
  }
  function answerHTML(a) {
    return a?.html || a?.answer_html || a?.text || a?.answer_text || "";
  }
  function correct(a) {
    return a?.is_correct === true || Number(a?.weight ?? a?.answer_weight ?? 0) > 0;
  }
  function letter(n) {
    let str = "";
    do {
      str = String.fromCharCode(65 + n % 26) + str;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return str;
  }
  function seededShuffle(items, seed) {
    let hash = 2166136261;
    for (const ch of seed) {
      hash ^= ch.charCodeAt(0);
      hash = Math.imul(hash, 16777619);
    }
    let state = hash >>> 0;
    function rnd() {
      state += 0x6d2b79f5;
      let v = Math.imul(state ^ (state >>> 15), 1 | state);
      v ^= v + Math.imul(v ^ (v >>> 7), 61 | v);
      return ((v ^ (v >>> 14)) >>> 0) / 4294967296;
    }
    const result = [...items];
    for (let i = result.length - 1; i > 0; i--) {
      const j = Math.floor(rnd() * (i + 1));
      [result[i], result[j]] = [result[j], result[i]];
    }
    return result;
  }
  function answersFor(q, opt) {
    const items = q.answers || [];
    return opt.shuffle && ["multiple_choice_question","multiple_answers_question"].includes(q.question_type)
      ? seededShuffle(items, `${courseId}:${quizId}:${q.id}:answers`)
      : [...items];
  }
  function includedQuestions(questions, opt) {
    const groups = new Set();
    return questions.filter(q => {
      if (!opt.onePerGroup || !q.quiz_group_id) return true;
      if (groups.has(q.quiz_group_id)) return false;
      groups.add(q.quiz_group_id);
      return true;
    });
  }
  function numericAnswer(q) {
    return (q.answers || []).map(a => {
      if (a.exact !== undefined) return a.margin ? `${a.exact} ± ${a.margin}` : `${a.exact}`;
      if (a.start !== undefined && a.end !== undefined) return `${a.start} to ${a.end}`;
      if (a.approximate !== undefined) return `Approximately ${a.approximate}`;
      return `${a.numerical_answer ?? a.text ?? ""}`;
    }).filter(Boolean).join(" | ");
  }
  function blankKey(q, id) {
    return (q.answers || [])
      .filter(a => `${a.blank_id}` === `${id}` && (a.weight === undefined || correct(a)))
      .map(a => plain(answerHTML(a))).filter(Boolean).join(" | ");
  }
  function fillBlanks(html, q, showKey) {
    const div = document.createElement("div");
    div.innerHTML = clean(html);
    const walker = document.createTreeWalker(div, NodeFilter.SHOW_TEXT);
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    for (const node of nodes) {
      const text = node.nodeValue || "";
      if (!/\[[^\]]+\]/.test(text)) continue;
      const f = document.createDocumentFragment();
      let pos = 0;
      const regex = /\[([^\]]+)\]/g;
      for (const m of text.matchAll(regex)) {
        if (m.index > pos) f.appendChild(document.createTextNode(text.slice(pos, m.index)));
        const span = document.createElement("span");
        span.className = "blank";
        span.textContent = "____________";
        f.appendChild(span);
        const key = showKey ? blankKey(q, m[1]) : "";
        if (key) f.appendChild(document.createTextNode(` (${key})`));
        pos = m.index + m[0].length;
      }
      if (pos < text.length) f.appendChild(document.createTextNode(text.slice(pos)));
      node.replaceWith(f);
    }
    return div.innerHTML;
  }
  function keyFor(q, opt) {
    const type = q.question_type;
    if (["multiple_choice_question","multiple_answers_question","true_false_question"].includes(type)) {
      return answersFor(q, opt).map((a, i) => correct(a) ? `${letter(i)}. ${plain(answerHTML(a))}` : "")
        .filter(Boolean).join(" | ") || "No correct choice identified";
    }
    if (type === "short_answer_question") {
      return (q.answers || []).map(a => plain(answerHTML(a))).filter(Boolean).join(" | ") || "Review manually";
    }
    if (["fill_in_multiple_blanks_question","multiple_dropdowns_question"].includes(type)) {
      return [...new Set((q.answers || []).map(a => a.blank_id).filter(Boolean))]
        .map(id => `${id}: ${blankKey(q, id)}`).join("; ") || "Review manually";
    }
    if (type === "matching_question") {
      return (q.answers || []).map((a, i) =>
        `${i + 1}. ${plain(a.match_question || a.answer_match_left || "")}: ${plain(a.text || a.answer_match_right || "")}`
      ).join("; ") || "Review manually";
    }
    if (type === "numerical_question") return numericAnswer(q) || "Review manually";
    if (type === "essay_question") return "Manual grading";
    if (type === "file_upload_question") return "File upload";
    if (type === "calculated_question") return "Calculated values vary";
    return "Review manually";
  }

  async function getJSON(url) {
    const r = await fetch(url, { credentials: "same-origin", headers: { Accept: "application/json" } });
    if (!r.ok) throw new Error(`Canvas API ${r.status}: ${(await r.text()).slice(0, 350)}`);
    return r.json();
  }
  async function getPages(url) {
    const out = [];
    let next = url;
    while (next) {
      const r = await fetch(next, { credentials: "same-origin", headers: { Accept: "application/json" } });
      if (!r.ok) throw new Error(`Canvas API ${r.status}: ${(await r.text()).slice(0, 350)}`);
      const data = await r.json();
      if (!Array.isArray(data)) throw new Error("Canvas returned unexpected question data.");
      out.push(...data);
      next = null;
      for (const part of (r.headers.get("Link") || "").split(",")) {
        if (/rel=["']?next["']?/.test(part)) {
          next = part.match(/<([^>]+)>/)?.[1] || null;
          break;
        }
      }
    }
    return out;
  }

  const existing = document.getElementById("quiz-printer-pro");
  if (existing) existing.remove();
  const dialog = document.createElement("div");
  dialog.id = "quiz-printer-pro";
  Object.assign(dialog.style, {
    position:"fixed",inset:"0",zIndex:"999999",display:"flex",alignItems:"center",
    justifyContent:"center",background:"rgba(0,0,0,.55)",fontFamily:"Arial,sans-serif"
  });
  dialog.innerHTML = `
    <section style="width:450px;max-width:95vw;max-height:94vh;overflow:auto;background:white;
      border-radius:10px;box-shadow:0 14px 40px #0004;color:#1e293b">
      <header style="padding:19px 20px;background:#f8fafc;border-bottom:1px solid #ddd">
        <h2 style="font-size:19px;margin:0">Canvas Quiz Printer Pro</h2>
        <p style="margin:6px 0 0;color:#64748b;font-size:13px">Printable quiz and native Word export</p>
      </header>
      <div style="padding:20px;display:flex;flex-direction:column;gap:14px;font-size:14px">
        <label><input id="qp-points" type="checkbox" checked> Show point values</label>
        <label><input id="qp-groups" type="checkbox" checked> One sample question per random group</label>
        <label><input id="qp-shuffle" type="checkbox"> Shuffle answer choices</label>
        <label><input id="qp-key" type="checkbox"> Include correct answers and answer key</label>
        <label><input id="qp-link" type="checkbox"> Include online quiz link</label>
        <div style="padding-top:12px;border-top:1px solid #ddd">
          <label for="qp-format" style="font-size:12px;font-weight:bold;display:block;margin-bottom:6px">OUTPUT FORMAT</label>
          <select id="qp-format" style="width:100%;padding:10px;background:white;border:1px solid #ccc;border-radius:6px">
            <option value="print">Print Preview / PDF</option>
            <option value="docx">Microsoft Word (.docx)</option>
          </select>
        </div>
        <div id="qp-status" role="status" style="font-size:13px;color:#475569;min-height:18px;overflow-wrap:anywhere"></div>
      </div>
      <footer style="padding:14px 20px;border-top:1px solid #ddd;background:#f8fafc;
        display:flex;gap:10px;justify-content:flex-end">
        <button id="qp-cancel" style="padding:9px 15px;background:white;border:1px solid #ccc;border-radius:6px;cursor:pointer">Cancel</button>
        <button id="qp-generate" style="padding:9px 18px;border:0;border-radius:6px;cursor:pointer;
          background:#003865;color:white;font-weight:bold">Generate</button>
      </footer>
    </section>`;
  document.body.appendChild(dialog);
  const sel = id => dialog.querySelector(id);
  const status = msg => { sel("#qp-status").textContent = msg; };
  const settings = await new Promise(resolve => {
    sel("#qp-cancel").onclick = () => { dialog.remove(); resolve(null); };
    sel("#qp-generate").onclick = () => {
      const opt = {
        points:sel("#qp-points").checked,
        onePerGroup:sel("#qp-groups").checked,
        shuffle:sel("#qp-shuffle").checked,
        key:sel("#qp-key").checked,
        link:sel("#qp-link").checked,
        format:sel("#qp-format").value
      };
      let preview = null;
      if (opt.format === "print") {
        preview = window.open("", "_blank");
        if (!preview) {
          status("Popup blocked. Allow popups for Canvas and try again.");
          return;
        }
        preview.document.body.textContent = "Preparing quiz...";
      }
      sel("#qp-generate").disabled = true;
      sel("#qp-generate").textContent = "Processing...";
      resolve({opt, preview});
    };
  });
  if (!settings) return;
  const {opt, preview} = settings;

  function responseLines(n) {
    return `<div class="response-lines">${Array.from({length:n}, () => '<div class="response-line"></div>').join("")}</div>`;
  }
  function matchingTable(q, key) {
    return `<table class="matching"><thead><tr><th>Item</th><th>Match</th></tr></thead><tbody>${
      (q.answers || []).map((a,i) => `<tr><td>${clean(a.match_question || a.answer_match_left || `Item ${i + 1}`)}</td>
        <td>________________________ ${key ? `<strong>${esc(plain(a.text || a.answer_match_right || ""))}</strong>` : ""}</td></tr>`).join("")
    }</tbody></table>`;
  }

  function printHTML(course, quiz, questions) {
    let n = 0;
    const keyRows = [];
    const blocks = includedQuestions(questions, opt).map(q => {
      const type = q.question_type;
      if (type === "text_only_question") return `<div class="text-only">${clean(q.question_text)}</div>`;
      n++;
      let question = clean(q.question_text);
      let answers = "";
      if (["fill_in_multiple_blanks_question","multiple_dropdowns_question"].includes(type))
        question = fillBlanks(question, q, opt.key);
      if (["multiple_choice_question","multiple_answers_question","true_false_question"].includes(type)) {
        answers = `<div class="answer-group">${answersFor(q, opt).map((a, i) =>
          `<div class="choice ${opt.key && correct(a) ? "correct" : ""}">
            <span class="choice-letter">${letter(i)}.</span><div>${clean(answerHTML(a))}
            ${opt.key && correct(a) ? '<strong class="check">✓</strong>' : ""}</div></div>`).join("")}</div>`;
      } else if (type === "short_answer_question") answers = responseLines(4);
      else if (type === "essay_question") answers = responseLines(12);
      else if (type === "matching_question") answers = matchingTable(q, opt.key);
      else if (type === "numerical_question") answers = "<p>Answer: __________________________</p>";
      else if (["fill_in_multiple_blanks_question","multiple_dropdowns_question"].includes(type)) answers = "";
      else answers = responseLines(5);
      if (opt.key) keyRows.push({n, text:keyFor(q, opt)});
      return `<section class="question"><div class="qnum">${n}. ${opt.points ? `<small>${esc(q.points_possible ?? "")} pts</small>` : ""}</div>
        <div class="qbody"><div class="prompt">${question}</div>${answers}</div></section>`;
    }).join("");
    const url = `${location.origin}/courses/${courseId}/quizzes/${quizId}`;
    return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(quiz.title || "Quiz")}</title>
      <style>
      @page{size:letter;margin:.65in}
      body{font:11pt/1.45 Arial,sans-serif;color:#111;max-width:850px;margin:auto;padding:18px}
      .toolbar{border:1px solid #cbd5e1;background:#f8fafc;padding:12px;font-size:13px;margin-bottom:22px}
      .toolbar button{background:#003865;color:white;border:0;border-radius:5px;padding:9px 16px;margin-right:12px;cursor:pointer}
      header.doc-header{border-bottom:2px solid #222;padding-bottom:14px;margin-bottom:20px}
      h1{font-size:21pt;margin:0 0 10px}.meta{font-size:10pt}.student{display:flex;justify-content:space-between;margin-top:18px;font-size:10pt}
      .instructions{background:#f7f7f7;border:1px solid #ddd;padding:10px;margin-bottom:20px}
      .text-only{padding:10px;border-left:3px solid #aaa;margin-bottom:18px;break-inside:avoid}
      .question{display:grid;grid-template-columns:45px minmax(0,1fr);gap:12px;margin-bottom:22px;
        break-inside:avoid;page-break-inside:avoid}
      .qnum{text-align:right;font-weight:bold}.qnum small{display:block;color:#666;font-size:8pt;font-weight:normal}
      .qbody{min-width:0}.qbody img{max-width:100%;height:auto}.prompt{break-after:avoid;page-break-after:avoid}
      .prompt p:first-child{margin-top:0}
      .answer-group{margin-top:10px;break-inside:avoid;page-break-inside:avoid}
      .choice{display:grid;grid-template-columns:28px minmax(0,1fr);padding:3px 0;break-inside:avoid;page-break-inside:avoid}
      .choice-letter{font-weight:bold}.choice p{margin:0}.correct{font-weight:bold}.check{color:#157a37;margin-left:8px}
      .response-line{height:25px;border-bottom:1px solid #bbb;margin:7px 0}
      .matching{width:100%;border-collapse:collapse;margin-top:10px}
      .matching td,.matching th{border:1px solid #ccc;padding:8px;text-align:left}
      .blank{display:inline-block;min-width:110px;border-bottom:1px solid #333}
      .key{break-before:page;page-break-before:always}.key h2{font-size:18pt;border-bottom:2px solid #222}
      .key-row{display:grid;grid-template-columns:40px 1fr;padding:6px 0;border-bottom:1px solid #ddd;break-inside:avoid}
      @media print{body{margin:0;padding:0;max-width:none}.toolbar{display:none!important}}
      </style></head><body>
      <div class="toolbar"><button onclick="window.print()">Print / Save PDF</button>
      Turn off browser headers and footers for the cleanest PDF.</div>
      <header class="doc-header"><h1>${esc(quiz.title || "Quiz")}</h1>
      <div class="meta"><div><b>Course:</b> ${esc(course.name)}</div>
      <div><b>Total Points:</b> ${esc(quiz.points_possible ?? "")}</div>
      ${opt.link ? `<div><b>Online:</b> ${esc(url)}</div>` : ""}</div>
      <div class="student"><span>Name: ______________________________</span><span>Score: ______________</span></div></header>
      ${quiz.description ? `<div class="instructions">${clean(quiz.description)}</div>` : ""}
      <main>${blocks}</main>
      ${opt.key ? `<section class="key"><h2>Answer Key</h2>${keyRows.map(k =>
        `<div class="key-row"><b>${k.n}.</b><span>${esc(k.text)}</span></div>`).join("")}</section>` : ""}
      </body></html>`;
  }

  function run(text, style = {}) {
    const pr = [];
    if (style.bold) pr.push("<w:b/>");
    if (style.italic) pr.push("<w:i/>");
    if (style.underline) pr.push('<w:u w:val="single"/>');
    if (style.super) pr.push('<w:vertAlign w:val="superscript"/>');
    if (style.sub) pr.push('<w:vertAlign w:val="subscript"/>');
    if (style.color) pr.push(`<w:color w:val="${style.color}"/>`);
    if (style.size) pr.push(`<w:sz w:val="${style.size}"/>`);
    const rPr = pr.length ? `<w:rPr>${pr.join("")}</w:rPr>` : "";
    return `<w:r>${rPr}${String(text ?? "").split(/\r?\n/).map((v,i) =>
      `${i ? "<w:br/>" : ""}<w:t xml:space="preserve">${esc(v)}</w:t>`).join("")}</w:r>`;
  }
  function paragraph(contents, config = {}) {
    const attrs = [];
    if (config.keepNext) attrs.push("<w:keepNext/>");
    attrs.push("<w:keepLines/>");
    if (config.newPage) attrs.push("<w:pageBreakBefore/>");
    attrs.push(`<w:spacing w:before="${config.before ?? 0}" w:after="${config.after ?? 90}" w:line="290" w:lineRule="auto"/>`);
    if (config.indent || config.hanging)
      attrs.push(`<w:ind w:left="${config.indent || 0}" w:hanging="${config.hanging || 0}"/>`);
    if (config.line) attrs.push('<w:pBdr><w:bottom w:val="single" w:sz="4" w:color="BBBBBB"/></w:pBdr>');
    return `<w:p><w:pPr>${attrs.join("")}</w:pPr>${contents || run(" ")}</w:p>`;
  }

  class NativeDocx {
    constructor() {
      this.images = [];
      this.imageCache = new Map();
      this.relations = [];
      this.imgCount = 0;
      this.drawingCount = 0;
      this.imageWarnings = 0;
    }
    async imageRun(element) {
      const src = element.getAttribute("src");
      const alt = element.getAttribute("alt") || "Image";
      if (!src) return run(`[${alt}]`);
      try {
        const srcUrl = new URL(src, location.href).href;
        let image = this.imageCache.get(srcUrl);
        if (!image) {
          status("Embedding quiz images...");
          const response = await fetch(srcUrl, { credentials:"include" });
          if (!response.ok) throw Error(`HTTP ${response.status}`);
          let blob = await response.blob();
          if (blob.size > 12 * 1024 * 1024) throw Error("Image exceeds 12 MB");
          let mime = blob.type.split(";")[0].toLowerCase();
          if (!["image/png","image/jpeg","image/gif"].includes(mime)) {
            const bitmap = await createImageBitmap(blob);
            const canvas = document.createElement("canvas");
            canvas.width = bitmap.width;
            canvas.height = bitmap.height;
            canvas.getContext("2d").drawImage(bitmap, 0, 0);
            bitmap.close();
            const b = await new Promise(resolve => canvas.toBlob(resolve, "image/png"));
            if (!b) throw Error("Image conversion failed");
            blob = b;
            mime = "image/png";
          }
          const bitmap = await createImageBitmap(blob);
          const width = bitmap.width;
          const height = bitmap.height;
          bitmap.close();
          const extension = mime === "image/jpeg" ? "jpg" : mime === "image/gif" ? "gif" : "png";
          const id = ++this.imgCount;
          const filename = `image${id}.${extension}`;
          const relId = `rIdImage${id}`;
          this.images.push({name:`word/media/${filename}`, data:new Uint8Array(await blob.arrayBuffer())});
          this.relations.push(`<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${filename}"/>`);
          image = {relId,width,height};
          this.imageCache.set(srcUrl,image);
        }
        const scale = Math.min(1, 580 / image.width, 720 / image.height);
        const cx = Math.round(image.width * scale * 9525);
        const cy = Math.round(image.height * scale * 9525);
        const id = ++this.drawingCount;
        return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="${cx}" cy="${cy}"/>
          <wp:docPr id="${id}" name="Quiz Image ${id}" descr="${esc(alt)}"/>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Quiz Image ${id}"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="${image.relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
          </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
      } catch (e) {
        this.imageWarnings++;
        return run(`[Image unavailable: ${alt}]`);
      }
    }
    async inline(nodes, style = {}) {
      let result = "";
      for (const node of nodes) {
        if (node.nodeType === Node.TEXT_NODE) {
          result += run(node.nodeValue, style);
          continue;
        }
        if (node.nodeType !== Node.ELEMENT_NODE) continue;
        const tag = node.tagName.toLowerCase();
        if (tag === "img") {result += await this.imageRun(node);continue;}
        if (tag === "br") {result += "<w:r><w:br/></w:r>";continue;}
        if (["script","style","iframe","object"].includes(tag)) continue;
        const next = {...style};
        if (["b","strong"].includes(tag)) next.bold = true;
        if (["i","em"].includes(tag)) next.italic = true;
        if (tag === "u") next.underline = true;
        if (tag === "sup") next.super = true;
        if (tag === "sub") next.sub = true;
        if (tag === "a") {next.underline = true;next.color = "0563C1";}
        result += await this.inline([...node.childNodes], next);
      }
      return result;
    }
    async richParagraphs(html) {
      const div = document.createElement("div");
      div.innerHTML = clean(html);
      const output = [];
      let pending = [];
      async function pushPending(builder) {
        if (pending.length) {
          const r = await builder.inline(pending);
          if (r) output.push(r);
          pending = [];
        }
      }
      const blocks = new Set(["P","DIV","SECTION","ARTICLE","H1","H2","H3","H4","H5","H6","BLOCKQUOTE","UL","OL","TABLE","PRE"]);
      const walk = async el => {
        for (const node of [...el.childNodes]) {
          if (node.nodeType !== Node.ELEMENT_NODE || !blocks.has(node.tagName)) {
            pending.push(node);continue;
          }
          await pushPending(this);
          if (["UL","OL"].includes(node.tagName)) {
            let i = Number(node.getAttribute("start")) || 1;
            for (const li of [...node.children]) {
              if (li.tagName !== "LI") continue;
              const label = node.tagName === "OL" ? `${i++}. ` : "• ";
              const content = await this.inline([...li.childNodes].filter(c =>
                c.nodeType !== Node.ELEMENT_NODE || !["UL","OL"].includes(c.tagName)));
              output.push(run(label) + content);
            }
          } else if (node.tagName === "TABLE") {
            for (const tr of node.querySelectorAll("tr")) {
              const cells = [...tr.children].filter(c => ["TD","TH"].includes(c.tagName));
              if (cells.length) {
                const parts = await Promise.all(cells.map(c => this.inline([...c.childNodes])));
                output.push(parts.join(run("  |  ")));
              }
            }
          } else if ([...node.children].some(c => blocks.has(c.tagName)) && node.tagName === "DIV") {
            await walk(node);
          } else {
            const style = /^H[1-6]$/.test(node.tagName) ? {bold:true} : {};
            output.push(await this.inline([...node.childNodes], style));
          }
        }
        await pushPending(this);
      };
      await walk(div);
      return output.filter(Boolean).length ? output.filter(Boolean) : [run(" ")];
    }
    async rich(html, config = {}, label = "") {
      const entries = await this.richParagraphs(html);
      return entries.map((r,i) => paragraph((i===0 ? label : "") + r, {
        indent:config.indent ?? 0,
        hanging:config.hanging ?? 0,
        before:i===0 ? (config.before ?? 0) : 0,
        after:i===entries.length-1 ? (config.after ?? 90) : 60,
        keepNext:i<entries.length-1 || Boolean(config.keepNext)
      }));
    }
    lines(n) {
      return Array.from({length:n}, () => paragraph(run(" "), {line:true,after:190}));
    }
    async create(course, quiz, questions) {
      const sections = [];
      const title = quiz.title || "Quiz";
      const url = `${location.origin}/courses/${courseId}/quizzes/${quizId}`;
      sections.push(paragraph(run(title,{bold:true,size:36}),{after:170}));
      sections.push(paragraph(run(`Course: ${course.name || ""}`),{after:50}));
      sections.push(paragraph(run(`Total Points: ${quiz.points_possible ?? ""}`),{after:50}));
      if (opt.link) sections.push(paragraph(run(`Online Quiz: ${url}`,{color:"0563C1",underline:true,size:18}),{after:130}));
      sections.push(paragraph(run("Name: ___________________________       Score: __________"),{after:260}));
      if (quiz.description) {
        sections.push(paragraph(run("Instructions",{bold:true}),{keepNext:true,after:90}));
        sections.push(...await this.rich(quiz.description,{after:160}));
      }
      let n=0;
      const keyRows=[];
      const list=includedQuestions(questions,opt);
      for (let i=0; i<list.length; i++) {
        const q=list[i];
        status(`Building Word document: ${i+1} of ${list.length}`);
        if (q.question_type === "text_only_question") {
          sections.push(...await this.rich(q.question_text,{before:120,after:130}));continue;
        }
        n++;
        const type = q.question_type;
        const prompt = ["fill_in_multiple_blanks_question","multiple_dropdowns_question"].includes(type)
          ? fillBlanks(q.question_text,q,opt.key) : q.question_text || "";
        sections.push(...await this.rich(prompt,
          {indent:380,hanging:380,before:160,after:95,keepNext:true},run(`${n}. `,{bold:true})));
        if (opt.points) sections.push(paragraph(run(`(${q.points_possible ?? 0} pts)`,{size:16,color:"666666"}),{
          indent:380,after:95,keepNext:true}));
        if (["multiple_choice_question","multiple_answers_question","true_false_question"].includes(type)) {
          const answers = answersFor(q,opt);
          if (!answers.length) sections.push(paragraph(run("[No choices returned by Canvas]"),{after:180}));
          for (let j=0;j<answers.length;j++) {
            const a = answers[j];
            const label = run(`${letter(j)}. `,{bold:true});
            const isRight = opt.key && correct(a);
            // Keep each answer (including its rich text) with the next answer,
            // except for the final answer. Word can still split oversized groups.
            const keepNext = j < answers.length-1 || isRight;
            sections.push(...await this.rich(answerHTML(a),{
              indent:700,hanging:280,
              after:j===answers.length-1 && !isRight ? 230 : 75,
              keepNext
            },label));
            if (isRight) sections.push(paragraph(run("Correct answer",{bold:true,color:"15803D",size:17}),{
              indent:700,after:j===answers.length-1 ? 220 : 75,
              keepNext:j<answers.length-1
            }));
          }
        } else if (type === "essay_question") {
          sections.push(...this.lines(12));
        } else if (type === "short_answer_question") {
          sections.push(...this.lines(4));
        } else if (["fill_in_multiple_blanks_question","multiple_dropdowns_question"].includes(type)) {
          sections.push(paragraph(run(" "),{after:160}));
        } else if (type === "matching_question") {
          const answers = q.answers || [];
          for (let j=0;j<answers.length;j++) {
            const a=answers[j];
            sections.push(paragraph(run(`${j+1}. ${plain(a.match_question || a.answer_match_left || "")}    ____________`),{
              indent:700,after:90,keepNext:j<answers.length-1 || opt.key}));
            if (opt.key) sections.push(paragraph(run(`Match: ${plain(a.text || a.answer_match_right || "")}`,{color:"15803D",size:17}),{
              indent:700,after:75,keepNext:j<answers.length-1}));
          }
        } else if (type === "numerical_question") {
          sections.push(paragraph(run("Answer: __________________________"),{indent:700,after:220}));
        } else {
          sections.push(...this.lines(5));
        }
        if (opt.key) keyRows.push({number:n,text:keyFor(q,opt)});
      }
      if (opt.key && keyRows.length) {
        sections.push(paragraph(run("Answer Key",{bold:true,size:30}),{newPage:true,keepNext:true,after:220}));
        for (const key of keyRows) sections.push(paragraph(
          run(`${key.number}. `,{bold:true})+run(key.text),{indent:380,hanging:380,after:120}));
      }
      const documentXML = `<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}">
        <w:body>${sections.join("")}
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
        <w:pgMar w:top="936" w:right="936" w:bottom="936" w:left="936"
          w:header="450" w:footer="450" w:gutter="0"/></w:sectPr></w:body></w:document>`;
      const stylesXML = `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${W}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
        <w:sz w:val="22"/></w:rPr></w:rPrDefault>
        <w:pPrDefault><w:pPr><w:spacing w:after="90" w:line="290" w:lineRule="auto"/></w:pPr>
        </w:pPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
        </w:styles>`;
      const rels = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PKG_REL}">
        <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        ${this.relations.join("")}</Relationships>`;
      const rootRels = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PKG_REL}">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>`;
      const contentTypes = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
        <Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        </Types>`;
      return zip([
        {name:"[Content_Types].xml",data:contentTypes},
        {name:"_rels/.rels",data:rootRels},
        {name:"word/document.xml",data:documentXML},
        {name:"word/styles.xml",data:stylesXML},
        {name:"word/_rels/document.xml.rels",data:rels},
        ...this.images
      ]);
    }
  }

  // Pure JavaScript ZIP (STORE method), producing a valid DOCX package.
  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let n=0;n<256;n++) {
      let c=n;
      for (let i=0;i<8;i++) c=c&1 ? 0xedb88320 ^ (c>>>1) : c>>>1;
      table[n]=c>>>0;
    }
    return table;
  })();
  function crc32(bytes) {
    let c=0xffffffff;
    for (const b of bytes) c=crcTable[(c^b)&255]^(c>>>8);
    return (c^0xffffffff)>>>0;
  }
  function zip(entries) {
    const enc=new TextEncoder();
    const fileParts=[];
    const central=[];
    let offset=0;
    const now=new Date();
    const dt=(now.getHours()<<11)|(now.getMinutes()<<5)|(Math.floor(now.getSeconds()/2));
    const dd=((Math.max(1980,now.getFullYear())-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate();
    for (const file of entries) {
      const name=enc.encode(file.name);
      const bytes=typeof file.data==="string" ? enc.encode(file.data) : file.data;
      const checksum=crc32(bytes);
      const header=new Uint8Array(30+name.length);
      const h=new DataView(header.buffer);
      h.setUint32(0,0x04034b50,true);
      h.setUint16(4,20,true);
      h.setUint16(6,0x0800,true);
      h.setUint16(8,0,true);
      h.setUint16(10,dt,true);
      h.setUint16(12,dd,true);
      h.setUint32(14,checksum,true);
      h.setUint32(18,bytes.length,true);
      h.setUint32(22,bytes.length,true);
      h.setUint16(26,name.length,true);
      h.setUint16(28,0,true);
      header.set(name,30);
      fileParts.push(header,bytes);
      const cent=new Uint8Array(46+name.length);
      const c=new DataView(cent.buffer);
      c.setUint32(0,0x02014b50,true);
      c.setUint16(4,20,true);
      c.setUint16(6,20,true);
      c.setUint16(8,0x0800,true);
      c.setUint16(10,0,true);
      c.setUint16(12,dt,true);
      c.setUint16(14,dd,true);
      c.setUint32(16,checksum,true);
      c.setUint32(20,bytes.length,true);
      c.setUint32(24,bytes.length,true);
      c.setUint16(28,name.length,true);
      c.setUint16(30,0,true);
      c.setUint16(32,0,true);
      c.setUint16(34,0,true);
      c.setUint16(36,0,true);
      c.setUint32(38,0,true);
      c.setUint32(42,offset,true);
      cent.set(name,46);
      central.push(cent);
      offset += header.length + bytes.length;
    }
    const cdSize=central.reduce((sum,x)=>sum+x.length,0);
    const end=new Uint8Array(22);
    const e=new DataView(end.buffer);
    e.setUint32(0,0x06054b50,true);
    e.setUint16(4,0,true);
    e.setUint16(6,0,true);
    e.setUint16(8,entries.length,true);
    e.setUint16(10,entries.length,true);
    e.setUint32(12,cdSize,true);
    e.setUint32(16,offset,true);
    e.setUint16(20,0,true);
    return new Blob([...fileParts,...central,end],{type:DOCX_MIME});
  }
  function downloadFile(blob,filename) {
    const url=URL.createObjectURL(blob);
    const link=document.createElement("a");
    link.href=url;
    link.download=filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(()=>URL.revokeObjectURL(url),10000);
  }
  try {
    status("Loading quiz information...");
    const [course,quiz]=await Promise.all([
      getJSON(`/api/v1/courses/${courseId}`),
      getJSON(`/api/v1/courses/${courseId}/quizzes/${quizId}`)
    ]);
    status("Loading quiz questions...");
    const questions=await getPages(`/api/v1/courses/${courseId}/quizzes/${quizId}/questions?per_page=100`);
    questions.sort((a,b)=>(a.position??0)-(b.position??0));
    if (!questions.length) throw Error("Canvas returned no quiz questions.");
    const filename=(quiz.title||"Quiz").replace(/[^a-z0-9]+/gi,"_").replace(/^_+|_+$/g,"")||"Quiz";
    if (opt.format==="docx") {
      const builder=new NativeDocx();
      status("Creating native Word document...");
      const blob=await builder.create(course,quiz,questions);
      downloadFile(blob,`${filename}.docx`);
      if (builder.imageWarnings) {
        status(`Word document saved. ${builder.imageWarnings} image(s) could not be embedded and are identified in the document.`);
        const gen=sel("#qp-generate");
        gen.textContent="Close";
        gen.disabled=false;
        gen.onclick=()=>dialog.remove();
      } else {
        dialog.remove();
      }
    } else {
      status("Preparing print preview...");
      preview.document.open();
      preview.document.write(printHTML(course,quiz,questions));
      preview.document.close();
      preview.focus();
      dialog.remove();
    }
  } catch(e) {
    console.error(e);
    status(`Error: ${e.message}`);
    sel("#qp-status").style.color="#b91c1c";
    const btn=sel("#qp-generate");
    btn.textContent="Generate";
    btn.disabled=false;
    if (preview && !preview.closed) preview.close();
  }
})();
