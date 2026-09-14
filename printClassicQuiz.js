(async () => {
  /* =============== Canvas Quiz -> Printable View (Pro v2.7) ===============
     - Single configuration modal
     - Print Preview and Microsoft Word (.doc)
     - Cleaner lettered answer choices
     - Optional Answer Key
     - Optional deterministic answer shuffling
     - Random-group handling without showing a notice in the output
     - Screen-only print tip
     - Word-friendly table layout
  ===================================================================== */

  /* ---------- UI: Spinner ---------- */
  function showSpinner(msg = 'Processing...') {
    document.getElementById('quiz-print-spinner')?.remove();

    const el = document.createElement('div');
    el.id = 'quiz-print-spinner';

    el.innerHTML = `
      <div style="
        position:fixed;
        inset:0;
        background:rgba(0,0,0,.25);
        z-index:999999;
        display:flex;
        align-items:center;
        justify-content:center;
      ">
        <div style="
          background:#fff;
          border-radius:10px;
          padding:14px 16px;
          min-width:260px;
          box-shadow:0 10px 30px rgba(0,0,0,.25);
          font:14px/1.3 system-ui,-apple-system,'Segoe UI',Roboto,Arial,sans-serif;
        ">
          <div style="display:flex;gap:10px;align-items:center;">
            <div style="
              width:18px;
              height:18px;
              border:3px solid #aaa;
              border-top-color:transparent;
              border-radius:50%;
              animation:quizPrintSpin .8s linear infinite;
            "></div>
            <div>${msg}</div>
          </div>
        </div>
      </div>

      <style>
        @keyframes quizPrintSpin {
          to { transform: rotate(360deg); }
        }
      </style>
    `;

    document.body.appendChild(el);
  }

  function hideSpinner() {
    document.getElementById('quiz-print-spinner')?.remove();
  }

  /* ---------- UI: Toast ---------- */
  function toast(msg, type = 'success') {
    const el = document.createElement('div');
    const bg = type === 'error' ? '#EF4444' : '#10B981';

    el.innerHTML = `<div style="position:fixed;left:50%;bottom:30px;transform:translateX(-50%);
      background:${bg};color:white;padding:12px 18px;border-radius:8px;
      box-shadow:0 10px 15px rgba(0,0,0,0.15);z-index:999999;
      font:13px system-ui,-apple-system,'Segoe UI',Roboto,Arial,sans-serif;font-weight:600;">
      ${msg}
    </div>`;

    document.body.appendChild(el);
    setTimeout(() => el.remove(), 3200);
  }

  /* ---------- UI: Configuration Modal ---------- */
  function getPrintOptions() {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.id = 'quiz-print-modal';

      Object.assign(overlay.style, {
        position: 'fixed',
        inset: '0',
        backgroundColor: 'rgba(0,0,0,0.5)',
        backdropFilter: 'blur(2px)',
        zIndex: '999999',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontFamily: 'system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif'
      });

      overlay.innerHTML = `
        <div style="background:white;width:100%;max-width:440px;border-radius:10px;
                    box-shadow:0 20px 25px rgba(0,0,0,0.15);overflow:hidden;">
          <div style="padding:18px 20px;border-bottom:1px solid #e5e7eb;background:#f9fafb;">
            <div style="margin:0;font-size:1.05rem;font-weight:700;color:#111827;">Quiz Print Options</div>
            <div style="margin-top:6px;font-size:0.86rem;color:#6b7280;">Choose your output and formatting.</div>
          </div>

          <div style="padding:18px 20px;display:flex;flex-direction:column;gap:12px;">
            <label style="display:flex;align-items:center;gap:10px;cursor:pointer;">
              <input type="checkbox" id="opt-points" checked style="width:16px;height:16px;">
              <span style="font-size:0.95rem;color:#374151;">Show point values</span>
            </label>

            <label style="display:flex;align-items:flex-start;gap:10px;cursor:pointer;">
              <input type="checkbox" id="opt-groups" checked style="width:16px;height:16px;margin-top:3px;">
              <div style="display:flex;flex-direction:column;">
                <span style="font-size:0.95rem;color:#374151;">Random draw groups: show 1 sample question</span>
                <span style="font-size:0.78rem;color:#6b7280;">If the quiz draws randomly from a pool, print one example item per group.</span>
              </div>
            </label>

            <label style="display:flex;align-items:flex-start;gap:10px;cursor:pointer;">
              <input type="checkbox" id="opt-shuffle" style="width:16px;height:16px;margin-top:3px;">
              <div style="display:flex;flex-direction:column;">
                <span style="font-size:0.95rem;color:#374151;">Shuffle answer choices</span>
                <span style="font-size:0.78rem;color:#6b7280;">Uses a consistent shuffle so the Answer Key still matches.</span>
              </div>
            </label>

            <label style="display:flex;align-items:center;gap:10px;cursor:pointer;">
              <input type="checkbox" id="opt-key" style="width:16px;height:16px;">
              <span style="font-size:0.95rem;color:#374151;">Include Answer Key</span>
            </label>

            <label style="display:flex;align-items:center;gap:10px;cursor:pointer;">
              <input type="checkbox" id="opt-link" style="width:16px;height:16px;">
              <span style="font-size:0.95rem;color:#374151;">Include link to online quiz</span>
            </label>

            <div style="margin-top:6px;padding-top:12px;border-top:1px solid #eee;">
              <label style="font-size:0.75rem;font-weight:800;color:#6b7280;display:block;margin-bottom:6px;letter-spacing:0.6px;">OUTPUT</label>
              <select id="opt-format" style="width:100%;padding:9px;border:1px solid #ccc;border-radius:6px;background:#fff;font-size:14px;">
                <option value="print">Print Preview (HTML/PDF)</option>
                <option value="word">Microsoft Word (.doc)</option>
              </select>
              <div style="margin-top:8px;font-size:0.78rem;color:#6b7280;">Word export uses Word-compatible HTML inside a .doc file.</div>
            </div>
          </div>

          <div style="padding:14px 20px;background:#f9fafb;border-top:1px solid #e5e7eb;display:flex;justify-content:flex-end;gap:12px;">
            <button id="btn-cancel" style="padding:8px 14px;background:white;border:1px solid #d1d5db;border-radius:8px;color:#374151;font-weight:700;cursor:pointer;">Cancel</button>
            <button id="btn-generate" style="padding:8px 14px;background:#008EE2;border:1px solid #008EE2;border-radius:8px;color:white;font-weight:800;cursor:pointer;">Generate</button>
          </div>
        </div>
      `;

      document.body.appendChild(overlay);

      const close = () => overlay.remove();

      overlay.querySelector('#btn-cancel').onclick = () => {
        close();
        resolve(null);
      };

      overlay.querySelector('#btn-generate').onclick = (e) => {
        const btn = e.currentTarget;
        btn.textContent = 'Working...';
        btn.disabled = true;

        const options = {
          showPoints: overlay.querySelector('#opt-points').checked,
          onePerGroup: overlay.querySelector('#opt-groups').checked,
          showKey: overlay.querySelector('#opt-key').checked,
          showLink: overlay.querySelector('#opt-link').checked,
          shuffleAnswers: overlay.querySelector('#opt-shuffle').checked,
          format: overlay.querySelector('#opt-format').value
        };

        resolve({ options, close });
      };
    });
  }

  /* ---------- Canvas Page Checks ---------- */
  function assertClassicQuizPage() {
    if (!/\/courses\/\d+\/quizzes\/\d+/.test(location.pathname)) {
      throw new Error('Open a Classic Quiz page first: /courses/:course_id/quizzes/:quiz_id');
    }

    const isNewQuizzes = !!document.querySelector('[data-testid="lti-launch-iframe"], iframe[src*="quizzes-next"]');

    if (isNewQuizzes) {
      throw new Error('This script works with Classic Quizzes. New Quizzes uses a different print workflow.');
    }
  }

  function getCourseAndQuizIds() {
    const match = location.pathname.match(/\/courses\/(\d+)\/quizzes\/(\d+)/);

    if (!match) {
      throw new Error('Could not determine the course and quiz IDs.');
    }

    return {
      course_id: match[1],
      quiz_id: match[2]
    };
  }

  /* ---------- Canvas API ---------- */
  async function canvasGet(url) {
    const response = await fetch(url, {
      headers: {
        Accept: 'application/json'
      }
    });

    if (!response.ok) {
      const body = await response.text().catch(() => '');
      throw new Error(`Canvas API error ${response.status}: ${body || response.statusText}`);
    }

    return response.json();
  }

  async function getAllPages(url) {
    const results = [];
    let next = url;

    while (next) {
      const response = await fetch(next, {
        headers: {
          Accept: 'application/json'
        }
      });

      if (!response.ok) {
        const body = await response.text().catch(() => '');
        throw new Error(`Canvas API error ${response.status}: ${body || response.statusText}`);
      }

      const data = await response.json();
      results.push(...data);

      const link = response.headers.get('Link');

      if (link) {
        const nextMatch = link.match(/<([^>]+)>;\s*rel="next"/);
        next = nextMatch ? nextMatch[1] : null;
      } else {
        next = null;
      }
    }

    return results;
  }

  /* ---------- HTML Helpers ---------- */
  function htmlSafe(value) {
    const el = document.createElement('div');
    el.textContent = value ?? '';
    return el.innerHTML;
  }

  function cleanCanvasHTML(html) {
    const wrapper = document.createElement('div');
    wrapper.innerHTML = html || '';

    wrapper.querySelectorAll('script,style,link,iframe').forEach((el) => el.remove());

    return wrapper.innerHTML;
  }

  function isCorrect(answer) {
    return !!(
      answer &&
      (
        answer.is_correct === true ||
        (
          typeof answer.weight === 'number' &&
          answer.weight > 0
        )
      )
    );
  }

  function optionLabel(index) {
    let label = '';
    let n = index;

    do {
      label = String.fromCharCode(65 + (n % 26)) + label;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);

    return label;
  }

  /* ---------- Deterministic Shuffle ---------- */
  function hashStringTo32BitInt(str) {
    let hash = 2166136261;

    for (let i = 0; i < str.length; i++) {
      hash ^= str.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }

    return hash >>> 0;
  }

  function mulberry32(seed) {
    let t = seed >>> 0;

    return function () {
      t += 0x6D2B79F5;

      let r = Math.imul(t ^ (t >>> 15), 1 | t);
      r ^= r + Math.imul(r ^ (r >>> 7), 61 | r);

      return ((r ^ (r >>> 14)) >>> 0) / 4294967296;
    };
  }

  function shuffleArrayDeterministic(array, seedString) {
    const arr = [...array];
    const random = mulberry32(hashStringTo32BitInt(seedString));

    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(random() * (i + 1));
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }

    return arr;
  }

  /* ---------- Question Rendering ---------- */
  function renderAnswerLines(count, mode) {
    const height = mode === 'word' ? 22 : 26;

    return `
      <div class="answer-lines">
        ${Array.from({ length: count })
          .map(() => `<div class="line" style="height:${height}px;"></div>`)
          .join('')}
      </div>
    `;
  }

  function renderLetteredOptions(answers, showKey) {
    const rows = (answers || [])
      .map((answer, index) => {
        const answerHtml = cleanCanvasHTML(answer.html || answer.text || '');
        const correct = showKey && isCorrect(answer);

        return `
          <tr class="${correct ? 'correct' : ''}">
            <td class="option-label">${optionLabel(index)}.</td>
            <td class="option-text">
              ${answerHtml}
              ${correct ? '<span class="inline-key-mark">✓</span>' : ''}
            </td>
          </tr>
        `;
      })
      .join('');

    return `
      <table class="options-table" cellpadding="0" cellspacing="0">
        <tbody>
          ${rows}
        </tbody>
      </table>
    `;
  }

  function renderMatching(answers, showKey, collectKey) {
    const items = (answers || [])
      .map((answer, index) => ({
        left: answer.match_question || answer.left || `Item ${index + 1}`,
        right: answer.text || answer.right || ''
      }));

    if (collectKey) {
      items.forEach((item, index) => {
        collectKey(`Item ${index + 1}`, item.right);
      });
    }

    return `
      <table class="matching" cellpadding="0" cellspacing="0">
        <thead>
          <tr>
            <th>Item</th>
            <th>Match</th>
          </tr>
        </thead>
        <tbody>
          ${items
            .map((item) => `
              <tr>
                <td>${cleanCanvasHTML(item.left)}</td>
                <td>
                  <div class="matching-line"></div>
                  ${showKey && item.right ? `<div class="key-note">(${htmlSafe(item.right)})</div>` : ''}
                </td>
              </tr>
            `)
            .join('')}
        </tbody>
      </table>
    `;
  }

  function renderBlanks(html, answers, showKey, collectKey, mode) {
    let body = html || '';

    const tokens = [...body.matchAll(/\[([^\]]*)\]/g)].map((match) => ({
      raw: match[0],
      id: match[1]
    }));

    const byBlank = {};

    (answers || []).forEach((answer) => {
      if (!answer.blank_id) return;

      byBlank[answer.blank_id] = byBlank[answer.blank_id] || [];
      byBlank[answer.blank_id].push(answer.text || answer.html || '');
    });

    if (tokens.length) {
      tokens.forEach((token) => {
        const values = byBlank[token.id] || [];

        const keyNote = showKey && values.length
          ? `<span class="key-note">(${htmlSafe(values.join(' | '))})</span>`
          : '';

        body = body.replace(token.raw, `<span class="blank-line"></span>${keyNote}`);

        if (collectKey && values.length) {
          collectKey(token.id, values.join(' | '));
        }
      });
    } else {
      body += renderAnswerLines(2, mode);

      const acceptedAnswers = (answers || [])
        .map((answer) => answer.text)
        .filter(Boolean);

      if (showKey && acceptedAnswers.length) {
        body += `<div class="key-note">Answers: ${htmlSafe(acceptedAnswers.join(' | '))}</div>`;

        if (collectKey) {
          collectKey('Answers', acceptedAnswers.join(' | '));
        }
      }
    }

    return body;
  }

  /* ---------- Main Document Builder ---------- */
  function buildDocumentHTML({
    course,
    quiz,
    questions,
    options,
    course_id,
    quiz_id
  }) {
    const {
      showPoints,
      onePerGroup,
      showKey,
      showLink,
      shuffleAnswers,
      format
    } = options;

    const mode = format === 'word' ? 'word' : 'print';

    const today = new Date().toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });

    let questionNumber = 0;
    const usedGroups = new Set();
    const answerKeyRows = [];

    function pushKeyRow(number, value) {
      const text = String(value || '').trim();
      if (!text) return;

      answerKeyRows.push({
        number,
        text
      });
    }

    /* ---------- Optional Online Link ---------- */
    const onlineLink = showLink
      ? `
        <div class="online-link">
          <strong>Online:</strong>
          ${
            mode === 'word'
              ? htmlSafe(location.href)
              : `<a href="${htmlSafe(location.href)}">${htmlSafe(location.href)}</a>`
          }
        </div>
      `
      : '';

    /* ---------- Header ---------- */
    const header = `
      <table class="document-header" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <div class="quiz-title">${htmlSafe(quiz.title || 'Quiz')}</div>

            ${onlineLink}

            <div class="quiz-meta">
              <div><strong>Course:</strong> ${htmlSafe(course.name || '')}</div>
              <div><strong>Total Points:</strong> ${htmlSafe(String(quiz.points_possible ?? ''))}</div>
              <div><strong>Date:</strong> ${htmlSafe(today)}</div>
            </div>

            <table class="student-info" cellpadding="0" cellspacing="0">
              <tr>
                <td>
                  <strong>Name:</strong>
                  <span class="field-line"></span>
                </td>
                <td>
                  <strong>Score:</strong>
                  <span class="field-line short"></span>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    `;

    /* ---------- Instructions ---------- */
    const instructions = quiz.description
      ? `
        <div class="instructions">
          <strong>Instructions:</strong>
          ${cleanCanvasHTML(quiz.description)}
        </div>
      `
      : '';

    /* ---------- Questions ---------- */
    const questionBlocks = questions
      .filter((question) => {
        if (!onePerGroup || !question.quiz_group_id) {
          return true;
        }

        if (usedGroups.has(question.quiz_group_id)) {
          return false;
        }

        usedGroups.add(question.quiz_group_id);
        return true;
      })
      .map((question) => {
        const type = (question.question_type || '').toLowerCase();
        const questionText = cleanCanvasHTML(question.question_text || '');

        /* Text-only Canvas question */
        if (type === 'text_only_question') {
          return `<div class="text-only-block">${questionText}</div>`;
        }

        questionNumber++;

        const points = typeof question.points_possible === 'number'
          ? question.points_possible
          : '';

        const pointsHtml = showPoints
          ? `<div class="question-points">${htmlSafe(String(points))} pts</div>`
          : '';

        let body = '';

        const collectKey = (label, value) => {
          pushKeyRow(questionNumber, `${label}: ${value}`);
        };

        /* ---------- MC / Multiple Answer / True-False ---------- */
        if (
          type.includes('multiple_answers') ||
          type.includes('multiple_choice') ||
          type === 'true_false_question'
        ) {
          let answers = question.answers && question.answers.length
            ? question.answers
            : [];

          if (shuffleAnswers && type !== 'true_false_question') {
            const seed = `${course_id}:${quiz_id}:${question.id || questionNumber}:answers`;
            answers = shuffleArrayDeterministic(answers, seed);
          }

          body = `
            ${questionText}
            ${renderLetteredOptions(answers, showKey)}
          `;

          if (showKey) {
            const correctAnswers = answers
              .map((answer, index) => ({
                answer,
                label: optionLabel(index)
              }))
              .filter(({ answer }) => isCorrect(answer))
              .map(({ answer, label }) => `${label}. ${answer.text || answer.html || ''}`);

            pushKeyRow(
              questionNumber,
              correctAnswers.join(' | ') || 'No correct answer flagged'
            );
          }
        }

        /* ---------- Short Answer ---------- */
        else if (type.includes('short_answer')) {
          body = `
            ${questionText}
            ${renderAnswerLines(4, mode)}
          `;

          if (showKey) {
            const answers = (question.answers || [])
              .map((answer) => answer.text)
              .filter(Boolean);

            pushKeyRow(
              questionNumber,
              answers.join(' | ') || 'Manual grading'
            );
          }
        }

        /* ---------- Essay ---------- */
        else if (type.includes('essay')) {
          body = `
            ${questionText}
            ${renderAnswerLines(12, mode)}
          `;

          if (showKey) {
            pushKeyRow(questionNumber, 'Manual grading');
          }
        }

        /* ---------- Fill in Blank / Dropdown ---------- */
        else if (type.includes('fill_in') || type.includes('dropdown')) {
          body = renderBlanks(
            questionText,
            question.answers,
            showKey,
            collectKey,
            mode
          );
        }

        /* ---------- Matching ---------- */
        else if (type.includes('matching')) {
          const answers = question.answers || [];

          body = `
            ${questionText}
            ${renderMatching(answers, showKey, collectKey)}
          `;
        }

        /* ---------- Numerical ---------- */
        else if (type.includes('numerical')) {
          body = `
            ${questionText}

            <div class="single-answer">
              <strong>Answer:</strong>
              <span class="single-answer-line"></span>
            </div>
          `;

          if (showKey) {
            const values = (question.answers || [])
              .map((answer) => {
                if (answer.exact !== undefined) {
                  return answer.margin
                    ? `${answer.exact} ± ${answer.margin}`
                    : String(answer.exact);
                }

                if (answer.numerical_answer !== undefined) {
                  return String(answer.numerical_answer);
                }

                return answer.text || '';
              })
              .filter(Boolean);

            pushKeyRow(questionNumber, values.join(' | '));
          }
        }

        /* ---------- File Upload ---------- */
        else if (type.includes('file_upload')) {
          body = `
            ${questionText}
            <div class="note-box">File Upload Question</div>
          `;

          if (showKey) {
            pushKeyRow(questionNumber, 'File upload');
          }
        }

        /* ---------- Fallback ---------- */
        else {
          body = `
            ${questionText}
            ${renderAnswerLines(5, mode)}
          `;
        }

        return `
          <table class="question-table" cellpadding="0" cellspacing="0">
            <tr>
              <td class="question-number">
                <div>${questionNumber}.</div>
                ${pointsHtml}
              </td>
              <td class="question-body">
                ${body}
              </td>
            </tr>
          </table>
        `;
      })
      .join('');

    /* ---------- Answer Key ---------- */
    const answerKey = showKey && answerKeyRows.length
      ? `
        <div class="page-break"></div>

        <section class="answer-key">
          <div class="answer-key-title">Answer Key</div>

          <table class="answer-key-table" cellpadding="0" cellspacing="0">
            ${answerKeyRows
              .map((row) => `
                <tr>
                  <td class="answer-key-number">${row.number}.</td>
                  <td class="answer-key-value">${htmlSafe(row.text)}</td>
                </tr>
              `)
              .join('')}
          </table>
        </section>
      `
      : '';

    /* ---------- Word XML ---------- */
    const wordXml = `
      <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
      </xml>
    `;

    /* ---------- CSS ---------- */
    const css = `
      @page {
        margin: 0.65in;
      }

      body {
        margin: 0;
        padding: 18px;
        color: #111;
        line-height: 1.45;
        font-family: Arial, Helvetica, sans-serif;
        font-size: 11pt;
      }

      @media print {
        body {
          padding: 0;
        }

        .screen-only {
          display: none !important;
        }
      }

      .screen-only {
        display: block;
      }

      /* ---------- Preview Toolbar ---------- */
      .preview-toolbar {
        background: #f4f4f4;
        border: 1px solid #d4d4d4;
        border-radius: 6px;
        padding: 9px 11px;
        margin-bottom: 16px;
        font-size: 9.5pt;
        color: #444;
      }

      .preview-toolbar button {
        appearance: none;
        border: 1px solid #bbb;
        background: #fff;
        padding: 6px 12px;
        border-radius: 5px;
        cursor: pointer;
        font-weight: bold;
        margin-right: 10px;
      }

      /* ---------- Header ---------- */
      .document-header {
        width: 100%;
        border-collapse: collapse;
        border-bottom: 2px solid #222;
        padding-bottom: 10px;
        margin-bottom: 16px;
      }

      .quiz-title {
        font-size: 21pt;
        font-weight: bold;
        margin-bottom: 7px;
      }

      .quiz-meta {
        font-size: 9.5pt;
        color: #333;
      }

      .quiz-meta div {
        margin-bottom: 2px;
      }

      .online-link {
        font-size: 9pt;
        margin-bottom: 6px;
        word-break: break-word;
      }

      .student-info {
        width: 100%;
        margin-top: 16px;
      }

      .student-info td {
        width: 50%;
        padding-right: 24px;
        font-size: 10pt;
      }

      .field-line {
        display: inline-block;
        width: 220px;
        border-bottom: 1px solid #222;
        height: 14px;
        margin-left: 7px;
      }

      .field-line.short {
        width: 100px;
      }

      /* ---------- Instructions ---------- */
      .instructions {
        background: #f5f5f5;
        border: 1px solid #ddd;
        padding: 10px 12px;
        margin: 0 0 22px;
        font-size: 10pt;
        line-height: 1.45;
      }

      .instructions p:first-child {
        margin-top: 6px;
      }

      .instructions p:last-child {
        margin-bottom: 4px;
      }

      /* ---------- Text Block ---------- */
      .text-only-block {
        margin: 14px 0 20px;
        padding: 10px 12px;
        border-left: 3px solid #aaa;
        page-break-inside: avoid;
      }

      /* ---------- Question ---------- */
      .question-table {
        width: 100%;
        border-collapse: collapse;
        margin: 0 0 20px;
        page-break-inside: avoid;
      }

      .question-number {
        width: 44px;
        vertical-align: top;
        text-align: right;
        padding-right: 11px;
        font-weight: bold;
        font-size: 11pt;
      }

      .question-points {
        margin-top: 4px;
        color: #666;
        font-size: 8pt;
        font-weight: normal;
        white-space: nowrap;
      }

      .question-body {
        vertical-align: top;
        font-size: 11pt;
      }

      .question-body p:first-child {
        margin-top: 0;
      }

      .question-body img {
        max-width: 100%;
        height: auto;
      }

      /* ---------- Answer Choices ---------- */
      .options-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 9px;
      }

      .options-table tr {
        page-break-inside: avoid;
      }

      .options-table td {
        padding-top: 3px;
        padding-bottom: 3px;
        vertical-align: top;
      }

      .option-label {
        width: 28px;
        padding-right: 5px;
        font-weight: bold;
        white-space: nowrap;
      }

      .option-text {
        padding-left: 0;
      }

      .option-text p {
        margin-top: 0;
        margin-bottom: 0;
      }

      .correct .option-text {
        font-weight: bold;
      }

      .inline-key-mark {
        color: #157a37;
        margin-left: 7px;
        font-weight: bold;
      }

      /* ---------- Written Answer Lines ---------- */
      .answer-lines {
        margin-top: 10px;
      }

      .answer-lines .line {
        border-bottom: 1px solid #bbb;
        margin-bottom: 7px;
      }

      .single-answer {
        margin-top: 12px;
      }

      .single-answer-line {
        display: inline-block;
        min-width: 220px;
        border-bottom: 1px solid #333;
        height: 14px;
        margin-left: 8px;
      }

      .blank-line {
        display: inline-block;
        width: 120px;
        border-bottom: 1px solid #222;
        height: 1em;
        margin: 0 5px;
      }

      /* ---------- Matching ---------- */
      .matching {
        width: 100%;
        border-collapse: collapse;
        margin-top: 10px;
        font-size: 9.5pt;
      }

      .matching th,
      .matching td {
        border: 1px solid #ccc;
        padding: 7px;
        text-align: left;
        vertical-align: top;
      }

      .matching th {
        background: #f3f3f3;
      }

      .matching-line {
        border-bottom: 1px solid #888;
        height: 18px;
      }

      /* ---------- Misc ---------- */
      .note-box {
        border: 1px dashed #999;
        padding: 9px;
        margin-top: 10px;
        color: #555;
        font-size: 9pt;
      }

      .key-note {
        color: #157a37;
        font-size: 9pt;
        font-weight: bold;
        margin-left: 5px;
      }

      /* ---------- Answer Key ---------- */
      .page-break {
        page-break-before: always;
      }

      .answer-key-title {
        font-size: 18pt;
        font-weight: bold;
        border-bottom: 2px solid #222;
        padding-bottom: 7px;
        margin-bottom: 10px;
      }

      .answer-key-table {
        width: 100%;
        border-collapse: collapse;
      }

      .answer-key-table td {
        padding: 6px 0;
        border-bottom: 1px solid #ddd;
        vertical-align: top;
      }

      .answer-key-number {
        width: 38px;
        font-weight: bold;
      }

      .answer-key-value {
        padding-left: 8px;
      }
    `;

    /* ---------- Browser-only Toolbar ---------- */
    const previewToolbar = mode === 'print'
      ? `
        <div class="preview-toolbar screen-only">
          <button onclick="window.print()">Print</button>
          <span>For the cleanest PDF, turn off "Headers and Footers" in the print dialog.</span>
        </div>
      `
      : '';

    return `
      <!DOCTYPE html>

      <html
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:w="urn:schemas-microsoft-com:office:word"
        xmlns="http://www.w3.org/TR/REC-html40"
      >
        <head>
          <meta charset="utf-8">

          <title>${htmlSafe(quiz.title || 'Quiz')}</title>

          ${mode === 'print'
            ? '<meta name="viewport" content="width=device-width,initial-scale=1">'
            : ''}

          ${mode === 'word'
            ? `<!--[if gte mso 9]>${wordXml}<![endif]-->`
            : ''}

          <style>
            ${css}
          </style>
        </head>

        <body>
          ${previewToolbar}
          ${header}
          ${instructions}

          <main>
            ${questionBlocks}
          </main>

          ${answerKey}
        </body>
      </html>
    `;
  }

  /* ---------- Output ---------- */
  function downloadAsWordDoc(filename, html) {
    const blob = new Blob(
      ['\ufeff', html],
      {
        type: 'application/vnd.ms-word;charset=utf-8'
      }
    );

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = `${filename}.doc`;

    document.body.appendChild(link);
    link.click();
    link.remove();

    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function openPreview(html) {
    const preview = window.open('', '_blank');

    if (!preview) {
      throw new Error('Popup blocked. Allow popups for Canvas and try again.');
    }

    preview.document.open();
    preview.document.write(html);
    preview.document.close();
  }

  /* ---------- Run ---------- */
  try {
    assertClassicQuizPage();

    const config = await getPrintOptions();
    if (!config) return;

    const { options, close } = config;

    showSpinner('Fetching quiz data...');

    const { course_id, quiz_id } = getCourseAndQuizIds();

    const [course, quiz] = await Promise.all([
      canvasGet(`/api/v1/courses/${course_id}`),
      canvasGet(`/api/v1/courses/${course_id}/quizzes/${quiz_id}`)
    ]);

    showSpinner('Fetching questions...');

    const questions = await getAllPages(
      `/api/v1/courses/${course_id}/quizzes/${quiz_id}/questions?per_page=100`
    );

    questions.sort(
      (a, b) =>
        (a.position || 0) -
        (b.position || 0)
    );

    showSpinner('Building document...');

    const html = buildDocumentHTML({
      course,
      quiz,
      questions,
      options,
      course_id,
      quiz_id
    });

    close();
    hideSpinner();

    const filename = (quiz.title || 'Quiz')
      .replace(/[^a-z0-9]+/gi, '_')
      .replace(/^_+|_+$/g, '');

    if (options.format === 'word') {
      downloadAsWordDoc(filename || 'Quiz', html);
      toast('Word document downloaded.');
    } else {
      openPreview(html);
      toast('Print preview opened in a new tab.');
    }
  } catch (error) {
    hideSpinner();

    document
      .getElementById('quiz-print-modal')
      ?.remove();

    console.error(error);

    toast(
      `Error: ${error.message}`,
      'error'
    );
  }
})();
