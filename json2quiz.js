(() => {
  // ===========================
  // Canvas Quiz JSON Updater
  // Unpublish -> Apply -> Re-publish
  // ===========================

  // --- CSRF token helper ---
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

  // --- Derive course and quiz ids from URL ---
  const courseMatch = location.pathname.match(/\/courses\/(\d+)/);
  const quizMatch = location.pathname.match(/\/quizzes\/(\d+)/);
  if (!courseMatch || !quizMatch) {
    alert("Could not detect course_id or quiz_id. Please run this on /courses/{id}/quizzes/{quiz_id}.");
    return;
  }
  const COURSE_ID = courseMatch[1];
  const QUIZ_ID = quizMatch[1];
  const BASE = `${location.origin}/api/v1`;
  const HEADERS = {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'X-CSRF-Token': getCsrfToken()
  };

  // --- Small helpers ---
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));

  async function fetchAll(url) {
    let results = [];
    let nextUrl = url;
    while (nextUrl) {
      const res = await fetch(nextUrl, { headers: HEADERS });
      if (!res.ok) throw new Error(`GET failed ${res.status}: ${await res.text()}`);
      const data = await res.json();
      results = results.concat(data);
      const link = res.headers.get('Link');
      if (link && /rel="next"/.test(link)) {
        const m = link.match(/<([^>]+)>;\s*rel="next"/);
        nextUrl = m ? m[1] : null;
      } else {
        nextUrl = null;
      }
    }
    return results;
  }

  async function getQuiz() {
    const res = await fetch(`${BASE}/courses/${COURSE_ID}/quizzes/${QUIZ_ID}`, { headers: HEADERS });
    if (!res.ok) throw new Error(`Failed to get quiz: ${res.status} ${await res.text()}`);
    return res.json();
  }

  async function updateQuiz(body) {
    const res = await fetch(`${BASE}/courses/${COURSE_ID}/quizzes/${QUIZ_ID}`, {
      method: 'PUT',
      headers: HEADERS,
      body: JSON.stringify({ quiz: body })
    });
    if (!res.ok) throw new Error(`Failed to update quiz: ${res.status} ${await res.text()}`);
    return res.json();
  }

  async function getAllQuestions() {
    return fetchAll(`${BASE}/courses/${COURSE_ID}/quizzes/${QUIZ_ID}/questions?per_page=100`);
  }

  async function deleteQuestion(questionId) {
    const res = await fetch(`${BASE}/courses/${COURSE_ID}/quizzes/${QUIZ_ID}/questions/${questionId}`, {
      method: 'DELETE',
      headers: HEADERS
    });
    if (!res.ok) throw new Error(`Failed to delete Q${questionId}: ${res.status} ${await res.text()}`);
    return true;
  }

  async function createQuestion(q) {
    const res = await fetch(`${BASE}/courses/${COURSE_ID}/quizzes/${QUIZ_ID}/questions`, {
      method: 'POST',
      headers: HEADERS,
      body: JSON.stringify({ question: q })
    });
    if (!res.ok) throw new Error(`Failed to create question: ${res.status} ${await res.text()}`);
    return res.json();
  }

  function mapJsonQuestionToCanvas(q, pts) {
    // q.type: "multiple_choice"; allow_multiple: boolean
    const type = (q.type === 'multiple_choice' && q.allow_multiple)
      ? 'multiple_answers_question'
      : 'multiple_choice_question';

    // Canvas answers use weight 100 for correct, 0 for incorrect
    const answers = (q.options || []).map((opt, idx) => ({
      answer_text: opt.text,
      answer_weight: opt.is_correct ? 100 : 0,
      answer_html: opt.text, // keep text in html field too
      // optional: answer_comments, answer_match_left/right not needed here
    }));

    return {
      question_name: q.title || '',
      question_text: q.text || '',
      question_type: type,
      points_possible: Number.isFinite(pts) ? pts : 1,
      answers
    };
  }

  // --- UI overlay ---
  const css = `
  .cqz-overlay{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:999999;}
  .cqz-panel{position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);width:min(900px,90vw);max-height:90vh;overflow:auto;background:#fff;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.3);padding:16px 18px;z-index:1000000;font-family:system-ui,Segoe UI,Roboto,Arial;}
  .cqz-row{display:flex;gap:12px;align-items:center;margin:10px 0;flex-wrap:wrap}
  .cqz-row label{white-space:nowrap;font-weight:600;}
  .cqz-textarea{width:100%;height:300px;font-family:ui-monospace,Consolas,monospace;font-size:13px;padding:8px;border:1px solid #ddd;border-radius:8px;}
  .cqz-btn{padding:8px 12px;border-radius:8px;border:1px solid #ccc;background:#fafafa;cursor:pointer}
  .cqz-btn.primary{background:#008aab;color:#fff;border-color:#008aab}
  .cqz-btn.danger{background:#ea0029;color:#fff;border-color:#ea0029}
  .cqz-footer{display:flex;gap:8px;justify-content:flex-end;margin-top:10px}
  .cqz-summary{background:#f7f7f8;border:1px solid #e6e6e8;border-radius:8px;padding:10px;font-size:13px;white-space:pre-wrap}
  .cqz-spinner{position:fixed;inset:0;display:none;align-items:center;justify-content:center;z-index:1000001;background:rgba(255,255,255,.65)}
  .cqz-spinner.show{display:flex}
  .cqz-spinner .dot{width:12px;height:12px;border-radius:50%;background:#510C76;margin:6px;animation:cqz-bounce 0.6s infinite alternate}
  .cqz-spinner .dot:nth-child(2){animation-delay:.2s}
  .cqz-spinner .dot:nth-child(3){animation-delay:.4s}
  @keyframes cqz-bounce{to{transform:translateY(-10px);opacity:.6}}
  .cqz-toast{position:fixed;top:20px;right:20px;background:#510C76;color:#fff;padding:10px 14px;border-radius:10px;box-shadow:0 8px 20px rgba(0,0,0,.25);z-index:1000002}
  `;
  const style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);

  const overlay = document.createElement('div');
  overlay.className = 'cqz-overlay';
  overlay.addEventListener('click', e => { if (e.target === overlay) document.body.removeChild(overlay); });

  const panel = document.createElement('div');
  panel.className = 'cqz-panel';
  panel.innerHTML = `
    <h2 style="margin:0 0 6px 0">Quiz JSON Updater</h2>
    <div class="cqz-row">
      <label>Course ID:</label><div>${COURSE_ID}</div>
      <label>Quiz ID:</label><div>${QUIZ_ID}</div>
      <label>Mode:</label>
      <div><input type="checkbox" id="cqz-append"> <label for="cqz-append" style="font-weight:500">Append questions (unchecked will replace all existing questions)</label></div>
    </div>
    <textarea class="cqz-textarea" id="cqz-json" placeholder="Paste your JSON here"></textarea>
    <div class="cqz-row">
      <button class="cqz-btn" id="cqz-preview">Preview</button>
      <button class="cqz-btn primary" id="cqz-apply">Unpublish → Apply JSON → Publish</button>
      <button class="cqz-btn danger" id="cqz-close">Close</button>
    </div>
    <div class="cqz-summary" id="cqz-summary" style="display:none"></div>
  `;
  overlay.appendChild(panel);
  document.body.appendChild(overlay);

  const spinner = document.createElement('div');
  spinner.className = 'cqz-spinner';
  spinner.innerHTML = `<div class="dot"></div><div class="dot"></div><div class="dot"></div>`;
  document.body.appendChild(spinner);
  const showSpinner = (on) => spinner.classList.toggle('show', !!on);

  function toast(msg, ms = 3500) {
    const t = document.createElement('div');
    t.className = 'cqz-toast';
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), ms);
  }

  const $json = panel.querySelector('#cqz-json');
  const $append = panel.querySelector('#cqz-append');
  const $summary = panel.querySelector('#cqz-summary');

  panel.querySelector('#cqz-close').onclick = () => overlay.remove();

  panel.querySelector('#cqz-preview').onclick = async () => {
    try {
      const raw = $json.value.trim();
      if (!raw) return alert('Please paste JSON first.');
      const j = JSON.parse(raw);
      const quiz = await getQuiz();
      const existingQs = await getAllQuestions();

      const intended = {
        time_limit_minutes: j.time_limit_minutes,
        attempts_allowed: j.attempts_allowed,
        points_per_question: j.points_per_question,
        total_new_questions: (j.questions || []).length,
        mode: $append.checked ? 'append' : 'replace',
        will_delete_existing: $append.checked ? 0 : existingQs.length,
        existing_questions: existingQs.length,
        current_title_will_remain: quiz.title
      };
      $summary.style.display = 'block';
      $summary.textContent =
        `Preview:\n` +
        `- Title (unchanged): ${intended.current_title_will_remain}\n` +
        `- Time limit: ${intended.time_limit_minutes}\n` +
        `- Attempts allowed: ${intended.attempts_allowed}\n` +
        `- Points per question: ${intended.points_per_question}\n` +
        `- JSON questions: ${intended.total_new_questions}\n` +
        `- Mode: ${intended.mode}\n` +
        `- Existing questions: ${intended.existing_questions}\n` +
        `- Will delete existing: ${intended.will_delete_existing}\n` +
        `- Will unpublish, apply, then publish.`;
    } catch (e) {
      console.error(e);
      alert(`Preview error: ${e.message}`);
    }
  };

  panel.querySelector('#cqz-apply').onclick = async () => {
    const confirmTxt = 'This will unpublish the quiz (if allowed), apply JSON (settings and questions), then publish again.\nContinue?';
    if (!confirm(confirmTxt)) return;

    try {
      const raw = $json.value.trim();
      if (!raw) return alert('Please paste JSON first.');
      const j = JSON.parse(raw);

      // Basic validation
      if (!Array.isArray(j.questions)) throw new Error('JSON "questions" must be an array.');
      const pts = Number(j.points_per_question ?? 1);
      const attempts = Number(j.attempts_allowed ?? 1);
      const timeLimit = Number(j.time_limit_minutes ?? 0);

      showSpinner(true);

      // 1) Unpublish quiz first (ignore title)
      let unpublished = false;
      try {
        await updateQuiz({ published: false });
        unpublished = true;
      } catch (e) {
        // If there are submissions, Canvas will reject unpublishing. Warn but continue.
        console.warn('Unpublish failed. Likely there are submissions. Proceeding while it stays published.', e);
        toast('Could not unpublish. Likely submissions exist. Proceeding anyway.');
      }

      // 2) Update quiz settings (do not send title)
      const quizUpdateBody = {};
      if (Number.isFinite(timeLimit)) quizUpdateBody.time_limit = timeLimit; // in minutes
      if (Number.isFinite(attempts)) quizUpdateBody.allowed_attempts = attempts;

      // Avoid changing title even if quiz_title exists
      await updateQuiz(quizUpdateBody);

      // 3) Replace or append questions
      const existing = await getAllQuestions();

      if (!$append.checked) {
        // Delete all existing questions
        for (const q of existing) {
          await deleteQuestion(q.id);
          await sleep(50); // gentle throttle
        }
      }

      // Create new questions from JSON
      for (const jq of j.questions) {
        const mapped = mapJsonQuestionToCanvas(jq, pts);
        await createQuestion(mapped);
        await sleep(50);
      }

      // 4) Publish again if we successfully unpublished earlier. If we could not unpublish, still try to publish to ensure published=true.
      try {
        await updateQuiz({ published: true });
      } catch (e) {
        console.warn('Re-publish failed.', e);
        toast('Warning: Could not publish the quiz. Please check the quiz status.');
      }

      showSpinner(false);
      toast('Done: JSON applied to quiz.');
      $summary.style.display = 'block';
      $summary.textContent = `Completed.\n- Unpublished attempted: ${unpublished}\n- Settings and questions updated.\n- Publish attempted: true\n\nYou can close this window.`;
    } catch (e) {
      showSpinner(false);
      console.error(e);
      alert(`Error: ${e.message}`);
    }
  };

})();
