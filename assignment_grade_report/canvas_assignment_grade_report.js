(async () => {
  'use strict';

  const GLOBAL_KEY = '__canvasAssignmentGradeReportAnimated';
  const APP_ID = 'canvas-assignment-grade-report-animated';

  const ROOT_ACCOUNT_ID = 1;
  const COURSE_SCOPE_ACCOUNT_IDS = [
    3,
    5
  ];
  // Four parallel course streams is usually a good balance between speed and
  // Canvas's dynamic API throttling. Retries below still handle HTTP 429s.
  const TERM_CONCURRENCY = 4;
  const COURSE_CONCURRENCY = 4;
  const MAX_RETRIES = 4;

  const CSV_HEADERS = [
    'StudentId',
    'StudentName',
    'EnrollmentTerm',
    'CourseCode',
    'CourseName',
    'AssignmentTitle',
    'AssignmentScore',
    'AssignmentPointsPossible',
    'AssignmentPercentageScore',
    'SubmissionStatus'
  ];

  const TEXT_FIELDS = new Set([
    'StudentId',
    'StudentName',
    'EnrollmentTerm',
    'CourseCode',
    'CourseName',
    'AssignmentTitle',
    'SubmissionStatus'
  ]);

  const OLD_GLOBAL_KEYS = [
    GLOBAL_KEY,
    '__canvasAssignmentGradeReportFast',
    '__canvasGradeReportFastFlat',
    '__canvasFastGradeReport',
    '__canvasAssignmentGradeReport'
  ];

  const OLD_APP_IDS = [
    APP_ID,
    'canvas-assignment-grade-report-fast',
    'canvas-grade-report-fast-flat',
    'canvas-fast-grade-report',
    'canvas-grade-report-clean',
    'canvas-assignment-grade-report-app'
  ];

  for (const key of OLD_GLOBAL_KEYS) {
    try {
      window[key]?.close?.(true);
    } catch (_) {
      // Ignore cleanup failures from older script versions.
    }
  }

  for (const id of OLD_APP_IDS) {
    document.getElementById(id)?.remove();
  }

  const previousBodyOverflow = document.body.style.overflow;
  const previousFocus = document.activeElement;

  const state = {
    terms: [],
    selectedTermIds: new Set(),

    courses: [],
    selectedCourseIds: new Set(),

    scopeAccounts: [],
    scopeAccountLoadError: '',
    selectedScopeAccountIds:
      new Set(
        COURSE_SCOPE_ACCOUNT_IDS.map(
          String
        )
      ),

    rows: [],

    controller: null,
    running: false,
    destroyed: false,

    operationName: '',
    operationStartedAt: 0,
    operationTimer: null,

    requestCount: 0,
    currentRequest: '',
    lastProgressAt: 0,

    completedUnits: 0,
    totalUnits: 0,

    activeCourses: new Map(),

    lastSearch: '',

    failedTerms: [],
    failedCourses: [],
    noSubmissionCourses: [],

    downloadedFilename: '',
    objectUrls: new Set(),

    warnings: {
      missingCanvasUserId: 0,
      missingUserName: 0,
      missingAssignment: 0
    }
  };

  class CanvasApiError extends Error {
    constructor(message, status = 0) {
      super(message);
      this.name = 'CanvasApiError';
      this.status = status;
    }
  }

  const host = document.createElement('div');
  host.id = APP_ID;

  const shadow = host.attachShadow({
    mode: 'open'
  });

  shadow.innerHTML = `
<style>
:host {
  all: initial;
}

*,
*::before,
*::after {
  box-sizing: border-box;
}

.overlay {
  position: fixed;
  inset: 0;
  z-index: 2147483647;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 10px;
  background: rgba(0, 0, 0, 0.58);
  color: #222;
  font-family: Arial, Helvetica, sans-serif;
}

.modal {
  width: min(1160px, 98vw);
  height: min(840px, 96vh);
  display: grid;
  grid-template-rows: auto minmax(0, 1fr) auto;
  overflow: hidden;
  background: #fff;
  border-radius: 9px;
  box-shadow: 0 20px 70px rgba(0, 0, 0, 0.45);
}

.header,
.footer {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 11px 14px;
  background: #f7f7f7;
}

.header {
  justify-content: space-between;
  border-bottom: 1px solid #ddd;
}

.footer {
  justify-content: space-between;
  flex-wrap: wrap;
  border-top: 1px solid #ddd;
}

.header h1 {
  margin: 0;
  font-size: 21px;
}

.body {
  min-height: 0;
  overflow: hidden;
  display: grid;
  grid-template-columns:
    minmax(280px, 0.78fr)
    minmax(500px, 1.45fr);
  grid-template-rows:
    minmax(0, 1fr)
    auto;
  gap: 10px;
  padding: 11px;
}

.panel {
  min-width: 0;
  min-height: 0;
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding: 10px;
  border: 1px solid #d0d0d0;
  border-radius: 7px;
  background: #fff;
}

.status-panel {
  grid-column: 1 / -1;
}

.panel h2 {
  margin: 0;
  font-size: 16px;
}

label {
  font-size: 13px;
  font-weight: 700;
}

input[type="text"] {
  width: 100%;
  min-height: 36px;
  padding: 7px 9px;
  border: 1px solid #777;
  border-radius: 5px;
  font: inherit;
}

button {
  min-height: 35px;
  padding: 7px 10px;
  border: 1px solid #666;
  border-radius: 5px;
  background: #f3f3f3;
  color: #111;
  font: inherit;
  font-weight: 700;
  cursor: pointer;
}

button.primary {
  border-color: #18467f;
  background: #245aa8;
  color: #fff;
}

button.danger {
  border-color: #7d2020;
  background: #a83232;
  color: #fff;
}

button:disabled {
  opacity: 0.48;
  cursor: not-allowed;
}

button:focus,
input:focus,
summary:focus {
  outline: 3px solid rgba(36, 90, 168, 0.28);
  outline-offset: 1px;
}

.controls,
.actions {
  display: flex;
  align-items: center;
  gap: 7px;
  flex-wrap: wrap;
}

.count {
  color: #444;
  font-size: 13px;
}

.field-heading {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
}

.field-heading label {
  margin: 0;
}

.scope-account-list {
  display: grid;
  grid-template-columns:
    repeat(2, minmax(0, 1fr));
  gap: 7px;
}

.scope-account-row {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  min-width: 0;
  padding: 8px 9px;
  border: 1px solid #c9d7e8;
  border-radius: 5px;
  background: #f5f8fc;
  font-size: 13px;
  font-weight: 400;
}

.scope-account-row span {
  overflow-wrap: anywhere;
}

.term-list,
.course-table-wrap {
  flex: 1 1 auto;
  min-height: 120px;
  overflow: auto;
  border: 1px solid #c9c9c9;
  border-radius: 5px;
  background: #fafafa;
}

.term-row {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  padding: 7px 8px;
  border-bottom: 1px solid #e3e3e3;
  font-size: 13px;
}

.term-row span,
td {
  overflow-wrap: anywhere;
}

table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
  font-size: 12px;
}

th,
td {
  padding: 7px;
  border-bottom: 1px solid #ddd;
  text-align: left;
  vertical-align: top;
}

th {
  position: sticky;
  top: 0;
  z-index: 1;
  background: #eee;
}

th:nth-child(1),
td:nth-child(1) {
  width: 55px;
  text-align: center;
}

th:nth-child(2),
td:nth-child(2) {
  width: 135px;
}

th:nth-child(3),
td:nth-child(3) {
  width: 185px;
}

th:nth-child(5),
td:nth-child(5) {
  width: 88px;
}

th:nth-child(6),
td:nth-child(6) {
  width: 150px;
}

.empty {
  padding: 20px;
  color: #555;
  text-align: center;
}

.activity-card {
  display: grid;
  grid-template-columns:
    auto
    minmax(0, 1fr)
    auto;
  align-items: center;
  gap: 10px;
  min-height: 58px;
  padding: 9px 10px;
  border: 1px solid #b8cbee;
  border-radius: 6px;
  background: #eef4ff;
}

.spinner {
  width: 24px;
  height: 24px;
  border: 3px solid rgba(36, 90, 168, 0.22);
  border-top-color: #245aa8;
  border-radius: 50%;
  opacity: 0;
  transform: scale(0.9);
}

.spinner.running {
  opacity: 1;
  animation: spin 0.85s linear infinite;
}

.activity-copy {
  min-width: 0;
}

.activity-main {
  font-weight: 700;
  overflow-wrap: anywhere;
}

.activity-sub {
  margin-top: 3px;
  color: #4e5b6d;
  font-size: 12px;
  overflow-wrap: anywhere;
}

.elapsed {
  min-width: 72px;
  color: #36506f;
  font-size: 12px;
  font-variant-numeric: tabular-nums;
  text-align: right;
}

.progress-shell {
  position: relative;
  height: 14px;
  overflow: hidden;
  border: 1px solid #9fb5d5;
  border-radius: 999px;
  background: #dfe9f7;
}

.progress-fill {
  position: absolute;
  inset: 0 auto 0 0;
  width: 0%;
  border-radius: inherit;
  background: linear-gradient(
    90deg,
    #245aa8,
    #4a82cf
  );
  transition: width 0.35s ease;
}

.progress-shell.running::after {
  content: '';
  position: absolute;
  top: 0;
  bottom: 0;
  width: 34%;
  background: linear-gradient(
    90deg,
    rgba(255, 255, 255, 0),
    rgba(255, 255, 255, 0.7),
    rgba(255, 255, 255, 0)
  );
  animation: sweep 1.25s ease-in-out infinite;
}

.progress-shell.indeterminate .progress-fill {
  width: 38%;
  animation: indeterminate 1.35s ease-in-out infinite;
}

.progress-labels {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  color: #536172;
  font-size: 11px;
}

.heartbeat {
  display: inline-flex;
  align-items: center;
  gap: 6px;
}

.heartbeat-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: #6f7d8d;
}

.heartbeat-dot.running {
  background: #2b7a35;
  animation: pulse 1.2s ease-in-out infinite;
}

.metrics {
  display: grid;
  grid-template-columns:
    repeat(5, minmax(110px, 1fr));
  gap: 7px;
}

.metric {
  padding: 7px 9px;
  border-radius: 5px;
  background: #f7f7f7;
}

.metric b {
  display: block;
  margin-bottom: 2px;
  color: #555;
  font-size: 12px;
}

.metric span {
  font-size: 15px;
  font-weight: 700;
}

details {
  padding-top: 7px;
  border-top: 1px solid #ddd;
}

summary {
  cursor: pointer;
  font-weight: 700;
}

.detail-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 9px;
  margin-top: 8px;
}

.log,
.summary {
  max-height: 135px;
  overflow: auto;
  padding: 8px;
  border: 1px solid #d4d4d4;
  border-radius: 5px;
  background: #fafafa;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  font-size: 12px;
  line-height: 1.4;
}

@keyframes spin {
  to {
    transform: rotate(360deg);
  }
}

@keyframes sweep {
  from {
    left: -42%;
  }

  to {
    left: 108%;
  }
}

@keyframes indeterminate {
  0% {
    left: -42%;
  }

  50% {
    left: 52%;
  }

  100% {
    left: 108%;
  }
}

@keyframes pulse {
  0%,
  100% {
    transform: scale(0.85);
    opacity: 0.55;
  }

  50% {
    transform: scale(1.25);
    opacity: 1;
  }
}

@media (prefers-reduced-motion: reduce) {
  .spinner.running,
  .progress-shell.running::after,
  .progress-shell.indeterminate .progress-fill,
  .heartbeat-dot.running {
    animation-duration: 2.5s;
  }
}

@media (max-width: 860px) {
  .overlay {
    padding: 5px;
  }

  .modal {
    width: 99vw;
    height: 98vh;
  }

  .body {
    grid-template-columns: 1fr;
    grid-template-rows:
      minmax(235px, 0.8fr)
      minmax(280px, 1fr)
      auto;
    overflow: auto;
  }

  .status-panel {
    grid-column: 1;
  }

  .metrics {
    grid-template-columns:
      repeat(2, minmax(120px, 1fr));
  }

  .detail-grid {
    grid-template-columns: 1fr;
  }

  .scope-account-list {
    grid-template-columns: 1fr;
  }

  .activity-card {
    grid-template-columns:
      auto
      minmax(0, 1fr);
  }

  .elapsed {
    grid-column: 2;
    text-align: left;
  }
}
</style>

<div class="overlay">
  <section
    class="modal"
    role="dialog"
    aria-modal="true"
    aria-labelledby="grade-report-title"
  >
    <header class="header">
      <h1 id="grade-report-title">
        Canvas Assignment Grade Report
      </h1>

      <button
        id="topClose"
        type="button"
        aria-label="Close"
      >
        ×
      </button>
    </header>

    <main class="body">
      <section class="panel">
        <h2>1. Choose scope</h2>

        <label for="courseCode">
          Course code contains
        </label>

        <input
          id="courseCode"
          type="text"
          placeholder="Example: COU67380"
          autocomplete="off"
        >

        <div class="field-heading">
          <label>
            Search within subaccounts
          </label>

          <span
            id="scopeAccountCount"
            class="count"
          ></span>
        </div>

        <div
          id="scopeAccountList"
          class="scope-account-list"
          role="group"
          aria-label="Course subaccounts"
        >
          <div class="empty">
            Loading account names…
          </div>
        </div>

        <label for="termSearch">
          Filter terms
        </label>

        <input
          id="termSearch"
          type="text"
          placeholder="Search term names"
          autocomplete="off"
        >

        <div class="controls">
          <button
            id="selectVisibleTerms"
            type="button"
          >
            Select Visible
          </button>

          <button
            id="clearTerms"
            type="button"
          >
            Clear Terms
          </button>

          <button
            id="reloadTerms"
            type="button"
          >
            Reload
          </button>

          <span
            id="termCount"
            class="count"
          ></span>
        </div>

        <div
          id="termList"
          class="term-list"
          role="group"
          aria-label="Enrollment terms"
        ></div>
      </section>

      <section class="panel">
        <h2>2. Review courses</h2>

        <div class="controls">
          <button
            id="selectAllCourses"
            type="button"
          >
            Select All
          </button>

          <button
            id="clearCourses"
            type="button"
          >
            Clear Courses
          </button>

          <span
            id="courseCount"
            class="count"
          ></span>
        </div>

        <div class="course-table-wrap">
          <table aria-label="Matched courses">
            <thead>
              <tr>
                <th>Use</th>
                <th>Term</th>
                <th>Course code</th>
                <th>Course name</th>
                <th>ID</th>
                <th>Account</th>
              </tr>
            </thead>

            <tbody id="courseTableBody"></tbody>
          </table>

          <div
            id="courseEmpty"
            class="empty"
          >
            Select terms and click Find Courses.
          </div>
        </div>
      </section>

      <section class="panel status-panel">
        <div class="activity-card">
          <div
            id="spinner"
            class="spinner"
            aria-hidden="true"
          ></div>

          <div class="activity-copy">
            <div
              id="activityMain"
              class="activity-main"
              role="status"
              aria-live="polite"
            >
              Loading terms…
            </div>

            <div
              id="activitySub"
              class="activity-sub"
            >
              Preparing the report tool.
            </div>
          </div>

          <div
            id="elapsed"
            class="elapsed"
          >
            00:00
          </div>
        </div>

        <div
          id="progressShell"
          class="progress-shell"
          role="progressbar"
          aria-valuemin="0"
          aria-valuemax="100"
          aria-valuenow="0"
          aria-label="Operation progress"
        >
          <div
            id="progressFill"
            class="progress-fill"
          ></div>
        </div>

        <div class="progress-labels">
          <span id="progressText">
            Waiting to start
          </span>

          <span class="heartbeat">
            <span
              id="heartbeatDot"
              class="heartbeat-dot"
            ></span>

            <span id="heartbeatText">
              Idle
            </span>
          </span>
        </div>

        <div class="metrics">
          <div class="metric">
            <b>Terms selected</b>
            <span id="metricTerms">0</span>
          </div>

          <div class="metric">
            <b>Courses selected</b>
            <span id="metricSelectedCourses">0</span>
          </div>

          <div class="metric">
            <b>Courses completed</b>
            <span id="metricCompletedCourses">0</span>
          </div>

          <div class="metric">
            <b>Report rows</b>
            <span id="metricRows">0</span>
          </div>

          <div class="metric">
            <b>API requests</b>
            <span id="metricRequests">0</span>
          </div>
        </div>

        <details id="details">
          <summary>
            Details, warnings, and final summary
          </summary>

          <div class="detail-grid">
            <div>
              <b>Status log</b>
              <div
                id="log"
                class="log"
              ></div>
            </div>

            <div>
              <b>Final summary</b>

              <div
                id="summary"
                class="summary"
              >
                No report has been generated.
              </div>
            </div>
          </div>
        </details>
      </section>
    </main>

    <footer class="footer">
      <div class="actions">
        <button
          id="findCourses"
          class="primary"
          type="button"
        >
          Find Courses
        </button>

        <button
          id="generateReport"
          class="primary"
          type="button"
        >
          Generate Report
        </button>

        <button
          id="cancel"
          class="danger"
          type="button"
        >
          Cancel
        </button>
      </div>

      <div class="actions">
        <button
          id="downloadPartial"
          type="button"
          hidden
        >
          Download Partial Results
        </button>

        <button
          id="downloadEmpty"
          type="button"
          hidden
        >
          Download Empty CSV
        </button>

        <button
          id="close"
          type="button"
        >
          Close
        </button>
      </div>
    </footer>
  </section>
</div>
`;

  document.body.appendChild(host);
  document.body.style.overflow = 'hidden';

  const getElement = id =>
    shadow.getElementById(id);

  const elements = {
    courseCode: getElement('courseCode'),
    scopeAccountList:
      getElement('scopeAccountList'),
    scopeAccountCount:
      getElement('scopeAccountCount'),
    termSearch: getElement('termSearch'),
    termList: getElement('termList'),
    termCount: getElement('termCount'),

    courseTableBody: getElement('courseTableBody'),
    courseEmpty: getElement('courseEmpty'),
    courseCount: getElement('courseCount'),

    activityMain: getElement('activityMain'),
    activitySub: getElement('activitySub'),
    elapsed: getElement('elapsed'),

    spinner: getElement('spinner'),

    progressShell: getElement('progressShell'),
    progressFill: getElement('progressFill'),
    progressText: getElement('progressText'),

    heartbeatDot: getElement('heartbeatDot'),
    heartbeatText: getElement('heartbeatText'),

    metricTerms: getElement('metricTerms'),
    metricSelectedCourses:
      getElement('metricSelectedCourses'),
    metricCompletedCourses:
      getElement('metricCompletedCourses'),
    metricRows: getElement('metricRows'),
    metricRequests: getElement('metricRequests'),

    details: getElement('details'),
    log: getElement('log'),
    summary: getElement('summary'),

    selectVisibleTerms:
      getElement('selectVisibleTerms'),
    clearTerms: getElement('clearTerms'),
    reloadTerms: getElement('reloadTerms'),

    selectAllCourses:
      getElement('selectAllCourses'),
    clearCourses: getElement('clearCourses'),

    findCourses: getElement('findCourses'),
    generateReport: getElement('generateReport'),
    cancel: getElement('cancel'),

    downloadPartial:
      getElement('downloadPartial'),
    downloadEmpty:
      getElement('downloadEmpty'),

    close: getElement('close'),
    topClose: getElement('topClose')
  };

  function stringValue(value) {
    return value == null
      ? ''
      : String(value);
  }

  function compareText(a, b) {
    return stringValue(a).localeCompare(
      stringValue(b),
      undefined,
      {
        sensitivity: 'base',
        numeric: true
      }
    );
  }

  function isAbortError(error) {
    return error?.name === 'AbortError';
  }

  function ensureNotAborted(signal) {
    if (signal?.aborted) {
      throw new DOMException(
        'Canceled',
        'AbortError'
      );
    }
  }

  function formatElapsed(milliseconds) {
    const totalSeconds = Math.max(
      0,
      Math.floor(milliseconds / 1000)
    );

    const minutes =
      Math.floor(totalSeconds / 60);

    const seconds =
      totalSeconds % 60;

    return (
      `${String(minutes).padStart(2, '0')}:` +
      `${String(seconds).padStart(2, '0')}`
    );
  }

  function setActivity(main, sub = '') {
    elements.activityMain.textContent = main;
    elements.activitySub.textContent = sub;
  }

  function addLog(message) {
    const line =
      document.createElement('div');

    line.textContent = message;

    elements.log.appendChild(line);

    elements.log.scrollTop =
      elements.log.scrollHeight;
  }

  function clearLog() {
    elements.log.replaceChildren();
  }

  function startOperation(
    name,
    totalUnits = 0
  ) {
    state.operationName = name;

    state.operationStartedAt =
      Date.now();

    state.lastProgressAt =
      Date.now();

    state.requestCount = 0;
    state.currentRequest = '';

    state.completedUnits = 0;
    state.totalUnits =
      Math.max(0, totalUnits);

    clearInterval(
      state.operationTimer
    );

    state.operationTimer =
      window.setInterval(
        () => {
          if (!state.running) {
            return;
          }

          const elapsedMs =
            Date.now() -
            state.operationStartedAt;

          const quietSeconds =
            Math.floor(
              (
                Date.now() -
                state.lastProgressAt
              ) /
              1000
            );

          elements.elapsed.textContent =
            formatElapsed(elapsedMs);

          elements.heartbeatText.textContent =
            quietSeconds >= 10
              ? (
                  `Still working, ` +
                  `${quietSeconds}s since last response`
                )
              : 'Working';

          const requestText =
            state.currentRequest
              ? (
                  `Current request: ` +
                  state.currentRequest
                )
              : (
                  'Waiting for Canvas ' +
                  'to respond.'
                );

          elements.activitySub.textContent =
            `${requestText} ` +
            `Elapsed ${formatElapsed(elapsedMs)}.`;
        },
        1000
      );

    updateProgressDisplay();
  }

  function stopOperation() {
    clearInterval(
      state.operationTimer
    );

    state.operationTimer = null;
    state.currentRequest = '';
    state.operationName = '';

    elements.spinner.classList.remove(
      'running'
    );

    elements.progressShell.classList.remove(
      'running',
      'indeterminate'
    );

    elements.heartbeatDot.classList.remove(
      'running'
    );

    elements.heartbeatText.textContent =
      'Idle';
  }

  function markProgress() {
    state.lastProgressAt =
      Date.now();
  }

  function updateProgressDisplay() {
    const hasTotal =
      state.totalUnits > 0;

    const ratio =
      hasTotal
        ? Math.min(
            1,
            state.completedUnits /
            state.totalUnits
          )
        : 0;

    const percent =
      Math.round(ratio * 100);

    elements.spinner.classList.toggle(
      'running',
      state.running
    );

    elements.progressShell.classList.toggle(
      'running',
      state.running
    );

    elements.progressShell.classList.toggle(
      'indeterminate',
      state.running && !hasTotal
    );

    elements.heartbeatDot.classList.toggle(
      'running',
      state.running
    );

    if (hasTotal) {
      elements.progressFill.style.width =
        `${percent}%`;

      elements.progressText.textContent =
        `${state.completedUnits} of ` +
        `${state.totalUnits} complete ` +
        `(${percent}%)`;

      elements.progressShell.setAttribute(
        'aria-valuenow',
        String(percent)
      );
    } else {
      elements.progressFill.style.width =
        '38%';

      elements.progressText.textContent =
        state.running
          ? (
              'Working. Canvas has not ' +
              'returned a measurable step yet.'
            )
          : 'Waiting to start';

      elements.progressShell.setAttribute(
        'aria-valuenow',
        '0'
      );
    }

    elements.metricRequests.textContent =
      state.requestCount;
  }

  function setMeasuredProgress(
    completed,
    total
  ) {
    state.completedUnits =
      Math.max(0, completed);

    state.totalUnits =
      Math.max(0, total);

    markProgress();
    updateProgressDisplay();
  }

  function updateCounts() {
    elements.scopeAccountCount.textContent =
      `${state.selectedScopeAccountIds.size} account` +
      `${
        state.selectedScopeAccountIds.size === 1
          ? ''
          : 's'
      } selected`;

    elements.termCount.textContent =
      `${state.selectedTermIds.size} term` +
      `${
        state.selectedTermIds.size === 1
          ? ''
          : 's'
      } selected`;

    elements.courseCount.textContent =
      `${state.selectedCourseIds.size} course` +
      `${
        state.selectedCourseIds.size === 1
          ? ''
          : 's'
      } selected`;

    elements.metricTerms.textContent =
      state.selectedTermIds.size;

    elements.metricSelectedCourses.textContent =
      state.selectedCourseIds.size;

    elements.metricCompletedCourses.textContent =
      state.totalUnits
        ? (
            `${state.completedUnits}/` +
            `${state.totalUnits}`
          )
        : '0';

    elements.metricRows.textContent =
      state.rows.length;

    elements.metricRequests.textContent =
      state.requestCount;

    elements.findCourses.disabled =
      state.running ||
      elements.courseCode.value
        .trim()
        .length < 3 ||
      state.selectedScopeAccountIds.size === 0 ||
      state.selectedTermIds.size === 0;

    elements.generateReport.disabled =
      state.running ||
      state.selectedCourseIds.size === 0;

    elements.cancel.disabled =
      !state.running;
  }

  function setRunning(running) {
    state.running = running;

    [
      elements.courseCode,
      elements.termSearch,
      elements.selectVisibleTerms,
      elements.clearTerms,
      elements.reloadTerms,
      elements.selectAllCourses,
      elements.clearCourses
    ].forEach(element => {
      element.disabled = running;
    });

    shadow
      .querySelectorAll(
        'input[type="checkbox"]'
      )
      .forEach(element => {
        element.disabled = running;
      });

    if (!running) {
      stopOperation();
    }

    updateCounts();
    updateProgressDisplay();
  }

  function parseLinkHeader(headerValue) {
    const links = {};

    if (!headerValue) {
      return links;
    }

    const regex =
      /<([^>]+)>\s*;\s*rel=(?:"([^"]+)"|([^,;\s]+))/g;

    let match;

    while (
      (match = regex.exec(headerValue))
    ) {
      const relationships =
        (match[2] || match[3] || '')
          .split(/\s+/)
          .filter(Boolean);

      for (
        const relationship
        of relationships
      ) {
        links[relationship] =
          match[1];
      }
    }

    return links;
  }

  function parseRetryAfter(value) {
    if (!value) {
      return null;
    }

    const seconds =
      Number(value);

    if (
      Number.isFinite(seconds) &&
      seconds >= 0
    ) {
      return seconds * 1000;
    }

    const date =
      Date.parse(value);

    return Number.isFinite(date)
      ? Math.max(
          0,
          date - Date.now()
        )
      : null;
  }

  function delay(
    milliseconds,
    signal
  ) {
    return new Promise(
      (resolve, reject) => {
        ensureNotAborted(signal);

        const timer =
          setTimeout(
            () => {
              cleanup();
              resolve();
            },
            milliseconds
          );

        const onAbort = () => {
          clearTimeout(timer);
          cleanup();

          reject(
            new DOMException(
              'Canceled',
              'AbortError'
            )
          );
        };

        const cleanup = () => {
          signal?.removeEventListener(
            'abort',
            onAbort
          );
        };

        signal?.addEventListener(
          'abort',
          onAbort,
          {
            once: true
          }
        );
      }
    );
  }

  async function fetchJson(
    input,
    {
      signal,
      label = 'Canvas data'
    } = {}
  ) {
    const url =
      new URL(
        input,
        window.location.origin
      );

    if (
      url.origin !==
      window.location.origin
    ) {
      throw new CanvasApiError(
        'Canvas returned a cross-origin pagination URL.'
      );
    }

    for (
      let attempt = 0;
      attempt <= MAX_RETRIES;
      attempt += 1
    ) {
      ensureNotAborted(signal);

      state.requestCount += 1;
      state.currentRequest = label;

      updateProgressDisplay();

      let response;

      try {
        response =
          await fetch(
            url.href,
            {
              method: 'GET',
              credentials:
                'same-origin',
              cache: 'no-store',
              headers: {
                Accept:
                  'application/json'
              },
              signal
            }
          );
      } catch (error) {
        if (isAbortError(error)) {
          throw error;
        }

        if (
          attempt ===
          MAX_RETRIES
        ) {
          throw new CanvasApiError(
            `Network error: ${error.message}`
          );
        }

        const waitTime =
          Math.min(
            1000 * 2 ** attempt,
            12000
          ) +
          Math.random() * 500;

        setActivity(
          `Retrying ${label}`,
          `Network error. Retrying in about ` +
          `${Math.ceil(waitTime / 1000)} seconds.`
        );

        addLog(
          `Retrying ${label} after a network error.`
        );

        await delay(
          waitTime,
          signal
        );

        continue;
      }

      if (!response.ok) {
        const bodyText =
          await response
            .text()
            .catch(() => '');

        const retryable =
          response.status === 429 ||
          response.status >= 500;

        if (
          retryable &&
          attempt < MAX_RETRIES
        ) {
          const waitTime =
            (
              parseRetryAfter(
                response.headers.get(
                  'Retry-After'
                )
              ) ??
              Math.min(
                1500 * 2 ** attempt,
                18000
              )
            ) +
            Math.random() * 700;

          setActivity(
            `Retrying ${label}`,
            `Canvas returned HTTP ` +
            `${response.status}. ` +
            `Retrying in about ` +
            `${Math.ceil(waitTime / 1000)} seconds.`
          );

          addLog(
            `Retrying ${label} after ` +
            `HTTP ${response.status}.`
          );

          await delay(
            waitTime,
            signal
          );

          continue;
        }

        let message =
          bodyText
            .replace(
              /<[^>]+>/g,
              ' '
            )
            .replace(
              /\s+/g,
              ' '
            )
            .trim();

        try {
          const parsed =
            JSON.parse(bodyText);

          message =
            parsed?.message ||
            parsed?.errors?.[0]?.message ||
            message;
        } catch (_) {
          // Use cleaned response text.
        }

        throw new CanvasApiError(
          `HTTP ${response.status}: ` +
          `${message || response.statusText}`,
          response.status
        );
      }

      try {
        const raw =
          await response.text();

        markProgress();

        return {
          response,
          body:
            raw
              ? JSON.parse(raw)
              : null
        };
      } catch (_) {
        throw new CanvasApiError(
          'Canvas returned invalid JSON.',
          response.status
        );
      }
    }

    throw new CanvasApiError(
      'Canvas request failed.'
    );
  }

  async function fetchAllPages(
    initialUrl,
    {
      signal,
      label,
      extractItems = body => body,
      onPage
    } = {}
  ) {
    const combined = [];

    let nextUrl =
      initialUrl;

    let pageNumber = 0;

    while (nextUrl) {
      const {
        response,
        body
      } =
        await fetchJson(
          nextUrl,
          {
            signal,
            label
          }
        );

      const items =
        extractItems(body);

      if (!Array.isArray(items)) {
        throw new CanvasApiError(
          'Canvas returned an unexpected response structure.'
        );
      }

      pageNumber += 1;

      combined.push(...items);

      markProgress();

      onPage?.({
        pageNumber,
        pageCount:
          items.length,
        totalCount:
          combined.length
      });

      nextUrl =
        parseLinkHeader(
          response.headers.get(
            'Link'
          )
        ).next || '';
    }

    return combined;
  }

  function createApiUrl(
    path,
    configure
  ) {
    const url =
      new URL(
        path,
        window.location.origin
      );

    configure?.(
      url.searchParams
    );

    return url.href;
  }

  async function runPool(
    items,
    limit,
    worker,
    signal
  ) {
    let nextIndex = 0;

    async function runner() {
      while (true) {
        ensureNotAborted(signal);

        const currentIndex =
          nextIndex++;

        if (
          currentIndex >=
          items.length
        ) {
          return;
        }

        await worker(
          items[currentIndex],
          currentIndex
        );
      }
    }

    await Promise.all(
      Array.from(
        {
          length: Math.min(
            limit,
            items.length
          )
        },
        runner
      )
    );
  }

  function getEffectiveTermStart(term) {
    const candidates = [
      term?.overrides
        ?.StudentEnrollment
        ?.start_at,

      term?.start_at
    ];

    for (
      const value
      of candidates
    ) {
      const parsed =
        Date.parse(value);

      if (
        Number.isFinite(parsed)
      ) {
        return parsed;
      }
    }

    return null;
  }

  function renderScopeAccounts() {
    elements.scopeAccountList
      .replaceChildren();

    if (!state.scopeAccounts.length) {
      const empty =
        document.createElement(
          'div'
        );

      empty.className =
        'empty';

      empty.textContent =
        state.scopeAccountLoadError ||
        'Loading account names…';

      elements.scopeAccountList
        .appendChild(empty);

      updateCounts();
      return;
    }

    for (
      const account
      of state.scopeAccounts
    ) {
      const accountId =
        String(account.id);

      const label =
        document.createElement(
          'label'
        );

      label.className =
        'scope-account-row';

      const checkbox =
        document.createElement(
          'input'
        );

      checkbox.type =
        'checkbox';

      checkbox.checked =
        state.selectedScopeAccountIds.has(
          accountId
        );

      checkbox.disabled =
        state.running;

      const textSpan =
        document.createElement(
          'span'
        );

      textSpan.textContent =
        `${account.name || 'Unnamed account'} ` +
        `(ID: ${account.id})`;

      checkbox.addEventListener(
        'change',
        () => {
          if (checkbox.checked) {
            state.selectedScopeAccountIds.add(
              accountId
            );
          } else {
            state.selectedScopeAccountIds.delete(
              accountId
            );
          }

          clearCourseResults(
            'Subaccount selection changed. ' +
            'Click Find Courses again.'
          );

          updateCounts();
        }
      );

      label.append(
        checkbox,
        textSpan
      );

      elements.scopeAccountList
        .appendChild(label);
    }

    updateCounts();
  }

  function renderTerms() {
    elements.termList
      .replaceChildren();

    const filter =
      elements.termSearch.value
        .trim()
        .toLowerCase();

    let visibleCount = 0;

    for (
      const term
      of state.terms
    ) {
      if (
        filter &&
        !stringValue(term.name)
          .toLowerCase()
          .includes(filter)
      ) {
        continue;
      }

      visibleCount += 1;

      const label =
        document.createElement(
          'label'
        );

      label.className =
        'term-row';

      const checkbox =
        document.createElement(
          'input'
        );

      checkbox.type =
        'checkbox';

      checkbox.checked =
        state.selectedTermIds.has(
          String(term.id)
        );

      checkbox.disabled =
        state.running;

      const textSpan =
        document.createElement(
          'span'
        );

      textSpan.textContent =
        `${term.name || 'Unnamed term'} ` +
        `(ID: ${term.id})`;

      checkbox.addEventListener(
        'change',
        () => {
          const termId =
            String(term.id);

          if (
            checkbox.checked
          ) {
            state.selectedTermIds.add(
              termId
            );
          } else {
            state.selectedTermIds.delete(
              termId
            );
          }

          clearCourseResults(
            'Term selection changed. ' +
            'Click Find Courses again.'
          );

          updateCounts();
        }
      );

      label.append(
        checkbox,
        textSpan
      );

      elements.termList.appendChild(
        label
      );
    }

    if (!visibleCount) {
      const empty =
        document.createElement(
          'div'
        );

      empty.className =
        'empty';

      empty.textContent =
        state.terms.length
          ? 'No terms match this filter.'
          : 'No eligible terms found.';

      elements.termList.appendChild(
        empty
      );
    }

    updateCounts();
  }

  function renderCourses() {
    elements.courseTableBody
      .replaceChildren();

    for (
      const course
      of state.courses
    ) {
      const row =
        document.createElement(
          'tr'
        );

      const checkboxCell =
        document.createElement(
          'td'
        );

      const checkbox =
        document.createElement(
          'input'
        );

      checkbox.type =
        'checkbox';

      checkbox.checked =
        state.selectedCourseIds.has(
          String(course.id)
        );

      checkbox.disabled =
        state.running;

      checkbox.setAttribute(
        'aria-label',
        `Include ${
          course.course_code ||
          course.name ||
          course.id
        }`
      );

      checkbox.addEventListener(
        'change',
        () => {
          const courseId =
            String(course.id);

          if (
            checkbox.checked
          ) {
            state.selectedCourseIds.add(
              courseId
            );
          } else {
            state.selectedCourseIds.delete(
              courseId
            );
          }

          updateCounts();
        }
      );

      checkboxCell.appendChild(
        checkbox
      );

      row.appendChild(
        checkboxCell
      );

      const values = [
        course._termName,
        course.course_code,
        course.name,
        course.id,
        course.account_name
      ];

      for (
        const value
        of values
      ) {
        const cell =
          document.createElement(
            'td'
          );

        cell.textContent =
          stringValue(value);

        row.appendChild(
          cell
        );
      }

      elements.courseTableBody.appendChild(
        row
      );
    }

    elements.courseEmpty.hidden =
      state.courses.length > 0;

    updateCounts();
  }

  function clearCourseResults(
    message = ''
  ) {
    if (state.running) {
      return;
    }

    state.courses = [];

    state.selectedCourseIds
      .clear();

    state.rows = [];

    state.completedUnits = 0;
    state.totalUnits = 0;

    renderCourses();

    if (message) {
      setActivity(
        message,
        'No report operation is currently running.'
      );
    }
  }

  async function loadScopeAccounts(
    signal
  ) {
    return Promise.all(
      COURSE_SCOPE_ACCOUNT_IDS.map(
        async configuredAccountId => {
          const accountUrl =
            createApiUrl(
              `/api/v1/accounts/${configuredAccountId}`
            );

          const descendantsUrl =
            createApiUrl(
              `/api/v1/accounts/${configuredAccountId}/sub_accounts`,
              params => {
                params.set(
                  'recursive',
                  'true'
                );

                params.set(
                  'per_page',
                  '100'
                );
              }
            );

          const [
            accountResult,
            descendants
          ] =
            await Promise.all([
              fetchJson(
                accountUrl,
                {
                  signal,

                  label:
                    `account ${configuredAccountId}`
                }
              ),

              fetchAllPages(
                descendantsUrl,
                {
                  signal,

                  label:
                    `subaccounts under ${configuredAccountId}`,

                  onPage:
                    ({
                      pageNumber,
                      totalCount
                    }) => {
                      setActivity(
                        `Reading account ${configuredAccountId}`,
                        `Received descendant page ${pageNumber}. ` +
                        `${totalCount} subaccounts found so far.`
                      );
                    }
                }
              )
            ]);

          const account =
            accountResult.body || {};

          const resolvedId =
            account.id ??
            configuredAccountId;

          return {
            id:
              resolvedId,

            name:
              account.name ||
              `Account ${resolvedId}`,

            accountIds: [
              resolvedId,

              ...descendants.map(
                descendant =>
                  descendant.id
              )
            ]
          };
        }
      )
    );
  }

  async function loadTerms() {
    if (state.running) {
      return;
    }

    state.controller =
      new AbortController();

    setRunning(true);

    startOperation(
      'Loading terms'
    );

    clearLog();

    state.scopeAccountLoadError = '';

    setActivity(
      'Loading enrollment terms',
      'The activity indicator will continue moving ' +
      'while Canvas prepares the response.'
    );

    try {
      const url =
        createApiUrl(
          `/api/v1/accounts/${ROOT_ACCOUNT_ID}/terms`,
          params => {
            params.set(
              'per_page',
              '100'
            );

            params.append(
              'workflow_state[]',
              'active'
            );

            params.append(
              'include[]',
              'overrides'
            );
          }
        );

      const [
        returnedTerms,
        returnedScopeAccounts
      ] =
        await Promise.all([
          fetchAllPages(
            url,
            {
              signal:
                state.controller.signal,

              label:
                'enrollment terms',

              extractItems:
                body =>
                  body?.enrollment_terms,

              onPage:
                ({
                  pageNumber,
                  totalCount
                }) => {
                  setActivity(
                    'Loading enrollment terms',
                    `Received page ${pageNumber}. ` +
                    `${totalCount} terms received so far.`
                  );
                }
            }
          ),

          loadScopeAccounts(
            state.controller.signal
          )
        ]);

      state.scopeAccounts =
        returnedScopeAccounts;

      const validScopeAccountIds =
        new Set(
          state.scopeAccounts.map(
            account =>
              String(account.id)
          )
        );

      state.selectedScopeAccountIds =
        new Set(
          [...state.selectedScopeAccountIds]
            .filter(
              accountId =>
                validScopeAccountIds.has(
                  accountId
                )
            )
        );

      renderScopeAccounts();

      const now =
        Date.now();

      let excludedFutureTerms = 0;

      state.terms =
        returnedTerms
          .filter(
            term =>
              term &&
              term.workflow_state !==
                'deleted'
          )
          .filter(
            term => {
              const effectiveStart =
                getEffectiveTermStart(
                  term
                );

              if (
                effectiveStart !== null &&
                effectiveStart > now
              ) {
                excludedFutureTerms += 1;
                return false;
              }

              return true;
            }
          )
          .sort(
            (a, b) =>
              Number(b.id) -
              Number(a.id)
          );

      const validTermIds =
        new Set(
          state.terms.map(
            term =>
              String(term.id)
          )
        );

      state.selectedTermIds =
        new Set(
          [...state.selectedTermIds]
            .filter(
              termId =>
                validTermIds.has(
                  termId
                )
            )
        );

      renderTerms();

      setActivity(
        'Enrollment terms loaded',
        `Loaded ${state.terms.length} ` +
        `active or past terms. ` +
        `Excluded ${excludedFutureTerms} ` +
        `future term${
          excludedFutureTerms === 1
            ? ''
            : 's'
        }.`
      );

      addLog(
        `Terms loaded from root account ` +
        `${ROOT_ACCOUNT_ID}. ` +
        'Terms are sorted by Term ID descending.'
      );

      addLog(
        'Course-search accounts: ' +
        state.scopeAccounts
          .map(
            account =>
              `${account.name} (${account.id})`
          )
          .join(', ') +
        '.'
      );
    } catch (error) {
      if (isAbortError(error)) {
        setActivity(
          'Term loading canceled',
          'No Canvas data was changed.'
        );
      } else {
        state.terms = [];

        state.scopeAccounts = [];
        state.scopeAccountLoadError =
          'Could not load course-search accounts.';

        renderScopeAccounts();
        renderTerms();

        setActivity(
          'Could not load terms',
          error.message
        );

        addLog(
          `Term error: ${error.message}`
        );

        elements.details.open =
          true;
      }
    } finally {
      state.controller = null;

      setRunning(false);
    }
  }

  function getSelectedScopeAccounts() {
    return state.scopeAccounts.filter(
      account =>
        state.selectedScopeAccountIds.has(
          String(account.id)
        )
    );
  }

  function getSelectedScopeAccountIds() {
    const accountIds =
      new Set();

    for (
      const account
      of getSelectedScopeAccounts()
    ) {
      for (
        const accountId
        of account.accountIds
      ) {
        accountIds.add(
          String(accountId)
        );
      }
    }

    return [
      ...accountIds
    ];
  }

  async function findCourses() {
    if (state.running) {
      return;
    }

    const searchValue =
      elements.courseCode.value
        .trim();

    if (
      searchValue.length < 3
    ) {
      setActivity(
        'More search text is needed',
        'Enter at least three characters in ' +
        'Course code contains.'
      );

      return;
    }

    const selectedTerms =
      state.terms.filter(
        term =>
          state.selectedTermIds.has(
            String(term.id)
          )
      );

    if (!selectedTerms.length) {
      setActivity(
        'No terms selected',
        'Select at least one enrollment term.'
      );

      return;
    }

    const selectedScopeAccounts =
      getSelectedScopeAccounts();

    if (!selectedScopeAccounts.length) {
      setActivity(
        'No subaccounts selected',
        'Select at least one course-search account.'
      );

      return;
    }

    state.lastSearch =
      searchValue;

    state.courses = [];

    state.selectedCourseIds
      .clear();

    state.rows = [];

    state.failedTerms = [];
    state.failedCourses = [];
    state.noSubmissionCourses = [];

    state.downloadedFilename = '';

    renderCourses();
    clearLog();

    state.controller =
      new AbortController();

    setRunning(true);

    startOperation(
      'Finding courses',
      selectedTerms.length
    );

    setActivity(
      'Searching selected terms',
      `Searching ${selectedTerms.length} ` +
      `term${
        selectedTerms.length === 1
          ? ''
          : 's'
      } for course codes containing ` +
      `“${searchValue}”.`
    );

    try {
      const accountIds =
        getSelectedScopeAccountIds();

      const selectedScopeNames =
        selectedScopeAccounts.map(
          account =>
            account.name ||
            `Account ${account.id}`
        );

      const selectedScopeDescription =
        selectedScopeNames.length === 1
          ? selectedScopeNames[0]
          : (
              selectedScopeNames
                .slice(0, -1)
                .join(', ') +
              ' and ' +
              selectedScopeNames.at(-1)
            );

      const foundCourses =
        new Map();

      const normalizedSearch =
        searchValue.toLowerCase();

      let completedTerms = 0;

      await runPool(
        selectedTerms,
        TERM_CONCURRENCY,
        async term => {
          try {
            setActivity(
              `Searching ${
                term.name ||
                term.id
              }`,
              `Canvas is checking ${selectedScopeDescription} ` +
              'and their descendant subaccounts.'
            );

            const url =
              createApiUrl(
                `/api/v1/accounts/${ROOT_ACCOUNT_ID}/courses`,
                params => {
                  params.set(
                    'enrollment_term_id',
                    term.id
                  );

                  params.set(
                    'search_term',
                    searchValue
                  );

                  params.set(
                    'search_by',
                    'course'
                  );

                  params.append(
                    'include[]',
                    'term'
                  );

                  params.append(
                    'include[]',
                    'account_name'
                  );

                  for (
                    const courseState
                    of [
                      'created',
                      'claimed',
                      'available',
                      'completed'
                    ]
                  ) {
                    params.append(
                      'state[]',
                      courseState
                    );
                  }

                  for (
                    const accountId
                    of accountIds
                  ) {
                    params.append(
                      'by_subaccounts[]',
                      accountId
                    );
                  }

                  params.set(
                    'per_page',
                    '100'
                  );
                }
              );

            const returnedCourses =
              await fetchAllPages(
                url,
                {
                  signal:
                    state.controller
                      .signal,

                  label:
                    `courses in ${
                      term.name ||
                      term.id
                    }`,

                  onPage:
                    ({
                      pageNumber,
                      totalCount
                    }) => {
                      setActivity(
                        `Searching ${
                          term.name ||
                          term.id
                        }`,
                        `Received page ${pageNumber}. ` +
                        `${totalCount} candidate courses ` +
                        `received so far.`
                      );
                    }
                }
              );

            let matchedCount = 0;

            for (
              const course
              of returnedCourses
            ) {
              const returnedTermId =
                course
                  ?.enrollment_term_id ??
                course?.term?.id;

              if (
                !course ||
                course.workflow_state ===
                  'deleted' ||
                String(returnedTermId) !==
                  String(term.id) ||
                !stringValue(
                  course.course_code
                )
                  .trim()
                  .toLowerCase()
                  .includes(
                    normalizedSearch
                  )
              ) {
                continue;
              }

              matchedCount += 1;

              if (
                !foundCourses.has(
                  String(course.id)
                )
              ) {
                foundCourses.set(
                  String(course.id),
                  {
                    ...course,

                    _termName:
                      term.name ||
                      '',

                    _termId:
                      Number(term.id)
                  }
                );
              }
            }

            addLog(
              `${term.name || term.id}: ` +
              `${matchedCount} matching course` +
              `${
                matchedCount === 1
                  ? ''
                  : 's'
              }.`
            );
          } catch (error) {
            if (
              isAbortError(error)
            ) {
              throw error;
            }

            state.failedTerms.push({
              name:
                term.name ||
                String(term.id),

              message:
                error.message
            });

            addLog(
              `${term.name || term.id} failed: ` +
              error.message
            );
          } finally {
            completedTerms += 1;

            setMeasuredProgress(
              completedTerms,
              selectedTerms.length
            );
          }
        },
        state.controller.signal
      );

      state.courses =
        [...foundCourses.values()]
          .sort(
            (a, b) =>
              b._termId -
                a._termId ||

              compareText(
                a.course_code,
                b.course_code
              ) ||

              compareText(
                a.name,
                b.name
              )
          );

      state.selectedCourseIds =
        new Set(
          state.courses.map(
            course =>
              String(course.id)
          )
        );

      renderCourses();

      setActivity(
        state.courses.length
          ? (
              `Found ${state.courses.length} ` +
              `matching course${
                state.courses.length === 1
                  ? ''
                  : 's'
              }`
            )
          : 'No matching courses found',

        state.courses.length
          ? (
              'All matched courses are selected. ' +
              'Review them, then click Generate Report.'
            )
          : (
              'Try different terms or a different ' +
              'course-code search value.'
            )
      );

      if (
        state.failedTerms.length
      ) {
        elements.details.open =
          true;
      }
    } catch (error) {
      if (isAbortError(error)) {
        setActivity(
          'Course search canceled',
          'Partial search results were not retained.'
        );
      } else {
        setActivity(
          'Course search failed',
          error.message
        );

        addLog(
          `Course search error: ${error.message}`
        );

        elements.details.open =
          true;
      }
    } finally {
      state.controller = null;

      setRunning(false);
    }
  }

  function updateGenerationStatus() {
    const activeCourseText =
      [...state.activeCourses.values()]
        .map(
          item =>
            `${item.label} ` +
            `(page ${item.page || 1})`
        )
        .join(', ');

    setActivity(
      `Generating report: ` +
      `${state.completedUnits}/` +
      `${state.totalUnits} courses complete`,

      activeCourseText
        ? (
            `Currently working on ` +
            `${activeCourseText}. ` +
            `${state.rows.length} rows collected.`
          )
        : (
            `${state.rows.length} rows collected.`
          )
    );

    setMeasuredProgress(
      state.completedUnits,
      state.totalUnits
    );

    updateCounts();
  }

  async function getCourseAssignments(
    course,
    signal
  ) {
    const courseLabel =
      course.course_code ||
      course.name ||
      String(course.id);

    const url =
      createApiUrl(
        `/api/v1/courses/${course.id}/assignments`,
        params => {
          params.set(
            'per_page',
            '100'
          );
        }
      );

    return fetchAllPages(
      url,
      {
        signal,

        label:
          `assignments for ${courseLabel}`,

        onPage:
          ({
            pageNumber,
            totalCount
          }) => {
            state.activeCourses.set(
              String(course.id),
              {
                label:
                  courseLabel,

                page:
                  pageNumber
              }
            );

            setActivity(
              `Reading assignments for ${courseLabel}`,
              `Received page ${pageNumber}. ` +
              `${totalCount} assignments received so far.`
            );

            updateGenerationStatus();
          }
      }
    );
  }

  async function getCourseSubmissions(
    course,
    signal
  ) {
    const courseLabel =
      course.course_code ||
      course.name ||
      String(course.id);

    const url =
      createApiUrl(
        `/api/v1/courses/${course.id}/students/submissions`,
        params => {
          params.append(
            'student_ids[]',
            'all'
          );

          params.append(
            'include[]',
            'user'
          );

          params.set(
            'per_page',
            '100'
          );
        }
      );

    return fetchAllPages(
      url,
      {
        signal,

        label:
          `submissions for ${courseLabel}`,

        onPage:
          ({
            pageNumber,
            totalCount
          }) => {
            state.activeCourses.set(
              String(course.id),
              {
                label:
                  courseLabel,

                page:
                  pageNumber
              }
            );

            setActivity(
              `Reading grades for ${courseLabel}`,
              `Received page ${pageNumber}. ` +
              `${totalCount} submission records ` +
              `received so far.`
            );

            updateGenerationStatus();
          }
      }
    );
  }

  function isExactTestStudent(user) {
    return [
      'name',
      'display_name',
      'short_name',
      'sortable_name'
    ].some(
      key =>
        stringValue(
          user?.[key]
        )
          .trim()
          .toLowerCase() ===
        'test student'
    );
  }

  function isCanvasTestStudent(
    submission
  ) {
    if (
      submission
        ?.is_test_student === true ||

      submission
        ?.test_student === true ||

      isExactTestStudent(
        submission?.user
      )
    ) {
      return true;
    }

    const enrollment =
      submission?.enrollment ||
      submission?.user_enrollment;

    const enrollmentType =
      stringValue(
        enrollment?.type ||
        enrollment?.role ||
        enrollment?.base_role_type
      )
        .toLowerCase()
        .replaceAll(
          '_',
          ''
        );

    return enrollmentType.includes(
      'studentview'
    );
  }

  function getSubmissionStatus(
    submission
  ) {
    if (
      submission?.excused === true
    ) {
      return 'excused';
    }

    if (
      submission?.missing === true
    ) {
      return 'missing';
    }

    if (
      submission?.late === true &&
      submission.workflow_state
    ) {
      return (
        `${submission.workflow_state} ` +
        '(late)'
      );
    }

    return (
      submission?.workflow_state ||
      ''
    );
  }

  function numberOrBlank(value) {
    if (
      value === null ||
      value === undefined ||
      value === ''
    ) {
      return '';
    }

    const number =
      Number(value);

    return Number.isFinite(number)
      ? number
      : '';
  }

  function calculatePercentage(
    score,
    pointsPossible
  ) {
    if (
      score === '' ||
      pointsPossible === '' ||
      !Number.isFinite(score) ||
      !Number.isFinite(pointsPossible) ||
      pointsPossible <= 0
    ) {
      return '';
    }

    return (
      Math.round(
        (
          score /
          pointsPossible
        ) *
        10000
      ) /
      100
    );
  }

  function deduplicateSubmissions(
    submissions,
    courseId
  ) {
    const submissionMap =
      new Map();

    let fallbackCounter = 0;

    for (
      const submission
      of submissions
    ) {
      if (
        !submission ||
        typeof submission !==
          'object'
      ) {
        continue;
      }

      const canvasUserId =
        submission.user_id ??
        submission.user?.id;

      const assignmentId =
        submission.assignment_id ??
        submission.assignment?.id;

      const key =
        canvasUserId == null ||
        assignmentId == null
          ? (
              `${courseId}|missing|` +
              `${++fallbackCounter}`
            )
          : (
              `${courseId}|` +
              `${canvasUserId}|` +
              `${assignmentId}`
            );

      const existing =
        submissionMap.get(key);

      const currentUpdated =
        Date.parse(
          submission.updated_at
        ) || 0;

      const existingUpdated =
        Date.parse(
          existing?.updated_at
        ) || 0;

      if (
        !existing ||
        currentUpdated >=
          existingUpdated
      ) {
        submissionMap.set(
          key,
          submission
        );
      }
    }

    return [
      ...submissionMap.values()
    ];
  }

  function createRowsForCourse(
    course,
    submissions
  ) {
    const rows = [];

    for (
      const submission
      of deduplicateSubmissions(
        submissions,
        course.id
      )
    ) {
      if (
        isCanvasTestStudent(
          submission
        )
      ) {
        continue;
      }

      const canvasUserId =
        submission.user_id ??
        submission.user?.id ??
        '';

      const studentName =
        submission.user?.name ??
        '';

      const assignment =
        submission.assignment;

      const assignmentTitle =
        assignment?.name ??
        '';

      if (
        canvasUserId === ''
      ) {
        state.warnings
          .missingCanvasUserId += 1;
      }

      if (!studentName) {
        state.warnings
          .missingUserName += 1;
      }

      if (!assignmentTitle) {
        state.warnings
          .missingAssignment += 1;
      }

      const score =
        numberOrBlank(
          submission.score
        );

      const pointsPossible =
        numberOrBlank(
          assignment?.points_possible
        );

      rows.push({
        StudentId:
          canvasUserId,

        StudentName:
          studentName,

        EnrollmentTerm:
          course._termName ||
          '',

        CourseCode:
          course.course_code ||
          '',

        CourseName:
          course.name ||
          '',

        AssignmentTitle:
          assignmentTitle,

        AssignmentScore:
          score,

        AssignmentPointsPossible:
          pointsPossible,

        AssignmentPercentageScore:
          calculatePercentage(
            score,
            pointsPossible
          ),

        SubmissionStatus:
          getSubmissionStatus(
            submission
          ),

        _termId:
          course._termId
      });
    }

    return rows;
  }

  async function processCourse(
    course,
    signal
  ) {
    const courseLabel =
      course.course_code ||
      course.name ||
      String(course.id);

    state.activeCourses.set(
      String(course.id),
      {
        label:
          courseLabel,

        page:
          1
      }
    );

    updateGenerationStatus();

    // Assignment objects can be relatively large. Asking Canvas to embed one
    // in every submission repeats the same data once per student. Fetch the
    // assignment list once, then join it to submissions locally instead.
    const assignments =
      await getCourseAssignments(
        course,
        signal
      );

    const assignmentById =
      new Map(
        assignments.map(
          assignment => [
            String(assignment.id),
            assignment
          ]
        )
      );

    const submissions =
      await getCourseSubmissions(
        course,
        signal
      );

    for (
      const submission
      of submissions
    ) {
      if (
        !submission.assignment &&
        submission.assignment_id != null
      ) {
        submission.assignment =
          assignmentById.get(
            String(
              submission.assignment_id
            )
          );
      }
    }

    if (!submissions.length) {
      state.noSubmissionCourses.push(
        courseLabel
      );

      addLog(
        `${courseLabel}: no submission records.`
      );

      return;
    }

    setActivity(
      `Preparing rows for ${courseLabel}`,
      `Processing ${submissions.length} ` +
      `submission records in the browser.`
    );

    const rows =
      createRowsForCourse(
        course,
        submissions
      );

    state.rows.push(...rows);

    markProgress();

    addLog(
      `${courseLabel}: ` +
      `${rows.length} report row` +
      `${
        rows.length === 1
          ? ''
          : 's'
      } from ` +
      `${submissions.length} returned ` +
      `submission record` +
      `${
        submissions.length === 1
          ? ''
          : 's'
      }.`
    );
  }

  function protectCsvText(value) {
    const raw =
      stringValue(value);

    return /^[=+\-@]/.test(raw)
      ? `'${raw}`
      : raw;
  }

  function quoteCsv(value) {
    return (
      `"${stringValue(value)
        .replace(
          /"/g,
          '""'
        )}"`
    );
  }

  function sortRows(rows) {
    return [...rows].sort(
      (a, b) =>
        b._termId -
          a._termId ||

        compareText(
          a.CourseCode,
          b.CourseCode
        ) ||

        compareText(
          a.CourseName,
          b.CourseName
        ) ||

        compareText(
          a.StudentName,
          b.StudentName
        ) ||

        compareText(
          a.AssignmentTitle,
          b.AssignmentTitle
        )
    );
  }

  function buildCsv(rows) {
    const lines = [
      CSV_HEADERS.join(',')
    ];

    for (
      const row
      of sortRows(rows)
    ) {
      const values =
        CSV_HEADERS.map(
          header => {
            const value =
              row[header];

            if (
              TEXT_FIELDS.has(header)
            ) {
              return quoteCsv(
                protectCsvText(
                  value
                )
              );
            }

            return (
              value === '' ||
              value == null
                ? ''
                : String(value)
            );
          }
        );

      lines.push(
        values.join(',')
      );
    }

    return (
      '\uFEFF' +
      lines.join('\r\n')
    );
  }

  function getFilename() {
    const date =
      new Date();

    const dateStamp =
      `${date.getFullYear()}-` +
      `${String(
        date.getMonth() + 1
      ).padStart(2, '0')}-` +
      `${String(
        date.getDate()
      ).padStart(2, '0')}`;

    const safeSearch =
      stringValue(
        state.lastSearch ||
        elements.courseCode.value
      )
        .trim()
        .replace(
          /[<>:"/\\|?*\x00-\x1F]/g,
          '_'
        )
        .replace(
          /[. ]+$/g,
          ''
        )
        .replace(
          /_+/g,
          '_'
        )
        .slice(
          0,
          100
        );

    const prefix =
      /[A-Za-z0-9]/.test(
        safeSearch
      )
        ? safeSearch
        : 'Canvas';

    return (
      `${prefix}_Assignment_Grades_` +
      `${dateStamp}.csv`
    );
  }

  function downloadCsv(
    rows,
    filename
  ) {
    setActivity(
      'Building CSV file',
      `Formatting ${rows.length} report row` +
      `${
        rows.length === 1
          ? ''
          : 's'
      } for download.`
    );

    const blob =
      new Blob(
        [
          buildCsv(rows)
        ],
        {
          type:
            'text/csv;charset=utf-8'
        }
      );

    const objectUrl =
      URL.createObjectURL(blob);

    state.objectUrls.add(
      objectUrl
    );

    const anchor =
      document.createElement(
        'a'
      );

    anchor.href =
      objectUrl;

    anchor.download =
      filename;

    anchor.style.display =
      'none';

    shadow.appendChild(
      anchor
    );

    anchor.click();
    anchor.remove();

    setTimeout(
      () => {
        URL.revokeObjectURL(
          objectUrl
        );

        state.objectUrls.delete(
          objectUrl
        );
      },
      1500
    );

    state.downloadedFilename =
      filename;
  }

  function renderFinalSummary(
    selectedCourseCount
  ) {
    const lines = [
      `Search subaccounts: ${
        getSelectedScopeAccounts()
          .map(
            account =>
              `${account.name} (${account.id})`
          )
          .join(', ') ||
        'None'
      }`,
      `Selected terms: ${state.selectedTermIds.size}`,
      `Matched courses: ${state.courses.length}`,
      `Selected courses: ${selectedCourseCount}`,
      `Courses completed: ${state.completedUnits}`,
      `Courses with no submissions: ${state.noSubmissionCourses.length}`,
      `Failed courses: ${state.failedCourses.length}`,
      `CSV rows created: ${state.rows.length}`,
      `API requests made: ${state.requestCount}`,
      `Downloaded filename: ${state.downloadedFilename || 'Not downloaded'}`,
      `Missing Canvas user IDs: ${state.warnings.missingCanvasUserId}`,
      `Missing student names: ${state.warnings.missingUserName}`,
      `Missing assignments: ${state.warnings.missingAssignment}`
    ];

    if (
      state.failedCourses.length
    ) {
      lines.push(
        '',
        'Failed courses:'
      );

      for (
        const failure
        of state.failedCourses
      ) {
        lines.push(
          `  ${failure.label}: ` +
          failure.message
        );
      }
    }

    if (
      state.failedTerms.length
    ) {
      lines.push(
        '',
        'Failed terms:'
      );

      for (
        const failure
        of state.failedTerms
      ) {
        lines.push(
          `  ${failure.name}: ` +
          failure.message
        );
      }
    }

    elements.summary.textContent =
      lines.join('\n');

    if (
      state.failedCourses.length ||
      state.failedTerms.length ||
      Object.values(
        state.warnings
      ).some(Boolean)
    ) {
      elements.details.open =
        true;
    }
  }

  async function generateReport() {
    if (state.running) {
      return;
    }

    const selectedCourses =
      state.courses.filter(
        course =>
          state.selectedCourseIds.has(
            String(course.id)
          )
      );

    if (!selectedCourses.length) {
      setActivity(
        'No courses selected',
        'Select at least one matched course.'
      );

      return;
    }

    state.rows = [];

    state.failedCourses = [];
    state.noSubmissionCourses = [];

    state.downloadedFilename =
      '';

    state.activeCourses.clear();

    state.warnings = {
      missingCanvasUserId: 0,
      missingUserName: 0,
      missingAssignment: 0
    };

    elements.downloadPartial.hidden =
      true;

    elements.downloadEmpty.hidden =
      true;

    clearLog();

    state.controller =
      new AbortController();

    setRunning(true);

    startOperation(
      'Generating report',
      selectedCourses.length
    );

    setActivity(
      'Starting grade collection',
      `Processing up to ${COURSE_CONCURRENCY} ` +
      `courses at a time. The animated bar and ` +
      `timer will continue while Canvas is working.`
    );

    try {
      await runPool(
        selectedCourses,
        COURSE_CONCURRENCY,
        async course => {
          const courseLabel =
            course.course_code ||
            course.name ||
            String(course.id);

          try {
            await processCourse(
              course,
              state.controller.signal
            );
          } catch (error) {
            if (
              isAbortError(error)
            ) {
              throw error;
            }

            state.failedCourses.push({
              label:
                courseLabel,

              message:
                error.message
            });

            addLog(
              `${courseLabel} failed: ` +
              error.message
            );
          } finally {
            state.activeCourses.delete(
              String(course.id)
            );

            state.completedUnits += 1;

            updateGenerationStatus();
          }
        },
        state.controller.signal
      );

      if (
        state.failedCourses.length ===
        selectedCourses.length
      ) {
        setActivity(
          'Every selected course failed',
          'No CSV file was downloaded.'
        );
      } else if (
        state.rows.length
      ) {
        const filename =
          getFilename();

        downloadCsv(
          state.rows,
          filename
        );

        setActivity(
          'Report complete',
          `Downloaded ${filename}. ` +
          `${state.rows.length} rows were created.`
        );
      } else {
        setActivity(
          'No report rows were returned',
          'Use Download Empty CSV to create ' +
          'a header-only file.'
        );

        elements.downloadEmpty.hidden =
          false;
      }
    } catch (error) {
      if (
        isAbortError(error)
      ) {
        setActivity(
          'Report generation canceled',
          'Partial results were not downloaded automatically.'
        );

        if (
          state.rows.length
        ) {
          elements.downloadPartial.hidden =
            false;
        }
      } else {
        setActivity(
          'Report generation failed',
          error.message
        );

        addLog(
          `Unexpected report error: ${error.message}`
        );

        if (
          state.rows.length
        ) {
          elements.downloadPartial.hidden =
            false;
        }
      }

      elements.details.open =
        true;
    } finally {
      state.controller = null;

      state.activeCourses.clear();

      setRunning(false);

      renderFinalSummary(
        selectedCourses.length
      );
    }
  }

  function cancelOperation() {
    if (
      state.running &&
      state.controller
    ) {
      setActivity(
        'Canceling operation',
        'Waiting for active Canvas requests to stop.'
      );

      state.controller.abort();
    }
  }

  function closeApplication(
    force = false
  ) {
    if (
      state.destroyed
    ) {
      return;
    }

    if (
      state.running &&
      !force &&
      !window.confirm(
        'An operation is running. Close and cancel it?'
      )
    ) {
      return;
    }

    state.destroyed =
      true;

    state.controller?.abort();

    clearInterval(
      state.operationTimer
    );

    for (
      const objectUrl
      of state.objectUrls
    ) {
      try {
        URL.revokeObjectURL(
          objectUrl
        );
      } catch (_) {
        // Ignore cleanup failures.
      }
    }

    state.objectUrls.clear();

    document.body.style.overflow =
      previousBodyOverflow;

    host.remove();

    if (
      window[GLOBAL_KEY]?.close ===
      closeApplication
    ) {
      delete window[GLOBAL_KEY];
    }

    try {
      if (
        previousFocus?.isConnected
      ) {
        previousFocus.focus();
      }
    } catch (_) {
      // Ignore focus restoration failures.
    }
  }

  elements.courseCode.addEventListener(
    'input',
    () => {
      clearCourseResults(
        'Course-code search changed. ' +
        'Click Find Courses again.'
      );

      updateCounts();
    }
  );

  elements.courseCode.addEventListener(
    'keydown',
    event => {
      if (
        event.key === 'Enter' &&
        !elements.findCourses.disabled
      ) {
        event.preventDefault();
        findCourses();
      }
    }
  );

  elements.termSearch.addEventListener(
    'input',
    renderTerms
  );

  elements.selectVisibleTerms.addEventListener(
    'click',
    () => {
      const filter =
        elements.termSearch.value
          .trim()
          .toLowerCase();

      for (
        const term
        of state.terms
      ) {
        if (
          !filter ||
          stringValue(term.name)
            .toLowerCase()
            .includes(filter)
        ) {
          state.selectedTermIds.add(
            String(term.id)
          );
        }
      }

      clearCourseResults(
        'Term selection changed. ' +
        'Click Find Courses again.'
      );

      renderTerms();
    }
  );

  elements.clearTerms.addEventListener(
    'click',
    () => {
      state.selectedTermIds.clear();

      clearCourseResults(
        'Term selection cleared.'
      );

      renderTerms();
    }
  );

  elements.reloadTerms.addEventListener(
    'click',
    loadTerms
  );

  elements.selectAllCourses.addEventListener(
    'click',
    () => {
      state.selectedCourseIds =
        new Set(
          state.courses.map(
            course =>
              String(course.id)
          )
        );

      renderCourses();
    }
  );

  elements.clearCourses.addEventListener(
    'click',
    () => {
      state.selectedCourseIds.clear();

      renderCourses();
    }
  );

  elements.findCourses.addEventListener(
    'click',
    findCourses
  );

  elements.generateReport.addEventListener(
    'click',
    generateReport
  );

  elements.cancel.addEventListener(
    'click',
    cancelOperation
  );

  elements.downloadPartial.addEventListener(
    'click',
    () => {
      if (
        !state.rows.length
      ) {
        return;
      }

      const filename =
        getFilename().replace(
          /\.csv$/i,
          '_PARTIAL.csv'
        );

      downloadCsv(
        state.rows,
        filename
      );

      setActivity(
        'Partial results downloaded',
        filename
      );

      renderFinalSummary(
        state.selectedCourseIds.size
      );
    }
  );

  elements.downloadEmpty.addEventListener(
    'click',
    () => {
      const filename =
        getFilename();

      downloadCsv(
        [],
        filename
      );

      setActivity(
        'Empty CSV downloaded',
        filename
      );

      renderFinalSummary(
        state.selectedCourseIds.size
      );
    }
  );

  elements.close.addEventListener(
    'click',
    () =>
      closeApplication(false)
  );

  elements.topClose.addEventListener(
    'click',
    () =>
      closeApplication(false)
  );

  shadow.addEventListener(
    'keydown',
    event => {
      if (
        event.key === 'Escape'
      ) {
        event.preventDefault();

        closeApplication(false);
      }
    }
  );

  window[GLOBAL_KEY] = {
    close:
      closeApplication,

    getSubmissionStatus,

    calculatePercentage,

    description:
      'StudentId contains the Canvas User ID. ' +
      'The progress area remains animated while ' +
      'Canvas requests are pending.'
  };

  updateCounts();
  updateProgressDisplay();

  renderScopeAccounts();

  elements.courseCode.focus();

  await loadTerms();
})();
