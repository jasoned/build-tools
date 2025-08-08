<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>DOCX to HTML Converter</title>
  <!-- Mammoth.js library -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
  <style>
    /* --- Basic styling for the body and overall layout --- */
    body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif, "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol";
        padding: 20px;
        background: #f4f7f6;
        display: flex;
        flex-direction: column;
        align-items: center;
        min-height: 100vh;
        margin: 0;
        box-sizing: border-box;
    }
    h1 { color: #333; margin-bottom: 20px; text-align: center; }
    #uploadArea { width: 90%; max-width: 800px; text-align: center; padding: 20px; border: 2px dashed #ccc; border-radius: 8px; background-color: #fafafa; transition: background-color 0.2s ease, border-color 0.2s ease; margin-bottom: 20px; }
    #uploadArea.drag-over { background-color: #e6f2ff; border-color: #007bff; }
    #upload { margin-top: 10px; }
    #output { background: white; padding: 20px; margin-top: 20px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); white-space: pre-wrap; width: 90%; max-width: 800px; overflow-x: auto; line-height: 1.6; color: #333; min-height: 150px; border: 1px solid #ddd; text-align: left; }

    /* --- Specific element styling --- */
    h1.wk-topic { color: #0056b3; border-bottom: 2px solid #0056b3; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; }
    div.LO { background-color: #e6f2ff; border-left: 5px solid #007bff; padding: 15px 20px; margin: 20px 0; border-radius: 8px; box-shadow: 0 2px 5px rgba(0, 123, 255, 0.2); }
    div.LO p { margin-bottom: 8px; line-height: 1.5; font-size: 1.05em; color: #2c3e50; }
    div.activities { background-color: #f7f9fc; border-left: 5px solid #28a745; padding: 15px 20px; margin: 20px 0; border-radius: 8px; box-shadow: 0 2px 5px rgba(40, 167, 69, 0.1); }
    div.assignments { background-color: #fff9e6; border-left: 5px solid #ffc107; padding: 15px 20px; margin: 20px 0; border-radius: 8px; box-shadow: 0 2px 5px rgba(255, 193, 7, 0.2); }
    .activities hr, .assignments hr { border: 0; height: 1px; background-image: linear-gradient(to right, rgba(0, 0, 0, 0), rgba(0, 0, 0, 0.2), rgba(0, 0, 0, 0)); margin: 20px 0; }

    h1.cdes { color: #6a0dad; border-bottom: 1px dashed #6a0dad; padding-bottom: 3px; font-size: 1.6em; }
    div.cdes { background-color: #f3e9fa; border-left: 4px solid #6a0dad; padding: 15px; margin: 10px 0 20px; border-radius: 5px; }
    h1.plo { color: #d9534f; border-bottom: 1px dashed #d9534f; padding-bottom: 3px; font-size: 1.6em; }
    div.plo { background-color: #fdf2f2; border-left: 4px solid #d9534f; padding: 15px; margin: 10px 0 20px; border-radius: 5px; }
    div.plo ul { list-style: disc inside; padding-left: 10px; }
    div.plo li { margin-bottom: 5px; }
    h1.clo { color: #5cb85c; border-bottom: 1px dashed #5cb85c; padding-bottom: 3px; font-size: 1.6em; }
    div.clo { background-color: #e9fae9; border-left: 4px solid #5cb85c; padding: 15px; margin: 10px 0 20px; border-radius: 5px; }
    div.clo ul { list-style: disc inside; padding-left: 10px; }
    div.clo li { margin-bottom: 5px; }
    h1.material { color: #f0ad4e; border-bottom: 1px dashed #f0ad4e; padding-bottom: 3px; font-size: 1.6em; }
    div.material { background-color: #fffbe6; border-left: 4px solid #f0ad4e; padding: 15px; margin: 10px 0 20px; border-radius: 5px; }

    /* Table look (you asked to keep borders) */
    table { border-collapse: collapse; width: 100%; }
    td, th { border: 1px solid #cfd4da; padding: 8px; vertical-align: top; }

    .action-buttons { display: flex; gap: 15px; margin-top: 20px; margin-bottom: 20px; flex-wrap: wrap; justify-content: center; }
    .action-buttons button { background-color: #007bff; color: white; padding: 10px 20px; border: none; border-radius: 8px; cursor: pointer; font-size: 1em; box-shadow: 0 4px 8px rgba(0,0,0,0.1); transition: background-color 0.3s ease, transform 0.2s ease; }
    .action-buttons button:hover { background-color: #0056b3; transform: translateY(-2px); }
    .action-buttons button:active { transform: translateY(0); }
    #copyHtmlButton { background-color: #6c757d; }
    #copyHtmlButton:hover { background-color: #5a6268; }
    #downloadHtmlButton { background-color: #28a745; }
    #downloadHtmlButton:hover { background-color: #218838; }
    #outputReportButton { background-color: #17a2b8; }
    #outputReportButton:hover { background-color: #138496; }
    #cleanupButton { background-color: #dc3545; }
    #cleanupButton:hover { background-color: #c82333; }
    #undoButton { background-color: #ffc107; color: #212529;}
    #undoButton:hover { background-color: #e0a800; }

    #reportOutput { background: #f8f9fa; padding: 20px; margin-top: 20px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); width: 90%; max-width: 800px; line-height: 1.6; color: #333; min-height: 50px; border: 1px solid #e2e6ea; text-align: left; position: relative; }
    #reportOutput .close-button { position: absolute; top: 10px; right: 15px; font-size: 1.5em; font-weight: bold; color: #888; cursor: pointer; background: none; border: none; padding: 0; line-height: 1; transition: color 0.2s ease; }
    #reportOutput .close-button:hover { color: #555; }
    #reportOutput table { width: 100%; border-collapse: collapse; margin-top: 15px; }
    #reportOutput th, #reportOutput td { border: 1px solid #dee2e6; padding: 8px; text-align: left; }
    #reportOutput th { background-color: #e9ecef; font-weight: bold; }
    #reportOutput ul { list-style: disc inside; padding-left: 20px; margin-top: 10px;}
    #reportOutput .error-message, .error-message { color: #dc3545; font-weight: bold; }
    #reportOutput .success-message { color: #28a745; font-weight: bold; }

    #wk-topics-list { background-color: #f0f8ff; border: 1px solid #b0e0e6; padding: 15px 20px; margin-bottom: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); width: calc(100% - 40px); max-width: 800px; margin-left: auto; margin-right: auto; }
    #wk-topics-list h2 { color: #2c3e50; margin-top: 0; margin-bottom: 10px; border-bottom: 1px dotted #ccc; padding-bottom: 5px; }
    #wk-topics-list ul { list-style: none; padding: 0; margin: 0; }
    #wk-topics-list li { margin-bottom: 5px; padding-left: 15px; position: relative; }
    #wk-topics-list li::before { content: '•'; color: #007bff; position: absolute; left: 0; top: 0; }

    .loading-message { font-style: italic; color: #555; margin-top: 10px; text-align: center; }
    #copyFeedback { margin-top: 10px; font-size: 0.9em; color: #555; opacity: 0; transition: opacity 0.3s ease-in-out; text-align: center; }
    #copyFeedback.show { opacity: 1; }
    #scrollToTopButton { position: fixed; bottom: 20px; right: 20px; width: 50px; height: 50px; background-color: #0056b3; color: white; border: none; border-radius: 50%; font-size: 24px; cursor: pointer; display: none; align-items: center; justify-content: center; box-shadow: 0 4px 8px rgba(0,0,0,0.2); z-index: 1000; transition: opacity 0.3s ease, transform 0.3s ease; opacity: 0; transform: translateY(10px); }
    #scrollToTopButton.show { display: flex; opacity: 1; transform: translateY(0); }
    #scrollToTopButton:hover { background-color: #007bff; }

    .modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 2000; display: none; align-items: center; justify-content: center; }
    .modal-content { background: white; padding: 30px; border-radius: 8px; text-align: center; max-width: 400px; box-shadow: 0 5px 15px rgba(0,0,0,0.3); }
    .modal-content p { margin-bottom: 20px; }
    .modal-buttons button { padding: 10px 20px; border-radius: 5px; border: none; cursor: pointer; margin: 0 10px; }
    #confirmCleanupButton { background-color: #dc3545; color: white; }
    #cancelCleanupButton { background-color: #6c757d; color: white; }

    #sessionRestoreBanner { display: none; position: fixed; top: 0; left: 0; width: 100%; background-color: #fffbe6; border-bottom: 1px solid #ffc107; padding: 15px; text-align: center; z-index: 3000; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    #sessionRestoreBanner button { padding: 8px 15px; border-radius: 5px; border: 1px solid #ccc; cursor: pointer; margin: 0 10px; }

    @media (max-width: 768px) {
        body { padding: 15px; }
        h1 { font-size: 1.8em; }
        #output, #reportOutput, #wk-topics-list, #uploadArea { padding: 15px; width: 95%; }
        .action-buttons { flex-direction: column; gap: 10px; }
    }
  </style>
</head>
<body>
  <div id="sessionRestoreBanner">
      You have a saved session.
      <button id="restoreSessionButton">Restore</button>
      <button id="dismissRestoreButton">Dismiss</button>
  </div>

  <h1>DOCX to HTML Converter</h1>

  <div class="action-buttons">
    <button id="copyHtmlButton">📋 Copy HTML</button>
    <button id="downloadHtmlButton">📥 Download HTML</button>
    <button id="outputReportButton">📊 Section Report</button>
    <button id="cleanupButton">🧹 Clean Up Formatting</button>
    <button id="undoButton" style="display: none;">↩️ Undo Cleanup</button>
  </div>
  
  <div id="copyFeedback"></div>
  <div id="reportOutput"></div>

  <div id="uploadArea">
      <p>Drag and drop a .docx file here, or click to select a file.</p>
      <input type="file" id="upload" accept=".docx" />
  </div>
  <div id="output"></div>

  <button id="scrollToTopButton">↑</button>

  <!-- Confirmation Modal HTML -->
  <div class="modal-overlay" id="cleanupModal">
      <div class="modal-content">
          <p>This will remove any content (and its heading) that is not in a standard section like "Learning Objectives" or "Activities."<br><br>You can undo this change immediately after. Are you sure?</p>
          <div class="modal-buttons">
              <button id="cancelCleanupButton">Cancel</button>
              <button id="confirmCleanupButton">Yes, Clean Up</button>
          </div>
      </div>
  </div>

  <script>
    // --- APP CONFIGURATION ---
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

    // --- HELPERS ---
    const textOnly = (el) => (el.textContent || '').trim().replace(/\s+/g, ' ');

    // Stash NESTED tables and replace them with placeholders.
    function stashNestedTables(doc) {
      const store = new Map();
      let counter = 0;
      // Any table with a table ancestor is considered nested and should be preserved.
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

    // Move children out of a cell into destination, respecting placeholders and lists
    function moveChildrenPreservingBlocks(src, dst) {
      let currentP = null;
      const flushP = () => { if (currentP && currentP.childNodes.length) dst.appendChild(currentP); currentP = null; };
      const ensureP = () => (currentP ||= document.createElement('p'));

      function walk(node) {
        while (node.firstChild) {
          const child = node.firstChild;
          node.removeChild(child);

          if (child.nodeType === Node.TEXT_NODE) {
            if (child.textContent.trim()) ensureP().appendChild(document.createTextNode(child.textContent));
            continue;
          }
          if (child.nodeType !== Node.ELEMENT_NODE) continue;

          // Our placeholders act like block elements
          if (child.matches('span[data-table-ph]')) { flushP(); dst.appendChild(child); continue; }

          // Lists are blocks
          if (child.tagName === 'UL' || child.tagName === 'OL') { flushP(); dst.appendChild(child); continue; }

          // If we encounter actual tables (rare here because we stashed nested), keep as block
          if (child.tagName === 'TABLE') { flushP(); dst.appendChild(child); continue; }

          // Paragraph-like containers: walk through to collect inline pieces and reveal placeholders
          if (child.tagName === 'P' || child.tagName === 'DIV' || child.tagName === 'SECTION' || child.tagName === 'ARTICLE') {
            walk(child);
            flushP();
            continue;
          }

          // Default: treat as inline to avoid invalid <p> nesting
          ensureP().appendChild(child);
        }
      }

      walk(src);
      flushP();
    }

    // --- SECTION/TABLE PROCESSING ---
    function processContentTables(doc, sectionHeadingText, wrapperClass, endWithHr = false) {
        const sectionHeadings = Array.from(doc.querySelectorAll('h1')).filter(h =>
            h.textContent.trim().toLowerCase().includes(sectionHeadingText.toLowerCase())
        );

        sectionHeadings.forEach(heading => {
            const mainWrapper = document.createElement('div');
            mainWrapper.className = wrapperClass;

            const tablesToProcess = [];
            let currentNode = heading.nextElementSibling;
            while (currentNode && currentNode.tagName !== 'H1') {
                if (currentNode.tagName === 'TABLE') tablesToProcess.push(currentNode);
                currentNode = currentNode.nextElementSibling;
            }
            tablesToProcess.forEach((table, index) => {
                const itemDiv = document.createElement('div');
                const rows = Array.from(table.querySelectorAll('tr'));
                if (rows.length > 0) {
                    // H2 title from FIRST cell text (strip wrappers)
                    const titleCell = rows[0].querySelector('td, th');
                    if (titleCell && titleCell.textContent.trim()) {
                        const h2 = document.createElement('h2');
                        h2.textContent = textOnly(titleCell);
                        itemDiv.appendChild(h2);
                    }
                    // Content from subsequent rows (preserve lists/placeholders)
                    for (let i = 1; i < rows.length; i++) {
                        const contentCell = rows[i].querySelector('td, th');
                        if (contentCell) moveChildrenPreservingBlocks(contentCell, itemDiv);
                    }
                }
                if (itemDiv.hasChildNodes()) mainWrapper.appendChild(itemDiv);
                const isLastItem = index === tablesToProcess.length - 1;
                if (endWithHr || !isLastItem) mainWrapper.appendChild(document.createElement('hr'));
            });
            if (mainWrapper.hasChildNodes()) {
                heading.insertAdjacentElement('afterend', mainWrapper);
                tablesToProcess.forEach(t => t.remove());
            }
        });
    }

    function wrapSectionContent(doc, headingText, wrapperClass) {
        const heading = Array.from(doc.querySelectorAll('h1')).find(h => h.textContent.trim().toLowerCase().includes(headingText.toLowerCase()));
        if (!heading) return;
        const wrapper = document.createElement('div');
        wrapper.className = wrapperClass;
        let currentNode = heading.nextElementSibling;
        const elementsToMove = [];
        while (currentNode && currentNode.tagName !== 'H1') { elementsToMove.push(currentNode); currentNode = currentNode.nextElementSibling; }
        elementsToMove.forEach(el => wrapper.appendChild(el));
        if (wrapper.hasChildNodes()) heading.insertAdjacentElement('afterend', wrapper);
    }

    function cleanupHtml(doc) {
        doc.querySelectorAll('p, a').forEach(el => { if (!el.textContent.trim() && !el.children.length) el.remove(); });
        return doc;
    }

    function createWeeklyTopicsList(doc) {
        const topicHeaders = doc.querySelectorAll('h1.wk-topic');
        if (topicHeaders.length === 0) return null;
        const listContainer = document.createElement('div');
        listContainer.id = 'wk-topics-list';
        const title = document.createElement('h2');
        title.textContent = 'Weekly Topics';
        listContainer.appendChild(title);
        const ul = document.createElement('ul');
        topicHeaders.forEach(h1 => { const li = document.createElement('li'); li.textContent = h1.textContent.trim().replace(/^Week\s+\d+:\s*/i, ''); ul.appendChild(li); });
        listContainer.appendChild(ul);
        return listContainer;
    }

    function processHtml(rawHtml) {
        const parser = new DOMParser();
        let doc = parser.parseFromString(rawHtml, 'text/html');

        // 1) Stash nested tables as placeholders so nothing touches them during transforms
        const tableStore = stashNestedTables(doc);

        // 2) Wrap high-level sections
        APP_CONFIG.topLevelSections.forEach(section => { wrapSectionContent(doc, section.name, section.className); });

        // 3) Convert LO table under each Week header to bullet paragraphs
        doc.querySelectorAll('h1.wk-topic').forEach(topic => {
            const nextH1 = topic.nextElementSibling;
            if (nextH1 && nextH1.tagName === 'H1' && nextH1.textContent.trim().toLowerCase().includes('learning objectives')) {
                const table = nextH1.nextElementSibling;
                if (table && table.tagName === 'TABLE') {
                    const wrapper = document.createElement('div');
                    wrapper.className = 'LO';
                    const seen = new Set();
                    Array.from(table.querySelectorAll('li')).forEach(li => {
                        const itemText = li.textContent.trim();
                        if (itemText && !seen.has(itemText)) { const p = document.createElement('p'); p.textContent = itemText; wrapper.appendChild(p); seen.add(itemText); }
                    });
                    if (wrapper.hasChildNodes()) { nextH1.insertAdjacentElement('afterend', wrapper); table.remove(); }
                }
            }
        });

        // 4) Build Activities/Assignments sections from their top-level tables
        processContentTables(doc, 'Activities and Resources', 'activities', false);
        processContentTables(doc, 'Assignments', 'assignments', true);

        // 5) Restore the stashed nested tables back in place
        restorePlaceholders(doc, tableStore);

        // 6) Final tidy + weekly list
        doc = cleanupHtml(doc);
        const weeklyTopicsList = createWeeklyTopicsList(doc);
        if (weeklyTopicsList) doc.body.prepend(weeklyTopicsList);
        return doc.body.innerHTML;
    }

    // --- REPORTING AND CLEANUP ACTIONS ---
    function generateSectionPresenceReport(htmlContent, reportOutputDiv) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlContent, 'text/html');
        reportOutputDiv.style.display = 'block';
        let issuesFound = false;
        const recommendations = [];
        const topLevelErrors = [];
        APP_CONFIG.topLevelSections.forEach(section => {
            if (doc.querySelectorAll(`.${section.className}`).length === 0) {
                issuesFound = true; const msg = `Top-Level section "${section.name}" is missing.`; topLevelErrors.push(msg); recommendations.push(msg);
            }
        });
        const wkTopicHeaders = doc.querySelectorAll('h1.wk-topic');
        const weeklyReportData = [];
        let anyUnwrappedContentFound = false;
        if (wkTopicHeaders.length === 0) { issuesFound = true; recommendations.push("No 'Weekly Topic' sections (h1.wk-topic) were found."); }
        wkTopicHeaders.forEach((wkTopic, index) => {
            const weekNumber = index + 1;
            const cleanWkTopicText = wkTopic.textContent.trim().replace(/^Week\s+\d+:\s*/i, '');
            let weekContentElements = [];
            let currentElement = wkTopic.nextElementSibling;
            while(currentElement && !currentElement.matches('h1.wk-topic')) { weekContentElements.push(currentElement); currentElement = currentElement.nextElementSibling; }
            const tempFragment = document.createDocumentFragment();
            weekContentElements.forEach(el => tempFragment.appendChild(el.cloneNode(true)));
            const weeklyCounts = {};
            APP_CONFIG.weeklySections.forEach(section => { weeklyCounts[section.className] = tempFragment.querySelectorAll(`.${section.className}`).length; });
            let unwrappedFound = false;
            Array.from(tempFragment.children).forEach(child => {
                if (child.nodeType === Node.ELEMENT_NODE && !child.matches(APP_CONFIG.allowedWrapperSelectors) && !child.matches('h1, hr')) {
                    if (child.textContent.trim() !== '') unwrappedFound = true;
                }
            });
            if (unwrappedFound) anyUnwrappedContentFound = true;
            weeklyReportData.push({weekNumber, cleanWkTopicText, counts: weeklyCounts, unwrapped: unwrappedFound});
            APP_CONFIG.weeklySections.forEach(section => {
                const count = weeklyCounts[section.className];
                const sectionName = section.name;
                if (count === 0) { issuesFound = true; recommendations.push(`Week ${weekNumber} is missing a ${sectionName} section.`); }
                if (count > 1) { issuesFound = true; recommendations.push(`Week ${weekNumber} has too many ${sectionName} sections (found ${count}, expected 1).`); }
            });
            if (unwrappedFound) { issuesFound = true; recommendations.push(`Week ${weekNumber} has unwrapped content.`); }
        });
        let reportHtml = '<h3>Document Structure Report<button class="close-button">&times;</button></h3>';
        if (issuesFound) {
            reportHtml += `<p class="error-message">Issues detected. Please review recommendations.</p>`;
            topLevelErrors.forEach(error => { reportHtml += `<p class="error-message" style="padding-left: 20px;">${error}</p>`; });
        }
        reportHtml += '<h4>Top-Level Sections</h4><table><thead><tr><th>Section</th><th>Present?</th><th>Item Count</th></tr></thead><tbody>';
        APP_CONFIG.topLevelSections.forEach(section => {
            const count = doc.querySelectorAll(`.${section.className}`).length;
            reportHtml += `<tr><td>${section.name}</td><td>${count > 0 ? 'Yes' : 'No'}</td><td>${count}</td></tr>`;
        });
        reportHtml += '</tbody></table>';
        reportHtml += '<h4>Weekly Sections</h4>';
        if (anyUnwrappedContentFound) { reportHtml += `<p class="error-message">Unwrapped content was found. This will be deleted if you use the "Clean Up" tool.</p>`; }
        reportHtml += '<table><thead><tr><th>Week</th><th>LO</th><th>A&R</th><th>Assignments</th><th>Unwrapped Content</th></tr></thead><tbody>';
        weeklyReportData.forEach(data => { reportHtml += `<tr><td>Week ${data.weekNumber}: ${data.cleanWkTopicText}</td><td>${data.counts.LO}</td><td>${data.counts.activities}</td><td>${data.counts.assignments}</td><td>${data.unwrapped ? '1' : '0'}</td></tr>`; });
        reportHtml += `</tbody></table><p><strong>${wkTopicHeaders.length} weeks found.</strong></p>`;
        reportHtml += '<h4>Summary & Recommendations</h4>';
        if (!issuesFound) { reportHtml += `<p class="success-message">Document structure looks good!</p>`; }
        else { reportHtml += `<ul>`; recommendations.forEach(rec => { reportHtml += `<li>${rec}</li>`; }); reportHtml += `</ul>`; }
        reportOutputDiv.innerHTML = reportHtml;
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
                nodesToRemove.add(node);
                if (lastH1) nodesToRemove.add(lastH1);
            }
        });
        nodesToRemove.forEach(node => node.remove());
        return body.innerHTML;
    }

    // --- MAIN APPLICATION LOGIC ---
    document.addEventListener('DOMContentLoaded', () => {
        const uploadArea = document.getElementById('uploadArea');
        const uploadInput = document.getElementById('upload');
        const outputDiv = document.getElementById('output');
        const actionButtons = document.querySelector('.action-buttons');
        const copyHtmlButton = document.getElementById('copyHtmlButton');
        const downloadHtmlButton = document.getElementById('downloadHtmlButton');
        const outputReportButton = document.getElementById('outputReportButton');
        const cleanupButton = document.getElementById('cleanupButton');
        const undoButton = document.getElementById('undoButton');
        const reportOutputDiv = document.getElementById('reportOutput');
        const copyFeedbackDiv = document.getElementById('copyFeedback');
        const scrollToTopButton = document.getElementById('scrollToTopButton');
        const cleanupModal = document.getElementById('cleanupModal');
        const confirmCleanupButton = document.getElementById('confirmCleanupButton');
        const cancelCleanupButton = document.getElementById('cancelCleanupButton');
        const sessionRestoreBanner = document.getElementById('sessionRestoreBanner');
        const restoreSessionButton = document.getElementById('restoreSessionButton');
        const dismissRestoreButton = document.getElementById('dismissRestoreButton');
        let htmlBeforeCleanup = null;
        const SESSION_STORAGE_KEY = 'docxConverterSession';

        function saveSession(html) { try { localStorage.setItem(SESSION_STORAGE_KEY, html); } catch (e) { console.error("Could not save session to localStorage:", e); } }

        function setUIState(state, errorMessage = '') {
            actionButtons.style.display = (state === 'loaded') ? 'flex' : 'none';
            reportOutputDiv.style.display = 'none';
            cleanupModal.style.display = 'none';
            undoButton.style.display = 'none';
            if (cleanupButton) cleanupButton.style.display = 'inline-block';
            if (state === 'initial') outputDiv.innerHTML = '<p>To begin, drop a .docx file here or click to select one.</p>';
            else if (state === 'loading') outputDiv.innerHTML = '<p class="loading-message">Converting...</p>';
            else if (state === 'error') outputDiv.innerHTML = `<p class="error-message">Error: ${errorMessage}</p>`;
        }
        
        function handleFile(file) {
            if (!file || !file.type.match(/officedocument\.wordprocessingml\.document/)) { alert("Please select a valid .docx file."); return; }
            setUIState('loading');
            const reader = new FileReader();
            reader.onload = function(e) {
                const arrayBuffer = e.target.result;
                mammoth.convertToHtml({ arrayBuffer }, { styleMap: APP_CONFIG.styleMap })
                    .then(result => {
                        const finalHtml = processHtml(result.value);

                        // *** ONLY CHANGE REQUESTED: set border="1" on every table right before rendering ***
                        const parser2 = new DOMParser();
                        const tempDoc = parser2.parseFromString(finalHtml, 'text/html');
                        tempDoc.querySelectorAll('table').forEach(tbl => tbl.setAttribute('border', '1'));
                        const finalWithBorders = tempDoc.body.innerHTML;

                        outputDiv.innerHTML = finalWithBorders;
                        saveSession(finalWithBorders);
                        setUIState('loaded');
                    })
                    .catch(err => { console.error("Conversion Error:", err); setUIState('error', err.message); });
            };
            reader.readAsArrayBuffer(file);
        }

        uploadInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => { uploadArea.addEventListener(eventName, (e) => { e.preventDefault(); e.stopPropagation(); }, false); });
        ['dragenter', 'dragover'].forEach(eventName => { uploadArea.addEventListener(eventName, () => uploadArea.classList.add('drag-over'), false); });
        ['dragleave', 'drop'].forEach(eventName => { uploadArea.addEventListener(eventName, () => uploadArea.classList.remove('drag-over'), false); });
        uploadArea.addEventListener('drop', (e) => handleFile(e.dataTransfer.files[0]), false);
        
        copyHtmlButton.addEventListener('click', () => {
            const contentToCopy = outputDiv.innerHTML;
            if (!contentToCopy) return;
            if (!navigator.clipboard) { copyFeedbackDiv.textContent = 'Clipboard API not supported by your browser.'; copyFeedbackDiv.classList.add('show'); setTimeout(() => copyFeedbackDiv.classList.remove('show'), 3000); return; }
            navigator.clipboard.writeText(contentToCopy).then(() => { copyFeedbackDiv.textContent = 'HTML copied to clipboard!'; })
              .catch(err => { console.error('Failed to copy HTML:', err); copyFeedbackDiv.textContent = 'Error copying HTML. See console for details.'; })
              .finally(() => { copyFeedbackDiv.classList.add('show'); setTimeout(() => copyFeedbackDiv.classList.remove('show'), 2000); });
        });

        downloadHtmlButton.addEventListener('click', () => {
            const htmlContent = outputDiv.innerHTML; if (!htmlContent) return;
            const blob = new Blob([htmlContent], { type: 'text/html' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a'); a.href = url; a.download = 'converted-document.html';
            document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
        });

        outputReportButton.addEventListener('click', () => { generateSectionPresenceReport(outputDiv.innerHTML, reportOutputDiv); reportOutputDiv.scrollIntoView({ behavior: 'smooth', block: 'start' }); });

        cleanupButton.addEventListener('click', () => cleanupModal.style.display = 'flex');
        cancelCleanupButton.addEventListener('click', () => cleanupModal.style.display = 'none');
        confirmCleanupButton.addEventListener('click', () => {
            htmlBeforeCleanup = outputDiv.innerHTML;
            const cleanedHtml = removeExtraContent(htmlBeforeCleanup);
            outputDiv.innerHTML = cleanedHtml; saveSession(cleanedHtml);
            cleanupModal.style.display = 'none'; reportOutputDiv.style.display = 'none';
            cleanupButton.style.display = 'none'; undoButton.style.display = 'inline-block';
        });
        undoButton.addEventListener('click', () => { if (htmlBeforeCleanup) { outputDiv.innerHTML = htmlBeforeCleanup; saveSession(htmlBeforeCleanup); htmlBeforeCleanup = null; } undoButton.style.display = 'none'; cleanupButton.style.display = 'inline-block'; });

        reportOutputDiv.addEventListener('click', (e) => { if (e.target.matches('.close-button')) reportOutputDiv.style.display = 'none'; });
        window.addEventListener('scroll', () => { scrollToTopButton.classList.toggle('show', window.scrollY > 300); });
        scrollToTopButton.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));

        restoreSessionButton.addEventListener('click', () => { const savedHtml = localStorage.getItem(SESSION_STORAGE_KEY); if (savedHtml) { outputDiv.innerHTML = savedHtml; setUIState('loaded'); } sessionRestoreBanner.style.display = 'none'; });
        dismissRestoreButton.addEventListener('click', () => { localStorage.removeItem(SESSION_STORAGE_KEY); sessionRestoreBanner.style.display = 'none'; });
        if (localStorage.getItem(SESSION_STORAGE_KEY)) sessionRestoreBanner.style.display = 'block';
        setUIState('initial');
    });
  </script>
</body>
</html>
