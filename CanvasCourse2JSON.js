(async () => {
  /************************************************************
   * Canvas Granular Course JSON Exporter
   * Run from inside a Canvas course.
   *
   * Read-only. Uses only GET requests.
   *
   * This revised version fixes the modal scrolling issue:
   * - Modal is constrained to the visible browser window
   * - Options pane scrolls independently
   * - Export JSON button stays sticky at the bottom
   ************************************************************/

  const APP_ID = "canvas-granular-json-exporter";
  const courseId = getCourseIdFromUrl();

  if (!courseId) {
    alert("Could not detect a Canvas course ID. Open a course page, then run this again.");
    return;
  }

  const apiBase = `${location.origin}/api/v1/courses/${courseId}`;

  const appState = {
    course: null,
    modules: [],
    errors: [],
    warnings: [],
    requestCount: 0,
    startedAt: new Date().toISOString()
  };

  const caches = {
    pageByUrl: new Map(),
    assignmentById: new Map(),
    discussionById: new Map(),
    quizById: new Map(),
    fileById: new Map(),
    genericByUrl: new Map()
  };

  removeExistingApp();
  const shell = createShell();
  log("Loading course and module list...");

  try {
    appState.course = await safeGet(`${apiBase}`, "course metadata");
    appState.modules = await getPaginated(`${apiBase}/modules?per_page=100`);
    appState.modules.sort((a, b) => numberSort(a.position, b.position));

    renderChooser(shell);
    log(`Loaded ${appState.modules.length} modules. Choose what to export.`);
  } catch (error) {
    logError("Startup failed", error);
  }

  function renderChooser(shell) {
    const body = shell.querySelector("[data-cje-body]");
    body.innerHTML = "";

    const courseName = appState.course && appState.course.name ? appState.course.name : `Course ${courseId}`;

    const title = document.createElement("div");
    title.style.fontWeight = "700";
    title.style.fontSize = "15px";
    title.style.marginBottom = "6px";
    title.textContent = courseName;

    const hint = document.createElement("div");
    hint.style.fontSize = "12px";
    hint.style.color = "#555";
    hint.style.marginBottom = "12px";
    hint.textContent = "Select a preset, choose modules, adjust content types, then export JSON.";

    const presetLabel = makeLabel("Export preset");
    const preset = document.createElement("select");
    preset.id = "cje-preset";
    preset.style.width = "100%";
    preset.style.marginBottom = "12px";
    preset.style.padding = "8px";
    preset.style.border = "1px solid #ccc";
    preset.style.borderRadius = "6px";
    preset.innerHTML = `
      <option value="everything">Export everything</option>
      <option value="selected_all">Selected modules, all item types</option>
      <option value="selected_pages">Selected modules, pages only</option>
      <option value="selected_assignments">Selected modules, assignments only</option>
      <option value="selected_pages_assignments">Selected modules, pages and assignments</option>
      <option value="custom">Custom</option>
    `;

    const moduleHeader = document.createElement("div");
    moduleHeader.style.display = "flex";
    moduleHeader.style.alignItems = "center";
    moduleHeader.style.justifyContent = "space-between";
    moduleHeader.style.margin = "10px 0 6px";

    const moduleTitle = document.createElement("strong");
    moduleTitle.textContent = "Modules";

    const moduleButtons = document.createElement("div");
    moduleButtons.style.display = "flex";
    moduleButtons.style.gap = "6px";

    const allButton = makeSmallButton("All");
    const noneButton = makeSmallButton("None");
    const invertButton = makeSmallButton("Invert");

    moduleButtons.append(allButton, noneButton, invertButton);
    moduleHeader.append(moduleTitle, moduleButtons);

    const moduleList = document.createElement("div");
    moduleList.id = "cje-module-list";
    moduleList.style.border = "1px solid #ddd";
    moduleList.style.borderRadius = "8px";
    moduleList.style.maxHeight = "180px";
    moduleList.style.overflow = "auto";
    moduleList.style.padding = "6px";
    moduleList.style.background = "#fafafa";

    for (const module of appState.modules) {
      const row = document.createElement("label");
      row.style.display = "flex";
      row.style.gap = "8px";
      row.style.alignItems = "flex-start";
      row.style.padding = "6px";
      row.style.borderBottom = "1px solid #eee";
      row.style.cursor = "pointer";

      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "cje-module-checkbox";
      cb.value = String(module.id);
      cb.checked = true;
      cb.dataset.moduleName = module.name || "";
      cb.addEventListener("change", () => {
        preset.value = "custom";
      });

      const text = document.createElement("span");
      text.style.lineHeight = "1.3";
      text.innerHTML = `<strong>${escapeHtml(module.position || "")}.</strong> ${escapeHtml(module.name || "Untitled module")} <span style="color:#777;">(${escapeHtml(module.items_count ?? "unknown")} items)</span>`;

      row.append(cb, text);
      moduleList.append(row);
    }

    const typeHeader = document.createElement("div");
    typeHeader.style.margin = "14px 0 6px";
    typeHeader.innerHTML = "<strong>Module item types to include</strong>";

    const typeGrid = document.createElement("div");
    typeGrid.style.display = "grid";
    typeGrid.style.gridTemplateColumns = "1fr 1fr";
    typeGrid.style.gap = "4px 12px";
    typeGrid.style.marginBottom = "12px";

    const typeOptions = [
      ["Page", "Pages"],
      ["Assignment", "Assignments"],
      ["Discussion", "Discussions"],
      ["Quiz", "Classic quizzes"],
      ["File", "File metadata"],
      ["ExternalUrl", "External URLs"],
      ["ExternalTool", "External tools"],
      ["SubHeader", "Text headers"],
      ["Other", "Other item types"]
    ];

    for (const [value, label] of typeOptions) {
      typeGrid.append(makeCheckboxRow("cje-type-checkbox", value, label, true));
    }

    const detailHeader = document.createElement("div");
    detailHeader.style.margin = "14px 0 6px";
    detailHeader.innerHTML = "<strong>Detail options</strong>";

    const detailGrid = document.createElement("div");
    detailGrid.style.display = "grid";
    detailGrid.style.gridTemplateColumns = "1fr";
    detailGrid.style.gap = "4px";
    detailGrid.style.marginBottom = "12px";

    detailGrid.append(
      makeOptionCheckbox("cje-include-html", "Include HTML descriptions and page bodies", true),
      makeOptionCheckbox("cje-include-plain-text", "Include plain text versions for AI review", true),
      makeOptionCheckbox("cje-include-raw", "Include raw Canvas API objects", false),
      makeOptionCheckbox("cje-include-unpublished", "Include unpublished module items if Canvas returns them", true)
    );

    const inventoryHeader = document.createElement("div");
    inventoryHeader.style.margin = "14px 0 6px";
    inventoryHeader.innerHTML = "<strong>Optional full course inventories</strong>";

    const inventoryNote = document.createElement("div");
    inventoryNote.style.fontSize = "12px";
    inventoryNote.style.color = "#666";
    inventoryNote.style.marginBottom = "6px";
    inventoryNote.textContent = "These inventories are outside the selected module filter. Useful for finding orphaned pages or assignments.";

    const inventoryGrid = document.createElement("div");
    inventoryGrid.style.display = "grid";
    inventoryGrid.style.gridTemplateColumns = "1fr";
    inventoryGrid.style.gap = "4px";
    inventoryGrid.style.marginBottom = "12px";

    inventoryGrid.append(
      makeOptionCheckbox("cje-inventory-pages", "Full pages inventory", true),
      makeOptionCheckbox("cje-inventory-assignments", "Full assignments inventory", true),
      makeOptionCheckbox("cje-inventory-discussions", "Full discussions inventory", true),
      makeOptionCheckbox("cje-inventory-quizzes", "Full classic quizzes inventory", true)
    );

    const actionBar = document.createElement("div");
    actionBar.style.position = "sticky";
    actionBar.style.bottom = "0";
    actionBar.style.background = "#fff";
    actionBar.style.borderTop = "1px solid #e5e5e5";
    actionBar.style.padding = "12px 0 0";
    actionBar.style.marginTop = "14px";
    actionBar.style.boxShadow = "0 -8px 14px rgba(255,255,255,.92)";
    actionBar.style.zIndex = "5";

    const exportButton = document.createElement("button");
    exportButton.textContent = "Export JSON";
    exportButton.style.width = "100%";
    exportButton.style.padding = "10px 12px";
    exportButton.style.border = "0";
    exportButton.style.borderRadius = "8px";
    exportButton.style.background = "#510C76";
    exportButton.style.color = "#fff";
    exportButton.style.fontWeight = "700";
    exportButton.style.cursor = "pointer";

    const closeButton = document.createElement("button");
    closeButton.textContent = "Close";
    closeButton.style.width = "100%";
    closeButton.style.marginTop = "8px";
    closeButton.style.padding = "8px 12px";
    closeButton.style.border = "1px solid #ccc";
    closeButton.style.borderRadius = "8px";
    closeButton.style.background = "#f7f7f7";
    closeButton.style.cursor = "pointer";
    closeButton.addEventListener("click", removeExistingApp);

    actionBar.append(exportButton, closeButton);

    body.append(
      title,
      hint,
      presetLabel,
      preset,
      moduleHeader,
      moduleList,
      typeHeader,
      typeGrid,
      detailHeader,
      detailGrid,
      inventoryHeader,
      inventoryNote,
      inventoryGrid,
      actionBar
    );

    allButton.addEventListener("click", () => {
      setAllModuleChecks(true);
      preset.value = "custom";
    });

    noneButton.addEventListener("click", () => {
      setAllModuleChecks(false);
      preset.value = "custom";
    });

    invertButton.addEventListener("click", () => {
      for (const cb of document.querySelectorAll(".cje-module-checkbox")) {
        cb.checked = !cb.checked;
      }
      preset.value = "custom";
    });

    preset.addEventListener("change", () => {
      applyPreset(preset.value);
    });

    for (const cb of document.querySelectorAll(".cje-type-checkbox, .cje-option-checkbox")) {
      cb.addEventListener("change", () => {
        if (preset.value !== "custom") {
          preset.value = "custom";
        }
      });
    }

    exportButton.addEventListener("click", async () => {
      exportButton.disabled = true;
      exportButton.textContent = "Exporting...";
      try {
        await runExport();
      } finally {
        exportButton.disabled = false;
        exportButton.textContent = "Export JSON";
      }
    });

    applyPreset("everything");
  }

  async function runExport() {
    appState.errors = [];
    appState.warnings = [];

    const settings = readSettings();

    if (!settings.selectedModuleIds.length) {
      log("Select at least one module before exporting.");
      return;
    }

    if (!settings.contentTypes.size) {
      log("Select at least one module item type before exporting.");
      return;
    }

    log("Starting export...");

    const selectedModules = appState.modules.filter(module => settings.selectedModuleIds.includes(String(module.id)));

    const moduleItemCounts = {
      total_seen: 0,
      included: 0,
      skipped_by_type: 0,
      skipped_unpublished: 0
    };

    const exportedModules = [];
    const orderedContent = [];

    for (let i = 0; i < selectedModules.length; i++) {
      const module = selectedModules[i];

      log(`Getting items for module ${i + 1} of ${selectedModules.length}: ${module.name}`);

      const items = await getPaginated(
        `${apiBase}/modules/${encodeURIComponent(module.id)}/items?per_page=100&include[]=content_details`
      );

      items.sort((a, b) => numberSort(a.position, b.position));
      moduleItemCounts.total_seen += items.length;

      const exportedItems = [];

      for (const item of items) {
        const typeKey = knownType(item.type) ? item.type : "Other";

        if (!settings.includeUnpublished && item.published === false) {
          moduleItemCounts.skipped_unpublished++;
          continue;
        }

        if (!settings.contentTypes.has(typeKey)) {
          moduleItemCounts.skipped_by_type++;
          continue;
        }

        log(`Exporting: ${module.name} > ${item.title || item.id}`);

        const enriched = await enrichModuleItem(item, settings);
        moduleItemCounts.included++;

        const orderedEntry = {
          module_id: module.id,
          module_name: module.name,
          module_position: module.position,
          module_item_id: item.id,
          module_item_title: item.title,
          module_item_type: item.type,
          module_item_position: item.position,
          content: enriched.content
        };

        exportedItems.push(enriched);
        orderedContent.push(orderedEntry);
      }

      exportedModules.push(normalizeModule(module, exportedItems, settings));
    }

    const inventories = {
      pages: [],
      assignments: [],
      discussions: [],
      classic_quizzes: []
    };

    if (settings.fullInventories.pages) {
      log("Getting full pages inventory...");
      const pages = await getPaginated(`${apiBase}/pages?per_page=100&include[]=body&sort=title&order=asc`);
      inventories.pages = pages.map(page => normalizePage(page, settings));
    }

    if (settings.fullInventories.assignments) {
      log("Getting full assignments inventory...");
      const assignments = await getPaginated(
        `${apiBase}/assignments?per_page=100&include[]=all_dates&include[]=overrides&include[]=can_edit`
      );
      inventories.assignments = assignments.map(assignment => normalizeAssignment(assignment, settings));
    }

    if (settings.fullInventories.discussions) {
      log("Getting full discussions inventory...");
      const discussions = await getPaginated(`${apiBase}/discussion_topics?per_page=100`);
      inventories.discussions = discussions.map(discussion => normalizeDiscussion(discussion, settings));
    }

    if (settings.fullInventories.quizzes) {
      log("Getting full classic quizzes inventory...");
      const quizzes = await getPaginated(`${apiBase}/quizzes?per_page=100`);
      inventories.classic_quizzes = quizzes.map(quiz => normalizeClassicQuiz(quiz, settings));
    }

    const exportData = {
      export_version: "2.1-granular-scrollable",
      generated_at: new Date().toISOString(),
      generated_from: {
        canvas_host: location.host,
        course_id: courseId,
        course_url: `${location.origin}/courses/${courseId}`,
        current_page_url: location.href
      },
      export_settings: settingsForJson(settings),
      notes: [
        "This is a read-only export created from Canvas API GET requests.",
        "Module order and module item order are preserved.",
        "Selected module exports include only the module item types chosen in the panel.",
        "Full inventories are not filtered by selected modules.",
        "File exports include metadata only, not binary file contents.",
        "External URL and External Tool exports include link or launch metadata only, not external-site content.",
        "New Quizzes commonly appear as assignments or external tools. Internal New Quiz question content is not included here."
      ],
      stats: {
        selected_modules: selectedModules.length,
        exported_modules: exportedModules.length,
        module_items_total_seen: moduleItemCounts.total_seen,
        module_items_included: moduleItemCounts.included,
        module_items_skipped_by_type: moduleItemCounts.skipped_by_type,
        module_items_skipped_unpublished: moduleItemCounts.skipped_unpublished,
        full_inventory_pages: inventories.pages.length,
        full_inventory_assignments: inventories.assignments.length,
        full_inventory_discussions: inventories.discussions.length,
        full_inventory_classic_quizzes: inventories.classic_quizzes.length,
        api_requests: appState.requestCount,
        errors: appState.errors.length,
        warnings: appState.warnings.length
      },
      course: normalizeCourse(appState.course, settings),
      modules: exportedModules,
      ordered_content: orderedContent,
      inventories,
      warnings: appState.warnings,
      errors: appState.errors
    };

    log("Creating JSON download...");

    const fileName = makeFileName(appState.course, settings);
    downloadJson(exportData, fileName);

    log(`Done. Downloaded ${fileName}. Included ${moduleItemCounts.included} module items.`);
    console.log("Canvas granular JSON export:", exportData);
  }

  async function enrichModuleItem(item, settings) {
    const exported = {
      id: item.id,
      module_id: item.module_id,
      position: item.position,
      title: item.title,
      type: item.type,
      indent: item.indent,
      published: item.published,
      content_id: item.content_id || null,
      page_url: item.page_url || null,
      html_url: item.html_url || null,
      external_url: item.external_url || null,
      new_tab: item.new_tab || null,
      completion_requirement: item.completion_requirement || null,
      content_details: item.content_details || null,
      content: null,
      fetch_status: "metadata_only",
      fetch_error: null
    };

    if (settings.includeRaw) {
      exported.raw_module_item = item;
    }

    try {
      if (item.type === "Page") {
        const pageUrl = item.page_url || extractPageUrl(item);
        const page = pageUrl ? await getPage(pageUrl) : await getFromModuleItemUrl(item);
        exported.content = page ? normalizePage(page, settings) : null;
        exported.fetch_status = page ? "ok" : "not_found";
        return exported;
      }

      if (item.type === "Assignment") {
        const assignment = item.content_id ? await getAssignment(item.content_id) : await getFromModuleItemUrl(item);
        exported.content = assignment ? normalizeAssignment(assignment, settings) : null;
        exported.fetch_status = assignment ? "ok" : "not_found";
        return exported;
      }

      if (item.type === "Discussion") {
        const discussion = item.content_id ? await getDiscussion(item.content_id) : await getFromModuleItemUrl(item);
        exported.content = discussion ? normalizeDiscussion(discussion, settings) : null;
        exported.fetch_status = discussion ? "ok" : "not_found";
        return exported;
      }

      if (item.type === "Quiz") {
        const quiz = item.content_id ? await getClassicQuiz(item.content_id) : await getFromModuleItemUrl(item);
        exported.content = quiz ? normalizeClassicQuiz(quiz, settings) : null;
        exported.fetch_status = quiz ? "ok" : "not_found";
        return exported;
      }

      if (item.type === "File") {
        const file = item.content_id ? await getFile(item.content_id, item.url) : await getFromModuleItemUrl(item);
        exported.content = file ? normalizeFile(file, settings) : null;
        exported.fetch_status = file ? "ok" : "not_found";
        return exported;
      }

      if (item.type === "ExternalUrl" || item.type === "ExternalTool") {
        exported.content = {
          title: item.title || null,
          external_url: item.external_url || null,
          html_url: item.html_url || null,
          url: item.url || null,
          new_tab: item.new_tab || null,
          content_details: item.content_details || null
        };
        exported.fetch_status = "metadata_only";
        return exported;
      }

      if (item.type === "SubHeader") {
        exported.content = {
          title: item.title || null
        };
        exported.fetch_status = "not_applicable";
        return exported;
      }

      const generic = await getFromModuleItemUrl(item);
      exported.content = generic;
      exported.fetch_status = generic ? "ok_generic" : "metadata_only";
      return exported;
    } catch (error) {
      exported.fetch_status = "error";
      exported.fetch_error = error.message || String(error);
      appState.errors.push({
        context: "enrich module item",
        item_id: item.id,
        item_title: item.title,
        item_type: item.type,
        message: error.message || String(error)
      });
      return exported;
    }
  }

  async function getPage(pageUrl) {
    const key = String(pageUrl);

    if (caches.pageByUrl.has(key)) {
      return caches.pageByUrl.get(key);
    }

    const page = await safeGet(`${apiBase}/pages/${encodeURIComponent(key)}`, `page ${key}`);
    if (page) caches.pageByUrl.set(key, page);
    return page;
  }

  async function getAssignment(assignmentId) {
    const key = String(assignmentId);

    if (caches.assignmentById.has(key)) {
      return caches.assignmentById.get(key);
    }

    const assignment = await safeGet(
      `${apiBase}/assignments/${encodeURIComponent(key)}?include[]=all_dates&include[]=overrides&include[]=can_edit`,
      `assignment ${key}`
    );

    if (assignment) caches.assignmentById.set(key, assignment);
    return assignment;
  }

  async function getDiscussion(discussionId) {
    const key = String(discussionId);

    if (caches.discussionById.has(key)) {
      return caches.discussionById.get(key);
    }

    const discussion = await safeGet(`${apiBase}/discussion_topics/${encodeURIComponent(key)}`, `discussion ${key}`);
    if (discussion) caches.discussionById.set(key, discussion);
    return discussion;
  }

  async function getClassicQuiz(quizId) {
    const key = String(quizId);

    if (caches.quizById.has(key)) {
      return caches.quizById.get(key);
    }

    const quiz = await safeGet(`${apiBase}/quizzes/${encodeURIComponent(key)}`, `classic quiz ${key}`);
    if (quiz) caches.quizById.set(key, quiz);
    return quiz;
  }

  async function getFile(fileId, itemUrl) {
    const key = String(fileId);

    if (caches.fileById.has(key)) {
      return caches.fileById.get(key);
    }

    let file = null;

    if (itemUrl && String(itemUrl).includes("/api/v1/")) {
      file = await safeGet(itemUrl, `file ${key}`);
    }

    if (!file) {
      file = await safeGet(`${location.origin}/api/v1/files/${encodeURIComponent(key)}`, `file ${key}`);
    }

    if (file) caches.fileById.set(key, file);
    return file;
  }

  async function getFromModuleItemUrl(item) {
    if (!item || !item.url || !String(item.url).includes("/api/v1/")) {
      return null;
    }

    const url = makeAbsoluteUrl(item.url);

    if (caches.genericByUrl.has(url)) {
      return caches.genericByUrl.get(url);
    }

    const data = await safeGet(url, `module item URL ${item.title || item.id}`);
    if (data) caches.genericByUrl.set(url, data);
    return data;
  }

  function normalizeCourse(course, settings) {
    if (!course) return null;

    const normalized = {
      id: course.id,
      name: course.name,
      course_code: course.course_code,
      workflow_state: course.workflow_state,
      start_at: course.start_at,
      end_at: course.end_at,
      time_zone: course.time_zone,
      default_view: course.default_view,
      public_syllabus: course.public_syllabus,
      syllabus_body_plain_text: settings.includePlainText ? htmlToText(course.syllabus_body || "") : null
    };

    if (settings.includeHtml) {
      normalized.syllabus_body_html = course.syllabus_body || "";
    }

    if (settings.includeRaw) {
      normalized.raw_course = course;
    }

    return normalized;
  }

  function normalizeModule(module, items, settings) {
    const normalized = {
      id: module.id,
      name: module.name,
      position: module.position,
      published: module.published,
      workflow_state: module.workflow_state,
      unlock_at: module.unlock_at,
      require_sequential_progress: module.require_sequential_progress,
      requirement_type: module.requirement_type,
      prerequisite_module_ids: module.prerequisite_module_ids || [],
      items_count_reported_by_canvas: module.items_count,
      exported_items_count: items.length,
      items
    };

    if (settings.includeRaw) {
      normalized.raw_module = module;
    }

    return normalized;
  }

  function normalizePage(page, settings) {
    const body = page.body || "";

    const normalized = {
      page_id: page.page_id,
      url: page.url,
      title: page.title,
      published: page.published,
      front_page: page.front_page,
      editor: page.editor || null,
      created_at: page.created_at,
      updated_at: page.updated_at,
      html_url: page.html_url || null
    };

    if (settings.includeHtml) {
      normalized.body_html = body;
    }

    if (settings.includePlainText) {
      normalized.body_plain_text = htmlToText(body);
    }

    if (settings.includeRaw) {
      normalized.raw_page = page;
    }

    return normalized;
  }

  function normalizeAssignment(assignment, settings) {
    const description = assignment.description || "";

    const normalized = {
      id: assignment.id,
      name: assignment.name,
      published: assignment.published,
      workflow_state: assignment.workflow_state,
      points_possible: assignment.points_possible,
      grading_type: assignment.grading_type,
      assignment_group_id: assignment.assignment_group_id,
      due_at: assignment.due_at,
      unlock_at: assignment.unlock_at,
      lock_at: assignment.lock_at,
      all_dates: assignment.all_dates || null,
      overrides: assignment.overrides || null,
      submission_types: assignment.submission_types || [],
      allowed_extensions: assignment.allowed_extensions || [],
      peer_reviews: assignment.peer_reviews || false,
      automatic_peer_reviews: assignment.automatic_peer_reviews || false,
      group_category_id: assignment.group_category_id || null,
      grade_group_students_individually: assignment.grade_group_students_individually || false,
      anonymous_grading: assignment.anonymous_grading || false,
      moderated_grading: assignment.moderated_grading || false,
      omit_from_final_grade: assignment.omit_from_final_grade || false,
      html_url: assignment.html_url || null,
      rubric: assignment.rubric || null,
      rubric_settings: assignment.rubric_settings || null
    };

    if (settings.includeHtml) {
      normalized.description_html = description;
    }

    if (settings.includePlainText) {
      normalized.description_plain_text = htmlToText(description);
    }

    if (settings.includeRaw) {
      normalized.raw_assignment = assignment;
    }

    return normalized;
  }

  function normalizeDiscussion(discussion, settings) {
    const message = discussion.message || discussion.description || "";

    const normalized = {
      id: discussion.id,
      title: discussion.title,
      published: discussion.published,
      workflow_state: discussion.workflow_state,
      posted_at: discussion.posted_at,
      delayed_post_at: discussion.delayed_post_at,
      lock_at: discussion.lock_at,
      assignment_id: discussion.assignment_id || null,
      discussion_type: discussion.discussion_type || null,
      pinned: discussion.pinned || false,
      locked: discussion.locked || false,
      html_url: discussion.html_url || null
    };

    if (settings.includeHtml) {
      normalized.message_html = message;
    }

    if (settings.includePlainText) {
      normalized.message_plain_text = htmlToText(message);
    }

    if (settings.includeRaw) {
      normalized.raw_discussion = discussion;
    }

    return normalized;
  }

  function normalizeClassicQuiz(quiz, settings) {
    const description = quiz.description || "";

    const normalized = {
      id: quiz.id,
      title: quiz.title,
      quiz_type: quiz.quiz_type,
      published: quiz.published,
      workflow_state: quiz.workflow_state,
      assignment_id: quiz.assignment_id || null,
      points_possible: quiz.points_possible,
      due_at: quiz.due_at,
      unlock_at: quiz.unlock_at,
      lock_at: quiz.lock_at,
      time_limit: quiz.time_limit,
      allowed_attempts: quiz.allowed_attempts,
      scoring_policy: quiz.scoring_policy,
      shuffle_answers: quiz.shuffle_answers,
      hide_results: quiz.hide_results,
      show_correct_answers: quiz.show_correct_answers,
      html_url: quiz.html_url || null
    };

    if (settings.includeHtml) {
      normalized.description_html = description;
    }

    if (settings.includePlainText) {
      normalized.description_plain_text = htmlToText(description);
    }

    if (settings.includeRaw) {
      normalized.raw_quiz = quiz;
    }

    return normalized;
  }

  function normalizeFile(file, settings) {
    const normalized = {
      id: file.id,
      uuid: file.uuid,
      display_name: file.display_name,
      filename: file.filename,
      content_type: file.content_type,
      size: file.size,
      folder_id: file.folder_id,
      url: file.url || null,
      preview_url: file.preview_url || null,
      created_at: file.created_at,
      updated_at: file.updated_at,
      unlock_at: file.unlock_at,
      locked: file.locked,
      hidden: file.hidden,
      locked_for_user: file.locked_for_user
    };

    if (settings.includeRaw) {
      normalized.raw_file = file;
    }

    return normalized;
  }

  async function canvasApiRequest(url) {
    appState.requestCount++;

    const response = await fetch(makeAbsoluteUrl(url), {
      method: "GET",
      credentials: "same-origin",
      headers: {
        "Accept": "application/json+canvas-string-ids"
      }
    });

    const text = await response.text();
    let data = null;

    if (text) {
      try {
        data = JSON.parse(text);
      } catch (error) {
        throw new Error(`Non-JSON response from ${url}: ${text.slice(0, 400)}`);
      }
    }

    if (!response.ok) {
      throw new Error(`Canvas API error ${response.status} ${response.statusText} from ${url}: ${text.slice(0, 700)}`);
    }

    return {
      data,
      headers: response.headers
    };
  }

  async function getPaginated(url) {
    let nextUrl = makeAbsoluteUrl(url);
    const results = [];

    while (nextUrl) {
      const { data, headers } = await canvasApiRequest(nextUrl);

      if (!Array.isArray(data)) {
        throw new Error(`Expected an array from ${nextUrl}`);
      }

      results.push(...data);
      nextUrl = getNextLink(headers.get("Link"));
    }

    return results;
  }

  async function safeGet(url, context) {
    try {
      const { data } = await canvasApiRequest(url);
      return data;
    } catch (error) {
      appState.errors.push({
        context,
        url,
        message: error.message || String(error)
      });
      log(`Could not get ${context}. Continuing.`);
      console.warn(`Canvas export warning for ${context}:`, error);
      return null;
    }
  }

  function getNextLink(linkHeader) {
    if (!linkHeader) return null;

    const links = linkHeader.split(",");
    for (const link of links) {
      const match = link.match(/<([^>]+)>;\s*rel="next"/);
      if (match) return match[1];
    }

    return null;
  }

  function makeAbsoluteUrl(url) {
    if (!url) return url;

    const value = String(url);

    if (value.startsWith("http://") || value.startsWith("https://")) {
      return value;
    }

    if (value.startsWith("/")) {
      return `${location.origin}${value}`;
    }

    return `${location.origin}/${value}`;
  }

  function readSettings() {
    const selectedModuleIds = Array.from(document.querySelectorAll(".cje-module-checkbox:checked"))
      .map(cb => cb.value);

    const contentTypes = new Set(
      Array.from(document.querySelectorAll(".cje-type-checkbox:checked"))
        .map(cb => cb.value)
    );

    return {
      preset: document.getElementById("cje-preset").value,
      selectedModuleIds,
      selectedModuleNames: selectedModuleIds.map(id => {
        const module = appState.modules.find(m => String(m.id) === String(id));
        return module ? module.name : id;
      }),
      contentTypes,
      includeHtml: document.getElementById("cje-include-html").checked,
      includePlainText: document.getElementById("cje-include-plain-text").checked,
      includeRaw: document.getElementById("cje-include-raw").checked,
      includeUnpublished: document.getElementById("cje-include-unpublished").checked,
      fullInventories: {
        pages: document.getElementById("cje-inventory-pages").checked,
        assignments: document.getElementById("cje-inventory-assignments").checked,
        discussions: document.getElementById("cje-inventory-discussions").checked,
        quizzes: document.getElementById("cje-inventory-quizzes").checked
      }
    };
  }

  function settingsForJson(settings) {
    return {
      preset: settings.preset,
      selectedModuleIds: settings.selectedModuleIds,
      selectedModuleNames: settings.selectedModuleNames,
      contentTypes: Array.from(settings.contentTypes),
      includeHtml: settings.includeHtml,
      includePlainText: settings.includePlainText,
      includeRaw: settings.includeRaw,
      includeUnpublished: settings.includeUnpublished,
      fullInventories: settings.fullInventories
    };
  }

  function applyPreset(presetValue) {
    const typeChecks = Array.from(document.querySelectorAll(".cje-type-checkbox"));
    const moduleChecks = Array.from(document.querySelectorAll(".cje-module-checkbox"));

    if (presetValue === "everything") {
      moduleChecks.forEach(cb => cb.checked = true);
      typeChecks.forEach(cb => cb.checked = true);
      setChecked("cje-inventory-pages", true);
      setChecked("cje-inventory-assignments", true);
      setChecked("cje-inventory-discussions", true);
      setChecked("cje-inventory-quizzes", true);
      setChecked("cje-include-html", true);
      setChecked("cje-include-plain-text", true);
      return;
    }

    if (presetValue === "selected_all") {
      typeChecks.forEach(cb => cb.checked = true);
      setInventories(false);
      return;
    }

    if (presetValue === "selected_pages") {
      setContentTypes(["Page"]);
      setInventories(false);
      return;
    }

    if (presetValue === "selected_assignments") {
      setContentTypes(["Assignment"]);
      setInventories(false);
      return;
    }

    if (presetValue === "selected_pages_assignments") {
      setContentTypes(["Page", "Assignment"]);
      setInventories(false);
    }
  }

  function setContentTypes(types) {
    const wanted = new Set(types);
    for (const cb of document.querySelectorAll(".cje-type-checkbox")) {
      cb.checked = wanted.has(cb.value);
    }
  }

  function setInventories(value) {
    setChecked("cje-inventory-pages", value);
    setChecked("cje-inventory-assignments", value);
    setChecked("cje-inventory-discussions", value);
    setChecked("cje-inventory-quizzes", value);
  }

  function setChecked(id, value) {
    const el = document.getElementById(id);
    if (el) el.checked = value;
  }

  function setAllModuleChecks(value) {
    for (const cb of document.querySelectorAll(".cje-module-checkbox")) {
      cb.checked = value;
    }
  }

  function knownType(type) {
    return ["Page", "Assignment", "Discussion", "Quiz", "File", "ExternalUrl", "ExternalTool", "SubHeader"].includes(type);
  }

  function extractPageUrl(item) {
    if (item.page_url) return item.page_url;

    for (const possible of [item.html_url, item.url]) {
      if (!possible) continue;
      const match = String(possible).match(/\/pages\/([^/?#]+)/);
      if (match) return decodeURIComponent(match[1]);
    }

    return null;
  }

  function htmlToText(html) {
    if (!html) return "";

    const doc = new DOMParser().parseFromString(String(html), "text/html");

    doc.querySelectorAll("script, style, noscript").forEach(el => el.remove());

    return (doc.body ? doc.body.textContent || "" : "")
      .replace(/\u00a0/g, " ")
      .replace(/[ \t]+/g, " ")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function downloadJson(data, fileName) {
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();

    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  function makeFileName(course, settings) {
    const courseName = course && course.name ? course.name : `course-${courseId}`;
    const safeName = String(courseName)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 80);

    const preset = String(settings.preset || "custom").replace(/[^a-z0-9]+/gi, "-").toLowerCase();
    const date = new Date().toISOString().replace(/[:.]/g, "-");

    return `canvas-course-export-${safeName}-${courseId}-${preset}-${date}.json`;
  }

  function createShell() {
    const overlay = document.createElement("div");
    overlay.id = APP_ID;
    overlay.style.position = "fixed";
    overlay.style.inset = "0";
    overlay.style.background = "rgba(0,0,0,.35)";
    overlay.style.zIndex = "999999";
    overlay.style.display = "flex";
    overlay.style.alignItems = "center";
    overlay.style.justifyContent = "center";
    overlay.style.padding = "18px";
    overlay.style.boxSizing = "border-box";

    const panel = document.createElement("div");
    panel.style.width = "820px";
    panel.style.maxWidth = "calc(100vw - 36px)";
    panel.style.height = "calc(100vh - 36px)";
    panel.style.maxHeight = "calc(100vh - 36px)";
    panel.style.background = "#fff";
    panel.style.borderRadius = "14px";
    panel.style.boxShadow = "0 18px 50px rgba(0,0,0,.3)";
    panel.style.overflow = "hidden";
    panel.style.fontFamily = "Arial, Helvetica, sans-serif";
    panel.style.color = "#222";
    panel.style.display = "flex";
    panel.style.flexDirection = "column";

    const header = document.createElement("div");
    header.style.flex = "0 0 auto";
    header.style.display = "flex";
    header.style.alignItems = "center";
    header.style.justifyContent = "space-between";
    header.style.padding = "14px 16px";
    header.style.background = "#510C76";
    header.style.color = "#fff";

    const headerTitle = document.createElement("div");
    headerTitle.style.fontWeight = "700";
    headerTitle.textContent = "Canvas Granular JSON Exporter";

    const close = document.createElement("button");
    close.textContent = "×";
    close.setAttribute("aria-label", "Close");
    close.style.border = "0";
    close.style.background = "transparent";
    close.style.color = "#fff";
    close.style.fontSize = "24px";
    close.style.cursor = "pointer";
    close.addEventListener("click", removeExistingApp);

    header.append(headerTitle, close);

    const content = document.createElement("div");
    content.style.flex = "1 1 auto";
    content.style.minHeight = "0";
    content.style.display = "grid";
    content.style.gridTemplateColumns = "minmax(0, 1fr) 300px";
    content.style.overflow = "hidden";

    const body = document.createElement("div");
    body.setAttribute("data-cje-body", "true");
    body.style.padding = "16px";
    body.style.overflowY = "auto";
    body.style.overflowX = "hidden";
    body.style.minHeight = "0";
    body.style.boxSizing = "border-box";

    const logWrap = document.createElement("div");
    logWrap.style.borderLeft = "1px solid #eee";
    logWrap.style.background = "#fafafa";
    logWrap.style.padding = "12px";
    logWrap.style.overflowY = "auto";
    logWrap.style.overflowX = "hidden";
    logWrap.style.minHeight = "0";
    logWrap.style.boxSizing = "border-box";

    const logTitle = document.createElement("div");
    logTitle.style.fontWeight = "700";
    logTitle.style.marginBottom = "8px";
    logTitle.textContent = "Status";

    const logBox = document.createElement("pre");
    logBox.id = "cje-log";
    logBox.style.whiteSpace = "pre-wrap";
    logBox.style.overflowWrap = "anywhere";
    logBox.style.fontFamily = "Consolas, Monaco, monospace";
    logBox.style.fontSize = "12px";
    logBox.style.margin = "0";
    logBox.style.color = "#333";

    logWrap.append(logTitle, logBox);
    content.append(body, logWrap);
    panel.append(header, content);
    overlay.append(panel);
    document.body.appendChild(overlay);

    return overlay;
  }

  function makeLabel(text) {
    const label = document.createElement("label");
    label.style.display = "block";
    label.style.fontWeight = "700";
    label.style.marginBottom = "4px";
    label.textContent = text;
    return label;
  }

  function makeSmallButton(text) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = text;
    button.style.fontSize = "12px";
    button.style.padding = "4px 7px";
    button.style.border = "1px solid #ccc";
    button.style.borderRadius = "6px";
    button.style.background = "#fff";
    button.style.cursor = "pointer";
    return button;
  }

  function makeCheckboxRow(className, value, labelText, checked) {
    const label = document.createElement("label");
    label.style.display = "flex";
    label.style.alignItems = "center";
    label.style.gap = "6px";
    label.style.cursor = "pointer";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = className;
    cb.value = value;
    cb.checked = checked;

    const text = document.createElement("span");
    text.textContent = labelText;

    label.append(cb, text);
    return label;
  }

  function makeOptionCheckbox(id, labelText, checked) {
    const label = document.createElement("label");
    label.style.display = "flex";
    label.style.alignItems = "center";
    label.style.gap = "6px";
    label.style.cursor = "pointer";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.id = id;
    cb.className = "cje-option-checkbox";
    cb.checked = checked;

    const text = document.createElement("span");
    text.textContent = labelText;

    label.append(cb, text);
    return label;
  }

  function log(message) {
    const logBox = document.getElementById("cje-log");
    const stamp = new Date().toLocaleTimeString();
    const line = `[${stamp}] ${message}`;

    if (logBox) {
      logBox.textContent += `${line}\n`;
      logBox.scrollTop = logBox.scrollHeight;
    }

    console.log(line);
  }

  function logError(message, error) {
    const detail = error && error.message ? error.message : String(error);
    appState.errors.push({
      context: message,
      message: detail
    });
    log(`${message}: ${detail}`);
    console.error(message, error);
  }

  function getCourseIdFromUrl() {
    const match = location.pathname.match(/\/courses\/(\d+)/);
    return match ? match[1] : null;
  }

  function removeExistingApp() {
    const existing = document.getElementById(APP_ID);
    if (existing) existing.remove();
  }

  function numberSort(a, b) {
    return Number(a || 0) - Number(b || 0);
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }
})();
