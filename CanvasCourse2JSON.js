(async () => {
  /************************************************************
   * Canvas Course AI JSON Exporter
   * Version 3.0.1
   *
   * Run from inside a Canvas course.
   * Read-only. Uses GET requests only.
   *
   * Design goals:
   * - Preserve module order and course structure
   * - Store Canvas content once in a canonical content_store
   * - Export Classic Quiz questions
   * - Export New Quiz metadata and items when available
   * - Export assignment groups, rubrics, outcomes, and navigation
   * - Build a manifest of Canvas files referenced by course content
   * - Do NOT download binary files
   * - Keep original HTML plus AI-readable text
   * - Keep raw Canvas API objects OFF by default
   ************************************************************/

  const APP_ID = "canvas-course-ai-json-exporter";
  const EXPORT_VERSION = "3.0.1-ai-review";

  const courseId = getCourseIdFromUrl();

  if (!courseId) {
    alert(
      "Could not detect a Canvas course ID. Open a course page, then run this again."
    );
    return;
  }

  const apiBase = `${location.origin}/api/v1/courses/${courseId}`;
  const newQuizApiBase = `${location.origin}/api/quiz/v1/courses/${courseId}`;

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
    classicQuizById: new Map(),
    classicQuizQuestionsById: new Map(),
    fileById: new Map(),
    genericByUrl: new Map()
  };

  const contentStore = {
    pages: new Map(),
    assignments: new Map(),
    discussions: new Map(),
    classic_quizzes: new Map(),
    new_quizzes: new Map(),
    files: new Map(),
    generic: new Map()
  };

  const fileReferences = new Map();

  removeExistingApp();

  const shell = createShell();

  log("Loading course and module list...");

  try {
    appState.course = await safeGet(
      `${apiBase}?include[]=syllabus_body&include[]=storage_quota_used_mb&include[]=post_manually`,
      "course metadata"
    );

    appState.modules = await safeGetPaginated(
      `${apiBase}/modules?per_page=100`,
      "course modules"
    );

    appState.modules.sort((a, b) =>
      numberSort(a.position, b.position)
    );

    renderChooser(shell);

    log(
      `Loaded ${appState.modules.length} modules. Choose what to export.`
    );
  } catch (error) {
    logError("Startup failed", error);
  }

  /************************************************************
   * USER INTERFACE
   ************************************************************/

  function renderChooser(shell) {
    const body = shell.querySelector("[data-cje-body]");

    body.innerHTML = "";

    const courseName =
      appState.course?.name || `Course ${courseId}`;

    const title = document.createElement("div");
    title.style.fontWeight = "700";
    title.style.fontSize = "15px";
    title.style.marginBottom = "6px";
    title.textContent = courseName;

    const hint = document.createElement("div");
    hint.style.fontSize = "12px";
    hint.style.color = "#555";
    hint.style.marginBottom = "12px";
    hint.textContent =
      "Recommended preset creates an AI-readable course map plus a referenced-file manifest. It does not download binary files.";

    const presetLabel = makeLabel("Export preset");

    const preset = document.createElement("select");
    preset.id = "cje-preset";
    preset.style.width = "100%";
    preset.style.marginBottom = "12px";
    preset.style.padding = "8px";
    preset.style.border = "1px solid #ccc";
    preset.style.borderRadius = "6px";

    preset.innerHTML = `
      <option value="ai_review">AI course review - recommended</option>
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

    moduleButtons.append(
      allButton,
      noneButton,
      invertButton
    );

    moduleHeader.append(
      moduleTitle,
      moduleButtons
    );

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

      cb.addEventListener(
        "change",
        markCustom
      );

      const text = document.createElement("span");

      text.style.lineHeight = "1.3";

      text.innerHTML =
        `<strong>${escapeHtml(module.position || "")}.</strong> ` +
        `${escapeHtml(module.name || "Untitled module")} ` +
        `<span style="color:#777;">(${escapeHtml(
          module.items_count ?? "unknown"
        )} items)</span>`;

      row.append(
        cb,
        text
      );

      moduleList.append(row);
    }

    const typeHeader = document.createElement("div");
    typeHeader.style.margin = "14px 0 6px";
    typeHeader.innerHTML =
      "<strong>Module item types to include</strong>";

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
      typeGrid.append(
        makeCheckboxRow(
          "cje-type-checkbox",
          value,
          label,
          true
        )
      );
    }

    const detailHeader = document.createElement("div");
    detailHeader.style.margin = "14px 0 6px";
    detailHeader.innerHTML =
      "<strong>Detail options</strong>";

    const detailGrid = document.createElement("div");

    detailGrid.style.display = "grid";
    detailGrid.style.gridTemplateColumns = "1fr";
    detailGrid.style.gap = "4px";
    detailGrid.style.marginBottom = "12px";

    detailGrid.append(
      makeOptionCheckbox(
        "cje-include-html",
        "Include original HTML",
        true
      ),

      makeOptionCheckbox(
        "cje-include-ai-text",
        "Include AI-readable text",
        true
      ),

      makeOptionCheckbox(
        "cje-include-raw",
        "Include raw Canvas API objects",
        false
      ),

      makeOptionCheckbox(
        "cje-include-unpublished",
        "Include unpublished module items if Canvas returns them",
        true
      ),

      makeOptionCheckbox(
        "cje-classic-quiz-questions",
        "Include Classic Quiz questions",
        true
      ),

      makeOptionCheckbox(
        "cje-new-quiz-items",
        "Include New Quiz items",
        true
      )
    );

    const inventoryHeader =
      document.createElement("div");

    inventoryHeader.style.margin =
      "14px 0 6px";

    inventoryHeader.innerHTML =
      "<strong>Course-wide data</strong>";

    const inventoryNote =
      document.createElement("div");

    inventoryNote.style.fontSize = "12px";
    inventoryNote.style.color = "#666";
    inventoryNote.style.marginBottom = "6px";

    inventoryNote.textContent =
      "These are not limited to selected modules. They help identify orphaned content and explain grading, navigation, outcomes, and assessments.";

    const inventoryGrid =
      document.createElement("div");

    inventoryGrid.style.display = "grid";
    inventoryGrid.style.gridTemplateColumns = "1fr";
    inventoryGrid.style.gap = "4px";
    inventoryGrid.style.marginBottom = "12px";

    inventoryGrid.append(
      makeOptionCheckbox(
        "cje-inventory-pages",
        "Full pages inventory",
        true
      ),

      makeOptionCheckbox(
        "cje-inventory-assignments",
        "Full assignments inventory",
        true
      ),

      makeOptionCheckbox(
        "cje-inventory-discussions",
        "Full discussions inventory",
        true
      ),

      makeOptionCheckbox(
        "cje-inventory-classic-quizzes",
        "Full Classic Quizzes inventory",
        true
      ),

      makeOptionCheckbox(
        "cje-inventory-new-quizzes",
        "Full New Quizzes inventory",
        true
      ),

      makeOptionCheckbox(
        "cje-inventory-assignment-groups",
        "Assignment groups and weighting",
        true
      ),

      makeOptionCheckbox(
        "cje-inventory-rubrics",
        "Course rubrics",
        true
      ),

      makeOptionCheckbox(
        "cje-inventory-outcomes",
        "Course outcome links",
        true
      ),

      makeOptionCheckbox(
        "cje-inventory-tabs",
        "Course navigation tabs",
        true
      ),

      makeOptionCheckbox(
        "cje-referenced-files",
        "Referenced Canvas file manifest",
        true
      ),

      makeOptionCheckbox(
        "cje-all-files",
        "Full Canvas Files inventory - metadata only",
        false
      )
    );

    const actionBar =
      document.createElement("div");

    actionBar.style.position = "sticky";
    actionBar.style.bottom = "0";
    actionBar.style.background = "#fff";
    actionBar.style.borderTop =
      "1px solid #e5e5e5";
    actionBar.style.padding = "12px 0 0";
    actionBar.style.marginTop = "14px";
    actionBar.style.boxShadow =
      "0 -8px 14px rgba(255,255,255,.92)";
    actionBar.style.zIndex = "5";

    const exportButton =
      document.createElement("button");

    exportButton.textContent = "Export JSON";
    exportButton.style.width = "100%";
    exportButton.style.padding = "10px 12px";
    exportButton.style.border = "0";
    exportButton.style.borderRadius = "8px";
    exportButton.style.background = "#510C76";
    exportButton.style.color = "#fff";
    exportButton.style.fontWeight = "700";
    exportButton.style.cursor = "pointer";

    const closeButton =
      document.createElement("button");

    closeButton.textContent = "Close";
    closeButton.style.width = "100%";
    closeButton.style.marginTop = "8px";
    closeButton.style.padding = "8px 12px";
    closeButton.style.border = "1px solid #ccc";
    closeButton.style.borderRadius = "8px";
    closeButton.style.background = "#f7f7f7";
    closeButton.style.cursor = "pointer";

    closeButton.addEventListener(
      "click",
      removeExistingApp
    );

    actionBar.append(
      exportButton,
      closeButton
    );

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

    allButton.addEventListener(
      "click",
      () => {
        setAllModuleChecks(true);
        markCustom();
      }
    );

    noneButton.addEventListener(
      "click",
      () => {
        setAllModuleChecks(false);
        markCustom();
      }
    );

    invertButton.addEventListener(
      "click",
      () => {
        for (
          const cb of document.querySelectorAll(
            ".cje-module-checkbox"
          )
        ) {
          cb.checked = !cb.checked;
        }

        markCustom();
      }
    );

    preset.addEventListener(
      "change",
      () => {
        applyPreset(preset.value);
      }
    );

    for (
      const cb of document.querySelectorAll(
        ".cje-type-checkbox, .cje-option-checkbox"
      )
    ) {
      cb.addEventListener(
        "change",
        markCustom
      );
    }

    exportButton.addEventListener(
      "click",
      async () => {
        exportButton.disabled = true;
        exportButton.textContent = "Exporting...";

        try {
          await runExport();
        } finally {
          exportButton.disabled = false;
          exportButton.textContent = "Export JSON";
        }
      }
    );

    applyPreset("ai_review");

    function markCustom() {
      if (preset.value !== "custom") {
        preset.value = "custom";
      }
    }
  }

  /************************************************************
   * MAIN EXPORT
   ************************************************************/

  async function runExport() {
    resetExportStores();

    const settings = readSettings();

    if (!settings.selectedModuleIds.length) {
      log(
        "Select at least one module before exporting."
      );
      return;
    }

    if (!settings.contentTypes.size) {
      log(
        "Select at least one module item type before exporting."
      );
      return;
    }

    log("Starting export...");

    const normalizedCourse =
      normalizeCourse(
        appState.course,
        settings
      );

    const selectedModules =
      appState.modules.filter(module =>
        settings.selectedModuleIds.includes(
          String(module.id)
        )
      );

    const moduleItemCounts = {
      total_seen: 0,
      included: 0,
      skipped_by_type: 0,
      skipped_unpublished: 0
    };

    const exportedModules = [];
    const orderedContent = [];

    for (
      let i = 0;
      i < selectedModules.length;
      i++
    ) {
      const module =
        selectedModules[i];

      log(
        `Getting items for module ${i + 1} of ${selectedModules.length}: ${module.name}`
      );

      const items =
        await safeGetPaginated(
          `${apiBase}/modules/${encodeURIComponent(
            module.id
          )}/items?per_page=100&include[]=content_details`,
          `module items for ${module.name}`
        );

      items.sort((a, b) =>
        numberSort(a.position, b.position)
      );

      moduleItemCounts.total_seen +=
        items.length;

      const exportedItems = [];

      for (const item of items) {
        const typeKey =
          knownType(item.type)
            ? item.type
            : "Other";

        if (
          !settings.includeUnpublished &&
          item.published === false
        ) {
          moduleItemCounts.skipped_unpublished++;
          continue;
        }

        if (
          !settings.contentTypes.has(typeKey)
        ) {
          moduleItemCounts.skipped_by_type++;
          continue;
        }

        log(
          `Exporting: ${module.name} > ${item.title || item.id}`
        );

        const enriched =
          await enrichModuleItem(
            item,
            settings,
            module
          );

        moduleItemCounts.included++;

        exportedItems.push(enriched);

        orderedContent.push({
          module_id:
            module.id,

          module_name:
            module.name,

          module_position:
            module.position,

          module_item_id:
            item.id,

          module_item_title:
            item.title,

          module_item_type:
            item.type,

          module_item_position:
            item.position,

          content_ref:
            enriched.content_ref || null,

          inline_content:
            enriched.inline_content || null
        });
      }

      exportedModules.push(
        normalizeModule(
          module,
          exportedItems,
          settings
        )
      );
    }

    const inventories = {
      pages: [],
      assignments: [],
      discussions: [],
      classic_quizzes: [],
      new_quizzes: []
    };

    const courseStructure = {
      assignment_groups: [],
      rubrics: [],
      outcomes: [],
      tabs: []
    };

    /**********************************************************
     * FULL PAGES INVENTORY
     **********************************************************/

    if (
      settings.fullInventories.pages
    ) {
      log(
        "Getting full pages inventory..."
      );

      const pages =
        await safeGetPaginated(
          `${apiBase}/pages?per_page=100&include[]=body&sort=title&order=asc`,
          "full pages inventory"
        );

      for (const page of pages) {
        inventories.pages.push(
          storePage(
            page,
            settings
          )
        );
      }
    }

    /**********************************************************
     * FULL ASSIGNMENTS INVENTORY
     **********************************************************/

    if (
      settings.fullInventories.assignments
    ) {
      log(
        "Getting full assignments inventory..."
      );

      const assignments =
        await safeGetPaginated(
          `${apiBase}/assignments?per_page=100&include[]=all_dates&include[]=overrides&include[]=can_edit`,
          "full assignments inventory"
        );

      for (
        const assignment of assignments
      ) {
        inventories.assignments.push(
          storeAssignment(
            assignment,
            settings
          )
        );
      }
    }

    /**********************************************************
     * FULL DISCUSSIONS INVENTORY
     **********************************************************/

    if (
      settings.fullInventories.discussions
    ) {
      log(
        "Getting full discussions inventory..."
      );

      const discussions =
        await safeGetPaginated(
          `${apiBase}/discussion_topics?per_page=100`,
          "full discussions inventory"
        );

      for (
        const discussion of discussions
      ) {
        inventories.discussions.push(
          storeDiscussion(
            discussion,
            settings
          )
        );
      }
    }

    /**********************************************************
     * CLASSIC QUIZZES
     **********************************************************/

    if (
      settings.fullInventories.classicQuizzes
    ) {
      log(
        "Getting full Classic Quizzes inventory..."
      );

      const quizzes =
        await safeGetPaginated(
          `${apiBase}/quizzes?per_page=100`,
          "full Classic Quizzes inventory"
        );

      const refs =
        await mapWithConcurrency(
          quizzes,
          4,
          async quiz => {
            let questions = null;

            if (
              settings.includeClassicQuizQuestions
            ) {
              questions =
                await getClassicQuizQuestions(
                  quiz.id
                );
            }

            return storeClassicQuiz(
              quiz,
              settings,
              questions
            );
          }
        );

      inventories.classic_quizzes.push(
        ...refs
      );
    }

    /**********************************************************
     * NEW QUIZZES
     **********************************************************/

    if (
      settings.fullInventories.newQuizzes
    ) {
      log(
        "Getting full New Quizzes inventory..."
      );

      const newQuizzes =
        await safeGetPaginated(
          `${newQuizApiBase}/quizzes?per_page=100`,
          "full New Quizzes inventory"
        );

      const refs =
        await mapWithConcurrency(
          newQuizzes,
          4,
          async quiz => {
            let items = null;

            if (
              settings.includeNewQuizItems
            ) {
              const quizKey =
                quiz.assignment_id ||
                quiz.id;

              items =
                await safeGetPaginated(
                  `${newQuizApiBase}/quizzes/${encodeURIComponent(
                    quizKey
                  )}/items?per_page=100`,
                  `New Quiz items for ${
                    quiz.title || quizKey
                  }`
                );
            }

            return storeNewQuiz(
              quiz,
              settings,
              items
            );
          }
        );

      inventories.new_quizzes.push(
        ...refs
      );
    }

    /**********************************************************
     * ASSIGNMENT GROUPS
     **********************************************************/

    if (
      settings.courseWide.assignmentGroups
    ) {
      log(
        "Getting assignment groups..."
      );

      const groups =
        await safeGetPaginated(
          `${apiBase}/assignment_groups?per_page=100`,
          "assignment groups"
        );

      courseStructure.assignment_groups =
        groups.map(
          normalizeAssignmentGroup
        );
    }

    /**********************************************************
     * RUBRICS
     **********************************************************/

    if (
      settings.courseWide.rubrics
    ) {
      log(
        "Getting course rubrics..."
      );

      const rubrics =
        await safeGetPaginated(
          `${apiBase}/rubrics?per_page=100`,
          "course rubrics"
        );

      courseStructure.rubrics =
        rubrics.map(rubric =>
          normalizeRubric(
            rubric,
            settings
          )
        );
    }

    /**********************************************************
     * OUTCOMES
     **********************************************************/

    if (
      settings.courseWide.outcomes
    ) {
      log(
        "Getting course outcome links..."
      );

      const outcomes =
        await safeGetPaginated(
          `${apiBase}/outcome_group_links?per_page=100&outcome_style=full&outcome_group_style=full`,
          "course outcome links"
        );

      courseStructure.outcomes =
        outcomes.map(outcomeLink =>
          normalizeOutcomeLink(
            outcomeLink,
            settings
          )
        );
    }

    /**********************************************************
     * NAVIGATION TABS
     **********************************************************/

    if (
      settings.courseWide.tabs
    ) {
      log(
        "Getting course navigation tabs..."
      );

      const tabs =
        await safeGetPaginated(
          `${apiBase}/tabs?per_page=100`,
          "course navigation tabs"
        );

      courseStructure.tabs =
        tabs.map(
          normalizeTab
        );
    }

    /**********************************************************
     * REFERENCED FILES
     **********************************************************/

    let referencedFiles = [];

    if (
      settings.courseWide.referencedFiles
    ) {
      log(
        `Resolving metadata for ${fileReferences.size} referenced Canvas files...`
      );

      referencedFiles =
        await buildReferencedFileManifest(
          settings
        );
    }

    /**********************************************************
     * ALL CANVAS FILES
     **********************************************************/

    let allCourseFiles = [];

    if (
      settings.courseWide.allFiles
    ) {
      log(
        "Getting full Canvas Files inventory - metadata only..."
      );

      const files =
        await safeGetPaginated(
          `${apiBase}/files?per_page=100`,
          "full Canvas Files inventory"
        );

      allCourseFiles =
        files.map(file =>
          normalizeFile(
            file,
            settings
          )
        );

      for (const file of files) {
        storeFile(
          file,
          settings
        );
      }
    }

    /**********************************************************
     * FINAL JSON OBJECT
     **********************************************************/

    const exportData = {
      export_version:
        EXPORT_VERSION,

      generated_at:
        new Date().toISOString(),

      generated_from: {
        canvas_host:
          location.host,

        course_id:
          courseId,

        course_url:
          `${location.origin}/courses/${courseId}`,

        current_page_url:
          location.href
      },

      export_settings:
        settingsForJson(settings),

      notes: [
        "This is a read-only export created from Canvas API GET requests.",
        "Module order and module item order are preserved.",
        "Canvas-native content is stored once in content_store and referenced from modules by content_ref.",
        "Classic Quiz questions are included when enabled and permitted by Canvas.",
        "New Quiz metadata and items are included when enabled and permitted by Canvas.",
        "Referenced Canvas files are represented by metadata and reference locations only. Binary file contents are not downloaded.",
        "The full Canvas Files inventory, when enabled, contains metadata only.",
        "External URLs and External Tools include launch metadata only. External-site content is not retrieved.",
        "Some Canvas APIs depend on permissions and feature availability. Any failures are recorded in errors and the export continues where possible."
      ],

      stats: {
        selected_modules:
          selectedModules.length,

        exported_modules:
          exportedModules.length,

        module_items_total_seen:
          moduleItemCounts.total_seen,

        module_items_included:
          moduleItemCounts.included,

        module_items_skipped_by_type:
          moduleItemCounts.skipped_by_type,

        module_items_skipped_unpublished:
          moduleItemCounts.skipped_unpublished,

        content_store_pages:
          contentStore.pages.size,

        content_store_assignments:
          contentStore.assignments.size,

        content_store_discussions:
          contentStore.discussions.size,

        content_store_classic_quizzes:
          contentStore.classic_quizzes.size,

        content_store_new_quizzes:
          contentStore.new_quizzes.size,

        referenced_canvas_files:
          referencedFiles.length,

        full_canvas_files_inventory:
          allCourseFiles.length,

        api_requests:
          appState.requestCount,

        errors:
          appState.errors.length,

        warnings:
          appState.warnings.length
      },

      course:
        normalizedCourse,

      modules:
        exportedModules,

      ordered_content:
        orderedContent,

      content_store:
        contentStoreForJson(),

      inventories,

      course_structure:
        courseStructure,

      files: {
        referenced_manifest:
          referencedFiles,

        all_course_files:
          allCourseFiles
      },

      warnings:
        appState.warnings,

      errors:
        appState.errors
    };

    log(
      "Creating JSON download..."
    );

    const fileName =
      makeFileName(
        appState.course,
        settings
      );

    downloadJson(
      exportData,
      fileName
    );

    log(
      `Done. Downloaded ${fileName}. ` +
      `Included ${moduleItemCounts.included} module items ` +
      `and found ${referencedFiles.length} referenced Canvas files.`
    );

    console.log(
      "Canvas Course AI JSON export:",
      exportData
    );
  }

  /************************************************************
   * MODULE ITEM ENRICHMENT
   ************************************************************/

  async function enrichModuleItem(
    item,
    settings,
    module
  ) {
    const exported = {
      id:
        item.id,

      module_id:
        item.module_id,

      position:
        item.position,

      title:
        item.title,

      type:
        item.type,

      indent:
        item.indent,

      published:
        item.published,

      content_id:
        item.content_id || null,

      page_url:
        item.page_url || null,

      html_url:
        item.html_url || null,

      external_url:
        item.external_url || null,

      new_tab:
        item.new_tab || null,

      completion_requirement:
        item.completion_requirement || null,

      content_details:
        item.content_details || null,

      content_ref:
        null,

      inline_content:
        null,

      fetch_status:
        "metadata_only",

      fetch_error:
        null
    };

    if (settings.includeRaw) {
      exported.raw_module_item =
        item;
    }

    const moduleReference = {
      source_type:
        "module_item",

      module_id:
        module.id,

      module_name:
        module.name,

      module_item_id:
        item.id,

      module_item_title:
        item.title,

      module_item_type:
        item.type
    };

    try {
      /******************************************************
       * PAGE
       ******************************************************/

      if (item.type === "Page") {
        const pageUrl =
          item.page_url ||
          extractPageUrl(item);

        const page =
          pageUrl
            ? await getPage(pageUrl)
            : await getFromModuleItemUrl(item);

        exported.content_ref =
          page
            ? storePage(
                page,
                settings
              )
            : null;

        exported.fetch_status =
          page
            ? "ok"
            : "not_found";

        return exported;
      }

      /******************************************************
       * ASSIGNMENT
       ******************************************************/

      if (
        item.type === "Assignment"
      ) {
        const assignment =
          item.content_id
            ? await getAssignment(
                item.content_id
              )
            : await getFromModuleItemUrl(
                item
              );

        exported.content_ref =
          assignment
            ? storeAssignment(
                assignment,
                settings
              )
            : null;

        exported.fetch_status =
          assignment
            ? "ok"
            : "not_found";

        return exported;
      }

      /******************************************************
       * DISCUSSION
       ******************************************************/

      if (
        item.type === "Discussion"
      ) {
        const discussion =
          item.content_id
            ? await getDiscussion(
                item.content_id
              )
            : await getFromModuleItemUrl(
                item
              );

        exported.content_ref =
          discussion
            ? storeDiscussion(
                discussion,
                settings
              )
            : null;

        exported.fetch_status =
          discussion
            ? "ok"
            : "not_found";

        return exported;
      }

      /******************************************************
       * CLASSIC QUIZ
       ******************************************************/

      if (item.type === "Quiz") {
        const quiz =
          item.content_id
            ? await getClassicQuiz(
                item.content_id
              )
            : await getFromModuleItemUrl(
                item
              );

        let questions = null;

        if (
          quiz &&
          settings.includeClassicQuizQuestions
        ) {
          questions =
            await getClassicQuizQuestions(
              quiz.id
            );
        }

        exported.content_ref =
          quiz
            ? storeClassicQuiz(
                quiz,
                settings,
                questions
              )
            : null;

        exported.fetch_status =
          quiz
            ? "ok"
            : "not_found";

        return exported;
      }

      /******************************************************
       * FILE
       ******************************************************/

      if (item.type === "File") {
        if (item.content_id) {
          recordFileReference(
            item.content_id,
            moduleReference
          );
        }

        const file =
          item.content_id
            ? await getFile(
                item.content_id,
                item.url
              )
            : await getFromModuleItemUrl(
                item
              );

        exported.content_ref =
          file
            ? storeFile(
                file,
                settings
              )
            : null;

        exported.fetch_status =
          file
            ? "ok"
            : "not_found";

        return exported;
      }

      /******************************************************
       * EXTERNAL URL / EXTERNAL TOOL
       ******************************************************/

      if (
        item.type === "ExternalUrl" ||
        item.type === "ExternalTool"
      ) {
        exported.inline_content = {
          title:
            item.title || null,

          external_url:
            item.external_url || null,

          html_url:
            item.html_url || null,

          url:
            item.url || null,

          new_tab:
            item.new_tab || null,

          content_details:
            item.content_details || null
        };

        exported.fetch_status =
          "metadata_only";

        return exported;
      }

      /******************************************************
       * SUBHEADER
       ******************************************************/

      if (
        item.type === "SubHeader"
      ) {
        exported.inline_content = {
          title:
            item.title || null
        };

        exported.fetch_status =
          "not_applicable";

        return exported;
      }

      /******************************************************
       * OTHER / GENERIC
       ******************************************************/

      const generic =
        await getFromModuleItemUrl(
          item
        );

      if (generic) {
        const key =
          String(
            item.url ||
            item.id
          );

        upsertContent(
          "generic",
          key,
          {
            key,
            title:
              item.title || null,
            data:
              generic
          }
        );

        exported.content_ref =
          makeRef(
            "generic",
            key
          );
      }

      exported.fetch_status =
        generic
          ? "ok_generic"
          : "metadata_only";

      return exported;
    } catch (error) {
      exported.fetch_status =
        "error";

      exported.fetch_error =
        error.message ||
        String(error);

      recordError({
        context:
          "enrich module item",

        item_id:
          item.id,

        item_title:
          item.title,

        item_type:
          item.type,

        message:
          error.message ||
          String(error)
      });

      return exported;
    }
  }

  /************************************************************
   * CONTENT STORE
   ************************************************************/

  function storePage(
    page,
    settings
  ) {
    const key =
      String(
        page.page_id ||
        page.url
      );

    const normalized =
      normalizePage(
        page,
        settings
      );

    upsertContent(
      "pages",
      key,
      normalized
    );

    return makeRef(
      "page",
      key
    );
  }

  function storeAssignment(
    assignment,
    settings
  ) {
    const key =
      String(
        assignment.id
      );

    const normalized =
      normalizeAssignment(
        assignment,
        settings
      );

    upsertContent(
      "assignments",
      key,
      normalized
    );

    return makeRef(
      "assignment",
      key
    );
  }

  function storeDiscussion(
    discussion,
    settings
  ) {
    const key =
      String(
        discussion.id
      );

    const normalized =
      normalizeDiscussion(
        discussion,
        settings
      );

    upsertContent(
      "discussions",
      key,
      normalized
    );

    return makeRef(
      "discussion",
      key
    );
  }

  function storeClassicQuiz(
    quiz,
    settings,
    questions = null
  ) {
    const key =
      String(
        quiz.id
      );

    const normalized =
      normalizeClassicQuiz(
        quiz,
        settings,
        questions
      );

    upsertContent(
      "classic_quizzes",
      key,
      normalized
    );

    return makeRef(
      "classic_quiz",
      key
    );
  }

  function storeNewQuiz(
    quiz,
    settings,
    items = null
  ) {
    const key =
      String(
        quiz.assignment_id ||
        quiz.id
      );

    const normalized =
      normalizeNewQuiz(
        quiz,
        settings,
        items
      );

    upsertContent(
      "new_quizzes",
      key,
      normalized
    );

    return makeRef(
      "new_quiz",
      key
    );
  }

  function storeFile(
    file,
    settings
  ) {
    const key =
      String(
        file.id
      );

    const normalized =
      normalizeFile(
        file,
        settings
      );

    upsertContent(
      "files",
      key,
      normalized
    );

    return makeRef(
      "file",
      key
    );
  }

  function upsertContent(
    bucket,
    key,
    value
  ) {
    const map =
      contentStore[bucket];

    if (!map) {
      throw new Error(
        `Unknown content store bucket: ${bucket}`
      );
    }

    const existing =
      map.get(
        String(key)
      );

    map.set(
      String(key),
      existing
        ? mergeDefined(
            existing,
            value
          )
        : value
    );
  }

  function mergeDefined(
    existing,
    incoming
  ) {
    const merged = {
      ...existing
    };

    for (
      const [key, value] of
      Object.entries(
        incoming || {}
      )
    ) {
      if (
        value !== undefined &&
        value !== null
      ) {
        merged[key] =
          value;
      }
    }

    return merged;
  }

  function makeRef(
    kind,
    key
  ) {
    return {
      kind,
      key:
        String(key)
    };
  }

  function contentStoreForJson() {
    return {
      pages:
        mapValuesSorted(
          contentStore.pages,
          value =>
            value.title ||
            value.url ||
            ""
        ),

      assignments:
        mapValuesSorted(
          contentStore.assignments,
          value =>
            value.name ||
            ""
        ),

      discussions:
        mapValuesSorted(
          contentStore.discussions,
          value =>
            value.title ||
            ""
        ),

      classic_quizzes:
        mapValuesSorted(
          contentStore.classic_quizzes,
          value =>
            value.title ||
            ""
        ),

      new_quizzes:
        mapValuesSorted(
          contentStore.new_quizzes,
          value =>
            value.title ||
            ""
        ),

      files:
        mapValuesSorted(
          contentStore.files,
          value =>
            value.display_name ||
            value.filename ||
            ""
        ),

      generic:
        Array.from(
          contentStore.generic.values()
        )
    };
  }

  function mapValuesSorted(
    map,
    getLabel
  ) {
    return Array.from(
      map.values()
    ).sort(
      (a, b) =>
        String(
          getLabel(a)
        ).localeCompare(
          String(
            getLabel(b)
          )
        )
    );
  }

  /************************************************************
   * NORMALIZERS
   ************************************************************/

  function normalizeCourse(
    course,
    settings
  ) {
    if (!course) {
      return null;
    }

    const syllabus =
      course.syllabus_body ||
      "";

    recordHtmlFileReferences(
      syllabus,
      {
        source_type:
          "syllabus",

        source_title:
          "Course Syllabus"
      }
    );

    const normalized = {
      id:
        course.id,

      name:
        course.name,

      course_code:
        course.course_code,

      workflow_state:
        course.workflow_state,

      start_at:
        course.start_at,

      end_at:
        course.end_at,

      time_zone:
        course.time_zone,

      default_view:
        course.default_view,

      public_syllabus:
        course.public_syllabus,

      public_syllabus_to_auth:
        course.public_syllabus_to_auth,

      storage_quota_used_mb:
        course.storage_quota_used_mb ??
        null,

      apply_assignment_group_weights:
        course.apply_assignment_group_weights ??
        null,

      grading_standard_id:
        course.grading_standard_id ??
        null,

      hide_final_grades:
        course.hide_final_grades ??
        null,

      restrict_enrollments_to_course_dates:
        course.restrict_enrollments_to_course_dates ??
        null,

      post_manually:
        course.post_manually ??
        null
    };

    addRichTextFields(
      normalized,
      "syllabus_body",
      syllabus,
      settings
    );

    if (settings.includeRaw) {
      normalized.raw_course =
        course;
    }

    return normalized;
  }

  function normalizeModule(
    module,
    items,
    settings
  ) {
    const normalized = {
      id:
        module.id,

      name:
        module.name,

      position:
        module.position,

      published:
        module.published,

      workflow_state:
        module.workflow_state,

      unlock_at:
        module.unlock_at,

      require_sequential_progress:
        module.require_sequential_progress,

      requirement_type:
        module.requirement_type,

      prerequisite_module_ids:
        module.prerequisite_module_ids ||
        [],

      items_count_reported_by_canvas:
        module.items_count,

      exported_items_count:
        items.length,

      items
    };

    if (settings.includeRaw) {
      normalized.raw_module =
        module;
    }

    return normalized;
  }

  function normalizePage(
    page,
    settings
  ) {
    const body =
      page.body ||
      "";

    recordHtmlFileReferences(
      body,
      {
        source_type:
          "page",

        source_id:
          page.page_id ||
          null,

        source_key:
          page.url ||
          null,

        source_title:
          page.title ||
          null
      }
    );

    const normalized = {
      page_id:
        page.page_id,

      url:
        page.url,

      title:
        page.title,

      published:
        page.published,

      front_page:
        page.front_page,

      editor:
        page.editor ||
        null,

      created_at:
        page.created_at,

      updated_at:
        page.updated_at,

      html_url:
        page.html_url ||
        null
    };

    addRichTextFields(
      normalized,
      "body",
      body,
      settings
    );

    if (settings.includeRaw) {
      normalized.raw_page =
        page;
    }

    return normalized;
  }

  function normalizeAssignment(
    assignment,
    settings
  ) {
    const description =
      assignment.description ||
      "";

    const source = {
      source_type:
        "assignment",

      source_id:
        assignment.id,

      source_title:
        assignment.name ||
        null
    };

    recordHtmlFileReferences(
      description,
      source
    );

    recordAttachmentFileReferences(
      assignment.attachments,
      source
    );

    const normalized = {
      id:
        assignment.id,

      name:
        assignment.name,

      published:
        assignment.published,

      workflow_state:
        assignment.workflow_state,

      points_possible:
        assignment.points_possible,

      grading_type:
        assignment.grading_type,

      assignment_group_id:
        assignment.assignment_group_id,

      due_at:
        assignment.due_at,

      unlock_at:
        assignment.unlock_at,

      lock_at:
        assignment.lock_at,

      all_dates:
        assignment.all_dates ||
        null,

      overrides:
        assignment.overrides ||
        null,

      submission_types:
        assignment.submission_types ||
        [],

      allowed_extensions:
        assignment.allowed_extensions ||
        [],

      peer_reviews:
        assignment.peer_reviews ||
        false,

      automatic_peer_reviews:
        assignment.automatic_peer_reviews ||
        false,

      intra_group_peer_reviews:
        assignment.intra_group_peer_reviews ||
        false,

      group_category_id:
        assignment.group_category_id ||
        null,

      grade_group_students_individually:
        assignment.grade_group_students_individually ||
        false,

      anonymous_grading:
        assignment.anonymous_grading ||
        false,

      moderated_grading:
        assignment.moderated_grading ||
        false,

      omit_from_final_grade:
        assignment.omit_from_final_grade ||
        false,

      only_visible_to_overrides:
        assignment.only_visible_to_overrides ||
        false,

      html_url:
        assignment.html_url ||
        null,

      external_tool_tag_attributes:
        assignment.external_tool_tag_attributes ||
        null,

      rubric:
        assignment.rubric ||
        null,

      rubric_settings:
        assignment.rubric_settings ||
        null
    };

    addRichTextFields(
      normalized,
      "description",
      description,
      settings
    );

    if (settings.includeRaw) {
      normalized.raw_assignment =
        assignment;
    }

    return normalized;
  }

  function normalizeDiscussion(
    discussion,
    settings
  ) {
    const message =
      discussion.message ||
      discussion.description ||
      "";

    const source = {
      source_type:
        "discussion",

      source_id:
        discussion.id,

      source_title:
        discussion.title ||
        null
    };

    recordHtmlFileReferences(
      message,
      source
    );

    recordAttachmentFileReferences(
      discussion.attachments,
      source
    );

    const normalized = {
      id:
        discussion.id,

      title:
        discussion.title,

      published:
        discussion.published,

      workflow_state:
        discussion.workflow_state,

      posted_at:
        discussion.posted_at,

      delayed_post_at:
        discussion.delayed_post_at,

      lock_at:
        discussion.lock_at,

      assignment_id:
        discussion.assignment_id ||
        null,

      discussion_type:
        discussion.discussion_type ||
        null,

      pinned:
        discussion.pinned ||
        false,

      locked:
        discussion.locked ||
        false,

      require_initial_post:
        discussion.require_initial_post ||
        false,

      group_category_id:
        discussion.group_category_id ||
        null,

      html_url:
        discussion.html_url ||
        null
    };

    addRichTextFields(
      normalized,
      "message",
      message,
      settings
    );

    if (settings.includeRaw) {
      normalized.raw_discussion =
        discussion;
    }

    return normalized;
  }

  function normalizeClassicQuiz(
    quiz,
    settings,
    questions = null
  ) {
    const description =
      quiz.description ||
      "";

    recordHtmlFileReferences(
      description,
      {
        source_type:
          "classic_quiz",

        source_id:
          quiz.id,

        source_title:
          quiz.title ||
          null
      }
    );

    const normalized = {
      id:
        quiz.id,

      title:
        quiz.title,

      quiz_type:
        quiz.quiz_type,

      published:
        quiz.published,

      workflow_state:
        quiz.workflow_state,

      assignment_id:
        quiz.assignment_id ||
        null,

      assignment_group_id:
        quiz.assignment_group_id ||
        null,

      points_possible:
        quiz.points_possible,

      due_at:
        quiz.due_at,

      unlock_at:
        quiz.unlock_at,

      lock_at:
        quiz.lock_at,

      time_limit:
        quiz.time_limit,

      allowed_attempts:
        quiz.allowed_attempts,

      scoring_policy:
        quiz.scoring_policy,

      shuffle_answers:
        quiz.shuffle_answers,

      one_question_at_a_time:
        quiz.one_question_at_a_time,

      cant_go_back:
        quiz.cant_go_back,

      hide_results:
        quiz.hide_results,

      show_correct_answers:
        quiz.show_correct_answers,

      show_correct_answers_at:
        quiz.show_correct_answers_at ||
        null,

      hide_correct_answers_at:
        quiz.hide_correct_answers_at ||
        null,

      html_url:
        quiz.html_url ||
        null
    };

    addRichTextFields(
      normalized,
      "description",
      description,
      settings
    );

    if (
      Array.isArray(questions)
    ) {
      normalized.questions =
        questions.map(
          question =>
            normalizeClassicQuizQuestion(
              question,
              settings,
              quiz
            )
        );
    }

    if (settings.includeRaw) {
      normalized.raw_quiz =
        quiz;
    }

    return normalized;
  }

  function normalizeClassicQuizQuestion(
    question,
    settings,
    quiz
  ) {
    const source = {
      source_type:
        "classic_quiz_question",

      source_id:
        question.id,

      source_title:
        question.question_name ||
        null,

      parent_quiz_id:
        quiz.id,

      parent_quiz_title:
        quiz.title ||
        null
    };

    const questionText =
      question.question_text ||
      "";

    recordHtmlFileReferences(
      questionText,
      source
    );

    for (
      const field of [
        "correct_comments",
        "incorrect_comments",
        "neutral_comments",
        "text_after_answers"
      ]
    ) {
      if (question[field]) {
        recordHtmlFileReferences(
          question[field],
          source
        );
      }
    }

    if (
      Array.isArray(
        question.answers
      )
    ) {
      for (
        const answer of
        question.answers
      ) {
        for (
          const value of
          Object.values(
            answer || {}
          )
        ) {
          if (
            typeof value ===
              "string" &&
            value.includes("<")
          ) {
            recordHtmlFileReferences(
              value,
              source
            );
          }
        }
      }
    }

    const normalized = {
      id:
        question.id,

      quiz_id:
        question.quiz_id,

      quiz_group_id:
        question.quiz_group_id ||
        null,

      position:
        question.position,

      question_name:
        question.question_name,

      question_type:
        question.question_type,

      points_possible:
        question.points_possible,

      answers:
        question.answers ||
        [],

      variables:
        question.variables ||
        null,

      formulas:
        question.formulas ||
        null,

      answer_tolerance:
        question.answer_tolerance ??
        null,

      formula_decimal_places:
        question.formula_decimal_places ??
        null,

      matches:
        question.matches ||
        null
    };

    addRichTextFields(
      normalized,
      "question_text",
      questionText,
      settings
    );

    addOptionalRichTextField(
      normalized,
      "correct_comments",
      question.correct_comments,
      settings
    );

    addOptionalRichTextField(
      normalized,
      "incorrect_comments",
      question.incorrect_comments,
      settings
    );

    addOptionalRichTextField(
      normalized,
      "neutral_comments",
      question.neutral_comments,
      settings
    );

    addOptionalRichTextField(
      normalized,
      "text_after_answers",
      question.text_after_answers,
      settings
    );

    if (settings.includeRaw) {
      normalized.raw_question =
        question;
    }

    return normalized;
  }

  function normalizeNewQuiz(
    quiz,
    settings,
    items = null
  ) {
    const instructions =
      quiz.instructions ||
      "";

    recordHtmlFileReferences(
      instructions,
      {
        source_type:
          "new_quiz",

        source_id:
          quiz.assignment_id ||
          quiz.id,

        source_title:
          quiz.title ||
          null
      }
    );

    const normalized = {
      id:
        quiz.id ||
        null,

      assignment_id:
        quiz.assignment_id ||
        null,

      title:
        quiz.title,

      assignment_group_id:
        quiz.assignment_group_id ||
        null,

      points_possible:
        quiz.points_possible,

      due_at:
        quiz.due_at,

      lock_at:
        quiz.lock_at,

      unlock_at:
        quiz.unlock_at,

      published:
        quiz.published,

      grading_type:
        quiz.grading_type,

      quiz_settings:
        quiz.quiz_settings ||
        null
    };

    addRichTextFields(
      normalized,
      "instructions",
      instructions,
      settings
    );

    if (
      Array.isArray(items)
    ) {
      normalized.items =
        items.map(
          item =>
            normalizeNewQuizItem(
              item,
              settings,
              quiz
            )
        );
    }

    if (settings.includeRaw) {
      normalized.raw_new_quiz =
        quiz;
    }

    return normalized;
  }

  function normalizeNewQuizItem(
    item,
    settings,
    quiz
  ) {
    const source = {
      source_type:
        "new_quiz_item",

      source_id:
        item.id,

      source_title:
        item.entry?.title ||
        item.entry?.name ||
        null,

      parent_quiz_id:
        quiz.assignment_id ||
        quiz.id,

      parent_quiz_title:
        quiz.title ||
        null
    };

    const entry =
      item.entry ||
      null;

    const normalizedEntry =
      normalizeNewQuizEntry(
        entry,
        settings,
        source
      );

    const normalized = {
      id:
        item.id,

      position:
        item.position,

      points_possible:
        item.points_possible,

      entry_type:
        item.entry_type,

      entry_editable:
        item.entry_editable,

      stimulus_quiz_entry_id:
        item.stimulus_quiz_entry_id ||
        null,

      status:
        item.status ||
        null,

      properties:
        item.properties ||
        null,

      entry:
        normalizedEntry
    };

    if (settings.includeRaw) {
      normalized.raw_item =
        item;
    }

    return normalized;
  }

  function normalizeNewQuizEntry(
    entry,
    settings,
    source
  ) {
    if (
      !entry ||
      typeof entry !== "object"
    ) {
      return entry;
    }

    const normalized = {
      ...entry
    };

    const richFields = [
      "item_body",
      "body",
      "instructions"
    ];

    for (
      const field of richFields
    ) {
      if (
        typeof entry[field] ===
        "string"
      ) {
        recordHtmlFileReferences(
          entry[field],
          source
        );

        delete normalized[field];

        addRichTextFields(
          normalized,
          field,
          entry[field],
          settings
        );
      }
    }

    if (
      entry.feedback &&
      typeof entry.feedback ===
        "object"
    ) {
      normalized.feedback = {};

      for (
        const [key, value] of
        Object.entries(
          entry.feedback
        )
      ) {
        if (
          typeof value ===
          "string"
        ) {
          recordHtmlFileReferences(
            value,
            source
          );

          addRichTextFields(
            normalized.feedback,
            key,
            value,
            settings
          );
        } else {
          normalized.feedback[key] =
            value;
        }
      }
    }

    if (
      entry.answer_feedback &&
      typeof entry.answer_feedback ===
        "object"
    ) {
      normalized.answer_feedback = {};

      for (
        const [key, value] of
        Object.entries(
          entry.answer_feedback
        )
      ) {
        if (
          typeof value ===
          "string"
        ) {
          recordHtmlFileReferences(
            value,
            source
          );

          normalized.answer_feedback[key] = {
            html:
              settings.includeHtml
                ? value
                : undefined,

            ai_text:
              settings.includeAiText
                ? htmlToAiText(
                    value
                  )
                : undefined
          };
        } else {
          normalized.answer_feedback[key] =
            value;
        }
      }
    }

    return normalized;
  }

  function normalizeAssignmentGroup(
    group
  ) {
    return {
      id:
        group.id,

      name:
        group.name,

      position:
        group.position,

      group_weight:
        group.group_weight,

      rules:
        group.rules ||
        null,

      sis_source_id:
        group.sis_source_id ||
        null
    };
  }

  function normalizeRubric(
    rubric,
    settings
  ) {
    const normalized = {
      id:
        rubric.id,

      title:
        rubric.title,

      context_id:
        rubric.context_id ||
        null,

      context_type:
        rubric.context_type ||
        null,

      points_possible:
        rubric.points_possible,

      reusable:
        rubric.reusable,

      read_only:
        rubric.read_only,

      free_form_criterion_comments:
        rubric.free_form_criterion_comments,

      criteria:
        rubric.criteria ||
        null
    };

    if (settings.includeRaw) {
      normalized.raw_rubric =
        rubric;
    }

    return normalized;
  }

  function normalizeOutcomeLink(
    link,
    settings
  ) {
    const normalized = {
      url:
        link.url ||
        null,

      context_id:
        link.context_id ||
        null,

      context_type:
        link.context_type ||
        null,

      assessed:
        link.assessed ??
        null,

      can_unlink:
        link.can_unlink ??
        null,

      outcome_group:
        link.outcome_group ||
        null,

      outcome:
        link.outcome
          ? {
              ...link.outcome
            }
          : null
    };

    if (
      normalized.outcome?.description
    ) {
      addOptionalRichTextField(
        normalized.outcome,
        "description",
        normalized.outcome.description,
        settings
      );
    }

    if (settings.includeRaw) {
      normalized.raw_outcome_link =
        link;
    }

    return normalized;
  }

  function normalizeTab(
    tab
  ) {
    return {
      id:
        tab.id,

      label:
        tab.label,

      type:
        tab.type,

      html_url:
        tab.html_url ||
        null,

      position:
        tab.position,

      hidden:
        tab.hidden ||
        false,

      visibility:
        tab.visibility ||
        null
    };
  }

  /************************************************************
   * FILE NORMALIZER
   *
   * Canvas sometimes returns MIME type as "content-type"
   * instead of "content_type".
   ************************************************************/

  function normalizeFile(
    file,
    settings
  ) {
    const normalized = {
      id:
        file.id,

      uuid:
        file.uuid,

      display_name:
        file.display_name,

      filename:
        file.filename,

      content_type:
        file.content_type ||
        file["content-type"] ||
        null,

      size:
        file.size,

      folder_id:
        file.folder_id,

      url:
        file.url ||
        null,

      preview_url:
        file.preview_url ||
        null,

      created_at:
        file.created_at,

      updated_at:
        file.updated_at,

      unlock_at:
        file.unlock_at,

      locked:
        file.locked,

      hidden:
        file.hidden,

      hidden_for_user:
        file.hidden_for_user,

      locked_for_user:
        file.locked_for_user,

      thumbnail_url:
        file.thumbnail_url ||
        null
    };

    if (settings.includeRaw) {
      normalized.raw_file =
        file;
    }

    return normalized;
  }

  /************************************************************
   * RICH TEXT
   ************************************************************/

  function addRichTextFields(
    target,
    baseName,
    html,
    settings
  ) {
    const value =
      html ||
      "";

    if (settings.includeHtml) {
      target[
        `${baseName}_html`
      ] = value;
    }

    if (
      settings.includeAiText
    ) {
      target[
        `${baseName}_ai_text`
      ] =
        htmlToAiText(value);
    }
  }

  function addOptionalRichTextField(
    target,
    baseName,
    html,
    settings
  ) {
    if (
      html == null ||
      html === ""
    ) {
      return;
    }

    addRichTextFields(
      target,
      baseName,
      html,
      settings
    );
  }

  /************************************************************
   * FILE REFERENCE DETECTION
   ************************************************************/

  function recordAttachmentFileReferences(
    attachments,
    source
  ) {
    if (
      !Array.isArray(
        attachments
      )
    ) {
      return;
    }

    for (
      const attachment of
      attachments
    ) {
      if (
        !attachment ||
        attachment.id == null
      ) {
        continue;
      }

      recordFileReference(
        attachment.id,
        {
          ...source,

          reference_type:
            "attachment",

          reference_label:
            attachment.display_name ||
            attachment.filename ||
            null
        }
      );

      caches.fileById.set(
        String(
          attachment.id
        ),
        attachment
      );
    }
  }

  function recordHtmlFileReferences(
    html,
    source
  ) {
    if (!html) {
      return;
    }

    let doc;

    try {
      doc =
        new DOMParser()
          .parseFromString(
            String(html),
            "text/html"
          );
    } catch {
      return;
    }

    const attrs = [
      "href",
      "src",
      "data-api-endpoint",
      "data-api-url",
      "data-canvas-previewable-url"
    ];

    for (
      const element of
      doc.querySelectorAll("*")
    ) {
      for (
        const attr of attrs
      ) {
        if (
          !element.hasAttribute(attr)
        ) {
          continue;
        }

        const value =
          element.getAttribute(attr);

        const ids =
          extractCanvasFileIdsFromString(
            value
          );

        for (const id of ids) {
          recordFileReference(
            id,
            {
              ...source,

              reference_type:
                element.tagName.toLowerCase(),

              reference_attribute:
                attr,

              reference_url:
                value,

              reference_label:
                getElementReferenceLabel(
                  element
                )
            }
          );
        }
      }
    }

    for (
      const id of
      extractCanvasFileIdsFromString(
        String(html)
      )
    ) {
      recordFileReference(
        id,
        {
          ...source,

          reference_type:
            "html_scan"
        }
      );
    }
  }

  function extractCanvasFileIdsFromString(
    value
  ) {
    if (!value) {
      return [];
    }

    const text =
      String(value);

    const ids =
      new Set();

    const patterns = [
      /\/api\/v1\/(?:courses\/\d+\/)?files\/(\d+)/gi,
      /\/courses\/\d+\/files\/(\d+)/gi,
      /\/files\/(\d+)(?:\/download|\/preview|\b|\?)/gi
    ];

    for (
      const pattern of patterns
    ) {
      let match;

      while (
        (
          match =
            pattern.exec(text)
        ) !== null
      ) {
        ids.add(
          match[1]
        );
      }
    }

    return Array.from(ids);
  }

  function getElementReferenceLabel(
    element
  ) {
    if (!element) {
      return null;
    }

    const alt =
      element.getAttribute?.(
        "alt"
      );

    if (alt) {
      return alt
        .trim()
        .slice(
          0,
          250
        );
    }

    const title =
      element.getAttribute?.(
        "title"
      );

    if (title) {
      return title
        .trim()
        .slice(
          0,
          250
        );
    }

    const text =
      element.textContent
        ?.replace(
          /\s+/g,
          " "
        )
        .trim();

    return text
      ? text.slice(
          0,
          250
        )
      : null;
  }

  function recordFileReference(
    fileId,
    reference
  ) {
    const key =
      String(fileId);

    if (
      !fileReferences.has(key)
    ) {
      fileReferences.set(
        key,
        {
          file_id:
            key,

          references:
            [],

          _dedupe:
            new Set()
        }
      );
    }

    const entry =
      fileReferences.get(key);

    const cleaned =
      removeUndefined(
        reference ||
        {}
      );

    const dedupeKey =
      JSON.stringify(cleaned);

    if (
      !entry._dedupe.has(
        dedupeKey
      )
    ) {
      entry._dedupe.add(
        dedupeKey
      );

      entry.references.push(
        cleaned
      );
    }
  }

  async function buildReferencedFileManifest(
    settings
  ) {
    const entries =
      Array.from(
        fileReferences.values()
      );

    return await mapWithConcurrency(
      entries,
      4,
      async entry => {
        const file =
          await getFile(
            entry.file_id,
            null
          );

        if (file) {
          storeFile(
            file,
            settings
          );
        }

        return {
          file_id:
            entry.file_id,

          file:
            file
              ? normalizeFile(
                  file,
                  settings
                )
              : null,

          references:
            entry.references
        };
      }
    );
  }

  function removeUndefined(
    obj
  ) {
    return Object.fromEntries(
      Object.entries(obj).filter(
        ([, value]) =>
          value !== undefined
      )
    );
  }

  /************************************************************
   * CANVAS CONTENT FETCHERS
   ************************************************************/

  async function getPage(
    pageUrl
  ) {
    const key =
      String(pageUrl);

    if (
      caches.pageByUrl.has(key)
    ) {
      return caches.pageByUrl.get(
        key
      );
    }

    const page =
      await safeGet(
        `${apiBase}/pages/${encodeURIComponent(
          key
        )}`,
        `page ${key}`
      );

    if (page) {
      caches.pageByUrl.set(
        key,
        page
      );
    }

    return page;
  }

  async function getAssignment(
    assignmentId
  ) {
    const key =
      String(assignmentId);

    if (
      caches.assignmentById.has(
        key
      )
    ) {
      return caches.assignmentById.get(
        key
      );
    }

    const assignment =
      await safeGet(
        `${apiBase}/assignments/${encodeURIComponent(
          key
        )}?include[]=all_dates&include[]=overrides&include[]=can_edit`,
        `assignment ${key}`
      );

    if (assignment) {
      caches.assignmentById.set(
        key,
        assignment
      );
    }

    return assignment;
  }

  async function getDiscussion(
    discussionId
  ) {
    const key =
      String(discussionId);

    if (
      caches.discussionById.has(
        key
      )
    ) {
      return caches.discussionById.get(
        key
      );
    }

    const discussion =
      await safeGet(
        `${apiBase}/discussion_topics/${encodeURIComponent(
          key
        )}`,
        `discussion ${key}`
      );

    if (discussion) {
      caches.discussionById.set(
        key,
        discussion
      );
    }

    return discussion;
  }

  async function getClassicQuiz(
    quizId
  ) {
    const key =
      String(quizId);

    if (
      caches.classicQuizById.has(
        key
      )
    ) {
      return caches.classicQuizById.get(
        key
      );
    }

    const quiz =
      await safeGet(
        `${apiBase}/quizzes/${encodeURIComponent(
          key
        )}`,
        `Classic Quiz ${key}`
      );

    if (quiz) {
      caches.classicQuizById.set(
        key,
        quiz
      );
    }

    return quiz;
  }

  async function getClassicQuizQuestions(
    quizId
  ) {
    const key =
      String(quizId);

    if (
      caches.classicQuizQuestionsById.has(
        key
      )
    ) {
      return caches.classicQuizQuestionsById.get(
        key
      );
    }

    const questions =
      await safeGetPaginated(
        `${apiBase}/quizzes/${encodeURIComponent(
          key
        )}/questions?per_page=100`,
        `Classic Quiz questions for ${key}`
      );

    caches.classicQuizQuestionsById.set(
      key,
      questions
    );

    return questions;
  }

  async function getFile(
    fileId,
    itemUrl
  ) {
    const key =
      String(fileId);

    if (
      caches.fileById.has(key)
    ) {
      return caches.fileById.get(
        key
      );
    }

    let file = null;

    if (
      itemUrl &&
      String(itemUrl).includes(
        "/api/v1/"
      )
    ) {
      file =
        await safeGet(
          itemUrl,
          `file ${key}`
        );
    }

    if (!file) {
      file =
        await safeGet(
          `${location.origin}/api/v1/files/${encodeURIComponent(
            key
          )}`,
          `file ${key}`
        );
    }

    if (file) {
      caches.fileById.set(
        key,
        file
      );
    }

    return file;
  }

  async function getFromModuleItemUrl(
    item
  ) {
    if (
      !item?.url ||
      !String(item.url).includes(
        "/api/"
      )
    ) {
      return null;
    }

    const url =
      makeAbsoluteUrl(
        item.url
      );

    if (
      caches.genericByUrl.has(
        url
      )
    ) {
      return caches.genericByUrl.get(
        url
      );
    }

    const data =
      await safeGet(
        url,
        `module item URL ${
          item.title || item.id
        }`
      );

    if (data) {
      caches.genericByUrl.set(
        url,
        data
      );
    }

    return data;
  }

  /************************************************************
   * CANVAS API
   ************************************************************/

  async function canvasApiRequest(
    url
  ) {
    appState.requestCount++;

    const response =
      await fetch(
        makeAbsoluteUrl(url),
        {
          method:
            "GET",

          credentials:
            "same-origin",

          headers: {
            Accept:
              "application/json+canvas-string-ids"
          }
        }
      );

    const text =
      await response.text();

    let data = null;

    if (text) {
      try {
        data =
          JSON.parse(text);
      } catch {
        throw new Error(
          `Non-JSON response from ${url}: ${text.slice(
            0,
            400
          )}`
        );
      }
    }

    if (!response.ok) {
      throw new Error(
        `Canvas API error ${response.status} ${response.statusText} from ${url}: ${text.slice(
          0,
          700
        )}`
      );
    }

    return {
      data,
      headers:
        response.headers
    };
  }

  async function getPaginated(
    url
  ) {
    let nextUrl =
      makeAbsoluteUrl(url);

    const results = [];

    while (nextUrl) {
      const {
        data,
        headers
      } =
        await canvasApiRequest(
          nextUrl
        );

      if (
        !Array.isArray(data)
      ) {
        throw new Error(
          `Expected an array from ${nextUrl}`
        );
      }

      results.push(
        ...data
      );

      nextUrl =
        getNextLink(
          headers.get("Link")
        );
    }

    return results;
  }

  async function safeGetPaginated(
    url,
    context
  ) {
    try {
      return await getPaginated(
        url
      );
    } catch (error) {
      recordError({
        context,
        url,
        message:
          error.message ||
          String(error)
      });

      log(
        `Could not get ${context}. Continuing.`
      );

      console.warn(
        `Canvas export warning for ${context}:`,
        error
      );

      return [];
    }
  }

  async function safeGet(
    url,
    context
  ) {
    try {
      const {
        data
      } =
        await canvasApiRequest(
          url
        );

      return data;
    } catch (error) {
      recordError({
        context,
        url,
        message:
          error.message ||
          String(error)
      });

      log(
        `Could not get ${context}. Continuing.`
      );

      console.warn(
        `Canvas export warning for ${context}:`,
        error
      );

      return null;
    }
  }

  async function mapWithConcurrency(
    items,
    limit,
    mapper
  ) {
    const source =
      Array.from(
        items ||
        []
      );

    const results =
      new Array(
        source.length
      );

    let nextIndex = 0;

    async function worker() {
      while (true) {
        const index =
          nextIndex++;

        if (
          index >=
          source.length
        ) {
          return;
        }

        results[index] =
          await mapper(
            source[index],
            index
          );
      }
    }

    const workerCount =
      Math.min(
        Math.max(
          1,
          limit
        ),
        Math.max(
          1,
          source.length
        )
      );

    const workers =
      Array.from(
        {
          length:
            workerCount
        },
        () =>
          worker()
      );

    await Promise.all(
      workers
    );

    return results;
  }

  function getNextLink(
    linkHeader
  ) {
    if (!linkHeader) {
      return null;
    }

    for (
      const link of
      linkHeader.split(",")
    ) {
      const match =
        link.match(
          /<([^>]+)>;\s*rel="next"/
        );

      if (match) {
        return match[1];
      }
    }

    return null;
  }

  function makeAbsoluteUrl(
    url
  ) {
    if (!url) {
      return url;
    }

    const value =
      String(url);

    if (
      value.startsWith(
        "http://"
      ) ||
      value.startsWith(
        "https://"
      )
    ) {
      return value;
    }

    if (
      value.startsWith("/")
    ) {
      return (
        `${location.origin}${value}`
      );
    }

    return (
      `${location.origin}/${value}`
    );
  }

  /************************************************************
   * SETTINGS
   ************************************************************/

  function readSettings() {
    const selectedModuleIds =
      Array.from(
        document.querySelectorAll(
          ".cje-module-checkbox:checked"
        )
      ).map(
        cb =>
          cb.value
      );

    const contentTypes =
      new Set(
        Array.from(
          document.querySelectorAll(
            ".cje-type-checkbox:checked"
          )
        ).map(
          cb =>
            cb.value
        )
      );

    return {
      preset:
        document.getElementById(
          "cje-preset"
        ).value,

      selectedModuleIds,

      selectedModuleNames:
        selectedModuleIds.map(
          id => {
            const module =
              appState.modules.find(
                m =>
                  String(m.id) ===
                  String(id)
              );

            return module
              ? module.name
              : id;
          }
        ),

      contentTypes,

      includeHtml:
        document.getElementById(
          "cje-include-html"
        ).checked,

      includeAiText:
        document.getElementById(
          "cje-include-ai-text"
        ).checked,

      includeRaw:
        document.getElementById(
          "cje-include-raw"
        ).checked,

      includeUnpublished:
        document.getElementById(
          "cje-include-unpublished"
        ).checked,

      includeClassicQuizQuestions:
        document.getElementById(
          "cje-classic-quiz-questions"
        ).checked,

      includeNewQuizItems:
        document.getElementById(
          "cje-new-quiz-items"
        ).checked,

      fullInventories: {
        pages:
          document.getElementById(
            "cje-inventory-pages"
          ).checked,

        assignments:
          document.getElementById(
            "cje-inventory-assignments"
          ).checked,

        discussions:
          document.getElementById(
            "cje-inventory-discussions"
          ).checked,

        classicQuizzes:
          document.getElementById(
            "cje-inventory-classic-quizzes"
          ).checked,

        newQuizzes:
          document.getElementById(
            "cje-inventory-new-quizzes"
          ).checked
      },

      courseWide: {
        assignmentGroups:
          document.getElementById(
            "cje-inventory-assignment-groups"
          ).checked,

        rubrics:
          document.getElementById(
            "cje-inventory-rubrics"
          ).checked,

        outcomes:
          document.getElementById(
            "cje-inventory-outcomes"
          ).checked,

        tabs:
          document.getElementById(
            "cje-inventory-tabs"
          ).checked,

        referencedFiles:
          document.getElementById(
            "cje-referenced-files"
          ).checked,

        allFiles:
          document.getElementById(
            "cje-all-files"
          ).checked
      }
    };
  }

  function settingsForJson(
    settings
  ) {
    return {
      preset:
        settings.preset,

      selectedModuleIds:
        settings.selectedModuleIds,

      selectedModuleNames:
        settings.selectedModuleNames,

      contentTypes:
        Array.from(
          settings.contentTypes
        ),

      includeHtml:
        settings.includeHtml,

      includeAiText:
        settings.includeAiText,

      includeRaw:
        settings.includeRaw,

      includeUnpublished:
        settings.includeUnpublished,

      includeClassicQuizQuestions:
        settings.includeClassicQuizQuestions,

      includeNewQuizItems:
        settings.includeNewQuizItems,

      fullInventories:
        settings.fullInventories,

      courseWide:
        settings.courseWide
    };
  }

  /************************************************************
   * PRESETS
   ************************************************************/

  function applyPreset(
    presetValue
  ) {
    const typeChecks =
      Array.from(
        document.querySelectorAll(
          ".cje-type-checkbox"
        )
      );

    const moduleChecks =
      Array.from(
        document.querySelectorAll(
          ".cje-module-checkbox"
        )
      );

    /******************************************************
     * AI REVIEW
     ******************************************************/

    if (
      presetValue ===
      "ai_review"
    ) {
      moduleChecks.forEach(
        cb =>
          cb.checked = true
      );

      typeChecks.forEach(
        cb =>
          cb.checked = true
      );

      setCoreInventories(true);
      setCourseWide(true);

      setChecked(
        "cje-all-files",
        true
      );

      setChecked(
        "cje-include-html",
        true
      );

      setChecked(
        "cje-include-ai-text",
        true
      );

      setChecked(
        "cje-include-raw",
        false
      );

      setChecked(
        "cje-include-unpublished",
        true
      );

      setChecked(
        "cje-classic-quiz-questions",
        true
      );

      setChecked(
        "cje-new-quiz-items",
        true
      );

      return;
    }

    /******************************************************
     * SELECTED MODULES, ALL TYPES
     ******************************************************/

    if (
      presetValue ===
      "selected_all"
    ) {
      typeChecks.forEach(
        cb =>
          cb.checked = true
      );

      setCoreInventories(false);
      setCourseWide(false);

      setChecked(
        "cje-include-html",
        true
      );

      setChecked(
        "cje-include-ai-text",
        true
      );

      setChecked(
        "cje-include-raw",
        false
      );

      return;
    }

    /******************************************************
     * SELECTED PAGES
     ******************************************************/

    if (
      presetValue ===
      "selected_pages"
    ) {
      setContentTypes(
        ["Page"]
      );

      setCoreInventories(false);
      setCourseWide(false);

      setChecked(
        "cje-include-raw",
        false
      );

      return;
    }

    /******************************************************
     * SELECTED ASSIGNMENTS
     ******************************************************/

    if (
      presetValue ===
      "selected_assignments"
    ) {
      setContentTypes(
        ["Assignment"]
      );

      setCoreInventories(false);
      setCourseWide(false);

      setChecked(
        "cje-include-raw",
        false
      );

      return;
    }

    /******************************************************
     * SELECTED PAGES + ASSIGNMENTS
     ******************************************************/

    if (
      presetValue ===
      "selected_pages_assignments"
    ) {
      setContentTypes(
        [
          "Page",
          "Assignment"
        ]
      );

      setCoreInventories(false);
      setCourseWide(false);

      setChecked(
        "cje-include-raw",
        false
      );
    }
  }

  function setCoreInventories(
    value
  ) {
    setChecked(
      "cje-inventory-pages",
      value
    );

    setChecked(
      "cje-inventory-assignments",
      value
    );

    setChecked(
      "cje-inventory-discussions",
      value
    );

    setChecked(
      "cje-inventory-classic-quizzes",
      value
    );

    setChecked(
      "cje-inventory-new-quizzes",
      value
    );
  }

  function setCourseWide(
    value
  ) {
    setChecked(
      "cje-inventory-assignment-groups",
      value
    );

    setChecked(
      "cje-inventory-rubrics",
      value
    );

    setChecked(
      "cje-inventory-outcomes",
      value
    );

    setChecked(
      "cje-inventory-tabs",
      value
    );

    setChecked(
      "cje-referenced-files",
      value
    );

    if (!value) {
      setChecked(
        "cje-all-files",
        false
      );
    }
  }

  function setContentTypes(
    types
  ) {
    const wanted =
      new Set(types);

    for (
      const cb of
      document.querySelectorAll(
        ".cje-type-checkbox"
      )
    ) {
      cb.checked =
        wanted.has(
          cb.value
        );
    }
  }

  function setChecked(
    id,
    value
  ) {
    const el =
      document.getElementById(id);

    if (el) {
      el.checked =
        value;
    }
  }

  function setAllModuleChecks(
    value
  ) {
    for (
      const cb of
      document.querySelectorAll(
        ".cje-module-checkbox"
      )
    ) {
      cb.checked =
        value;
    }
  }

  /************************************************************
   * MODULE HELPERS
   ************************************************************/

  function knownType(
    type
  ) {
    return [
      "Page",
      "Assignment",
      "Discussion",
      "Quiz",
      "File",
      "ExternalUrl",
      "ExternalTool",
      "SubHeader"
    ].includes(type);
  }

  function extractPageUrl(
    item
  ) {
    if (item.page_url) {
      return item.page_url;
    }

    for (
      const possible of [
        item.html_url,
        item.url
      ]
    ) {
      if (!possible) {
        continue;
      }

      const match =
        String(possible).match(
          /\/pages\/([^/?#]+)/
        );

      if (match) {
        return decodeURIComponent(
          match[1]
        );
      }
    }

    return null;
  }

  /************************************************************
   * HTML TO AI-READABLE TEXT
   ************************************************************/

  function htmlToAiText(
    html
  ) {
    if (!html) {
      return "";
    }

    const doc =
      new DOMParser()
        .parseFromString(
          String(html),
          "text/html"
        );

    doc
      .querySelectorAll(
        "script, style, noscript"
      )
      .forEach(
        el =>
          el.remove()
      );

    function render(
      node,
      depth = 0
    ) {
      if (
        node.nodeType ===
        Node.TEXT_NODE
      ) {
        return (
          node.nodeValue ||
          ""
        ).replace(
          /\s+/g,
          " "
        );
      }

      if (
        node.nodeType !==
        Node.ELEMENT_NODE
      ) {
        return "";
      }

      const tag =
        node.tagName.toLowerCase();

      const children =
        () =>
          Array.from(
            node.childNodes
          )
            .map(
              child =>
                render(
                  child,
                  depth + 1
                )
            )
            .join("");

      if (
        /^h[1-6]$/.test(tag)
      ) {
        const level =
          Number(tag[1]);

        return (
          `\n${"#".repeat(level)} ` +
          `${cleanInline(children())}\n\n`
        );
      }

      if (
        tag === "p" ||
        tag === "div" ||
        tag === "section" ||
        tag === "article"
      ) {
        const text =
          cleanInline(
            children()
          );

        return text
          ? `\n${text}\n\n`
          : "\n";
      }

      if (tag === "br") {
        return "\n";
      }

      if (tag === "li") {
        const text =
          cleanInline(
            children()
          );

        return text
          ? `\n- ${text}`
          : "";
      }

      if (
        tag === "ul" ||
        tag === "ol"
      ) {
        return (
          `${children()}\n`
        );
      }

      if (tag === "a") {
        const text =
          cleanInline(
            children()
          ) ||
          cleanInline(
            node.getAttribute(
              "title"
            ) ||
            ""
          ) ||
          "link";

        const href =
          node.getAttribute(
            "href"
          );

        return href
          ? `[${text}](${href})`
          : text;
      }

      if (tag === "img") {
        const alt =
          cleanInline(
            node.getAttribute(
              "alt"
            ) ||
            ""
          );

        const src =
          node.getAttribute(
            "src"
          );

        if (
          alt &&
          src
        ) {
          return (
            `[Image: ${alt}](${src})`
          );
        }

        if (alt) {
          return (
            `[Image: ${alt}]`
          );
        }

        if (src) {
          return (
            `[Image](${src})`
          );
        }

        return "[Image]";
      }

      if (tag === "table") {
        const rows =
          Array.from(
            node.querySelectorAll(
              ":scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr, :scope > tr"
            )
          );

        if (!rows.length) {
          return (
            `\n${cleanInline(
              children()
            )}\n`
          );
        }

        return (
          "\n" +
          rows
            .map(
              row => {
                const cells =
                  Array.from(
                    row.children
                  ).filter(
                    cell =>
                      [
                        "TD",
                        "TH"
                      ].includes(
                        cell.tagName
                      )
                  );

                return cells
                  .map(
                    cell =>
                      cleanInline(
                        Array.from(
                          cell.childNodes
                        )
                          .map(
                            child =>
                              render(
                                child,
                                depth + 1
                              )
                          )
                          .join("")
                      )
                  )
                  .join(" | ");
              }
            )
            .join("\n") +
          "\n\n"
        );
      }

      if (
        tag === "strong" ||
        tag === "b"
      ) {
        const text =
          cleanInline(
            children()
          );

        return text
          ? `**${text}**`
          : "";
      }

      if (
        tag === "em" ||
        tag === "i"
      ) {
        const text =
          cleanInline(
            children()
          );

        return text
          ? `*${text}*`
          : "";
      }

      return children();
    }

    return render(doc.body)
      .replace(
        /[ \t]+\n/g,
        "\n"
      )
      .replace(
        /\n[ \t]+/g,
        "\n"
      )
      .replace(
        /\n{3,}/g,
        "\n\n"
      )
      .replace(
        /[ \t]{2,}/g,
        " "
      )
      .trim();
  }

  function cleanInline(
    text
  ) {
    return String(
      text ||
      ""
    )
      .replace(
        /\u00a0/g,
        " "
      )
      .replace(
        /[ \t\r\n]+/g,
        " "
      )
      .trim();
  }

  /************************************************************
   * JSON DOWNLOAD
   ************************************************************/

  function downloadJson(
    data,
    fileName
  ) {
    const json =
      JSON.stringify(
        data,
        null,
        2
      );

    const blob =
      new Blob(
        [json],
        {
          type:
            "application/json;charset=utf-8"
        }
      );

    const url =
      URL.createObjectURL(blob);

    const link =
      document.createElement("a");

    link.href =
      url;

    link.download =
      fileName;

    document.body.appendChild(
      link
    );

    link.click();
    link.remove();

    setTimeout(
      () =>
        URL.revokeObjectURL(
          url
        ),
      30000
    );
  }

  function makeFileName(
    course,
    settings
  ) {
    const courseName =
      course?.name ||
      `course-${courseId}`;

    const safeName =
      String(courseName)
        .toLowerCase()
        .replace(
          /[^a-z0-9]+/g,
          "-"
        )
        .replace(
          /^-+|-+$/g,
          ""
        )
        .slice(
          0,
          80
        );

    const preset =
      String(
        settings.preset ||
        "custom"
      )
        .replace(
          /[^a-z0-9]+/gi,
          "-"
        )
        .toLowerCase();

    const date =
      new Date()
        .toISOString()
        .replace(
          /[:.]/g,
          "-"
        );

    return (
      `canvas-course-ai-export-` +
      `${safeName}-` +
      `${courseId}-` +
      `${preset}-` +
      `${date}.json`
    );
  }

  /************************************************************
   * EXPORT STATE
   ************************************************************/

  function resetExportStores() {
    appState.errors = [];
    appState.warnings = [];

    for (
      const map of
      Object.values(
        contentStore
      )
    ) {
      map.clear();
    }

    fileReferences.clear();
  }

  function recordError(
    errorObject
  ) {
    appState.errors.push(
      errorObject
    );
  }

  /************************************************************
   * MODAL SHELL
   ************************************************************/

  function createShell() {
    const overlay =
      document.createElement("div");

    overlay.id =
      APP_ID;

    overlay.style.position =
      "fixed";

    overlay.style.inset =
      "0";

    overlay.style.background =
      "rgba(0,0,0,.35)";

    overlay.style.zIndex =
      "999999";

    overlay.style.display =
      "flex";

    overlay.style.alignItems =
      "center";

    overlay.style.justifyContent =
      "center";

    overlay.style.padding =
      "18px";

    overlay.style.boxSizing =
      "border-box";

    const panel =
      document.createElement("div");

    panel.style.width =
      "900px";

    panel.style.maxWidth =
      "calc(100vw - 36px)";

    panel.style.height =
      "calc(100vh - 36px)";

    panel.style.maxHeight =
      "calc(100vh - 36px)";

    panel.style.background =
      "#fff";

    panel.style.borderRadius =
      "14px";

    panel.style.boxShadow =
      "0 18px 50px rgba(0,0,0,.3)";

    panel.style.overflow =
      "hidden";

    panel.style.fontFamily =
      "Arial, Helvetica, sans-serif";

    panel.style.color =
      "#222";

    panel.style.display =
      "flex";

    panel.style.flexDirection =
      "column";

    const header =
      document.createElement("div");

    header.style.flex =
      "0 0 auto";

    header.style.display =
      "flex";

    header.style.alignItems =
      "center";

    header.style.justifyContent =
      "space-between";

    header.style.padding =
      "14px 16px";

    header.style.background =
      "#510C76";

    header.style.color =
      "#fff";

    const headerTitle =
      document.createElement("div");

    headerTitle.style.fontWeight =
      "700";

    headerTitle.textContent =
      "Canvas Course AI JSON Exporter";

    const close =
      document.createElement("button");

    close.textContent =
      "×";

    close.setAttribute(
      "aria-label",
      "Close"
    );

    close.style.border =
      "0";

    close.style.background =
      "transparent";

    close.style.color =
      "#fff";

    close.style.fontSize =
      "24px";

    close.style.cursor =
      "pointer";

    close.addEventListener(
      "click",
      removeExistingApp
    );

    header.append(
      headerTitle,
      close
    );

    const content =
      document.createElement("div");

    content.style.flex =
      "1 1 auto";

    content.style.minHeight =
      "0";

    content.style.display =
      "grid";

    content.style.gridTemplateColumns =
      "minmax(0, 1fr) 320px";

    content.style.overflow =
      "hidden";

    const body =
      document.createElement("div");

    body.setAttribute(
      "data-cje-body",
      "true"
    );

    body.style.padding =
      "16px";

    body.style.overflowY =
      "auto";

    body.style.overflowX =
      "hidden";

    body.style.minHeight =
      "0";

    body.style.boxSizing =
      "border-box";

    const logWrap =
      document.createElement("div");

    logWrap.style.borderLeft =
      "1px solid #eee";

    logWrap.style.background =
      "#fafafa";

    logWrap.style.padding =
      "12px";

    logWrap.style.overflowY =
      "auto";

    logWrap.style.overflowX =
      "hidden";

    logWrap.style.minHeight =
      "0";

    logWrap.style.boxSizing =
      "border-box";

    const logTitle =
      document.createElement("div");

    logTitle.style.fontWeight =
      "700";

    logTitle.style.marginBottom =
      "8px";

    logTitle.textContent =
      "Status";

    const logBox =
      document.createElement("pre");

    logBox.id =
      "cje-log";

    logBox.style.whiteSpace =
      "pre-wrap";

    logBox.style.overflowWrap =
      "anywhere";

    logBox.style.fontFamily =
      "Consolas, Monaco, monospace";

    logBox.style.fontSize =
      "12px";

    logBox.style.margin =
      "0";

    logBox.style.color =
      "#333";

    logWrap.append(
      logTitle,
      logBox
    );

    content.append(
      body,
      logWrap
    );

    panel.append(
      header,
      content
    );

    overlay.append(panel);

    document.body.appendChild(
      overlay
    );

    return overlay;
  }

  /************************************************************
   * UI HELPERS
   ************************************************************/

  function makeLabel(
    text
  ) {
    const label =
      document.createElement("label");

    label.style.display =
      "block";

    label.style.fontWeight =
      "700";

    label.style.marginBottom =
      "4px";

    label.textContent =
      text;

    return label;
  }

  function makeSmallButton(
    text
  ) {
    const button =
      document.createElement("button");

    button.type =
      "button";

    button.textContent =
      text;

    button.style.fontSize =
      "12px";

    button.style.padding =
      "4px 7px";

    button.style.border =
      "1px solid #ccc";

    button.style.borderRadius =
      "6px";

    button.style.background =
      "#fff";

    button.style.cursor =
      "pointer";

    return button;
  }

  function makeCheckboxRow(
    className,
    value,
    labelText,
    checked
  ) {
    const label =
      document.createElement("label");

    label.style.display =
      "flex";

    label.style.alignItems =
      "center";

    label.style.gap =
      "6px";

    label.style.cursor =
      "pointer";

    const cb =
      document.createElement("input");

    cb.type =
      "checkbox";

    cb.className =
      className;

    cb.value =
      value;

    cb.checked =
      checked;

    const text =
      document.createElement("span");

    text.textContent =
      labelText;

    label.append(
      cb,
      text
    );

    return label;
  }

  function makeOptionCheckbox(
    id,
    labelText,
    checked
  ) {
    const label =
      document.createElement("label");

    label.style.display =
      "flex";

    label.style.alignItems =
      "center";

    label.style.gap =
      "6px";

    label.style.cursor =
      "pointer";

    const cb =
      document.createElement("input");

    cb.type =
      "checkbox";

    cb.id =
      id;

    cb.className =
      "cje-option-checkbox";

    cb.checked =
      checked;

    const text =
      document.createElement("span");

    text.textContent =
      labelText;

    label.append(
      cb,
      text
    );

    return label;
  }

  /************************************************************
   * LOGGING
   ************************************************************/

  function log(
    message
  ) {
    const logBox =
      document.getElementById(
        "cje-log"
      );

    const stamp =
      new Date()
        .toLocaleTimeString();

    const line =
      `[${stamp}] ${message}`;

    if (logBox) {
      logBox.textContent +=
        `${line}\n`;

      logBox.scrollTop =
        logBox.scrollHeight;
    }

    console.log(line);
  }

  function logError(
    message,
    error
  ) {
    const detail =
      error?.message ||
      String(error);

    recordError({
      context:
        message,

      message:
        detail
    });

    log(
      `${message}: ${detail}`
    );

    console.error(
      message,
      error
    );
  }

  /************************************************************
   * GENERAL HELPERS
   ************************************************************/

  function getCourseIdFromUrl() {
    const match =
      location.pathname.match(
        /\/courses\/(\d+)/
      );

    return match
      ? match[1]
      : null;
  }

  function removeExistingApp() {
    const existing =
      document.getElementById(
        APP_ID
      );

    if (existing) {
      existing.remove();
    }
  }

  function numberSort(
    a,
    b
  ) {
    return (
      Number(
        a ||
        0
      ) -
      Number(
        b ||
        0
      )
    );
  }

  function escapeHtml(
    value
  ) {
    return String(value)
      .replaceAll(
        "&",
        "&amp;"
      )
      .replaceAll(
        "<",
        "&lt;"
      )
      .replaceAll(
        ">",
        "&gt;"
      )
      .replaceAll(
        '"',
        "&quot;"
      )
      .replaceAll(
        "'",
        "&#039;"
      );
  }
})();
