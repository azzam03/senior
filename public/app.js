/* DATA CLASSIFICATION & GOVERNANCE PLATFORM
 * This file owns browser behavior only. The visible app structure lives in index.html,
 * the visual system lives in styles.css, and this script connects the UI to backend APIs,
 * SSE classification streams, upload handling, review workflows, and export actions.
 */

const state = {
  user: null,
  systems: [],
  currentSystem: null,
  activeTab: "overview",
  records: [],
  personalRecords: [],
  originalColumns: [],
  page: 1,
  pageSize: 25,
  totalPages: 1,
  totalRecords: 0,
  search: "",
  sortBy: "rowIndex",
  sortDir: "asc",
  viewMode: "classification",
  classifierStream: null,
  classifierJobId: null,
  classifierMode: null,
  liveUpdatedRecordId: null,
};

const qs = (selector, root = document) => root.querySelector(selector);
const qsa = (selector, root = document) => Array.from(root.querySelectorAll(selector));

document.addEventListener("DOMContentLoaded", boot);

/* Bootstrapping connects static HTML to live server state. It decides whether to show
 * authentication or the authenticated SaaS shell and wires all event handlers once.
 */
async function boot() {
  bindAuthEvents();
  bindShellEvents();
  bindSystemEvents();
  bindDataEvents();
  bindGovernanceEvents();

  try {
    const session = await api("/api/auth/me");
    state.user = session.user;
    showApp();
    await loadDashboard();
  } catch (_error) {
    showAuth();
  }
}

/* API helper centralizes request behavior so authentication, JSON parsing,
 * and error display are consistent across every workflow.
 */
async function api(url, options = {}) {
  const response = await fetch(url, {
    credentials: "same-origin",
    headers: options.body instanceof FormData ? undefined : { "Content-Type": "application/json" },
    ...options,
  });
  const contentType = response.headers.get("content-type") || "";
  const data = contentType.includes("application/json") ? await response.json() : await response.text();

  // Session expired while the user was on the page.  Redirect to login so
  // they don't see a confusing series of toast errors — their work is safe in
  // the database; they just need to sign in again.
  if (response.status === 401 && state.user) {
    state.user = null;
    window.location.replace("/");
    // Return a never-resolving promise so the calling code doesn't run
    // after the redirect is initiated.
    return new Promise(() => {});
  }

  if (!response.ok) {
    throw new Error(data.error || data.message || "Request failed");
  }
  return data;
}

async function downloadSystemExport() {
  if (!state.currentSystem) return;
  closeActionsMenu();
  const button = qs("#exportExcelButton");
  button.disabled = true;

  try {
    const response = await fetch(`/api/systems/${state.currentSystem.id}/export`, {
      credentials: "same-origin",
    });

    if (response.status === 401 && state.user) {
      state.user = null;
      window.location.replace("/");
      return;
    }

    if (!response.ok) {
      throw new Error(await exportErrorMessage(response));
    }

    const contentType = response.headers.get("content-type") || "";
    if (contentType.includes("text/html")) {
      const text = await response.text();
      throw new Error(
        `Export endpoint returned HTML instead of XLSX. ` +
        `Check the API route and server configuration. ` +
        `(status ${response.status})`
      );
    }

    const blob = await response.blob();
    if (!blob.size) throw new Error("The exported workbook was empty.");

    const fileName = fileNameFromContentDisposition(response.headers.get("content-disposition"))
      || `${safeFileStem(state.currentSystem.name)}_classification_report.xlsx`;
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast("Report download started.");
  } finally {
    button.disabled = false;
  }
}

async function exportErrorMessage(response) {
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    const payload = await response.json().catch(() => ({}));
    return payload.error || payload.message || `Export failed with status ${response.status}.`;
  }
  const text = await response.text().catch(() => "");
  return text || `Export failed with status ${response.status}.`;
}

function fileNameFromContentDisposition(header) {
  const value = String(header || "");
  const utf8Match = value.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match) return decodeURIComponent(utf8Match[1]);
  const asciiMatch = value.match(/filename="?([^";]+)"?/i);
  return asciiMatch ? asciiMatch[1] : "";
}

function safeFileStem(value) {
  return (
    String(value || "system")
      .trim()
      .replace(/\s+/g, "_")
      .replace(/[^a-z0-9_]/gi, "")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "") || "system"
  );
}

function bindAuthEvents() {
  qsa(".auth-tab").forEach((button) => {
    button.addEventListener("click", () => switchAuthTab(button.dataset.authTab));
  });
  qsa(".auth-link").forEach((button) => {
    button.addEventListener("click", () => switchAuthTab(button.dataset.authLink));
  });

  qs("#accountTypeSelect").addEventListener("change", (event) => {
    qs("#companyNameField").classList.toggle("hidden", event.target.value !== "Company");
  });

  qs("#loginForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const body = formJson(event.currentTarget);
    try {
      const result = await api("/api/auth/login", { method: "POST", body: JSON.stringify(body) });
      state.user = result.user;
      window.location.replace("/app");
    } catch (error) {
      toast(error.message);
    }
  });

  qs("#signupForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const body = formJson(event.currentTarget);
    try {
      const result = await api("/api/auth/signup", { method: "POST", body: JSON.stringify(body) });
      state.user = result.user;
      window.location.replace("/app");
    } catch (error) {
      toast(error.message);
    }
  });

  qs("#forgotForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    try {
      const result = await api("/api/auth/forgot-password", {
        method: "POST",
        body: JSON.stringify(formJson(event.currentTarget)),
      });
      qs("#resetTokenOutput").textContent = result.resetToken
        ? `Reset token: ${result.resetToken}`
        : result.message;
      if (result.resetToken) qs("#resetTokenInput").value = result.resetToken;
      switchAuthTab("reset");
    } catch (error) {
      toast(error.message);
    }
  });

  qs("#resetForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    try {
      await api("/api/auth/reset-password", { method: "POST", body: JSON.stringify(formJson(event.currentTarget)) });
      toast("Password reset complete.");
      switchAuthTab("login");
    } catch (error) {
      toast(error.message);
    }
  });
}

function switchAuthTab(tab) {
  qsa(".auth-tab").forEach((button) => button.classList.toggle("active", button.dataset.authTab === tab));
  qsa(".auth-link").forEach((button) => button.classList.toggle("active", button.dataset.authLink === tab));
  qsa("[data-auth-panel]").forEach((panel) => panel.classList.toggle("active", panel.dataset.authPanel === tab));
}

function bindShellEvents() {
  qs("#logoutButton").addEventListener("click", async () => {
    await api("/api/auth/logout", { method: "POST", body: JSON.stringify({}) }).catch(() => null);
    state.user = null;
    state.currentSystem = null;
    state.records = [];
    state.personalRecords = [];
    window.location.replace("/");
  });

  qs("#dashboardNav").addEventListener("click", showDashboard);
  qs("#backToDashboard").addEventListener("click", showDashboard);
  qs("#openAddSystemModal").addEventListener("click", () => qs("#addSystemModal").showModal());
  qs("#closeAddSystemModal").addEventListener("click", () => qs("#addSystemModal").close());
  qs("#cancelAddSystem").addEventListener("click", () => qs("#addSystemModal").close());

  qs("#addSystemForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    try {
      const result = await api("/api/systems", {
        method: "POST",
        body: JSON.stringify(formJson(form)),
      });
      qs("#addSystemModal").close();
      form.reset();
      await loadDashboard();
      await openSystem(result.system.id);
    } catch (error) {
      toast(error.message);
    }
  });
}

function bindSystemEvents() {
  qsa(".tab").forEach((button) => {
    button.addEventListener("click", () => switchSystemTab(button.dataset.tab));
  });
}

/* System Data is the operational center. These handlers manage upload visibility,
 * data retrieval, live SSE classification, sorting, paging, and Excel export.
 */
function bindDataEvents() {
  qs("#metadataFileInput").addEventListener("change", (event) => {
    const fileName = event.target.files?.[0]?.name;
    if (fileName) {
      setClassificationStatus(`Ready to import ${fileName}`, 0);
      qs("#uploadForm").requestSubmit();
    }
  });

  qs("#importExcelButton").addEventListener("click", () => {
    closeActionsMenu();
    qs("#metadataFileInput").click();
  });
  qs("#importFilePrimaryButton").addEventListener("click", () => qs("#metadataFileInput").click());

  qs("#uploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!state.currentSystem) return;
    const fileInput = qs("#metadataFileInput");
    if (!fileInput.files.length) {
      toast("Choose an Excel or CSV file first.");
      return;
    }
    const formData = new FormData();
    formData.append("metadataFile", fileInput.files[0]);
    try {
      setClassificationStatus(`Importing ${fileInput.files[0].name}`, 0);
      const result = await api(`/api/systems/${state.currentSystem.id}/upload`, { method: "POST", body: formData });
      state.currentSystem.lastUploadFileName = result.fileName;
      state.originalColumns = result.originalColumns;
      updateFileNameLabels(result.fileName);
      toast(`Imported ${result.rowsImported} rows from ${result.fileName}.`);
      await reloadCurrentSystem();
      await loadRecords();
      renderOverview(result.summary);
      setClassificationStatus("Ready for classification", 0);
    } catch (error) {
      toast(error.message);
      setClassificationStatus("Import failed", 0);
    }
  });

  qs("#classifyPageButton").addEventListener("click", () => {
    closeActionsMenu();
    startClassification("page");
  });
  qs("#classifyAllButton").addEventListener("click", () => {
    closeActionsMenu();
    startClassification("all");
  });
  qs("#exportExcelButton").addEventListener("click", (event) => {
    event.stopPropagation();
    downloadSystemExport().catch((error) => toast(error.message || "Export failed."));
  });

  qs("#recordSearch").addEventListener("input", debounce((event) => {
    state.search = event.target.value;
    state.page = 1;
    loadRecords();
  }, 250));

  qsa("[data-view-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      state.viewMode = button.dataset.viewMode;
      qsa("[data-view-mode]").forEach((item) => item.classList.toggle("active", item === button));
      renderRecordsTable();
    });
  });

  qs("#prevPageButton").addEventListener("click", () => {
    if (state.page > 1) {
      state.page -= 1;
      loadRecords();
    }
  });

  qs("#nextPageButton").addEventListener("click", () => {
    if (state.page < state.totalPages) {
      state.page += 1;
      loadRecords();
    }
  });
}

function bindGovernanceEvents() {
  qs("#contextContent").addEventListener("input", (event) => {
    const words = wordCount(event.target.value);
    qs("#contextWordCount").textContent = `${words} / 200 words`;
    qs("#contextWordCount").classList.toggle("warning", words > 200);
  });

  qs("#contextForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!state.currentSystem) return;
    const form = event.currentTarget;
    const body = formJson(form);
    if (wordCount(body.content) > 200) {
      toast("Context content must be 200 words or fewer.");
      return;
    }
    try {
      await api(`/api/systems/${state.currentSystem.id}/context`, { method: "POST", body: JSON.stringify(body) });
      form.reset();
      qs("#contextWordCount").textContent = "0 / 200 words";
      await loadContext();
    } catch (error) {
      toast(error.message);
    }
  });

  qs("#pdplForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!state.currentSystem) return;
    try {
      await api(`/api/systems/${state.currentSystem.id}/pdpl`, {
        method: "PUT",
        body: JSON.stringify(formJson(event.currentTarget)),
      });
      toast("PDPL notes saved.");
      await loadPdpl();
    } catch (error) {
      toast(error.message);
    }
  });

  qs("#linkForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!state.currentSystem) return;
    const form = event.currentTarget;
    try {
      await api(`/api/systems/${state.currentSystem.id}/links`, {
        method: "POST",
        body: JSON.stringify(formJson(form)),
      });
      form.reset();
      await loadLinks();
    } catch (error) {
      toast(error.message);
    }
  });

  qs("#approvePersonalPageButton").addEventListener("click", approvePersonalOnPage);
}

function showAuth() {
  qs("#authScreen").classList.remove("hidden");
  qs("#appShell").classList.add("hidden");
  if (window.location.pathname !== "/") {
    window.history.replaceState({}, "", "/");
  }
}

function showApp() {
  qs("#authScreen").classList.add("hidden");
  qs("#appShell").classList.remove("hidden");
  qs("#dashboardPage").classList.add("active");
  qs("#systemPage").classList.remove("active");
  qs("#welcomeTitle").textContent = `Welcome, ${state.user.name}`;
  qs("#profileName").textContent = state.user.name;
  qs("#profileEmail").textContent = state.user.email;
  qs("#userAvatar").textContent = state.user.avatarInitials || "U";
  if (window.location.pathname !== "/app") {
    window.history.replaceState({}, "", "/app");
  }
}

async function loadDashboard() {
  const [summary, systemsResult] = await Promise.all([api("/api/dashboard"), api("/api/systems")]);
  state.systems = systemsResult.systems;
  qs("#statTotalSystems").textContent = summary.totalSystems;
  qs("#statInProgress").textContent = summary.systemsInProgress;
  qs("#statCompleted").textContent = summary.completedSystems;
  qs("#statCompletion").textContent = `${summary.completionPercentage}%`;
  renderSystems();
}

function renderSystems() {
  const grid = qs("#systemsGrid");
  const template = qs("#systemCardTemplate");
  grid.replaceChildren();
  qs("#emptySystems").classList.toggle("hidden", state.systems.length > 0);

  for (const system of state.systems) {
    const fragment = template.content.cloneNode(true);
    const card = qs(".system-card", fragment);
    qs("h3", card).textContent = system.name;
    qs(".status-pill", card).textContent = system.status;
    qs(".system-card-meta", card).innerHTML = "";
    qs(".system-card-meta", card).append(
      textLine(`Owner: ${system.owner}`),
      textLine(`DBA: ${system.dba}`),
      textLine(`Group: ${system.systemGroup}`)
    );
    qs("progress", card).value = system.summary.classificationProgress || 0;
    qs(".records-count", card).textContent = `${system.summary.totalRecords || 0} records`;
    qs(".file-chip", card).textContent = system.lastUploadFileName || "No file uploaded";
    qs(".open-system", card).addEventListener("click", () => openSystem(system.id));
    qs(".delete-system", card).addEventListener("click", () => deleteSystem(system));
    grid.append(card);
  }
}

function textLine(text) {
  const span = document.createElement("span");
  span.textContent = text;
  return span;
}

async function deleteSystem(system) {
  const confirmed = window.confirm(
    `Delete "${system.name}" and all related uploaded rows, classifications, PDPL notes, context, links, and jobs?\n\nThis cannot be undone.`
  );
  if (!confirmed) return;

  try {
    await api(`/api/systems/${system.id}`, { method: "DELETE" });
    if (state.currentSystem?.id === system.id) {
      state.currentSystem = null;
      showDashboard();
    }
    toast(`Deleted "${system.name}".`);
    await loadDashboard();
  } catch (error) {
    toast(error.message || "Delete failed. Please try again.");
  }
}

async function openSystem(systemId) {
  try {
    const result = await api(`/api/systems/${systemId}`);
    state.currentSystem = result.system;
    state.originalColumns = result.originalColumns;
    state.page = 1;
    state.search = "";
    qs("#recordSearch").value = "";
    renderSystemHeader();
    showSystemPage();
    await switchSystemTab("overview");
  } catch (error) {
    toast(error.message || "Could not open system. Please try again.");
  }
}

async function reloadCurrentSystem() {
  if (!state.currentSystem) return;
  try {
    const result = await api(`/api/systems/${state.currentSystem.id}`);
    state.currentSystem = result.system;
    state.originalColumns = result.originalColumns;
    renderSystemHeader();
  } catch (error) {
    // Non-fatal — the system may have been deleted in another tab.
    console.warn("reloadCurrentSystem failed:", error.message);
  }
}

function renderSystemHeader() {
  const system = state.currentSystem;
  qs("#systemNameTitle").textContent = system.name;
  qs("#systemOwnerMeta").textContent = `Owner: ${system.owner}`;
  qs("#systemDbaMeta").textContent = `DBA: ${system.dba}`;
  qs("#systemOwnerEmailMeta").textContent = system.ownerEmail;
  qs("#systemGroupMeta").textContent = `Group: ${system.systemGroup}`;
  qs("#systemStatusPill").textContent = system.status;
  updateFileNameLabels(system.lastUploadFileName || "No file uploaded");
}

function updateFileNameLabels(fileName) {
  const value = fileName || "No file uploaded";
  qs("#activeFileChip").textContent = value;
  qs("#dataTabFileName").textContent = value;
  qs("#classificationFileName").textContent = value;
  qs("#overviewFileName").textContent = value;
  qs("#systemDataActiveFile").textContent = value;
}

function showDashboard() {
  closeClassifierStream();
  qs("#dashboardPage").classList.add("active");
  qs("#systemPage").classList.remove("active");
  loadDashboard().catch((error) => toast(error.message || "Dashboard failed to load."));
}

function showSystemPage() {
  qs("#dashboardPage").classList.remove("active");
  qs("#systemPage").classList.add("active");
}

async function switchSystemTab(tab) {
  state.activeTab = tab;
  qsa(".tab").forEach((button) => button.classList.toggle("active", button.dataset.tab === tab));
  const panelMap = {
    overview: "#overviewTab",
    "system-data": "#systemDataTab",
    "personal-data": "#personalDataTab",
    "system-context": "#systemContextTab",
    pdpl: "#pdplTab",
    "system-links": "#systemLinksTab",
  };
  Object.entries(panelMap).forEach(([key, selector]) => qs(selector).classList.toggle("active", key === tab));

  if (!state.currentSystem) return;
  try {
    if (tab === "overview") await loadOverview();
    if (tab === "system-data") await loadRecords();
    if (tab === "personal-data") await loadPersonalData();
    if (tab === "system-context") await loadContext();
    if (tab === "pdpl") await loadPdpl();
    if (tab === "system-links") await loadLinks();
  } catch (error) {
    toast(error.message || "Failed to load tab content. Please try again.");
  }
}

async function loadOverview() {
  await reloadCurrentSystem();
  renderOverview(state.currentSystem.summary);
}

function renderOverview(summary) {
  if (!summary) return;
  const classificationProgress = clampPercent(summary.classificationProgress);
  qs("#overviewClassified").textContent = summary.classified;
  qs("#overviewClassifiedPct").textContent = `${summary.classifiedPercentage}%`;
  qs("#overviewPersonal").textContent = summary.personal;
  qs("#overviewPersonalPct").textContent = `${summary.personalPercentage}%`;
  qs("#overviewPending").textContent = summary.pending;
  qs("#overviewPendingPct").textContent = `${summary.pendingPercentage}%`;
  qs("#overviewReviewQueue").textContent = summary.reviewQueue;
  qs("#overviewLowConfidence").textContent = summary.lowConfidence;
  qs("#overviewPolicyCount").textContent = summary.policyRecommendationCount;
  qs("#overviewTablesPersonal").textContent = summary.tablesWithPersonalData;
  qs("#overviewProgressCircle").setAttribute("stroke-dasharray", `${classificationProgress} ${100 - classificationProgress}`);
  qs("#overviewProgressLabel").textContent = `${classificationProgress}%`;
  updateClassificationMetrics({ ...summary, classificationProgress });
  updateFileNameLabels(state.currentSystem?.lastUploadFileName || "No file uploaded");
  renderChart(qs("#confidentialityChart"), summary.confidentiality || {}, [
    "Confidential",
    "Secret",
    "Top Secret",
    "Public",
    "Pending",
  ]);
  renderChart(qs("#personalChart"), summary.personalDistribution || {}, ["Has Personal Data", "No Personal Data"]);
  renderPersonalTypeSummary(summary.personalDataTypes || {});
}

function renderPersonalTypeSummary(values) {
  const container = qs("#personalTypeSummary");
  container.replaceChildren();
  const entries = Object.entries(values || {})
    .map(([label, count]) => [label, Number(count || 0)])
    .filter(([, count]) => count > 0)
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  const total = entries.reduce((sum, [, count]) => sum + count, 0);

  if (!entries.length) {
    const empty = document.createElement("div");
    empty.className = "personal-type-empty";
    empty.textContent = "No personal-data categories have been identified yet.";
    container.append(empty);
    return;
  }

  const overview = document.createElement("div");
  overview.className = "personal-type-overview";
  overview.append(
    personalTypeMetric("Categories", entries.length),
    personalTypeMetric("Personal Fields", total),
    personalTypeMetric("PDPL Review", `${entries.length} tracks`)
  );
  container.append(overview);

  for (const [labelText, countValue] of entries) {
    const percentage = clampPercent((countValue / Math.max(1, total)) * 100);
    const row = document.createElement("article");
    row.className = "personal-type-row";

    const header = document.createElement("div");
    header.className = "personal-type-row-header";
    const label = document.createElement("strong");
    label.textContent = labelText;
    const count = document.createElement("span");
    count.textContent = `${countValue} fields`;
    header.append(label, count);

    const bar = document.createElement("div");
    bar.className = "personal-type-bar";
    const fill = document.createElement("i");
    fill.style.width = `${percentage}%`;
    bar.append(fill);

    const detail = document.createElement("div");
    detail.className = "personal-type-detail";
    const share = document.createElement("span");
    share.textContent = `${percentage}% of personal-data fields`;
    const guidance = document.createElement("em");
    guidance.textContent = pdplGuidanceForType(labelText);
    detail.append(share, guidance);

    row.append(header, bar, detail);
    container.append(row);
  }
}

function personalTypeMetric(label, value) {
  const item = document.createElement("div");
  const span = document.createElement("span");
  const strong = document.createElement("strong");
  span.textContent = label;
  strong.textContent = value;
  item.append(span, strong);
  return item;
}

function pdplGuidanceForType(type) {
  const text = String(type || "").toLowerCase();
  if (/health|medical|biometric|genetic|religion|disability/.test(text)) {
    return "Enhanced safeguards, explicit lawful basis, strict access.";
  }
  if (/government|national|identifier|id|passport|license/.test(text)) {
    return "Masking, restricted access, purpose limitation.";
  }
  if (/financial|salary|bank|payment|card|income/.test(text)) {
    return "Restricted processing, retention controls, audit trail.";
  }
  if (/contact|email|phone|address/.test(text)) {
    return "Notice, lawful basis, retention and access review.";
  }
  if (/name|individual|reference|demographic/.test(text)) {
    return "Lawful basis, minimization, subject-rights readiness.";
  }
  return "Confirm lawful basis, retention, access control, and PDPL rights.";
}

function renderChart(container, values, preferredOrder = []) {
  container.replaceChildren();
  const orderedLabels = [
    ...preferredOrder,
    ...Object.keys(values).filter((label) => !preferredOrder.includes(label)),
  ];
  const entries = orderedLabels.map((label) => [label, Number(values[label] || 0)]);
  const total = entries.reduce((sum, [, count]) => sum + count, 0);

  if (!entries.length || total === 0) {
    const empty = document.createElement("div");
    empty.className = "chart-empty";
    empty.textContent = "No data available yet";
    container.append(empty);
    return;
  }

  for (const [label, count] of entries) {
    const percentage = clampPercent((count / total) * 100);
    const row = document.createElement("div");
    row.className = `chart-row ${chartClass(label)}`;
    const labelNode = document.createElement("span");
    labelNode.textContent = label;
    const bar = document.createElement("div");
    bar.className = "chart-bar";
    const fill = document.createElement("i");
    fill.style.width = `${percentage}%`;
    bar.append(fill);
    const countNode = document.createElement("strong");
    countNode.textContent = `${count} (${percentage}%)`;
    row.append(labelNode, bar, countNode);
    container.append(row);
  }
}

function clampPercent(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  return Math.max(0, Math.min(100, Math.round(number)));
}

function chartClass(label) {
  return `chart-${String(label).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "item"}`;
}

async function loadRecords() {
  if (!state.currentSystem) return;
  const params = new URLSearchParams({
    page: state.page,
    pageSize: state.pageSize,
    search: state.search,
    sortBy: state.sortBy,
    sortDir: state.sortDir,
  });
  const result = await api(`/api/systems/${state.currentSystem.id}/records?${params.toString()}`);
  state.records = result.records;
  state.originalColumns = result.originalColumns;
  state.totalPages = result.totalPages;
  state.totalRecords = result.total;
  renderRecordsTable();
  renderOverview(result.summary);
}

function renderRecordsTable() {
  renderTable({
    head: qs("#recordsTableHead"),
    body: qs("#recordsTableBody"),
    records: state.records,
    columns: systemDataColumns(),
    editable: state.viewMode === "edit",
  });
  qs("#recordCounter").textContent = `${state.totalRecords} records`;
  qs("#paginationLabel").textContent = `Page ${state.page} of ${state.totalPages}`;
  qs("#prevPageButton").disabled = state.page <= 1;
  qs("#nextPageButton").disabled = state.page >= state.totalPages;
}

function systemDataColumns() {
  const original = state.originalColumns.map((column) => ({
    key: `original.${column}`,
    label: column,
    originalColumn: column,
    sortable: true,
  }));
  return [
    ...original,
    { key: "confidentiality", label: "Confidentiality", sortable: true, editable: true },
    { key: "confReason", label: "Conf Reason", className: "reason-cell", editable: true },
    { key: "personalData", label: "Personal Data", sortable: true, editable: true },
    { key: "personalReason", label: "Personal Reason", className: "reason-cell", editable: true },
    { key: "personalDataType", label: "Personal Data Type", sortable: true },
    { key: "review", label: "Review" },
  ];
}

async function loadPersonalData() {
  if (!state.currentSystem) return;
  const result = await api(`/api/systems/${state.currentSystem.id}/records?personalOnly=true&page=1&pageSize=100`);
  state.personalRecords = result.records;
  renderPersonalDataTable();
}

function renderPersonalDataTable() {
  renderTable({
    head: qs("#personalTableHead"),
    body: qs("#personalTableBody"),
    records: state.personalRecords,
    columns: [
      { key: "tableName", label: "Table" },
      { key: "columnName", label: "Column" },
      { key: "dataType", label: "Data Type" },
      { key: "confidentiality", label: "Confidentiality" },
      { key: "confReason", label: "Conf Reason", className: "reason-cell" },
      { key: "personalReason", label: "Personal Reason", className: "reason-cell" },
      { key: "personalDataType", label: "Personal Data Type" },
      { key: "pseudonymizable", label: "Can be pseudonymized" },
      { key: "anonymizable", label: "Can be anonymized" },
      { key: "specialCategory", label: "Special Category" },
      { key: "auditTrail", label: "Audit Trail", className: "reason-cell" },
      { key: "needsReview", label: "Needs Review" },
      { key: "confidenceScore", label: "Confidence Score" },
      { key: "policyRecommendation", label: "Policy Recommendation", className: "policy-cell" },
      { key: "personalApproval", label: "Approval" },
    ],
    editable: false,
  });
}

/* Table rendering is shared by System Data and Personal Data so sorting, badges,
 * edit controls, and workflow actions stay consistent across the platform.
 */
function renderTable({ head, body, records, columns, editable }) {
  head.replaceChildren();
  body.replaceChildren();

  const headerRow = document.createElement("tr");
  for (const column of columns) {
    const th = document.createElement("th");
    if (column.sortable) {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = sortLabel(column.label, column.originalColumn || column.key);
      button.addEventListener("click", () => sortBy(column.originalColumn || column.key));
      th.append(button);
    } else {
      th.textContent = column.label;
    }
    headerRow.append(th);
  }
  head.append(headerRow);

  for (const record of records) {
    const tr = document.createElement("tr");
    if (record.id === state.liveUpdatedRecordId) tr.classList.add("row-live-update");
    for (const column of columns) {
      const td = document.createElement("td");
      if (column.className) td.className = column.className;
      renderCell(td, record, column, editable);
      tr.append(td);
    }
    body.append(tr);
  }
}

function renderCell(td, record, column, editable) {
  if (column.originalColumn) {
    td.textContent = record.original[column.originalColumn] ?? "";
    return;
  }

  if (column.key === "review") {
    td.append(reviewControl(record, editable));
    return;
  }

  if (column.key === "personalApproval") {
    td.append(personalApprovalControl(record));
    return;
  }

  if (editable && column.editable) {
    td.append(editControl(record, column.key));
    return;
  }

  if (column.key === "confidentiality") {
    td.append(badge(record.confidentiality || "Pending", confidentialityClass(record.confidentiality)));
    return;
  }

  if (["personalData", "pseudonymizable", "anonymizable", "specialCategory"].includes(column.key)) {
    const value = record[column.key] || "No";
    td.append(badge(value, value === "Yes" ? "badge-yes" : "badge-no"));
    return;
  }

  if (column.key === "needsReview" || column.key === "pushToClient") {
    const value = record[column.key] ? "Yes" : "No";
    td.append(badge(value, value === "Yes" ? "badge-yes" : "badge-no"));
    return;
  }

  if (column.key === "confidenceScore") {
    td.textContent = record.confidenceScore == null ? "" : `${Math.round(Number(record.confidenceScore) * 100)}%`;
    return;
  }

  td.textContent = record[column.key] ?? "";
}

function reviewControl(record, editable) {
  const reviewed = record.systemReviewStatus === "Approved";
  const button = document.createElement("button");
  button.className = `review-toggle ${reviewed ? "reviewed" : ""}`;
  button.type = "button";
  button.textContent = reviewed ? "Approved" : "Review";
  button.title = editable ? "Toggle review state" : "Switch to Edit View to update review state";
  button.disabled = !editable;
  if (editable) {
    button.addEventListener("click", () => {
      updateRecord(record.id, { reviewed: !reviewed }).catch((error) => toast(error.message));
    });
  }
  return button;
}

function personalApprovalControl(record) {
  const approved = record.personalReviewStatus === "Approved";
  const button = document.createElement("button");
  button.className = `review-toggle ${approved ? "reviewed" : ""}`;
  button.type = "button";
  button.textContent = approved ? "Approved" : "Approve";
  button.title = "Update PDPL obligation review status for this personal data item";
  button.addEventListener("click", () => {
    updateRecord(record.id, { personalApproved: !approved }).catch((error) => toast(error.message));
  });
  return button;
}

function editControl(record, key) {
  const wrapper = document.createElement("div");
  wrapper.className = "inline-edit";
  let input;
  if (key === "confidentiality") {
    input = document.createElement("select");
    ["Public", "Confidential", "Secret", "Top Secret", "Pending"].forEach((value) => input.add(new Option(value, value)));
    input.value = record.confidentiality || "Pending";
  } else if (key === "personalData") {
    input = document.createElement("select");
    ["Yes", "No"].forEach((value) => input.add(new Option(value, value)));
    input.value = record.personalData || "No";
  } else if (key === "confReason" || key === "personalReason") {
    input = document.createElement("textarea");
    input.rows = 3;
    input.value = record[key] || "";
  } else if (key === "reviewStatus") {
    input = document.createElement("select");
    ["Unreviewed", "Pending Review", "In Review", "Approved", "Rejected"].forEach((value) => input.add(new Option(value, value)));
    input.value = record.reviewStatus || "Unreviewed";
  } else if (key === "needsReview" || key === "pushToClient") {
    input = document.createElement("select");
    [["Yes", "true"], ["No", "false"]].forEach(([label, value]) => input.add(new Option(label, value)));
    input.value = record[key] ? "true" : "false";
  } else {
    input = document.createElement("input");
    input.value = record[key] || "";
  }

  input.addEventListener("change", () => {
    const raw = input.value;
    const value = key === "needsReview" || key === "pushToClient" ? raw === "true" : raw;
    updateRecord(record.id, { [key]: value }).catch((error) => toast(error.message));
  });
  wrapper.append(input);
  return wrapper;
}

async function updateRecord(recordId, updates) {
  try {
    const result = await api(`/api/records/${recordId}`, { method: "PUT", body: JSON.stringify(updates) });
    // Update in-memory caches so UI reflects the change immediately.
    state.records = state.records.map((r) => (r.id === recordId ? result.record : r));
    state.personalRecords = state.personalRecords.map((r) => (r.id === recordId ? result.record : r));

    // Re-render whichever table is visible so the user sees the update.
    if (state.activeTab === "system-data") renderRecordsTable();
    if (state.activeTab === "personal-data") {
      renderPersonalDataTable();
      // Reload to ensure counts and approval state are accurate.
      await loadPersonalData();
    }
    if (state.activeTab === "pdpl") await loadPdpl();

    // Always refresh overview charts regardless of current tab so numbers are
    // accurate when the user switches back.
    renderOverview(result.summary);
  } catch (error) {
    toast(error.message || "Failed to save changes.");
  }
}

async function approvePersonalOnPage() {
  if (!state.currentSystem) return;
  const recordIds = state.personalRecords
    .filter((record) => record.personalReviewStatus !== "Approved")
    .map((record) => record.id);
  if (!recordIds.length) {
    toast("All personal-data records on this page are already approved.");
    return;
  }
  const result = await api(`/api/systems/${state.currentSystem.id}/personal-approvals`, {
    method: "POST",
    body: JSON.stringify({ recordIds }),
  });
  state.personalRecords = result.records;
  renderPersonalDataTable();
  renderOverview(result.summary);
  toast(`Approved ${recordIds.length} personal-data record${recordIds.length === 1 ? "" : "s"}.`);
  if (state.activeTab === "pdpl") await loadPdpl();
}

function sortBy(key) {
  if (state.sortBy === key) {
    state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
  } else {
    state.sortBy = key;
    state.sortDir = "asc";
  }
  loadRecords();
}

function sortLabel(label, key) {
  if (state.sortBy !== key) return label;
  return `${label} ${state.sortDir === "asc" ? "ASC" : "DESC"}`;
}

/* SSE classification gives demos a real-time feel. The server streams row updates,
 * progress, and summary changes as each batch completes.
 */
async function startClassification(mode) {
  if (!state.currentSystem) return;
  closeClassifierStream();
  const body = {
    mode,
    page: state.page,
    pageSize: Math.min(50, state.pageSize),
    search: mode === "all" ? "" : state.search,
    sortBy: state.sortBy,
    sortDir: state.sortDir,
  };
  const fileName = state.currentSystem.lastUploadFileName || "No file uploaded";
  qs(".classification-panel").classList.add("is-running");
  setClassificationStatus(mode === "all" ? `Classifying all rows in ${fileName}` : `Classifying current page in ${fileName}`, 0);
  let jobResult;
  try {
    jobResult = await api(`/api/systems/${state.currentSystem.id}/classification-jobs`, {
      method: "POST",
      body: JSON.stringify(body),
    });
  } catch (error) {
    toast(error.message);
    setClassificationStatus("Classification could not start", 0);
    qs(".classification-panel").classList.remove("is-running");
    return;
  }

  state.classifierJobId = jobResult.job.id;
  state.classifierMode = jobResult.job.mode;
  state.classifierStream = new EventSource(
    `/api/systems/${state.currentSystem.id}/classification-jobs/${jobResult.job.id}/stream`
  );

  state.classifierStream.addEventListener("start", (event) => {
    const data = JSON.parse(event.data);
    updateFileNameLabels(fileName);
    handleClassificationJobUpdate(data, { renderTable: false });
  });

  state.classifierStream.addEventListener("row", async (event) => {
    const data = JSON.parse(event.data);
    await handleClassificationJobUpdate(data);
  });

  state.classifierStream.addEventListener("progress", async (event) => {
    const data = JSON.parse(event.data);
    await handleClassificationJobUpdate(data);
  });

  state.classifierStream.addEventListener("warning", (event) => {
    const data = JSON.parse(event.data);
    if (data.code === "OPENAI_NO_CREDITS") {
      toast(data.message || "No OpenAI API credits are available.", {
        variant: "danger",
        persistent: true,
      });
      return;
    }
    toast(data.message || "Classification continued with fallback rules.");
  });

  state.classifierStream.addEventListener("heartbeat", () => {
    // Keeps long classification runs alive behind proxies without changing UI state.
  });

  state.classifierStream.addEventListener("done", async (event) => {
    const data = JSON.parse(event.data);
    const job = data.job || {};
    setClassificationStatus(`Complete: ${job.total || 0} records`, 100);
    renderOverview(data.summary);
    closeClassifierStream();
    qs(".classification-panel").classList.remove("is-running");
    await reloadCurrentSystem();
    await loadRecords();
  });

  state.classifierStream.addEventListener("classification-error", async (event) => {
    const data = event.data ? JSON.parse(event.data) : null;
    const message = data?.job?.errorMessage || data?.error || "Classification failed";
    toast(message);
    setClassificationStatus("Classification stopped", 0);
    qs(".classification-panel").classList.remove("is-running");
    closeClassifierStream();
    await reloadCurrentSystem().catch(() => null);
    await loadRecords().catch(() => null);
  });

  state.classifierStream.addEventListener("error", async () => {
    toast("Classification connection interrupted. Saved rows were kept.");
    setClassificationStatus("Connection interrupted; saved rows kept", qs("#classificationProgress").value);
    qs(".classification-panel").classList.remove("is-running");
    closeClassifierStream();
    await reloadCurrentSystem().catch(() => null);
    await loadRecords().catch(() => null);
  });
}

async function handleClassificationJobUpdate(data, options = {}) {
  const job = data.job || {};
  const record = data.record || null;
  const total = Number(job.total || 0);
  const processed = Number(job.processed || 0);
  const percentage = Number(job.percentage || clampPercent((processed / Math.max(1, total)) * 100));
  const currentPage = Number(job.currentPage || state.page || 1);
  const pageLabel = state.classifierMode === "all" ? ` - page ${currentPage}` : "";
  setClassificationStatus(`Processed ${processed} of ${total}${pageLabel}`, percentage);
  if (data.summary) renderOverview(data.summary);

  if (state.classifierMode === "all" && currentPage && currentPage !== state.page) {
    state.page = currentPage;
    await loadRecords();
  }

  if (record) {
    state.records = state.records.map((item) => (item.id === record.id ? record : item));
    state.liveUpdatedRecordId = record.id;
  }

  if (options.renderTable !== false) renderRecordsTable();
  if (record) {
    setTimeout(() => {
      if (state.liveUpdatedRecordId === record.id) {
        state.liveUpdatedRecordId = null;
        renderRecordsTable();
      }
    }, 1200);
  }
}

function closeClassifierStream() {
  if (state.classifierStream) {
    state.classifierStream.close();
    state.classifierStream = null;
  }
  state.classifierJobId = null;
  state.classifierMode = null;
  const panel = qs(".classification-panel");
  if (panel) panel.classList.remove("is-running");
}

function closeActionsMenu() {
  const menu = qs(".actions-menu");
  if (menu) menu.removeAttribute("open");
}

function setClassificationStatus(text, percentage) {
  const statusEl = qs("#classificationStatus");
  // The element starts hidden in the HTML; reveal it the first time it is set
  // so the user can see classification progress and completion messages.
  statusEl.classList.remove("hidden");
  statusEl.textContent = text;
  const value = Number(percentage || 0);
  qs("#classificationProgress").value = value;
  qs("#classificationPercentLabel").textContent = `${Math.round(value)}% classified`;
  qs("#pendingPercentLabel").textContent = `${Math.max(0, 100 - Math.round(value))}% pending`;
}

function updateClassificationMetrics(summary) {
  const classified = Number(summary?.classificationProgress || 0);
  const pending = summary?.totalRecords ? Math.max(0, 100 - classified) : 0;
  qs("#classificationProgress").value = classified;
  qs("#classificationPercentLabel").textContent = `${classified}% classified`;
  qs("#pendingPercentLabel").textContent = `${pending}% pending`;
}

async function loadContext() {
  const result = await api(`/api/systems/${state.currentSystem.id}/context`);
  const container = qs("#contextCards");
  container.replaceChildren();
  if (!result.points.length) {
    container.append(emptyCard("No context points yet.", "Add system context to improve classification relevance."));
    return;
  }
  for (const point of result.points) {
    const card = document.createElement("article");
    const header = document.createElement("header");
    const title = document.createElement("strong");
    title.textContent = point.tag;
    const remove = document.createElement("button");
    remove.className = "btn btn-ghost";
    remove.type = "button";
    remove.textContent = "Delete";
    remove.addEventListener("click", async () => {
      await api(`/api/context/${point.id}`, { method: "DELETE" });
      await loadContext();
    });
    const content = document.createElement("p");
    content.textContent = point.content;
    header.append(title, remove);
    card.append(header, content);
    container.append(card);
  }
}

async function loadPdpl() {
  const result = await api(`/api/systems/${state.currentSystem.id}/pdpl`);
  qs("#pdplCompliant").textContent = result.status.compliant;
  qs("#pdplNeedsReview").textContent = result.status.needsReview;
  qs("#pdplNonCompliant").textContent = result.status.nonCompliant;
  qs("#pdplForm").governanceNotes.value = result.notes.governanceNotes || "";
  qs("#pdplForm").dataSubjectRightsCoverage.value = result.notes.dataSubjectRightsCoverage || "Needs Review";
  qs("#pdplForm").consentTrackingStatus.value = result.notes.consentTrackingStatus || "Needs Review";
  qs("#pdplForm").crossBorderTransferFlags.value = result.notes.crossBorderTransferFlags || "No Flags Recorded";

  const body = qs("#pdplObligationsBody");
  body.replaceChildren();
  for (const item of result.obligations) {
    const tr = document.createElement("tr");
    [item.tableName, item.columnName, item.personalDataType, item.policyRecommendation, item.reviewStatus].forEach((value, index) => {
      const td = document.createElement("td");
      if (index === 4) {
        td.append(badge(value || "Needs Review", value === "Approved" ? "badge-yes" : "badge-pending"));
      } else {
        td.textContent = value || "";
      }
      tr.append(td);
    });
    body.append(tr);
  }
}

async function loadLinks() {
  const result = await api(`/api/systems/${state.currentSystem.id}/links`);
  const container = qs("#systemLinksCards");
  container.replaceChildren();
  if (!result.links.length) {
    container.append(emptyCard("No links stored yet.", "Add documentation, API, or governance policy references."));
    return;
  }
  for (const link of result.links) {
    const card = document.createElement("article");
    const header = document.createElement("header");
    const title = document.createElement("strong");
    title.textContent = link.title;
    const actions = document.createElement("div");
    const open = document.createElement("a");
    open.className = "btn btn-secondary";
    open.href = link.url;
    open.target = "_blank";
    open.rel = "noreferrer";
    open.textContent = "Open";
    const remove = document.createElement("button");
    remove.className = "btn btn-ghost";
    remove.type = "button";
    remove.textContent = "Delete";
    remove.addEventListener("click", async () => {
      await api(`/api/links/${link.id}`, { method: "DELETE" });
      await loadLinks();
    });
    const category = badge(link.category, "badge-pending");
    const description = document.createElement("p");
    description.textContent = link.description || link.url;
    actions.append(open, remove);
    header.append(title, actions);
    card.append(header, category, description);
    container.append(card);
  }
}

function emptyCard(title, description) {
  const card = document.createElement("article");
  const strong = document.createElement("strong");
  strong.textContent = title;
  const p = document.createElement("p");
  p.textContent = description;
  card.append(strong, p);
  return card;
}

function badge(text, className) {
  const span = document.createElement("span");
  span.className = `badge ${className}`;
  span.textContent = text;
  return span;
}

function confidentialityClass(value) {
  if (value === "Public") return "badge-public";
  if (value === "Confidential") return "badge-confidential";
  if (value === "Secret") return "badge-secret";
  if (value === "Top Secret") return "badge-top-secret";
  return "badge-pending";
}

function formJson(form) {
  const data = new FormData(form);
  const result = {};
  for (const [key, value] of data.entries()) {
    result[key] = value;
  }
  qsa("input[type='checkbox']", form).forEach((input) => {
    result[input.name] = input.checked;
  });
  return result;
}

function wordCount(value) {
  return String(value || "").trim().split(/\s+/).filter(Boolean).length;
}

function debounce(fn, wait) {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}

function toast(message, options = {}) {
  const node = qs("#toast");
  const persistent = Boolean(options.persistent);
  if (node.dataset.locked === "true" && !persistent) return;
  node.className = "toast";
  node.replaceChildren();
  node.dataset.locked = persistent ? "true" : "false";
  node.setAttribute("role", options.variant === "danger" ? "alert" : "status");
  if (options.variant === "danger") node.classList.add("toast-danger");
  if (persistent) node.classList.add("toast-persistent");

  const text = document.createElement("span");
  text.textContent = message;
  node.append(text);

  if (persistent) {
    const close = document.createElement("button");
    close.className = "toast-close";
    close.type = "button";
    close.setAttribute("aria-label", "Dismiss notification");
    close.textContent = "x";
    close.addEventListener("click", () => {
      node.classList.remove("show");
      node.dataset.locked = "false";
      clearTimeout(node._timer);
    });
    node.append(close);
  }

  node.classList.add("show");
  clearTimeout(node._timer);
  if (!persistent) {
    node._timer = setTimeout(() => node.classList.remove("show"), 4200);
  }
}
