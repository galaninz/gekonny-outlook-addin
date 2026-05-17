/* ============================================================
   Gekonny Subject Builder — task pane logic
   ============================================================ */

/* ----- CONFIG: change these two values if anything moves ----- */
var CONFIG = {
  // Helper-flow HTTP URL (Power Automate "GetActiveProjects"). Full URL incl. &sig=
  PROJECTS_ENDPOINT: "https://defaultd8bc567963cc4849af903e6e3f8795.cc.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a267216360ff4b788436407b67580369/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=gX1bour4ERee8isgpJsthBwnI1Va3WLwcWB33tkjSB4",
  // Where new bids are forwarded for intake
  INTAKE_ADDRESS: "build@gekonny.com"
};
/* ------------------------------------------------------------- */

var state = {
  projects: [],          // [{id, name, code, status}]
  selectedProject: null, // chosen project object
  newBidType: "GC"       // GC | SUB
};

/* Office is ready */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    initUI();
    loadProjects();
  }
});

/* ===================== UI WIRING ===================== */
function initUI() {
  // tabs
  byId("tabExisting").addEventListener("click", function () { switchTab("existing"); });
  byId("tabNew").addEventListener("click", function () { switchTab("new"); });

  // existing-project controls
  byId("projectSearch").addEventListener("input", onProjectSearch);
  byId("projectSearch").addEventListener("focus", onProjectSearch);
  byId("projectClear").addEventListener("click", clearProject);
  byId("typeSelect").addEventListener("change", renderExistingPreview);
  byId("descInput").addEventListener("input", renderExistingPreview);
  byId("applyExisting").addEventListener("click", applyExisting);

  // new-bid controls
  byId("segGC").addEventListener("click", function () { setBidType("GC"); });
  byId("segSUB").addEventListener("click", function () { setBidType("SUB"); });
  byId("bidName").addEventListener("input", renderNewPreview);
  byId("applyNew").addEventListener("click", applyNew);

  byId("intakeAddr").textContent = CONFIG.INTAKE_ADDRESS;

  // close project list when clicking elsewhere
  document.addEventListener("click", function (e) {
    if (!e.target.closest("#projectSearch") && !e.target.closest("#projectList")) {
      byId("projectList").hidden = true;
    }
  });

  renderExistingPreview();
  renderNewPreview();
}

function switchTab(which) {
  var existing = which === "existing";
  byId("tabExisting").classList.toggle("tab-active", existing);
  byId("tabNew").classList.toggle("tab-active", !existing);
  byId("panelExisting").hidden = !existing;
  byId("panelNew").hidden = existing;
}

/* ===================== LOAD PROJECTS ===================== */
function loadProjects() {
  setProjectStatus("Loading projects…", false);

  if (CONFIG.PROJECTS_ENDPOINT.indexOf("PASTE_") === 0) {
    setProjectStatus("Project list not configured yet.", true);
    return;
  }

  fetch(CONFIG.PROJECTS_ENDPOINT, { method: "GET" })
    .then(function (r) {
      if (!r.ok) { throw new Error("HTTP " + r.status); }
      return r.json();
    })
    .then(function (data) {
      // helper-flow returns an array of {id,name,code,status}
      var list = Array.isArray(data) ? data : (data.value || []);
      state.projects = list
        .filter(function (p) { return p && p.code; })
        .sort(function (a, b) { return (b.code || "").localeCompare(a.code || ""); });
      setProjectStatus(state.projects.length + " projects loaded.", false);
    })
    .catch(function (err) {
      setProjectStatus("Could not load projects (" + err.message + "). You can still type a code manually.", true);
    });
}

/* ===================== PROJECT SEARCH ===================== */
function onProjectSearch() {
  var q = byId("projectSearch").value.trim().toLowerCase();
  var listEl = byId("projectList");
  listEl.innerHTML = "";

  var matches = state.projects.filter(function (p) {
    return (p.code || "").toLowerCase().indexOf(q) > -1 ||
           (p.name || "").toLowerCase().indexOf(q) > -1;
  }).slice(0, 30);

  if (state.projects.length === 0) {
    addEmptyRow(listEl, "No project list available — type the code into the description if needed.");
  } else if (matches.length === 0) {
    addEmptyRow(listEl, "No match.");
  } else {
    matches.forEach(function (p) {
      var row = document.createElement("div");
      row.className = "project-item";
      row.innerHTML = '<div class="code">' + esc(p.code) + '</div>' +
                      '<div class="name">' + esc(p.name || "") + '</div>';
      row.addEventListener("click", function () { selectProject(p); });
      listEl.appendChild(row);
    });
  }
  listEl.hidden = false;
}

function addEmptyRow(listEl, text) {
  var row = document.createElement("div");
  row.className = "project-item";
  row.innerHTML = '<div class="empty">' + esc(text) + '</div>';
  listEl.appendChild(row);
}

function selectProject(p) {
  state.selectedProject = p;
  byId("projectSearch").value = "";
  byId("projectList").hidden = true;
  byId("projectSearch").hidden = true;
  byId("projectChosenText").textContent = p.code + "  —  " + (p.name || "");
  byId("projectChosen").hidden = false;
  renderExistingPreview();
}

function clearProject() {
  state.selectedProject = null;
  byId("projectChosen").hidden = true;
  byId("projectSearch").hidden = false;
  byId("projectSearch").value = "";
  byId("projectSearch").focus();
  renderExistingPreview();
}

/* ===================== PREVIEW BUILDERS ===================== */
function buildExistingSubject() {
  if (!state.selectedProject) { return null; }
  var code = state.selectedProject.code;
  var type = byId("typeSelect").value;
  var desc = byId("descInput").value.trim();
  var subject = "[" + code + "] [" + type + "]";
  if (desc) { subject += " " + desc; }
  return subject;
}

function renderExistingPreview() {
  var subject = buildExistingSubject();
  byId("previewExisting").textContent = subject || "—";
  byId("applyExisting").disabled = !subject;
}

function buildNewSubject() {
  var name = byId("bidName").value.trim();
  var subject = "RFP " + state.newBidType + " -";
  if (name) { subject += " " + name; }
  return subject;
}

function renderNewPreview() {
  byId("previewNew").textContent = buildNewSubject();
}

function setBidType(t) {
  state.newBidType = t;
  byId("segGC").classList.toggle("seg-active", t === "GC");
  byId("segSUB").classList.toggle("seg-active", t === "SUB");
  renderNewPreview();
}

/* ===================== APPLY TO SUBJECT ===================== */
function applyExisting() {
  var subject = buildExistingSubject();
  if (!subject) { return; }
  setSubject(subject);
}

function applyNew() {
  var subject = buildNewSubject();
  setSubject(subject);
}

function setSubject(subject) {
  Office.context.mailbox.item.subject.setAsync(subject, function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      showToast("Subject applied ✓", "ok");
    } else {
      showToast("Could not set subject: " + (res.error ? res.error.message : "unknown"), "err");
    }
  });
}

/* ===================== HELPERS ===================== */
function byId(id) { return document.getElementById(id); }

function esc(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
  });
}

function setProjectStatus(msg, isError) {
  var el = byId("projectStatus");
  el.textContent = msg;
  el.classList.toggle("error", !!isError);
}

var toastTimer = null;
function showToast(msg, kind) {
  var t = byId("toast");
  t.textContent = msg;
  t.className = "toast " + (kind || "ok");
  t.hidden = false;
  if (toastTimer) { clearTimeout(toastTimer); }
  toastTimer = setTimeout(function () { t.hidden = true; }, 3200);
}
