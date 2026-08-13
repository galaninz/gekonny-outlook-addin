/* Gekonny Subject Builder — v2 (project → folder → existing item)
   New item:        [Type] Address - description [CODE]
   Existing thread: [Type] Address - item name [CODE] {Ref} */

var CONFIG = {
  PROJECTS_ENDPOINT: "https://defaultd8bc567963cc4849af903e6e3f8795.cc.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a267216360ff4b788436407b67580369/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=gX1bour4ERee8isgpJsthBwnI1Va3WLwcWB33tkjSB4",
  ITEMS_ENDPOINT: "PASTE_ITEMS_FLOW_URL_HERE",
  INTAKE_ADDRESS: "build@gekonny.com"
};

var TYPES = [["Drawing","01 Drawings"],["Specification","02 Specifications"],["Submittal","03 Submittals"],["RFI","04 RFIs"],["Schedule","05 Schedule"],["Takeoff","06 Takeoff"],["Meeting Minutes","07 Meeting Minutes"],["Photo","08 Photos"],["Permit","09 Permits & Violations"],["Report","10 Reports & Punchlists"],["Insurance","11 Insurance"],["Agreement","12 Agreements & Contracts"],["Lien Waiver","13 Lien Waivers"],["CO","15 Change Orders"],["PO","16 Purchase Orders"],["Warranty","17 Warranty"],["Requisition","18 Requisitions"],["Team Doc","19 Team Documents"]];

var state = { projects: [], selectedProject: null, newBidType: "GC", items: [], selectedItem: null, itemsKey: "" };

Office.onReady(function (info) { if (info.host === Office.HostType.Outlook) { initUI(); loadProjects(); } });

function initUI() {
  byId("tabExisting").addEventListener("click", function () { switchTab("existing"); });
  byId("tabNew").addEventListener("click", function () { switchTab("new"); });
  byId("projectSearch").addEventListener("input", onProjectSearch);
  byId("projectSearch").addEventListener("focus", onProjectSearch);
  byId("projectClear").addEventListener("click", clearProject);
  byId("typeSelect").addEventListener("change", onTypeChange);
  byId("descInput").addEventListener("input", renderExistingPreview);
  byId("applyExisting").addEventListener("click", applyExisting);
  byId("modeNew").addEventListener("click", function () { setMode("new"); });
  byId("modeThread").addEventListener("click", function () { setMode("thread"); });
  byId("itemSearch").addEventListener("input", renderItemList);
  byId("itemReload").addEventListener("click", function () { loadItems(true); });
  byId("segGC").addEventListener("click", function () { setBidType("GC"); });
  byId("segSUB").addEventListener("click", function () { setBidType("SUB"); });
  byId("bidName").addEventListener("input", renderNewPreview);
  byId("applyNew").addEventListener("click", applyNew);
  byId("intakeAddr").textContent = CONFIG.INTAKE_ADDRESS;
  document.addEventListener("click", function (e) {
    if (!e.target.closest("#projectSearch") && !e.target.closest("#projectList")) { byId("projectList").hidden = true; }
  });
  renderExistingPreview(); renderNewPreview();
}

function switchTab(which) {
  var ex = which === "existing";
  byId("tabExisting").classList.toggle("tab-active", ex);
  byId("tabNew").classList.toggle("tab-active", !ex);
  byId("panelExisting").hidden = !ex;
  byId("panelNew").hidden = ex;
}

function setMode(m) {
  var thread = m === "thread";
  byId("modeNew").classList.toggle("seg-active", !thread);
  byId("modeThread").classList.toggle("seg-active", thread);
  byId("blockNewItem").hidden = thread;
  byId("blockThread").hidden = !thread;
  state.selectedItem = null;
  if (thread) { loadItems(false); }
  renderExistingPreview();
}
function isThreadMode() { return byId("modeThread").classList.contains("seg-active"); }

function loadProjects() {
  setStatus("projectStatus", "Loading projects…", false);
  fetch(CONFIG.PROJECTS_ENDPOINT, { method: "GET" })
    .then(function (r) { if (!r.ok) { throw new Error("HTTP " + r.status); } return r.json(); })
    .then(function (data) {
      var list = Array.isArray(data) ? data : (data.value || []);
      state.projects = list.filter(function (p) { return p && p.code; }).sort(function (a, b) { return (a.name || "").localeCompare(b.name || ""); });
      setStatus("projectStatus", state.projects.length + " projects loaded.", false);
    })
    .catch(function (err) { setStatus("projectStatus", "Could not load projects (" + err.message + ").", true); });
}

function onProjectSearch() {
  var q = byId("projectSearch").value.trim().toLowerCase();
  var listEl = byId("projectList"); listEl.innerHTML = "";
  var matches = state.projects.filter(function (p) {
    return (p.code || "").toLowerCase().indexOf(q) > -1 || (p.name || "").toLowerCase().indexOf(q) > -1;
  }).slice(0, 30);
  if (!matches.length) {
    listEl.innerHTML = '<div class="project-item"><div class="empty">No match.</div></div>';
  } else {
    matches.forEach(function (p) {
      var row = document.createElement("div");
      row.className = "project-item";
      row.innerHTML = '<div style="font-weight:600">' + esc(p.name || p.code) + '</div><div style="opacity:.7;font-size:12px">' + esc(p.code) + '</div>';
      row.addEventListener("click", function () { selectProject(p); });
      listEl.appendChild(row);
    });
  }
  listEl.hidden = false;
}

function selectProject(p) {
  state.selectedProject = p;
  byId("projectSearch").value = ""; byId("projectList").hidden = true; byId("projectSearch").hidden = true;
  byId("projectChosenText").textContent = (p.name || p.code) + " — " + p.code;
  byId("projectChosen").hidden = false;
  state.selectedItem = null;
  if (isThreadMode()) { loadItems(false); }
  renderExistingPreview();
}

function clearProject() {
  state.selectedProject = null; state.selectedItem = null; state.items = [];
  byId("projectChosen").hidden = true; byId("projectSearch").hidden = false;
  byId("projectSearch").value = ""; byId("projectSearch").focus();
  renderItemList(); renderExistingPreview();
}

function onTypeChange() { state.selectedItem = null; if (isThreadMode()) { loadItems(false); } renderExistingPreview(); }

function loadItems(force) {
  var p = state.selectedProject, type = byId("typeSelect").value;
  if (!p) { setStatus("itemStatus", "Pick a project first.", false); state.items = []; renderItemList(); return; }
  if (CONFIG.ITEMS_ENDPOINT.indexOf("PASTE_") === 0) { setStatus("itemStatus", "Item lookup not configured (ITEMS_ENDPOINT).", true); state.items = []; renderItemList(); return; }
  var key = p.code + "|" + type;
  if (!force && key === state.itemsKey && state.items.length) { renderItemList(); return; }
  setStatus("itemStatus", "Loading items…", false);
  var url = CONFIG.ITEMS_ENDPOINT + (CONFIG.ITEMS_ENDPOINT.indexOf("?") > -1 ? "&" : "?") + "code=" + encodeURIComponent(p.code) + "&type=" + encodeURIComponent(type);
  fetch(url, { method: "GET" })
    .then(function (r) { if (!r.ok) { throw new Error("HTTP " + r.status); } return r.json(); })
    .then(function (data) {
      var list = Array.isArray(data) ? data : (data.value || data.items || []);
      state.items = list.filter(function (i) { return i && i.name; });
      state.itemsKey = key;
      setStatus("itemStatus", state.items.length + " items in " + folderLabel(type) + ".", false);
      renderItemList();
    })
    .catch(function (err) { state.items = []; setStatus("itemStatus", "Could not load items (" + err.message + ").", true); renderItemList(); });
}

function renderItemList() {
  var box = byId("itemList"); box.innerHTML = "";
  var q = byId("itemSearch").value.trim().toLowerCase();
  var rows = state.items.filter(function (i) { return !q || (i.name || "").toLowerCase().indexOf(q) > -1 || (i.ref || "").toLowerCase().indexOf(q) > -1; });
  if (!rows.length) { box.innerHTML = '<div class="project-item"><div class="empty">No items.</div></div>'; return; }
  rows.slice(0, 50).forEach(function (i) {
    var row = document.createElement("div");
    row.className = "project-item" + (state.selectedItem && state.selectedItem.id === i.id ? " item-active" : "");
    var meta = [];
    if (i.ref) { meta.push(i.ref); }
    if (i.status) { meta.push(i.status); }
    if (i.date) { meta.push(i.date); }
    row.innerHTML = '<div style="font-weight:600">' + esc(i.name) + '</div>' + (meta.length ? '<div style="opacity:.7;font-size:12px">' + esc(meta.join(" · ")) + '</div>' : '');
    row.addEventListener("click", function () { state.selectedItem = i; renderItemList(); renderExistingPreview(); });
    box.appendChild(row);
  });
}

function folderLabel(v) { for (var i = 0; i < TYPES.length; i++) { if (TYPES[i][0] === v) { return TYPES[i][1]; } } return v; }

function buildExistingSubject() {
  var p = state.selectedProject;
  if (!p) { return null; }
  var type = byId("typeSelect").value;
  if (isThreadMode()) {
    var it = state.selectedItem;
    if (!it) { return null; }
    var s = "[" + type + "] " + (p.name || p.code) + " - " + it.name + " [" + p.code + "]";
    if (it.ref) { s += " {" + it.ref + "}"; }
    return s;
  }
  var desc = byId("descInput").value.trim();
  var subj = "[" + type + "] " + (p.name || p.code);
  if (desc) { subj += " - " + desc; }
  return subj + " [" + p.code + "]";
}

function renderExistingPreview() {
  var s = buildExistingSubject();
  byId("previewExisting").textContent = s || "—";
  byId("applyExisting").disabled = !s;
}

function buildNewSubject() {
  var name = byId("bidName").value.trim();
  return "[RFP " + state.newBidType + "]" + (name ? " " + name : "");
}
function renderNewPreview() { byId("previewNew").textContent = buildNewSubject(); }
function setBidType(t) {
  state.newBidType = t;
  byId("segGC").classList.toggle("seg-active", t === "GC");
  byId("segSUB").classList.toggle("seg-active", t === "SUB");
  renderNewPreview();
}
function applyExisting() { var s = buildExistingSubject(); if (s) { setSubject(s); } }
function applyNew() { setSubject(buildNewSubject()); }
function setSubject(subject) {
  Office.context.mailbox.item.subject.setAsync(subject, function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded) { showToast("Subject applied ✓", "ok"); }
    else { showToast("Could not set subject: " + (res.error ? res.error.message : "unknown"), "err"); }
  });
}
function byId(id) { return document.getElementById(id); }
function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
function setStatus(id, msg, isError) { var el = byId(id); if (!el) { return; } el.textContent = msg; el.classList.toggle("error", !!isError); }
var toastTimer = null;
function showToast(msg, kind) {
  var t = byId("toast"); t.textContent = msg; t.className = "toast " + (kind || "ok"); t.hidden = false;
  if (toastTimer) { clearTimeout(toastTimer); }
  toastTimer = setTimeout(function () { t.hidden = true; }, 3200);
}
