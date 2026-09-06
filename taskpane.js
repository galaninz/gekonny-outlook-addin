/* Gekonny Subject Builder — v3.0
   Works inside Outlook (Apply to subject), as a web page, and as an
   installed app on phone or desktop (Copy subject / Open in Mail).

   New item:        [Type] Address - description [CODE]
   Existing thread: [Type] Address - item name [CODE] {#itemId}
*/

var CONFIG = {
  PROJECTS_ENDPOINT: "https://defaultd8bc567963cc4849af903e6e3f8795.cc.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/a267216360ff4b788436407b67580369/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=gX1bour4ERee8isgpJsthBwnI1Va3WLwcWB33tkjSB4",
  ITEMS_ENDPOINT: "https://defaultd8bc567963cc4849af903e6e3f8795.cc.environment.api.powerplatform.com/powerautomate/automations/direct/cu/19/workflows/23f3a10557784d41bde6b691334b3180/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=3uMez77ZGfTDxmWojlyyOErK1fyMM1Zo-9lqtzXmmSU",
  INTAKE_ADDRESS: "build@gekonny.com"
};

var TYPES = [["Drawing","01 Drawings"],["Specification","02 Specifications"],["Submittal","03 Submittals"],["RFI","04 RFIs"],["Schedule","05 Schedule"],["Takeoff","06 Takeoff"],["Meeting Minutes","07 Meeting Minutes"],["Photo","08 Photos"],["Permit","09 Permits & Violations"],["Report","10 Reports & Punchlists"],["Insurance","11 Insurance"],["Agreement","12 Agreements & Contracts"],["Lien Waiver","13 Lien Waivers"],["CO","15 Change Orders"],["PO","16 Purchase Orders"],["Warranty","17 Warranty"],["Requisition","18 Requisitions"],["Team Doc","19 Team Documents"]];

var state = { projects: [], selectedProject: null, newBidType: "GC", items: [], selectedItem: null, itemsKey: "",
              meetProject: null, queue: [], archive: [] };

var IN_OUTLOOK = false;
var booted = false;

/* ---------- boot ---------------------------------------------------- */

function boot(inOutlook) {
  if (booted) { return; }
  booted = true;
  IN_OUTLOOK = !!inOutlook;
  try { initUI(); } catch (e) { showToast("UI error: " + e.message, "err"); }
  loadProjects();
}

var officeIsOutlook = false;

function markOutlook() {
  IN_OUTLOOK = true;
  setText("applyExisting", "Apply to subject");
  setText("applyNew", "Apply to subject");
  showMailButtons(false);
}

/* Installed as an app: the service worker keeps the panel opening instantly
   and working with no signal. Inside Outlook it is not registered. */
if ("serviceWorker" in navigator && location.protocol === "https:") {
  window.addEventListener("load", function () {
    navigator.serviceWorker.register("sw.js").catch(function () {});
  });
}

function showMailButtons(show) {
  var a = byId("mailRowExisting"), b = byId("mailRowNew");
  if (a) { a.hidden = !show; }
  if (b) { b.hidden = !show; }
}

/* Two deliberate choices, because the platforms differ.

   "Open in Outlook" uses Outlook's own ms-outlook:// scheme, so it lands in
   Outlook regardless of which mail app the phone treats as default. That
   matters on iPhone, where mailto: silently goes to Apple Mail and company
   mail is not set up there.

   "Other mail app" is a plain mailto:, so the phone decides — on Android
   that means the usual app chooser. */

function openInOutlook(subject, to) {
  if (!subject) { showToast("Pick a project first.", "err"); return; }

  var left = false;
  function markLeft() { left = true; }
  document.addEventListener("visibilitychange", markLeft);
  window.addEventListener("pagehide", markLeft);

  setTimeout(function () {
    document.removeEventListener("visibilitychange", markLeft);
    window.removeEventListener("pagehide", markLeft);
    if (!left) { showToast("Outlook did not open. Try Other mail app.", "err"); }
  }, 1200);

  var url = "ms-outlook://compose?subject=" + encodeURIComponent(subject);
  if (to) { url += "&to=" + encodeURIComponent(to); }
  navTo(url);
}

function openInMail(subject, to) {
  if (!subject) { showToast("Pick a project first.", "err"); return; }
  window.location.href = "mailto:" + encodeURIComponent(to || "") +
                         "?subject=" + encodeURIComponent(subject);
}

/* A synthetic link click survives iOS standalone mode, where assigning
   location.href for a custom scheme is sometimes ignored. */
function navTo(url) {
  var a = document.createElement("a");
  a.href = url;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(function () { if (a.parentNode) { a.parentNode.removeChild(a); } }, 0);
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function (info) {
    officeIsOutlook = !!(info && info.host === Office.HostType.Outlook && Office.context && Office.context.mailbox);
    if (officeIsOutlook && booted) { markOutlook(); }
  });
}

function start() { boot(officeIsOutlook); }
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", start);
} else {
  start();
}

/* ---------- UI wiring ------------------------------------------------ */

function on(id, evt, fn) {
  var el = byId(id);
  if (el) { el.addEventListener(evt, fn); }
  return el;
}

function initUI() {
  on("tabExisting", "click", function () { switchTab("existing"); });
  on("tabNew", "click", function () { switchTab("new"); });
  on("projectSearch", "input", onProjectSearch);
  on("projectSearch", "focus", onProjectSearch);
  on("projectClear", "click", clearProject);
  on("manualUse", "click", useManualProject);
  on("typeSelect", "change", onTypeChange);
  on("descInput", "input", renderExistingPreview);
  on("applyExisting", "click", applyExisting);
  on("copyExisting", "click", function () { copyText(buildExistingSubject()); });
  on("modeNew", "click", function () { setMode("new"); });
  on("modeThread", "click", function () { setMode("thread"); });
  on("itemSearch", "input", renderItemList);
  on("itemReload", "click", function () { loadItems(true); });
  on("segGC", "click", function () { setBidType("GC"); });
  on("segSUB", "click", function () { setBidType("SUB"); });
  on("bidName", "input", renderNewPreview);
  on("applyNew", "click", applyNew);
  on("copyNew", "click", function () { copyText(buildNewSubject()); });
  on("mailExisting", "click", function () { openInOutlook(buildExistingSubject()); });
  on("mailExistingAny", "click", function () { openInMail(buildExistingSubject()); });
  on("mailNew", "click", function () { openInOutlook(buildNewSubject()); });
  on("mailNewAny", "click", function () { openInMail(buildNewSubject()); });

  on("tabMeeting", "click", function () { switchTab("meeting"); });
  on("meetProjectSearch", "input", onMeetProjectSearch);
  on("meetProjectSearch", "focus", onMeetProjectSearch);
  on("meetProjectClear", "click", clearMeetProject);
  on("meetAdd", "click", addQueueRow);
  on("meetDesc", "keydown", function (e) { if (e.key === "Enter") { addQueueRow(); } });
  on("meetTo", "keydown", function (e) { if (e.key === "Enter") { addQueueRow(); } });
  on("backlogToggle", "click", toggleBacklog);
  fillMeetTypes();
  loadQueue();
  renderQueue();

  var intake = byId("intakeAddr");
  if (intake) { intake.textContent = CONFIG.INTAKE_ADDRESS; }

  if (!IN_OUTLOOK) {
    setText("applyExisting", "Copy subject");
    setText("applyNew", "Copy subject");
    showMailButtons(true);
  }

  document.addEventListener("click", function (e) {
    if (!e.target.closest("#projectSearch") && !e.target.closest("#projectList")) {
      var pl = byId("projectList");
      if (pl) { pl.hidden = true; }
    }
  });

  renderExistingPreview();
  renderNewPreview();
}

function switchTab(which) {
  var tabs = { existing: "Existing", "new": "New", meeting: "Meeting" };
  Object.keys(tabs).forEach(function (key) {
    var on = key === which;
    var t = byId("tab" + tabs[key]), p = byId("panel" + tabs[key]);
    if (t) { t.classList.toggle("tab-active", on); }
    if (p) { p.hidden = !on; }
  });
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
function isThreadMode() {
  var el = byId("modeThread");
  return !!el && el.classList.contains("seg-active");
}

/* ---------- projects ------------------------------------------------- */

function loadProjects() {
  setStatus("projectStatus", "Loading projects…", false);
  fetch(CONFIG.PROJECTS_ENDPOINT, { method: "GET" })
    .then(function (r) { if (!r.ok) { throw new Error("HTTP " + r.status); } return r.json(); })
    .then(function (data) {
      var list = Array.isArray(data) ? data : (data.value || data.projects || []);
      state.projects = list.filter(function (p) { return p && p.code; }).sort(function (a, b) {
        return (a.name || "").localeCompare(b.name || "");
      });
      if (!state.projects.length) { throw new Error("empty list"); }
      setStatus("projectStatus", state.projects.length + " projects loaded.", false);
      showManual(false);
    })
    .catch(function (err) {
      setStatus("projectStatus", "Could not load projects (" + err.message + ").", true);
      showManual(true);
    });
}

function showManual(show) {
  var box = byId("manualBox");
  if (box) { box.hidden = !show; }
}

function useManualProject() {
  var name = (byId("manualName").value || "").trim();
  var code = (byId("manualCode").value || "").trim().toUpperCase();
  if (!code) { showToast("Enter the project code.", "err"); return; }
  selectProject({ name: name || code, code: code });
}

function onProjectSearch() {
  var q = byId("projectSearch").value.trim().toLowerCase();
  var listEl = byId("projectList");
  listEl.innerHTML = "";
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
  byId("projectSearch").value = "";
  byId("projectList").hidden = true;
  byId("projectSearch").hidden = true;
  showManual(false);
  if (!state.projects.length) { setStatus("projectStatus", "", false); }
  byId("projectChosenText").textContent = (p.name || p.code) + " — " + p.code;
  byId("projectChosen").hidden = false;
  state.selectedItem = null;
  if (isThreadMode()) { loadItems(false); }
  renderExistingPreview();
}

function clearProject() {
  state.selectedProject = null;
  state.selectedItem = null;
  state.items = [];
  byId("projectChosen").hidden = true;
  byId("projectSearch").hidden = false;
  byId("projectSearch").value = "";
  byId("projectSearch").focus();
  if (!state.projects.length) { showManual(true); }
  renderItemList();
  renderExistingPreview();
}

function onTypeChange() {
  state.selectedItem = null;
  if (isThreadMode()) { loadItems(false); }
  renderExistingPreview();
}

/* ---------- existing items ------------------------------------------- */

function loadItems(force) {
  var p = state.selectedProject, type = byId("typeSelect").value;
  if (!p) { setStatus("itemStatus", "Pick a project first.", false); state.items = []; renderItemList(); return; }
  if (CONFIG.ITEMS_ENDPOINT.indexOf("PASTE_") === 0) {
    setStatus("itemStatus", "Item lookup not configured yet (ITEMS_ENDPOINT).", true);
    state.items = []; renderItemList(); return;
  }
  var key = p.code + "|" + type;
  if (!force && key === state.itemsKey && state.items.length) { renderItemList(); return; }
  setStatus("itemStatus", "Loading items…", false);
  var url = CONFIG.ITEMS_ENDPOINT + (CONFIG.ITEMS_ENDPOINT.indexOf("?") > -1 ? "&" : "?") +
            "code=" + encodeURIComponent(p.code) + "&type=" + encodeURIComponent(type);
  fetch(url, { method: "GET" })
    .then(function (r) { if (!r.ok) { throw new Error("HTTP " + r.status); } return r.json(); })
    .then(function (data) {
      var list = Array.isArray(data) ? data : (data.value || data.items || []);
      state.items = list.filter(function (i) { return i && i.name; });
      state.itemsKey = key;
      setStatus("itemStatus", state.items.length + " items in " + folderLabel(type) + ".", false);
      renderItemList();
    })
    .catch(function (err) {
      state.items = [];
      setStatus("itemStatus", "Could not load items (" + err.message + ").", true);
      renderItemList();
    });
}

function renderItemList() {
  var box = byId("itemList");
  if (!box) { return; }
  box.innerHTML = "";
  var q = (byId("itemSearch") ? byId("itemSearch").value : "").trim().toLowerCase();
  var rows = state.items.filter(function (i) {
    return !q || (i.name || "").toLowerCase().indexOf(q) > -1 || (i.ref || "").toLowerCase().indexOf(q) > -1;
  });
  if (!rows.length) { box.innerHTML = '<div class="project-item"><div class="empty">No items.</div></div>'; return; }
  rows.slice(0, 50).forEach(function (i) {
    var row = document.createElement("div");
    row.className = "project-item" + (state.selectedItem && state.selectedItem.id === i.id ? " item-active" : "");
    var meta = [];
    if (i.ref && i.ref.charAt(0) !== "#") { meta.push(i.ref); }
    if (i.status) { meta.push(i.status); }
    if (i.date) { meta.push(i.date); }
    row.innerHTML = '<div style="font-weight:600">' + esc(i.name) + '</div>' +
      (meta.length ? '<div style="opacity:.7;font-size:12px">' + esc(meta.join(" · ")) + '</div>' : '');
    row.addEventListener("click", function () { state.selectedItem = i; renderItemList(); renderExistingPreview(); });
    box.appendChild(row);
  });
}

function folderLabel(v) {
  for (var i = 0; i < TYPES.length; i++) { if (TYPES[i][0] === v) { return TYPES[i][1]; } }
  return v;
}

/* ---------- subject building ----------------------------------------- */

function buildExistingSubject() {
  var p = state.selectedProject;
  if (!p) { return null; }
  var type = byId("typeSelect").value;
  if (isThreadMode()) {
    var it = state.selectedItem;
    if (!it) { return null; }
    var s = "[" + type + "] " + (p.name || p.code) + " - " + it.name + " [" + p.code + "]";
    if (it.id) { s += " {#" + it.id + "}"; }
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
  var c = byId("copyExisting");
  if (c) { c.disabled = !s; }
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
  if (!IN_OUTLOOK || !Office.context.mailbox.item || !Office.context.mailbox.item.subject) {
    copyText(subject);
    return;
  }
  try {
    Office.context.mailbox.item.subject.setAsync(subject, function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) { showToast("Subject applied ✓", "ok"); }
      else { copyText(subject, "Could not set subject — copied instead."); }
    });
  } catch (e) {
    copyText(subject, "Could not set subject — copied instead.");
  }
}

/* ---------- clipboard ------------------------------------------------- */

function copyText(text, altMessage) {
  if (!text) { return; }
  var done = function () { showToast(altMessage || "Subject copied — paste it with Ctrl/Cmd+V ✓", "ok"); };
  var fail = function () { showToast("Copy blocked — select the preview text and copy it manually.", "err"); };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(done, function () { legacyCopy(text) ? done() : fail(); });
  } else {
    legacyCopy(text) ? done() : fail();
  }
}

function legacyCopy(text) {
  try {
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.setAttribute("readonly", "");
    ta.style.position = "fixed";
    ta.style.top = "-1000px";
    document.body.appendChild(ta);
    ta.select();
    var ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  } catch (e) { return false; }
}

/* ---------- helpers --------------------------------------------------- */

function byId(id) { return document.getElementById(id); }
function setText(id, t) { var el = byId(id); if (el) { el.textContent = t; } }
function esc(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
  });
}
function setStatus(id, msg, isError) {
  var el = byId(id);
  if (!el) { return; }
  el.textContent = msg;
  el.classList.toggle("error", !!isError);
}
var toastTimer = null;
function showToast(msg, kind) {
  var t = byId("toast");
  if (!t) { return; }
  t.textContent = msg;
  t.className = "toast " + (kind || "ok");
  t.hidden = false;
  if (toastTimer) { clearTimeout(toastTimer); }
  toastTimer = setTimeout(function () { t.hidden = true; }, 3600);
}

/* ---------- queue -----------------------------------------------------

   A to-do list of subjects you build during a meeting and work through
   afterwards. It has no access to mail: "Open" hands the subject to
   Outlook and you send it yourself, which is the whole point — nothing
   leaves this panel on its own.

   Rows can point at different projects; an office meeting covering several
   jobs is exactly the case this is for, so the list groups by project.

   Storage is this device's browser only. Deliberate: no server, no second
   copy of the truth to keep in sync. The cost is that the phone and the
   desktop keep separate lists, and clearing browser data clears the queue.
   --------------------------------------------------------------------- */

var Q_KEY = "gk_queue_v1";
var Q_ARCHIVE_KEY = "gk_queue_archive_v1";
var ARCHIVE_DAYS = 30;

function fillMeetTypes() {
  var sel = byId("meetType");
  if (!sel || sel.options.length) { return; }
  TYPES.forEach(function (t) {
    var o = document.createElement("option");
    o.value = t[0];
    o.textContent = t[1];
    if (t[0] === "RFI") { o.selected = true; }
    sel.appendChild(o);
  });
}

function onMeetProjectSearch() {
  var q = byId("meetProjectSearch").value.trim().toLowerCase();
  var listEl = byId("meetProjectList");
  listEl.innerHTML = "";
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
      row.addEventListener("click", function () { selectMeetProject(p); });
      listEl.appendChild(row);
    });
  }
  listEl.hidden = false;
}

function selectMeetProject(p) {
  state.meetProject = p;
  byId("meetProjectSearch").value = "";
  byId("meetProjectSearch").hidden = true;
  byId("meetProjectList").hidden = true;
  byId("meetProjectChosenText").textContent = (p.name || p.code) + " — " + p.code;
  byId("meetProjectChosen").hidden = false;
  var d = byId("meetDesc");
  if (d) { d.focus(); }
}

function clearMeetProject() {
  state.meetProject = null;
  byId("meetProjectChosen").hidden = true;
  byId("meetProjectSearch").hidden = false;
  byId("meetProjectSearch").value = "";
  byId("meetProjectSearch").focus();
}


/* Returns the first address that does not look like an address, or "". */
function badEmails(list) {
  var parts = String(list || "").split(/[,;]/);
  for (var i = 0; i < parts.length; i++) {
    var a = parts[i].trim();
    if (!a) { continue; }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a)) { return a; }
  }
  return "";
}

function uid() {
  return String(Date.now()) + "-" + Math.random().toString(36).slice(2, 8);
}

/* Storage can throw outright — private windows, and some managed browsers
   block it entirely. Never let that take the panel down with it. */
function qRead(key) {
  try {
    var raw = localStorage.getItem(key);
    var rows = raw ? JSON.parse(raw) : [];
    return Array.isArray(rows) ? rows : [];
  } catch (e) { return []; }
}

function qWrite(key, rows) {
  try {
    localStorage.setItem(key, JSON.stringify(rows));
    return true;
  } catch (e) {
    setStatus("meetStatus", "This browser will not let the panel save the queue, so it will be gone when you close it.", true);
    return false;
  }
}

function loadQueue() {
  state.queue = qRead(Q_KEY);
  state.archive = qRead(Q_ARCHIVE_KEY);
  pruneArchive();
}

/* The backlog answers "what did I just cross off" — a month is generous for
   that and keeps the list from growing without limit. */
function pruneArchive() {
  var cutoff = Date.now() - ARCHIVE_DAYS * 86400000;
  var kept = state.archive.filter(function (r) {
    var t = Date.parse(r.at || "");
    return isNaN(t) ? true : t >= cutoff;
  });
  if (kept.length !== state.archive.length) {
    state.archive = kept;
    qWrite(Q_ARCHIVE_KEY, state.archive);
  }
}

function saveQueue() { qWrite(Q_KEY, state.queue); }
function saveArchive() { qWrite(Q_ARCHIVE_KEY, state.archive); }

/* ---------- adding ---------------------------------------------------- */

function addQueueRow() {
  var p = state.meetProject;
  if (!p) { showToast("Pick a project for this task.", "err"); return; }

  var desc = byId("meetDesc").value.trim();
  if (!desc) { showToast("Describe the task.", "err"); return; }

  var to = byId("meetTo").value.trim();
  var bad = badEmails(to);
  if (bad) { showToast("Check this address: " + bad, "err"); return; }

  var type = byId("meetType").value;
  state.queue.push({
    id: uid(),
    projectName: p.name || p.code,
    code: p.code,
    type: type,
    description: desc,
    to: to,
    subject: "[" + type + "] " + (p.name || p.code) + " - " + desc + " [" + p.code + "]",
    at: new Date().toISOString()
  });
  saveQueue();

  byId("meetDesc").value = "";
  byId("meetTo").value = "";
  setStatus("meetStatus", "", false);
  renderQueue();
  byId("meetDesc").focus();
}

/* ---------- acting on a row ------------------------------------------- */

function qFind(id) {
  for (var i = 0; i < state.queue.length; i++) {
    if (state.queue[i].id === id) { return i; }
  }
  return -1;
}

function queueOpen(id) {
  var i = qFind(id);
  if (i < 0) { return; }
  var row = state.queue[i];

  if (IN_OUTLOOK) {
    try {
      var mb = Office.context.mailbox;
      if (mb && mb.displayNewMessageForm) {
        mb.displayNewMessageForm({
          toRecipients: row.to ? row.to.split(/[,;]/).map(function (s) { return s.trim(); }).filter(Boolean) : [],
          subject: row.subject
        });
        return;
      }
    } catch (e) { /* fall through */ }
    setSubject(row.subject);
    return;
  }

  openInOutlook(row.subject, row.to);
}

/* Crossed off for good: it leaves the queue and lands in the backlog, so
   there is always a record of what was just cleared. */
function queueArchive(id, reason) {
  var i = qFind(id);
  if (i < 0) { return; }
  var row = state.queue.splice(i, 1)[0];
  row.reason = reason;
  row.at = new Date().toISOString();
  state.archive.unshift(row);
  saveQueue();
  saveArchive();
  renderQueue();
  showToast(reason === "sent" ? "Crossed off ✓" : "Removed — it is in the backlog.", "ok");
}

function queueRestore(id) {
  for (var i = 0; i < state.archive.length; i++) {
    if (state.archive[i].id === id) {
      var row = state.archive.splice(i, 1)[0];
      delete row.reason;
      state.queue.push(row);
      saveQueue();
      saveArchive();
      renderQueue();
      showToast("Back in the queue.", "ok");
      return;
    }
  }
}

/* ---------- rendering -------------------------------------------------- */

function qButton(label, cls, fn) {
  var b = document.createElement("button");
  b.type = "button";
  b.className = "q-btn" + (cls ? " " + cls : "");
  b.textContent = label;
  b.addEventListener("click", fn);
  return b;
}

function qRowShell(row, archived) {
  var el = document.createElement("div");
  el.className = "meet-row" + (archived ? " q-archived" : "");
  var main = document.createElement("div");
  main.className = "meet-main";
  var subj = document.createElement("div");
  subj.className = "meet-subject";
  subj.textContent = row.subject;
  main.appendChild(subj);
  if (row.to) {
    var to = document.createElement("div");
    to.className = "meet-to";
    to.textContent = "→ " + row.to;
    main.appendChild(to);
  }
  el.appendChild(main);
  return { el: el, main: main };
}

function renderQueue() {
  var box = byId("meetList");
  if (!box) { return; }
  box.innerHTML = "";

  var n = state.queue.length;
  setText("meetCount", n ? "(" + n + ")" : "");

  if (!n) {
    box.innerHTML = '<div class="meet-empty">Nothing queued. Add tasks one by one above.</div>';
  } else {
    var lastCode = null;
    state.queue.forEach(function (row) {
      if (row.code !== lastCode) {
        lastCode = row.code;
        var h = document.createElement("div");
        h.className = "q-group";
        h.textContent = row.projectName + " · " + row.code;
        box.appendChild(h);
      }

      var shell = qRowShell(row, false);
      var acts = document.createElement("div");
      acts.className = "q-actions";
      acts.appendChild(qButton("Open", "", function () { queueOpen(row.id); }));
      acts.appendChild(qButton("Sent", "q-btn-done", function () { queueArchive(row.id, "sent"); }));
      acts.appendChild(qButton("×", "q-btn-drop", function () { queueArchive(row.id, "dropped"); }));
      shell.el.appendChild(acts);
      box.appendChild(shell.el);
    });
  }

  renderBacklog();
}

function renderBacklog() {
  var card = byId("backlogCard"), box = byId("backlogList");
  if (!card || !box) { return; }

  var n = state.archive.length;
  card.hidden = !n;
  setText("backlogLabel", "Backlog (" + n + ")");
  if (!n) { return; }

  box.innerHTML = "";
  state.archive.forEach(function (row) {
    var shell = qRowShell(row, true);
    var stamp = document.createElement("div");
    stamp.className = "q-stamp";
    stamp.textContent = (row.reason === "sent" ? "Sent" : "Removed") + " · " + shortDate(row.at);
    shell.main.appendChild(stamp);

    var acts = document.createElement("div");
    acts.className = "q-actions";
    acts.appendChild(qButton("Put back", "", function () { queueRestore(row.id); }));
    shell.el.appendChild(acts);
    box.appendChild(shell.el);
  });
}

function toggleBacklog() {
  var box = byId("backlogList"), chev = byId("backlogChev"), hint = byId("backlogHint");
  if (!box) { return; }
  var open = box.hidden;
  box.hidden = !open;
  if (hint) { hint.hidden = !open; }
  if (chev) { chev.classList.toggle("open", open); }
}

function shortDate(iso) {
  var t = Date.parse(iso || "");
  if (isNaN(t)) { return ""; }
  var d = new Date(t);
  try {
    return d.toLocaleDateString(undefined, { month: "short", day: "numeric" });
  } catch (e) {
    return d.getMonth() + 1 + "/" + d.getDate();
  }
}
