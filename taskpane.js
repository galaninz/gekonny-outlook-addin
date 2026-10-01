/* Gekonny Subject Builder — v3.3
   Works inside Outlook (Apply to subject), as a web page, and as an
   installed app on phone or desktop (Copy subject / Open in Mail).

   New item:        [Type] Address - description [CODE]
   Existing thread: [Type] Address - item name [CODE] {#itemId}

   v3.3: opened on an email you are READING, the panel shows "Add to Monday"
   instead — the email (with its history) goes into Monday through the flow
   "Add Email to Monday" (CONFIG.LINK_ENDPOINT).
*/

var CONFIG = {
  PROJECTS_ENDPOINT: "https://defaultd8bc567963cc4849af903e6e3f8795.cc.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/a267216360ff4b788436407b67580369/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=gX1bour4ERee8isgpJsthBwnI1Va3WLwcWB33tkjSB4",
  /* Flow "Add Email to Monday (Subject Builder)" — HTTP POST URL */
  LINK_ENDPOINT: "https://defaultd8bc567963cc4849af903e6e3f8795.cc.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/29/workflows/73f8e9ae43204772ba20fac0000f58ba/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=s2U5STeEQcGg8_u0yDmJWNlp_fS1AaeJ23LumLy92fQ",
  ITEMS_ENDPOINT: "https://defaultd8bc567963cc4849af903e6e3f8795.cc.environment.api.powerplatform.com/powerautomate/automations/direct/cu/19/workflows/23f3a10557784d41bde6b691334b3180/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=3uMez77ZGfTDxmWojlyyOErK1fyMM1Zo-9lqtzXmmSU"
};

/* The first value of each pair is the tag that goes into the subject, and
   the mail intake (PC v2) keys on it to pick the folder — so a tag added
   here that PC v2 does not know files nowhere. Keep the two in step.
   The second value is only the label shown in the dropdown, which is why
   it reads "21 Scope of Work" while the folder is 21_ScopeOfWork.
   RFP -> 14_BidsAndProposals: on a tender (RFP / Bid Pipeline) PC v2 turns
   an [RFP] email into Bid Invitation rows, one per sub in To. */
var TYPES = [["Drawing","01 Drawings"],["Specification","02 Specifications"],["Submittal","03 Submittals"],["RFI","04 RFIs"],["Schedule","05 Schedule"],["Takeoff","06 Takeoff"],["Meeting Minutes","07 Meeting Minutes"],["Photo","08 Photos"],["Permit","09 Permits & Violations"],["Report","10 Reports & Punchlists"],["Insurance","11 Insurance"],["Agreement","12 Agreements & Contracts"],["Lien Waiver","13 Lien Waivers"],["RFP","14 Bids & Proposals"],["CO","15 Change Orders"],["PO","16 Purchase Orders"],["Warranty","17 Warranty"],["Requisition","18 Requisitions"],["Team Doc","19 Team Documents"],["Trash","20 Trash & Debris"],["Scope","21 Scope of Work"],["Invoice","22 Invoices"],["Inspection","23 Inspections"]];

/* Numbered documents: the mail intake (PC v2) numbers them in "Number #"
   and files each one into its own Dropbox folder, e.g. 15_CO/CO#1 - Subject.
   Same prefixes as PC v2 uses. */
var NUM_PREFIX = { "Submittal": "SUB", "RFI": "RFI", "CO": "CO", "PO": "PO", "Requisition": "REQ", "Invoice": "INV" };
function numLabel(type, num) {
  if (!num) { return ""; }
  return (NUM_PREFIX[type] || type) + "#" + num;
}

var state = { projects: [], selectedProject: null, items: [], selectedItem: null, itemsKey: "",
              meetProject: null, queue: [], archive: [], loadedFromQueue: null,
              readRecips: [], readBusy: false };

var IN_OUTLOOK = false;
var booted = false;
var READ_MODE = false;        /* v3.3: panel opened on an email being read */
var readMsg = null;
var readHandlerAdded = false;

/* ---------- boot ---------------------------------------------------- */

function boot(inOutlook) {
  if (booted) { return; }
  booted = true;
  IN_OUTLOOK = !!inOutlook;
  try { initUI(); } catch (e) { showToast("UI error: " + e.message, "err"); }
  if (IN_OUTLOOK && detectReadMode()) { enterReadMode(false); }
  loadProjects();
}

var officeIsOutlook = false;

function markOutlook() {
  IN_OUTLOOK = true;
  setText("applyExisting", "Apply to subject");
  showMailButtons(false);
  if (detectReadMode()) { enterReadMode(false); }
}

/* Installed as an app: the service worker keeps the panel opening instantly
   and working with no signal. Inside Outlook it is not registered. */
if ("serviceWorker" in navigator && location.protocol === "https:") {
  window.addEventListener("load", function () {
    navigator.serviceWorker.register("sw.js").catch(function () {});
  });
}

function showMailButtons(show) {
  var a = byId("mailRowExisting");
  if (a) { a.hidden = !show; }
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
  on("fromQueue", "change", onFromQueuePick);
  on("fromQueueDone", "click", crossOffLoaded);
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
  on("mailExisting", "click", function () { openInOutlook(buildExistingSubject(), loadedRecipient()); });
  on("mailExistingAny", "click", function () { openInMail(buildExistingSubject(), loadedRecipient()); });

  on("tabMeeting", "click", function () { switchTab("meeting"); });
  on("meetProjectSearch", "input", onMeetProjectSearch);
  on("meetProjectSearch", "focus", onMeetProjectSearch);
  on("meetProjectClear", "click", clearMeetProject);
  on("meetAdd", "click", addQueueRow);
  on("meetDesc", "keydown", function (e) { if (e.key === "Enter") { addQueueRow(); } });
  on("meetTo", "keydown", function (e) { if (e.key === "Enter") { addQueueRow(); } });
  on("backlogToggle", "click", toggleBacklog);
  fillTypes("typeSelect");
  fillTypes("meetType");
  loadQueue();
  renderQueue();

  if (!IN_OUTLOOK) {
    setText("applyExisting", "Copy subject");
    showMailButtons(true);
  }

  document.addEventListener("click", function (e) {
    if (!e.target.closest("#projectSearch") && !e.target.closest("#projectList")) {
      var pl = byId("projectList");
      if (pl) { pl.hidden = true; }
    }
  });

  renderExistingPreview();
}

function switchTab(which) {
  var tabs = { existing: "Existing", meeting: "Meeting" };
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
  if (READ_MODE) { renderReadRecipients(); }
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
  var type = byId("typeSelect") ? byId("typeSelect").value : "";
  var rows = state.items.filter(function (i) {
    return !q || (i.name || "").toLowerCase().indexOf(q) > -1 || (i.ref || "").toLowerCase().indexOf(q) > -1 ||
           numLabel(type, i.num).toLowerCase().indexOf(q) > -1;
  });
  if (!rows.length) { box.innerHTML = '<div class="project-item"><div class="empty">No items.</div></div>'; return; }
  rows.slice(0, 50).forEach(function (i) {
    var row = document.createElement("div");
    row.className = "project-item" + (state.selectedItem && state.selectedItem.id === i.id ? " item-active" : "");
    var meta = [];
    if (i.num) { meta.push(numLabel(type, i.num)); }
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
    var nl = numLabel(type, it.num);
    var s = "[" + type + "] " + (p.name || p.code) + " - " + (nl ? nl + " " : "") + it.name + " [" + p.code + "]";
    if (it.id) { s += " {#" + it.id + "}"; }
    return s;
  }
  var desc = byId("descInput").value.trim();
  var subj = "[" + type + "] " + (p.name || p.code);
  if (desc) { subj += " - " + desc; }
  return subj + " [" + p.code + "]";
}

function renderExistingPreview() {
  if (READ_MODE) {
    var r = buildReadPreview();
    byId("previewExisting").textContent = r || "—";
    byId("applyExisting").disabled = !r || state.readBusy;
    return;
  }
  var s = buildExistingSubject();
  byId("previewExisting").textContent = s || "—";
  byId("applyExisting").disabled = !s;
  var c = byId("copyExisting");
  if (c) { c.disabled = !s; }

  /* The moment the form stops matching the task that was loaded, the link is
     stale — drop it, so "Cross off" can never strike the wrong row. */
  if (state.loadedFromQueue && s !== state.loadedFromQueue.subject) {
    unloadFromQueue();
  }
}

function applyExisting() {
  if (READ_MODE) { addToMonday(); return; }
  var s = buildExistingSubject(); if (s) { setSubject(s); }
}

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

/* ---------- read mode: Add to Monday (v3.3) ---------------------------

   The panel opened on an email you are READING (it cannot change that
   email's subject). Pick the project and type and press "Add to Monday":
   the flow "Add Email to Monday" puts the email, with its history, into
   Monday — for RFP one Bid Invitation per chosen outside recipient, for any
   other type a new row in the type's group or the existing row you pick.
   It also stores the conversation, so replies land on that row and the
   Follow-up from Monday answers in the same thread. */

/* READ_MODE, readMsg, readHandlerAdded are declared at the top (boot can run first). */

function detectReadMode() {
  try {
    var it = Office.context.mailbox.item;
    return !!it && typeof it.subject === "string";
  } catch (e) { return false; }
}

function cleanSubject(s) {
  return String(s || "").replace(/^(\s*(re|fw|fwd|aw)\s*:\s*)+/i, "").trim();
}

function addrList(arr) {
  var out = [];
  (arr || []).forEach(function (a) {
    if (a && a.emailAddress) { out.push(String(a.emailAddress).toLowerCase().trim()); }
  });
  return out;
}

function collectReadMessage() {
  var it = Office.context.mailbox.item;
  var from = it.from || it.sender || {};
  var dt = it.dateTimeCreated ? new Date(it.dateTimeCreated) : null;
  return {
    subject: it.subject || "",
    from: String(from.emailAddress || "").toLowerCase().trim(),
    fromName: from.displayName || "",
    to: addrList(it.to),
    cc: addrList(it.cc),
    date: dt ? dt.toISOString() : "",
    dateText: dt ? dt.toLocaleString() : "",
    internetMessageId: it.internetMessageId || "",
    conversationId: it.conversationId || ""
  };
}

/* Outside addresses of this email: To and the sender are ticked, Cc is not. */
function readExternals() {
  var seen = {}, out = [];
  function add(e, ticked) {
    e = String(e || "").toLowerCase().trim();
    if (!e || /@gekonny\.com$/.test(e) || seen[e]) { return; }
    seen[e] = true;
    out.push({ email: e, checked: ticked });
  }
  readMsg.to.forEach(function (e) { add(e, true); });
  add(readMsg.from, true);
  readMsg.cc.forEach(function (e) { add(e, false); });
  return out;
}

function enterReadMode(fromItemChange) {
  READ_MODE = true;
  readMsg = collectReadMessage();
  state.readRecips = readExternals();
  var panel = byId("panelExisting");
  var box = byId("readBox");
  if (!box) {
    box = document.createElement("div");
    box.id = "readBox";
    box.className = "card";
    panel.insertBefore(box, panel.firstChild);
    var rc = document.createElement("div");
    rc.id = "readRecipients";
    rc.className = "card";
    rc.hidden = true;
    panel.insertBefore(rc, document.querySelector("#panelExisting .preview-box"));
    var res = document.createElement("div");
    res.id = "readResult";
    res.className = "hint";
    panel.appendChild(res);
  }
  box.innerHTML = '<div class="field-label">This email</div>' +
    '<div style="font-weight:600">' + esc(readMsg.subject || "(no subject)") + '</div>' +
    '<div class="hint">' + esc((readMsg.fromName ? readMsg.fromName + " · " : "") + readMsg.from +
    (readMsg.dateText ? " · " + readMsg.dateText : "")) + '</div>';
  var d = byId("descInput");
  if (d && (fromItemChange || !d.value)) { d.value = cleanSubject(readMsg.subject); }
  var lbl = document.querySelector("#panelExisting .preview-label");
  if (lbl) { lbl.textContent = "Will be added to Monday"; }
  var c = byId("copyExisting");
  if (c) { c.hidden = true; }
  showMailButtons(false);
  var mr = byId("mailRowExisting");
  if (mr) { mr.style.display = "none"; }   /* .mail-row CSS overrides [hidden] */
  setText("applyExisting", "Add to Monday");
  if (CONFIG.LINK_ENDPOINT.indexOf("PASTE_") === 0) {
    setStatus("readResult", "Add to Monday is not configured yet (LINK_ENDPOINT).", true);
  } else {
    setStatus("readResult", "", false);
  }
  renderReadRecipients();
  renderExistingPreview();

  /* pinned panel: follow the email the user switches to */
  if (!readHandlerAdded && Office.context.mailbox.addHandlerAsync && Office.EventType && Office.EventType.ItemChanged) {
    readHandlerAdded = true;
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
      if (detectReadMode()) { state.selectedItem = null; enterReadMode(true); }
    });
  }
}

function renderReadRecipients() {
  var rc = byId("readRecipients");
  if (!rc) { return; }
  var isRfp = byId("typeSelect").value === "RFP";
  /* RFP always makes (or finds) one invitation per sub — no "continue existing" */
  var seg = document.querySelector("#panelExisting .segmented");
  /* style.display, not .hidden: the CSS gives .segmented its own display */
  if (seg) {
    seg.style.display = isRfp ? "none" : "";
    if (seg.previousElementSibling) { seg.previousElementSibling.style.display = isRfp ? "none" : ""; }
  }
  if (isRfp && isThreadMode()) { setMode("new"); }
  rc.hidden = !isRfp;
  if (!isRfp) { return; }
  rc.innerHTML = "";
  var head = document.createElement("div");
  head.className = "field-label";
  head.textContent = "Bid Invitation for";
  rc.appendChild(head);
  if (!state.readRecips.length) {
    var none = document.createElement("div");
    none.className = "hint error";
    none.textContent = "No outside address in this email.";
    rc.appendChild(none);
    return;
  }
  state.readRecips.forEach(function (r) {
    var row = document.createElement("label");
    row.style.display = "flex";
    row.style.alignItems = "center";
    row.style.gap = "8px";
    row.style.margin = "6px 0";
    var cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = r.checked;
    cb.addEventListener("change", function () { r.checked = cb.checked; renderExistingPreview(); });
    var sp = document.createElement("span");
    sp.textContent = r.email;
    row.appendChild(cb);
    row.appendChild(sp);
    rc.appendChild(row);
  });
}

function buildReadPreview() {
  var p = state.selectedProject;
  if (!p || !readMsg) { return null; }
  var type = byId("typeSelect").value;
  if (type === "RFP") {
    var chosen = state.readRecips.filter(function (r) { return r.checked; });
    if (!chosen.length) { return null; }
    return "Bid Invitations · " + (p.name || p.code) + " → " +
           chosen.map(function (r) { return r.email; }).join(", ");
  }
  if (isThreadMode()) {
    var it = state.selectedItem;
    if (!it) { return null; }
    var nl = numLabel(type, it.num);
    return folderLabel(type) + " · existing row: " + (nl ? nl + " " : "") + it.name;
  }
  var desc = byId("descInput").value.trim() || cleanSubject(readMsg.subject) || "Email";
  return folderLabel(type) + " · new row: " + desc + " (" + (p.name || p.code) + ")";
}

function addToMonday() {
  var p = state.selectedProject, type = byId("typeSelect").value;
  if (!p) { showToast("Pick a project first.", "err"); return; }
  if (state.readBusy) { return; }
  if (CONFIG.LINK_ENDPOINT.indexOf("PASTE_") === 0) { showToast("Add to Monday is not configured yet.", "err"); return; }
  var payload = {
    code: p.code, type: type, mode: "new", itemId: "",
    description: byId("descInput").value.trim(),
    recipients: [],
    user: (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || "",
    message: {}
  };
  if (type === "RFP") {
    payload.recipients = state.readRecips.filter(function (r) { return r.checked; }).map(function (r) { return r.email; });
    if (!payload.recipients.length) { showToast("Tick at least one subcontractor.", "err"); return; }
  } else if (isThreadMode()) {
    if (!state.selectedItem) { showToast("Pick the existing row.", "err"); return; }
    payload.mode = "existing";
    payload.itemId = String(state.selectedItem.id || "");
  }
  state.readBusy = true;
  setText("applyExisting", "Adding…");
  renderExistingPreview();
  setStatus("readResult", "Reading the email…", false);

  var finish = function () {
    state.readBusy = false;
    setText("applyExisting", "Add to Monday");
    renderExistingPreview();
  };
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (res) {
    var body = (res && res.status === Office.AsyncResultStatus.Succeeded) ? (res.value || "") : "";
    if (body.length > 60000) { body = body.slice(0, 60000); }
    payload.message = {
      subject: readMsg.subject, from: readMsg.from, fromName: readMsg.fromName,
      to: readMsg.to, cc: readMsg.cc, date: readMsg.date,
      internetMessageId: readMsg.internetMessageId, conversationId: readMsg.conversationId,
      body: body
    };
    setStatus("readResult", "Adding to Monday…", false);
    fetch(CONFIG.LINK_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    })
      .then(function (r) {
        return r.json().catch(function () { return { ok: false, error: "HTTP " + r.status }; });
      })
      .then(function (data) { showReadResult(data || {}); })
      .catch(function (err) {
        setStatus("readResult", "Could not reach Monday (" + err.message + "). Nothing was added — try again.", true);
      })
      .then(finish, finish);
  });
}

function showReadResult(data) {
  var el = byId("readResult");
  if (!el) { return; }
  var items = data.items || [];
  if (!data.ok && !items.length) {
    setStatus("readResult", "Not added: " + (data.error || "unknown error"), true);
    return;
  }
  el.classList.toggle("error", !data.ok);
  el.innerHTML = "";
  var head = document.createElement("div");
  head.style.fontWeight = "600";
  head.textContent = data.ok ? "Added to Monday ✓" : ("Partly added: " + (data.error || ""));
  el.appendChild(head);
  items.forEach(function (i) {
    var a = document.createElement("a");
    a.href = "https://gekonny.monday.com/boards/" + i.board + "/pulses/" + i.id;
    a.target = "_blank";
    a.rel = "noopener";
    a.textContent = (i.name || ("#" + i.id)) + (String(i.existing) === "true" ? " (already there — email added)" : "");
    a.style.display = "block";
    el.appendChild(a);
  });
  if (data.threadFound === false || String(data.threadFound) === "false") {
    var n = document.createElement("div");
    n.textContent = "build@ has no copy of this email, so a Follow-up from Monday will start a new email.";
    el.appendChild(n);
  }
  if (data.ok) { showToast("Added to Monday ✓", "ok"); }
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

/* Both type pickers are built from TYPES. They used to disagree — one was
   filled from here, the other was hand-written in the HTML — which meant
   adding a folder in one place silently left the other a list short. */
function fillTypes(id) {
  var sel = byId(id);
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
  renderFromQueue();
}

/* ---------- picking a queued task on the Subject tab -------------------

   The queue is where tasks are collected; the Subject tab is where one is
   turned into an email. Picking here fills the normal form rather than
   bypassing it, so the preview, Apply and the mail buttons all keep
   working exactly as they do for anything typed by hand.
   --------------------------------------------------------------------- */

function renderFromQueue() {
  var card = byId("fromQueueCard"), sel = byId("fromQueue");
  if (!card || !sel) { return; }

  var n = state.queue.length;
  card.hidden = !n;
  setText("fromQueueCount", n ? "(" + n + ")" : "");
  if (!n) { unloadFromQueue(); return; }

  var keep = state.loadedFromQueue ? state.loadedFromQueue.id : "";
  sel.innerHTML = "";
  var none = document.createElement("option");
  none.value = "";
  none.textContent = "— pick a task —";
  sel.appendChild(none);

  state.queue.forEach(function (row) {
    var o = document.createElement("option");
    o.value = row.id;
    o.textContent = row.code + " · " + row.type + " · " + row.description;
    sel.appendChild(o);
  });
  sel.value = keep;
  if (sel.value !== keep) { unloadFromQueue(); }
}

function onFromQueuePick() {
  var id = byId("fromQueue").value;
  if (!id) { unloadFromQueue(); return; }

  var i = qFind(id);
  if (i < 0) { unloadFromQueue(); renderFromQueue(); return; }
  var row = state.queue[i];

  /* Prefer the real project record — it carries the id the item lookup
     needs. Fall back to what the row itself remembers. */
  var p = null;
  for (var k = 0; k < state.projects.length; k++) {
    if (state.projects[k].code === row.code) { p = state.projects[k]; break; }
  }
  if (!p) { p = { name: row.projectName, code: row.code }; }

  setMode("new");
  selectProject(p);
  byId("typeSelect").value = row.type;
  byId("descInput").value = row.description;

  state.loadedFromQueue = row;
  renderExistingPreview();

  var done = byId("fromQueueDone");
  if (done) { done.hidden = false; }
  byId("fromQueue").value = row.id;
}

function unloadFromQueue() {
  state.loadedFromQueue = null;
  var done = byId("fromQueueDone"), sel = byId("fromQueue");
  if (done) { done.hidden = true; }
  if (sel && sel.value) { sel.value = ""; }
}

function loadedRecipient() {
  return state.loadedFromQueue ? state.loadedFromQueue.to : "";
}

function crossOffLoaded() {
  if (!state.loadedFromQueue) { return; }
  var id = state.loadedFromQueue.id;
  unloadFromQueue();
  queueArchive(id, "sent");
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
