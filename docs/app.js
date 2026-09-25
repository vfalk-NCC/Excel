// ---------------------------------------------------------------------
// Anteckningar & att göra – 3D-viewer-extension för Trimble Connect
// ---------------------------------------------------------------------
// Anteckningar och att göra-punkter som kan kopplas till objekt och en
// sparad kameravy i 3D-modellen, samt (valfritt) till ett planerat objekt
// i 4D-planering. Allt kan exporteras till Excel eller Markdown.
//
// Lagring:
//   - Med GitHub-token: projects/<projekt-id>/field_notes.json i det
//     privata repot vfalk-NCC/4D-data – samma datalager och samma
//     projektmapp som 4D-planering/4D-dashboard, via github-storage.js.
//   - Utan token: localStorage i den här webbläsaren (bara för dig).
//
// Koppling mot 4D-planering:
//   - Läser plan_items.json (skrivskyddat) för att kunna koppla en post
//     till ett planerat objekt och visa dess status/datum.
//   - Kopplar man markerade 3D-objekt som redan är planerade i
//     4D-planering föreslås det planerade objektet automatiskt (matchning
//     på objektets externa ID/IFC GUID, samma nyckel som 4D-planering).
//   - "Skicka till 4D-planering" lägger posten som en kommentar i
//     plan_item_comments.json, så den syns under 💬 i 4D-planering.
// ---------------------------------------------------------------------

const SETTINGS_KEY = "tcnotes-settings";
// 4D-planering hostas på samma origin (vfalk-ncc.github.io), så dess
// sparade token går att återanvända utan att användaren klistrar in den igen.
const PLAN_SETTINGS_KEY = "4dplan-settings";

const STATUS_LABELS = {
  ej_planerad: "Ej planerad",
  planerad: "Planerad",
  pagaende: "Pågående",
  forsenad: "Försenad",
  klar: "Klar",
  pausad: "Pausad",
};
const PRIORITY_LABELS = { hog: "Hög", normal: "Normal", lag: "Låg" };

let API = null;
let projectId = null;
let projectName = "";
let settings = { userName: "", githubToken: "" };

let items = [];                 // alla anteckningar + att göra (radformat, snake_case)
let planItems = [];             // plan_items.json från 4D-planering
let planById = new Map();
let planByObjectId = new Map(); // externt objekt-ID -> planerat objekt
let planLabelToId = new Map();  // datalist-text -> plan_item id

let activeTab = "note";
let editingId = null;
let draftObjects = [];          // [{model_id, object_id, object_name}]
let draftCamera = null;         // kameravy från posten som redigeras

let selectionRuntime = [];      // [{modelId, objectRuntimeIds}] - senaste markeringen
let selectionKeys = new Set();  // "modelId|externtId" för "Endast markerade objekt"
let selectionKeysDirty = true;

/* ---------------------------------------------------------------------
   Init
   ------------------------------------------------------------------- */
window.addEventListener("DOMContentLoaded", initApp);

async function initApp() {
  loadSettings();
  bindUI();
  switchTab("note");

  try {
    API = await TrimbleConnectWorkspace.connect(window.parent, onWorkspaceEvent, 30000);
    const project = await API.project.getProject();
    projectId = project.id;
    projectName = project.name || "";
  } catch (e) {
    console.error("Kunde inte ansluta till Trimble Connect", e);
    showWarning("⚠️ Kunde inte ansluta till Trimble Connect. Öppna extensionen från 3D-visaren i ett projekt.");
    return;
  }

  updateStorageBadge();
  await refreshAll();
  await onSelectionChanged();
}

async function refreshAll() {
  const btn = document.getElementById("btnRefresh");
  btn.disabled = true;
  btn.classList.add("spinning");
  try {
    await Promise.all([loadItems(), loadPlanItems()]);
    buildPlanItemOptions();
    render();
  } catch (e) {
    console.error(e);
    showWarning("⚠️ Kunde inte hämta data: " + e.message);
  } finally {
    btn.disabled = false;
    btn.classList.remove("spinning");
  }
}

function onWorkspaceEvent(event) {
  if (event === "viewer.onSelectionChanged" || event === "extension.onSelectionChanged") {
    onSelectionChanged();
  }
}

/* ---------------------------------------------------------------------
   Inställningar
   ------------------------------------------------------------------- */
function loadSettings() {
  try {
    const raw = window.localStorage.getItem(SETTINGS_KEY);
    if (raw) settings = { ...settings, ...JSON.parse(raw) };
  } catch (e) { /* ignorera */ }
  if (!settings.userName) {
    const plan = readPlanSettings();
    if (plan.userName) settings.userName = plan.userName;
  }
}

function readPlanSettings() {
  try {
    return JSON.parse(window.localStorage.getItem(PLAN_SETTINGS_KEY) || "{}") || {};
  } catch (e) {
    return {};
  }
}

/** Egen token i första hand, annars 4D-planeringens. null = lokalt läge. */
function resolveToken() {
  if (settings.githubToken) return { token: settings.githubToken, source: "egen" };
  const planToken = readPlanSettings().githubToken;
  if (planToken) return { token: planToken, source: "4D-planering" };
  return null;
}

function token() {
  const t = resolveToken();
  return t ? t.token : null;
}

function updateStorageBadge() {
  const badge = document.getElementById("storageBadge");
  const t = resolveToken();
  if (t) {
    badge.innerText = `Delad lagring (4D-data)${t.source === "4D-planering" ? " · token från 4D-planering" : ""}`;
    hideWarning();
  } else {
    badge.innerText = "Lokalt läge";
    showWarning("Lokalt läge – sparas bara i den här webbläsaren. Ange GitHub-token under ⚙ för att dela med projektet och koppla mot 4D-planering.");
  }
}

function openSettings() {
  document.getElementById("sUserName").value = settings.userName || "";
  document.getElementById("sToken").value = settings.githubToken || "";
  const dlg = document.getElementById("settingsDialog");
  dlg.returnValue = "";
  dlg.showModal();
}

async function onSettingsClosed() {
  const dlg = document.getElementById("settingsDialog");
  if (dlg.returnValue !== "save") return;
  const hadToken = token();
  settings.userName = document.getElementById("sUserName").value.trim();
  settings.githubToken = document.getElementById("sToken").value.trim();
  try {
    window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  } catch (e) { /* ignorera */ }
  updateStorageBadge();
  if (hadToken !== token() && projectId) await refreshAll();
}

/* ---------------------------------------------------------------------
   Lagring
   ------------------------------------------------------------------- */
function projectPath(file) {
  return `projects/${encodeURIComponent(projectId)}/${file}`;
}
function localKey() {
  return `tcnotes-items-${projectId}`;
}

async function loadItems() {
  const t = token();
  if (t) {
    items = await ghReadJSON(t, projectPath("field_notes.json"));
  } else {
    try {
      items = JSON.parse(window.localStorage.getItem(localKey()) || "[]");
    } catch (e) {
      items = [];
    }
  }
}

/** Kör mutateFn på listan, lokalt direkt (optimistiskt) och sedan mot lagringen. */
async function persist(mutateFn, message) {
  items = mutateFn(items.slice());
  render();
  const t = token();
  if (!t) {
    try {
      window.localStorage.setItem(localKey(), JSON.stringify(items));
    } catch (e) {
      alert("Kunde inte spara lokalt: " + e.message);
    }
    return;
  }
  setSaveStatus("Sparar…");
  try {
    items = await ghWriteJSON(t, projectPath("field_notes.json"), mutateFn, message);
    setSaveStatus("");
    render();
  } catch (e) {
    console.error(e);
    setSaveStatus("");
    alert("Kunde inte spara: " + e.message);
    await loadItems();
    render();
  }
}

async function loadPlanItems() {
  const t = token();
  planItems = t ? await ghReadJSON(t, projectPath("plan_items.json")) : [];
  planById = new Map(planItems.map(p => [String(p.id), p]));
  planByObjectId = new Map(planItems.filter(p => p.object_id).map(p => [String(p.object_id), p]));
}

function planLabel(p) {
  return [p.object_name || p.object_id, p.activity, p.area].filter(Boolean).join(" · ");
}

function buildPlanItemOptions() {
  const list = document.getElementById("planItemOptions");
  list.innerHTML = "";
  planLabelToId = new Map();
  const seen = new Map();
  planItems.forEach(p => {
    let label = planLabel(p) || String(p.id);
    const n = (seen.get(label) || 0) + 1;
    seen.set(label, n);
    if (n > 1) label = `${label} (${n})`;
    planLabelToId.set(label, String(p.id));
    const opt = document.createElement("option");
    opt.value = label;
    list.appendChild(opt);
  });
  const input = document.getElementById("fPlanItem");
  input.disabled = planItems.length === 0;
  input.placeholder = planItems.length
    ? "Sök planerat objekt (valfritt)…"
    : (token() ? "Inga planerade objekt i 4D-planering för projektet" : "Kräver GitHub-token (⚙)");
}

function labelForPlanId(id) {
  for (const [label, pid] of planLabelToId) if (pid === String(id)) return label;
  return "";
}

/* ---------------------------------------------------------------------
   Markering i 3D
   ------------------------------------------------------------------- */
async function onSelectionChanged() {
  if (!API) return;
  try {
    selectionRuntime = await API.viewer.getSelection();
  } catch (e) {
    selectionRuntime = [];
  }
  const count = selectionRuntime.reduce((n, m) => n + (m.objectRuntimeIds || []).length, 0);
  document.getElementById("selCount").innerText = count;
  document.getElementById("btnAttachSelection").disabled = count === 0;
  selectionKeysDirty = true;
  if (document.getElementById("onlySelected").checked) {
    await ensureSelectionKeys();
    render();
  }
}

async function ensureSelectionKeys() {
  if (!selectionKeysDirty) return;
  const keys = new Set();
  for (const group of selectionRuntime) {
    if (!group.objectRuntimeIds || group.objectRuntimeIds.length === 0) continue;
    try {
      const ext = await API.viewer.convertToObjectIds(group.modelId, group.objectRuntimeIds);
      ext.forEach(id => { if (id) keys.add(`${group.modelId}|${id}`); });
    } catch (e) {
      console.warn("Kunde inte läsa objekt-ID:n för markeringen", e);
    }
  }
  selectionKeys = keys;
  selectionKeysDirty = false;
}

function objectName(props) {
  if (!props) return "";
  if (props.product && props.product.name) return props.product.name;
  for (const pset of props.properties || []) {
    for (const prop of pset.properties || []) {
      if (/^name$/i.test(prop.name) && prop.value) return String(prop.value);
    }
  }
  return "";
}

/** Läser den aktuella markeringen som [{model_id, object_id, object_name}]. */
async function readSelectionObjects() {
  const result = [];
  for (const group of selectionRuntime) {
    const runtimeIds = group.objectRuntimeIds || [];
    if (runtimeIds.length === 0) continue;
    const ext = await API.viewer.convertToObjectIds(group.modelId, runtimeIds);
    let props = [];
    // Namn är bara för visning - hoppa över för jättemarkeringar.
    if (runtimeIds.length <= 200) {
      try {
        props = await API.viewer.getObjectProperties(group.modelId, runtimeIds);
      } catch (e) { /* namn är valfria */ }
    }
    const propsById = new Map(props.map(p => [p.id, p]));
    runtimeIds.forEach((rid, i) => {
      if (!ext[i]) return;
      result.push({
        model_id: group.modelId,
        object_id: String(ext[i]),
        object_name: objectName(propsById.get(rid)) || null,
      });
    });
  }
  return result;
}

/** Markerar postens objekt i 3D och flyttar kameran (sparad vy om den finns). */
async function showInModel(rec) {
  let objs = rec.objects || [];
  const plan = rec.plan_item_id ? planById.get(String(rec.plan_item_id)) : null;
  if (objs.length === 0 && plan && plan.model_id && plan.object_id) {
    objs = [{ model_id: plan.model_id, object_id: plan.object_id }];
  }

  const byModel = {};
  objs.forEach(o => { (byModel[o.model_id] = byModel[o.model_id] || []).push(o.object_id); });

  const modelObjectIds = [];
  let missing = 0;
  for (const modelId of Object.keys(byModel)) {
    const runtimeIds = await convertToRuntimeIdsSafe(modelId, byModel[modelId]);
    const valid = runtimeIds.filter(id => id !== undefined && id !== null);
    missing += runtimeIds.length - valid.length;
    if (valid.length) modelObjectIds.push({ modelId, objectRuntimeIds: valid });
  }

  if (modelObjectIds.length) {
    await API.viewer.setSelection({ modelObjectIds }, "set");
  }
  if (rec.camera) {
    await API.viewer.setCamera(rec.camera, { animationTime: 600 });
  } else if (modelObjectIds.length) {
    await API.viewer.setCamera({ modelObjectIds }, { animationTime: 600 });
  } else if (objs.length) {
    alert("Hittade inte de kopplade objekten i den modell som är inläst just nu. Öppna rätt modell/version och försök igen.");
    return;
  }
  if (missing > 0) {
    setSaveStatus(`${missing} av objekten finns inte i den inlästa modellen.`, 4000);
  }
}

/** convertToObjectRuntimeIds som tål att enstaka objekt saknas i modellen. */
async function convertToRuntimeIdsSafe(modelId, objectIds) {
  try {
    return await API.viewer.convertToObjectRuntimeIds(modelId, objectIds);
  } catch (e) {
    const out = [];
    for (const id of objectIds) {
      try {
        const [rid] = await API.viewer.convertToObjectRuntimeIds(modelId, [id]);
        out.push(rid);
      } catch (err) {
        out.push(undefined);
      }
    }
    return out;
  }
}

/* ---------------------------------------------------------------------
   UI
   ------------------------------------------------------------------- */
function bindUI() {
  document.querySelectorAll(".tab").forEach(tab => {
    tab.onclick = () => switchTab(tab.dataset.tab);
  });
  document.getElementById("btnRefresh").onclick = refreshAll;
  document.getElementById("btnSettings").onclick = openSettings;
  document.getElementById("settingsDialog").addEventListener("close", onSettingsClosed);

  document.getElementById("btnNew").onclick = () => openEditor(null);
  document.getElementById("btnCancel").onclick = closeEditor;
  document.getElementById("btnSave").onclick = onSave;
  document.getElementById("btnAttachSelection").onclick = onAttachSelection;
  document.getElementById("btnClearObjects").onclick = () => { draftObjects = []; renderDraftObjects(); };
  document.getElementById("fPlanItem").oninput = renderPlanItemInfo;
  document.getElementById("fTitle").onkeydown = (e) => { if (e.key === "Enter") onSave(); };

  document.getElementById("search").oninput = render;
  document.getElementById("hideDone").onchange = render;
  document.getElementById("onlySelected").onchange = async () => {
    await ensureSelectionKeys();
    render();
  };

  document.getElementById("btnExportXlsx").onclick = exportXlsx;
  document.getElementById("btnExportMd").onclick = exportMarkdown;
  document.getElementById("btnCopy").onclick = copyText;
}

function switchTab(tab) {
  activeTab = tab;
  document.querySelectorAll(".tab").forEach(t => t.classList.toggle("active", t.dataset.tab === tab));
  document.querySelectorAll(".todo-only").forEach(el => el.classList.toggle("hidden", tab !== "todo"));
  document.getElementById("btnNew").innerText = tab === "todo" ? "+ Ny att göra-punkt" : "+ Ny anteckning";
  closeEditor();
  render();
}

function openEditor(rec) {
  editingId = rec ? rec.id : null;
  const type = rec ? rec.type : activeTab;
  document.getElementById("fTitle").value = rec ? rec.title || "" : "";
  document.getElementById("fBody").value = rec ? rec.body || "" : "";
  document.getElementById("fDue").value = rec ? rec.due_date || "" : "";
  document.getElementById("fAssignee").value = rec ? rec.assignee || "" : "";
  document.getElementById("fPriority").value = rec ? rec.priority || "normal" : "normal";
  document.getElementById("fPlanItem").value = rec && rec.plan_item_id ? labelForPlanId(rec.plan_item_id) : "";
  document.getElementById("fSaveView").checked = !!(rec && rec.camera);
  document.getElementById("todoFields").classList.toggle("hidden", type !== "todo");
  draftObjects = rec ? (rec.objects || []).slice() : [];
  draftCamera = rec ? rec.camera || null : null;
  renderDraftObjects();
  renderPlanItemInfo();
  document.getElementById("btnSave").innerText = rec ? "Spara ändringar" : "Spara";
  document.getElementById("editor").classList.remove("hidden");
  document.getElementById("btnNew").classList.add("hidden");
  document.getElementById("fTitle").focus();
}

function closeEditor() {
  editingId = null;
  draftObjects = [];
  draftCamera = null;
  document.getElementById("editor").classList.add("hidden");
  document.getElementById("btnNew").classList.remove("hidden");
}

async function onAttachSelection() {
  const btn = document.getElementById("btnAttachSelection");
  btn.disabled = true;
  try {
    const objs = await readSelectionObjects();
    const keys = new Set(draftObjects.map(o => `${o.model_id}|${o.object_id}`));
    objs.forEach(o => {
      if (!keys.has(`${o.model_id}|${o.object_id}`)) draftObjects.push(o);
    });
    renderDraftObjects();

    // Föreslå det planerade objektet i 4D-planering om markeringen redan är planerad där.
    const planInput = document.getElementById("fPlanItem");
    if (!planInput.value) {
      const match = objs.map(o => planByObjectId.get(o.object_id)).find(Boolean);
      if (match) {
        planInput.value = labelForPlanId(match.id);
        renderPlanItemInfo();
      }
    }
  } catch (e) {
    alert("Kunde inte läsa markeringen: " + e.message);
  } finally {
    btn.disabled = false;
  }
}

function renderDraftObjects() {
  const el = document.getElementById("fObjects");
  el.innerHTML = "";
  if (draftObjects.length > 0) el.appendChild(chip(objectsSummary(draftObjects)));
  document.getElementById("btnClearObjects").classList.toggle("hidden", draftObjects.length === 0);
}

function selectedPlanId() {
  const label = document.getElementById("fPlanItem").value.trim();
  return label ? planLabelToId.get(label) || null : null;
}

function renderPlanItemInfo() {
  const info = document.getElementById("fPlanItemInfo");
  const label = document.getElementById("fPlanItem").value.trim();
  const id = selectedPlanId();
  if (!label) info.innerText = "";
  else if (!id) info.innerText = "Välj ett objekt ur listan för att koppla.";
  else info.innerText = "✓ " + planSummary(planById.get(id));
}

async function onSave() {
  const title = document.getElementById("fTitle").value.trim();
  const body = document.getElementById("fBody").value.trim();
  if (!title && !body) {
    alert("Skriv en rubrik eller text.");
    return;
  }
  if (document.getElementById("fPlanItem").value.trim() && !selectedPlanId()) {
    alert("Det valda 4D-objektet finns inte – välj ur listan eller töm fältet.");
    return;
  }

  let camera = null;
  if (document.getElementById("fSaveView").checked) {
    try {
      // Ny vy vid ny post; vid redigering behålls den sparade vyn om den finns.
      camera = draftCamera || await API.viewer.getCamera();
    } catch (e) {
      console.warn("Kunde inte läsa kameravyn", e);
    }
  }

  const existing = editingId ? items.find(r => r.id === editingId) : null;
  const type = existing ? existing.type : activeTab;
  const now = new Date().toISOString();
  const rec = {
    ...(existing || {}),
    id: existing ? existing.id : ghNewId(),
    project_id: projectId,
    type,
    title,
    body,
    due_date: type === "todo" ? document.getElementById("fDue").value || null : null,
    assignee: type === "todo" ? document.getElementById("fAssignee").value.trim() || null : null,
    priority: type === "todo" ? document.getElementById("fPriority").value : null,
    done: existing ? !!existing.done : false,
    done_at: existing ? existing.done_at || null : null,
    plan_item_id: selectedPlanId(),
    objects: draftObjects.slice(),
    camera,
    author: existing ? existing.author : (settings.userName || null),
    created_at: existing ? existing.created_at : now,
    updated_at: now,
  };

  closeEditor();
  await persist(
    arr => {
      const idx = arr.findIndex(r => r.id === rec.id);
      if (idx >= 0) arr[idx] = rec; else arr.push(rec);
      return arr;
    },
    existing ? `Uppdatera ${type === "todo" ? "att göra" : "anteckning"}` : `Ny ${type === "todo" ? "att göra-punkt" : "anteckning"}`
  );
}

async function onToggleDone(rec, done) {
  const now = new Date().toISOString();
  await persist(
    arr => arr.map(r => r.id === rec.id ? { ...r, done, done_at: done ? now : null, updated_at: now } : r),
    done ? "Att göra klar" : "Att göra återöppnad"
  );
}

async function onDelete(rec) {
  if (!confirm(`Radera "${rec.title || "(utan rubrik)"}"?`)) return;
  await persist(arr => arr.filter(r => r.id !== rec.id), "Radera anteckning");
}

/** Lägger posten som kommentar på det planerade objektet i 4D-planering. */
async function onSendTo4D(rec) {
  const t = token();
  const plan = planById.get(String(rec.plan_item_id));
  if (!t || !plan) return;
  if (!confirm(`Lägga till som kommentar på "${planLabel(plan)}" i 4D-planering?`)) return;

  const lines = [`${rec.type === "todo" ? "☑️ Att göra" : "📝 Anteckning"}: ${rec.title || ""}`.trim()];
  if (rec.body) lines.push(rec.body);
  if (rec.type === "todo") {
    const extra = [
      rec.assignee && `Ansvarig: ${rec.assignee}`,
      rec.due_date && `Förfaller: ${rec.due_date}`,
      rec.priority && rec.priority !== "normal" && `Prioritet: ${PRIORITY_LABELS[rec.priority]}`,
    ].filter(Boolean);
    if (extra.length) lines.push(extra.join(" · "));
  }

  setSaveStatus("Skickar till 4D-planering…");
  try {
    const comment = {
      id: ghNewId(),
      created_at: new Date().toISOString(),
      plan_item_id: plan.id,
      parent_comment_id: null,
      author: settings.userName || rec.author || "Anonym",
      body: lines.join("\n"),
    };
    await ghWriteJSON(t, projectPath("plan_item_comments.json"), arr => [...arr, comment], "Ny kommentar (från Anteckningar)");
    setSaveStatus("");
    await persist(
      arr => arr.map(r => r.id === rec.id ? { ...r, sent_to_4d_at: comment.created_at } : r),
      "Markera anteckning som skickad till 4D-planering"
    );
    setSaveStatus("✓ Skickad till 4D-planering", 3000);
  } catch (e) {
    setSaveStatus("");
    alert("Kunde inte skicka: " + e.message);
  }
}

/* ---------------------------------------------------------------------
   Lista
   ------------------------------------------------------------------- */
function matchesSelection(rec) {
  if ((rec.objects || []).some(o => selectionKeys.has(`${o.model_id}|${o.object_id}`))) return true;
  const plan = rec.plan_item_id ? planById.get(String(rec.plan_item_id)) : null;
  return !!(plan && selectionKeys.has(`${plan.model_id}|${plan.object_id}`));
}

function searchText(rec) {
  const plan = rec.plan_item_id ? planById.get(String(rec.plan_item_id)) : null;
  return [
    rec.title, rec.body, rec.assignee, rec.author,
    plan && planLabel(plan),
    ...(rec.objects || []).map(o => o.object_name),
  ].filter(Boolean).join(" ").toLowerCase();
}

/** Posterna av en viss typ efter sökning och filter - samma urval som listan och exporten använder. */
function filteredItems(type) {
  const q = document.getElementById("search").value.trim().toLowerCase();
  const onlySelected = document.getElementById("onlySelected").checked;
  const hideDone = document.getElementById("hideDone").checked;
  let out = items.filter(r => r.type === type);
  if (q) out = out.filter(r => searchText(r).includes(q));
  if (onlySelected) out = out.filter(matchesSelection);
  if (type === "todo" && hideDone) out = out.filter(r => !r.done);
  return sortItems(out, type);
}

function sortItems(arr, type) {
  if (type === "todo") {
    const prio = { hog: 0, normal: 1, lag: 2 };
    return arr.sort((a, b) =>
      (Number(!!a.done) - Number(!!b.done)) ||
      ((a.due_date || "9999") < (b.due_date || "9999") ? -1 : (a.due_date || "9999") > (b.due_date || "9999") ? 1 : 0) ||
      ((prio[a.priority] ?? 1) - (prio[b.priority] ?? 1)) ||
      (b.created_at || "").localeCompare(a.created_at || "")
    );
  }
  return arr.sort((a, b) => (b.updated_at || "").localeCompare(a.updated_at || ""));
}

function render() {
  document.getElementById("countNote").innerText = items.filter(r => r.type === "note").length;
  document.getElementById("countTodo").innerText = items.filter(r => r.type === "todo" && !r.done).length;

  const list = document.getElementById("list");
  list.innerHTML = "";
  const rows = filteredItems(activeTab);
  const highlight = selectionKeys.size > 0 && !selectionKeysDirty;
  rows.forEach(rec => list.appendChild(renderCard(rec, highlight && matchesSelection(rec))));
  const empty = document.getElementById("emptyHint");
  empty.classList.toggle("hidden", rows.length > 0);
  empty.innerText = items.some(r => r.type === activeTab) ? "Inget matchar filtret." : "Inget här än.";
}

function renderCard(rec, isSelectedMatch) {
  const li = document.createElement("li");
  li.className = "card" + (rec.done ? " done" : "") + (isSelectedMatch ? " selected-match" : "");

  const head = el("div", "card-head");
  if (rec.type === "todo") {
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!rec.done;
    cb.title = rec.done ? "Markera som ej klar" : "Markera som klar";
    cb.onchange = () => onToggleDone(rec, cb.checked);
    head.appendChild(cb);
  }
  head.appendChild(el("div", "card-title", rec.title || "(utan rubrik)"));
  li.appendChild(head);

  if (rec.body) li.appendChild(el("div", "card-body", rec.body));

  const meta = el("div", "card-meta");
  if (rec.type === "todo") {
    if (rec.priority && rec.priority !== "normal") {
      meta.appendChild(el("span", "prio-" + rec.priority, `Prio ${PRIORITY_LABELS[rec.priority]}`));
      meta.appendChild(document.createTextNode(" · "));
    }
    if (rec.due_date) {
      const overdue = !rec.done && rec.due_date < todayIso();
      meta.appendChild(el("span", overdue ? "overdue" : "", `Förfaller ${rec.due_date}`));
      meta.appendChild(document.createTextNode(" · "));
    }
    if (rec.assignee) meta.appendChild(document.createTextNode(`${rec.assignee} · `));
  }
  meta.appendChild(document.createTextNode(
    [rec.author, formatDateTime(rec.updated_at || rec.created_at)].filter(Boolean).join(" · ")
  ));
  li.appendChild(meta);

  const chips = el("div", "chips");
  if ((rec.objects || []).length) chips.appendChild(chip(objectsSummary(rec.objects)));
  if (rec.camera) chips.appendChild(chip("🎥 Sparad vy"));
  if (rec.plan_item_id) {
    const plan = planById.get(String(rec.plan_item_id));
    if (plan) {
      chips.appendChild(chip("4D: " + planSummary(plan), "plan" + (plan.status === "forsenad" ? " status-forsenad" : "")));
    } else if (token()) {
      chips.appendChild(chip("4D-objektet finns inte längre", "warn"));
    }
  }
  if (rec.sent_to_4d_at) chips.appendChild(chip("💬 Skickad till 4D", "plan"));
  if (chips.childNodes.length) li.appendChild(chips);

  const actions = el("div", "card-actions");
  const canShow = (rec.objects || []).length || rec.camera || rec.plan_item_id;
  if (canShow) actions.appendChild(iconButton("👁", "Visa i modellen", () => showInModel(rec).catch(e => alert(e.message))));
  if (rec.plan_item_id && token() && planById.has(String(rec.plan_item_id))) {
    actions.appendChild(iconButton("💬", "Skicka som kommentar till 4D-planering", () => onSendTo4D(rec)));
  }
  actions.appendChild(iconButton("✏️", "Redigera", () => openEditor(rec)));
  actions.appendChild(iconButton("🗑️", "Radera", () => onDelete(rec)));
  li.appendChild(actions);
  return li;
}

/* ---------------------------------------------------------------------
   Export
   ------------------------------------------------------------------- */
function planColumns(rec) {
  const plan = rec.plan_item_id ? planById.get(String(rec.plan_item_id)) : null;
  return {
    "4D-objekt": plan ? planLabel(plan) : "",
    "4D-status": plan ? STATUS_LABELS[plan.status] || plan.status || "" : "",
    "4D-start": plan ? plan.start_date || "" : "",
    "4D-slut": plan ? plan.end_date || "" : "",
  };
}

function objectColumns(rec) {
  const objs = rec.objects || [];
  return {
    "Kopplade objekt": objs.map(o => o.object_name || o.object_id).join(", "),
    "Objekt-ID": objs.map(o => o.object_id).join(", "),
  };
}

function exportXlsx() {
  if (typeof XLSX === "undefined") {
    alert("Excel-biblioteket kunde inte laddas. Använd Markdown-exporten i stället.");
    return;
  }
  const notes = filteredItems("note").map(r => ({
    Rubrik: r.title || "",
    Text: r.body || "",
    Författare: r.author || "",
    Skapad: formatDateTime(r.created_at),
    Uppdaterad: formatDateTime(r.updated_at),
    ...planColumns(r),
    ...objectColumns(r),
  }));
  const todos = filteredItems("todo").map(r => ({
    Klar: r.done ? "Ja" : "Nej",
    Rubrik: r.title || "",
    Text: r.body || "",
    Ansvarig: r.assignee || "",
    Förfallodatum: r.due_date || "",
    Prioritet: PRIORITY_LABELS[r.priority] || "",
    "Klar datum": r.done_at ? formatDateTime(r.done_at) : "",
    Författare: r.author || "",
    Skapad: formatDateTime(r.created_at),
    ...planColumns(r),
    ...objectColumns(r),
  }));
  if (notes.length === 0 && todos.length === 0) {
    alert("Inget att exportera.");
    return;
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheetWithWidths(todos, ["Klar", "Rubrik", "Text"]), "Att göra");
  XLSX.utils.book_append_sheet(wb, sheetWithWidths(notes, ["Rubrik", "Text"]), "Anteckningar");
  XLSX.writeFile(wb, `${fileBaseName()}.xlsx`);
}

function sheetWithWidths(rows, fallbackHeader) {
  const sheet = rows.length ? XLSX.utils.json_to_sheet(rows) : XLSX.utils.aoa_to_sheet([fallbackHeader]);
  const header = rows.length ? Object.keys(rows[0]) : fallbackHeader;
  sheet["!cols"] = header.map(h => ({
    wch: Math.min(60, Math.max(h.length, ...rows.map(r => String(r[h] || "").length)) + 2),
  }));
  return sheet;
}

function buildMarkdown() {
  const lines = [`# Anteckningar & att göra${projectName ? " – " + projectName : ""}`, "", `Exporterad ${formatDateTime(new Date().toISOString())}`, ""];

  const todos = filteredItems("todo");
  lines.push("## Att göra", "");
  if (!todos.length) lines.push("_Inga punkter._", "");
  todos.forEach(r => {
    const meta = [
      r.priority && r.priority !== "normal" && `prio ${PRIORITY_LABELS[r.priority].toLowerCase()}`,
      r.due_date && `förfaller ${r.due_date}`,
      r.assignee && `ansvarig ${r.assignee}`,
    ].filter(Boolean);
    lines.push(`- [${r.done ? "x" : " "}] **${r.title || "(utan rubrik)"}**${meta.length ? " – " + meta.join(", ") : ""}`);
    if (r.body) r.body.split("\n").forEach(l => lines.push(`  ${l}`));
    refLines(r).forEach(l => lines.push(`  ${l}`));
  });
  lines.push("");

  const notes = filteredItems("note");
  lines.push("## Anteckningar", "");
  if (!notes.length) lines.push("_Inga anteckningar._", "");
  notes.forEach(r => {
    lines.push(`### ${r.title || "(utan rubrik)"}`);
    lines.push(`_${[r.author, formatDateTime(r.updated_at || r.created_at)].filter(Boolean).join(" · ")}_`, "");
    if (r.body) lines.push(r.body, "");
    refLines(r).forEach(l => lines.push(l));
    lines.push("");
  });
  return lines.join("\n");
}

function refLines(r) {
  const out = [];
  const plan = r.plan_item_id ? planById.get(String(r.plan_item_id)) : null;
  if (plan) out.push(`- 4D: ${planSummary(plan)}`);
  if ((r.objects || []).length) out.push(`- Objekt: ${r.objects.map(o => o.object_name || o.object_id).join(", ")}`);
  return out;
}

function exportMarkdown() {
  const blob = new Blob([buildMarkdown()], { type: "text/markdown;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${fileBaseName()}.md`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

async function copyText() {
  const text = buildMarkdown();
  try {
    await navigator.clipboard.writeText(text);
  } catch (e) {
    // Urklipps-API:t kan vara blockerat i Trimbles iframe - fall tillbaka.
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
  }
  setSaveStatus("✓ Kopierat till urklipp", 2500);
}

function fileBaseName() {
  const safe = projectName.replace(/[\\/:*?"<>|]+/g, "").trim();
  return `Anteckningar-${safe ? safe + "-" : ""}${todayIso()}`;
}

/* ---------------------------------------------------------------------
   Hjälpfunktioner
   ------------------------------------------------------------------- */
function el(tag, className, text) {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text !== undefined) e.textContent = text;
  return e;
}

function chip(text, extraClass) {
  const c = el("span", "chip" + (extraClass ? " " + extraClass : ""), text);
  c.title = text;
  return c;
}

function iconButton(icon, title, onClick) {
  const b = el("button", "", icon);
  b.type = "button";
  b.title = title;
  b.onclick = onClick;
  return b;
}

function objectsSummary(objs) {
  const names = objs.map(o => o.object_name).filter(Boolean);
  if (objs.length === 1) return `🧱 ${names[0] || "1 objekt"}`;
  const shown = names.slice(0, 2).join(", ");
  return `🧱 ${objs.length} objekt${shown ? " (" + shown + (names.length > 2 ? ", …" : "") + ")" : ""}`;
}

function planSummary(p) {
  if (!p) return "";
  const status = STATUS_LABELS[p.status] || p.status || "";
  const dates = p.start_date || p.end_date ? `${p.start_date || "?"} → ${p.end_date || "?"}` : "";
  return [planLabel(p), status, dates].filter(Boolean).join(" · ");
}

function todayIso() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function formatDateTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d)) return "";
  return d.toLocaleString("sv-SE", { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
}

function showWarning(text) {
  const w = document.getElementById("connectionWarning");
  w.innerText = text;
  w.classList.remove("hidden");
}

function hideWarning() {
  document.getElementById("connectionWarning").classList.add("hidden");
}

let saveStatusTimer = null;
function setSaveStatus(text, clearAfterMs) {
  const s = document.getElementById("saveStatus");
  clearTimeout(saveStatusTimer);
  s.innerText = text;
  s.classList.toggle("hidden", !text);
  if (text && clearAfterMs) saveStatusTimer = setTimeout(() => setSaveStatus(""), clearAfterMs);
}
