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

// Visas i headern - gör det lätt att se om Trimble kör senaste versionen
// (GitHub Pages cachar filerna ~10 min). Räkna upp vid varje release och
// uppdatera ?v= i index.html samtidigt.
const APP_VERSION = "2026-09-26.6";
const SETTINGS_KEY = "tcnotes-settings";
const DEFAULT_BUBBLE_SIZE = 256;
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
let settings = { userName: "", githubToken: "", showBubbles: false, bubbleSize: DEFAULT_BUBBLE_SIZE, bubbleSource: "generated", bubbleScale: true, bubbleRefDistance: null };

let items = [];                 // alla anteckningar + att göra (radformat, snake_case)
let planItems = [];             // plan_items.json från 4D-planering
let planById = new Map();
let planByObjectId = new Map(); // externt objekt-ID -> planerat objekt
let planLabelToId = new Map();  // datalist-text -> plan_item id

let activeTab = "note";
let editingId = null;
let draftObjects = [];          // [{model_id, object_id, object_name}]
let draftCamera = null;         // kameravy från posten som redigeras
let draftPin = null;            // {x, y, z} i meter - egen punkt för 3D-bubblan
let pickArmed = false;          // väntar på ett klick i modellen för draftPin

let selectionRuntime = [];      // [{modelId, objectRuntimeIds}] - senaste markeringen
let selectionKeys = new Set();  // "modelId|externtId" för "Endast markerade objekt"
let selectionKeysDirty = true;

/* ---------------------------------------------------------------------
   Init
   ------------------------------------------------------------------- */
window.addEventListener("DOMContentLoaded", initApp);

async function initApp() {
  document.getElementById("versionBadge").innerText = "v" + APP_VERSION;
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
    // Oberoende av varandra: ett fel på 4D-planeringens data ska inte
    // hindra anteckningarna från att visas.
    const [notesRes, planRes] = await Promise.allSettled([loadItems(), loadPlanItems()]);
    buildPlanItemOptions();
    render();
    const failed = [notesRes, planRes].find(r => r.status === "rejected");
    if (failed) {
      console.error(failed.reason);
      showWarning("⚠️ Kunde inte hämta data: " + failed.reason.message);
    } else {
      updateStorageBadge();
    }
  } finally {
    btn.disabled = false;
    btn.classList.remove("spinning");
  }
}

function onWorkspaceEvent(event, arg) {
  // Händelsedata kommer som { data: ... } (EventArgument i Workspace API).
  const data = arg && typeof arg === "object" && "data" in arg ? arg.data : arg;
  if (event === "viewer.onSelectionChanged" || event === "extension.onSelectionChanged") {
    onSelectionChanged();
  } else if (event === "viewer.onPicked") {
    onPointPicked(data);
  } else if (event === "viewer.onCameraChanged") {
    onCameraChanged(data);
  } else if (event === "viewer.onIconPicked") {
    onBubblePicked(data);
  } else if (event === "viewer.onModelStateChanged" || event === "viewer.onModelReset") {
    // En modell laddades/togs bort - positioner kan ha tillkommit eller försvunnit.
    bubbleState.bboxCache.clear();
    scheduleBubbles();
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
  // Tidigare standardstorlek (96) gav oläsligt små bubblor i Trimble.
  if (!settings.bubbleSizeV2) {
    settings.bubbleSize = DEFAULT_BUBBLE_SIZE;
    settings.bubbleSizeV2 = true;
  }
  if (!settings.userName) {
    const plan = readPlanSettings();
    if (plan.userName) settings.userName = plan.userName;
  }
}

function saveSettings() {
  try {
    window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  } catch (e) { /* ignorera */ }
}

function readPlanSettings() {
  try {
    return JSON.parse(window.localStorage.getItem(PLAN_SETTINGS_KEY) || "{}") || {};
  } catch (e) {
    return {};
  }
}

/** Egen token i första hand, annars 4D-planeringens. null = lokalt läge. */
// Token:ar som GitHub svarat 401 (ogiltig/utgången) på i den här sessionen.
const rejectedTokens = new Set();

/**
 * Egen token i första hand, annars 4D-planeringens. En token som GitHub
 * avvisat hoppas över, så att t.ex. en gammal egen token inte blockerar när
 * 4D-planeringens fungerar. Är alla avvisade returneras den första ändå,
 * markerad `invalid` (appen ska inte tyst byta till lokalt läge).
 * null = ingen token alls = lokalt läge.
 */
function resolveToken() {
  const candidates = [];
  if (settings.githubToken) candidates.push({ token: settings.githubToken, source: "egen" });
  const planToken = readPlanSettings().githubToken;
  if (planToken && planToken !== settings.githubToken) candidates.push({ token: planToken, source: "4D-planering" });
  if (!candidates.length) return null;
  return candidates.find(c => !rejectedTokens.has(c.token)) || { ...candidates[0], invalid: true };
}

function isAuthError(e) {
  return /\b401\b/.test(String(e && e.message));
}

/** Kör fn(token) och provar nästa token om GitHub svarar 401. */
async function withToken(fn) {
  for (;;) {
    const t = resolveToken();
    if (!t) throw new Error("Ingen GitHub-token.");
    try {
      return await fn(t.token);
    } catch (e) {
      if (!isAuthError(e) || t.invalid) {
        if (isAuthError(e)) {
          updateStorageBadge();
          throw new Error("GitHub-token är ogiltig eller har gått ut (401). Ange en ny token under ⚙ (samma som i 4D-planering).");
        }
        throw e;
      }
      rejectedTokens.add(t.token);
      updateStorageBadge();
    }
  }
}

function token() {
  const t = resolveToken();
  return t ? t.token : null;
}

function updateStorageBadge() {
  const badge = document.getElementById("storageBadge");
  const t = resolveToken();
  if (t && t.invalid) {
    badge.innerText = "Delad lagring – token ogiltig";
    showWarning("⚠️ GitHub-token är ogiltig eller har gått ut (401). Klicka ⚙ och ange en ny token (samma som i 4D-planering).");
  } else if (t) {
    badge.innerText = `Delad lagring (4D-data)${t.source === "4D-planering" ? " · token från 4D-planering" : ""}`;
    hideWarning();
  } else {
    badge.innerText = "Lokalt läge";
    showWarning("Lokalt läge – sparas bara i den här webbläsaren. Ange GitHub-token under ⚙ för att dela med projektet och koppla mot 4D-planering.");
  }
}

function openSettings() {
  // Tålig mot saknade fält, så att dialogen alltid öppnas.
  const set = (id, prop, value) => { const el = document.getElementById(id); if (el) el[prop] = value; };
  set("sUserName", "value", settings.userName || "");
  set("sToken", "value", settings.githubToken || "");
  set("sBubbleSize", "value", settings.bubbleSize);
  set("sBubbleSource", "value", settings.bubbleSource);
  set("sBubbleScale", "checked", settings.bubbleScale !== false);
  const dlg = document.getElementById("settingsDialog");
  dlg.returnValue = "";
  try {
    dlg.showModal();
  } catch (e) {
    dlg.setAttribute("open", "");
  }
}

async function onSettingsClosed() {
  const dlg = document.getElementById("settingsDialog");
  if (dlg.returnValue !== "save") return;
  const hadToken = token();
  rejectedTokens.clear();
  settings.userName = document.getElementById("sUserName").value.trim();
  settings.githubToken = document.getElementById("sToken").value.trim();
  settings.bubbleSize = Number(document.getElementById("sBubbleSize").value) || DEFAULT_BUBBLE_SIZE;
  settings.bubbleSource = document.getElementById("sBubbleSource").value;
  settings.bubbleScale = document.getElementById("sBubbleScale").checked;
  saveSettings();
  updateStorageBadge();
  scheduleBubbles();
  if (projectId) await refreshAll();
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
    items = await withToken(tok => ghReadJSON(tok, projectPath("field_notes.json")));
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
    items = await withToken(tok => ghWriteJSON(tok, projectPath("field_notes.json"), mutateFn, message));
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
  planItems = [];
  planById = new Map();
  planByObjectId = new Map();
  if (!token()) return;
  planItems = await withToken(tok => ghReadJSON(tok, projectPath("plan_items.json")));
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
  document.getElementById("btnPickPoint").onclick = () => setPickArmed(!pickArmed);
  document.getElementById("btnClearPin").onclick = () => { draftPin = null; renderDraftPin(); };
  document.getElementById("fTitle").onkeydown = (e) => { if (e.key === "Enter") onSave(); };

  document.getElementById("search").oninput = render;
  document.getElementById("hideDone").onchange = render;
  const showBubbles = document.getElementById("showBubbles");
  showBubbles.checked = !!settings.showBubbles;
  showBubbles.onchange = () => {
    settings.showBubbles = showBubbles.checked;
    // Ny normalvy varje gång bubblorna slås på - den vy man står i då är
    // den man vill läsa dem i.
    if (showBubbles.checked) settings.bubbleRefDistance = null;
    saveSettings();
    scheduleBubbles(0);
  };
  document.getElementById("btnBubbleNormal").onclick = async () => {
    try {
      const cam = await API.viewer.getCamera();
      if (cam && cam.position) bubbleState.cameraPos = cam.position;
    } catch (e) { /* använd senast kända kameraposition */ }
    if (!setReferenceDistance()) setBubbleHint("Inga bubblor att anpassa i den här vyn.");
  };
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
  draftPin = rec ? rec.pin || null : null;
  setPickArmed(false);
  renderDraftObjects();
  renderDraftPin();
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
  draftPin = null;
  setPickArmed(false);
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
    pin: draftPin,
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
    await withToken(tok => ghWriteJSON(tok, projectPath("plan_item_comments.json"), arr => [...arr, comment], "Ny kommentar (från Anteckningar)"));
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
  scheduleBubbles();
}

function renderCard(rec, isSelectedMatch) {
  const li = document.createElement("li");
  li.dataset.id = rec.id;
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
  if (rec.pin) chips.appendChild(chip("📍 Egen bubbelpunkt"));
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
   3D-bubblor
   ---------------------------------------------------------------------
   Trimbles text-markups har fast utseende (bara text + en färg). I stället
   ritas varje bubbla som en egen PNG på en canvas och läggs i 3D-vyn med
   viewer.addIcon(). Bilden är dubbelt så hög som själva bubblan och spetsen
   slutar exakt i bildens mitt, så att spetsen pekar på ankarpunkten även om
   Trimble centrerar ikonen på positionen.

   Ankarpunkt (meter, samma som getObjectBoundingBoxes och PointIcon):
     1. postens egen punkt (📍), annars
     2. mitten av ovansidan på de kopplade objektens gemensamma bounding box,
        annars
     3. det kopplade 4D-objektet.
   Poster med samma ankarpunkt slås ihop till en bubbla med "+N".
   ------------------------------------------------------------------- */
const bubbleState = {
  icons: [],                // PointIcon[] som den här extensionen lagt till
  iconToRecs: new Map(),    // icon-id -> [post-id]
  nextId: 1,
  bboxCache: new Map(),     // "modelId|objektId" -> Box3 | null (finns inte i inläst modell)
  imageCache: new Map(),
  groups: [],               // [{pos, recs}] - senast beräknade ankarpunkter
  cameraPos: null,          // kamerans position (meter), för storlek efter avstånd
  signature: "",            // vad som visas just nu, för att slippa onödiga omritningar
  cameraTimer: null,
  timer: null,
  running: false,
  pending: false,
};

const BUBBLE_COLORS = { note: "#0b5fff", todo: "#d97706", overdue: "#c62828", done: "#1b7f3b" };

function scheduleBubbles(delay = 300) {
  clearTimeout(bubbleState.timer);
  bubbleState.timer = setTimeout(updateBubbles, delay);
}

async function updateBubbles() {
  if (!API) return;
  if (bubbleState.running) {
    bubbleState.pending = true;
    return;
  }
  bubbleState.running = true;
  try {
    const recs = settings.showBubbles ? filteredItems(activeTab) : [];
    const groups = await groupByAnchor(recs);
    bubbleState.groups = groups;
    if (groups.length && !bubbleState.cameraPos) {
      try {
        const cam = await API.viewer.getCamera();
        if (cam && cam.position) bubbleState.cameraPos = cam.position;
      } catch (e) { /* storlek utan avstånd */ }
    }
    await applyBubbleIcons();

    const placed = groups.reduce((n, g) => n + g.recs.length, 0);
    setBubbleHint(settings.showBubbles
      ? (recs.length === 0 ? "" : placed < recs.length
        ? `${placed} av ${recs.length} poster visas i 3D. Övriga saknar kopplade objekt/punkt, eller så finns objekten inte i den inlästa modellen.`
        : "")
      : "");
  } catch (e) {
    console.error("Kunde inte visa bubblor i 3D", e);
    setBubbleHint("Kunde inte visa bubblor i 3D: " + e.message);
  } finally {
    bubbleState.running = false;
    if (bubbleState.pending) {
      bubbleState.pending = false;
      scheduleBubbles(0);
    }
  }
}

/**
 * Lägger ut bubblorna för bubbleState.groups. Trimble ritar ikoner med fast
 * storlek på skärmen, så utan åtgärd blir en bubbla enorm i förhållande till
 * modellen när man zoomar ut. Därför räknas storleken om efter kamerans
 * avstånd till bubblan.
 *
 * Avståndet jämförs med ett referensavstånd ("normalvyn") i stället för med
 * fasta meter - då spelar det ingen roll vilken enhet/skala Trimble
 * rapporterar kamerans position i. Referensen sätts automatiskt från vyn när
 * bubblorna slås på, och kan sättas om med "Normalstorlek här".
 *
 * Blir bubblan för liten för att texten ska gå att läsa (mindre än 60 % av
 * grundstorleken) visas en liten nål i stället. Storleken avrundas i steg så
 * att vi inte ritar om vid varje liten kamerarörelse.
 */
async function applyBubbleIcons() {
  const base = settings.bubbleSize;
  const cam = bubbleState.cameraPos;
  const dist = g => Math.hypot(g.pos.x - cam.x, g.pos.y - cam.y, g.pos.z - cam.z);
  if (cam && bubbleState.groups.length && !(settings.bubbleRefDistance > 0)) {
    setReferenceDistance(false);
  }
  const ref = settings.bubbleRefDistance;
  const scale = settings.bubbleScale !== false && cam && ref > 0;

  const specs = bubbleState.groups.map(g => {
    if (!scale) return { g, pin: false, size: base };
    const f = ref / Math.max(ref / 20, dist(g));
    if (f < 0.6) return { g, pin: true, size: Math.max(32, quantize(base * 0.45)) };
    return { g, pin: false, size: quantize(base * Math.min(1.5, f)) };
  });

  const signature = JSON.stringify([base, settings.bubbleSource, specs.map(sp =>
    [sp.g.pos.x, sp.g.pos.y, sp.g.pos.z, sp.pin, sp.size, sp.g.recs.map(r => [r.id, r.updated_at, r.done, (planById.get(String(r.plan_item_id)) || {}).status].join()).join()])]);
  if (signature === bubbleState.signature) return;

  const icons = [];
  const iconToRecs = new Map();
  specs.forEach(sp => {
    const id = bubbleState.nextId++;
    const iconPath = sp.pin ? pinImage(sp.g.recs, sp.size) : bubbleImage(sp.g.recs, sp.size);
    icons.push({ id, iconPath, position: sp.g.pos, size: sp.size });
    iconToRecs.set(id, sp.g.recs.map(r => r.id));
  });

  const old = bubbleState.icons;
  bubbleState.icons = icons;
  bubbleState.iconToRecs = iconToRecs;
  bubbleState.signature = signature;
  // Lägg till de nya innan de gamla tas bort, så att bubblorna inte blinkar.
  if (icons.length) await API.viewer.addIcon(icons);
  if (old.length) {
    try {
      await API.viewer.removeIcon(old);
    } catch (e) {
      console.warn("Kunde inte ta bort gamla bubblor", e);
    }
  }
}

/** Sätter referensavståndet (normalstorlek) till medianavståndet i nuvarande vy. */
function setReferenceDistance(redraw = true) {
  const cam = bubbleState.cameraPos;
  if (!cam || !bubbleState.groups.length) return false;
  const d = bubbleState.groups
    .map(g => Math.hypot(g.pos.x - cam.x, g.pos.y - cam.y, g.pos.z - cam.z))
    .sort((a, b) => a - b);
  const median = d[Math.floor(d.length / 2)];
  if (!(median > 0)) return false;
  settings.bubbleRefDistance = median;
  saveSettings();
  if (redraw) {
    bubbleState.signature = "";
    applyBubbleIcons().catch(e => console.warn(e));
  }
  return true;
}

function quantize(size) {
  return Math.max(8, Math.round(size / 8) * 8);
}

function onCameraChanged(camera) {
  if (!camera || !camera.position) return;
  bubbleState.cameraPos = camera.position;
  if (!settings.showBubbles || settings.bubbleScale === false || !bubbleState.groups.length) return;
  if (bubbleState.cameraTimer) return;       // max en omräkning per 250 ms
  bubbleState.cameraTimer = setTimeout(async () => {
    bubbleState.cameraTimer = null;
    if (bubbleState.running) return;         // updateBubbles tar med nya kameran ändå
    try {
      await applyBubbleIcons();
    } catch (e) {
      console.warn("Kunde inte uppdatera bubblornas storlek", e);
    }
  }, 250);
}

function setBubbleHint(text) {
  const h = document.getElementById("bubbleHint");
  h.innerText = text;
  h.classList.toggle("hidden", !text);
}

/** Grupperar poster på gemensam ankarpunkt: [{pos, recs}]. */
async function groupByAnchor(recs) {
  const objsByRec = new Map();
  recs.forEach(rec => {
    if (rec.pin) return;
    let objs = rec.objects || [];
    const plan = rec.plan_item_id ? planById.get(String(rec.plan_item_id)) : null;
    if (objs.length === 0 && plan && plan.model_id && plan.object_id) {
      objs = [{ model_id: plan.model_id, object_id: String(plan.object_id) }];
    }
    if (objs.length) objsByRec.set(rec.id, objs);
  });
  await loadBoundingBoxes([].concat(...objsByRec.values()));

  const groups = new Map();
  recs.forEach(rec => {
    let pos = rec.pin || null;
    if (!pos && objsByRec.has(rec.id)) {
      const boxes = objsByRec.get(rec.id)
        .map(o => bubbleState.bboxCache.get(`${o.model_id}|${o.object_id}`))
        .filter(Boolean);
      if (boxes.length) {
        const min = { x: Math.min(...boxes.map(b => b.min.x)), y: Math.min(...boxes.map(b => b.min.y)) };
        const max = { x: Math.max(...boxes.map(b => b.max.x)), y: Math.max(...boxes.map(b => b.max.y)), z: Math.max(...boxes.map(b => b.max.z)) };
        pos = { x: (min.x + max.x) / 2, y: (min.y + max.y) / 2, z: max.z };
      }
    }
    if (!pos) return;
    const key = [pos.x, pos.y, pos.z].map(v => Math.round(v * 20)).join(",");
    if (!groups.has(key)) groups.set(key, { pos: { x: pos.x, y: pos.y, z: pos.z }, recs: [] });
    groups.get(key).recs.push(rec);
  });
  return [...groups.values()];
}

/** Hämtar bounding boxes för objekt som inte redan finns i cachen. */
async function loadBoundingBoxes(objs) {
  const byModel = {};
  objs.forEach(o => {
    const key = `${o.model_id}|${o.object_id}`;
    if (bubbleState.bboxCache.has(key)) return;
    bubbleState.bboxCache.set(key, null);
    (byModel[o.model_id] = byModel[o.model_id] || []).push(o.object_id);
  });
  for (const modelId of Object.keys(byModel)) {
    const objectIds = byModel[modelId];
    const runtimeIds = await convertToRuntimeIdsSafe(modelId, objectIds);
    const runtimeToObject = new Map();
    runtimeIds.forEach((rid, i) => { if (rid !== undefined && rid !== null) runtimeToObject.set(rid, objectIds[i]); });
    if (runtimeToObject.size === 0) continue;
    try {
      const boxes = await API.viewer.getObjectBoundingBoxes(modelId, [...runtimeToObject.keys()]);
      boxes.forEach(b => {
        const objectId = runtimeToObject.get(b.id);
        if (objectId !== undefined && b.boundingBox) bubbleState.bboxCache.set(`${modelId}|${objectId}`, b.boundingBox);
      });
    } catch (e) {
      console.warn("Kunde inte läsa objektens position", e);
    }
  }
}

function bubbleKind(rec) {
  if (rec.type !== "todo") return "note";
  if (rec.done) return "done";
  return rec.due_date && rec.due_date < todayIso() ? "overdue" : "todo";
}

/**
 * Trimble verkar cacha ikonen per bildadress (iconPath) - inklusive dess
 * storlek. Samma bild med ny storlek visades därför med ny storlek i ett
 * ögonblick och föll sedan tillbaka till den gamla. Varje storlek får därför
 * en egen, unik bildadress (se sizeTag/markSize).
 */
function cachedImage(key, draw) {
  const cache = bubbleState.imageCache;
  if (!cache.has(key)) {
    if (cache.size > 300) cache.clear();    // ~100 kB per bild - håll minnet nere
    cache.set(key, draw());
  }
  return cache.get(key);
}

function staticIconUrl(kind, size) {
  return new URL(`callouts/${kind}.png?s=${size}`, window.location.href).href;
}

/** Ritar en nästan osynlig pixel vars position beror på storleken -> unik bild per storlek. */
function markSize(canvas, size) {
  const ctx = canvas.getContext("2d");
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.fillStyle = "rgba(0,0,0,0.004)";
  ctx.fillRect(size % canvas.width, canvas.height - 1 - Math.floor(size / canvas.width) % 4, 1, 1);
}

function bubbleImage(recs, size) {
  const rec = recs[0];
  const kind = bubbleKind(rec);
  if (settings.bubbleSource === "static") {
    return staticIconUrl(kind, size);
  }
  const plan = rec.plan_item_id ? planById.get(String(rec.plan_item_id)) : null;
  let sub;
  if (rec.type === "todo") {
    sub = [
      rec.done ? "Klar" : rec.due_date ? (kind === "overdue" ? `Försenad · ${rec.due_date}` : `Förfaller ${rec.due_date}`) : "",
      rec.assignee,
    ].filter(Boolean).join(" · ");
  } else {
    sub = [rec.author, formatDateTime(rec.updated_at || rec.created_at).slice(0, 10)].filter(Boolean).join(" · ");
  }
  const spec = {
    kind,
    title: rec.title || (rec.body || "").split("\n")[0] || "(utan rubrik)",
    sub,
    plan: plan ? [plan.object_name || plan.object_id, STATUS_LABELS[plan.status] || plan.status].filter(Boolean).join(" · ") : "",
    planLate: !!(plan && plan.status === "forsenad"),
    more: recs.length - 1,
  };
  return cachedImage(JSON.stringify([spec, size]), () => drawBubble(spec, size));
}

function pinImage(recs, size) {
  const kind = bubbleKind(recs[0]);
  if (settings.bubbleSource === "static") {
    return staticIconUrl(kind, size);
  }
  return cachedImage(`pin:${kind}:${size}`, () => drawPin(kind, size));
}

function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

/** Radbryter text på ord, max `maxLines` rader (sista raden kortas med …). */
function wrapText(font, text, maxWidth, maxLines) {
  const ctx = document.createElement("canvas").getContext("2d");
  ctx.font = font;
  const words = String(text).split(/\s+/).filter(Boolean);
  const out = [];
  let line = "";
  for (let i = 0; i < words.length; i++) {
    const test = line ? line + " " + words[i] : words[i];
    if (ctx.measureText(test).width <= maxWidth || !line) {
      line = test;
      continue;
    }
    if (out.length === maxLines - 1) {
      line = words.slice(i - line.split(" ").length, words.length).join(" ");
      break;
    }
    out.push(line);
    line = words[i];
  }
  out.push(fitText(ctx, line, maxWidth));
  return out;
}

function fitText(ctx, text, maxWidth) {
  if (ctx.measureText(text).width <= maxWidth) return text;
  let t = text;
  while (t.length > 1 && ctx.measureText(t + "…").width > maxWidth) t = t.slice(0, -1);
  return t + "…";
}

/** Ikonsymbol: bock-ruta för att göra, textrader för anteckning. */
function drawGlyph(ctx, kind, cx, cy, r) {
  ctx.fillStyle = BUBBLE_COLORS[kind];
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.fill();
  ctx.strokeStyle = "#fff";
  ctx.lineWidth = r * 0.2;
  ctx.lineCap = "round";
  ctx.lineJoin = "round";
  ctx.beginPath();
  if (kind === "note") {
    [-0.38, 0, 0.38].forEach((dy, i) => {
      ctx.moveTo(cx - r * 0.45, cy + dy * r);
      ctx.lineTo(cx + r * (i === 2 ? 0.15 : 0.45), cy + dy * r);
    });
  } else if (kind === "overdue") {
    ctx.moveTo(cx, cy - r * 0.45);
    ctx.lineTo(cx, cy + r * 0.1);
    ctx.moveTo(cx, cy + r * 0.42);
    ctx.lineTo(cx, cy + r * 0.43);
  } else {
    ctx.moveTo(cx - r * 0.4, cy + r * 0.02);
    ctx.lineTo(cx - r * 0.1, cy + r * 0.32);
    ctx.lineTo(cx + r * 0.42, cy - r * 0.3);
  }
  ctx.stroke();
}

/** Ritar en callout-bubbla och returnerar den som PNG-data-URL. */
/** Texturstorlek för en bubbla som visas i `size`: ~1,5x, som tvåpotens, 256-2048 px. */
function bubbleTextureSize(size) {
  const dpr = Math.max(1, Math.min(3, window.devicePixelRatio || 1));
  const target = (size || DEFAULT_BUBBLE_SIZE) * dpr * 1.5;
  let px = 256;
  while (px < target && px < 2048) px *= 2;
  return px;
}

function drawBubble(spec, sizeTag) {
  // Trimble ritar ikoner som kvadrater - en icke-kvadratisk bild trycks ihop
  // (smal text). Därför en kvadratisk canvas med bubblan i övre halvan och
  // spetsen exakt i mitten, på ankarpunkten.
  //
  // Upplösningen anpassas till visningsstorleken: en mycket större bild än
  // den visas i skalas ner av 3D-motorn utan utjämning, och då hackas texten
  // sönder (oläslig). Runt 1,5x visningsstorleken ger skarp text, avrundat
  // till en tvåpotens (säkrast för WebGL-texturer).
  const PX = bubbleTextureSize(sizeTag);
  const W = 300;
  const S = PX / W;
  const pad = 12, tip = 14, radius = 10;
  const TITLE = 26, LINE = 21;              // radhöjder (logiska enheter)
  const lines = [spec.sub, spec.plan].filter(Boolean);
  const titleFont = '700 20px "Segoe UI", Arial, sans-serif';
  const badgeW = spec.more > 0 ? 38 : 0;
  const textLeft = 4 + 6 + pad + 28;
  const titleLines = wrapText(titleFont, spec.title, (W - 4) - pad - badgeW - textLeft, 2);
  const H = pad * 2 + TITLE * titleLines.length + lines.length * LINE;
  const canvas = document.createElement("canvas");
  canvas.width = PX;
  canvas.height = PX;
  const ctx = canvas.getContext("2d");
  ctx.scale(S, S);
  const x = 4, y = W / 2 - tip - H, w = W - 8;
  const color = BUBBLE_COLORS[spec.kind];

  // Kropp + spets som en form, med skugga
  ctx.save();
  ctx.shadowColor = "rgba(0,0,0,0.28)";
  ctx.shadowBlur = 6;
  ctx.shadowOffsetY = 2;
  roundRect(ctx, x, y, w, H, radius);
  ctx.moveTo(W / 2 - tip, y + H);
  ctx.lineTo(W / 2, y + H + tip);
  ctx.lineTo(W / 2 + tip, y + H);
  ctx.fillStyle = "#ffffff";
  ctx.fill();
  ctx.restore();

  // Färgad vänsterkant
  ctx.save();
  roundRect(ctx, x, y, w, H, radius);
  ctx.clip();
  ctx.fillStyle = color;
  ctx.fillRect(x, y, 6, H);
  ctx.restore();

  // Ram (röd om 4D-objektet är försenat)
  roundRect(ctx, x, y, w, H, radius);
  ctx.strokeStyle = spec.planLate ? BUBBLE_COLORS.overdue : "rgba(0,0,0,0.12)";
  ctx.lineWidth = spec.planLate ? 2 : 1;
  ctx.stroke();

  const font = '"Segoe UI", Arial, sans-serif';
  const textX = textLeft;
  const titleY = y + pad + TITLE / 2;
  drawGlyph(ctx, spec.kind, x + 6 + pad + 11, titleY, 12);

  // Stora, feta typsnitt och full kontrast - bubblan visas ofta liten.
  ctx.textBaseline = "middle";
  ctx.fillStyle = "#000000";
  ctx.font = titleFont;
  titleLines.forEach((t, i) => {
    const ty = titleY + i * TITLE;
    ctx.fillText(t, textX, ty);
    if (spec.kind === "done") {
      ctx.strokeStyle = "#4b5563";
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(textX, ty + 1);
      ctx.lineTo(textX + ctx.measureText(t).width, ty + 1);
      ctx.stroke();
    }
  });

  ctx.font = `600 16px ${font}`;
  lines.forEach((line, i) => {
    const isPlan = line === spec.plan && i === lines.length - 1 && spec.plan;
    ctx.fillStyle = isPlan ? (spec.planLate ? "#b71c1c" : "#14532d")
      : spec.kind === "overdue" && i === 0 ? "#b71c1c" : "#1f2937";
    ctx.fillText(fitText(ctx, (isPlan ? "4D: " : "") + line, x + w - pad - textX), textX, y + pad + TITLE * titleLines.length + LINE / 2 + i * LINE);
  });

  if (spec.more > 0) {
    ctx.fillStyle = color;
    roundRect(ctx, x + w - pad - badgeW + 4, titleY - 12, badgeW - 4, 24, 12);
    ctx.fill();
    ctx.fillStyle = "#fff";
    ctx.font = `700 15px ${font}`;
    ctx.textAlign = "center";
    ctx.fillText(`+${spec.more}`, x + w - pad - badgeW / 2 + 2, titleY);
    ctx.textAlign = "left";
  }
  if (sizeTag) markSize(canvas, sizeTag);
  return canvas.toDataURL("image/png");
}

/** Enkel nål-ikon (för "Enkel ikon"-läget; genereras till docs/callouts/*.png). */
function drawPin(kind, sizeTag) {
  const size = 256;                         // kvadratisk, se drawBubble
  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  const ctx = canvas.getContext("2d");
  const cx = size / 2, r = 44, cy = size / 2 - 16 - r * 1.35; // spetsen i bildens mitt
  ctx.save();
  ctx.shadowColor = "rgba(0,0,0,0.3)";
  ctx.shadowBlur = 6;
  ctx.beginPath();
  ctx.arc(cx, cy, r, Math.PI * 0.8, Math.PI * 2.2);
  ctx.lineTo(cx, size / 2);
  ctx.closePath();
  ctx.fillStyle = "#fff";
  ctx.fill();
  ctx.restore();
  drawGlyph(ctx, kind, cx, cy, r - 6);
  if (sizeTag) markSize(canvas, sizeTag);
  return canvas.toDataURL("image/png");
}

/** Klick på en bubbla i 3D: visa och markera posten i listan. */
function onBubblePicked(data) {
  const icons = Array.isArray(data) ? data : data ? [data] : [];
  for (const icon of icons) {
    const recIds = bubbleState.iconToRecs.get(icon && icon.id);
    if (recIds && recIds.length) {
      focusRecords(recIds);
      return;
    }
  }
}

function focusRecords(recIds) {
  const first = items.find(r => r.id === recIds[0]);
  if (!first) return;
  if (first.type !== activeTab) switchTab(first.type);
  let firstCard = null;
  recIds.forEach(id => {
    const card = document.querySelector(`.card[data-id="${CSS.escape(id)}"]`);
    if (!card) return;
    card.classList.remove("flash");
    void card.offsetWidth;                  // starta om animationen
    card.classList.add("flash");
    firstCard = firstCard || card;
  });
  if (firstCard) firstCard.scrollIntoView({ behavior: "smooth", block: "center" });
}

/* ----- Egen punkt för bubblan (📍) ----- */
function setPickArmed(on) {
  pickArmed = on;
  const btn = document.getElementById("btnPickPoint");
  btn.classList.toggle("armed", on);
  btn.innerText = on ? "📍 Klicka i modellen… (avbryt)" : "📍 Välj punkt för 3D-bubbla";
}

function onPointPicked(detail) {
  if (!pickArmed || !detail || !detail.position) return;
  const p = detail.position;
  draftPin = { x: Number(p.x), y: Number(p.y), z: Number(p.z) };
  setPickArmed(false);
  renderDraftPin();
}

function renderDraftPin() {
  document.getElementById("fPinInfo").innerText = draftPin ? "✓ Punkt vald" : "";
  document.getElementById("btnClearPin").classList.toggle("hidden", !draftPin);
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
