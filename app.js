const STORAGE_KEY = "cutlist-optimizer-projects-v2";
const FIREBASE_CONFIG_KEY = "cutlist-optimizer-firebase-config-v1";
const EPSILON = 0.0001;
const INCH_TO_MM = 25.4;
const FOOT_TO_MM = 304.8;

const DEFAULTS = {
  modelUnits: "mm",
  unitScale: 1,
  thicknessQuarters: [4, 5, 6, 8],
  globalThicknessOverride: "auto",
  kerfMm: 3.2,
  pricePerBoardFoot: 9.5,
  defaultGrainLock: true,
  milling: {
    thicknessMm: 3.2,
    widthMm: 3.2,
    lengthMm: 25.4,
    boardEndTrimMm: 50.8,
    ripMarginMm: 1.6,
  },
  planningWidthsIn: [4, 6, 8, 10, 12],
  planningLengthsFt: [6, 8, 10],
  inventory: [
    { thicknessQuarter: 4, widthIn: 8, lengthFt: 8, quantity: 20 },
    { thicknessQuarter: 6, widthIn: 8, lengthFt: 10, quantity: 10 },
  ],
};

const state = {
  objText: "",
  rawParts: [],
  parts: [],
  partOverrides: {},
  planningResult: null,
  inventoryResult: null,
  viewer: null,
  sort: {
    key: "name",
    direction: "asc",
  },
  activeTab: "planner",
  firebase: {
    connected: false,
    mode: "local",
    app: null,
    auth: null,
    db: null,
    user: null,
    config: null,
  },
};

const dom = {
  tabPlanner: document.querySelector("#tab-planner"),
  tabSettings: document.querySelector("#tab-settings"),
  tabPanels: [...document.querySelectorAll(".tab-panel")],

  projectName: document.querySelector("#project-name"),
  saveProject: document.querySelector("#save-project"),
  newProject: document.querySelector("#new-project"),
  projectSelect: document.querySelector("#project-select"),
  loadProject: document.querySelector("#load-project"),
  deleteProject: document.querySelector("#delete-project"),
  syncStatusIcon: document.querySelector("#sync-status-icon"),
  storageModeIcon: document.querySelector("#storage-mode-icon"),

  firebaseApiKey: document.querySelector("#firebase-api-key"),
  firebaseAuthDomain: document.querySelector("#firebase-auth-domain"),
  firebaseProjectId: document.querySelector("#firebase-project-id"),
  firebaseAppId: document.querySelector("#firebase-app-id"),
  firebaseStorageBucket: document.querySelector("#firebase-storage-bucket"),
  firebaseMessagingSenderId: document.querySelector("#firebase-messaging-sender-id"),
  firebaseSaveConfig: document.querySelector("#firebase-save-config"),
  firebaseConnect: document.querySelector("#firebase-connect"),
  firebaseUseLocal: document.querySelector("#firebase-use-local"),
  firebaseStatus: document.querySelector("#firebase-status"),

  objFile: document.querySelector("#obj-file"),
  modelUnits: document.querySelector("#model-units"),
  unitScale: document.querySelector("#unit-scale"),
  thicknessOptions: document.querySelector("#thickness-options"),
  globalThicknessOverride: document.querySelector("#global-thickness-override"),
  applyThicknessOverride: document.querySelector("#apply-thickness-override"),
  clearPartOverrides: document.querySelector("#clear-part-overrides"),
  kerf: document.querySelector("#kerf"),
  pricePerBoardFoot: document.querySelector("#price-per-board-foot"),
  defaultGrainLock: document.querySelector("#default-grain-lock"),

  allowThickness: document.querySelector("#allow-thickness"),
  allowWidth: document.querySelector("#allow-width"),
  allowLength: document.querySelector("#allow-length"),
  boardEndTrim: document.querySelector("#board-end-trim"),
  ripMargin: document.querySelector("#rip-margin"),

  planningWidths: document.querySelector("#planning-widths"),
  planningLengths: document.querySelector("#planning-lengths"),
  inventoryInfinite: document.querySelector("#inventory-infinite"),

  inventoryBody: document.querySelector("#inventory-body"),
  addInventory: document.querySelector("#add-inventory"),
  inventoryRowTemplate: document.querySelector("#inventory-row-template"),

  analyze: document.querySelector("#analyze"),
  plan: document.querySelector("#plan"),
  inventoryPlan: document.querySelector("#inventory-plan"),
  status: document.querySelector("#status"),
  modelViewer: document.querySelector("#model-viewer"),
  resetView: document.querySelector("#reset-view"),

  partsSummary: document.querySelector("#parts-summary"),
  partsTable: document.querySelector("#parts-table"),
  partsHeaderSortables: [...document.querySelectorAll("#parts-table thead th[data-sort-key]")],
  partsTableBody: document.querySelector("#parts-table tbody"),

  planningSummary: document.querySelector("#planning-summary"),
  planningLayouts: document.querySelector("#planning-layouts"),
  inventorySummary: document.querySelector("#inventory-summary"),
  lumberYardSuggestions: document.querySelector("#lumber-yard-suggestions"),
  inventoryLayouts: document.querySelector("#inventory-layouts"),
};

init();

function init() {
  wireEvents();
  seedDefaultProjectInputs();
  restoreFirebaseConfigInputs();
  initTabs();
  initPartsSorting();
  updateStorageModeIndicator();
  updateSyncStatusIndicator();
  refreshProjectSelect();
  initModelViewer();
}

function wireEvents() {
  dom.tabPlanner.addEventListener("click", () => switchTab("planner"));
  dom.tabSettings.addEventListener("click", () => switchTab("settings"));
  dom.objFile.addEventListener("change", handleObjFile);
  dom.modelUnits.addEventListener("change", refreshViewerFromSettings);
  dom.unitScale.addEventListener("change", refreshViewerFromSettings);
  dom.thicknessOptions.addEventListener("input", syncGlobalThicknessOverrideOptions);
  dom.addInventory.addEventListener("click", () => addInventoryRow());
  dom.analyze.addEventListener("click", runAnalyze);
  dom.plan.addEventListener("click", runPlanning);
  dom.inventoryPlan.addEventListener("click", runInventoryPlan);
  dom.saveProject.addEventListener("click", saveProject);
  dom.loadProject.addEventListener("click", loadSelectedProject);
  dom.deleteProject.addEventListener("click", deleteSelectedProject);
  dom.newProject.addEventListener("click", clearProject);
  dom.applyThicknessOverride.addEventListener("click", applyGlobalOverrideToAllParts);
  dom.clearPartOverrides.addEventListener("click", clearAllPartOverrides);
  dom.resetView.addEventListener("click", resetViewerCamera);
  dom.firebaseSaveConfig.addEventListener("click", saveFirebaseConfigFromInputs);
  dom.firebaseConnect.addEventListener("click", connectFirebase);
  dom.firebaseUseLocal.addEventListener("click", useLocalStorageBackend);
  dom.inventoryInfinite.addEventListener("change", setInventoryQuantityMode);
}

function initTabs() {
  switchTab(state.activeTab);
}

function switchTab(tabName) {
  state.activeTab = tabName;
  dom.tabPlanner.classList.toggle("active", tabName === "planner");
  dom.tabSettings.classList.toggle("active", tabName === "settings");
  for (const panel of dom.tabPanels) {
    const panelTab = panel.getAttribute("data-tab");
    panel.classList.toggle("hidden", panelTab !== tabName);
  }
}

function initPartsSorting() {
  for (const th of dom.partsHeaderSortables) {
    th.addEventListener("click", () => {
      const key = th.dataset.sortKey;
      if (!key) {
        return;
      }
      if (state.sort.key === key) {
        state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
      } else {
        state.sort.key = key;
        state.sort.direction = "asc";
      }
      if (state.parts.length) {
        renderPartsTable(state.parts, collectInputs());
      }
    });
  }
  refreshSortHeaderStyles();
}

function seedDefaultProjectInputs() {
  dom.modelUnits.value = DEFAULTS.modelUnits;
  dom.unitScale.value = String(DEFAULTS.unitScale);
  dom.thicknessOptions.value = DEFAULTS.thicknessQuarters.join(",");
  syncGlobalThicknessOverrideOptions(DEFAULTS.globalThicknessOverride);
  dom.kerf.value = String(DEFAULTS.kerfMm);
  dom.pricePerBoardFoot.value = String(DEFAULTS.pricePerBoardFoot);
  dom.defaultGrainLock.checked = DEFAULTS.defaultGrainLock;

  dom.allowThickness.value = String(DEFAULTS.milling.thicknessMm);
  dom.allowWidth.value = String(DEFAULTS.milling.widthMm);
  dom.allowLength.value = String(DEFAULTS.milling.lengthMm);
  dom.boardEndTrim.value = String(DEFAULTS.milling.boardEndTrimMm);
  dom.ripMargin.value = String(DEFAULTS.milling.ripMarginMm);

  dom.planningWidths.value = DEFAULTS.planningWidthsIn.join(",");
  dom.planningLengths.value = DEFAULTS.planningLengthsFt.join(",");
  dom.inventoryInfinite.checked = true;

  dom.inventoryBody.innerHTML = "";
  for (const row of DEFAULTS.inventory) {
    addInventoryRow(row);
  }
  setInventoryQuantityMode();
}

async function handleObjFile(event) {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }
  state.objText = await file.text();
  const inputs = collectInputs();
  refreshViewerModel(state.objText, unitToMmFactor(inputs.modelUnits) * inputs.unitScale);
  setStatus(`Loaded ${file.name}. Click "Analyze Model" to parse parts.`, "ok");
}

function refreshViewerFromSettings() {
  if (!state.objText) {
    return;
  }
  const inputs = collectInputs();
  refreshViewerModel(state.objText, unitToMmFactor(inputs.modelUnits) * inputs.unitScale);
}

function clearProject() {
  state.objText = "";
  state.rawParts = [];
  state.parts = [];
  state.partOverrides = {};
  state.planningResult = null;
  state.inventoryResult = null;

  dom.projectName.value = "";
  dom.projectSelect.value = "";
  dom.objFile.value = "";

  seedDefaultProjectInputs();
  renderPartsSummary([]);
  renderPartsTable([], collectInputs());
  clearResults();
  clearViewerModel();
  switchTab("planner");
  setStatus("Started a new project with default settings.", "ok");
}

function clearResults() {
  dom.planningSummary.innerHTML = "";
  dom.planningLayouts.innerHTML = "";
  dom.inventorySummary.innerHTML = "";
  dom.inventoryLayouts.innerHTML = "";
  dom.lumberYardSuggestions.innerHTML = "";
}

function restoreFirebaseConfigInputs() {
  const config = readFirebaseConfigLocal();
  if (!config) {
    setFirebaseStatus("Firebase not connected.");
    return;
  }
  dom.firebaseApiKey.value = config.apiKey || "";
  dom.firebaseAuthDomain.value = config.authDomain || "";
  dom.firebaseProjectId.value = config.projectId || "";
  dom.firebaseAppId.value = config.appId || "";
  dom.firebaseStorageBucket.value = config.storageBucket || "";
  dom.firebaseMessagingSenderId.value = config.messagingSenderId || "";
  setFirebaseStatus("Firebase config loaded locally. Click Connect Firebase.");
}

function readFirebaseConfigFromInputs() {
  return {
    apiKey: dom.firebaseApiKey.value.trim(),
    authDomain: dom.firebaseAuthDomain.value.trim(),
    projectId: dom.firebaseProjectId.value.trim(),
    appId: dom.firebaseAppId.value.trim(),
    storageBucket: dom.firebaseStorageBucket.value.trim(),
    messagingSenderId: dom.firebaseMessagingSenderId.value.trim(),
  };
}

function validateFirebaseConfig(config) {
  const required = ["apiKey", "authDomain", "projectId", "appId"];
  const missing = required.filter((field) => !config[field]);
  return {
    ok: !missing.length,
    missing,
  };
}

function saveFirebaseConfigFromInputs() {
  const config = readFirebaseConfigFromInputs();
  const validation = validateFirebaseConfig(config);
  if (!validation.ok) {
    setFirebaseStatus(`Missing required Firebase fields: ${validation.missing.join(", ")}`, "error");
    return;
  }
  localStorage.setItem(FIREBASE_CONFIG_KEY, JSON.stringify(config));
  setFirebaseStatus("Firebase config saved locally for this browser.", "ok");
}

function readFirebaseConfigLocal() {
  try {
    const raw = localStorage.getItem(FIREBASE_CONFIG_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    console.error(error);
    return null;
  }
}

function setFirebaseStatus(message, type = "") {
  dom.firebaseStatus.className = `status ${type}`.trim();
  dom.firebaseStatus.textContent = message;
}

function activeStorageBackend() {
  return state.firebase.connected && state.firebase.mode === "firebase" ? "firebase" : "local";
}

function updateStorageModeIndicator() {
  if (!dom.storageModeIcon) {
    return;
  }
  const backend = activeStorageBackend();
  dom.storageModeIcon.classList.toggle("cloud", backend === "firebase");
  dom.storageModeIcon.classList.toggle("local", backend !== "firebase");
  dom.storageModeIcon.textContent = backend === "firebase" ? "C" : "L";
  dom.storageModeIcon.title =
    backend === "firebase"
      ? "Storage backend: Firebase Cloud"
      : "Storage backend: Local Browser Storage";
}

function updateSyncStatusIndicator() {
  if (!dom.syncStatusIcon) {
    return;
  }
  const healthy = state.firebase.connected && state.firebase.mode === "firebase";
  dom.syncStatusIcon.classList.toggle("green", healthy);
  dom.syncStatusIcon.classList.toggle("red", !healthy);
  dom.syncStatusIcon.title = healthy
    ? "Sync status: connected to Firebase cloud"
    : "Sync status: local mode or cloud disconnected";
}

async function connectFirebase() {
  if (!window.firebase) {
    setFirebaseStatus("Firebase scripts failed to load.", "error");
    return;
  }

  const config = readFirebaseConfigFromInputs();
  const validation = validateFirebaseConfig(config);
  if (!validation.ok) {
    setFirebaseStatus(`Missing required Firebase fields: ${validation.missing.join(", ")}`, "error");
    return;
  }

  try {
    saveFirebaseConfigFromInputs();
    const appName = `cutlist-${config.projectId}`;
    let app = window.firebase.apps.find((item) => item.name === appName);
    if (!app) {
      app = window.firebase.initializeApp(config, appName);
    }

    const auth = app.auth();
    await auth.signInAnonymously();
    const user = auth.currentUser;
    const db = app.firestore();

    state.firebase.connected = true;
    state.firebase.mode = "firebase";
    state.firebase.app = app;
    state.firebase.auth = auth;
    state.firebase.db = db;
    state.firebase.user = user;
    state.firebase.config = config;

    updateStorageModeIndicator();
    updateSyncStatusIndicator();
    await refreshProjectSelect();
    setFirebaseStatus(`Connected to Firebase project "${config.projectId}".`, "ok");
    setStatus("Switched project storage backend to Firebase cloud.", "ok");
  } catch (error) {
    console.error(error);
    state.firebase.connected = false;
    state.firebase.mode = "local";
    state.firebase.app = null;
    state.firebase.auth = null;
    state.firebase.db = null;
    state.firebase.user = null;
    state.firebase.config = null;
    updateStorageModeIndicator();
    updateSyncStatusIndicator();
    setFirebaseStatus(`Firebase connect failed: ${error.message || "Unknown error"}`, "error");
  }
}

async function useLocalStorageBackend() {
  state.firebase.mode = "local";
  updateStorageModeIndicator();
  updateSyncStatusIndicator();
  await refreshProjectSelect();
  setFirebaseStatus("Using local browser storage. Firebase connection is idle.");
  setStatus("Switched project storage backend to local browser storage.", "ok");
}

async function saveProject() {
  const name = (dom.projectName.value || "").trim();
  if (!name) {
    setStatus("Enter a project name before saving.", "error");
    return;
  }

  const data = collectInputs();
  data.partOverrides = state.partOverrides;

  const projects = await readProjectsActive();
  const existing = projects.find((project) => project.name === name);
  const ownerUid = state.firebase.user?.uid || null;
  const payload = {
    id: existing ? existing.id : crypto.randomUUID(),
    name,
    savedAt: new Date().toISOString(),
    objText: state.objText,
    ownerUid,
    inputs: data,
  };

  if (activeStorageBackend() === "firebase") {
    await saveProjectFirebase(payload);
  } else {
    saveProjectLocal(payload);
  }
  await refreshProjectSelect(payload.id);
  setStatus(
    `Saved project "${name}" to ${
      activeStorageBackend() === "firebase" ? "Firebase cloud" : "local browser storage"
    }.`,
    "ok"
  );
}

async function loadSelectedProject() {
  const selectedId = dom.projectSelect.value;
  if (!selectedId) {
    setStatus("Choose a saved project first.", "error");
    return;
  }

  const project = await readProjectByIdActive(selectedId);
  if (!project) {
    setStatus("Saved project could not be loaded.", "error");
    return;
  }

  dom.projectName.value = project.name || "";
  state.objText = project.objText || "";
  restoreInputs(project.inputs || {});
  runAnalyze();
  setStatus(`Loaded project "${project.name}".`, "ok");
}

async function deleteSelectedProject() {
  const selectedId = dom.projectSelect.value;
  if (!selectedId) {
    setStatus("Choose a saved project first.", "error");
    return;
  }

  const project = await readProjectByIdActive(selectedId);
  if (activeStorageBackend() === "firebase") {
    await deleteProjectFirebase(selectedId);
  } else {
    deleteProjectLocal(selectedId);
  }
  await refreshProjectSelect();
  if (project) {
    setStatus(`Deleted project "${project.name}".`, "ok");
  }
}

function readProjectsLocal() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (error) {
    console.error(error);
    return [];
  }
}

async function readProjectsFirebase() {
  if (!state.firebase.connected || !state.firebase.db) {
    return [];
  }
  const uid = state.firebase.user?.uid || null;
  let ref = state.firebase.db.collection("projects");
  if (uid) {
    ref = ref.where("ownerUid", "==", uid);
  }
  const snapshot = await ref.get();
  return snapshot.docs
    .map((doc) => ({ id: doc.id, ...doc.data() }))
    .sort((a, b) => (b.savedAt || "").localeCompare(a.savedAt || ""));
}

async function readProjectsActive() {
  if (activeStorageBackend() === "firebase") {
    return readProjectsFirebase();
  }
  return readProjectsLocal();
}

async function readProjectByIdActive(id) {
  if (activeStorageBackend() === "firebase") {
    if (!state.firebase.connected || !state.firebase.db) {
      return null;
    }
    const doc = await state.firebase.db.collection("projects").doc(id).get();
    if (!doc.exists) {
      return null;
    }
    const data = doc.data() || {};
    if (state.firebase.user?.uid && data.ownerUid && data.ownerUid !== state.firebase.user.uid) {
      return null;
    }
    return { id: doc.id, ...data };
  }
  return readProjectsLocal().find((item) => item.id === id) || null;
}

function saveProjectLocal(payload) {
  const projects = readProjectsLocal();
  const existing = projects.find((project) => project.id === payload.id);
  const next = existing
    ? projects.map((project) => (project.id === payload.id ? payload : project))
    : [...projects, payload];
  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
}

function deleteProjectLocal(id) {
  const projects = readProjectsLocal();
  const next = projects.filter((entry) => entry.id !== id);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
}

async function saveProjectFirebase(payload) {
  if (!state.firebase.connected || !state.firebase.db) {
    throw new Error("Firebase is not connected.");
  }
  await state.firebase.db.collection("projects").doc(payload.id).set(payload, { merge: true });
}

async function deleteProjectFirebase(id) {
  if (!state.firebase.connected || !state.firebase.db) {
    throw new Error("Firebase is not connected.");
  }
  await state.firebase.db.collection("projects").doc(id).delete();
}

async function refreshProjectSelect(selectedId = "") {
  let projects = [];
  try {
    projects = (await readProjectsActive()).sort((a, b) => (b.savedAt || "").localeCompare(a.savedAt || ""));
  } catch (error) {
    console.error(error);
    setStatus(`Could not load project list: ${error.message || "unknown error"}`, "error");
  }
  dom.projectSelect.innerHTML = '<option value="">Select a project...</option>';
  for (const project of projects) {
    const option = document.createElement("option");
    option.value = project.id;
    option.textContent = `${project.name} (${new Date(project.savedAt).toLocaleString()})`;
    dom.projectSelect.append(option);
  }
  if (selectedId) {
    dom.projectSelect.value = selectedId;
  }
}

function restoreInputs(inputs) {
  dom.modelUnits.value = inputs.modelUnits ?? DEFAULTS.modelUnits;
  dom.unitScale.value = String(inputs.unitScale ?? DEFAULTS.unitScale);

  const quarters = Array.isArray(inputs.thicknessOptionsQuarters)
    ? inputs.thicknessOptionsQuarters
    : DEFAULTS.thicknessQuarters;
  dom.thicknessOptions.value = quarters.join(",");
  syncGlobalThicknessOverrideOptions(inputs.globalThicknessOverride ?? "auto");

  dom.kerf.value = String(inputs.kerfMm ?? DEFAULTS.kerfMm);
  dom.pricePerBoardFoot.value = String(inputs.pricePerBoardFoot ?? DEFAULTS.pricePerBoardFoot);
  dom.defaultGrainLock.checked =
    typeof inputs.defaultGrainLock === "boolean"
      ? inputs.defaultGrainLock
      : DEFAULTS.defaultGrainLock;

  const milling = inputs.milling || {};
  dom.allowThickness.value = String(milling.thicknessMm ?? DEFAULTS.milling.thicknessMm);
  dom.allowWidth.value = String(milling.widthMm ?? DEFAULTS.milling.widthMm);
  dom.allowLength.value = String(milling.lengthMm ?? DEFAULTS.milling.lengthMm);
  dom.boardEndTrim.value = String(milling.boardEndTrimMm ?? DEFAULTS.milling.boardEndTrimMm);
  dom.ripMargin.value = String(milling.ripMarginMm ?? DEFAULTS.milling.ripMarginMm);

  dom.planningWidths.value = (
    Array.isArray(inputs.planningWidthsIn) ? inputs.planningWidthsIn : DEFAULTS.planningWidthsIn
  ).join(",");
  dom.planningLengths.value = (
    Array.isArray(inputs.planningLengthsFt) ? inputs.planningLengthsFt : DEFAULTS.planningLengthsFt
  ).join(",");
  dom.inventoryInfinite.checked =
    typeof inputs.inventoryInfinite === "boolean" ? inputs.inventoryInfinite : true;

  dom.inventoryBody.innerHTML = "";
  const inventory = Array.isArray(inputs.inventory) ? inputs.inventory : DEFAULTS.inventory;
  for (const row of inventory) {
    addInventoryRow(row);
  }
  setInventoryQuantityMode();

  state.partOverrides = typeof inputs.partOverrides === "object" && inputs.partOverrides
    ? inputs.partOverrides
    : {};
}

function collectInputs() {
  const thicknessOptionsQuarters = parseQuarterList(
    dom.thicknessOptions.value,
    DEFAULTS.thicknessQuarters
  );
  const globalThicknessOverride = parseGlobalOverride(
    dom.globalThicknessOverride.value,
    thicknessOptionsQuarters
  );

  return {
    modelUnits: dom.modelUnits.value || DEFAULTS.modelUnits,
    unitScale: getPositiveNumber(dom.unitScale.value, DEFAULTS.unitScale),
    thicknessOptionsQuarters,
    globalThicknessOverride,
    kerfMm: getNonNegativeNumber(dom.kerf.value, DEFAULTS.kerfMm),
    pricePerBoardFoot: getNonNegativeNumber(dom.pricePerBoardFoot.value, DEFAULTS.pricePerBoardFoot),
    defaultGrainLock: Boolean(dom.defaultGrainLock.checked),
    milling: {
      thicknessMm: getNonNegativeNumber(dom.allowThickness.value, DEFAULTS.milling.thicknessMm),
      widthMm: getNonNegativeNumber(dom.allowWidth.value, DEFAULTS.milling.widthMm),
      lengthMm: getNonNegativeNumber(dom.allowLength.value, DEFAULTS.milling.lengthMm),
      boardEndTrimMm: getNonNegativeNumber(dom.boardEndTrim.value, DEFAULTS.milling.boardEndTrimMm),
      ripMarginMm: getNonNegativeNumber(dom.ripMargin.value, DEFAULTS.milling.ripMarginMm),
    },
    planningWidthsIn: parseNumberList(dom.planningWidths.value, DEFAULTS.planningWidthsIn),
    planningLengthsFt: parseNumberList(dom.planningLengths.value, DEFAULTS.planningLengthsFt),
    inventoryInfinite: Boolean(dom.inventoryInfinite.checked),
    inventory: readInventoryRows(Boolean(dom.inventoryInfinite.checked)),
  };
}

function syncGlobalThicknessOverrideOptions(preferred = null) {
  const quarters = parseQuarterList(dom.thicknessOptions.value, DEFAULTS.thicknessQuarters);
  const current = preferred ?? dom.globalThicknessOverride.value ?? "auto";
  dom.globalThicknessOverride.innerHTML = "";

  const autoOption = document.createElement("option");
  autoOption.value = "auto";
  autoOption.textContent = "Auto";
  dom.globalThicknessOverride.append(autoOption);

  for (const quarter of quarters) {
    const option = document.createElement("option");
    option.value = String(quarter);
    option.textContent = `${quarter}/4 (${formatMm(quarterToMm(quarter), 1)})`;
    dom.globalThicknessOverride.append(option);
  }

  if (current === "auto" || quarters.includes(Number(current))) {
    dom.globalThicknessOverride.value = String(current);
  } else {
    dom.globalThicknessOverride.value = "auto";
  }
}

function parseGlobalOverride(raw, quarters) {
  if (raw === "auto") {
    return "auto";
  }
  const parsed = Number(raw);
  if (Number.isInteger(parsed) && quarters.includes(parsed)) {
    return parsed;
  }
  return "auto";
}

function addInventoryRow(initial = {}) {
  const fragment = dom.inventoryRowTemplate.content.cloneNode(true);
  const row = fragment.querySelector("tr");

  for (const [field, defaultValue] of [
    ["thicknessQuarter", 4],
    ["widthIn", 8],
    ["lengthFt", 8],
    ["quantity", ""],
  ]) {
    const input = row.querySelector(`input[data-field="${field}"]`);
    const value = initial[field] ?? defaultValue;
    input.value = value === "" || value == null ? "" : String(value);
  }

  row.querySelector('button[data-field="delete"]').addEventListener("click", () => {
    row.remove();
  });

  dom.inventoryBody.append(row);
  setInventoryQuantityMode();
}

function setInventoryQuantityMode() {
  const infinite = Boolean(dom.inventoryInfinite?.checked);
  for (const row of dom.inventoryBody.querySelectorAll("tr")) {
    const quantityInput = row.querySelector('input[data-field="quantity"]');
    if (!quantityInput) {
      continue;
    }
    quantityInput.disabled = infinite;
    quantityInput.placeholder = infinite ? "ignored (infinite)" : "unlimited";
  }
}

function readInventoryRows(inventoryInfinite = false) {
  const rows = [];
  for (const row of dom.inventoryBody.querySelectorAll("tr")) {
    const thicknessQuarter = Math.max(
      1,
      Math.floor(Number(row.querySelector('input[data-field="thicknessQuarter"]').value))
    );
    const widthIn = getPositiveNumber(
      row.querySelector('input[data-field="widthIn"]').value,
      null
    );
    const lengthFt = getPositiveNumber(
      row.querySelector('input[data-field="lengthFt"]').value,
      null
    );
    const quantityRaw = row.querySelector('input[data-field="quantity"]').value.trim();
    const quantity = inventoryInfinite
      ? null
      : quantityRaw === ""
        ? null
        : Math.max(1, Math.floor(Number(quantityRaw)));

    if (thicknessQuarter && widthIn && lengthFt) {
      rows.push({ thicknessQuarter, widthIn, lengthFt, quantity });
    }
  }
  return rows;
}

function runAnalyze() {
  const analysis = analyzeFromCurrentInputs({ clearExistingResults: true });
  if (!analysis) {
    return;
  }
  setStatus(`Analyzed ${state.parts.length} parts (${analysis.inputs.modelUnits} source -> metric output).`, "ok");
}

function runPlanning() {
  const analysis = analyzeFromCurrentInputs({ clearExistingResults: false });
  if (!analysis) {
    return;
  }

  const boardCatalog = buildPlanningCatalog(analysis.inputs);
  state.planningResult = optimizeCutPlan(state.parts, boardCatalog, analysis.inputs);

  renderPlanSummary(
    dom.planningSummary,
    state.planningResult,
    "Planning stock requirement",
    analysis.inputs.pricePerBoardFoot
  );
  renderLayouts(dom.planningLayouts, state.planningResult.boards);
  setStatus("Planning stock optimization completed.", "ok");
}

function runInventoryPlan() {
  const analysis = analyzeFromCurrentInputs({ clearExistingResults: false });
  if (!analysis) {
    return;
  }

  if (!analysis.inputs.inventory.length) {
    setStatus(
      "Add at least one lumber inventory row (size options) before recalculating inventory mode.",
      "error"
    );
    return;
  }

  const boardCatalog = buildInventoryCatalog(analysis.inputs.inventory);
  state.inventoryResult = optimizeCutPlan(state.parts, boardCatalog, analysis.inputs);

  renderPlanSummary(
    dom.inventorySummary,
    state.inventoryResult,
    "Inventory fit result",
    analysis.inputs.pricePerBoardFoot
  );
  renderLayouts(dom.inventoryLayouts, state.inventoryResult.boards);

  const suggestions = buildYardSuggestions(state.parts, analysis.inputs.inventory, analysis.inputs.kerfMm);
  renderYardSuggestions(suggestions);

  if (state.inventoryResult.unmetParts.length) {
    const unmetPlan = optimizeCutPlan(
      state.parts.filter((part) =>
        state.inventoryResult.unmetParts.some((unmet) => unmet.partId && unmet.partId === part.id)
      ),
      boardCatalog.map((item) => ({ ...item, quantity: null })),
      analysis.inputs
    );
    renderAdditionalNeeds(dom.inventorySummary, unmetPlan, analysis.inputs.pricePerBoardFoot);
  }

  setStatus("Lumber yard recalculation completed.", "ok");
}

function analyzeFromCurrentInputs({ clearExistingResults }) {
  if (!state.objText) {
    setStatus("Load an OBJ file before analyzing.", "error");
    return null;
  }

  const inputs = collectInputs();
  if (!inputs.thicknessOptionsQuarters.length) {
    setStatus("Add at least one stock thickness quarter option.", "error");
    return null;
  }

  const scaleToMm = unitToMmFactor(inputs.modelUnits) * inputs.unitScale;
  const parsed = parseObjObjects(state.objText, scaleToMm);
  refreshViewerModel(state.objText, scaleToMm);
  if (!parsed.length) {
    setStatus('No mesh objects found. Ensure OBJ has faces grouped with "o" or "g".', "error");
    return null;
  }

  state.rawParts = parsed;
  pruneOverridesToKnownParts(parsed);

  if (inputs.globalThicknessOverride !== "auto" && !hasExplicitPartOverrides()) {
    applyGlobalOverrideInMemory(inputs.globalThicknessOverride);
  }

  state.parts = assignPartsForStock(state.rawParts, inputs, state.partOverrides);

  renderPartsSummary(state.parts);
  renderPartsTable(state.parts, inputs);

  if (clearExistingResults) {
    state.planningResult = null;
    state.inventoryResult = null;
    clearResults();
  }

  return { inputs };
}

function hasExplicitPartOverrides() {
  return Object.values(state.partOverrides).some(
    (entry) => entry && (entry.thicknessOverrideQuarter != null || typeof entry.grainLock === "boolean")
  );
}

function applyGlobalOverrideInMemory(quarter) {
  for (const part of state.rawParts) {
    const current = state.partOverrides[part.id] || {};
    state.partOverrides[part.id] = {
      ...current,
      thicknessOverrideQuarter: quarter,
    };
  }
}

function applyGlobalOverrideToAllParts() {
  if (!state.rawParts.length) {
    setStatus("Analyze your model first, then apply the global override.", "error");
    return;
  }

  const override = parseGlobalOverride(
    dom.globalThicknessOverride.value,
    parseQuarterList(dom.thicknessOptions.value, DEFAULTS.thicknessQuarters)
  );

  for (const part of state.rawParts) {
    const current = state.partOverrides[part.id] || {};
    if (override === "auto") {
      delete current.thicknessOverrideQuarter;
      if (!Object.keys(current).length) {
        delete state.partOverrides[part.id];
      } else {
        state.partOverrides[part.id] = current;
      }
    } else {
      state.partOverrides[part.id] = {
        ...current,
        thicknessOverrideQuarter: override,
      };
    }
  }

  const inputs = collectInputs();
  state.parts = assignPartsForStock(state.rawParts, inputs, state.partOverrides);
  renderPartsSummary(state.parts);
  renderPartsTable(state.parts, inputs);
  clearResults();
  setStatus(
    override === "auto"
      ? "Cleared global thickness override from all parts."
      : `Applied ${override}/4 thickness override to all parts.`,
    "ok"
  );
}

function clearAllPartOverrides() {
  state.partOverrides = {};
  if (!state.rawParts.length) {
    setStatus("Cleared stored part overrides.", "ok");
    return;
  }
  const inputs = collectInputs();
  state.parts = assignPartsForStock(state.rawParts, inputs, state.partOverrides);
  renderPartsSummary(state.parts);
  renderPartsTable(state.parts, inputs);
  clearResults();
  setStatus("Cleared per-part grain and thickness overrides.", "ok");
}

function pruneOverridesToKnownParts(parts) {
  const ids = new Set(parts.map((part) => part.id));
  for (const key of Object.keys(state.partOverrides)) {
    if (!ids.has(key)) {
      delete state.partOverrides[key];
    }
  }
}

function parseObjObjects(text, scaleToMm) {
  const lines = text.split(/\r?\n/);
  const vertices = [null];
  const objectToVertexIndexes = new Map();
  let currentObject = "Unlabeled";

  ensureObject(currentObject);

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }

    if (line.startsWith("v ")) {
      const [, xs, ys, zs] = line.split(/\s+/);
      const x = Number(xs);
      const y = Number(ys);
      const z = Number(zs);
      if (Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(z)) {
        vertices.push([x * scaleToMm, y * scaleToMm, z * scaleToMm]);
      }
      continue;
    }

    if (line.startsWith("o ") || line.startsWith("g ")) {
      currentObject = line.slice(2).trim() || "Unlabeled";
      ensureObject(currentObject);
      continue;
    }

    if (line.startsWith("f ")) {
      const tokens = line.split(/\s+/).slice(1);
      const indexes = objectToVertexIndexes.get(currentObject);
      for (const token of tokens) {
        const rawIndex = token.split("/")[0];
        if (!rawIndex) {
          continue;
        }
        let index = Number(rawIndex);
        if (!Number.isInteger(index)) {
          continue;
        }
        if (index < 0) {
          index = vertices.length + index;
        }
        if (index > 0 && index < vertices.length) {
          indexes.add(index);
        }
      }
    }
  }

  const parts = [];
  let counter = 1;
  for (const [name, indexSet] of objectToVertexIndexes.entries()) {
    if (!indexSet.size) {
      continue;
    }

    const points = [];
    for (const index of indexSet) {
      const point = vertices[index];
      if (point) {
        points.push(point);
      }
    }
    if (!points.length) {
      continue;
    }

    const [x, y, z] = getBoundingBoxDimensions(points);
    const safeName = name || `Part ${counter}`;
    parts.push({
      id: `${slugify(safeName)}-${counter}`,
      name: safeName,
      xMm: x,
      yMm: y,
      zMm: z,
    });
    counter += 1;
  }

  return parts;

  function ensureObject(name) {
    if (!objectToVertexIndexes.has(name)) {
      objectToVertexIndexes.set(name, new Set());
    }
  }
}

function initModelViewer() {
  if (!dom.modelViewer || !window.THREE || !window.THREE.OrbitControls) {
    setStatus("3D preview unavailable (viewer library failed to load).", "error");
    return;
  }

  const renderer = new THREE.WebGLRenderer({
    canvas: dom.modelViewer,
    antialias: true,
    alpha: false,
  });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));

  const scene = new THREE.Scene();
  scene.background = new THREE.Color(0xf8f2e6);

  const camera = new THREE.PerspectiveCamera(45, 1, 1, 300000);
  camera.position.set(1200, 900, 1200);

  const controls = new THREE.OrbitControls(camera, renderer.domElement);
  controls.enableDamping = true;
  controls.dampingFactor = 0.08;
  controls.rotateSpeed = 0.85;
  controls.zoomSpeed = 0.9;
  controls.panSpeed = 0.8;
  controls.target.set(0, 0, 0);

  scene.add(new THREE.AmbientLight(0xffffff, 0.75));
  const keyLight = new THREE.DirectionalLight(0xffffff, 0.7);
  keyLight.position.set(1, 2, 1);
  scene.add(keyLight);
  const fillLight = new THREE.DirectionalLight(0xfff6df, 0.35);
  fillLight.position.set(-1.4, -0.8, 0.7);
  scene.add(fillLight);

  const modelGroup = new THREE.Group();
  scene.add(modelGroup);

  const grid = new THREE.GridHelper(4000, 40, 0x9f8a70, 0xd7c8b0);
  grid.position.y = 0;
  scene.add(grid);

  state.viewer = {
    renderer,
    scene,
    camera,
    controls,
    modelGroup,
    grid,
    originalCamera: null,
    observer: null,
    rafId: null,
  };

  const observer = new ResizeObserver(() => resizeViewer());
  observer.observe(dom.modelViewer.parentElement);
  state.viewer.observer = observer;

  const animate = () => {
    if (!state.viewer) {
      return;
    }
    state.viewer.controls.update();
    state.viewer.renderer.render(state.viewer.scene, state.viewer.camera);
    state.viewer.rafId = requestAnimationFrame(animate);
  };
  animate();
  resizeViewer();
}

function resizeViewer() {
  if (!state.viewer || !dom.modelViewer) {
    return;
  }
  const rect = dom.modelViewer.getBoundingClientRect();
  const width = Math.max(1, Math.floor(rect.width));
  const height = Math.max(1, Math.floor(rect.height));
  state.viewer.renderer.setSize(width, height, false);
  state.viewer.camera.aspect = width / height;
  state.viewer.camera.updateProjectionMatrix();
}

function clearViewerModel() {
  if (!state.viewer) {
    return;
  }
  while (state.viewer.modelGroup.children.length) {
    const child = state.viewer.modelGroup.children[0];
    state.viewer.modelGroup.remove(child);
    if (child.geometry) {
      child.geometry.dispose();
    }
    if (child.material) {
      if (Array.isArray(child.material)) {
        child.material.forEach((material) => material.dispose());
      } else {
        child.material.dispose();
      }
    }
  }
}

function refreshViewerModel(objText, scaleToMm) {
  if (!state.viewer) {
    return;
  }
  if (!objText) {
    clearViewerModel();
    return;
  }

  const meshes = parseObjMeshesForViewer(objText, scaleToMm);
  clearViewerModel();

  const palette = [
    0xa0592a,
    0xba7b43,
    0x7e8f46,
    0x6d597a,
    0x3a6b58,
    0x5f6e53,
    0x6b4f3f,
    0x43617a,
  ];

  meshes.forEach((meshDef, index) => {
    const geometry = new THREE.BufferGeometry();
    geometry.setAttribute(
      "position",
      new THREE.Float32BufferAttribute(meshDef.positions, 3)
    );
    geometry.computeVertexNormals();

    const material = new THREE.MeshStandardMaterial({
      color: palette[index % palette.length],
      roughness: 0.72,
      metalness: 0.03,
      side: THREE.DoubleSide,
    });
    const mesh = new THREE.Mesh(geometry, material);
    mesh.name = meshDef.name;
    state.viewer.modelGroup.add(mesh);
  });

  frameViewerOnModel();
}

function parseObjMeshesForViewer(text, scaleToMm) {
  const lines = text.split(/\r?\n/);
  const vertices = [null];
  const objects = new Map();
  let currentObject = "Unlabeled";
  ensureObject(currentObject);

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }

    if (line.startsWith("v ")) {
      const [, xs, ys, zs] = line.split(/\s+/);
      const x = Number(xs);
      const y = Number(ys);
      const z = Number(zs);
      if (Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(z)) {
        vertices.push([x * scaleToMm, y * scaleToMm, z * scaleToMm]);
      }
      continue;
    }

    if (line.startsWith("o ") || line.startsWith("g ")) {
      currentObject = line.slice(2).trim() || "Unlabeled";
      ensureObject(currentObject);
      continue;
    }

    if (line.startsWith("f ")) {
      const tokens = line.split(/\s+/).slice(1);
      const indexes = [];
      for (const token of tokens) {
        const rawIndex = token.split("/")[0];
        if (!rawIndex) {
          continue;
        }
        let index = Number(rawIndex);
        if (!Number.isInteger(index)) {
          continue;
        }
        if (index < 0) {
          index = vertices.length + index;
        }
        if (index > 0 && index < vertices.length) {
          indexes.push(index);
        }
      }

      if (indexes.length < 3) {
        continue;
      }

      const triangles = objects.get(currentObject);
      for (let i = 1; i < indexes.length - 1; i += 1) {
        triangles.push([indexes[0], indexes[i], indexes[i + 1]]);
      }
    }
  }

  const meshes = [];
  for (const [name, triangles] of objects.entries()) {
    if (!triangles.length) {
      continue;
    }

    const positions = [];
    for (const [a, b, c] of triangles) {
      for (const index of [a, b, c]) {
        const point = vertices[index];
        if (point) {
          positions.push(point[0], point[1], point[2]);
        }
      }
    }
    if (positions.length) {
      meshes.push({ name, positions });
    }
  }
  return meshes;

  function ensureObject(name) {
    if (!objects.has(name)) {
      objects.set(name, []);
    }
  }
}

function frameViewerOnModel() {
  if (!state.viewer) {
    return;
  }
  const { camera, controls, modelGroup } = state.viewer;
  const box = new THREE.Box3().setFromObject(modelGroup);
  if (box.isEmpty()) {
    camera.position.set(1200, 900, 1200);
    controls.target.set(0, 0, 0);
    controls.update();
    return;
  }

  const center = box.getCenter(new THREE.Vector3());
  const size = box.getSize(new THREE.Vector3());
  const maxDim = Math.max(size.x, size.y, size.z);
  const fitDistance = Math.max(400, maxDim / (2 * Math.tan((camera.fov * Math.PI) / 360)) * 1.4);

  camera.near = Math.max(1, maxDim / 1000);
  camera.far = Math.max(10000, fitDistance * 12);
  camera.position.set(center.x + fitDistance, center.y + fitDistance * 0.68, center.z + fitDistance);
  camera.updateProjectionMatrix();

  controls.target.copy(center);
  controls.update();
  state.viewer.originalCamera = {
    position: camera.position.clone(),
    target: controls.target.clone(),
  };
}

function resetViewerCamera() {
  if (!state.viewer) {
    return;
  }
  const { camera, controls, originalCamera } = state.viewer;
  if (!originalCamera) {
    frameViewerOnModel();
    return;
  }
  camera.position.copy(originalCamera.position);
  controls.target.copy(originalCamera.target);
  controls.update();
}

function getBoundingBoxDimensions(points) {
  let minX = Infinity;
  let maxX = -Infinity;
  let minY = Infinity;
  let maxY = -Infinity;
  let minZ = Infinity;
  let maxZ = -Infinity;

  for (const [x, y, z] of points) {
    minX = Math.min(minX, x);
    maxX = Math.max(maxX, x);
    minY = Math.min(minY, y);
    maxY = Math.max(maxY, y);
    minZ = Math.min(minZ, z);
    maxZ = Math.max(maxZ, z);
  }

  return [maxX - minX, maxY - minY, maxZ - minZ].map((value) => Number(value.toFixed(3)));
}

function assignPartsForStock(rawParts, inputs, partOverrides) {
  const quarters = [...inputs.thicknessOptionsQuarters].sort((a, b) => a - b);
  const oriented = [];

  for (const rawPart of rawParts) {
    const override = partOverrides[rawPart.id] || {};
    const overrideQuarter =
      override.thicknessOverrideQuarter == null ? null : Number(override.thicknessOverrideQuarter);
    const grainLock =
      typeof override.grainLock === "boolean" ? override.grainLock : inputs.defaultGrainLock;

    const canonical = canonicalizePartAxes(rawPart);
    const netLengthMm = canonical.x.value;
    const netWidthMm = canonical.y.value;
    const netThicknessMm = canonical.z.value;
    const roughLengthMm = netLengthMm + inputs.milling.lengthMm;
    const roughWidthMm = netWidthMm + inputs.milling.widthMm;
    const roughThicknessMm = netThicknessMm + inputs.milling.thicknessMm;
    const thicknessPlan = resolveStockPlan(roughThicknessMm, quarters, overrideQuarter);

    const base = {
      id: rawPart.id,
      name: rawPart.name,
      rawMm: { x: rawPart.xMm, y: rawPart.yMm, z: rawPart.zMm },
      netLengthMm: roundTo(netLengthMm, 2),
      netWidthMm: roundTo(netWidthMm, 2),
      netThicknessMm: roundTo(netThicknessMm, 2),
      roughLengthMm: roundTo(roughLengthMm, 2),
      roughWidthMm: roundTo(roughWidthMm, 2),
      roughThicknessMm: roundTo(roughThicknessMm, 2),
      orientation: `X<=${canonical.x.axis} (grain), Y<=${canonical.y.axis}, Z<=${canonical.z.axis}`,
      grainLock,
      thicknessOverrideQuarter: overrideQuarter,
    };

    if (!thicknessPlan.ok) {
      oriented.push({
        ...base,
        stockQuarter: null,
        stockThicknessMm: null,
        layers: 0,
        thicknessWasteMm: null,
        status: "invalid",
        reason:
          overrideQuarter != null
            ? `Override ${overrideQuarter}/4 is not available in thickness options.`
            : "No valid stock thickness options for this part.",
      });
      continue;
    }

    oriented.push({
      ...base,
      stockQuarter: thicknessPlan.stockQuarter,
      stockThicknessMm: roundTo(thicknessPlan.stockThicknessMm, 2),
      layers: thicknessPlan.layers,
      thicknessWasteMm: roundTo(thicknessPlan.wasteMm, 2),
      status: "ok",
      reason: "",
    });
  }

  return oriented.sort((a, b) => {
    if (a.stockQuarter == null && b.stockQuarter != null) {
      return 1;
    }
    if (a.stockQuarter != null && b.stockQuarter == null) {
      return -1;
    }
    const quarterA = a.stockQuarter ?? 0;
    const quarterB = b.stockQuarter ?? 0;
    return quarterA - quarterB || b.roughLengthMm * b.roughWidthMm - a.roughLengthMm * a.roughWidthMm;
  });
}

function canonicalizePartAxes(rawPart) {
  const dims = [
    { axis: "X", value: rawPart.xMm },
    { axis: "Y", value: rawPart.yMm },
    { axis: "Z", value: rawPart.zMm },
  ].sort((a, b) => b.value - a.value);

  return {
    x: dims[0],
    y: dims[1],
    z: dims[2],
  };
}

function resolveStockPlan(roughThicknessMm, quarters, overrideQuarter) {
  if (!quarters.length) {
    return { ok: false, reason: "No stock thickness options." };
  }

  if (overrideQuarter != null) {
    if (!quarters.includes(overrideQuarter)) {
      return { ok: false, reason: "Override not in thickness options." };
    }
    const stockThicknessMm = quarterToMm(overrideQuarter);
    const layers = Math.max(1, Math.ceil(roughThicknessMm / stockThicknessMm));
    return {
      ok: true,
      stockQuarter: overrideQuarter,
      stockThicknessMm,
      layers,
      wasteMm: layers * stockThicknessMm - roughThicknessMm,
    };
  }

  let best = null;
  for (const quarter of quarters) {
    const stockThicknessMm = quarterToMm(quarter);
    const layers = Math.max(1, Math.ceil(roughThicknessMm / stockThicknessMm));
    const wasteMm = layers * stockThicknessMm - roughThicknessMm;
    const candidate = { stockQuarter: quarter, stockThicknessMm, layers, wasteMm };
    if (!best) {
      best = candidate;
      continue;
    }

    if (candidate.layers < best.layers) {
      best = candidate;
      continue;
    }
    if (candidate.layers === best.layers && candidate.wasteMm < best.wasteMm - EPSILON) {
      best = candidate;
      continue;
    }
    if (
      candidate.layers === best.layers &&
      nearlyEqual(candidate.wasteMm, best.wasteMm) &&
      candidate.stockQuarter < best.stockQuarter
    ) {
      best = candidate;
    }
  }

  return { ok: true, ...best };
}

function buildPlanningCatalog(inputs) {
  const catalog = [];
  for (const quarter of inputs.thicknessOptionsQuarters) {
    for (const widthIn of inputs.planningWidthsIn) {
      for (const lengthFt of inputs.planningLengthsFt) {
        catalog.push({
          thicknessQuarter: quarter,
          widthIn,
          lengthFt,
          widthMm: widthIn * INCH_TO_MM,
          lengthMm: lengthFt * FOOT_TO_MM,
          quantity: null,
          source: "planner",
        });
      }
    }
  }
  return catalog;
}

function buildInventoryCatalog(inventoryRows) {
  return inventoryRows.map((row) => ({
    thicknessQuarter: row.thicknessQuarter,
    widthIn: row.widthIn,
    lengthFt: row.lengthFt,
    widthMm: row.widthIn * INCH_TO_MM,
    lengthMm: row.lengthFt * FOOT_TO_MM,
    quantity: row.quantity,
    source: "inventory",
  }));
}

function optimizeCutPlan(parts, boardCatalog, inputs) {
  const spacingMm = inputs.kerfMm + inputs.milling.ripMarginMm;
  const endTrimMm = inputs.milling.boardEndTrimMm;
  const unmetParts = [];

  const blanks = [];
  for (const part of parts) {
    if (part.status !== "ok" || !part.stockQuarter || part.layers < 1) {
      unmetParts.push({
        partId: part.id,
        partName: part.name,
        reason: part.reason || "Part is missing valid stock assignment.",
      });
      continue;
    }

    for (let layer = 1; layer <= part.layers; layer += 1) {
      blanks.push({
        id: `${part.id}-L${layer}`,
        partId: part.id,
        basePartName: part.name,
        name: part.layers > 1 ? `${part.name} (lam ${layer}/${part.layers})` : part.name,
        widthMm: part.roughWidthMm,
        lengthMm: part.roughLengthMm,
        stockQuarter: part.stockQuarter,
        grainLock: part.grainLock,
      });
    }
  }

  const grouped = groupBy(blanks, (blank) => String(blank.stockQuarter));
  const boards = [];
  let boardCounter = 1;

  for (const [quarterKey, group] of grouped.entries()) {
    const quarter = Number(quarterKey);
    const types = boardCatalog
      .filter((row) => row.thicknessQuarter === quarter)
      .map((row, index) => ({
        ...row,
        typeId: `${quarter}-${row.widthIn}-${row.lengthFt}-${index}`,
        remaining: row.quantity == null ? Infinity : row.quantity,
      }));

    if (!types.length) {
      for (const blank of group) {
        unmetParts.push({
          partId: blank.partId,
          partName: blank.name,
          reason: `No catalog/inventory board type for ${quarter}/4 stock.`,
        });
      }
      continue;
    }

    const sorted = [...group].sort(
      (a, b) =>
        Math.max(b.widthMm, b.lengthMm) - Math.max(a.widthMm, a.lengthMm) ||
        b.widthMm * b.lengthMm - a.widthMm * a.lengthMm
    );

    const openBoards = [];

    for (const blank of sorted) {
      let placement = findBestPlacementAcrossBoards(openBoards, blank, spacingMm);

      if (!placement) {
        const boardType = chooseBoardType(types, blank, spacingMm, endTrimMm);
        if (!boardType) {
          unmetParts.push({
            partId: blank.partId,
            partName: blank.name,
            reason: "No available board can fit this rough blank.",
          });
          continue;
        }

        boardType.remaining -= 1;
        const board = createBoardFromType(boardType, `B${boardCounter++}`, endTrimMm);
        openBoards.push(board);
        boards.push(board);
        placement = findBestPlacementOnBoard(board, blank, spacingMm);
      }

      if (!placement) {
        unmetParts.push({
          partId: blank.partId,
          partName: blank.name,
          reason: "Placement solver failed to place this blank.",
        });
        continue;
      }

      placeBlankOnBoard(placement.board, blank, placement, spacingMm);
    }
  }

  const boardUsage = new Map();
  for (const board of boards) {
    const key = boardKey(board);
    const prev = boardUsage.get(key) || {
      count: 0,
      boardFeetEach: boardFeetForBoard(board),
      thicknessQuarter: board.thicknessQuarter,
      widthIn: board.widthIn,
      lengthFt: board.lengthFt,
    };
    prev.count += 1;
    boardUsage.set(key, prev);
  }

  const stockAreaMm2 = sum(boards.map((board) => board.widthMm * board.lengthMm));
  const usedAreaMm2 = sum(
    boards.flatMap((board) => board.placements.map((placement) => placement.widthMm * placement.lengthMm))
  );
  const yieldPercent = stockAreaMm2 ? (usedAreaMm2 / stockAreaMm2) * 100 : 0;

  const totalBoardFeet = sum(boards.map((board) => boardFeetForBoard(board)));
  const estimatedCost = totalBoardFeet * inputs.pricePerBoardFoot;
  const stockVolumeM3 = sum(
    boards.map(
      (board) =>
        (quarterToMm(board.thicknessQuarter) * board.widthMm * board.lengthMm) / 1_000_000_000
    )
  );

  return {
    boards,
    unmetParts,
    boardUsage,
    stockAreaMm2,
    usedAreaMm2,
    yieldPercent,
    totalBoardFeet,
    estimatedCost,
    stockVolumeM3,
  };
}

function createBoardFromType(type, id, endTrimMm) {
  const trim = Math.max(0, Math.min(endTrimMm, Math.max(0, type.lengthMm - 1)));
  const trimOffset = trim / 2;
  const usableLengthMm = Math.max(0, type.lengthMm - trim);
  return {
    id,
    source: type.source,
    thicknessQuarter: type.thicknessQuarter,
    widthIn: type.widthIn,
    lengthFt: type.lengthFt,
    widthMm: type.widthMm,
    lengthMm: type.lengthMm,
    trimTotalMm: trim,
    trimOffsetMm: trimOffset,
    usableLengthMm,
    placements: [],
    freeRects: [{ x: 0, y: trimOffset, w: type.widthMm, h: usableLengthMm }],
  };
}

function chooseBoardType(types, blank, spacingMm, endTrimMm) {
  let best = null;
  for (const type of types) {
    if (type.remaining <= 0) {
      continue;
    }
    const usableLengthMm = type.lengthMm - endTrimMm;
    if (usableLengthMm <= EPSILON) {
      continue;
    }

    const options = buildBlankOrientationOptions(blank);
    let bestFitScore = null;
    for (const option of options) {
      if (option.widthMm + spacingMm > type.widthMm + EPSILON) {
        continue;
      }
      if (option.lengthMm + spacingMm > usableLengthMm + EPSILON) {
        continue;
      }
      const areaWaste = type.widthMm * usableLengthMm - option.widthMm * option.lengthMm;
      const sideWaste = (type.widthMm - option.widthMm) + (usableLengthMm - option.lengthMm);
      const score = areaWaste + sideWaste * 10;
      if (bestFitScore == null || score < bestFitScore) {
        bestFitScore = score;
      }
    }

    if (bestFitScore == null) {
      continue;
    }

    if (!best || bestFitScore < best.score) {
      best = { score: bestFitScore, type };
    }
  }
  return best?.type ?? null;
}

function findBestPlacementAcrossBoards(boards, blank, spacingMm) {
  let best = null;
  for (const board of boards) {
    const candidate = findBestPlacementOnBoard(board, blank, spacingMm);
    if (!candidate) {
      continue;
    }
    if (!best || comparePlacementScores(candidate, best) < 0) {
      best = candidate;
    }
  }
  return best;
}

function findBestPlacementOnBoard(board, blank, spacingMm) {
  let best = null;
  const options = buildBlankOrientationOptions(blank);

  for (let rectIndex = 0; rectIndex < board.freeRects.length; rectIndex += 1) {
    const rect = board.freeRects[rectIndex];
    for (const option of options) {
      const neededW = option.widthMm + spacingMm;
      const neededL = option.lengthMm + spacingMm;
      if (neededW > rect.w + EPSILON || neededL > rect.h + EPSILON) {
        continue;
      }

      const shortFit = Math.min(rect.w - neededW, rect.h - neededL);
      const longFit = Math.max(rect.w - neededW, rect.h - neededL);
      const areaFit = rect.w * rect.h - neededW * neededL;

      const candidate = {
        board,
        rectIndex,
        x: rect.x,
        y: rect.y,
        widthMm: option.widthMm,
        lengthMm: option.lengthMm,
        usedW: neededW,
        usedL: neededL,
        rotated: option.rotated,
        shortFit,
        longFit,
        areaFit,
      };

      if (!best || comparePlacementScores(candidate, best) < 0) {
        best = candidate;
      }
    }
  }

  return best;
}

function comparePlacementScores(a, b) {
  if (a.shortFit !== b.shortFit) {
    return a.shortFit - b.shortFit;
  }
  if (a.longFit !== b.longFit) {
    return a.longFit - b.longFit;
  }
  return a.areaFit - b.areaFit;
}

function placeBlankOnBoard(board, blank, placement) {
  const usedRect = { x: placement.x, y: placement.y, w: placement.usedW, h: placement.usedL };
  board.freeRects = splitFreeRects(board.freeRects, usedRect);
  board.freeRects = pruneContainedRects(board.freeRects).filter(
    (rect) => rect.w > EPSILON && rect.h > EPSILON
  );

  board.placements.push({
    blankId: blank.id,
    partId: blank.partId,
    partName: blank.name,
    x: placement.x,
    y: placement.y,
    widthMm: placement.widthMm,
    lengthMm: placement.lengthMm,
    rotated: placement.rotated,
    grainLock: blank.grainLock,
  });
}

function splitFreeRects(freeRects, usedRect) {
  const next = [];
  for (const rect of freeRects) {
    if (!rectanglesIntersect(rect, usedRect)) {
      next.push(rect);
      continue;
    }

    if (usedRect.x > rect.x + EPSILON) {
      next.push({ x: rect.x, y: rect.y, w: usedRect.x - rect.x, h: rect.h });
    }
    if (usedRect.x + usedRect.w < rect.x + rect.w - EPSILON) {
      next.push({
        x: usedRect.x + usedRect.w,
        y: rect.y,
        w: rect.x + rect.w - (usedRect.x + usedRect.w),
        h: rect.h,
      });
    }
    if (usedRect.y > rect.y + EPSILON) {
      next.push({ x: rect.x, y: rect.y, w: rect.w, h: usedRect.y - rect.y });
    }
    if (usedRect.y + usedRect.h < rect.y + rect.h - EPSILON) {
      next.push({
        x: rect.x,
        y: usedRect.y + usedRect.h,
        w: rect.w,
        h: rect.y + rect.h - (usedRect.y + usedRect.h),
      });
    }
  }
  return next;
}

function rectanglesIntersect(a, b) {
  return !(
    b.x >= a.x + a.w - EPSILON ||
    b.x + b.w <= a.x + EPSILON ||
    b.y >= a.y + a.h - EPSILON ||
    b.y + b.h <= a.y + EPSILON
  );
}

function buildBlankOrientationOptions(blank) {
  const options = [{ widthMm: blank.widthMm, lengthMm: blank.lengthMm, rotated: false }];
  if (!blank.grainLock && !nearlyEqual(blank.widthMm, blank.lengthMm)) {
    options.push({ widthMm: blank.lengthMm, lengthMm: blank.widthMm, rotated: true });
  }
  return options;
}

function pruneContainedRects(rects) {
  return rects.filter((rect, index) => {
    for (let i = 0; i < rects.length; i += 1) {
      if (i === index) {
        continue;
      }
      const other = rects[i];
      const contained =
        rect.x >= other.x - EPSILON &&
        rect.y >= other.y - EPSILON &&
        rect.x + rect.w <= other.x + other.w + EPSILON &&
        rect.y + rect.h <= other.y + other.h + EPSILON;
      if (contained) {
        return false;
      }
    }
    return true;
  });
}

function buildYardSuggestions(parts, inventory, kerfMm) {
  const suggestions = [];
  const longThresholdMm = 8 * FOOT_TO_MM;
  const carryTargetMm = 6 * FOOT_TO_MM;

  const byQuarter = groupBy(
    parts.filter((part) => part.status === "ok" && part.stockQuarter),
    (part) => String(part.stockQuarter)
  );

  for (const row of inventory) {
    const boardLengthMm = row.lengthFt * FOOT_TO_MM;
    if (boardLengthMm <= longThresholdMm + EPSILON) {
      continue;
    }

    const partsForQuarter = byQuarter.get(String(row.thicknessQuarter)) || [];
    if (!partsForQuarter.length) {
      continue;
    }

    const maxRequiredLenMm = Math.max(...partsForQuarter.map((part) => part.roughLengthMm));
    if (maxRequiredLenMm > carryTargetMm + EPSILON) {
      continue;
    }

    const segmentsMm = splitBoardLength(boardLengthMm, kerfMm, carryTargetMm);
    suggestions.push({
      row,
      maxRequiredLenMm,
      segmentsMm,
    });
  }

  return suggestions;
}

function splitBoardLength(totalLengthMm, kerfMm, targetMaxMm) {
  let cuts = Math.max(1, Math.ceil(totalLengthMm / targetMaxMm));
  while (cuts < 20) {
    const usable = totalLengthMm - (cuts - 1) * kerfMm;
    const segment = usable / cuts;
    if (segment <= targetMaxMm + EPSILON) {
      return Array.from({ length: cuts }, () => roundTo(segment, 2));
    }
    cuts += 1;
  }
  return [roundTo(totalLengthMm, 2)];
}

function renderYardSuggestions(suggestions) {
  if (!suggestions.length) {
    dom.lumberYardSuggestions.innerHTML =
      '<p class="muted">No long-board recut suggestions meet the less-than-6-foot carry target.</p>';
    return;
  }

  dom.lumberYardSuggestions.innerHTML = "<h3>Lumber Yard Recut Suggestions</h3>";
  const list = document.createElement("ul");
  list.className = "compact";

  for (const suggestion of suggestions) {
    const item = document.createElement("li");
    const segmentsText = suggestion.segmentsMm
      .map((segmentMm) => `${formatMm(segmentMm, 1)} (${formatFeet(mmToFeet(segmentMm), 2)})`)
      .join(" + ");
    item.textContent =
      `${suggestion.row.thicknessQuarter}/4 x ${formatInches(suggestion.row.widthIn)} x ${formatFeet(
        suggestion.row.lengthFt,
        1
      )}: max required part length is ${formatMm(suggestion.maxRequiredLenMm, 1)}. ` +
      `Suggested yard split: ${segmentsText}.`;
    list.append(item);
  }

  dom.lumberYardSuggestions.append(list);
}

function renderPartsSummary(parts) {
  const valid = parts.filter((part) => part.status === "ok").length;
  const invalid = parts.length - valid;
  const layers = sum(parts.map((part) => part.layers || 0));

  dom.partsSummary.innerHTML = "";
  for (const text of [
    `${parts.length} total parts detected`,
    `${valid} parts with valid stock assignments`,
    `${invalid} parts unassigned`,
    `${layers} total rough blanks including lamination layers`,
  ]) {
    const box = summaryBox(text);
    dom.partsSummary.append(box);
  }
}

function renderPartsTable(parts, inputs) {
  dom.partsTableBody.innerHTML = "";
  const quarters = inputs.thicknessOptionsQuarters;
  const sortedParts = getSortedParts(parts);
  refreshSortHeaderStyles();

  for (const part of sortedParts) {
    const row = document.createElement("tr");

    appendTextCell(
      row,
      part.name,
      `Raw X=${formatMm(part.rawMm.x)}, Raw Y=${formatMm(part.rawMm.y)}, Raw Z=${formatMm(part.rawMm.z)}`
    );
    appendTextCell(row, formatMm(part.netLengthMm));
    appendTextCell(row, formatMm(part.netWidthMm));
    appendTextCell(row, formatMm(part.netThicknessMm));
    appendTextCell(row, formatMm(part.roughLengthMm));
    appendTextCell(row, formatMm(part.roughWidthMm));
    appendTextCell(row, formatMm(part.roughThicknessMm));

    appendTextCell(
      row,
      part.stockQuarter
        ? `${part.stockQuarter}/4 (${formatMm(part.stockThicknessMm, 1)})`
        : `— ${part.reason}`
    );
    appendTextCell(row, part.layers ? String(part.layers) : "—");

    const grainCell = document.createElement("td");
    const grainInput = document.createElement("input");
    grainInput.type = "checkbox";
    grainInput.checked = Boolean(part.grainLock);
    grainInput.title = "Prevent 90° rotation in nesting";
    grainInput.addEventListener("change", () => {
      const current = state.partOverrides[part.id] || {};
      state.partOverrides[part.id] = { ...current, grainLock: grainInput.checked };
      updatePartsFromOverrides();
    });
    grainCell.append(grainInput);
    row.append(grainCell);

    const overrideCell = document.createElement("td");
    const overrideSelect = document.createElement("select");
    const autoOption = document.createElement("option");
    autoOption.value = "auto";
    autoOption.textContent = "Auto";
    overrideSelect.append(autoOption);

    for (const quarter of quarters) {
      const option = document.createElement("option");
      option.value = String(quarter);
      option.textContent = `${quarter}/4`;
      overrideSelect.append(option);
    }

    overrideSelect.value = part.thicknessOverrideQuarter == null ? "auto" : String(part.thicknessOverrideQuarter);
    overrideSelect.addEventListener("change", () => {
      const next = overrideSelect.value === "auto" ? null : Number(overrideSelect.value);
      const current = state.partOverrides[part.id] || {};
      if (next == null) {
        delete current.thicknessOverrideQuarter;
      } else {
        current.thicknessOverrideQuarter = next;
      }
      if (!Object.keys(current).length) {
        delete state.partOverrides[part.id];
      } else {
        state.partOverrides[part.id] = current;
      }
      updatePartsFromOverrides();
    });
    overrideCell.append(overrideSelect);
    row.append(overrideCell);

    appendTextCell(row, part.orientation, "", "orientation");
    dom.partsTableBody.append(row);
  }
}

function updatePartsFromOverrides() {
  if (!state.rawParts.length) {
    return;
  }
  const inputs = collectInputs();
  state.parts = assignPartsForStock(state.rawParts, inputs, state.partOverrides);
  renderPartsSummary(state.parts);
  renderPartsTable(state.parts, inputs);
  clearResults();
  setStatus("Part overrides updated.", "ok");
}

function appendTextCell(row, text, title = "", col = "") {
  const cell = document.createElement("td");
  cell.textContent = text;
  if (title) {
    cell.title = title;
  }
  if (col) {
    cell.dataset.col = col;
  }
  row.append(cell);
}

function getSortedParts(parts) {
  const { key, direction } = state.sort;
  const sign = direction === "asc" ? 1 : -1;
  const sorted = [...parts].sort((a, b) => {
    const av = getPartSortValue(a, key);
    const bv = getPartSortValue(b, key);

    if (typeof av === "number" && typeof bv === "number") {
      if (Number.isNaN(av) && Number.isNaN(bv)) {
        return 0;
      }
      if (Number.isNaN(av)) {
        return 1;
      }
      if (Number.isNaN(bv)) {
        return -1;
      }
      return (av - bv) * sign;
    }

    return String(av).localeCompare(String(bv), undefined, { sensitivity: "base" }) * sign;
  });
  return sorted;
}

function getPartSortValue(part, key) {
  switch (key) {
    case "name":
      return part.name || "";
    case "netLengthMm":
    case "netWidthMm":
    case "netThicknessMm":
    case "roughLengthMm":
    case "roughWidthMm":
    case "roughThicknessMm":
      return Number(part[key] ?? Number.NaN);
    case "stockQuarter":
      return part.stockQuarter == null ? Number.POSITIVE_INFINITY : Number(part.stockQuarter);
    case "layers":
      return Number(part.layers ?? Number.NaN);
    case "grainLock":
      return part.grainLock ? 1 : 0;
    case "thicknessOverrideQuarter":
      return part.thicknessOverrideQuarter == null
        ? Number.POSITIVE_INFINITY
        : Number(part.thicknessOverrideQuarter);
    case "orientation":
      return part.orientation || "";
    default:
      return part.name || "";
  }
}

function refreshSortHeaderStyles() {
  for (const th of dom.partsHeaderSortables) {
    const active = th.dataset.sortKey === state.sort.key;
    th.classList.toggle("sort-active", active);
    if (active) {
      th.title = `Sorted ${state.sort.direction} (click to toggle)`;
    } else {
      th.title = "Click to sort";
    }
  }
}

function renderPlanSummary(target, result, title, pricePerBoardFoot) {
  target.innerHTML = "";
  const root = document.createElement("div");
  root.className = "summary-grid";

  root.append(
    summaryBox(title),
    summaryBox(
      `${result.boards.length} boards used, ${formatNumber(result.totalBoardFeet, 2)} board feet total`
    ),
    summaryBox(`Estimated lumber cost: ${formatCurrency(result.estimatedCost)}`),
    summaryBox(
      `Used area: ${formatNumber(result.usedAreaMm2 / 1_000_000, 3)} m² of ${formatNumber(
        result.stockAreaMm2 / 1_000_000,
        3
      )} m² (${formatNumber(result.yieldPercent, 1)}% yield)`
    ),
    summaryBox(`Estimated stock volume: ${formatNumber(result.stockVolumeM3, 4)} m³`)
  );

  if (result.unmetParts.length) {
    root.append(summaryBox(`${result.unmetParts.length} parts are currently unmet`, "warning"));
  } else {
    root.append(summaryBox("All parts are successfully allocated.", "ok"));
  }

  const usageTable = document.createElement("table");
  usageTable.innerHTML =
    "<thead><tr><th>Stock Size</th><th>Board Count</th><th>Board Feet</th><th>Estimated Cost</th></tr></thead>";
  const usageBody = document.createElement("tbody");
  for (const [key, entry] of [...result.boardUsage.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    const lineBoardFeet = entry.boardFeetEach * entry.count;
    const lineCost = lineBoardFeet * pricePerBoardFoot;
    const row = document.createElement("tr");
    row.innerHTML = `<td>${key}</td><td>${entry.count}</td><td>${formatNumber(
      lineBoardFeet,
      2
    )}</td><td>${formatCurrency(lineCost)}</td>`;
    usageBody.append(row);
  }
  usageTable.append(usageBody);
  root.append(usageTable);

  if (result.unmetParts.length) {
    const header = document.createElement("h3");
    header.textContent = "Unmet Parts";
    root.append(header);
    const list = document.createElement("ul");
    list.className = "compact";
    for (const unmet of result.unmetParts) {
      const item = document.createElement("li");
      item.textContent = `${unmet.partName}: ${unmet.reason}`;
      list.append(item);
    }
    root.append(list);
  }

  target.append(root);
}

function renderAdditionalNeeds(target, additionalPlan, pricePerBoardFoot) {
  const block = document.createElement("div");
  block.className = "summary-box";
  if (!additionalPlan.boardUsage.size) {
    block.textContent = "Additional stock recommendation could not be generated for all unmet parts.";
    target.append(block);
    return;
  }

  const entries = [];
  let extraBoardFeet = 0;
  for (const [key, entry] of additionalPlan.boardUsage.entries()) {
    const bf = entry.boardFeetEach * entry.count;
    extraBoardFeet += bf;
    entries.push(`${entry.count} x ${key}`);
  }

  block.textContent =
    `Additional boards suggested for unmet parts: ${entries.join("; ")}. ` +
    `Extra board feet: ${formatNumber(extraBoardFeet, 2)}. ` +
    `Extra estimated cost: ${formatCurrency(extraBoardFeet * pricePerBoardFoot)}.`;
  target.append(block);
}

function renderLayouts(target, boards) {
  target.innerHTML = "";
  if (!boards.length) {
    target.innerHTML = '<p class="muted">No board layouts available.</p>';
    return;
  }

  const colors = [
    "#bc6c25",
    "#dda15e",
    "#606c38",
    "#283618",
    "#7f5539",
    "#9c6644",
    "#386641",
    "#1d3557",
    "#6d597a",
    "#2a9d8f",
  ];

  boards.forEach((board, boardIndex) => {
    const card = document.createElement("article");
    card.className = "board-card";

    const title = document.createElement("h4");
    title.textContent = `${board.id} • ${board.thicknessQuarter}/4 x ${formatInches(
      board.widthIn
    )} x ${formatFeet(board.lengthFt, 1)}`;
    card.append(title);

    const subtitle = document.createElement("p");
    subtitle.className = "muted";
    subtitle.textContent = `Metric: ${formatMm(board.widthMm, 1)} x ${formatMm(board.lengthMm, 1)} (${board.source})`;
    card.append(subtitle);

    const scale = 460 / board.widthMm;
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("class", "board-svg");
    svg.setAttribute("viewBox", `0 0 ${board.widthMm} ${board.lengthMm}`);
    svg.setAttribute("preserveAspectRatio", "none");
    svg.style.height = `${Math.max(140, board.lengthMm * scale)}px`;

    const boardRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    boardRect.setAttribute("x", "0");
    boardRect.setAttribute("y", "0");
    boardRect.setAttribute("width", String(board.widthMm));
    boardRect.setAttribute("height", String(board.lengthMm));
    boardRect.setAttribute("fill", "#f4e6ce");
    boardRect.setAttribute("stroke", "#a48a6a");
    boardRect.setAttribute("stroke-width", String(Math.max(0.8, board.widthMm * 0.002)));
    svg.append(boardRect);

    if (board.trimTotalMm > EPSILON) {
      const topTrim = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      topTrim.setAttribute("x", "0");
      topTrim.setAttribute("y", "0");
      topTrim.setAttribute("width", String(board.widthMm));
      topTrim.setAttribute("height", String(board.trimOffsetMm));
      topTrim.setAttribute("fill", "#d7c4a8");
      topTrim.setAttribute("fill-opacity", "0.45");
      svg.append(topTrim);

      const bottomTrim = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      bottomTrim.setAttribute("x", "0");
      bottomTrim.setAttribute("y", String(board.lengthMm - board.trimOffsetMm));
      bottomTrim.setAttribute("width", String(board.widthMm));
      bottomTrim.setAttribute("height", String(board.trimOffsetMm));
      bottomTrim.setAttribute("fill", "#d7c4a8");
      bottomTrim.setAttribute("fill-opacity", "0.45");
      svg.append(bottomTrim);
    }

    board.placements.forEach((placement, placementIndex) => {
      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.setAttribute("x", String(placement.x));
      rect.setAttribute("y", String(placement.y));
      rect.setAttribute("width", String(placement.widthMm));
      rect.setAttribute("height", String(placement.lengthMm));
      rect.setAttribute("fill", colors[(placementIndex + boardIndex) % colors.length]);
      rect.setAttribute("fill-opacity", "0.86");
      rect.setAttribute("stroke", "#ffffff");
      rect.setAttribute("stroke-width", String(Math.max(0.6, board.widthMm * 0.0015)));
      svg.append(rect);

      const label = document.createElementNS("http://www.w3.org/2000/svg", "text");
      label.setAttribute("x", String(placement.x + 3));
      label.setAttribute("y", String(placement.y + 10));
      label.setAttribute("font-size", String(Math.max(8, board.widthMm * 0.04)));
      label.setAttribute("fill", "#fff");
      label.textContent = shortenPartName(placement.partName);
      svg.append(label);
    });

    card.append(svg);

    const boardYield =
      board.placements.reduce((acc, item) => acc + item.widthMm * item.lengthMm, 0) /
      (board.widthMm * board.lengthMm);

    const caption = document.createElement("p");
    caption.className = "muted";
    caption.textContent = `${board.placements.length} blanks, ${formatNumber(boardYield * 100, 1)}% board yield`;
    card.append(caption);

    target.append(card);
  });
}

function boardKey(board) {
  return `${board.thicknessQuarter}/4 x ${formatInches(board.widthIn)} x ${formatFeet(board.lengthFt, 1)}`;
}

function boardFeetForBoard(board) {
  const thicknessIn = board.thicknessQuarter / 4;
  return (thicknessIn * board.widthIn * (board.lengthFt * 12)) / 144;
}

function summaryBox(text, kind = "default") {
  const box = document.createElement("div");
  box.className = "summary-box";
  if (kind === "warning") {
    box.style.background = "#fdf1e6";
    box.style.borderColor = "#dfc7a6";
  }
  if (kind === "ok") {
    box.style.background = "#ebf7ef";
    box.style.borderColor = "#b8d6c4";
  }
  box.textContent = text;
  return box;
}

function getPositiveNumber(value, fallback) {
  const num = Number(value);
  if (Number.isFinite(num) && num > 0) {
    return num;
  }
  return fallback;
}

function getNonNegativeNumber(value, fallback) {
  const num = Number(value);
  if (Number.isFinite(num) && num >= 0) {
    return num;
  }
  return fallback;
}

function parseNumberList(text, fallback = []) {
  const values = text
    .split(",")
    .map((chunk) => Number(chunk.trim()))
    .filter((num) => Number.isFinite(num) && num > 0);
  if (!values.length) {
    return [...fallback];
  }
  return [...new Set(values)].sort((a, b) => a - b);
}

function parseQuarterList(text, fallback = []) {
  const values = text
    .split(",")
    .map((chunk) => Math.max(1, Math.round(Number(chunk.trim()))))
    .filter((num) => Number.isInteger(num) && num > 0);
  if (!values.length) {
    return [...fallback];
  }
  return [...new Set(values)].sort((a, b) => a - b);
}

function unitToMmFactor(unit) {
  switch (unit) {
    case "mm":
      return 1;
    case "cm":
      return 10;
    case "m":
      return 1000;
    case "in":
      return INCH_TO_MM;
    case "ft":
      return FOOT_TO_MM;
    default:
      return 1;
  }
}

function quarterToMm(quarter) {
  return (quarter / 4) * INCH_TO_MM;
}

function mmToFeet(mm) {
  return mm / FOOT_TO_MM;
}

function groupBy(items, picker) {
  const map = new Map();
  for (const item of items) {
    const key = picker(item);
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(item);
  }
  return map;
}

function sum(values) {
  return values.reduce((acc, value) => acc + value, 0);
}

function nearlyEqual(a, b) {
  return Math.abs(a - b) <= EPSILON;
}

function roundTo(value, digits = 2) {
  const p = 10 ** digits;
  return Math.round(value * p) / p;
}

function formatNumber(value, digits = 2) {
  return Number(value).toFixed(digits);
}

function formatMm(value, digits = 1) {
  return `${formatNumber(value, digits)} mm`;
}

function formatInches(value, digits = 2) {
  return `${formatNumber(value, digits)}"`;
}

function formatFeet(value, digits = 1) {
  return `${formatNumber(value, digits)}'`;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value);
}

function shortenPartName(name) {
  return name.length > 20 ? `${name.slice(0, 17)}…` : name;
}

function slugify(text) {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "part";
}

function setStatus(message, type = "") {
  dom.status.className = `status ${type}`.trim();
  dom.status.textContent = message;
}
