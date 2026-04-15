// ─────────────────────────────────────────────────────────────────────────────
// Firebase default credentials
// Fill these in to enable auto-connect on startup.
// Firebase API keys are safe to embed in client-side code — they identify your
// project but do not grant access. Security is enforced by Firestore Security
// Rules and Authentication, not by keeping the API key secret.
// ─────────────────────────────────────────────────────────────────────────────
const DEFAULT_FIREBASE_CONFIG = {
  apiKey: "AIzaSyBWt2qIalVgZsu8ZPsoBY--I6N4qX1i6vM",
  authDomain: "cut-list-optimizer-4b4c5.firebaseapp.com",
  projectId: "cut-list-optimizer-4b4c5",
  appId: "1:520866072644:web:4e0ed71abd041b07f91411",
  storageBucket: "cut-list-optimizer-4b4c5.appspot.com",      // optional
  messagingSenderId: "585858585858",  // optional
};

// ─────────────────────────────────────────────────────────────────────────────
// Firebase App Check
// Protects your backend from abuse and unauthorized clients.
//
// Web (GitHub Pages):
//   1. Firebase Console → App Check → register your web app → reCAPTCHA v3
//   2. Google reCAPTCHA Admin (g.co/recaptcha/admin) → create v3 site key
//      for your domain (e.g. yourusername.github.io) AND localhost
//   3. Paste the site key below.
//
// App Check enforcement:
//   Start in MONITOR mode (Firebase Console → App Check → Apps → overflow menu).
//   Watch traffic for ~1 week before switching to ENFORCE.  Enforcing too early
//   can lock out legitimate users if something is misconfigured.
//
// Cordova (future):
//   Replace ReCaptchaV3Provider with the native Play Integrity / App Attest
//   provider in the Cordova build.  The web key is ignored by native builds.
//
// App Check is automatically skipped on localhost (reCAPTCHA v3 won't work there).
// Leave RECAPTCHA_SITE_KEY as "" to also skip it on deployed builds.
// ─────────────────────────────────────────────────────────────────────────────
const RECAPTCHA_SITE_KEY = "6LehgrcsAAAAABkZHwD2M9fngmCtW7YptYUj7UsG";

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────
const STORAGE_KEY        = "cutlist-optimizer-projects-v2";
const FIREBASE_CONFIG_KEY = "cutlist-optimizer-firebase-config-v1";
const EPSILON        = 0.0001;
const INCH_TO_MM     = 25.4;
const FOOT_TO_MM     = 304.8;
// Extra width added per panel strip to allow edge-jointing each glue face before assembly.
const PANEL_JOINT_MM = 3.2;

const DEFAULTS = {
  modelUnits: "cm",
  unitScale: 1,
  thicknessQuarters: [4, 5, 6, 8],
  globalThicknessOverride: "auto",
  kerfMm: 3.2,
  pricePerBoardFoot: 9.5,
  defaultGrainLock: true,
  maxPlanerWidthIn: 12, // 0 = no restriction
  milling: {
    thicknessMm: 3.2,
    widthMm: 3.2,
    lengthMm: 25.4,
    boardEndTrimMm: 50.8,
    ripMarginMm: 1.6,
  },
  planningWidthMinIn:  4,
  planningWidthMaxIn:  12,
  planningLengthMinFt: 6,
  planningLengthMaxFt: 10,
  // Default inventory shown only on very first load (new project starts empty)
  inventory: [
    { thicknessQuarter: 4, widthIn: 8, lengthFt: 8,  quantity: 20 },
    { thicknessQuarter: 6, widthIn: 8, lengthFt: 10, quantity: 10 },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// Application state
// ─────────────────────────────────────────────────────────────────────────────
const state = {
  objText: "",
  rawParts: [],
  parts: [],
  partOverrides: {},
  planningResult: null,
  inventoryResult: null,
  viewer: null,
  sort: { key: "name", direction: "asc" },
  activeTab: "planning",
  lastSavedSnapshot: null, // set on save/load to detect unsaved changes
  layoutScale: 0.75, // board diagram scale; 1.0 = full size, default 0.75 (25% smaller)
  firebase: {
    connected: false,
    mode: "local",
    app: null,
    auth: null,
    db: null,
    user: null,
    config: null,
    role: null,      // "standard" | "admin" | null
    authReady: false, // true after onAuthStateChanged fires once
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// DOM references
// ─────────────────────────────────────────────────────────────────────────────
const dom = {
  // Tabs
  tabPlanning:     document.querySelector("#tab-planning"),
  tabLumberYard:   document.querySelector("#tab-lumber-yard"),
  tabWorkshop:     document.querySelector("#tab-workshop"),
  tabSettings:     document.querySelector("#tab-settings"),
  tabInstructions: document.querySelector("#tab-instructions"),
  workshopSourceNote: document.querySelector("#workshop-source-note"),
  workshopContent:    document.querySelector("#workshop-content"),
  workshopPrint:      document.querySelector("#workshop-print"),
  tabContents:     [...document.querySelectorAll(".tab-content")],

  // Project
  projectSection: document.querySelector("#project-section"),
  projectName:  document.querySelector("#project-name"),
  saveProject:  document.querySelector("#save-project"),
  newProject:   document.querySelector("#new-project"),
  projectSelect: document.querySelector("#project-select"),
  loadProject:  document.querySelector("#load-project"),
  deleteProject: document.querySelector("#delete-project"),
  syncStatusIcon:  document.querySelector("#sync-status-icon"),

  // Status
  status: document.querySelector("#status"),

  // Admin tab
  firebaseStatus:    document.querySelector("#firebase-status"),
  adminRefreshUsers: document.querySelector("#admin-refresh-users"),
  adminUsersList:    document.querySelector("#admin-users-list"),

  // Model + Material settings
  objFile:            document.querySelector("#obj-file"),
  modelUnits:         document.querySelector("#model-units"),
  unitsDetectedNote:  document.querySelector("#units-detected-note"),
  unitScale:               document.querySelector("#unit-scale"),
  thicknessOptions:        document.querySelector("#thickness-options"),
  globalThicknessOverride: document.querySelector("#global-thickness-override"),
  applyThicknessOverride:  document.querySelector("#apply-thickness-override"),
  clearPartOverrides:      document.querySelector("#clear-part-overrides"),
  kerf:                    document.querySelector("#kerf"),
  pricePerBoardFoot:       document.querySelector("#price-per-board-foot"),
  defaultGrainLock:        document.querySelector("#default-grain-lock"),

  // Milling allowances
  allowThickness: document.querySelector("#allow-thickness"),
  allowWidth:     document.querySelector("#allow-width"),
  allowLength:    document.querySelector("#allow-length"),
  boardEndTrim:   document.querySelector("#board-end-trim"),
  ripMargin:         document.querySelector("#rip-margin"),
  maxPlanerWidth:    document.querySelector("#max-planer-width"),

  // Planning catalog range
  planningWidthMin:  document.querySelector("#planning-width-min"),
  planningWidthMax:  document.querySelector("#planning-width-max"),
  planningLengthMin: document.querySelector("#planning-length-min"),
  planningLengthMax: document.querySelector("#planning-length-max"),

  // Action buttons
  analyze:       document.querySelector("#analyze"),
  plan:          document.querySelector("#plan"),
  inventoryPlan: document.querySelector("#inventory-plan"),

  // Viewer
  modelViewer: document.querySelector("#model-viewer"),
  resetView:   document.querySelector("#reset-view"),

  // Parts
  partsSummary:        document.querySelector("#parts-summary"),
  partsTable:          document.querySelector("#parts-table"),
  partsHeaderSortables: [...document.querySelectorAll("#parts-table thead th[data-sort-key]")],
  partsTableBody:      document.querySelector("#parts-table tbody"),

  // Planning result
  planningSummary: document.querySelector("#planning-summary"),
  planningLayouts: document.querySelector("#planning-layouts"),

  // Inventory
  inventoryInfinite:    document.querySelector("#inventory-infinite"),
  inventoryBody:        document.querySelector("#inventory-body"),
  addInventory:         document.querySelector("#add-inventory"),
  recalcTipWrap:        document.querySelector("#recalc-tip-wrap"),
  inventoryRowTemplate: document.querySelector("#inventory-row-template"),

  // Lumber yard result
  inventorySummary:    document.querySelector("#inventory-summary"),
  lumberYardSuggestions: document.querySelector("#lumber-yard-suggestions"),
  inventoryLayouts:    document.querySelector("#inventory-layouts"),

  // Modal
  modal:       document.querySelector("#modal"),
  modalMsg:    document.querySelector("#modal-msg"),
  modalSave:   document.querySelector("#modal-save"),
  modalOk:     document.querySelector("#modal-ok"),
  modalCancel: document.querySelector("#modal-cancel"),

  // Auth modals
  authLoginModal:          document.querySelector("#auth-login-modal"),
  authLoginEmail:          document.querySelector("#auth-login-email"),
  authLoginPassword:       document.querySelector("#auth-login-password"),
  authLoginSubmit:         document.querySelector("#auth-login-submit"),
  authLoginCancel:         document.querySelector("#auth-login-cancel"),
  authLoginError:          document.querySelector("#auth-login-error"),
  authShowReset:           document.querySelector("#auth-show-reset"),
  authShowSignup:          document.querySelector("#auth-show-signup"),

  authSignupModal:         document.querySelector("#auth-signup-modal"),
  authSignupEmail:         document.querySelector("#auth-signup-email"),
  authSignupPassword:      document.querySelector("#auth-signup-password"),
  authSignupConfirm:       document.querySelector("#auth-signup-confirm"),
  authSignupSubmit:        document.querySelector("#auth-signup-submit"),
  authSignupCancel:        document.querySelector("#auth-signup-cancel"),
  authSignupError:         document.querySelector("#auth-signup-error"),
  authShowLoginFromSignup: document.querySelector("#auth-show-login-from-signup"),

  authResetModal:          document.querySelector("#auth-reset-modal"),
  authResetEmail:          document.querySelector("#auth-reset-email"),
  authResetSubmit:         document.querySelector("#auth-reset-submit"),
  authResetCancel:         document.querySelector("#auth-reset-cancel"),
  authResetMessage:        document.querySelector("#auth-reset-message"),

  // Mobile tab hamburger
  tabNavMobile:       document.querySelector("#tab-nav-mobile"),
  tabHamburger:       document.querySelector("#tab-hamburger"),
  tabHamburgerLabel:  document.querySelector("#tab-hamburger-label"),
  tabDropdown:        document.querySelector("#tab-dropdown"),

  // User menu (topbar)
  userMenuWrap:       document.querySelector("#user-menu-wrap"),
  userMenuBtn:        document.querySelector("#user-menu-btn"),
  userAvatarSilhouette: document.querySelector("#user-avatar-silhouette"),
  userAvatarInitials:   document.querySelector("#user-avatar-initials"),
  userMenuDropdown:   document.querySelector("#user-menu-dropdown"),
  userMenuInfo:       document.querySelector("#user-menu-info"),
  userMenuEmail:      document.querySelector("#user-menu-email"),
  userMenuRole:       document.querySelector("#user-menu-role"),
  userMenuDivider:    document.querySelector("#user-menu-divider"),
  userMenuSignin:     document.querySelector("#user-menu-signin"),
  userMenuSignout:    document.querySelector("#user-menu-signout"),
};

// ─────────────────────────────────────────────────────────────────────────────
// Bootstrap
// ─────────────────────────────────────────────────────────────────────────────
init();

function init() {
  wireEvents();
  seedDefaultProjectInputs();
  initTabs();
  initPartsSorting();
  updateSyncStatusIndicator();
  updateUserMenuUI();           // show unauthenticated state immediately
  updateSettingsTabVisibility(); // hide Settings tab until admin confirmed
  updateRecalcButtonState();    // ghost Recalculate until a plan exists
  refreshProjectSelect();
  initModelViewer();
  initAuth(); // async — sets up onAuthStateChanged listener
}

function wireEvents() {
  // Tabs
  dom.tabPlanning.addEventListener("click",     () => switchTab("planning"));
  dom.tabLumberYard.addEventListener("click",   () => switchTab("lumber-yard"));
  dom.tabWorkshop.addEventListener("click",     () => switchTab("workshop"));
  dom.tabSettings.addEventListener("click",     () => switchTab("settings"));
  dom.tabInstructions.addEventListener("click", () => switchTab("instructions"));

  // OBJ file + viewer settings
  dom.objFile.addEventListener("change", handleObjFile);
  dom.modelUnits.addEventListener("change", refreshViewerFromSettings);
  dom.unitScale.addEventListener("change", refreshViewerFromSettings);

  // Thickness options list — fix: wrap in arrow fn so no Event leaks as `preferred`
  dom.thicknessOptions.addEventListener("input", () => syncGlobalThicknessOverrideOptions());

  // Inventory
  dom.addInventory.addEventListener("click", () => addInventoryRow());
  dom.inventoryInfinite.addEventListener("change", setInventoryQuantityMode);

  // Action buttons
  dom.analyze.addEventListener("click",       runAnalyze);
  dom.plan.addEventListener("click",          runPlanning);
  dom.inventoryPlan.addEventListener("click", runInventoryPlan);
  dom.workshopPrint?.addEventListener("click", printWorkshopPDF);

  // Project management
  dom.saveProject.addEventListener("click",   saveProject);
  dom.loadProject.addEventListener("click",   loadSelectedProject);
  dom.deleteProject.addEventListener("click", deleteSelectedProject);
  dom.newProject.addEventListener("click",    clearProject);
  dom.projectSelect.addEventListener("change", updateProjectSelectButtons);

  // Override buttons
  dom.applyThicknessOverride.addEventListener("click", applyGlobalOverrideToAllParts);
  dom.clearPartOverrides.addEventListener("click",     clearAllPartOverrides);

  // Viewer
  dom.resetView.addEventListener("click", resetViewerCamera);

  // Admin tab
  dom.adminRefreshUsers.addEventListener("click", loadAdminUsersAndProjects);

  // Mobile tab hamburger
  dom.tabHamburger.addEventListener("click", (e) => {
    e.stopPropagation();
    const open = dom.tabDropdown.classList.toggle("hidden") === false;
    dom.tabHamburger.setAttribute("aria-expanded", String(open));
  });
  document.addEventListener("click", () => {
    if (!dom.tabDropdown.classList.contains("hidden")) {
      dom.tabDropdown.classList.add("hidden");
      dom.tabHamburger.setAttribute("aria-expanded", "false");
    }
  });
  dom.tabDropdown.querySelectorAll(".tab-dropdown-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      switchTab(btn.dataset.tabTarget);
      dom.tabDropdown.classList.add("hidden");
      dom.tabHamburger.setAttribute("aria-expanded", "false");
    });
  });

  // User menu toggle
  dom.userMenuBtn.addEventListener("click", toggleUserMenu);
  document.addEventListener("click", (e) => {
    if (!dom.userMenuWrap.contains(e.target)) closeUserMenu();
  });
  dom.userMenuSignin.addEventListener("click",  () => { closeUserMenu(); openLoginModal(); });
  dom.userMenuSignout.addEventListener("click", () => { closeUserMenu(); signOutUser(); });

  // Login modal
  dom.authLoginSubmit.addEventListener("click",   handleLogin);
  dom.authLoginCancel.addEventListener("click",   closeAllAuthModals);
  dom.authShowReset.addEventListener("click",     () => { closeAllAuthModals(); openResetModal(); });
  dom.authShowSignup.addEventListener("click",    () => { closeAllAuthModals(); openSignupModal(); });
  dom.authLoginPassword.addEventListener("keydown", (e) => { if (e.key === "Enter") handleLogin(); });

  // Signup modal
  dom.authSignupSubmit.addEventListener("click",  handleSignup);
  dom.authSignupCancel.addEventListener("click",  closeAllAuthModals);
  dom.authShowLoginFromSignup.addEventListener("click", () => { closeAllAuthModals(); openLoginModal(); });
  dom.authSignupConfirm.addEventListener("keydown", (e) => { if (e.key === "Enter") handleSignup(); });

  // Reset modal
  dom.authResetSubmit.addEventListener("click",  handlePasswordReset);
  dom.authResetCancel.addEventListener("click",  closeAllAuthModals);
  dom.authResetEmail.addEventListener("keydown", (e) => { if (e.key === "Enter") handlePasswordReset(); });

  // Modal
  dom.modalSave.addEventListener("click",   () => resolveModal("save"));
  dom.modalOk.addEventListener("click",     () => resolveModal(_modalMode === "unsaved" ? "discard" : true));
  dom.modalCancel.addEventListener("click", () => resolveModal(_modalMode === "unsaved" ? "cancel"  : false));
  dom.modal.addEventListener("keydown", (e) => {
    if (e.key === "Enter")  resolveModal(_modalMode === "unsaved" ? "discard" : true);
    if (e.key === "Escape") resolveModal(_modalMode === "unsaved" ? "cancel"  : false);
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Modal
// ─────────────────────────────────────────────────────────────────────────────
let _modalReturnFocus = null;
let _modalResolve     = null; // set when modal is in confirm/unsaved mode
let _modalMode        = "info"; // "info" | "confirm" | "unsaved"

// Info modal (one OK button)
function showModal(message, fromElement = null) {
  _modalMode = "info";
  dom.modalMsg.textContent      = message;
  dom.modalOk.textContent       = "OK";
  dom.modalSave.style.display   = "none";
  dom.modalCancel.style.display = "none";
  dom.modal.classList.remove("hidden");
  _modalReturnFocus = fromElement || document.activeElement;
  _modalResolve     = null;
  dom.modalOk.focus();
}

// Confirm modal — returns a Promise<boolean>
function showConfirm(message, okLabel = "Confirm", fromElement = null) {
  _modalMode = "confirm";
  dom.modalMsg.textContent      = message;
  dom.modalOk.textContent       = okLabel;
  dom.modalSave.style.display   = "none";
  dom.modalCancel.style.display = "";
  dom.modal.classList.remove("hidden");
  _modalReturnFocus = fromElement || document.activeElement;
  dom.modalOk.focus();
  return new Promise((resolve) => { _modalResolve = resolve; });
}

// 3-button unsaved-changes dialog — returns Promise<"save"|"discard"|"cancel">
function showUnsavedChangesDialog(fromElement = null) {
  _modalMode = "unsaved";
  const name = (dom.projectName.value || "").trim();
  dom.modalMsg.textContent      = name
    ? `"${name}" has unsaved changes.`
    : "You have unsaved changes.";
  dom.modalSave.textContent     = "Save";
  dom.modalSave.style.display   = "";
  dom.modalOk.textContent       = "Discard";
  dom.modalCancel.style.display = "";
  dom.modalCancel.textContent   = "Cancel";
  dom.modal.classList.remove("hidden");
  _modalReturnFocus = fromElement || document.activeElement;
  dom.modalSave.focus();
  return new Promise((resolve) => { _modalResolve = resolve; });
}

function resolveModal(result) {
  _modalMode = "info";
  dom.modalSave.style.display   = "none";
  dom.modalCancel.style.display = "none";
  dom.modal.classList.add("hidden");
  if (_modalReturnFocus) {
    try { _modalReturnFocus.focus(); } catch (_) {}
    _modalReturnFocus = null;
  }
  if (_modalResolve) {
    const cb = _modalResolve;
    _modalResolve = null;
    cb(result);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Tab management
// ─────────────────────────────────────────────────────────────────────────────
function initTabs() {
  switchTab(state.activeTab);
}

function switchTab(tabName) {
  state.activeTab = tabName;
  dom.tabPlanning.classList.toggle("active",     tabName === "planning");
  dom.tabLumberYard.classList.toggle("active",   tabName === "lumber-yard");
  dom.tabWorkshop.classList.toggle("active",     tabName === "workshop");
  dom.tabSettings.classList.toggle("active",     tabName === "settings");
  dom.tabInstructions.classList.toggle("active", tabName === "instructions");
  for (const content of dom.tabContents) {
    content.classList.toggle("hidden", content.getAttribute("data-tab") !== tabName);
  }

  // Sync mobile hamburger label and active item
  if (dom.tabHamburgerLabel) {
    const activeItem = dom.tabDropdown?.querySelector(`[data-tab-target="${tabName}"]`);
    dom.tabHamburgerLabel.textContent = activeItem?.textContent ?? tabName;
    dom.tabDropdown?.querySelectorAll(".tab-dropdown-item")
      .forEach((btn) => btn.classList.toggle("active", btn.dataset.tabTarget === tabName));
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Settings: seed / collect / restore
// ─────────────────────────────────────────────────────────────────────────────
function seedDefaultProjectInputs() {
  dom.modelUnits.value  = DEFAULTS.modelUnits;
  dom.unitScale.value   = String(DEFAULTS.unitScale);
  dom.thicknessOptions.value = DEFAULTS.thicknessQuarters.join(",");
  syncGlobalThicknessOverrideOptions(DEFAULTS.globalThicknessOverride);
  dom.kerf.value              = String(DEFAULTS.kerfMm);
  dom.pricePerBoardFoot.value = String(DEFAULTS.pricePerBoardFoot);
  dom.defaultGrainLock.checked = DEFAULTS.defaultGrainLock;

  dom.allowThickness.value = String(DEFAULTS.milling.thicknessMm);
  dom.allowWidth.value     = String(DEFAULTS.milling.widthMm);
  dom.allowLength.value    = String(DEFAULTS.milling.lengthMm);
  dom.boardEndTrim.value   = String(DEFAULTS.milling.boardEndTrimMm);
  dom.ripMargin.value      = String(DEFAULTS.milling.ripMarginMm);

  dom.planningWidthMin.value  = String(DEFAULTS.planningWidthMinIn);
  dom.planningWidthMax.value  = String(DEFAULTS.planningWidthMaxIn);
  dom.planningLengthMin.value = String(DEFAULTS.planningLengthMinFt);
  dom.planningLengthMax.value = String(DEFAULTS.planningLengthMaxFt);

  // Default inventory rows shown only on first/initial load
  dom.inventoryInfinite.checked = true;
  dom.inventoryBody.innerHTML = "";
  for (const row of DEFAULTS.inventory) {
    addInventoryRow(row);
  }
  setInventoryQuantityMode();
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
    unitScale:  getPositiveNumber(dom.unitScale.value, DEFAULTS.unitScale),
    thicknessOptionsQuarters,
    globalThicknessOverride,
    kerfMm:           getNonNegativeNumber(dom.kerf.value,              DEFAULTS.kerfMm),
    pricePerBoardFoot: getNonNegativeNumber(dom.pricePerBoardFoot.value, DEFAULTS.pricePerBoardFoot),
    defaultGrainLock: Boolean(dom.defaultGrainLock.checked),
    milling: {
      thicknessMm:  getNonNegativeNumber(dom.allowThickness.value, DEFAULTS.milling.thicknessMm),
      widthMm:      getNonNegativeNumber(dom.allowWidth.value,     DEFAULTS.milling.widthMm),
      lengthMm:     getNonNegativeNumber(dom.allowLength.value,    DEFAULTS.milling.lengthMm),
      boardEndTrimMm: getNonNegativeNumber(dom.boardEndTrim.value, DEFAULTS.milling.boardEndTrimMm),
      ripMarginMm:  getNonNegativeNumber(dom.ripMargin.value,      DEFAULTS.milling.ripMarginMm),
    },
    maxPlanerWidthIn: getNonNegativeNumber(dom.maxPlanerWidth.value, DEFAULTS.maxPlanerWidthIn),
    planningWidthMinIn:  getPositiveNumber(dom.planningWidthMin.value,  DEFAULTS.planningWidthMinIn),
    planningWidthMaxIn:  getPositiveNumber(dom.planningWidthMax.value,  DEFAULTS.planningWidthMaxIn),
    planningLengthMinFt: getPositiveNumber(dom.planningLengthMin.value, DEFAULTS.planningLengthMinFt),
    planningLengthMaxFt: getPositiveNumber(dom.planningLengthMax.value, DEFAULTS.planningLengthMaxFt),
    inventoryInfinite: Boolean(dom.inventoryInfinite.checked),
    inventory: readInventoryRows(Boolean(dom.inventoryInfinite.checked)),
    partOverrides: { ...state.partOverrides }, // snapshot — restored on project load
  };
}

function restoreInputs(inputs) {
  dom.modelUnits.value = inputs.modelUnits ?? DEFAULTS.modelUnits;
  dom.unitScale.value  = String(inputs.unitScale ?? DEFAULTS.unitScale);

  const quarters = Array.isArray(inputs.thicknessOptionsQuarters)
    ? inputs.thicknessOptionsQuarters
    : DEFAULTS.thicknessQuarters;
  dom.thicknessOptions.value = quarters.join(",");
  syncGlobalThicknessOverrideOptions(inputs.globalThicknessOverride ?? "auto");

  dom.kerf.value              = String(inputs.kerfMm              ?? DEFAULTS.kerfMm);
  dom.pricePerBoardFoot.value = String(inputs.pricePerBoardFoot   ?? DEFAULTS.pricePerBoardFoot);
  dom.defaultGrainLock.checked =
    typeof inputs.defaultGrainLock === "boolean" ? inputs.defaultGrainLock : DEFAULTS.defaultGrainLock;

  const milling = inputs.milling || {};
  dom.allowThickness.value = String(milling.thicknessMm  ?? DEFAULTS.milling.thicknessMm);
  dom.allowWidth.value     = String(milling.widthMm      ?? DEFAULTS.milling.widthMm);
  dom.allowLength.value    = String(milling.lengthMm     ?? DEFAULTS.milling.lengthMm);
  dom.boardEndTrim.value   = String(milling.boardEndTrimMm ?? DEFAULTS.milling.boardEndTrimMm);
  dom.ripMargin.value      = String(milling.ripMarginMm  ?? DEFAULTS.milling.ripMarginMm);
  dom.maxPlanerWidth.value = String(inputs.maxPlanerWidthIn ?? DEFAULTS.maxPlanerWidthIn);

  // Support legacy projects saved with planningWidthsIn / planningLengthsFt arrays
  if (Array.isArray(inputs.planningWidthsIn) && inputs.planningWidthsIn.length) {
    dom.planningWidthMin.value = String(Math.min(...inputs.planningWidthsIn));
    dom.planningWidthMax.value = String(Math.max(...inputs.planningWidthsIn));
  } else {
    dom.planningWidthMin.value = String(inputs.planningWidthMinIn  ?? DEFAULTS.planningWidthMinIn);
    dom.planningWidthMax.value = String(inputs.planningWidthMaxIn  ?? DEFAULTS.planningWidthMaxIn);
  }
  if (Array.isArray(inputs.planningLengthsFt) && inputs.planningLengthsFt.length) {
    dom.planningLengthMin.value = String(Math.min(...inputs.planningLengthsFt));
    dom.planningLengthMax.value = String(Math.max(...inputs.planningLengthsFt));
  } else {
    dom.planningLengthMin.value = String(inputs.planningLengthMinFt ?? DEFAULTS.planningLengthMinFt);
    dom.planningLengthMax.value = String(inputs.planningLengthMaxFt ?? DEFAULTS.planningLengthMaxFt);
  }
  // maxTransportLengthFt was merged into planningLengthMaxFt — silently ignored on old project loads

  dom.inventoryInfinite.checked =
    typeof inputs.inventoryInfinite === "boolean" ? inputs.inventoryInfinite : true;

  dom.inventoryBody.innerHTML = "";
  const inventory = Array.isArray(inputs.inventory) ? inputs.inventory : [];
  for (const row of inventory) {
    addInventoryRow(row);
  }
  setInventoryQuantityMode();

  state.partOverrides =
    typeof inputs.partOverrides === "object" && inputs.partOverrides
      ? inputs.partOverrides
      : {};
}

// ─────────────────────────────────────────────────────────────────────────────
// Global thickness override UI helpers
// ─────────────────────────────────────────────────────────────────────────────
function syncGlobalThicknessOverrideOptions(preferred = null) {
  const quarters = parseQuarterList(dom.thicknessOptions.value, DEFAULTS.thicknessQuarters);
  const current  = preferred ?? dom.globalThicknessOverride.value ?? "auto";
  dom.globalThicknessOverride.innerHTML = "";

  const autoOpt = document.createElement("option");
  autoOpt.value = "auto";
  autoOpt.textContent = "Auto";
  dom.globalThicknessOverride.append(autoOpt);

  for (const quarter of quarters) {
    const opt = document.createElement("option");
    opt.value = String(quarter);
    opt.textContent = `${quarter}/4 (${formatMm(quarterToMm(quarter), 1)})`;
    dom.globalThicknessOverride.append(opt);
  }

  if (current === "auto" || quarters.includes(Number(current))) {
    dom.globalThicknessOverride.value = String(current);
  } else {
    dom.globalThicknessOverride.value = "auto";
  }
}

function parseGlobalOverride(raw, quarters) {
  if (raw === "auto") return "auto";
  const parsed = Number(raw);
  if (Number.isInteger(parsed) && quarters.includes(parsed)) return parsed;
  return "auto";
}

// ─────────────────────────────────────────────────────────────────────────────
// Inventory rows
// ─────────────────────────────────────────────────────────────────────────────
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
    input.value = (value === "" || value == null) ? "" : String(value);
  }

  row.querySelector('button[data-field="delete"]').addEventListener("click", () => {
    row.remove();
    updateRecalcButtonState();
  });

  dom.inventoryBody.append(row);
  setInventoryQuantityMode();
}

function setInventoryQuantityMode() {
  const infinite = Boolean(dom.inventoryInfinite?.checked);
  for (const row of dom.inventoryBody.querySelectorAll("tr")) {
    const qty = row.querySelector('input[data-field="quantity"]');
    if (!qty) continue;
    qty.disabled = infinite;
    qty.placeholder = infinite ? "ignored (infinite)" : "unlimited";
  }
  updateRecalcButtonState();
}

function readInventoryRows(inventoryInfinite = false) {
  const rows = [];
  for (const row of dom.inventoryBody.querySelectorAll("tr")) {
    const thicknessQuarter = Math.max(
      1,
      Math.floor(Number(row.querySelector('input[data-field="thicknessQuarter"]').value))
    );
    const widthIn  = getPositiveNumber(row.querySelector('input[data-field="widthIn"]').value,  null);
    const lengthFt = getPositiveNumber(row.querySelector('input[data-field="lengthFt"]').value, null);
    const qRaw     = row.querySelector('input[data-field="quantity"]').value.trim();
    const quantity = inventoryInfinite
      ? null
      : qRaw === ""
        ? null
        : Math.max(1, Math.floor(Number(qRaw)));

    if (thicknessQuarter && widthIn && lengthFt) {
      rows.push({ thicknessQuarter, widthIn, lengthFt, quantity });
    }
  }
  return rows;
}

// ─────────────────────────────────────────────────────────────────────────────
// Project management
// ─────────────────────────────────────────────────────────────────────────────

// ── Dirty-state tracking ─────────────────────────────────────────────────────

function makeProjectSnapshot() {
  return {
    name:          (dom.projectName.value || "").trim(),
    objTextLen:    state.objText.length,
    partOverrides: JSON.stringify(state.partOverrides),
  };
}

function markProjectClean() {
  state.lastSavedSnapshot = makeProjectSnapshot();
}

function isProjectDirty() {
  // Completely blank slate — nothing to lose
  const hasContent = state.objText ||
                     (dom.projectName.value || "").trim() ||
                     Object.keys(state.partOverrides).length;
  if (!hasContent) return false;

  // Content exists but was never saved/loaded — always dirty
  if (!state.lastSavedSnapshot) return true;

  const snap = state.lastSavedSnapshot;
  const cur  = makeProjectSnapshot();
  return cur.name !== snap.name ||
         cur.objTextLen !== snap.objTextLen ||
         cur.partOverrides !== snap.partOverrides;
}

function updateProjectSelectButtons() {
  const hasSelection = !!dom.projectSelect.value;
  dom.loadProject.disabled   = !hasSelection;
  dom.deleteProject.disabled = !hasSelection;
}

async function clearProject() {
  if (isProjectDirty()) {
    const choice = await showUnsavedChangesDialog(dom.newProject);
    if (choice === "cancel") return;
    if (choice === "save") {
      await saveProject();
      // If save failed (no name, network error, etc.) don't wipe the project
      if (isProjectDirty()) return;
    }
    // choice === "discard" — fall through and wipe
  }

  state.objText           = "";
  state.rawParts          = [];
  state.parts             = [];
  state.partOverrides     = {};
  state.planningResult    = null;
  state.inventoryResult   = null;
  state.lastSavedSnapshot = null;

  dom.projectName.value   = "";
  dom.projectSelect.value = "";
  dom.objFile.value       = "";
  dom.modelUnits.disabled = false;
  if (dom.unitsDetectedNote) dom.unitsDetectedNote.style.display = "none";

  seedDefaultProjectInputs();

  // New project starts with empty inventory
  dom.inventoryBody.innerHTML = "";
  dom.inventoryInfinite.checked = true;
  setInventoryQuantityMode();

  renderPartsSummary([]);
  renderPartsTable([], collectInputs());
  clearResults();
  clearViewerModel();
  switchTab("planning");
  setStatus("Started a new project with default settings.", "ok");
}

function updateRecalcButtonState() {
  const hasParts     = state.parts.length > 0;
  const hasInventory = dom.inventoryInfinite.checked ||
                       dom.inventoryBody.querySelectorAll("tr").length > 0;
  const enabled = hasParts && hasInventory;
  dom.inventoryPlan.disabled = !enabled;

  let reason = "";
  if (!hasParts && !hasInventory) {
    reason = "Analyze a model and configure inventory first";
  } else if (!hasParts) {
    reason = "Analyze a model to extract parts first";
  } else if (!hasInventory) {
    reason = "Enable infinite inventory or add inventory rows first";
  }
  if (dom.recalcTipWrap) dom.recalcTipWrap.title = reason;
}

function clearResults() {
  dom.planningSummary.innerHTML       = "";
  dom.planningLayouts.innerHTML       = "";
  dom.inventorySummary.innerHTML      = "";
  dom.inventoryLayouts.innerHTML      = "";
  dom.lumberYardSuggestions.innerHTML = "";
  updateRecalcButtonState();
  renderWorkshopTab();
}

async function saveProject() {
  const name = (dom.projectName.value || "").trim();
  if (!name) {
    setStatus("Enter a project name before saving.", "error");
    return;
  }

  const data       = collectInputs();
  data.partOverrides = state.partOverrides;

  const projects = await readProjectsActive();
  const existing = projects.find((p) => p.name === name);
  const ownerUid = state.firebase.user?.uid || null;

  // Serialize inventoryResult for Firestore:
  //  • strip freeRects (large, only needed during packing)
  //  • convert boardUsage Map → plain object (Firestore cannot store Map)
  const finalInventory = state.inventoryResult
    ? {
        ...state.inventoryResult,
        boards:     state.inventoryResult.boards.map(({ freeRects: _, ...b }) => b),
        boardUsage: Object.fromEntries(state.inventoryResult.boardUsage),
      }
    : null;

  const payload  = {
    id:       existing ? existing.id : crypto.randomUUID(),
    name,
    savedAt:  new Date().toISOString(),
    objText:  state.objText,
    ownerUid,
    inputs:   data,
    finalInventory,
  };

  try {
    if (activeStorageBackend() === "firebase") {
      await saveProjectFirebase(payload);
    } else {
      saveProjectLocal(payload);
    }
    await refreshProjectSelect(payload.id);
    markProjectClean();
    setStatus(
      `Saved project "${name}" to ${
        activeStorageBackend() === "firebase" ? "Firebase cloud" : "local browser storage"
      }.`,
      "ok"
    );
  } catch (error) {
    console.error(error);
    setStatus(`Save failed: ${error.message || "Unknown error"}`, "error");
  }
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
  state.objText         = project.objText || "";
  restoreInputs(project.inputs || {});

  if (state.objText) {
    runAnalyze();
  } else {
    renderPartsSummary([]);
    renderPartsTable([], collectInputs());
    clearResults();
    setStatus(`Loaded project "${project.name}" (no OBJ file stored — re-analyze to load one).`, "ok");
  }

  // Restore finalInventory AFTER runAnalyze() so it isn't wiped by clearExistingResults.
  // Convert boardUsage from plain object (Firestore) back to Map.
  // Restore finalInventory AFTER runAnalyze() so it isn't wiped by clearExistingResults.
  // Convert boardUsage from plain object (Firestore) back to Map.
  const fi = project.finalInventory ?? null;
  state.inventoryResult = fi
    ? { ...fi, boardUsage: new Map(Object.entries(fi.boardUsage ?? {})) }
    : null;

  if (state.inventoryResult) {
    const inputs = collectInputs();
    renderPlanSummary(dom.inventorySummary, state.inventoryResult, "Inventory fit result", inputs.pricePerBoardFoot);
    renderLayouts(dom.inventoryLayouts, state.inventoryResult.boards);
    renderWorkshopTab();
  }

  markProjectClean();
}

async function deleteSelectedProject() {
  const selectedId = dom.projectSelect.value;
  if (!selectedId) {
    setStatus("Choose a saved project first.", "error");
    return;
  }

  const selectedName = dom.projectSelect.options[dom.projectSelect.selectedIndex]?.text || "this project";
  const confirmed = await showConfirm(
    `Delete "${selectedName}"? This cannot be undone.`,
    "Delete",
    dom.deleteProject
  );
  if (!confirmed) return;

  try {
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
  } catch (error) {
    console.error(error);
    setStatus(`Delete failed: ${error.message || "Unknown error"}`, "error");
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Storage backends
// ─────────────────────────────────────────────────────────────────────────────
function activeStorageBackend() {
  return state.firebase.connected && state.firebase.mode === "firebase" ? "firebase" : "local";
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
  if (!state.firebase.connected || !state.firebase.db) return [];
  const uid = state.firebase.user?.uid || null;
  let ref = state.firebase.db.collection("projects");
  if (uid) ref = ref.where("ownerUid", "==", uid);
  const snapshot = await ref.get();
  return snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
}

async function readProjectsActive() {
  if (activeStorageBackend() === "firebase") return readProjectsFirebase();
  return readProjectsLocal();
}

async function readProjectByIdActive(id) {
  if (activeStorageBackend() === "firebase") {
    if (!state.firebase.connected || !state.firebase.db) return null;
    const doc = await state.firebase.db.collection("projects").doc(id).get();
    if (!doc.exists) return null;
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
  const existing = projects.find((p) => p.id === payload.id);
  const next = existing
    ? projects.map((p) => (p.id === payload.id ? payload : p))
    : [...projects, payload];
  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
}

function deleteProjectLocal(id) {
  const projects = readProjectsLocal();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(projects.filter((p) => p.id !== id)));
}

async function saveProjectFirebase(payload) {
  if (!state.firebase.connected || !state.firebase.db) {
    throw new Error("Firebase is not connected.");
  }
  // Guard against Firestore 1 MB document size limit
  const size = JSON.stringify(payload).length;
  if (size > 900_000) {
    throw new Error(
      `Project data (~${Math.round(size / 1024)} KB) is near or over Firebase's 1 MB limit. ` +
        "Try saving without the OBJ file or reduce model complexity."
    );
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
    projects = (await readProjectsActive()).sort(
      (a, b) => (b.savedAt || "").localeCompare(a.savedAt || "")
    );
  } catch (error) {
    console.error(error);
    setStatus(`Could not load project list: ${error.message || "unknown error"}`, "error");
  }
  dom.projectSelect.innerHTML = '<option value="">Select a project…</option>';
  for (const project of projects) {
    const opt = document.createElement("option");
    opt.value = project.id;
    opt.textContent = `${project.name} (${new Date(project.savedAt).toLocaleString()})`;
    dom.projectSelect.append(opt);
  }
  if (selectedId) dom.projectSelect.value = selectedId;
  updateProjectSelectButtons();
}

// ─────────────────────────────────────────────────────────────────────────────
// Firebase connection
// ─────────────────────────────────────────────────────────────────────────────
// ─── Firebase config helpers ───────────────────────────────────────────────
// Config comes from DEFAULT_FIREBASE_CONFIG (hardcoded in app.js) or from a
// previously-saved copy in localStorage (for users who configured it before
// the hardcoded default was added).

function validateFirebaseConfig(config) {
  const required = ["apiKey", "authDomain", "projectId", "appId"];
  const missing  = required.filter((f) => !config[f]);
  return { ok: !missing.length, missing };
}

function readFirebaseConfigLocal() {
  try {
    const raw = localStorage.getItem(FIREBASE_CONFIG_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    console.error(e);
    return null;
  }
}

/** Returns the best available Firebase config, or null if none is valid. */
function getActiveFirebaseConfig() {
  if (validateFirebaseConfig(DEFAULT_FIREBASE_CONFIG).ok) return DEFAULT_FIREBASE_CONFIG;
  const saved = readFirebaseConfigLocal();
  if (saved && validateFirebaseConfig(saved).ok) return saved;
  return null;
}

function setFirebaseStatus(message, type = "") {
  dom.firebaseStatus.className  = `status ${type}`.trim();
  dom.firebaseStatus.textContent = message;
}

// ─────────────────────────────────────────────────────────────────────────────
// Auth — initialization + state listener
// ─────────────────────────────────────────────────────────────────────────────

function initAuth() {
  if (!window.firebase) return;
  const config = getActiveFirebaseConfig();
  if (!config) {
    setFirebaseStatus(
      "Add your Firebase credentials to DEFAULT_FIREBASE_CONFIG in app.js to enable cloud sync."
    );
    return;
  }

  const appName = `cutlist-${config.projectId}`;
  let app = window.firebase.apps.find((a) => a.name === appName);
  if (!app) app = window.firebase.initializeApp(config, appName);

  // ── App Check (must activate before auth/firestore are used) ──────────────
  // Skip App Check on localhost — reCAPTCHA v3 requires a registered public domain
  // and will always fail locally.  Firebase never enforces App Check on localhost
  // anyway (it can't verify a loopback address), so this is safe to skip.
  const isLocalhost = ["localhost", "127.0.0.1", ""].includes(location.hostname);
  if (RECAPTCHA_SITE_KEY && !isLocalhost && window.firebase.appCheck) {
    try {
      app.appCheck().activate(
        new firebase.appCheck.ReCaptchaV3Provider(RECAPTCHA_SITE_KEY),
        true // auto-refresh tokens
      );
      console.info("Firebase App Check activated (reCAPTCHA v3).");
    } catch (err) {
      // Non-fatal: App Check failure should not block the app from loading.
      // In MONITOR mode Firebase still allows traffic; in ENFORCE mode it won't.
      console.warn("App Check activation failed:", err.message);
    }
  } else if (isLocalhost) {
    console.info("App Check skipped on localhost.");
  }

  const auth = app.auth();
  const db   = app.firestore();

  // Store references immediately so the Settings-tab connect button still works
  state.firebase.app    = app;
  state.firebase.auth   = auth;
  state.firebase.db     = db;
  state.firebase.config = config;

  // Single source-of-truth for auth state
  auth.onAuthStateChanged((user) => handleAuthStateChange(user));
}

async function handleAuthStateChange(user) {
  state.firebase.authReady = true;

  if (user) {
    state.firebase.user      = user;
    state.firebase.connected = true;
    state.firebase.mode      = "firebase";

    const role = await fetchOrCreateUserRole(user);
    state.firebase.role = role;
  } else {
    state.firebase.user      = null;
    state.firebase.connected = false;
    state.firebase.mode      = "local";
    state.firebase.role      = null;

    // If Firebase is configured, nudge the user to sign in for cloud sync
    if (getActiveFirebaseConfig()) {
      setFirebaseStatus(
        "Sign in with the account button (↑) to sync projects to the cloud."
      );
    }
  }

  updateSyncStatusIndicator();
  updateUserMenuUI();
  updateSettingsTabVisibility();
  updateProjectSectionVisibility();
  await refreshProjectSelect();
}

async function fetchOrCreateUserRole(user) {
  try {
    const ref  = state.firebase.db.collection("users").doc(user.uid);
    const snap = await ref.get();
    if (snap.exists) return snap.data().role || "standard";

    // First sign-in: create document with role = standard
    await ref.set({
      email:       user.email,
      displayName: user.displayName || null,
      role:        "standard",
      createdAt:   firebase.firestore.FieldValue.serverTimestamp(),
    });
    return "standard";
  } catch (err) {
    console.error("fetchOrCreateUserRole failed:", err.code, err.message);
    setFirebaseStatus(
      `Could not read/write user profile: ${err.message} — check Firestore security rules.`,
      "error"
    );
    return "standard"; // safe fallback — app continues working
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Auth — user menu UI
// ─────────────────────────────────────────────────────────────────────────────

function updateUserMenuUI() {
  const user     = state.firebase.user;
  const role     = state.firebase.role;
  const signedIn = !!user;

  // Avatar: silhouette when signed out, initials when signed in
  dom.userAvatarSilhouette.classList.toggle("hidden", signedIn);
  dom.userAvatarInitials.classList.toggle("hidden", !signedIn);
  if (signedIn && user.email) {
    dom.userAvatarInitials.textContent = user.email.charAt(0).toUpperCase();
  }
  dom.userMenuBtn.classList.toggle("unauthenticated", !signedIn);

  // Dropdown info
  dom.userMenuInfo.classList.toggle("hidden",    !signedIn);
  dom.userMenuDivider.classList.toggle("hidden", !signedIn);
  if (signedIn) {
    dom.userMenuEmail.textContent = user.email || "";
    dom.userMenuRole.textContent  = role || "standard";
    dom.userMenuRole.classList.toggle("admin", role === "admin");
  }

  // Signin / signout
  dom.userMenuSignin.classList.toggle("hidden",  signedIn);
  dom.userMenuSignout.classList.toggle("hidden", !signedIn);
}

function toggleUserMenu() {
  const willOpen = dom.userMenuDropdown.classList.toggle("hidden") === false;
  dom.userMenuBtn.setAttribute("aria-expanded", String(willOpen));
}

function closeUserMenu() {
  dom.userMenuDropdown.classList.add("hidden");
  dom.userMenuBtn.setAttribute("aria-expanded", "false");
}

// ─────────────────────────────────────────────────────────────────────────────
// Admin — Users & Projects panel
// ─────────────────────────────────────────────────────────────────────────────
async function loadAdminUsersAndProjects() {
  const container = dom.adminUsersList;
  if (!container) return;
  if (!state.firebase.db || state.firebase.role !== "admin") {
    container.innerHTML = '<p class="muted">Not available — sign in as an admin first.</p>';
    return;
  }
  container.innerHTML = '<p class="muted">Loading…</p>';

  try {
    // Fetch all users
    const usersSnap    = await state.firebase.db.collection("users").get();
    const usersById    = new Map();
    for (const doc of usersSnap.docs) {
      usersById.set(doc.id, { uid: doc.id, ...doc.data() });
    }

    // Fetch all projects (requires admin-read rule in Firestore)
    const projectsSnap = await state.firebase.db.collection("projects").get();
    const projectsByOwner = new Map();
    for (const doc of projectsSnap.docs) {
      const data  = doc.data();
      const owner = data.ownerUid || "(unknown)";
      if (!projectsByOwner.has(owner)) projectsByOwner.set(owner, []);
      projectsByOwner.get(owner).push({ id: doc.id, name: data.name || doc.id, updatedAt: data.updatedAt });
    }

    container.innerHTML = "";

    if (usersById.size === 0) {
      container.innerHTML = '<p class="muted">No users found.</p>';
      return;
    }

    // Build a table — one row per user
    const table = document.createElement("table");
    table.style.cssText = "width:100%;border-collapse:collapse;font-size:0.88rem";

    const thead = table.createTHead();
    const hrow  = thead.insertRow();
    for (const h of ["Email", "Role", "UID", "Projects"]) {
      const th = document.createElement("th");
      th.textContent = h;
      th.style.cssText = "text-align:left;padding:4px 8px;border-bottom:1px solid var(--border)";
      hrow.append(th);
    }

    const tbody = table.createTBody();
    for (const [uid, user] of [...usersById.entries()].sort((a, b) =>
      (a[1].email || "").localeCompare(b[1].email || "")
    )) {
      const projects = projectsByOwner.get(uid) || [];
      const tr = tbody.insertRow();
      tr.style.verticalAlign = "top";

      const tdStyle = "padding:4px 8px;border-bottom:1px solid var(--border)";
      const emailTd = tr.insertCell(); emailTd.style.cssText = tdStyle;
      emailTd.textContent = user.email || "—";

      const roleTd  = tr.insertCell(); roleTd.style.cssText = tdStyle;
      roleTd.textContent = user.role || "standard";

      const uidTd   = tr.insertCell(); uidTd.style.cssText = `${tdStyle};font-size:0.78rem;color:var(--text-muted,#888);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap`;
      uidTd.title   = uid;
      uidTd.textContent = uid;

      const projTd  = tr.insertCell(); projTd.style.cssText = tdStyle;
      if (projects.length === 0) {
        projTd.textContent = "—";
      } else {
        const ul = document.createElement("ul");
        ul.style.cssText = "margin:0;padding:0 0 0 14px";
        for (const p of projects.sort((a, b) => (a.name || "").localeCompare(b.name || ""))) {
          const li = document.createElement("li");
          li.textContent = p.name;
          ul.append(li);
        }
        projTd.append(ul);
      }

      tbody.append(tr);
    }

    // Show any projects with unknown owners at the bottom
    const unknownProjects = projectsByOwner.get("(unknown)") || [];
    if (unknownProjects.length) {
      const tr = tbody.insertRow();
      const tdStyle = "padding:4px 8px;border-bottom:1px solid var(--border)";
      const emailTd = tr.insertCell(); emailTd.colSpan = 3;
      emailTd.style.cssText = `${tdStyle};color:var(--text-muted,#888)`;
      emailTd.textContent = "(unknown owner)";
      const projTd = tr.insertCell(); projTd.style.cssText = tdStyle;
      projTd.textContent = unknownProjects.map((p) => p.name).join(", ");
    }

    container.append(table);
    container.insertAdjacentHTML("beforeend",
      `<p class="muted" style="margin-top:8px">${usersById.size} user(s) · ${projectsSnap.size} project(s) total</p>`
    );
  } catch (err) {
    console.error("Admin panel error:", err);
    container.innerHTML =
      `<p class="status error">Failed to load: ${err.message}. ` +
      `Check that the Firestore rules allow admins to read all projects.</p>`;
  }
}

function updateSettingsTabVisibility() {
  const isAdmin = state.firebase.role === "admin";
  dom.tabSettings.classList.toggle("hidden", !isAdmin);
  // Mirror visibility in the mobile dropdown
  const mobileAdminItem = dom.tabDropdown?.querySelector('[data-tab-target="settings"]');
  if (mobileAdminItem) mobileAdminItem.classList.toggle("hidden", !isAdmin);
  // If currently on the Settings tab but no longer admin, redirect to Planning
  if (!isAdmin && state.activeTab === "settings") {
    switchTab("planning");
  }
}

function updateProjectSectionVisibility() {
  const signedIn = !!state.firebase.user;
  dom.projectSection.classList.toggle("hidden", !signedIn);
}

// ─────────────────────────────────────────────────────────────────────────────
// Auth — modal helpers
// ─────────────────────────────────────────────────────────────────────────────

function openLoginModal() {
  closeAllAuthModals();
  dom.authLoginEmail.value    = "";
  dom.authLoginPassword.value = "";
  setAuthError(dom.authLoginError, "");
  dom.authLoginModal.classList.remove("hidden");
  dom.authLoginEmail.focus();
}

function openSignupModal() {
  closeAllAuthModals();
  dom.authSignupEmail.value    = "";
  dom.authSignupPassword.value = "";
  dom.authSignupConfirm.value  = "";
  setAuthError(dom.authSignupError, "");
  dom.authSignupModal.classList.remove("hidden");
  dom.authSignupEmail.focus();
}

function openResetModal() {
  closeAllAuthModals();
  dom.authResetEmail.value = "";
  setAuthMessage(dom.authResetMessage, "", "");
  dom.authResetModal.classList.remove("hidden");
  dom.authResetEmail.focus();
}

function closeAllAuthModals() {
  dom.authLoginModal.classList.add("hidden");
  dom.authSignupModal.classList.add("hidden");
  dom.authResetModal.classList.add("hidden");
}

function setAuthError(el, msg) {
  el.textContent = msg;
  el.classList.toggle("hidden", !msg);
}

function setAuthMessage(el, msg, type) {
  el.textContent = msg;
  el.className   = `status ${type}`.trim();
  el.classList.toggle("hidden", !msg);
}

// ─────────────────────────────────────────────────────────────────────────────
// Auth — action handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleLogin() {
  const email = dom.authLoginEmail.value.trim();
  const pass  = dom.authLoginPassword.value;
  if (!email || !pass) {
    setAuthError(dom.authLoginError, "Enter your email and password.");
    return;
  }
  dom.authLoginSubmit.disabled = true;
  try {
    await state.firebase.auth.signInWithEmailAndPassword(email, pass);
    // onAuthStateChanged fires → handleAuthStateChange() does the rest
    closeAllAuthModals();
  } catch (err) {
    setAuthError(dom.authLoginError, friendlyAuthError(err.code));
  } finally {
    dom.authLoginSubmit.disabled = false;
  }
}

async function handleSignup() {
  const email   = dom.authSignupEmail.value.trim();
  const pass    = dom.authSignupPassword.value;
  const confirm = dom.authSignupConfirm.value;
  if (!email || !pass) {
    setAuthError(dom.authSignupError, "Enter an email and password.");
    return;
  }
  if (pass !== confirm) {
    setAuthError(dom.authSignupError, "Passwords do not match.");
    return;
  }
  if (pass.length < 6) {
    setAuthError(dom.authSignupError, "Password must be at least 6 characters.");
    return;
  }
  dom.authSignupSubmit.disabled = true;
  try {
    await state.firebase.auth.createUserWithEmailAndPassword(email, pass);
    // onAuthStateChanged fires → fetchOrCreateUserRole() writes role:standard
    closeAllAuthModals();
  } catch (err) {
    setAuthError(dom.authSignupError, friendlyAuthError(err.code));
  } finally {
    dom.authSignupSubmit.disabled = false;
  }
}

async function handlePasswordReset() {
  const email = dom.authResetEmail.value.trim();
  if (!email) {
    setAuthMessage(dom.authResetMessage, "Enter your email address.", "error");
    return;
  }
  dom.authResetSubmit.disabled = true;
  try {
    await state.firebase.auth.sendPasswordResetEmail(email);
    setAuthMessage(
      dom.authResetMessage,
      "Reset email sent — check your inbox (and spam folder).",
      "ok"
    );
    dom.authResetCancel.textContent = "Close";
  } catch (err) {
    setAuthMessage(dom.authResetMessage, friendlyAuthError(err.code), "error");
  } finally {
    dom.authResetSubmit.disabled = false;
  }
}

async function signOutUser() {
  if (!state.firebase.auth) return;
  await state.firebase.auth.signOut();
  // onAuthStateChanged fires → handleAuthStateChange(null) cleans up state
}

function friendlyAuthError(code) {
  const map = {
    "auth/user-not-found":         "No account found for that email.",
    "auth/wrong-password":         "Incorrect password.",
    "auth/invalid-email":          "Invalid email address.",
    "auth/email-already-in-use":   "An account with that email already exists.",
    "auth/weak-password":          "Password is too weak (min 6 characters).",
    "auth/too-many-requests":      "Too many attempts. Try again later.",
    "auth/network-request-failed": "Network error. Check your connection.",
    "auth/invalid-credential":     "Incorrect email or password.",
  };
  return map[code] || "Authentication error. Please try again.";
}

function updateSyncStatusIndicator() {
  if (!dom.syncStatusIcon) return;
  const healthy = state.firebase.connected && state.firebase.mode === "firebase";
  dom.syncStatusIcon.classList.toggle("green", healthy);
  dom.syncStatusIcon.classList.toggle("red",  !healthy);
  dom.syncStatusIcon.title = healthy
    ? "Sync status: connected to cloud"
    : "Sync status: local mode or cloud disconnected";
}

// ─────────────────────────────────────────────────────────────────────────────
// OBJ file handling + unit detection
// ─────────────────────────────────────────────────────────────────────────────
async function handleObjFile(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  state.objText = await file.text();

  const detectedUnit = detectObjUnits(state.objText);
  if (detectedUnit) {
    dom.modelUnits.value    = detectedUnit;
    dom.modelUnits.disabled = true;
    if (dom.unitsDetectedNote) {
      dom.unitsDetectedNote.textContent = `Auto-detected from file`;
      dom.unitsDetectedNote.style.display = "";
    }
    setStatus(
      `Loaded ${file.name}. Units auto-detected: ${detectedUnit} (locked). Click "Analyze Model" to parse parts.`,
      "ok"
    );
  } else {
    dom.modelUnits.disabled = false;
    if (dom.unitsDetectedNote) dom.unitsDetectedNote.style.display = "none";
    setStatus(`Loaded ${file.name}. Set units manually, then click "Analyze Model".`, "ok");
  }

  const inputs = collectInputs();
  refreshViewerModel(state.objText, unitToMmFactor(inputs.modelUnits) * inputs.unitScale);
}

/**
 * Scan OBJ comment lines for unit declarations.
 * Fusion 360 exports: "# scale 1 unit = 1 cm"
 * Generic exports:    "# units: mm"  /  "# unit = in"  /  "# Millimeters"
 */
function detectObjUnits(text) {
  const lines = text.split(/\r?\n/);
  for (let i = 0; i < Math.min(60, lines.length); i++) {
    const line = lines[i].trim();
    if (!line.startsWith("#")) continue;
    const lower = line.toLowerCase();

    // Fusion 360: "# scale 1 unit = 1 cm"
    const scaleMatch = lower.match(/scale\s+1\s+unit\s*=\s*1\s+(\w+)/);
    if (scaleMatch) {
      const u = normalizeUnit(scaleMatch[1]);
      if (u) return u;
    }

    // Generic: "# units: mm" or "# unit = mm"
    const unitMatch = lower.match(/\bunits?\s*[:=]\s*(\w+)/);
    if (unitMatch) {
      const u = normalizeUnit(unitMatch[1]);
      if (u) return u;
    }

    // Named words
    if (lower.includes("millimeter")) return "mm";
    if (lower.includes("centimeter")) return "cm";
    if (/\bmetre(s)?\b|\bmeter(s)?\b/.test(lower)) return "m";
    if (lower.includes("inches") || /\binch(es)?\b/.test(lower)) return "in";
    if (lower.includes("feet") || /\bfoot\b/.test(lower)) return "ft";
  }
  return null;
}

function normalizeUnit(str) {
  const map = {
    mm: "mm", millimeter: "mm", millimeters: "mm", millimetre: "mm", millimetres: "mm",
    cm: "cm", centimeter: "cm", centimeters: "cm", centimetre: "cm", centimetres: "cm",
    m:  "m",  meter: "m",  meters: "m",  metre: "m",  metres: "m",
    in: "in", inch: "in", inches: "in",
    ft: "ft", foot: "ft", feet: "ft",
  };
  return map[(str || "").toLowerCase()] || null;
}

function refreshViewerFromSettings() {
  if (!state.objText) return;
  const inputs = collectInputs();
  refreshViewerModel(state.objText, unitToMmFactor(inputs.modelUnits) * inputs.unitScale);
}

// ─────────────────────────────────────────────────────────────────────────────
// Analysis / Planning / Inventory actions
// ─────────────────────────────────────────────────────────────────────────────
function runAnalyze() {
  const analysis = analyzeFromCurrentInputs({ clearExistingResults: true });
  if (!analysis) return;
  const msg = `Analyzed ${state.parts.length} part(s) (${analysis.inputs.modelUnits} source → metric output).`;
  setStatus(msg, "ok");
  showModal(`Analysis complete: ${state.parts.length} part(s) found.`, dom.analyze);
}

function runPlanning() {
  const analysis = analyzeFromCurrentInputs({ clearExistingResults: false });
  if (!analysis) return;

  const boardCatalog = buildPlanningCatalog(analysis.inputs);
  state.planningResult = optimizeCutPlan(state.parts, boardCatalog, analysis.inputs);

  renderPlanSummary(
    dom.planningSummary,
    state.planningResult,
    "Planning stock requirement",
    analysis.inputs.pricePerBoardFoot
  );
  renderLayouts(dom.planningLayouts, state.planningResult.boards);
  updateRecalcButtonState();
  renderWorkshopTab();
  setStatus("Planning stock optimization completed.", "ok");

  const boardCount = state.planningResult.boards.length;
  const cost       = formatCurrency(state.planningResult.estimatedCost);
  showModal(`Planning complete: ${boardCount} board(s), estimated cost ${cost}.`, dom.plan);
}

function runInventoryPlan() {
  const analysis = analyzeFromCurrentInputs({ clearExistingResults: false });
  if (!analysis) return;

  const { inventoryInfinite, inventory } = analysis.inputs;

  // Require rows if not infinite
  if (!inventory.length && !inventoryInfinite) {
    setStatus(
      "Add at least one lumber inventory row, or enable Assume infinite inventory quantities.",
      "error"
    );
    return;
  }

  // Build the board catalog:
  //  • rows present (infinite or not) → use those rows
  //  • infinite but no rows → fall back to planning catalog (any standard size, unlimited)
  let boardCatalog;
  if (!inventory.length && inventoryInfinite) {
    boardCatalog = buildPlanningCatalog(analysis.inputs);
  } else {
    boardCatalog = buildInventoryCatalog(inventory);
  }

  // Re-assign parts using only thickness quarters available in the catalog.
  // This ensures that if the planner used 8/4 stock but the inventory only has
  // 4/4, parts are reassigned with layers=2 (lamination) rather than left unmet.
  const inventoryQuarters = [...new Set(boardCatalog.map((r) => r.thicknessQuarter))].sort((a, b) => a - b);
  const inventoryInputs = { ...analysis.inputs, thicknessOptionsQuarters: inventoryQuarters };
  const inventoryParts  = assignPartsForStock(state.rawParts, inventoryInputs, state.partOverrides);

  state.inventoryResult = optimizeCutPlan(inventoryParts, boardCatalog, inventoryInputs);

  renderPlanSummary(
    dom.inventorySummary,
    state.inventoryResult,
    "Inventory fit result",
    analysis.inputs.pricePerBoardFoot
  );
  renderLayouts(dom.inventoryLayouts, state.inventoryResult.boards);

  const suggestions = buildYardSuggestions(
    inventoryParts,
    analysis.inputs.inventory,
    analysis.inputs.kerfMm,
    analysis.inputs.planningLengthMaxFt
  );
  renderYardSuggestions(suggestions);

  if (state.inventoryResult.unmetParts.length) {
    const unmetPlan = optimizeCutPlan(
      inventoryParts.filter((part) =>
        state.inventoryResult.unmetParts.some(
          (u) => u.partId && u.partId === part.id
        )
      ),
      boardCatalog.map((item) => ({ ...item, quantity: null })),
      inventoryInputs
    );
    renderAdditionalNeeds(dom.inventorySummary, unmetPlan, analysis.inputs.pricePerBoardFoot);
  }

  renderWorkshopTab();
  setStatus("Lumber yard recalculation completed.", "ok");

  const boardCount  = state.inventoryResult.boards.length;
  const unmetCount  = state.inventoryResult.unmetParts.length;
  const modalMsg    = unmetCount > 0
    ? `Recalculation complete: ${boardCount} board(s) used, ${unmetCount} part(s) unmet.`
    : `Recalculation complete: ${boardCount} board(s), all parts allocated.`;
  showModal(modalMsg, dom.inventoryPlan);
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
  const parsed    = parseObjObjects(state.objText, scaleToMm);

  // Only re-render viewer when model or scale actually changed
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
    state.planningResult  = null;
    state.inventoryResult = null;
    clearResults();
  }

  return { inputs };
}

function hasExplicitPartOverrides() {
  return Object.values(state.partOverrides).some(
    (entry) =>
      entry &&
      (entry.thicknessOverrideQuarter != null ||
       typeof entry.grainLock === "boolean" ||
       typeof entry.grainDir  === "string")
  );
}

function applyGlobalOverrideInMemory(quarter) {
  for (const part of state.rawParts) {
    const current = state.partOverrides[part.id] || {};
    state.partOverrides[part.id] = { ...current, thicknessOverrideQuarter: quarter };
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
      state.partOverrides[part.id] = { ...current, thicknessOverrideQuarter: override };
    }
  }

  const inputs = collectInputs();
  state.parts  = assignPartsForStock(state.rawParts, inputs, state.partOverrides);
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
  state.parts  = assignPartsForStock(state.rawParts, inputs, state.partOverrides);
  renderPartsSummary(state.parts);
  renderPartsTable(state.parts, inputs);
  clearResults();
  setStatus("Cleared per-part grain and thickness overrides.", "ok");
}

function pruneOverridesToKnownParts(parts) {
  const ids = new Set(parts.map((p) => p.id));
  for (const key of Object.keys(state.partOverrides)) {
    if (!ids.has(key)) delete state.partOverrides[key];
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// OBJ parsing — bounding-box extraction
// ─────────────────────────────────────────────────────────────────────────────
function parseObjObjects(text, scaleToMm) {
  const lines    = text.split(/\r?\n/);
  const vertices = [null]; // 1-based indexing
  const objectToVertexIndexes = new Map();
  let currentObject = "Unlabeled";
  ensureObj(currentObject);

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;

    if (line.startsWith("v ")) {
      const [, xs, ys, zs] = line.split(/\s+/);
      const x = Number(xs), y = Number(ys), z = Number(zs);
      if (Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(z)) {
        vertices.push([x * scaleToMm, y * scaleToMm, z * scaleToMm]);
      }
      continue;
    }

    if (line.startsWith("o ") || line.startsWith("g ")) {
      currentObject = line.slice(2).trim() || "Unlabeled";
      ensureObj(currentObject);
      continue;
    }

    if (line.startsWith("f ")) {
      const tokens  = line.split(/\s+/).slice(1);
      const indexes = objectToVertexIndexes.get(currentObject);
      for (const token of tokens) {
        const rawIndex = token.split("/")[0];
        if (!rawIndex) continue;
        let index = Number(rawIndex);
        if (!Number.isInteger(index)) continue;
        if (index < 0) index = vertices.length + index;
        if (index > 0 && index < vertices.length) indexes.add(index);
      }
    }
  }

  const parts = [];
  let counter  = 1;
  for (const [name, indexSet] of objectToVertexIndexes.entries()) {
    if (!indexSet.size) continue;
    const points = [];
    for (const idx of indexSet) {
      const pt = vertices[idx];
      if (pt) points.push(pt);
    }
    if (!points.length) continue;
    const [rawX, rawY, rawZ] = getBoundingBoxDimensions(points);
    const pcaDims = computePcaDimensions(points); // [length, width, thickness] sorted largest→smallest, or null
    const curved  = computeCurvatureFlag(points);
    const safeName = name || `Part ${counter}`;
    parts.push({
      id:   `${slugify(safeName)}-${counter}`,
      name: safeName,
      xMm: rawX, yMm: rawY, zMm: rawZ, // AABB raw values — shown in hover tooltip
      pcaDims, // true dimensions from PCA; null when < 4 vertices
      curved,  // true when geometry appears to have curved surfaces
    });
    counter += 1;
  }
  return parts;

  function ensureObj(name) {
    if (!objectToVertexIndexes.has(name)) objectToVertexIndexes.set(name, new Set());
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 3D Viewer
// ─────────────────────────────────────────────────────────────────────────────
function initModelViewer() {
  if (!dom.modelViewer || !window.THREE || !window.THREE.OrbitControls) {
    setStatus("3D preview unavailable (viewer library failed to load).", "error");
    return;
  }

  const renderer = new THREE.WebGLRenderer({ canvas: dom.modelViewer, antialias: true, alpha: false });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));

  const scene = new THREE.Scene();
  scene.background = new THREE.Color(0xf8f2e6);

  const camera = new THREE.PerspectiveCamera(45, 1, 1, 300000);
  camera.position.set(1200, 900, 1200);

  const controls = new THREE.OrbitControls(camera, renderer.domElement);
  controls.enableDamping  = true;
  controls.dampingFactor  = 0.08;
  controls.rotateSpeed    = 0.85;
  controls.zoomSpeed      = 0.9;
  controls.panSpeed       = 0.8;
  controls.target.set(0, 0, 0);

  scene.add(new THREE.AmbientLight(0xffffff, 0.75));
  const keyLight = new THREE.DirectionalLight(0xffffff, 0.7);
  keyLight.position.set(1, 2, 1);
  scene.add(keyLight);
  const fillLight = new THREE.DirectionalLight(0xfff6df, 0.35);
  fillLight.position.set(-1.4, -0.8, 0.7);
  scene.add(fillLight);

  const modelGroup = new THREE.Group();
  // Fusion 360 exports OBJ with Z-up coordinates; Three.js uses Y-up.
  // Rotating -90° around X maps the design's Z-up to Three.js's Y-up,
  // correcting the 90° tilt that would otherwise appear in the viewer.
  modelGroup.rotation.x = -Math.PI / 2;
  scene.add(modelGroup);

  const grid = new THREE.GridHelper(4000, 40, 0x9f8a70, 0xd7c8b0);
  scene.add(grid);

  state.viewer = {
    renderer, scene, camera, controls, modelGroup, grid,
    originalCamera: null,
    observer: null,
    rafId: null,
    lastObjText: null,
    lastScaleToMm: null,
  };

  const observer = new ResizeObserver(() => resizeViewer());
  observer.observe(dom.modelViewer.parentElement);
  state.viewer.observer = observer;

  const animate = () => {
    if (!state.viewer) return;
    state.viewer.controls.update();
    state.viewer.renderer.render(state.viewer.scene, state.viewer.camera);
    state.viewer.rafId = requestAnimationFrame(animate);
  };
  animate();
  resizeViewer();
}

function resizeViewer() {
  if (!state.viewer || !dom.modelViewer) return;
  const rect   = dom.modelViewer.getBoundingClientRect();
  const width  = Math.max(1, Math.floor(rect.width));
  const height = Math.max(1, Math.floor(rect.height));
  state.viewer.renderer.setSize(width, height, false);
  state.viewer.camera.aspect = width / height;
  state.viewer.camera.updateProjectionMatrix();
}

function clearViewerModel() {
  if (!state.viewer) return;
  while (state.viewer.modelGroup.children.length) {
    const child = state.viewer.modelGroup.children[0];
    state.viewer.modelGroup.remove(child);
    if (child.geometry) child.geometry.dispose();
    if (child.material) {
      if (Array.isArray(child.material)) {
        child.material.forEach((m) => m.dispose());
      } else {
        child.material.dispose();
      }
    }
  }
  state.viewer.lastObjText   = null;
  state.viewer.lastScaleToMm = null;
}

function refreshViewerModel(objText, scaleToMm) {
  if (!state.viewer) return;

  // Skip re-render if neither the model text nor the scale changed
  if (
    objText === state.viewer.lastObjText &&
    nearlyEqual(scaleToMm, state.viewer.lastScaleToMm || 0)
  ) {
    return;
  }
  state.viewer.lastObjText   = objText;
  state.viewer.lastScaleToMm = scaleToMm;

  if (!objText) {
    clearViewerModel();
    return;
  }

  const meshes = parseObjMeshesForViewer(objText, scaleToMm);
  clearViewerModel();
  // Restore cache after clear
  state.viewer.lastObjText   = objText;
  state.viewer.lastScaleToMm = scaleToMm;

  const palette = [0xa0592a, 0xba7b43, 0x7e8f46, 0x6d597a, 0x3a6b58, 0x5f6e53, 0x6b4f3f, 0x43617a];

  meshes.forEach((meshDef, index) => {
    const geometry = new THREE.BufferGeometry();
    geometry.setAttribute("position", new THREE.Float32BufferAttribute(meshDef.positions, 3));
    geometry.computeVertexNormals();

    const material = new THREE.MeshStandardMaterial({
      color:     palette[index % palette.length],
      roughness: 0.72,
      metalness: 0.03,
      side:      THREE.DoubleSide,
    });
    const mesh  = new THREE.Mesh(geometry, material);
    mesh.name   = meshDef.name;
    state.viewer.modelGroup.add(mesh);
  });

  frameViewerOnModel();
}

function parseObjMeshesForViewer(text, scaleToMm) {
  const lines    = text.split(/\r?\n/);
  const vertices = [null];
  const objects  = new Map();
  let currentObject = "Unlabeled";
  ensureObj(currentObject);

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;

    if (line.startsWith("v ")) {
      const [, xs, ys, zs] = line.split(/\s+/);
      const x = Number(xs), y = Number(ys), z = Number(zs);
      if (Number.isFinite(x) && Number.isFinite(y) && Number.isFinite(z)) {
        vertices.push([x * scaleToMm, y * scaleToMm, z * scaleToMm]);
      }
      continue;
    }

    if (line.startsWith("o ") || line.startsWith("g ")) {
      currentObject = line.slice(2).trim() || "Unlabeled";
      ensureObj(currentObject);
      continue;
    }

    if (line.startsWith("f ")) {
      const tokens  = line.split(/\s+/).slice(1);
      const indexes = [];
      for (const token of tokens) {
        const rawIndex = token.split("/")[0];
        if (!rawIndex) continue;
        let index = Number(rawIndex);
        if (!Number.isInteger(index)) continue;
        if (index < 0) index = vertices.length + index;
        if (index > 0 && index < vertices.length) indexes.push(index);
      }
      if (indexes.length < 3) continue;
      const triangles = objects.get(currentObject);
      for (let i = 1; i < indexes.length - 1; i++) {
        triangles.push([indexes[0], indexes[i], indexes[i + 1]]);
      }
    }
  }

  const meshes = [];
  for (const [name, triangles] of objects.entries()) {
    if (!triangles.length) continue;
    const positions = [];
    for (const [a, b, c] of triangles) {
      for (const idx of [a, b, c]) {
        const pt = vertices[idx];
        if (pt) positions.push(pt[0], pt[1], pt[2]);
      }
    }
    if (positions.length) meshes.push({ name, positions });
  }
  return meshes;

  function ensureObj(name) {
    if (!objects.has(name)) objects.set(name, []);
  }
}

function frameViewerOnModel() {
  if (!state.viewer) return;
  const { camera, controls, modelGroup } = state.viewer;
  const box = new THREE.Box3().setFromObject(modelGroup);
  if (box.isEmpty()) {
    camera.position.set(1200, 900, 1200);
    controls.target.set(0, 0, 0);
    controls.update();
    return;
  }

  const center = box.getCenter(new THREE.Vector3());

  // Use the bounding sphere radius so the fit works correctly regardless of
  // which axis is longest and from whichever diagonal the camera approaches.
  const sphere = new THREE.Sphere();
  box.getBoundingSphere(sphere);
  const radius = sphere.radius;

  // Effective FOV: on portrait viewports the horizontal FOV is narrower, so use
  // that as the limiting dimension to keep the model fully visible.
  const vFovRad   = (camera.fov * Math.PI) / 180;
  const effFovRad = camera.aspect >= 1
    ? vFovRad
    : 2 * Math.atan(Math.tan(vFovRad / 2) * camera.aspect);

  // Distance so the bounding sphere fills ~80% of the viewport (1.25 = 20% padding).
  const distance = Math.max(400, (radius / Math.tan(effFovRad / 2)) * 1.25);

  // 3/4 view: equal X/Z offset, slightly lower Y so the model reads naturally.
  const dx = 1, dy = 0.65, dz = 1;
  const dLen = Math.sqrt(dx * dx + dy * dy + dz * dz);

  camera.near = Math.max(1, radius / 100);
  camera.far  = Math.max(10000, distance * 10);
  camera.position.set(
    center.x + (dx / dLen) * distance,
    center.y + (dy / dLen) * distance,
    center.z + (dz / dLen) * distance,
  );
  camera.updateProjectionMatrix();

  controls.target.copy(center);
  controls.update();

  // Store for Reset View — only on first load of a new model
  if (!state.viewer.originalCamera) {
    state.viewer.originalCamera = {
      position: camera.position.clone(),
      target:   controls.target.clone(),
    };
  }
}

function resetViewerCamera() {
  if (!state.viewer) return;
  const { camera, controls, originalCamera } = state.viewer;
  if (!originalCamera) {
    frameViewerOnModel();
    return;
  }
  camera.position.copy(originalCamera.position);
  controls.target.copy(originalCamera.target);
  controls.update();
}

// ─────────────────────────────────────────────────────────────────────────────
// Geometry helpers
// ─────────────────────────────────────────────────────────────────────────────
/**
 * Curved-surface heuristic.
 *
 * For a prismatic (rectangular) part every vertex lies on one of the 6 flat faces,
 * so it is near an extreme (within 15 % of range) on at least one axis.
 * Vertices that are NOT near any axis extreme can only appear on curved geometry.
 * If more than 15 % of vertices are "fully interior" we flag the part as curved.
 *
 * Returns true  → likely curved surface
 *         false → appears prismatic
 */
function computeCurvatureFlag(points) {
  if (!points || points.length < 12) return false;

  let minX = Infinity, maxX = -Infinity;
  let minY = Infinity, maxY = -Infinity;
  let minZ = Infinity, maxZ = -Infinity;
  for (const [x, y, z] of points) {
    if (x < minX) minX = x; if (x > maxX) maxX = x;
    if (y < minY) minY = y; if (y > maxY) maxY = y;
    if (z < minZ) minZ = z; if (z > maxZ) maxZ = z;
  }
  const rx = maxX - minX, ry = maxY - minY, rz = maxZ - minZ;
  // Skip degenerate / nearly-flat geometry
  if (rx < 0.1 || ry < 0.1 || rz < 0.1) return false;

  const margin = 0.15; // within 15 % of range from a face = "on that face"
  let interior = 0;
  for (const [x, y, z] of points) {
    const nearX = (x - minX) / rx < margin || (maxX - x) / rx < margin;
    const nearY = (y - minY) / ry < margin || (maxY - y) / ry < margin;
    const nearZ = (z - minZ) / rz < margin || (maxZ - z) / rz < margin;
    if (!nearX && !nearY && !nearZ) interior++;
  }
  return (interior / points.length) > 0.15;
}

function getBoundingBoxDimensions(points) {
  let minX = Infinity,  maxX = -Infinity;
  let minY = Infinity,  maxY = -Infinity;
  let minZ = Infinity,  maxZ = -Infinity;
  for (const [x, y, z] of points) {
    if (x < minX) minX = x; if (x > maxX) maxX = x;
    if (y < minY) minY = y; if (y > maxY) maxY = y;
    if (z < minZ) minZ = z; if (z > maxZ) maxZ = z;
  }
  return [maxX - minX, maxY - minY, maxZ - minZ].map((v) => Number(v.toFixed(3)));
}

// ─────────────────────────────────────────────────────────────────────────────
// PCA dimension analysis
// For parts placed at an angle in the design, axis-aligned bounding boxes
// (AABB) over-report dimensions. PCA finds the true principal axes of the
// vertex cloud via Jacobi eigenvalue iteration and measures extents along
// those axes — giving the actual board length/width/thickness regardless of
// how the part is rotated in 3D space.
// ─────────────────────────────────────────────────────────────────────────────
function computePcaDimensions(points) {
  const n = points.length;
  if (n < 4) return null; // not enough points for meaningful PCA

  // Centroid
  let cx = 0, cy = 0, cz = 0;
  for (const [x, y, z] of points) { cx += x; cy += y; cz += z; }
  cx /= n; cy /= n; cz /= n;

  // 3×3 symmetric covariance matrix
  let c00 = 0, c01 = 0, c02 = 0, c11 = 0, c12 = 0, c22 = 0;
  for (const [x, y, z] of points) {
    const dx = x - cx, dy = y - cy, dz = z - cz;
    c00 += dx * dx; c01 += dx * dy; c02 += dx * dz;
    c11 += dy * dy; c12 += dy * dz; c22 += dz * dz;
  }
  c00 /= n; c01 /= n; c02 /= n; c11 /= n; c12 /= n; c22 /= n;

  // Jacobi eigenvalue iteration — columns of V converge to eigenvectors
  const M = [[c00, c01, c02], [c01, c11, c12], [c02, c12, c22]];
  const V = [[1, 0, 0], [0, 1, 0], [0, 0, 1]];

  for (let iter = 0; iter < 60; iter++) {
    let maxAbs = 0, p = 0, q = 1;
    for (let i = 0; i < 3; i++) {
      for (let j = i + 1; j < 3; j++) {
        if (Math.abs(M[i][j]) > maxAbs) { maxAbs = Math.abs(M[i][j]); p = i; q = j; }
      }
    }
    if (maxAbs < 1e-12) break;

    const theta = (M[q][q] - M[p][p]) / (2 * M[p][q]);
    const t     = (theta >= 0 ? 1 : -1) / (Math.abs(theta) + Math.sqrt(1 + theta * theta));
    const c     = 1 / Math.sqrt(1 + t * t);
    const s     = t * c;

    const Mpp = M[p][p] - t * M[p][q];
    const Mqq = M[q][q] + t * M[p][q];
    M[p][p] = Mpp; M[q][q] = Mqq; M[p][q] = M[q][p] = 0;
    for (let r = 0; r < 3; r++) {
      if (r !== p && r !== q) {
        const Mrp = c * M[r][p] - s * M[r][q];
        const Mrq = s * M[r][p] + c * M[r][q];
        M[r][p] = M[p][r] = Mrp; M[r][q] = M[q][r] = Mrq;
      }
    }
    for (let r = 0; r < 3; r++) {
      const Vrp = c * V[r][p] - s * V[r][q];
      const Vrq = s * V[r][p] + c * V[r][q];
      V[r][p] = Vrp; V[r][q] = Vrq;
    }
  }

  // Project vertices onto each eigenvector (column of V); measure extents
  const dims = [0, 1, 2].map((col) => {
    const ev = [V[0][col], V[1][col], V[2][col]];
    let minP = Infinity, maxP = -Infinity;
    for (const [x, y, z] of points) {
      const proj = (x - cx) * ev[0] + (y - cy) * ev[1] + (z - cz) * ev[2];
      if (proj < minP) minP = proj;
      if (proj > maxP) maxP = proj;
    }
    return maxP - minP;
  });

  dims.sort((a, b) => b - a); // largest → smallest: [length, width, thickness]
  return dims.map((d) => Number(d.toFixed(3)));
}

// ─────────────────────────────────────────────────────────────────────────────
// Part assignment
// ─────────────────────────────────────────────────────────────────────────────
function assignPartsForStock(rawParts, inputs, partOverrides) {
  const quarters = [...inputs.thicknessOptionsQuarters].sort((a, b) => a - b);
  const oriented = [];

  for (const rawPart of rawParts) {
    const override       = partOverrides[rawPart.id] || {};
    const overrideQuarter =
      override.thicknessOverrideQuarter == null
        ? null
        : Number(override.thicknessOverrideQuarter);
    // Grain direction — 3-way enum: "long" | "mid" | "free"
    // "long" = grain along longest axis, nesting cannot rotate (was grainLock: true)
    // "mid"  = grain along middle axis, blank pre-rotated before nesting
    // "free" = nesting may rotate freely (was grainLock: false)
    // Backward compat: old saved overrides may have grainLock boolean instead of grainDir string.
    let grainDir;
    if (typeof override.grainDir === "string") {
      grainDir = override.grainDir;
    } else if (typeof override.grainLock === "boolean") {
      grainDir = override.grainLock ? "long" : "free";
    } else {
      grainDir = inputs.defaultGrainLock ? "long" : "free";
    }
    const excluded = Boolean(override.excluded);

    const canonical       = canonicalizePartAxes(rawPart);
    const netLengthMm     = canonical.x.value;
    const netWidthMm      = canonical.y.value;
    const netThicknessMm  = canonical.z.value;
    const roughLengthMm   = netLengthMm    + inputs.milling.lengthMm;
    const roughWidthMm    = netWidthMm     + inputs.milling.widthMm;
    const roughThicknessMm = netThicknessMm + inputs.milling.thicknessMm;
    const thicknessPlan   = resolveStockPlan(roughThicknessMm, quarters, overrideQuarter);

    // Orientation string: show "PCA-corrected" only when PCA dims differ measurably
    // from the sorted AABB dims, indicating the part is genuinely angled in 3D space.
    let orientation;
    if (rawPart.pcaDims) {
      const aabbSorted = [rawPart.xMm, rawPart.yMm, rawPart.zMm].sort((a, b) => b - a);
      const pca = rawPart.pcaDims;
      const isAngled = aabbSorted.some(
        (v, i) => v > 0.1 && Math.abs(v - pca[i]) / v > 0.02 // >2% deviation on any axis
      );
      orientation = isAngled
        ? "PCA-corrected (angled part)"
        : `X≤${canonical.x.axis} (grain), Y≤${canonical.y.axis}, Z≤${canonical.z.axis}`;
    } else {
      orientation = `X≤${canonical.x.axis} (grain), Y≤${canonical.y.axis}, Z≤${canonical.z.axis}`;
    }

    const base = {
      id: rawPart.id,
      name: rawPart.name,
      rawMm: { x: rawPart.xMm, y: rawPart.yMm, z: rawPart.zMm }, // always AABB for hover display
      curved: Boolean(rawPart.curved),
      netLengthMm:      roundTo(netLengthMm, 2),
      netWidthMm:       roundTo(netWidthMm, 2),
      netThicknessMm:   roundTo(netThicknessMm, 2),
      roughLengthMm:    roundTo(roughLengthMm, 2),
      roughWidthMm:     roundTo(roughWidthMm, 2),
      roughThicknessMm: roundTo(roughThicknessMm, 2),
      orientation,
      grainDir,
      excluded,
      thicknessOverrideQuarter: overrideQuarter,
    };

    if (!thicknessPlan.ok) {
      oriented.push({
        ...base,
        stockQuarter: null, stockThicknessMm: null, layers: 0, thicknessWasteMm: null,
        status: "invalid",
        reason: overrideQuarter != null
          ? `Override ${overrideQuarter}/4 is not available in thickness options.`
          : "No valid stock thickness options for this part.",
      });
      continue;
    }

    oriented.push({
      ...base,
      stockQuarter:     thicknessPlan.stockQuarter,
      stockThicknessMm: roundTo(thicknessPlan.stockThicknessMm, 2),
      layers:           thicknessPlan.layers,
      thicknessWasteMm: roundTo(thicknessPlan.wasteMm, 2),
      status: "ok",
      reason: "",
    });
  }

  return oriented.sort((a, b) => {
    if (a.stockQuarter == null && b.stockQuarter != null) return 1;
    if (a.stockQuarter != null && b.stockQuarter == null) return -1;
    const qa = a.stockQuarter ?? 0, qb = b.stockQuarter ?? 0;
    return qa - qb || b.roughLengthMm * b.roughWidthMm - a.roughLengthMm * a.roughWidthMm;
  });
}

function canonicalizePartAxes(rawPart) {
  // Use PCA dims when available: they give true board dimensions for angled parts.
  // PCA dims are already sorted largest→smallest (length, width, thickness).
  if (rawPart.pcaDims) {
    const [l, w, t] = rawPart.pcaDims;
    return {
      x: { axis: "PC1", value: l },
      y: { axis: "PC2", value: w },
      z: { axis: "PC3", value: t },
    };
  }
  // Fallback: AABB sort (only for parts with < 4 vertices)
  const dims = [
    { axis: "X", value: rawPart.xMm },
    { axis: "Y", value: rawPart.yMm },
    { axis: "Z", value: rawPart.zMm },
  ].sort((a, b) => b.value - a.value);
  return { x: dims[0], y: dims[1], z: dims[2] };
}

function resolveStockPlan(roughThicknessMm, quarters, overrideQuarter) {
  if (!quarters.length) return { ok: false, reason: "No stock thickness options." };

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
    if (!best) { best = candidate; continue; }
    if (candidate.layers < best.layers)                                               { best = candidate; continue; }
    if (candidate.layers === best.layers && candidate.wasteMm < best.wasteMm - EPSILON) { best = candidate; continue; }
    if (candidate.layers === best.layers && nearlyEqual(candidate.wasteMm, best.wasteMm) &&
        candidate.stockQuarter < best.stockQuarter) {
      best = candidate;
    }
  }

  if (!best) return { ok: false, reason: "Could not find a valid stock option." };
  return { ok: true, ...best };
}

// ─────────────────────────────────────────────────────────────────────────────
// Board catalogs
// ─────────────────────────────────────────────────────────────────────────────
function buildPlanningCatalog(inputs) {
  const widths  = generateIntRange(inputs.planningWidthMinIn,  inputs.planningWidthMaxIn);
  const lengths = generateIntRange(inputs.planningLengthMinFt, inputs.planningLengthMaxFt);
  const catalog = [];
  for (const quarter of inputs.thicknessOptionsQuarters) {
    for (const widthIn of widths) {
      for (const lengthFt of lengths) {
        catalog.push({
          thicknessQuarter: quarter,
          widthIn,
          lengthFt,
          widthMm:  widthIn  * INCH_TO_MM,
          lengthMm: lengthFt * FOOT_TO_MM,
          quantity: null,
          source:   "planner",
        });
      }
    }
  }
  return catalog;
}

function buildInventoryCatalog(inventoryRows) {
  return inventoryRows.map((row) => ({
    thicknessQuarter: row.thicknessQuarter,
    widthIn:   row.widthIn,
    lengthFt:  row.lengthFt,
    widthMm:   row.widthIn  * INCH_TO_MM,
    lengthMm:  row.lengthFt * FOOT_TO_MM,
    quantity:  row.quantity,
    source:    "inventory",
  }));
}

function generateIntRange(min, max) {
  const lo = Math.max(1, Math.ceil(min));
  const hi = Math.max(lo, Math.floor(max));
  const result = [];
  for (let i = lo; i <= hi; i++) result.push(i);
  return result;
}

// ─────────────────────────────────────────────────────────────────────────────
// Cut-plan optimizer (MaxRects-style 2-D nesting)
// ─────────────────────────────────────────────────────────────────────────────
function optimizeCutPlan(parts, boardCatalog, inputs) {
  const spacingMm        = inputs.kerfMm + inputs.milling.ripMarginMm;
  const endTrimMm        = inputs.milling.boardEndTrimMm;
  const maxPlanerWidthMm = (inputs.maxPlanerWidthIn ?? 0) * INCH_TO_MM;
  const unmetParts = [];

  // Pre-compute the widest available board per thickness quarter.
  // Parts whose rough width exceeds this limit are split into glued-up panel strips.
  const maxWidthByQuarter = new Map();
  for (const type of boardCatalog) {
    const cur = maxWidthByQuarter.get(type.thicknessQuarter) ?? 0;
    if (type.widthMm > cur) maxWidthByQuarter.set(type.thicknessQuarter, type.widthMm);
  }
  const globalMaxWidthMm = boardCatalog.length
    ? Math.max(...boardCatalog.map((t) => t.widthMm))
    : 0;

  const blanks = [];
  for (const part of parts) {
    if (part.excluded) continue; // user explicitly excluded from planning
    if (part.status !== "ok" || !part.stockQuarter || part.layers < 1) {
      unmetParts.push({
        partId: part.id,
        partName: part.name,
        reason: part.reason || "Part is missing valid stock assignment.",
      });
      continue;
    }

    // Max board width for this part's stock quarter (fall back to global max for upsized cases)
    const maxBoardWidthMm = maxWidthByQuarter.get(part.stockQuarter) ?? globalMaxWidthMm;

    for (let layer = 1; layer <= part.layers; layer++) {
      // "mid" grain: swap rough dims so the blank enters nesting already rotated 90°,
      // then lock rotation to preserve that orientation.
      const blankWidth  = part.grainDir === "mid" ? part.roughLengthMm : part.roughWidthMm;
      const blankLength = part.grainDir === "mid" ? part.roughWidthMm  : part.roughLengthMm;

      if (maxBoardWidthMm > 0 && blankWidth > maxBoardWidthMm + EPSILON) {
        // ── Panel glue-up: part is wider than any available board ──────────
        // Split into the minimum number of strips that each fit on a single board.
        // Each strip gets PANEL_JOINT_MM of extra width for edge-jointing the glue faces.
        let numStrips = Math.ceil(blankWidth / maxBoardWidthMm);
        let stripWidth = blankWidth / numStrips + PANEL_JOINT_MM;
        if (stripWidth > maxBoardWidthMm) numStrips++; // safety: if joint allowance pushes over
        stripWidth = roundTo(blankWidth / numStrips + PANEL_JOINT_MM, 1);

        const layerNote = part.layers > 1 ? ` lam ${layer}/${part.layers}` : "";
        for (let strip = 1; strip <= numStrips; strip++) {
          blanks.push({
            id:              `${part.id}-L${layer}-P${strip}`,
            partId:          part.id,
            basePartName:    part.name,
            name:            `${part.name} (panel ${strip}/${numStrips}${layerNote})`,
            widthMm:         stripWidth,
            lengthMm:        blankLength,
            stockQuarter:    part.stockQuarter,
            grainLock:       part.grainDir !== "free",
            panelStripCount: numStrips,
            panelStripIndex: strip,
            panelFullWidthMm: blankWidth,
          });
        }
      } else {
        blanks.push({
          id:           `${part.id}-L${layer}`,
          partId:       part.id,
          basePartName: part.name,
          name: part.layers > 1 ? `${part.name} (lam ${layer}/${part.layers})` : part.name,
          widthMm:      blankWidth,
          lengthMm:     blankLength,
          stockQuarter: part.stockQuarter,
          grainLock:    part.grainDir !== "free",
        });
      }
    }
  }

  const grouped      = groupBy(blanks, (b) => String(b.stockQuarter));
  const boards       = [];
  let   boardCounter = 1;

  for (const [quarterKey, group] of grouped.entries()) {
    const quarter = Number(quarterKey);
    let types = boardCatalog
      .filter((row) => row.thicknessQuarter === quarter)
      .map((row, index) => ({
        ...row,
        typeId:      `${quarter}-${row.widthIn}-${row.lengthFt}-${index}`,
        remaining:   row.quantity == null ? Infinity : row.quantity,
        upsized:     false,
        neededQuarter: quarter,
      }));

    // When no exact-thickness boards are in inventory, try the next thicker quarter.
    // Using thicker stock is wasteful but better than leaving parts unmet.
    if (!types.length) {
      const allQuarters = [...new Set(boardCatalog.map((r) => r.thicknessQuarter))].sort((a, b) => a - b);
      const nextThicker = allQuarters.find((q) => q > quarter);
      if (nextThicker != null) {
        types = boardCatalog
          .filter((row) => row.thicknessQuarter === nextThicker)
          .map((row, index) => ({
            ...row,
            typeId:        `${nextThicker}-${row.widthIn}-${row.lengthFt}-${index}`,
            remaining:     row.quantity == null ? Infinity : row.quantity,
            upsized:       true,
            neededQuarter: quarter,
          }));
      }
    }

    if (!types.length) {
      for (const blank of group) {
        unmetParts.push({
          partId:   blank.partId,
          partName: blank.name,
          reason:   `No catalog/inventory board type for ${quarter}/4 stock (no thicker alternative available).`,
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

    for (let i = 0; i < sorted.length; i++) {
      const blank    = sorted[i];
      const remaining = sorted.slice(i); // used for density scoring
      let placement = findBestPlacementAcrossBoards(openBoards, blank, spacingMm);

      if (!placement) {
        const boardType = chooseBoardType(types, blank, spacingMm, endTrimMm, remaining, maxPlanerWidthMm);
        if (!boardType) {
          unmetParts.push({
            partId:   blank.partId,
            partName: blank.name,
            reason:   "No available board can fit this rough blank.",
          });
          continue;
        }
        boardType.remaining -= 1;
        const board = createBoardFromType(boardType, `B${boardCounter++}`, endTrimMm);
        if (boardType.upsized) {
          board.upsized      = true;
          board.neededQuarter = boardType.neededQuarter;
        }
        openBoards.push(board);
        boards.push(board);
        placement = findBestPlacementOnBoard(board, blank, spacingMm);
      }

      if (!placement) {
        unmetParts.push({
          partId:   blank.partId,
          partName: blank.name,
          reason:   "Placement solver failed to place this blank.",
        });
        continue;
      }

      placeBlankOnBoard(placement.board, blank, placement);
    }
  }

  const boardUsage = new Map();
  for (const board of boards) {
    const key  = boardKey(board);
    const prev = boardUsage.get(key) || {
      count: 0,
      boardFeetEach:    boardFeetForBoard(board),
      thicknessQuarter: board.thicknessQuarter,
      widthIn:          board.widthIn,
      lengthFt:         board.lengthFt,
    };
    prev.count += 1;
    boardUsage.set(key, prev);
  }

  const stockAreaMm2 = sum(boards.map((b) => b.widthMm * b.lengthMm));
  const usedAreaMm2  = sum(
    boards.flatMap((b) => b.placements.map((p) => p.widthMm * p.lengthMm))
  );
  const yieldPercent   = stockAreaMm2 ? (usedAreaMm2 / stockAreaMm2) * 100 : 0;
  const totalBoardFeet = sum(boards.map(boardFeetForBoard));
  const estimatedCost  = totalBoardFeet * inputs.pricePerBoardFoot;

  const upsizedBoards = boards.filter((b) => b.upsized);

  return {
    boards, unmetParts, upsizedBoards, boardUsage,
    stockAreaMm2, usedAreaMm2, yieldPercent,
    totalBoardFeet, estimatedCost,
  };
}

function createBoardFromType(type, id, endTrimMm) {
  const trim        = Math.max(0, Math.min(endTrimMm, Math.max(0, type.lengthMm - 1)));
  const trimOffset  = trim / 2;
  const usableLengthMm = Math.max(0, type.lengthMm - trim);
  return {
    id,
    source:           type.source,
    thicknessQuarter: type.thicknessQuarter,
    widthIn:          type.widthIn,
    lengthFt:         type.lengthFt,
    widthMm:          type.widthMm,
    lengthMm:         type.lengthMm,
    trimTotalMm:      trim,
    trimOffsetMm:     trimOffset,
    usableLengthMm,
    placements: [],
    freeRects:  [{ x: 0, y: trimOffset, w: type.widthMm, h: usableLengthMm }],
  };
}

// chooseBoardType — density-based scoring
// Instead of minimising waste-per-blank (which always picks the smallest board),
// we score each board type by how much of the remaining work can potentially fit
// on it. This naturally prefers longer/wider boards when many blanks remain,
// producing better yield across the full job.
function chooseBoardType(types, blank, spacingMm, endTrimMm, remainingBlanks = [], maxPlanerWidthMm = 0) {
  const eligible = [];

  for (const type of types) {
    if (type.remaining <= 0) continue;
    const usableLengthMm = type.lengthMm - endTrimMm;
    if (usableLengthMm <= EPSILON) continue;

    // Current blank must fit — otherwise this type is not eligible
    const fits = buildBlankOrientationOptions(blank).some(
      (opt) =>
        opt.widthMm  + spacingMm <= type.widthMm    + EPSILON &&
        opt.lengthMm + spacingMm <= usableLengthMm  + EPSILON
    );
    if (!fits) continue;
    eligible.push({ type, usableLengthMm });
  }

  if (!eligible.length) return null;

  // When a max planer width is set, try narrow boards first.
  // Only consider over-width boards if no narrow board can satisfy the blank.
  const narrowEligible = maxPlanerWidthMm > 0
    ? eligible.filter(({ type }) => type.widthMm <= maxPlanerWidthMm + EPSILON)
    : eligible;
  const pool = narrowEligible.length > 0 ? narrowEligible : eligible;

  let best = null;
  for (const { type, usableLengthMm } of pool) {
    const boardArea = type.widthMm * usableLengthMm;

    // Sum area of remaining blanks (including current) that could fit on this board type
    let fittableArea = 0;
    for (const rb of remainingBlanks) {
      if (
        buildBlankOrientationOptions(rb).some(
          (opt) =>
            opt.widthMm  + spacingMm <= type.widthMm    + EPSILON &&
            opt.lengthMm + spacingMm <= usableLengthMm  + EPSILON
        )
      ) {
        fittableArea += rb.widthMm * rb.lengthMm;
      }
    }

    // Fill ratio: how well remaining work fills this board. Capped at 1 (can't overfill).
    const fillRatio = Math.min(1, fittableArea / boardArea);

    if (!best || fillRatio > best.fillRatio) best = { fillRatio, type };
  }

  return best?.type ?? null;
}

function findBestPlacementAcrossBoards(boards, blank, spacingMm) {
  let best = null;
  for (const board of boards) {
    const candidate = findBestPlacementOnBoard(board, blank, spacingMm);
    if (!candidate) continue;
    if (!best || comparePlacementScores(candidate, best) < 0) best = candidate;
  }
  return best;
}

function findBestPlacementOnBoard(board, blank, spacingMm) {
  let best    = null;
  const options = buildBlankOrientationOptions(blank);

  for (let rectIndex = 0; rectIndex < board.freeRects.length; rectIndex++) {
    const rect = board.freeRects[rectIndex];
    for (const option of options) {
      const neededW = option.widthMm  + spacingMm;
      const neededL = option.lengthMm + spacingMm;
      if (neededW > rect.w + EPSILON || neededL > rect.h + EPSILON) continue;

      const shortFit = Math.min(rect.w - neededW, rect.h - neededL);
      const longFit  = Math.max(rect.w - neededW, rect.h - neededL);
      const areaFit  = rect.w * rect.h - neededW * neededL;

      const candidate = {
        board, rectIndex,
        x: rect.x, y: rect.y,
        widthMm:  option.widthMm,
        lengthMm: option.lengthMm,
        usedW:    neededW,
        usedL:    neededL,
        rotated:  option.rotated,
        shortFit, longFit, areaFit,
      };
      if (!best || comparePlacementScores(candidate, best) < 0) best = candidate;
    }
  }
  return best;
}

function comparePlacementScores(a, b) {
  if (a.shortFit !== b.shortFit) return a.shortFit - b.shortFit;
  if (a.longFit  !== b.longFit)  return a.longFit  - b.longFit;
  return a.areaFit - b.areaFit;
}

function placeBlankOnBoard(board, blank, placement) {
  const usedRect = { x: placement.x, y: placement.y, w: placement.usedW, h: placement.usedL };
  board.freeRects = splitFreeRects(board.freeRects, usedRect);
  board.freeRects = pruneContainedRects(board.freeRects).filter(
    (r) => r.w > EPSILON && r.h > EPSILON
  );
  board.placements.push({
    blankId:  blank.id,
    partId:   blank.partId,
    partName: blank.name,
    x: placement.x,
    y: placement.y,
    widthMm:  placement.widthMm,
    lengthMm: placement.lengthMm,
    rotated:  placement.rotated,
    grainLock: blank.grainLock,
    // Panel strip fields — only present when part is wider than available stock
    panelStripCount:  blank.panelStripCount,
    panelStripIndex:  blank.panelStripIndex,
    panelFullWidthMm: blank.panelFullWidthMm,
  });
}

function splitFreeRects(freeRects, usedRect) {
  const next = [];
  for (const rect of freeRects) {
    if (!rectanglesIntersect(rect, usedRect)) { next.push(rect); continue; }
    if (usedRect.x > rect.x + EPSILON) {
      next.push({ x: rect.x, y: rect.y, w: usedRect.x - rect.x, h: rect.h });
    }
    if (usedRect.x + usedRect.w < rect.x + rect.w - EPSILON) {
      next.push({ x: usedRect.x + usedRect.w, y: rect.y, w: rect.x + rect.w - (usedRect.x + usedRect.w), h: rect.h });
    }
    if (usedRect.y > rect.y + EPSILON) {
      next.push({ x: rect.x, y: rect.y, w: rect.w, h: usedRect.y - rect.y });
    }
    if (usedRect.y + usedRect.h < rect.y + rect.h - EPSILON) {
      next.push({ x: rect.x, y: usedRect.y + usedRect.h, w: rect.w, h: rect.y + rect.h - (usedRect.y + usedRect.h) });
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
    for (let i = 0; i < rects.length; i++) {
      if (i === index) continue;
      const other = rects[i];
      if (
        rect.x >= other.x - EPSILON &&
        rect.y >= other.y - EPSILON &&
        rect.x + rect.w <= other.x + other.w + EPSILON &&
        rect.y + rect.h <= other.y + other.h + EPSILON
      ) return false;
    }
    return true;
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Lumber yard suggestions
// ─────────────────────────────────────────────────────────────────────────────
function buildYardSuggestions(parts, inventory, kerfMm, maxTransportLengthFt = DEFAULTS.planningLengthMaxFt) {
  const suggestions      = [];
  const maxTransportMm   = maxTransportLengthFt * FOOT_TO_MM;

  const byQuarter = groupBy(
    parts.filter((p) => p.status === "ok" && p.stockQuarter),
    (p) => String(p.stockQuarter)
  );

  for (const row of inventory) {
    const boardLengthMm = row.lengthFt * FOOT_TO_MM;
    // Only suggest cuts when the board exceeds the max transport length
    if (boardLengthMm <= maxTransportMm + EPSILON) continue;

    const partsForQuarter = byQuarter.get(String(row.thicknessQuarter)) || [];
    if (!partsForQuarter.length) continue;

    const maxRequiredLenMm = Math.max(...partsForQuarter.map((p) => p.roughLengthMm));
    // Only suggest if all parts for this quarter fit within transport length
    if (maxRequiredLenMm > maxTransportMm + EPSILON) continue;

    const segmentsMm = splitBoardLength(boardLengthMm, kerfMm, maxTransportMm);
    suggestions.push({ row, maxRequiredLenMm, segmentsMm, maxTransportLengthFt });
  }

  return suggestions;
}

function splitBoardLength(totalLengthMm, kerfMm, targetMaxMm) {
  let cuts = Math.max(1, Math.ceil(totalLengthMm / targetMaxMm));
  while (cuts < 20) {
    const usable  = totalLengthMm - (cuts - 1) * kerfMm;
    const segment = usable / cuts;
    if (segment <= targetMaxMm + EPSILON) {
      return Array.from({ length: cuts }, () => roundTo(segment, 2));
    }
    cuts++;
  }
  return [roundTo(totalLengthMm, 2)];
}

// ─────────────────────────────────────────────────────────────────────────────
// Rendering
// ─────────────────────────────────────────────────────────────────────────────
function renderYardSuggestions(suggestions) {
  if (!suggestions.length) {
    dom.lumberYardSuggestions.innerHTML =
      '<p class="muted">No recut suggestions — all inventory boards are within the max transport length.</p>';
    return;
  }
  const maxFt = suggestions[0].maxTransportLengthFt ?? DEFAULTS.planningLengthMaxFt;
  dom.lumberYardSuggestions.innerHTML = `<h3>Lumber Yard Recut Suggestions <span class="muted" style="font-weight:400;font-size:0.88rem">(boards exceeding ${maxFt} ft transport limit)</span></h3>`;
  const list = document.createElement("ul");
  list.className = "compact";
  for (const sug of suggestions) {
    const item = document.createElement("li");
    const segsText = sug.segmentsMm
      .map((mm) => `${formatMm(mm, 1)} (${formatFeet(mmToFeet(mm), 2)})`)
      .join(" + ");
    item.textContent =
      `${sug.row.thicknessQuarter}/4 × ${formatInches(sug.row.widthIn)} × ${formatFeet(sug.row.lengthFt, 1)}: ` +
      `max required part length ${formatMm(sug.maxRequiredLenMm, 1)}. ` +
      `Suggested yard split: ${segsText}.`;
    list.append(item);
  }
  dom.lumberYardSuggestions.append(list);
}

function renderPartsSummary(parts) {
  const excluded = parts.filter((p) => p.excluded).length;
  const active   = parts.filter((p) => !p.excluded);
  const valid    = active.filter((p) => p.status === "ok").length;
  const invalid  = active.length - valid;
  const layers   = sum(active.map((p) => p.layers || 0));

  dom.partsSummary.innerHTML = "";
  const boxes = [
    `${parts.length} total parts detected`,
    `${active.length} active parts (${valid} with valid stock, ${invalid} unassigned)`,
    `${layers} total rough blanks including lamination layers`,
  ];
  if (excluded) boxes.push(`${excluded} part(s) excluded from planning`);
  for (const text of boxes) dom.partsSummary.append(summaryBox(text));
}

function renderPartsTable(parts, inputs) {
  dom.partsTableBody.innerHTML = "";
  const quarters     = inputs.thicknessOptionsQuarters;
  const sortedParts  = getSortedParts(parts);
  refreshSortHeaderStyles();

  for (const part of sortedParts) {
    const row = document.createElement("tr");
    if (part.excluded) row.classList.add("part-excluded");

    // Exclude checkbox (first column)
    const excludeCell  = document.createElement("td");
    const excludeInput = document.createElement("input");
    excludeInput.type    = "checkbox";
    excludeInput.checked = Boolean(part.excluded);
    excludeInput.title   = "Exclude this part from planning calculations";
    excludeInput.addEventListener("change", () => {
      const current = state.partOverrides[part.id] || {};
      if (excludeInput.checked) {
        state.partOverrides[part.id] = { ...current, excluded: true };
      } else {
        delete current.excluded;
        if (Object.keys(current).length) {
          state.partOverrides[part.id] = current;
        } else {
          delete state.partOverrides[part.id];
        }
      }
      updatePartsFromOverrides();
    });
    excludeCell.append(excludeInput);
    row.append(excludeCell);

    // Part name cell — append curved-surface badge when detected
    const nameCell = document.createElement("td");
    nameCell.title = `Raw X=${formatMm(part.rawMm.x)}, Raw Y=${formatMm(part.rawMm.y)}, Raw Z=${formatMm(part.rawMm.z)}`;
    nameCell.textContent = part.name;
    if (part.curved) {
      const badge = document.createElement("span");
      badge.className = "curved-badge";
      badge.title = "This part appears to have curved geometry. " +
        "The blank dimensions shown are the bounding box — correct for ordering stock, " +
        "but the finished part will require shaping (band saw, router, etc.) after rough milling.";
      badge.textContent = " ⌒ curved";
      nameCell.append(badge);
    }
    row.append(nameCell);
    appendTextCell(row, formatMm(part.netLengthMm));
    appendTextCell(row, formatMm(part.netWidthMm));
    appendTextCell(row, formatMm(part.netThicknessMm));
    appendTextCell(row, formatMm(part.roughLengthMm));
    appendTextCell(row, formatMm(part.roughWidthMm));
    appendTextCell(row, formatMm(part.roughThicknessMm));
    appendTextCell(row,
      part.stockQuarter
        ? `${part.stockQuarter}/4 (${formatMm(part.stockThicknessMm, 1)})`
        : `— ${part.reason}`
    );
    appendTextCell(row, part.layers ? String(part.layers) : "—");

    // Grain direction select — Long / Mid / Free
    const grainCell   = document.createElement("td");
    const grainSelect = document.createElement("select");
    grainSelect.title = "Long: grain along longest axis (no rotation). Mid: grain along middle axis (blank pre-rotated). Free: nesting may rotate.";
    for (const [val, label] of [["long", "Long"], ["mid", "Mid"], ["free", "Free"]]) {
      const opt = document.createElement("option");
      opt.value       = val;
      opt.textContent = label;
      grainSelect.append(opt);
    }
    grainSelect.value = part.grainDir || "long";
    grainSelect.addEventListener("change", () => {
      const current = state.partOverrides[part.id] || {};
      state.partOverrides[part.id] = { ...current, grainDir: grainSelect.value };
      updatePartsFromOverrides();
    });
    grainCell.append(grainSelect);
    row.append(grainCell);

    // Per-part thickness override select
    const overrideCell   = document.createElement("td");
    const overrideSelect = document.createElement("select");
    const autoOpt        = document.createElement("option");
    autoOpt.value       = "auto";
    autoOpt.textContent = "Auto";
    overrideSelect.append(autoOpt);
    for (const quarter of quarters) {
      const opt = document.createElement("option");
      opt.value = String(quarter);
      opt.textContent = `${quarter}/4`;
      overrideSelect.append(opt);
    }
    overrideSelect.value = part.thicknessOverrideQuarter == null ? "auto" : String(part.thicknessOverrideQuarter);
    overrideSelect.addEventListener("change", () => {
      const next    = overrideSelect.value === "auto" ? null : Number(overrideSelect.value);
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
  if (!state.rawParts.length) return;
  const inputs = collectInputs();
  state.parts  = assignPartsForStock(state.rawParts, inputs, state.partOverrides);
  renderPartsSummary(state.parts);
  renderPartsTable(state.parts, inputs);
  clearResults();
  setStatus("Part overrides updated.", "ok");
}

function appendTextCell(row, text, title = "", col = "") {
  const cell = document.createElement("td");
  cell.textContent = text;
  if (title) cell.title = title;
  if (col)   cell.dataset.col = col;
  row.append(cell);
}

function getSortedParts(parts) {
  const { key, direction } = state.sort;
  const sign   = direction === "asc" ? 1 : -1;
  return [...parts].sort((a, b) => {
    const av = getPartSortValue(a, key);
    const bv = getPartSortValue(b, key);
    if (typeof av === "number" && typeof bv === "number") {
      if (Number.isNaN(av) && Number.isNaN(bv)) return 0;
      if (Number.isNaN(av)) return 1;
      if (Number.isNaN(bv)) return -1;
      return (av - bv) * sign;
    }
    return String(av).localeCompare(String(bv), undefined, { sensitivity: "base" }) * sign;
  });
}

function getPartSortValue(part, key) {
  switch (key) {
    case "name":                 return part.name || "";
    case "netLengthMm":
    case "netWidthMm":
    case "netThicknessMm":
    case "roughLengthMm":
    case "roughWidthMm":
    case "roughThicknessMm":     return Number(part[key] ?? Number.NaN);
    case "stockQuarter":         return part.stockQuarter == null ? Infinity : Number(part.stockQuarter);
    case "layers":               return Number(part.layers ?? Number.NaN);
    case "grainDir":             return part.grainDir || "";
    case "thicknessOverrideQuarter":
      return part.thicknessOverrideQuarter == null ? Infinity : Number(part.thicknessOverrideQuarter);
    case "orientation":          return part.orientation || "";
    default:                     return part.name || "";
  }
}

function initPartsSorting() {
  for (const th of dom.partsHeaderSortables) {
    th.addEventListener("click", () => {
      const key = th.dataset.sortKey;
      if (!key) return;
      if (state.sort.key === key) {
        state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
      } else {
        state.sort.key       = key;
        state.sort.direction = "asc";
      }
      if (state.parts.length) renderPartsTable(state.parts, collectInputs());
    });
  }
  refreshSortHeaderStyles();
}

function refreshSortHeaderStyles() {
  for (const th of dom.partsHeaderSortables) {
    const active = th.dataset.sortKey === state.sort.key;
    th.classList.toggle("sort-active", active);
    th.title = active
      ? `Sorted ${state.sort.direction} (click to toggle)`
      : "Click to sort";
  }
}

function renderPlanSummary(target, result, title, pricePerBoardFoot) {
  target.innerHTML = "";
  const root = document.createElement("div");
  root.className = "summary-grid";

  root.append(
    summaryBox(title),
    summaryBox(`${result.boards.length} boards used, ${formatNumber(result.totalBoardFeet, 2)} board feet total`),
    summaryBox(`Estimated lumber cost: ${formatCurrency(result.estimatedCost)}`),
    summaryBox(
      `Used area: ${formatNumber(result.usedAreaMm2 / 1_000_000, 3)} m² of ` +
      `${formatNumber(result.stockAreaMm2 / 1_000_000, 3)} m² (${formatNumber(result.yieldPercent, 1)}% yield)`
    )
  );

  if (result.upsizedBoards?.length) {
    const groups = new Map();
    for (const b of result.upsizedBoards) {
      const k = `${b.neededQuarter}/4→${b.thicknessQuarter}/4`;
      groups.set(k, (groups.get(k) || 0) + 1);
    }
    const detail = [...groups.entries()].map(([k, n]) => `${n}× ${k}`).join(", ");
    root.append(summaryBox(
      `${result.upsizedBoards.length} board(s) use thicker-than-needed stock (inventory gap): ${detail}. ` +
      `Extra thickness will be milled off — wasteful but parts will fit.`,
      "warning"
    ));
  }

  if (result.unmetParts.length) {
    root.append(summaryBox(`${result.unmetParts.length} part(s) are currently unmet`, "warning"));
  } else {
    root.append(summaryBox("All parts are successfully allocated.", "ok"));
  }

  // Board usage table with totals row
  const usageTable = document.createElement("table");
  usageTable.innerHTML =
    "<thead><tr><th>Stock Size</th><th>Board Count</th><th>Board Feet</th><th>Estimated Cost</th></tr></thead>";
  const usageBody  = document.createElement("tbody");
  const usageFoot  = document.createElement("tfoot");

  let totalBoards = 0, totalBf = 0, totalCost = 0;

  for (const [key, entry] of [...result.boardUsage.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    const lineBf   = entry.boardFeetEach * entry.count;
    const lineCost = lineBf * pricePerBoardFoot;
    totalBoards += entry.count;
    totalBf     += lineBf;
    totalCost   += lineCost;

    const row = document.createElement("tr");
    row.innerHTML = `<td>${key}</td><td>${entry.count}</td><td>${formatNumber(lineBf, 2)}</td><td>${formatCurrency(lineCost)}</td>`;
    usageBody.append(row);
  }

  const totalRow = document.createElement("tr");
  totalRow.innerHTML = `<td><strong>Total</strong></td><td><strong>${totalBoards}</strong></td><td><strong>${formatNumber(totalBf, 2)}</strong></td><td><strong>${formatCurrency(totalCost)}</strong></td>`;
  usageFoot.append(totalRow);

  usageTable.append(usageBody, usageFoot);
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

  const entries    = [];
  let   extraBf    = 0;
  for (const [key, entry] of additionalPlan.boardUsage.entries()) {
    const bf = entry.boardFeetEach * entry.count;
    extraBf += bf;
    entries.push(`${entry.count} × ${key}`);
  }
  block.textContent =
    `Additional boards suggested for unmet parts: ${entries.join("; ")}. ` +
    `Extra board feet: ${formatNumber(extraBf, 2)}. ` +
    `Extra estimated cost: ${formatCurrency(extraBf * pricePerBoardFoot)}.`;
  target.append(block);
}

// ─────────────────────────────────────────────────────────────────────────────
// Workshop tab — cut guide
// ─────────────────────────────────────────────────────────────────────────────

function renderWorkshopTab() {
  if (!dom.workshopContent) return;

  // Prefer inventory result (real boards); fall back to planning result.
  const result = state.inventoryResult || state.planningResult;

  if (!result || !result.boards.length) {
    dom.workshopContent.innerHTML = "";
    if (dom.workshopSourceNote) {
      dom.workshopSourceNote.textContent =
        "Run Plan Stock or Recalculate to generate the workshop guide.";
    }
    return;
  }

  const sourceLabel = state.inventoryResult ? "Lumber Yard Recalculate" : "Plan Stock";
  if (dom.workshopSourceNote) {
    dom.workshopSourceNote.innerHTML =
      `Showing cut guide from <strong>${sourceLabel}</strong>. ` +
      `Run Recalculate after visiting the lumber yard to get board-specific instructions.`;
  }

  // Build a fast lookup: partId → part data
  const partsMap = new Map(state.parts.map((p) => [p.id, p]));
  const maxPlanerWidthIn = collectInputs().maxPlanerWidthIn;

  // Global draw scale (same logic as renderLayouts — widest board = 180 px)
  const maxWidthMm = Math.max(...result.boards.map((b) => b.widthMm), 1);
  const drawScale  = (180 / maxWidthMm) * state.layoutScale;

  dom.workshopContent.innerHTML = "";

  result.boards.forEach((board) => {
    const card = document.createElement("div");
    card.className = "workshop-board-card";

    // ── Header ──────────────────────────────────────────────────
    const title = document.createElement("h3");
    title.textContent =
      `${board.id} · ${board.thicknessQuarter}/4 × ${formatInches(board.widthIn)} × ${formatFeet(board.lengthFt, 1)}`;
    card.append(title);

    const sub = document.createElement("p");
    sub.className = "muted";
    sub.style.margin = "0 0 8px";
    sub.textContent =
      `${formatMm(board.widthMm, 0)} wide × ${formatMm(board.lengthMm, 0)} long · ${board.source}` +
      (board.upsized ? ` · ⚠ upsized from ${board.neededQuarter}/4` : "");
    card.append(sub);

    // ── Board diagram (same SVG as layout view) ─────────────────
    const svgHeight = Math.max(50, board.widthMm * drawScale);
    const svg = buildBoardSvg(board, svgHeight);
    card.append(svg);

    // ── Parts table ─────────────────────────────────────────────
    const partsSectionHead = document.createElement("div");
    partsSectionHead.className = "workshop-section-heading";
    partsSectionHead.textContent = "Parts on this board";
    card.append(partsSectionHead);

    card.append(buildWorkshopPartsTable(board, partsMap));

    // ── Cut sequence ────────────────────────────────────────────
    const cutHead = document.createElement("div");
    cutHead.className = "workshop-section-heading";
    cutHead.textContent = "Recommended cut sequence";
    card.append(cutHead);

    const steps = buildCutSequence(board, partsMap, maxPlanerWidthIn);
    const ol = document.createElement("ol");
    ol.className = "workshop-steps";
    steps.forEach((step, i) => {
      const li = document.createElement("li");
      li.className = "workshop-step";

      const num = document.createElement("span");
      num.className = "workshop-step-num";
      num.textContent = `${i + 1}.`;

      const tool = document.createElement("span");
      tool.className = `workshop-tool ${step.toolClass}`;
      tool.textContent = step.tool;

      const txt = document.createElement("span");
      txt.textContent = step.text;

      li.append(num, tool, txt);
      ol.append(li);
    });
    card.append(ol);

    // ── Final milling reference ──────────────────────────────────
    const finalBox = buildFinalMillingBox(board, partsMap);
    if (finalBox) card.append(finalBox);

    dom.workshopContent.append(card);
  });

  // ── Consolidated schedule ────────────────────────────────────
  const phases = buildConsolidatedSchedule(result, partsMap, maxPlanerWidthIn);
  if (phases.length) {
    dom.workshopContent.append(renderConsolidatedSchedule(phases));
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Workshop PDF print
// ─────────────────────────────────────────────────────────────────────────────

function printWorkshopPDF() {
  const allCards = dom.workshopContent?.querySelectorAll(".workshop-board-card");
  if (!allCards || !allCards.length) {
    setStatus("No workshop boards to print. Run Plan Stock or Recalculate first.", "error");
    return;
  }

  const projectName = (dom.projectName.value || "").trim() || "Workshop Plan";
  const cssHref     = document.querySelector('link[rel="stylesheet"]')?.href ?? "";

  // Format date/time for footer
  const now = new Date();
  const dateStr = now.toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
  const footerText = `${projectName} — Workshop Cut Guide · Printed ${dateStr} at ${timeStr}`;

  // Board cards (all except the consolidated schedule card, which gets its own last page)
  const boardCards = [...allCards].filter((c) => !c.classList.contains("workshop-consolidated-card"));
  const consolidatedCard = dom.workshopContent?.querySelector(".workshop-consolidated-card");

  // Each board card = one page; consolidated card = final page
  const pageSections = [
    ...boardCards.map((card) =>
      `<section class="workshop-board-card print-board" style="page-break-after:always;break-after:page;">${card.innerHTML}</section>`
    ),
    consolidatedCard
      ? `<section class="workshop-board-card print-board">${consolidatedCard.innerHTML}</section>`
      : "",
  ].join("\n");

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>${projectName} — Workshop Plan</title>
  ${cssHref ? `<link rel="stylesheet" href="${cssHref}">` : ""}
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      font-size: 13px;
      color: #2c2416;
      margin: 0;
      padding: 0 0 28px;
      background: #fff;
    }
    .workshop-board-card.print-board {
      border: none;
      border-radius: 0;
      padding: 20px 28px;
      box-shadow: none;
      background: #fff;
    }
    .print-footer {
      display: none;
    }
    @media print {
      body { margin: 0; padding: 0 0 32px; }
      .workshop-board-card.print-board {
        padding: 12px 20px;
        border: none !important;
        box-shadow: none !important;
        background: #fff !important;
      }
      .workshop-tool {
        border: 1px solid #999 !important;
        background: #fff !important;
        color: #333 !important;
      }
      .print-footer {
        display: block;
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        padding: 4px 20px;
        font-size: 9px;
        color: #888;
        border-top: 1px solid #ddd;
        background: #fff;
        text-align: center;
      }
    }
  </style>
</head>
<body>
  ${pageSections}
  <div class="print-footer">${footerText}</div>
  <script>
    window.addEventListener("load", function() { window.print(); });
    setTimeout(function() { window.print(); }, 800);
  <\/script>
</body>
</html>`;

  const win = window.open("", "_blank", "width=900,height=700");
  if (!win) {
    setStatus("Pop-up blocked — please allow pop-ups for this page and try again.", "error");
    return;
  }
  win.document.write(html);
  win.document.close();
}

// Build the same SVG used in renderLayouts but returns the element (no scale row).
function buildBoardSvg(board, svgHeight) {
  const colors = [
    "#bc6c25","#dda15e","#606c38","#283618",
    "#7f5539","#9c6644","#386641","#1d3557",
    "#6d597a","#2a9d8f",
  ];

  const svg = document.createElementNS("http://www.w3.org/2000/svg","svg");
  svg.setAttribute("class","board-svg");
  svg.setAttribute("viewBox",`0 0 ${board.lengthMm} ${board.widthMm}`);
  svg.setAttribute("preserveAspectRatio","none");
  svg.style.height = `${svgHeight}px`;

  const bg = document.createElementNS("http://www.w3.org/2000/svg","rect");
  bg.setAttribute("x","0"); bg.setAttribute("y","0");
  bg.setAttribute("width",String(board.lengthMm));
  bg.setAttribute("height",String(board.widthMm));
  bg.setAttribute("fill","#f4e6ce");
  bg.setAttribute("stroke","#a48a6a");
  bg.setAttribute("stroke-width",String(Math.max(0.8, board.widthMm * 0.008)));
  svg.append(bg);

  if (board.trimTotalMm > EPSILON) {
    for (const [x, w] of [
      [0, board.trimOffsetMm],
      [board.lengthMm - board.trimOffsetMm, board.trimOffsetMm],
    ]) {
      const trim = document.createElementNS("http://www.w3.org/2000/svg","rect");
      trim.setAttribute("x",String(x)); trim.setAttribute("y","0");
      trim.setAttribute("width",String(w)); trim.setAttribute("height",String(board.widthMm));
      trim.setAttribute("fill","#d7c4a8"); trim.setAttribute("fill-opacity","0.55");
      svg.append(trim);
    }
  }

  board.placements.forEach((placement, idx) => {
    const svgX = placement.y;
    const svgY = placement.x;
    const svgW = placement.lengthMm;
    const svgH = placement.widthMm;

    const rect = document.createElementNS("http://www.w3.org/2000/svg","rect");
    rect.setAttribute("x",String(svgX)); rect.setAttribute("y",String(svgY));
    rect.setAttribute("width",String(svgW)); rect.setAttribute("height",String(svgH));
    rect.setAttribute("fill", colors[idx % colors.length]);
    rect.setAttribute("fill-opacity","0.86");
    rect.setAttribute("stroke","#ffffff");
    rect.setAttribute("stroke-width",String(Math.max(0.6, board.widthMm * 0.006)));
    svg.append(rect);

    const labelPad = svgW * 0.03;
    const fontSize = Math.min(
      Math.max(4, svgH * 0.45),
      Math.max(4, svgW * 0.06),
      Math.max(4, board.widthMm * 0.10)
    );
    const label = document.createElementNS("http://www.w3.org/2000/svg","text");
    label.setAttribute("x",String(svgX + labelPad));
    label.setAttribute("y",String(svgY + svgH / 2));
    label.setAttribute("text-anchor","start");
    label.setAttribute("dominant-baseline","central");
    label.setAttribute("font-size",String(fontSize));
    label.setAttribute("fill","#fff");
    label.textContent = shortenPartName(placement.partName);
    svg.append(label);
  });

  return svg;
}

// Parts mini-table: one row per placement showing rough + net dims.
function buildWorkshopPartsTable(board, partsMap) {
  const table = document.createElement("table");
  table.className = "workshop-parts-table";

  const thead = table.createTHead();
  const hr = thead.insertRow();
  for (const h of ["Part", "Rough L × W × T", "Net L × W × T", "Grain"]) {
    const th = document.createElement("th");
    th.textContent = h;
    hr.append(th);
  }

  const tbody = table.createTBody();
  // Sort placements left-to-right (by y = position along board length)
  const sorted = [...board.placements].sort((a, b) => a.y - b.y);

  for (const p of sorted) {
    const part = partsMap.get(p.partId);
    const tr = tbody.insertRow();

    const nameTd = tr.insertCell();
    nameTd.textContent = p.partName; // includes lam note when relevant

    const roughTd = tr.insertCell();
    if (part) {
      const isPanelStrip  = p.panelStripCount > 1;
      const isLaminated   = part.layers > 1;
      const perLayerThick = isLaminated ? roundTo(part.roughThicknessMm / part.layers, 1) : part.roughThicknessMm;
      const thickStr      = isLaminated
        ? `${formatMm(perLayerThick, 0)}/layer (${part.layers} layers → ${formatMm(part.roughThicknessMm, 0)} after glue-up)`
        : formatMm(part.roughThicknessMm, 0);
      roughTd.textContent = isPanelStrip
        ? `${formatMm(p.lengthMm, 0)} × ${formatMm(p.widthMm, 0)} × ${thickStr}` +
          ` (strip ${p.panelStripIndex}/${p.panelStripCount} → panel ${formatMm(p.panelFullWidthMm, 0)} W)`
        : `${formatMm(p.lengthMm, 0)} × ${formatMm(p.widthMm, 0)} × ${thickStr}`;
    } else {
      roughTd.textContent = `${formatMm(p.lengthMm, 0)} × ${formatMm(p.widthMm, 0)} × —`;
    }

    const netTd = tr.insertCell();
    netTd.textContent = part
      ? `${formatMm(part.netLengthMm, 0)} × ${formatMm(part.netWidthMm, 0)} × ${formatMm(part.netThicknessMm, 0)}`
      : "—";

    const grainTd = tr.insertCell();
    grainTd.textContent = part
      ? ({ long: "Long ▶", mid: "Mid ▶", free: "Free ↻" }[part.grainDir] ?? "Long ▶")
      : "—";
  }

  return table;
}

// ─────────────────────────────────────────────────────────────────────────────
// Shared workshop helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Group placements on a board into cross-cut sections.
 * Placements whose y-ranges overlap share a section (can be cut from one piece).
 */
function buildSections(board) {
  const sorted = [...board.placements].sort((a, b) => a.y - b.y);
  const sections = [];
  for (const p of sorted) {
    const pEnd = p.y + p.lengthMm;
    const last  = sections[sections.length - 1];
    if (last && p.y < last.endY + EPSILON) {
      last.endY = Math.max(last.endY, pEnd);
      last.placements.push(p);
    } else {
      sections.push({ startY: p.y, endY: pEnd, placements: [p] });
    }
  }
  return sections;
}

/** Group an array into a Map keyed by keyFn(item). */
function groupBy(arr, keyFn) {
  const map = new Map();
  for (const item of arr) {
    const key = keyFn(item);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(item);
  }
  return map;
}

/** Max rough thickness needed across all placements on a board (full assembled thickness). */
function maxRoughThickForBoard(board, partsMap) {
  return board.placements.reduce((mx, p) => {
    const pt = partsMap.get(p.partId);
    return Math.max(mx, pt ? pt.roughThicknessMm : 0);
  }, 0) || quarterToMm(board.thicknessQuarter);
}

/**
 * Per-layer planing target for a board.
 * For single-layer parts this equals roughThicknessMm.
 * For laminated parts each blank is planed to roughThicknessMm / layers,
 * then layers are glued together AFTER cutting.
 */
function planeThickForBoard(board, partsMap) {
  return board.placements.reduce((mx, p) => {
    const pt = partsMap.get(p.partId);
    if (!pt) return mx;
    const perLayer = roundTo(pt.roughThicknessMm / Math.max(1, pt.layers), 1);
    return Math.max(mx, perLayer);
  }, 0) || quarterToMm(board.thicknessQuarter);
}

/**
 * Generate a step-by-step cut sequence for a single board.
 * Steps are objects: { tool, toolClass, text }
 */
function buildCutSequence(board, partsMap, maxPlanerWidthIn = 0) {
  const steps   = [];
  const spacer  = (tool, toolClass, text) => steps.push({ tool, toolClass, text });

  const stockMm          = quarterToMm(board.thicknessQuarter);
  const maxRoughThick    = maxRoughThickForBoard(board, partsMap);
  const maxPlanerWidthMm = maxPlanerWidthIn * INCH_TO_MM;
  // planeThickMm is the target for the PLANER at this stage.
  // For laminated parts the blanks are planed to roughThicknessMm / layers — the layers
  // are glued together AFTER cutting, then flattened as a glued assembly.
  const planeThickMm  = planeThickForBoard(board, partsMap);
  const trimEach      = board.trimOffsetMm ?? 25.4;
  const hasLam        = board.placements.some((p) => (partsMap.get(p.partId)?.layers ?? 1) > 1);

  // ── Initial milling ──────────────────────────────────────────
  spacer("Inspect", "tool-inspect",
    "Check for cupping, bowing, twist, and surface defects. " +
    "Mark any knots or checks to route around when laying out blanks.");

  spacer("Jointer", "tool-jointer",
    "Face joint one face flat. This becomes your reference face (face against the planer bed).");

  // Re-saw suggestion: if stock is more than ~12 mm thicker than the per-layer target
  const resawExcess = stockMm - planeThickMm;
  if (resawExcess > 12) {
    const resawTarget = roundTo(planeThickMm + 3, 0.5);
    spacer("Band saw", "tool-bandsaw",
      `Re-saw to ≈${formatMm(resawTarget, 0)} — stock is ${formatMm(stockMm, 0)} ` +
      `but blanks only need ${formatMm(planeThickMm, 0)} before glue-up. ` +
      `Re-sawing saves ${formatMm(resawExcess - 3, 0)} of planer travel. ` +
      `Save the off-cut for thinner parts.`);
    spacer("Jointer", "tool-jointer",
      "Light face-joint pass on the re-sawn face to remove saw marks before planing.");
  }

  const lamThickNote = hasLam
    ? ` (per-layer target — blanks will be glued up to full thickness after cutting)`
    : "";
  const boardOverWidth = maxPlanerWidthMm > 0 && board.widthMm > maxPlanerWidthMm + EPSILON;
  const multiPassNote = boardOverWidth
    ? ` ⚠ Board is ${formatInches(board.widthIn)} wide — wider than your ${formatInches(maxPlanerWidthIn)}" planer capacity. ` +
      `Rip the board into strips ≤ ${formatInches(maxPlanerWidthIn)}" wide before planing, ` +
      `then edge-glue them back to width after planing if needed. ` +
      `Alternatively, use a wide drum sander or hand planes.`
    : "";
  spacer("Planer", "tool-planer",
    `Plane to ${formatMm(planeThickMm, 1)}${lamThickNote}. ` +
    `Take light passes (≤ 1 mm each). Flip between faces to keep even tension.${multiPassNote}`);

  spacer("Jointer", "tool-jointer",
    "Joint one long edge straight. This is your reference edge (fence against the rip fence).");

  spacer("Miter saw", "tool-mitersaw",
    `Trim the reference end only: cut ≈${formatMm(trimEach, 0)} to square up and remove end checks. ` +
    `Leave the far end intact for now — it will be cleaned up as part of the final cross-cut. ` +
    `All blank measurements below are taken from this trimmed reference end.`);

  // ── Blank cuts ───────────────────────────────────────────────
  const sections = buildSections(board);

  if (sections.length > 1) {
    spacer("Note", "tool-note",
      `Cross-cut the board into ${sections.length} sections first, then rip each section to width. ` +
      `Measurements below are from the trimmed reference end.`);
  }

  sections.forEach((sec, si) => {
    // Cross-cut position from trimmed reference end
    const crossCutPos = roundTo(sec.endY - trimEach, 0.5);
    const sectionLen  = roundTo(sec.endY - sec.startY, 0.5);
    const names = sec.placements.map((p) => shortenPartName(p.partName)).join(", ");

    if (sections.length > 1) {
      spacer("Miter saw", "tool-mitersaw",
        `Cross-cut section ${si + 1} at ${formatMm(crossCutPos, 0)} from reference end ` +
        `— yields a ${formatMm(sectionLen, 0)}-long piece containing: ${names}.`);
    }

    // Within the section, rip each blank in order of x (width position)
    const byX = [...sec.placements].sort((a, b) => a.x - b.x);
    byX.forEach((p, pi) => {
      const part = partsMap.get(p.partId);
      // Thickness of THIS blank = per-layer target (for laminated parts) or full rough thickness
      const blankThickMm = part && part.layers > 1
        ? roundTo(part.roughThicknessMm / part.layers, 1)
        : (part ? part.roughThicknessMm : planeThickMm);
      const lamNote = (part && part.layers > 1)
        ? ` · layer ${p.blankId?.split("-L")[1] ?? (pi + 1)} of ${part.layers} (glue-up after all blanks are cut)`
        : "";
      const crossCutNote = sectionLen > p.lengthMm + 1
        ? ` (cross-cut to ${formatMm(p.lengthMm, 0)} first)`
        : "";

      spacer("Table saw", "tool-tablesaw",
        `Rip to ${formatMm(p.widthMm, 0)} wide${crossCutNote} → ` +
        `blank: ${formatMm(p.lengthMm, 0)} L × ${formatMm(p.widthMm, 0)} W × ` +
        `${formatMm(blankThickMm, 0)} T — "${p.partName}"${lamNote}.`);
    });
  });

  // ── Panel glue-up note ───────────────────────────────────────
  // Collect unique panel parts on this board (identified by partId + strip count).
  const panelPartsOnBoard = new Map();
  for (const p of board.placements) {
    if (p.panelStripCount > 1 && !panelPartsOnBoard.has(p.partId)) {
      const pt = partsMap.get(p.partId);
      panelPartsOnBoard.set(p.partId, {
        name:       pt ? pt.name : p.partName,
        count:      p.panelStripCount,
        fullWidthMm: p.panelFullWidthMm,
        stripsHere: board.placements.filter((q) => q.partId === p.partId).length,
      });
    }
  }
  if (panelPartsOnBoard.size) {
    for (const panel of panelPartsOnBoard.values()) {
      const allOnThis = panel.stripsHere === panel.count;
      const locationNote = allOnThis
        ? `All ${panel.count} strips are on this board.`
        : `${panel.stripsHere} of ${panel.count} strips are on this board — collect remaining strips from other boards.`;
      spacer("Jointer", "tool-jointer",
        `Edge-joint one face of each "${panel.name}" strip (the faces that will be glued). ` +
        `${locationNote}`);
      spacer("Note", "tool-note",
        `Dry-fit all ${panel.count} strips, then glue up to form a panel ` +
        `≈${formatMm(panel.fullWidthMm, 0)} wide. Clamp overnight. ` +
        `After cure, face-joint and plane the panel to rough thickness before final milling.`);
    }
  }

  // ── Lamination glue-up (comes after ALL blanks are cut) ─────
  const lamGroups = new Map(); // partId → { name, layers, roughThickMm, perLayerMm }
  for (const p of board.placements) {
    const pt = partsMap.get(p.partId);
    if (pt && pt.layers > 1 && !lamGroups.has(p.partId)) {
      lamGroups.set(p.partId, {
        name:         pt.name,
        layers:       pt.layers,
        roughThickMm: pt.roughThicknessMm,
        perLayerMm:   roundTo(pt.roughThicknessMm / pt.layers, 1),
        stripsOnBoard: board.placements.filter((q) => q.partId === p.partId).length,
      });
    }
  }
  if (lamGroups.size) {
    const lamNames = [...lamGroups.values()].map((l) => `"${l.name}"`).join(", ");
    spacer("Note", "tool-note",
      `Collect all layer blanks for ${lamNames} — they may be spread across multiple boards. ` +
      `Do NOT begin glue-up until every layer blank is cut.`);
    for (const lam of lamGroups.values()) {
      spacer("Jointer", "tool-jointer",
        `"${lam.name}" — edge-joint the mating face of each of the ${lam.layers} layer blanks ` +
        `(the face that will contact the next layer).`);
      spacer("Note", "tool-note",
        `Dry-fit all ${lam.layers} layers of "${lam.name}" and check alignment. ` +
        `Apply glue to mating faces, assemble, and clamp overnight. ` +
        `Each layer is ≈${formatMm(lam.perLayerMm, 1)} thick; ` +
        `glued assembly target: ≈${formatMm(lam.roughThickMm, 1)} rough.`);
      spacer("Jointer", "tool-jointer",
        `After cure — scrape squeeze-out from "${lam.name}", then take one light face-joint pass ` +
        `on both faces to flatten. Final rough thickness: ${formatMm(lam.roughThickMm, 1)}.`);
    }
  }

  return steps;
}

/**
 * Build a "Final Milling Reference" box showing net target dimensions per unique part.
 */
function buildFinalMillingBox(board, partsMap) {
  const seenIds = new Set();
  const rows    = [];

  const sorted = [...board.placements].sort((a, b) => a.y - b.y);
  for (const p of sorted) {
    if (seenIds.has(p.partId)) continue;
    seenIds.add(p.partId);
    const part = partsMap.get(p.partId);
    if (!part) continue;
    rows.push(part);
  }
  if (!rows.length) return null;

  const box = document.createElement("div");
  box.className = "workshop-final-milling";

  const heading = document.createElement("h4");
  heading.textContent = "Final milling — target net dimensions";
  box.append(heading);

  const table = document.createElement("table");
  table.className = "workshop-parts-table";
  table.style.background = "transparent";

  const thead = table.createTHead();
  const hr = thead.insertRow();
  for (const h of ["Part", "Plane to (T)", "Rip to (W)", "Cross-cut to (L)"]) {
    const th = document.createElement("th");
    th.textContent = h;
    hr.append(th);
  }

  const tbody = table.createTBody();
  for (const part of rows) {
    const tr = tbody.insertRow();
    tr.insertCell().textContent = part.name;
    tr.insertCell().textContent = formatMm(part.netThicknessMm, 1);
    tr.insertCell().textContent = formatMm(part.netWidthMm, 1);
    tr.insertCell().textContent = formatMm(part.netLengthMm, 1);
  }
  box.append(table);

  const note = document.createElement("p");
  note.className = "muted";
  note.style.marginTop = "6px";
  note.textContent =
    "Mill to net dimensions in this order: face joint → plane to net T → " +
    "joint edge → rip to net W → cross-cut to net L.";
  box.append(note);

  return box;
}

// ─────────────────────────────────────────────────────────────────────────────
// Consolidated Mill Schedule
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Build a cross-board milling schedule optimised to minimise tool adjustments.
 * Returns an array of phase objects: { tool, toolClass, heading, note, groups[] }
 * Each group: { setting, items[] }  where item: { label, detail }
 */
function buildConsolidatedSchedule(result, partsMap, maxPlanerWidthIn = 0) {
  const boards = result.boards;
  if (!boards.length) return [];
  const maxPlanerWidthMm = maxPlanerWidthIn * INCH_TO_MM;

  const phases = [];
  const boardDesc = (b) =>
    `${b.thicknessQuarter}/4 × ${formatInches(b.widthIn)} × ${formatFeet(b.lengthFt, 1)}`;

  // ── Phase 1: Inspect all boards ─────────────────────────────────────────
  phases.push({
    tool: "Inspect", toolClass: "tool-inspect",
    heading: "Inspect all boards",
    note: "Before any milling, check every board for cupping, bowing, twist, and surface defects. Mark knots and checks to work around during layout.",
    groups: [{ setting: null, items: boards.map((b) => ({ label: b.id, detail: boardDesc(b) })) }],
  });

  // ── Phase 2: Face joint all boards ──────────────────────────────────────
  phases.push({
    tool: "Jointer", toolClass: "tool-jointer",
    heading: "Face joint all boards",
    note: "One flat reference face per board. Mark the jointed face. No fence or depth change between boards.",
    groups: [{ setting: null, items: boards.map((b) => ({ label: b.id, detail: boardDesc(b) })) }],
  });

  // ── Phase 3: Re-saw oversize boards (optional) ───────────────────────────
  const resawItems = boards.map((b) => {
    const stockMm = quarterToMm(b.thicknessQuarter);
    const needed  = maxRoughThickForBoard(b, partsMap);
    const excess  = stockMm - needed;
    if (excess <= 12) return null;
    const target = roundTo(needed + 3, 1);
    return { boardId: b.id, stockMm, needed, target };
  }).filter(Boolean);

  if (resawItems.length) {
    const byTarget = groupBy(resawItems, (r) => roundTo(r.target, 1));
    phases.push({
      tool: "Band saw", toolClass: "tool-bandsaw",
      heading: "Re-saw oversize boards — thickest target first",
      note: "Do all re-saws before planing. Move blade down only. Save off-cuts for thinner parts.",
      groups: [...byTarget.entries()].sort((a, b) => b[0] - a[0]).map(([t, items]) => ({
        setting: `Re-saw to ≈${formatMm(t, 0)}`,
        items: items.map((r) => ({
          label: r.boardId,
          detail: `${formatMm(r.stockMm, 0)} stock → saves ${formatMm(r.stockMm - r.target, 0)} of planer travel`,
        })),
      })),
    });
    phases.push({
      tool: "Jointer", toolClass: "tool-jointer",
      heading: "Light face-joint re-sawn faces",
      note: "One light pass on each re-sawn face only, to remove saw marks before planing.",
      groups: [{ setting: null, items: resawItems.map((r) => ({ label: r.boardId, detail: "re-sawn face" })) }],
    });
  }

  // ── Phase 4: Plane to per-layer thickness — thickest group first ─────────
  // For laminated parts, each blank is planed to roughThicknessMm / layers.
  // The blanks are glued up to full thickness AFTER all cutting is done.
  const overWidthBoards = maxPlanerWidthMm > 0
    ? boards.filter((b) => b.widthMm > maxPlanerWidthMm + EPSILON)
    : [];

  const planeGroups = groupBy(boards, (b) => roundTo(planeThickForBoard(b, partsMap), 1));
  const planeNote = overWidthBoards.length
    ? `Set to the thickest group first, then step the planer down — never back up. ` +
      `For laminated parts the target shown is per layer. Take ≤ 1 mm passes per face. Alternate faces. ` +
      `⚠ ${overWidthBoards.map((b) => b.id).join(", ")} exceed your ${formatInches(maxPlanerWidthIn)}" planer capacity ` +
      `— rip those boards to ≤ ${formatInches(maxPlanerWidthIn)}" strips before planing (see individual board guides).`
    : "Set to the thickest group first, then step the planer down — never back up. For laminated parts the target shown is per layer. Take ≤ 1 mm passes per face. Alternate faces.";

  phases.push({
    tool: "Planer", toolClass: "tool-planer",
    heading: "Plane all boards to (per-layer) thickness — thickest first",
    note: planeNote,
    groups: [...planeGroups.entries()].sort((a, b) => b[0] - a[0]).map(([t, bds], i) => ({
      setting: `${formatMm(t, 1)}${i === 0 ? "  ← set here first" : ""}`,
      items: bds.map((b) => {
        const bHasLam     = b.placements.some((p) => (partsMap.get(p.partId)?.layers ?? 1) > 1);
        const overWidth   = maxPlanerWidthMm > 0 && b.widthMm > maxPlanerWidthMm + EPSILON;
        return {
          label:  b.id,
          detail: `${b.thicknessQuarter}/4" stock → ${formatMm(t, 1)}${bHasLam ? " per layer" : ""}` +
                  (overWidth ? ` ⚠ rip to strips first (${formatInches(b.widthIn)} > ${formatInches(maxPlanerWidthIn)}" max)` : ""),
        };
      }),
    })),
  });

  // ── Phase 5: Joint reference edge — all boards ───────────────────────────
  phases.push({
    tool: "Jointer", toolClass: "tool-jointer",
    heading: "Joint reference edge — all boards",
    note: "One straight edge per board. This edge registers against the table saw rip fence for all subsequent rip cuts.",
    groups: [{ setting: null, items: boards.map((b) => ({ label: b.id, detail: "" })) }],
  });

  // ── Phase 6: Trim reference end only — all boards ───────────────────────
  phases.push({
    tool: "Miter saw", toolClass: "tool-mitersaw",
    heading: "Trim reference end only — all boards",
    note: "Square up one end (the reference end) and remove end checks. Leave the far end intact — it will be cleaned up as the final cross-cut on each board. This preserves recovery options if a cross-cut measurement is off.",
    groups: [{
      setting: null,
      items: boards.map((b) => {
        const trim   = b.trimOffsetMm ?? 25.4;
        const usable = b.usableLengthMm ?? (b.lengthMm - 2 * trim);
        return { label: b.id, detail: `trim ≈${formatMm(trim, 0)} from reference end → ≈${formatMm(usable, 0)} usable` };
      }),
    }],
  });

  // ── Phase 7: Cross-cut sections — longest stop position first ────────────
  const allCuts = [];
  for (const b of boards) {
    const sections = buildSections(b);
    if (sections.length <= 1) continue;
    const trim = b.trimOffsetMm ?? 25.4;
    // Every section except the last needs a cross-cut; measure from trimmed reference end
    sections.slice(0, -1).forEach((sec, si) => {
      const pos  = roundTo(sec.endY - trim, 1);
      const len  = roundTo(sec.endY - sec.startY, 1);
      const names = sec.placements.map((p) => shortenPartName(p.partName)).join(", ");
      allCuts.push({ boardId: b.id, pos, len, names, si });
    });
  }

  if (allCuts.length) {
    const cutGroups = groupBy(allCuts, (c) => roundTo(c.pos, 1));
    phases.push({
      tool: "Miter saw", toolClass: "tool-mitersaw",
      heading: "Cross-cut sections — longest measurement first",
      note: "Set stop block to the longest measurement first, then move it inward only. Process all boards at each stop before adjusting.",
      groups: [...cutGroups.entries()].sort((a, b) => b[0] - a[0]).map(([pos, cuts]) => ({
        setting: `Stop at ${formatMm(pos, 0)} from reference end`,
        items: cuts.map((c) => ({
          label: c.boardId,
          detail: `section ${c.si + 1}: ${formatMm(c.len, 0)} long — ${c.names}`,
        })),
      })),
    });
  }

  // ── Phase 8: Rip all blanks — widest fence setting first ─────────────────
  const allRips = [];
  for (const b of boards) {
    const roughT = maxRoughThickForBoard(b, partsMap);
    for (const p of b.placements) {
      const pt = partsMap.get(p.partId);
      allRips.push({
        boardId: b.id,
        name:    p.partName,
        w:       roundTo(p.widthMm,   1),
        l:       roundTo(p.lengthMm,  1),
        t:       roundTo(pt ? pt.roughThicknessMm : roughT, 1),
      });
    }
  }

  const ripGroups = groupBy(allRips, (r) => r.w);
  phases.push({
    tool: "Table saw", toolClass: "tool-tablesaw",
    heading: "Rip all blanks to rough width — widest fence first",
    note: "Set fence to widest dimension first. Move fence inward only — never outward. Process every part at each setting before adjusting the fence.",
    groups: [...ripGroups.entries()].sort((a, b) => b[0] - a[0]).map(([w, rips]) => ({
      setting: `Fence at ${formatMm(w, 0)}`,
      items: rips.map((r) => ({
        label: r.boardId,
        detail: `${r.name}: ${formatMm(r.l, 0)} × ${formatMm(r.w, 0)} × ${formatMm(r.t, 0)}`,
      })),
    })),
  });

  // ── Phase 9: Panel glue-ups (parts wider than any single board) ─────────
  // Collect all panel parts across all boards, along with which board each strip is on.
  const panelMap = new Map(); // partId → { name, count, fullWidthMm, stripBoards: Map<stripIndex→boardId> }
  for (const b of boards) {
    for (const p of b.placements) {
      if (!(p.panelStripCount > 1)) continue;
      if (!panelMap.has(p.partId)) {
        const pt = partsMap.get(p.partId);
        panelMap.set(p.partId, {
          name:        pt ? pt.name : p.partName,
          count:       p.panelStripCount,
          fullWidthMm: p.panelFullWidthMm,
          stripBoards: new Map(),
        });
      }
      panelMap.get(p.partId).stripBoards.set(p.panelStripIndex, b.id);
    }
  }

  if (panelMap.size) {
    // Jointer: edge-joint glue faces — group all panels together (one jointer session)
    phases.push({
      tool: "Jointer", toolClass: "tool-jointer",
      heading: "Edge-joint glue faces — panel glue-ups",
      note: "Joint one face (the glue edge) of each panel strip. For best colour and grain match, keep strips from the same board together.",
      groups: [...panelMap.values()].map((panel) => {
        const stripList = [...panel.stripBoards.entries()]
          .sort((a, b) => a[0] - b[0])
          .map(([idx, bid]) => ({ label: `Strip ${idx}/${panel.count}`, detail: `from ${bid}` }));
        return { setting: `"${panel.name}" — ${panel.count} strips → panel ${formatMm(panel.fullWidthMm, 0)} W`, items: stripList };
      }),
    });

    // Glue-up: one group per panel part, widest panel first (more clamps / longest open time)
    const byWidth = [...panelMap.values()].sort((a, b) => b.fullWidthMm - a.fullWidthMm);
    phases.push({
      tool: "Note", toolClass: "tool-note",
      heading: "Glue up panels — widest first",
      note: "Dry-fit strips before gluing. Apply glue to both mating edges. Clamp overnight. After cure, face-joint and plane each panel to rough thickness before final milling.",
      groups: byWidth.map((panel) => ({
        setting: `"${panel.name}" — ${panel.count} strips`,
        items: [{ label: "Target panel width", detail: `≈${formatMm(panel.fullWidthMm, 0)} rough → ${formatMm(panel.fullWidthMm - PANEL_JOINT_MM * panel.count, 0)} after edge-jointing glued faces` }],
      })),
    });
  }

  // ── Final phases: Thickness lamination (if any) ──────────────────────────
  // These come AFTER all blanks are cut so every layer is ready before glue-up begins.
  const lamParts = new Map();
  for (const b of boards) {
    for (const p of b.placements) {
      const pt = partsMap.get(p.partId);
      if (pt && pt.layers > 1 && !lamParts.has(pt.id)) {
        lamParts.set(pt.id, {
          name:         pt.name,
          layers:       pt.layers,
          netThickMm:   pt.netThicknessMm,
          roughThickMm: pt.roughThicknessMm,
          perLayerMm:   roundTo(pt.roughThicknessMm / pt.layers, 1),
        });
      }
    }
  }
  if (lamParts.size) {
    const lamList = [...lamParts.values()];

    // Step A — collect and confirm all layer blanks are cut
    phases.push({
      tool: "Note", toolClass: "tool-note",
      heading: "Collect all layer blanks before gluing",
      note: "Every layer blank for every laminated part must be cut before any glue-up begins. Gather matching layers and label them.",
      groups: [{
        setting: null,
        items: lamList.map((l) => ({
          label: l.name,
          detail: `${l.layers} layers × ≈${formatMm(l.perLayerMm, 1)} each → ≈${formatMm(l.roughThickMm, 1)} rough after glue-up`,
        })),
      }],
    });

    // Step B — edge-joint the mating faces
    phases.push({
      tool: "Jointer", toolClass: "tool-jointer",
      heading: "Edge-joint mating faces of lamination layer blanks",
      note: "Joint only the faces that will be glued (not all faces). One pass per glue face is enough — just remove mill marks and get the surface flat.",
      groups: [{
        setting: null,
        items: lamList.map((l) => ({
          label: l.name,
          detail: `joint ${l.layers - 1} glue face(s) per set`,
        })),
      }],
    });

    // Step C — glue-up and clamp (thickest first = most clamps needed)
    phases.push({
      tool: "Note", toolClass: "tool-note",
      heading: "Glue up and clamp — thickest assembly first",
      note: "Dry-fit before gluing. Apply glue to both mating faces. Align grain. Clamp with cauls to keep the assembly flat. Let cure fully before unclamping (overnight minimum).",
      groups: [{
        setting: null,
        items: [...lamList].sort((a, b) => b.roughThickMm - a.roughThickMm).map((l) => ({
          label: l.name,
          detail: `${l.layers} layers clamped → ≈${formatMm(l.roughThickMm, 1)} rough`,
        })),
      }],
    });

    // Step D — flatten after cure
    phases.push({
      tool: "Jointer", toolClass: "tool-jointer",
      heading: "Flatten laminated blanks after cure",
      note: "Scrape squeeze-out first. Take one light face-joint pass on both faces to flatten the glue joint. The blank is now at rough thickness and ready for final milling.",
      groups: [{
        setting: null,
        items: lamList.map((l) => ({
          label: l.name,
          detail: `flatten both faces → ${formatMm(l.roughThickMm, 1)} rough → final mill to ${formatMm(l.netThickMm, 1)} net`,
        })),
      }],
    });
  }

  return phases;
}

/**
 * Render the consolidated schedule phases into a DOM card element.
 */
function renderConsolidatedSchedule(phases) {
  const card = document.createElement("div");
  card.className = "workshop-board-card workshop-consolidated-card";

  const heading = document.createElement("h3");
  heading.textContent = "Consolidated Mill Schedule";
  card.append(heading);

  const intro = document.createElement("p");
  intro.className = "muted";
  intro.style.marginTop = "0";
  intro.textContent =
    "Process all boards together through each phase to minimise tool adjustments. " +
    "Complete each phase for all boards before moving to the next.";
  card.append(intro);

  phases.forEach((phase, idx) => {
    const phaseDiv = document.createElement("div");
    phaseDiv.className = "consolidated-phase";

    // Phase header row
    const head = document.createElement("div");
    head.className = "consolidated-phase-head";

    const num = document.createElement("span");
    num.className = "consolidated-phase-num";
    num.textContent = `Phase ${idx + 1}`;

    const badge = document.createElement("span");
    badge.className = `workshop-tool ${phase.toolClass}`;
    badge.textContent = phase.tool;

    const title = document.createElement("span");
    title.className = "consolidated-phase-title";
    title.textContent = phase.heading;

    head.append(num, badge, title);
    phaseDiv.append(head);

    if (phase.note) {
      const noteEl = document.createElement("p");
      noteEl.className = "muted consolidated-phase-note";
      noteEl.textContent = phase.note;
      phaseDiv.append(noteEl);
    }

    for (const group of phase.groups) {
      const groupDiv = document.createElement("div");
      groupDiv.className = "consolidated-group";

      if (group.setting) {
        const settingEl = document.createElement("div");
        settingEl.className = "consolidated-group-setting";
        settingEl.textContent = `⚙ ${group.setting}`;
        groupDiv.append(settingEl);
      }

      const ul = document.createElement("ul");
      ul.className = "consolidated-items";
      for (const item of group.items) {
        const li = document.createElement("li");
        const lbl = document.createElement("strong");
        lbl.textContent = item.label;
        li.append(lbl);
        if (item.detail) li.append(` — ${item.detail}`);
        ul.append(li);
      }
      groupDiv.append(ul);
      phaseDiv.append(groupDiv);
    }

    card.append(phaseDiv);
  });

  return card;
}

function renderLayouts(target, boards) {
  target.innerHTML = "";
  if (!boards.length) {
    target.innerHTML = '<p class="muted">No board layouts available.</p>';
    return;
  }

  // ── Scale control — rendered as a sibling BEFORE the board grid ──────────
  // Remove any existing scale row for this target so re-renders stay clean
  const existingScaleRow = target.previousElementSibling;
  if (existingScaleRow && existingScaleRow.dataset.scaleFor === target.id) {
    existingScaleRow.remove();
  }

  const scaleRow = document.createElement("div");
  scaleRow.dataset.scaleFor = target.id;
  scaleRow.style.cssText =
    "display:flex;align-items:center;gap:10px;margin-bottom:10px;font-size:0.88rem;";

  const scaleLabel = document.createElement("span");
  scaleLabel.textContent = "Diagram scale:";

  const slider = document.createElement("input");
  slider.type  = "range";
  slider.min   = "0.20";
  slider.max   = "2.00";
  slider.step  = "0.05";
  slider.value = String(state.layoutScale);
  slider.style.cssText = "flex:1;min-width:80px;max-width:220px;cursor:pointer;";

  const pct = document.createElement("span");
  pct.style.minWidth = "3.4em";
  pct.textContent    = `${Math.round(state.layoutScale * 100)}%`;

  slider.addEventListener("input", () => {
    state.layoutScale = Number(slider.value);
    renderLayouts(target, boards); // re-render with new scale
  });

  scaleRow.append(scaleLabel, slider, pct);
  target.parentElement.insertBefore(scaleRow, target);

  // ── Board cards ──────────────────────────────────────────────────────────
  const colors = [
    "#bc6c25", "#dda15e", "#606c38", "#283618",
    "#7f5539", "#9c6644", "#386641", "#1d3557",
    "#6d597a", "#2a9d8f",
  ];

  // Global scale: derive from the widest board so all cards are proportional.
  // baseScale targets ~180px display height for the widest board at scale 1.0.
  const maxWidthMm = Math.max(...boards.map((b) => b.widthMm), 1);
  const baseScale  = 180 / maxWidthMm;
  const drawScale  = baseScale * state.layoutScale;

  boards.forEach((board, boardIndex) => {
    const card = document.createElement("article");
    card.className = "board-card";

    const title = document.createElement("h4");
    title.textContent = `${board.id} · ${board.thicknessQuarter}/4 × ${formatInches(board.widthIn)} × ${formatFeet(board.lengthFt, 1)}`;
    card.append(title);

    const subtitle = document.createElement("p");
    subtitle.className = "muted";
    subtitle.textContent = `Metric: ${formatMm(board.widthMm, 1)} × ${formatMm(board.lengthMm, 1)} (${board.source})`;
    card.append(subtitle);
    const svgHeight = Math.max(60, board.widthMm * drawScale);

    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("class", "board-svg");
    svg.setAttribute("viewBox", `0 0 ${board.lengthMm} ${board.widthMm}`);
    svg.setAttribute("preserveAspectRatio", "none");
    svg.style.height = `${svgHeight}px`;

    // Board background
    const boardRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    boardRect.setAttribute("x", "0"); boardRect.setAttribute("y", "0");
    boardRect.setAttribute("width",  String(board.lengthMm));
    boardRect.setAttribute("height", String(board.widthMm));
    boardRect.setAttribute("fill", "#f4e6ce");
    boardRect.setAttribute("stroke", "#a48a6a");
    boardRect.setAttribute("stroke-width", String(Math.max(0.8, board.widthMm * 0.008)));
    svg.append(boardRect);

    // End-trim zones at left and right ends of the board
    if (board.trimTotalMm > EPSILON) {
      for (const [x, w] of [
        [0, board.trimOffsetMm],
        [board.lengthMm - board.trimOffsetMm, board.trimOffsetMm],
      ]) {
        const trim = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        trim.setAttribute("x", String(x)); trim.setAttribute("y", "0");
        trim.setAttribute("width", String(w)); trim.setAttribute("height", String(board.widthMm));
        trim.setAttribute("fill", "#d7c4a8"); trim.setAttribute("fill-opacity", "0.55");
        svg.append(trim);
      }
    }

    board.placements.forEach((placement, placementIndex) => {
      // Coordinate swap: board x-axis (width) → SVG y-axis; board y-axis (length) → SVG x-axis
      const svgX = placement.y;
      const svgY = placement.x;
      const svgW = placement.lengthMm;
      const svgH = placement.widthMm;

      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.setAttribute("x",      String(svgX));
      rect.setAttribute("y",      String(svgY));
      rect.setAttribute("width",  String(svgW));
      rect.setAttribute("height", String(svgH));
      rect.setAttribute("fill",         colors[(placementIndex + boardIndex) % colors.length]);
      rect.setAttribute("fill-opacity", "0.86");
      rect.setAttribute("stroke",       "#ffffff");
      rect.setAttribute("stroke-width", String(Math.max(0.6, board.widthMm * 0.006)));
      svg.append(rect);

      // Label: left-aligned, reads left-to-right, no rotation
      const labelPad = svgW * 0.03;
      const fontSize = Math.min(
        Math.max(4, svgH * 0.45),        // fit within part height
        Math.max(4, svgW * 0.06),        // don't dominate part width
        Math.max(4, board.widthMm * 0.10) // limit by board scale
      );

      const label = document.createElementNS("http://www.w3.org/2000/svg", "text");
      label.setAttribute("x",                String(svgX + labelPad));
      label.setAttribute("y",                String(svgY + svgH / 2));
      label.setAttribute("text-anchor",      "start");
      label.setAttribute("dominant-baseline","central");
      label.setAttribute("font-size",        String(fontSize));
      label.setAttribute("fill",             "#fff");
      label.textContent = shortenPartName(placement.partName);
      svg.append(label);
    });

    card.append(svg);

    const boardYield =
      board.placements.reduce((acc, p) => acc + p.widthMm * p.lengthMm, 0) /
      (board.widthMm * board.lengthMm);
    const caption = document.createElement("p");
    caption.className = "muted";
    caption.textContent = `${board.placements.length} blank(s) · ${formatNumber(boardYield * 100, 1)}% board yield`;
    card.append(caption);

    target.append(card);
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// UI helpers
// ─────────────────────────────────────────────────────────────────────────────
function boardKey(board) {
  return `${board.thicknessQuarter}/4 × ${formatInches(board.widthIn)} × ${formatFeet(board.lengthFt, 1)}`;
}

function boardFeetForBoard(board) {
  const thicknessIn = board.thicknessQuarter / 4;
  return (thicknessIn * board.widthIn * (board.lengthFt * 12)) / 144;
}

function summaryBox(text, kind = "default") {
  const box = document.createElement("div");
  box.className = kind === "default" ? "summary-box" : `summary-box ${kind}`;
  box.textContent = text;
  return box;
}

function setStatus(message, type = "") {
  dom.status.className  = `status ${type}`.trim();
  dom.status.textContent = message;
}

// ─────────────────────────────────────────────────────────────────────────────
// Numeric helpers
// ─────────────────────────────────────────────────────────────────────────────
function getPositiveNumber(value, fallback) {
  const num = Number(value);
  return (Number.isFinite(num) && num > 0) ? num : fallback;
}

function getNonNegativeNumber(value, fallback) {
  const num = Number(value);
  return (Number.isFinite(num) && num >= 0) ? num : fallback;
}

function parseNumberList(text, fallback = []) {
  const values = text.split(",").map((c) => Number(c.trim())).filter((n) => Number.isFinite(n) && n > 0);
  return values.length ? [...new Set(values)].sort((a, b) => a - b) : [...fallback];
}

function parseQuarterList(text, fallback = []) {
  const values = text
    .split(",")
    .map((c) => Math.max(1, Math.round(Number(c.trim()))))
    .filter((n) => Number.isInteger(n) && n > 0);
  return values.length ? [...new Set(values)].sort((a, b) => a - b) : [...fallback];
}

// ─────────────────────────────────────────────────────────────────────────────
// Unit / math helpers
// ─────────────────────────────────────────────────────────────────────────────
function unitToMmFactor(unit) {
  switch (unit) {
    case "mm": return 1;
    case "cm": return 10;
    case "m":  return 1000;
    case "in": return INCH_TO_MM;
    case "ft": return FOOT_TO_MM;
    default:   return 1;
  }
}

function quarterToMm(quarter)   { return (quarter / 4) * INCH_TO_MM; }
function mmToFeet(mm)            { return mm / FOOT_TO_MM; }

function groupBy(items, picker) {
  const map = new Map();
  for (const item of items) {
    const key = picker(item);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(item);
  }
  return map;
}

function sum(values) { return values.reduce((acc, v) => acc + v, 0); }

function nearlyEqual(a, b) { return Math.abs(a - b) <= EPSILON; }

function roundTo(value, digits = 2) {
  const p = 10 ** digits;
  return Math.round(value * p) / p;
}

// ─────────────────────────────────────────────────────────────────────────────
// Formatting helpers
// ─────────────────────────────────────────────────────────────────────────────
function formatNumber(value, digits = 2) { return Number(value).toFixed(digits); }
function formatMm(value, digits = 1)     { return `${formatNumber(value, digits)} mm`; }
function formatInches(value, digits = 2) { return `${formatNumber(value, digits)}"`; }
function formatFeet(value, digits = 1)   { return `${formatNumber(value, digits)}'`; }
function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value);
}

function shortenPartName(name) {
  return name.length > 20 ? `${name.slice(0, 17)}…` : name;
}

function slugify(text) {
  return text.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "part";
}
