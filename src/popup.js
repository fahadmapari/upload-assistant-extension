// ── State ──────────────────────────────────────────────
let config = {};
let tours = [];
let selectedTour = null;
let filteredTours = [];

const $ = (id) => document.getElementById(id);

// Column letters are configured explicitly in the setup panel — no auto-detection.

// ── Dummy data (used for fields not yet coming from sheet) ─
const DUMMY = {
  tourType: "Private",
  activityType: "Walking",
  subType: "Historical",
  description:
    "A fascinating guided tour of the ancient Colosseum and Roman Forum.",
  willSee: "Colosseum\nRoman Forum\nPalatine Hill",
  willLearn: "Roman history\nAncient architecture",
  mandatoryInfo: "Comfortable walking shoes required",
  recommendedInfo: "Water bottle,Sunscreen,Camera",
  included: "Licensed guide,Skip-the-line tickets",
  notIncluded: "Meals,Hotel transfers",
  activityFor: "Everyone",
  voucherType: "Digital",
  noOfPax: "15",
  guideLanguageInstant: "English",
  guideLanguageRequest: "French",
  longitude: "12.4922",
  latitude: "41.8902",
  meetingPoint: "Colosseum main entrance",
  pickupInstructions: "Look for guide with blue flag",
  endPoint: "Roman Forum exit",
  tags: "History",
  priceModel: "Fixed Rate",
  currency: "EUR",
  extraHour: "30",
  extraHourB2C: "25",
  extraHourRequest: "28",
  extraHourRequestB2C: "22",
  holidaySupplement: "15",
  weekendSupplement: "10",
  startTime: "09:00",
  endTime: "13:00",
  isB2CEnabled: "true",
  isB2BEnabled: "true",
};

// ── Init ───────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", async () => {
  config = await loadConfig();

  if (config.sheetId && config.apiKey) {
    $("sheetId").value = config.sheetId;
    $("sheetTab").value = config.sheetTab || "Sheet1";
    $("apiKey").value = config.apiKey;
    // Restore column inputs
    const cols = [
      ["colTitle", "A"],
      ["colDocUrl", "B"],
      ["colCountry", ""],
      ["colCity", ""],
      ["colDuration", ""],
      ["colServiceType", ""],
      ["colRate", ""],
      ["colRateRequest", ""],
      ["colRateB2C", ""],
      ["colRateRequestB2C", ""],
      ["colCancellation", ""],
      ["colCancellationRequest", ""],
      ["colRelease", ""],
      ["colReleaseRequest", ""],
    ];
    cols.forEach(([id, fallback]) => {
      const el = $(id);
      if (el) el.value = config[id] || fallback;
    });
    showPanel("panelList");
    if (!config.sheetName) fetchAndStoreSheetTitle().then(updateConfigBar);
    updateConfigBar();
    loadTours(false);
  }

  $("settingsBtn").addEventListener("click", () => {
    if (tours.length) $("setupCloseBtn").style.display = "flex";
    showPanel("panelSetup");
  });
  $("setupCloseBtn").addEventListener("click", () => showPanel("panelList"));
  $("saveConfigBtn").addEventListener("click", saveAndLoad);

  // Auto-save config on any input change (no fetch)
  [
    "sheetId",
    "sheetTab",
    "apiKey",
    "colTitle",
    "colDocUrl",
    "colCountry",
    "colCity",
    "colDuration",
    "colServiceType",
    "colRate",
    "colRateRequest",
    "colRateB2C",
    "colRateRequestB2C",
    "colCancellation",
    "colCancellationRequest",
    "colRelease",
    "colReleaseRequest",
  ].forEach((id) => {
    const el = $(id);
    if (el) el.addEventListener("input", autoSaveConfig);
  });
  $("searchInput").addEventListener("input", (e) =>
    filterTours(e.target.value),
  );
  $("refreshBtn").addEventListener("click", () => loadTours(true));
  $("selectFillBtn").addEventListener("click", () => {
    if (selectedTour) goToFillPanel(selectedTour);
  });
  $("backBtn").addEventListener("click", () => showPanel("panelList"));
  $("openSectionBtn").addEventListener("click", openAllSections);
  $("startFillBtn").addEventListener("click", startFill);

  $("tourList").addEventListener("click", (e) => {
    const card = e.target.closest(".tour-card");
    if (card) selectTour(parseInt(card.dataset.row));
  });
});

// ── Config ─────────────────────────────────────────────
function loadConfig() {
  return new Promise((resolve) => {
    chrome.storage.local.get(["tourExtConfig"], (r) =>
      resolve(r.tourExtConfig || {}),
    );
  });
}

function saveConfig(cfg) {
  return new Promise((resolve) => {
    chrome.storage.local.set({ tourExtConfig: cfg }, resolve);
  });
}

function buildConfig() {
  const col = (id) => $(id).value.trim() || "";
  return {
    sheetId: $("sheetId").value.trim(),
    sheetTab: $("sheetTab").value.trim() || "Sheet1",
    apiKey: $("apiKey").value.trim(),
    colTitle: col("colTitle") || "A",
    colDocUrl: col("colDocUrl") || "B",
    colCountry: col("colCountry"),
    colCity: col("colCity"),
    colDuration: col("colDuration"),
    colServiceType: col("colServiceType"),
    colRate: col("colRate"),
    colRateRequest: col("colRateRequest"),
    colRateB2C: col("colRateB2C"),
    colRateRequestB2C: col("colRateRequestB2C"),
    colCancellation: col("colCancellation"),
    colCancellationRequest: col("colCancellationRequest"),
    colRelease: col("colRelease"),
    colReleaseRequest: col("colReleaseRequest"),
  };
}

function debounce(fn, ms) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

const autoSaveConfig = debounce(async () => {
  const cfg = buildConfig();
  config = cfg;
  await saveConfig(cfg);
}, 600);

async function saveAndLoad() {
  const cfg = buildConfig();

  if (!cfg.sheetId || !cfg.apiKey) {
    showToast("Sheet ID and API Key are required", "error");
    return;
  }

  config = cfg;
  await saveConfig(cfg);
  await fetchAndStoreSheetTitle();
  updateConfigBar();
  showPanel("panelList");
  loadTours(true);
}

async function fetchAndStoreSheetTitle() {
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${config.sheetId}?fields=properties.title&key=${config.apiKey}`;
    const res = await fetch(url);
    if (!res.ok) return;
    const data = await res.json();
    const title = data.properties?.title;
    if (title) {
      config.sheetName = title;
      await saveConfig(config);
    }
  } catch {
    /* silently ignore — title is cosmetic */
  }
}

function updateConfigBar() {
  $("configBar").style.display = "flex";
  const name = config.sheetName || config.sheetId.slice(0, 20) + "…";
  const tab = config.sheetTab || "Sheet1";
  $("configSheetInfo").textContent = `${name}  ·  ${tab}`;
  $("configConnStatus").textContent = "● connected";
  $("configConnStatus").className = "connected";
}

// ── Tour cache ─────────────────────────────────────────
function saveCachedTours(data) {
  return new Promise((resolve) =>
    chrome.storage.local.set({ tourExtCache: data }, resolve),
  );
}
function loadCachedTours() {
  return new Promise((resolve) =>
    chrome.storage.local.get(["tourExtCache"], (r) =>
      resolve(r.tourExtCache || null),
    ),
  );
}

// ── Google Sheets fetch ────────────────────────────────
// Fetches all columns, maps by header name using COLUMN_MAP
async function loadTours(forceFresh = false) {
  setStatus("busy");
  renderTourList([]);

  if (!forceFresh) {
    const cached = await loadCachedTours();
    if (cached && cached.length) {
      tours = cached;
      filteredTours = [...tours];
      $("tourCountLabel").textContent =
        `${tours.length} tour${tours.length !== 1 ? "s" : ""} found (cached)`;
      renderTourList(filteredTours);
      setStatus("ready");
      return;
    }
  }

  $("tourCountLabel").textContent = "Fetching from Google Sheets…";

  try {
    // Build the range from only the configured columns
    // Always wrap tab name in single quotes to handle spaces and special characters.
    // The entire range must be encoded AFTER the quotes are added.
    const range = `'${config.sheetTab}'!A1:Z500`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${config.sheetId}/values/${encodeURIComponent(range)}?key=${config.apiKey}`;

    console.log("[TourExt] Fetching range:", range);
    console.log(
      "[TourExt] Full URL (key hidden):",
      url.replace(config.apiKey, "***"),
    );
    const res = await fetch(url);
    console.log("[TourExt] Response status:", res.status);
    if (!res.ok) {
      const err = await res.json();
      console.error(
        "[TourExt] API error response:",
        JSON.stringify(err, null, 2),
      );
      throw new Error(err.error?.message || `HTTP ${res.status}`);
    }

    const data = await res.json();
    const rows = data.values || [];
    console.log("[TourExt] Rows returned:", rows.length);

    if (rows.length < 2) {
      $("tourCountLabel").textContent = "0 tours found";
      renderTourList([]);
      setStatus("ready");
      return;
    }

    const headers = rows[0].map((h) => h.toLowerCase().trim());

    tours = rows
      .slice(1)
      .map((row, i) => buildTour(headers, row, i + 2))
      .filter((t) => t.title);

    filteredTours = [...tours];
    await saveCachedTours(tours);
    $("tourCountLabel").textContent =
      `${tours.length} tour${tours.length !== 1 ? "s" : ""} found`;
    renderTourList(filteredTours);
    setStatus("ready");
  } catch (e) {
    console.error("[TourExt] Sheet fetch failed:", e.message);
    console.error(
      "[TourExt] Config used — sheetId:",
      config.sheetId,
      "| sheetTab:",
      config.sheetTab,
      "| apiKey set:",
      !!config.apiKey,
    );
    $("tourCountLabel").textContent = "Failed to load";
    showToast(`Error: ${e.message}`, "error");
    setStatus("error");
  }
}

function buildTour(headers, row, rowNum) {
  // Helper: get cell value by configured column letter (e.g. "A" → index 0)
  const col = (cfgKey) => {
    const letter = config[cfgKey];
    if (!letter) return "";
    const idx = colLetterToIndex(letter);
    return (row[idx] || "").trim();
  };

  return {
    rowNum,
    title: col("colTitle"),
    docUrl: col("colDocUrl"),
    country: col("colCountry"),
    city: col("colCity"),
    duration: col("colDuration"),
    serviceType: col("colServiceType"),
    rate: col("colRate"),
    rateRequest: col("colRateRequest"),
    rateB2C: col("colRateB2C"),
    rateRequestB2C: col("colRateRequestB2C"),
    cancellation: col("colCancellation"),
    cancellationRequest: col("colCancellationRequest"),
    release: col("colRelease"),
    releaseRequest: col("colReleaseRequest"),
  };
}

// Convert column letter(s) to 0-based index: A→0, B→1, Z→25, AA→26, etc.
function colLetterToIndex(letters) {
  letters = letters.toUpperCase().trim();
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index - 1;
}

// ── Tour List render ───────────────────────────────────
function renderTourList(list) {
  const container = $("tourList");

  if (!list.length) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📋</div>
        ${tours.length ? "No tours match your search" : "No tours found in sheet"}
      </div>`;
    return;
  }

  container.innerHTML = list
    .map(
      (t) => `
    <div class="tour-card" data-row="${t.rowNum}">
      <div class="tour-card-title">${t.title}</div>
      <div class="tour-card-meta">
        <span class="tag">Row ${t.rowNum}</span>
        ${t.country ? `<span class="tag green">${t.country}</span>` : ""}
        ${t.city ? `<span class="tag green">${t.city}</span>` : ""}
        ${t.docUrl ? `<span class="tag blue">Doc linked</span>` : `<span class="tag">No doc</span>`}
      </div>
    </div>
  `,
    )
    .join("");
}

function selectTour(rowNum) {
  selectedTour = tours.find((t) => t.rowNum === rowNum);
  document.querySelectorAll(".tour-card").forEach((c) => {
    c.classList.toggle("selected", parseInt(c.dataset.row) === rowNum);
  });
  $("selectFillBtn").disabled = false;
  $("selectFillBtn").textContent =
    `Fill "${(selectedTour.title || "").slice(0, 26)}" →`;
}

function filterTours(query) {
  const q = query.toLowerCase();
  filteredTours = q
    ? tours.filter(
        (t) =>
          (t.title || "").toLowerCase().includes(q) ||
          (t.country || "").toLowerCase().includes(q) ||
          (t.city || "").toLowerCase().includes(q),
      )
    : [...tours];
  $("tourCountLabel").textContent =
    `${filteredTours.length} of ${tours.length} tours`;
  renderTourList(filteredTours);
}

// ── Google Docs fetch (via chrome.identity OAuth) ──────
// Requests an OAuth token silently, then calls the Docs API.
// Raw document JSON is returned — parsing happens elsewhere.
async function fetchGoogleDoc(docUrl) {
  const match = docUrl && docUrl.match(/\/d\/([\w-]+)/);
  if (!match) return null;
  const docId = match[1];

  // Get OAuth token via chrome.identity (uses the user's logged-in Google account)
  const token = await new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, (t) => {
      if (chrome.runtime.lastError)
        reject(new Error(chrome.runtime.lastError.message));
      else resolve(t);
    });
  });

  const res = await fetch(`https://docs.googleapis.com/v1/documents/${docId}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error?.message || `Docs API ${res.status}`);
  }

  // Return raw document object — caller is responsible for parsing
  return await res.json();
}

// ── Fill Panel ──────────────────────────────────────────
const FILL_FIELDS = [
  { key: "title", label: "Tour Title", source: "sheet" },
  { key: "serviceType", label: "Product Type", source: "sheet" },
  { key: "country", label: "Country", source: "sheet" },
  { key: "city", label: "City", source: "sheet" },
  { key: "duration", label: "Duration", source: "sheet" },
  { key: "rate", label: "B2B Price (Instant)", source: "sheet" },
  { key: "rateRequest", label: "B2B Price (On Request)", source: "sheet" },
  { key: "rateB2C", label: "B2C Price (Instant)", source: "sheet" },
  { key: "rateRequestB2C", label: "B2C Price (On Request)", source: "sheet" },
  { key: "cancellation", label: "Cancellation (Instant)", source: "sheet" },
  {
    key: "cancellationRequest",
    label: "Cancellation (On Request)",
    source: "sheet",
  },
  { key: "release", label: "Cut Off (Instant)", source: "sheet" },
  { key: "releaseRequest", label: "Cut Off (On Request)", source: "sheet" },
  // Below: still using dummy data until doc parsing is implemented
  { key: "description", label: "Description (Quill)", source: "dummy" },
  { key: "tourType", label: "Tour Type", source: "dummy" },
  { key: "activityType", label: "Activity Type", source: "dummy" },
  { key: "subType", label: "Sub Type", source: "dummy" },
  { key: "willSee", label: "You Will See", source: "dummy" },
  { key: "willLearn", label: "You Will Learn", source: "dummy" },
  { key: "mandatoryInfo", label: "Mandatory Information", source: "dummy" },
  { key: "recommendedInfo", label: "Recommended Information", source: "dummy" },
  { key: "included", label: "Included", source: "dummy" },
  { key: "notIncluded", label: "Not Included", source: "dummy" },
  { key: "activityFor", label: "Activity For", source: "dummy" },
  { key: "voucherType", label: "Voucher Type", source: "dummy" },
  { key: "noOfPax", label: "No of Pax", source: "dummy" },
  {
    key: "guideLanguageInstant",
    label: "Guide Language (Instant)",
    source: "dummy",
  },
  {
    key: "guideLanguageRequest",
    label: "Guide Language (Request)",
    source: "dummy",
  },
  { key: "longitude", label: "Longitude", source: "dummy" },
  { key: "latitude", label: "Latitude", source: "dummy" },
  { key: "meetingPoint", label: "Meeting Point", source: "dummy" },
  { key: "pickupInstructions", label: "Pickup Instructions", source: "dummy" },
  { key: "endPoint", label: "End Point", source: "dummy" },
  { key: "tags", label: "Tags", source: "dummy" },
  { key: "priceModel", label: "Price Model", source: "dummy" },
  { key: "currency", label: "Currency", source: "dummy" },
  { key: "extraHour", label: "Extra Hour (Instant)", source: "dummy" },
  { key: "extraHourB2C", label: "Extra Hour B2C", source: "dummy" },
  { key: "extraHourRequest", label: "Extra Hour (Request)", source: "dummy" },
  {
    key: "extraHourRequestB2C",
    label: "Extra Hour B2C (Request)",
    source: "dummy",
  },
  { key: "holidaySupplement", label: "Holiday Supplement %", source: "dummy" },
  { key: "weekendSupplement", label: "Weekend Supplement %", source: "dummy" },
  { key: "startTime", label: "Start Time", source: "dummy" },
  { key: "endTime", label: "End Time", source: "dummy" },
  { key: "isB2CEnabled", label: "B2C Enabled", source: "dummy" },
  { key: "isB2BEnabled", label: "B2B Enabled", source: "dummy" },
];

function goToFillPanel(tour) {
  $("previewName").textContent = tour.title;

  // Show sheet-sourced values in the preview
  const previewRows = [
    ["Country", tour.country],
    ["City", tour.city],
    ["Duration", tour.duration],
    ["Product Type", tour.serviceType],
    ["B2B Instant", tour.rate ? `€${tour.rate}` : null],
    ["B2B Request", tour.rateRequest ? `€${tour.rateRequest}` : null],
    ["B2C Instant", tour.rateB2C ? `€${tour.rateB2C}` : null],
    ["B2C Request", tour.rateRequestB2C ? `€${tour.rateRequestB2C}` : null],
  ].filter(([, v]) => v);

  $("previewFields").innerHTML = `
    ${previewRows
      .map(
        ([k, v]) => `
      <div class="preview-row">
        <span class="preview-key">${k}</span>
        <span class="preview-val">${v}</span>
      </div>`,
      )
      .join("")}
    <div class="preview-row">
      <span class="preview-key">Doc</span>
      <span class="preview-val">${
        tour.docUrl
          ? `<a href="${tour.docUrl}" target="_blank" style="color:var(--accent)">View ↗</a>`
          : '<span style="color:var(--muted)">Not linked</span>'
      }</span>
    </div>
    <div class="preview-row" style="margin-top:4px">
      <span class="preview-key" style="color:var(--warning)">⚠ Note</span>
      <span class="preview-val" style="color:var(--warning);font-size:10px">Non-sheet fields use dummy data</span>
    </div>
  `;

  // Checklist — label source per field
  $("fieldsChecklist").innerHTML = FILL_FIELDS.map(
    (f) => `
    <div class="check-item" id="check-${f.key}">
      <div class="check-dot"></div>
      <span>${f.label}</span>
      <span style="margin-left:auto;font-size:9px;font-family:var(--mono);color:${f.source === "sheet" ? "var(--success)" : "var(--muted)"}">
        ${f.source === "sheet" ? "sheet" : "dummy"}
      </span>
    </div>
  `,
  ).join("");

  showPanel("panelFill");
}

// ── Open All Sections ───────────────────────────────────
async function openAllSections() {
  try {
    const [tab] = await chrome.tabs.query({
      active: true,
      currentWindow: true,
    });
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => {
        document
          .querySelectorAll("mat-expansion-panel-header")
          .forEach((h, i) => {
            if (h.getAttribute("aria-expanded") !== "true") {
              setTimeout(() => h.click(), i * 600);
            }
          });
      },
    });
    showToast("Opening sections — wait 2s then autofill", "info");
  } catch (e) {
    showToast("Could not open sections: " + e.message, "error");
  }
}

// ── Autofill ────────────────────────────────────────────
async function startFill() {
  if (!selectedTour) return;

  $("startFillBtn").disabled = true;
  $("startFillBtn").innerHTML = '<div class="spinner"></div> Filling…';
  setStatus("busy");

  // Merge: dummy base → overridden by real sheet values where available
  const fillData = {
    ...DUMMY,
    title: selectedTour.title,
    serviceType: selectedTour.serviceType || DUMMY.serviceType,
    country: selectedTour.country || DUMMY.country,
    city: selectedTour.city || DUMMY.city,
    duration: selectedTour.duration || DUMMY.duration,
    rate: selectedTour.rate || DUMMY.rate,
    rateRequest: selectedTour.rateRequest || DUMMY.rateRequest,
    rateB2C: selectedTour.rateB2C || DUMMY.rateB2C,
    rateRequestB2C: selectedTour.rateRequestB2C || DUMMY.rateRequestB2C,
    cancellation: selectedTour.cancellation || DUMMY.cancellation,
    cancellationRequest: selectedTour.cancellationRequest || null,
    release: selectedTour.release || DUMMY.release,
    releaseRequest: selectedTour.releaseRequest || null,
  };

  try {
    const [tab] = await chrome.tabs.query({
      active: true,
      currentWindow: true,
    });
    const result = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: injectTourData,
      args: [fillData],
    });

    const outcome = result?.[0]?.result;
    if (outcome?.success) {
      outcome.filled?.forEach((key) => {
        const el = $(`check-${key}`);
        if (el) el.className = "check-item done";
      });
      outcome.failed?.forEach((key) => {
        const el = $(`check-${key}`);
        if (el) el.className = "check-item error";
      });
      showToast(`✓ Filled ${outcome.filled?.length || 0} fields`, "success");
      setStatus("ready");
    } else {
      throw new Error(outcome?.error || "Unknown error");
    }
  } catch (e) {
    showToast("Fill failed: " + e.message, "error");
    setStatus("error");
  }

  $("startFillBtn").disabled = false;
  $("startFillBtn").innerHTML = "▶ Start Autofill";
}

// ── Injected into admin page ────────────────────────────
function injectTourData(tour) {
  const filled = [];
  const failed = [];

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function setNativeInput(el, value) {
    if (!el) return false;
    try {
      const proto =
        el.tagName === "TEXTAREA"
          ? window.HTMLTextAreaElement.prototype
          : window.HTMLInputElement.prototype;
      const setter = Object.getOwnPropertyDescriptor(proto, "value").set;
      setter.call(el, value);
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
      el.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
      return true;
    } catch {
      return false;
    }
  }

  function fillByControl(controlName, value) {
    if (!value && value !== 0) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el || el.disabled || el.getAttribute("readonly") === "readonly")
      return;
    const tag = el.tagName.toLowerCase();
    if (tag === "input" || tag === "textarea") {
      if (setNativeInput(el, value)) filled.push(controlName);
      else failed.push(controlName);
    }
  }

  async function fillNgSelect(controlName, value) {
    if (!value) return;
    const container = document.querySelector(
      `[formcontrolname="${controlName}"]`,
    );
    if (!container) {
      failed.push(controlName);
      return;
    }
    try {
      container.click();
      await sleep(250);
      const searchInput = container.querySelector(".ng-input input");
      if (searchInput && !searchInput.readOnly && !searchInput.disabled) {
        setNativeInput(searchInput, value);
        await sleep(350);
      }
      const options = [
        ...document.querySelectorAll(".ng-option:not(.ng-option-disabled)"),
      ];
      const match =
        options.find((o) => o.textContent.trim() === value) ||
        options.find((o) =>
          o.textContent.trim().toLowerCase().includes(value.toLowerCase()),
        );
      if (match) {
        match.click();
        filled.push(controlName);
      } else {
        document.body.click();
        failed.push(controlName);
      }
      await sleep(150);
    } catch {
      failed.push(controlName);
    }
  }

  function fillQuill(value) {
    if (!value) return;
    const editor = document.querySelector("quill-editor .ql-editor");
    if (!editor) {
      failed.push("description");
      return;
    }
    try {
      editor.focus();
      editor.innerHTML = `<p>${value.replace(/\n/g, "</p><p>")}</p>`;
      editor.dispatchEvent(new Event("input", { bubbles: true }));
      editor.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
      filled.push("description");
    } catch {
      failed.push("description");
    }
  }

  function setCheckbox(controlName, value) {
    if (!value) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el) return;
    const shouldCheck = ["true", "yes", "1"].includes(
      String(value).toLowerCase(),
    );
    if (el.checked !== shouldCheck) el.click();
    filled.push(controlName);
  }

  async function runFill() {
    // Text / number inputs
    fillByControl("tourTitle", tour.title);
    fillByControl("descriptionWillSee", tour.willSee);
    fillByControl("descriptionLearn", tour.willLearn);
    fillByControl("mandatoryInformation", tour.mandatoryInfo);
    fillByControl("recommendedInformation", tour.recommendedInfo);
    fillByControl("included", tour.included);
    fillByControl("notIncluded", tour.notIncluded);
    fillByControl("noOfPax", tour.noOfPax);
    fillByControl("longitude", tour.longitude);
    fillByControl("latitude", tour.latitude);
    fillByControl("meetingPoint", tour.meetingPoint);
    fillByControl("pickupInstructions", tour.pickupInstructions);
    fillByControl("endPoint", tour.endPoint);
    fillByControl("duration", tour.duration);
    // Prices
    fillByControl("rate", tour.rate);
    fillByControl("rateB2C", tour.rateB2C);
    fillByControl("rate_request", tour.rateRequest);
    fillByControl("rate_requestB2C", tour.rateRequestB2C);
    fillByControl("extraHourCharges", tour.extraHour);
    fillByControl("extraHourChargesB2C", tour.extraHourB2C);
    fillByControl("extraHourCharges_request", tour.extraHourRequest);
    fillByControl("extraHourCharges_requestB2C", tour.extraHourRequestB2C);
    fillByControl("publicHolidaySurchargePercentage", tour.holidaySupplement);
    fillByControl("weekendSupplementPercentage", tour.weekendSupplement);
    // Schedule
    fillByControl("startTime", tour.startTime);
    fillByControl("endTime", tour.endTime);
    // Cancellation & cut off — two separate fields for instant vs on-request
    fillByControl("cancellation", tour.cancellation);
    fillByControl("cancellation_request", tour.cancellationRequest);
    fillByControl("release", tour.release);
    fillByControl("release_request", tour.releaseRequest);

    // Quill
    fillQuill(tour.description);

    // Checkboxes
    setCheckbox("isB2CEnabled", tour.isB2CEnabled);
    setCheckbox("isB2BEnabled", tour.isB2BEnabled);

    // ng-selects (sequential)
    await fillNgSelect("serviceType", tour.serviceType);
    await fillNgSelect("tourType", tour.tourType);
    await fillNgSelect("activityType", tour.activityType);
    await fillNgSelect("subType", tour.subType);
    await fillNgSelect("activityFor", tour.activityFor);
    await fillNgSelect("voucherType", tour.voucherType);
    await fillNgSelect("countryId", tour.country);
    await fillNgSelect("cityId", tour.city);
    await fillNgSelect("isFixedModel", tour.priceModel);
    await fillNgSelect("currency", tour.currency);
    await fillNgSelect("tourGuideLanguageList", tour.guideLanguageInstant);
    await fillNgSelect(
      "tourGuideLanguageList_request",
      tour.guideLanguageRequest,
    );
    await fillNgSelect("tagsList", tour.tags);

    return { success: true, filled, failed };
  }

  return runFill();
}

// ── UI helpers ─────────────────────────────────────────
function showPanel(id) {
  document
    .querySelectorAll(".panel")
    .forEach((p) => p.classList.remove("active"));
  $(id).classList.add("active");
}

function setStatus(state) {
  $("statusDot").className = `status-dot ${state}`;
}

let toastTimer;
function showToast(msg, type = "success") {
  const toast = $("toast");
  $("toastIcon").textContent =
    { success: "✓", error: "✕", info: "ℹ" }[type] || "✓";
  $("toastMsg").textContent = msg;
  toast.className = `toast ${type} show`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), 3200);
}
