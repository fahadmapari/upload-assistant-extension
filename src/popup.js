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
  activityType: "City Tours",
  subType: "Walking Tours",
  description:
    "A fascinating guided tour of the ancient Colosseum and Roman Forum.",
  willSee: "Colosseum\nRoman Forum\nPalatine Hill",
  willLearn: "Roman history\nAncient architecture",
  mandatoryInfo: "Comfortable walking shoes required",
  recommendedInfo: "Water bottle,Sunscreen,Camera",
  included: "Licensed guide,Skip-the-line tickets",
  notIncluded: "Meals,Hotel transfers",
  activityFor: "All",
  voucherType: "Printed or E-Voucher Accepted",
  noOfPax: "15",
  guideLanguageInstant: "English",
  guideLanguageRequest: [
    "Arabic", "Belarusian", "Bosnian", "Bulgarian", "Cantonese", "Chinese Mandarin",
    "Croatian", "Czech", "Danish", "Dutch", "Estonian", "Finnish", "French", "German",
    "Greek", "Hebrew", "Hindi", "Hungarian", "Icelandic", "Indonesian", "Italian",
    "Japanese", "Korean", "Latvian", "Lithuanian", "Maltese", "Norwegian", "Persian",
    "Polish", "Portuguese", "Romanian", "Russian", "Serbian", "Slovak", "Slovenian",
    "Spanish", "Swedish", "Taiwanese", "Thai", "Turkish", "Ukranian", "Vietnamese",
  ],
  longitude: "12.4922",
  latitude: "41.8902",
  meetingPoint: "Colosseum main entrance",
  pickupInstructions: "",
  endPoint: "Roman Forum exit",
  tags: "Walk",
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

  if (config.sheetId) {
    $("sheetId").value  = config.sheetId;
    $("sheetTab").value = config.sheetTab || "Sheet1";
    if (config.apiKey) $("apiKey").value = config.apiKey; // legacy field
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
  $("startFillBtn").addEventListener("click", startFill);

  $("tourList").addEventListener("click", (e) => {
    const card = e.target.closest(".tour-card");
    if (card) selectTour(parseInt(card.dataset.row));
  });
});

// ── Config ─────────────────────────────────────────────
function loadConfig() {
  return new Promise((resolve) => {
    chrome.storage.local.get(["tourExtConfig"], (r) => {
      if (chrome.runtime.lastError) {
        console.error("[TourExt] loadConfig error:", chrome.runtime.lastError.message);
        resolve({});
      } else {
        resolve(r.tourExtConfig || {});
      }
    });
  });
}

function saveConfig(cfg) {
  return new Promise((resolve) => {
    chrome.storage.local.set({ tourExtConfig: cfg }, () => {
      if (chrome.runtime.lastError) {
        console.error("[TourExt] saveConfig error:", chrome.runtime.lastError.message);
      }
      resolve();
    });
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
  try {
    const cfg = buildConfig();
    config = cfg;
    await saveConfig(cfg);
  } catch (e) {
    console.error("[TourExt] autoSaveConfig error:", e.message);
  }
}, 600);

async function saveAndLoad() {
  const cfg = buildConfig();

  if (!cfg.sheetId) {
    showToast("Sheet ID is required", "error");
    return;
  }

  config = cfg;
  await saveConfig(cfg);
  await fetchAndStoreSheetTitle();
  updateConfigBar();
  showPanel("panelList");
  loadTours(true);
}

// ── OAuth token helper (shared by Sheets + Docs) ───────
function getOAuthToken(interactive = false) {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive }, (token) => {
      if (chrome.runtime.lastError)
        reject(new Error(chrome.runtime.lastError.message));
      else resolve(token);
    });
  });
}

function removeCachedToken(token) {
  return new Promise(resolve => chrome.identity.removeCachedAuthToken({ token }, () => {
    if (chrome.runtime.lastError) {
      console.warn("[TourExt] removeCachedAuthToken error:", chrome.runtime.lastError.message);
    }
    resolve();
  }));
}

// Fetch with automatic token-refresh on 401/403 (handles stale/insufficient-scope tokens)
async function authedFetch(url, interactive = true) {
  let token = await getOAuthToken(interactive);
  let res   = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });

  if (res.status === 401 || res.status === 403) {
    await removeCachedToken(token);
    token = await getOAuthToken(true); // force fresh interactive auth with current scopes
    res   = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  }

  return res;
}

async function fetchAndStoreSheetTitle() {
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${config.sheetId}?fields=properties.title`;
    const res = await authedFetch(url, false);
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
    chrome.storage.local.set({ tourExtCache: data }, () => {
      if (chrome.runtime.lastError) {
        console.error("[TourExt] saveCachedTours error:", chrome.runtime.lastError.message);
      }
      resolve();
    }),
  );
}
function loadCachedTours() {
  return new Promise((resolve) =>
    chrome.storage.local.get(["tourExtCache"], (r) => {
      if (chrome.runtime.lastError) {
        console.error("[TourExt] loadCachedTours error:", chrome.runtime.lastError.message);
        resolve(null);
      } else {
        resolve(r.tourExtCache || null);
      }
    }),
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
    const range = `'${config.sheetTab}'!A:Z`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${config.sheetId}/values/${encodeURIComponent(range)}`;

    console.log("[TourExt] Fetching range:", range);
    const res = await authedFetch(url, true);
    console.log("[TourExt] Response status:", res.status);
    if (!res.ok) {
      let errMsg = `HTTP ${res.status}`;
      try {
        const err = await res.json();
        console.error("[TourExt] API error response:", JSON.stringify(err, null, 2));
        errMsg = err.error?.message || errMsg;
      } catch (jsonErr) {
        console.error("[TourExt] Could not parse error response body:", jsonErr.message);
      }
      throw new Error(errMsg);
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
    console.error("[TourExt] Config — sheetId:", config.sheetId, "| sheetTab:", config.sheetTab);
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

// ── Virtual tour list ──────────────────────────────────
// Row height = card padding(20) + title(18) + meta margin(4) + meta(16) + inter-card gap(6) = 64px
const V_ROW    = 64;
const V_OVERSCAN = 4;
let vData      = [];
let vPrevStart = -1;

function tourCardHTML(t) {
  const sel = selectedTour?.rowNum === t.rowNum ? " selected" : "";
  return `<div class="tour-card${sel}" data-row="${t.rowNum}" style="margin-bottom:6px">
    <div class="tour-card-title">${t.title}</div>
    <div class="tour-card-meta">
      <span class="tag">Row ${t.rowNum}</span>
      ${t.country ? `<span class="tag green">${t.country}</span>` : ""}
      ${t.city    ? `<span class="tag green">${t.city}</span>`    : ""}
      ${t.docUrl  ? `<span class="tag blue">Doc linked</span>`    : `<span class="tag">No doc</span>`}
    </div>
  </div>`;
}

function renderTourList(list) {
  vData = list;
  vPrevStart = -1;
  const container = $("tourList");
  container.onscroll = null;

  if (!list.length) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📋</div>
        ${tours.length ? "No tours match your search" : "No tours found in sheet"}
      </div>`;
    return;
  }

  // flex-shrink:0 is critical — prevents the flex parent from collapsing the scroller
  // height, which would eliminate overflow and disable scrolling entirely.
  const totalH = list.length * V_ROW - 6;
  container.innerHTML =
    `<div id="vScroller" style="position:relative;width:100%;height:${totalH}px;flex-shrink:0">` +
    `<div id="vItems" style="position:absolute;left:0;right:0;top:0"></div>` +
    `</div>`;

  container.onscroll = () => requestAnimationFrame(renderVisible);
  renderVisible();
}

function renderVisible() {
  const container = $("tourList");
  const vItems = $("vItems");
  if (!vData.length || !vItems) return;

  const scrollTop = container.scrollTop;
  const viewH     = container.clientHeight || 280;

  const start = Math.max(0, Math.floor(scrollTop / V_ROW) - V_OVERSCAN);
  const end   = Math.min(vData.length - 1, Math.ceil((scrollTop + viewH) / V_ROW) + V_OVERSCAN - 1);

  if (start === vPrevStart) return;
  vPrevStart = start;

  vItems.style.top = start * V_ROW + "px";
  vItems.innerHTML = vData.slice(start, end + 1).map(tourCardHTML).join("");
}

function selectTour(rowNum) {
  selectedTour = tours.find((t) => t.rowNum === rowNum);
  if (!selectedTour) {
    console.error("[TourExt] selectTour: no tour found for rowNum", rowNum);
    return;
  }
  // Update visible cards only — off-screen cards get correct class on next renderVisible()
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
async function fetchGoogleDoc(docUrl) {
  const match = docUrl && docUrl.match(/\/d\/([\w-]+)/);
  if (!match) {
    console.warn("[TourExt] fetchGoogleDoc: invalid or missing doc URL:", docUrl);
    return null;
  }
  const docId = match[1];

  try {
    const res = await authedFetch(`https://docs.googleapis.com/v1/documents/${docId}`, true);

    if (!res.ok) {
      let errMsg = `Docs API ${res.status}`;
      try {
        const err = await res.json();
        errMsg = err.error?.message || errMsg;
      } catch (jsonErr) {
        console.error("[TourExt] fetchGoogleDoc: could not parse error response:", jsonErr.message);
      }
      throw new Error(errMsg);
    }

    // Return raw document object — caller is responsible for parsing
    return await res.json();
  } catch (e) {
    console.error("[TourExt] fetchGoogleDoc failed for docId", docId, ":", e.message);
    throw e;
  }
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

// ── Autofill ────────────────────────────────────────────
async function startFill() {
  if (!selectedTour) return;

  $("startFillBtn").disabled = true;
  $("startFillBtn").innerHTML = '<div class="spinner"></div> Filling…';
  setStatus("busy");

  // Compute start/end dates at fill time
  const _now = new Date();
  const _end = new Date(_now);
  _end.setFullYear(_end.getFullYear() + 2);
  const _fmt = (d) => `${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}/${d.getFullYear()}`;

  // Merge: dummy base → overridden by real sheet values where available
  const fillData = {
    ...DUMMY,
    title: selectedTour.title,
    serviceType: "Guide",
    subType: selectedTour.serviceType === "Driver-Guide" ? "Driver Guide" : "Walking Tours",
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
    startDate: _fmt(_now),
    endDate: _fmt(_end),
  };

  try {
    const [tab] = await chrome.tabs.query({
      active: true,
      currentWindow: true,
    });
    if (!tab) {
      throw new Error("No active tab found");
    }
    const result = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: injectTourData,
      args: [fillData],
      world: "MAIN",
    });

    const outcome = result?.[0]?.result;
    if (outcome?.success) {
      outcome.filled?.forEach((key) => {
        const el = $(`check-${key}`);
        if (el) el.className = "check-item done";
      });
      outcome.failed?.forEach((key) => {
        const el = $(`check-${key}`);
        if (el) {
          el.className = "check-item error";
          const msg = outcome.errors?.[key];
          if (msg) {
            el.title = msg;
            const lbl = el.querySelector("span:not(.check-dot)");
            if (lbl) lbl.textContent += `  — ${msg.slice(0, 60)}`;
          }
        }
      });
      const failCount = outcome.failed?.length || 0;
      const fillCount = outcome.filled?.length || 0;
      showToast(
        failCount
          ? `✓ ${fillCount} filled  ✕ ${failCount} failed`
          : `✓ Filled ${fillCount} fields`,
        failCount ? "error" : "success",
      );
      setStatus("ready");
    } else {
      throw new Error(outcome?.error || "Unknown error");
    }
  } catch (e) {
    console.error("[TourExt] startFill failed:", e.message);
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
  const errors = {};

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

  async function ensurePanelOpen(el) {
    const panel = el.closest("mat-expansion-panel");
    if (!panel) return;
    const header = panel.querySelector("mat-expansion-panel-header");
    if (!header) return;
    if (header.getAttribute("aria-expanded") !== "true") {
      header.click();
      await sleep(600);
    }
  }

  async function fillNgbDatepicker(controlName, dateStr) {
    if (!dateStr) return;
    const [mm, dd, yyyy] = dateStr.split("/").map(Number);
    const input = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!input) { failed.push(controlName); return; }
    try {
      await ensurePanelOpen(input);
      input.click();
      await sleep(500);

      // Datepicker uses container="body" so it appends to <body>
      const picker = document.querySelector("ngb-datepicker.ngb-dp-body, ngb-datepicker.show");
      if (!picker) {
        errors[controlName] = "Datepicker popup did not open";
        failed.push(controlName);
        return;
      }

      const yearSel = picker.querySelector("select[aria-label='Select year']");
      const monthSel = picker.querySelector("select[aria-label='Select month']");

      // Set year first (affects which months are selectable)
      if (yearSel && yearSel.value !== String(yyyy)) {
        yearSel.value = String(yyyy);
        yearSel.dispatchEvent(new Event("change", { bubbles: true }));
        await sleep(300);
      }
      // Then set month
      if (monthSel && monthSel.value !== String(mm)) {
        monthSel.value = String(mm);
        monthSel.dispatchEvent(new Event("change", { bubbles: true }));
        await sleep(300);
      }

      // Click the matching day — skip "outside" (adjacent-month) cells
      const dayEls = picker.querySelectorAll(".ngb-dp-day:not(.disabled)");
      let clicked = false;
      for (const dayEl of dayEls) {
        const inner = dayEl.querySelector("[ngbdatepickerdayview]");
        if (!inner || inner.classList.contains("outside")) continue;
        if (inner.textContent.trim() === String(dd)) {
          dayEl.click();
          clicked = true;
          break;
        }
      }

      if (clicked) {
        filled.push(controlName);
      } else {
        errors[controlName] = `Day ${dd} not found in picker (month ${mm}/${yyyy})`;
        document.body.click();
        failed.push(controlName);
      }
      await sleep(200);
    } catch (e) {
      errors[controlName] = e.message;
      failed.push(controlName);
    }
  }

  async function fillByControl(controlName, value) {
    if (!value && value !== 0) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el || el.disabled || el.getAttribute("readonly") === "readonly")
      return;
    await ensurePanelOpen(el);
    const tag = el.tagName.toLowerCase();
    if (tag === "input" || tag === "textarea") {
      if (setNativeInput(el, value)) filled.push(controlName);
      else failed.push(controlName);
    }
  }

  // Retrieve the ng-select Angular component instance for a specific host element.
  // el.__ngContext__ is the PARENT form's LView (shared by all ng-selects), so we must
  // match the found component back to `el` via its host-element getter/property.
  function getNgSelectComp(el) {
    // 1. Try Angular's official debug API (works in dev builds and Ivy production)
    try {
      if (typeof ng !== "undefined" && typeof ng.getComponent === "function") {
        const c = ng.getComponent(el);
        if (c && typeof c.open === "function" && c.itemsList) return c;
      }
    } catch (_) {}

    // 2. Search the parent LView for the component whose host element is `el`
    const lView = el.__ngContext__;
    if (!lView || !Array.isArray(lView)) return null;
    for (const item of lView) {
      if (
        item &&
        typeof item.open === "function" &&
        item.itemsList &&
        typeof item.select === "function"
      ) {
        // Verify this instance belongs to our element, not another ng-select
        const hostEl =
          item.element ||                        // ng-select public .element getter
          item._elementRef?.nativeElement ||     // common Angular DI pattern
          item.elementRef?.nativeElement;        // alternative accessor
        if (hostEl === el) return item;
      }
    }
    return null;
  }

  async function fillNgSelect(controlName, value) {
    if (!value) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el) { failed.push(controlName); return; }
    try {
      await ensurePanelOpen(el);
      const comp = getNgSelectComp(el);
      if (comp) {
        // Primary path: use the ng-select component API directly
        comp.open();
        await sleep(500);
        const items = comp.itemsList.items || [];
        const match =
          items.find((i) => i.label === value) ||
          items.find((i) => i.label && i.label.toLowerCase().includes(value.toLowerCase()));
        if (match) {
          comp.select(match);
          comp.close();
          filled.push(controlName);
        } else {
          const available = items.map((i) => i.label).join(", ") || "(no items loaded)";
          errors[controlName] = `No match for "${value}". Available: ${available.slice(0, 120)}`;
          comp.close();
          failed.push(controlName);
        }
      } else {
        // Fallback: DOM click approach (for non-Ivy or unrecognised components)
        errors[controlName] = "Angular context not found — used DOM fallback";
        const trigger = el.querySelector(".ng-select-container") || el;
        trigger.click();
        await sleep(300);
        const searchInput = el.querySelector(".ng-input input");
        if (searchInput && !searchInput.readOnly && !searchInput.disabled) {
          setNativeInput(searchInput, value);
          await sleep(400);
        }
        let options = [];
        for (let i = 0; i < 10; i++) {
          options = [...document.querySelectorAll(".ng-option:not(.ng-option-disabled)")];
          if (options.length > 0) break;
          await sleep(200);
        }
        const match =
          options.find((o) => o.textContent.trim() === value) ||
          options.find((o) => o.textContent.trim().toLowerCase().includes(value.toLowerCase()));
        if (match) { delete errors[controlName]; match.click(); filled.push(controlName); }
        else {
          errors[controlName] = `DOM fallback: no option matching "${value}"`;
          document.body.click(); failed.push(controlName);
        }
      }
      await sleep(200);
    } catch (e) {
      errors[controlName] = e.message;
      console.error("[TourExt] fillNgSelect failed for", controlName, ":", e.message);
      failed.push(controlName);
    }
  }

  async function fillNgSelectMultiple(controlName, values) {
    if (!values || !values.length) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el) { failed.push(controlName); return; }
    try {
      await ensurePanelOpen(el);
      const comp = getNgSelectComp(el);
      if (comp) {
        comp.open();
        await sleep(500);
        const items = comp.itemsList.items || [];
        for (const value of values) {
          const match =
            items.find((i) => i.label === value) ||
            items.find((i) => i.label && i.label.toLowerCase().includes(value.toLowerCase()));
          if (match && !match.selected) {
            comp.select(match);
            await sleep(50);
          }
        }
        comp.close();
        filled.push(controlName);
      } else {
        // Fallback: DOM click approach
        for (const value of values) {
          el.click();
          await sleep(200);
          const searchInput = el.querySelector(".ng-input input");
          if (searchInput && !searchInput.readOnly && !searchInput.disabled) {
            setNativeInput(searchInput, value);
            await sleep(300);
          }
          const options = [...document.querySelectorAll(".ng-option:not(.ng-option-disabled)")];
          const match =
            options.find((o) => o.textContent.trim() === value) ||
            options.find((o) => o.textContent.trim().toLowerCase().includes(value.toLowerCase()));
          if (match) { match.click(); await sleep(100); }
          else { document.body.click(); await sleep(100); }
        }
        filled.push(controlName);
      }
    } catch (e) {
      errors[controlName] = e.message;
      console.error("[TourExt] fillNgSelectMultiple failed for", controlName, ":", e.message);
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
    } catch (e) {
      console.error("[TourExt] fillQuill failed:", e.message);
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
    await fillByControl("tourTitle", tour.title);
    await fillByControl("descriptionWillSee", tour.willSee);
    await fillByControl("descriptionLearn", tour.willLearn);
    await fillByControl("mandatoryInformation", tour.mandatoryInfo);
    await fillByControl("recommendedInformation", tour.recommendedInfo);
    await fillByControl("included", tour.included);
    await fillByControl("notIncluded", tour.notIncluded);
    await fillByControl("noOfPax", tour.noOfPax);
    await fillByControl("longitude", tour.longitude);
    await fillByControl("latitude", tour.latitude);
    await fillByControl("meetingPoint", tour.meetingPoint);
    await fillByControl("pickupInstructions", tour.pickupInstructions);
    await fillByControl("endPoint", tour.endPoint);
    await fillByControl("duration", tour.duration);
    // Prices
    await fillByControl("rate", tour.rate);
    await fillByControl("rateB2C", tour.rateB2C);
    await fillByControl("rate_request", tour.rateRequest);
    await fillByControl("rate_requestB2C", tour.rateRequestB2C);
    await fillByControl("extraHourCharges", tour.extraHour);
    await fillByControl("extraHourChargesB2C", tour.extraHourB2C);
    await fillByControl("extraHourCharges_request", tour.extraHourRequest);
    await fillByControl("extraHourCharges_requestB2C", tour.extraHourRequestB2C);
    await fillByControl("publicHolidaySurchargePercentage", tour.holidaySupplement);
    await fillByControl("weekendSupplementPercentage", tour.weekendSupplement);
    // Schedule
    await fillNgbDatepicker("startDate", tour.startDate);
    await fillNgbDatepicker("endDate", tour.endDate);
    await fillByControl("startTime", tour.startTime);
    await fillByControl("endTime", tour.endTime);
    // Cancellation & cut off — two separate fields for instant vs on-request
    await fillByControl("cancellation", tour.cancellation);
    await fillByControl("cancellation_request", tour.cancellationRequest);
    await fillByControl("release", tour.release);
    await fillByControl("release_request", tour.releaseRequest);

    // Quill
    fillQuill(tour.description);

    // Checkboxes
    setCheckbox("isB2CEnabled", tour.isB2CEnabled);
    setCheckbox("isB2BEnabled", tour.isB2BEnabled);

    // ng-selects (sequential) — cascading: serviceType → activityType → subType
    await fillNgSelect("serviceType", tour.serviceType);
    await sleep(900); // wait for activityType options to cascade-load
    await fillNgSelect("tourType", tour.tourType);
    await fillNgSelect("activityType", tour.activityType);
    await sleep(900); // wait for subType options to cascade-load
    await fillNgSelect("subType", tour.subType);
    await fillNgSelect("activityFor", tour.activityFor);
    await fillNgSelect("voucherType", tour.voucherType);
    await fillNgSelect("countryId", tour.country);
    await fillNgSelect("cityId", tour.city);
    await fillNgSelect("isFixedModel", tour.priceModel);
    await fillNgSelect("currency", tour.currency);
    await fillNgSelect("tourGuideLanguageList", tour.guideLanguageInstant);
    await fillNgSelectMultiple("tourGuideLanguageList_request", tour.guideLanguageRequest);
    await fillNgSelect("tagsList", tour.tags);

    return { success: true, filled, failed, errors };
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
