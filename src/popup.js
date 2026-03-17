// ── State ──────────────────────────────────────────────
let config = {};
let tours = [];
let selectedTour = null;
let filteredTours = [];
let currentDocTour = null; // parsed tour data from Google Doc
let showReadyOnly = false;

// Fields populated from the Google Doc (not sheet or default)
const DOC_FIELDS = new Set([
  "description", "willSee", "willLearn",
  "included", "notIncluded", "mandatoryInfo",
  "meetingPoint", "endPoint", "longitude", "latitude",
]);

const $ = (id) => document.getElementById(id);

// Column letters are configured explicitly in the setup panel — no auto-detection.

// ── Default data (used for fields not yet coming from sheet) ─
const DEFAULT = {

  activityType: "City Tours",
  subType: "Walking Tours",
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

  startTime: "09:00",
  endTime: "13:00",
};

// ── Init ───────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", async () => {
  const licGate  = document.getElementById("panelLicense");
  const licInput = document.getElementById("licenseKeyInput");
  const licBtn   = document.getElementById("licenseActivateBtn");
  const licError = document.getElementById("licenseError");

  // Attach the click handler FIRST — before any async work
  licBtn.addEventListener("click", async () => {
    const key = licInput.value.trim();
    if (!key) {
      licInput.classList.add("invalid");
      licError.textContent = "Please enter your license key.";
      return;
    }
    licInput.classList.remove("invalid");
    licError.textContent = "";
    licBtn.disabled = true;
    licBtn.innerHTML = '<span class="license-spinner"></span> Verifying…';

    try {
      const result = await window.LicenseManager.activateLicense(key);
      if (result.ok) {
        chrome.runtime.sendMessage({ type: "LICENSE_ACTIVATED" }, () => void chrome.runtime.lastError);
        licGate.classList.add("hidden");
        initApp();
      } else {
        licBtn.disabled = false;
        licBtn.textContent = "Activate";
        licInput.classList.add("invalid");
        licError.textContent = result.error;
      }
    } catch (err) {
      licBtn.disabled = false;
      licBtn.textContent = "Activate";
      licError.textContent = "Unexpected error. Please try again.";
      console.error("[License]", err);
    }
  });

  licInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") licBtn.click();
  });

  // ── License check ──
  let licResult;
  try {
    licResult = await window.LicenseManager.checkLicense();
  } catch (err) {
    console.error("[License] checkLicense error:", err);
    return; // gate stays visible (shown by default in HTML)
  }

  if (licResult.licensed) {
    licGate.classList.add("hidden");
    initApp();
  } else if (licResult.reason === "expired") {
    // Show a checking state — don't force key re-entry until we know the key is actually invalid
    licBtn.disabled = true;
    licInput.disabled = true;
    licBtn.innerHTML = '<span class="license-spinner"></span> Checking license…';

    try {
      const reVerifyResult = await window.LicenseManager.silentReVerify(licResult.key);
      if (reVerifyResult === true) {
        // Key is still valid — schedule next alarm and go straight to the app
        chrome.runtime.sendMessage({ type: "LICENSE_ACTIVATED" }, () => void chrome.runtime.lastError);
        licGate.classList.add("hidden");
        initApp();
      } else {
        // false = revoked, null = network error — show the gate
        licBtn.disabled = false;
        licInput.disabled = false;
        licBtn.textContent = "Activate";
        if (reVerifyResult === false) {
          licError.textContent = "Your license key is no longer valid. Please enter a new key.";
        }
        // null (offline): gate shows with empty error — user can try entering the key
      }
    } catch {
      licBtn.disabled = false;
      licInput.disabled = false;
      licBtn.textContent = "Activate";
    }
  }
  // reason: no_record | corrupt | tampered → gate stays visible as-is
});

async function initApp() {
  const { version } = chrome.runtime.getManifest();
  const vEl = document.getElementById("extVersion");
  if (vEl) vEl.textContent = `v${version}`;

  config = await loadConfig();

  if (config.sheetId) {
    $("sheetId").value  = config.sheetId;
    $("sheetTab").value = config.sheetTab || "Sheet1";
    if (config.apiKey) $("apiKey").value = config.apiKey; // legacy field
    // Restore column inputs
    const cols = [
      ["colTitle", "F"],
      ["colDocUrl", "F"],
      ["colCountry", "A"],
      ["colCity", "B"],
      ["colDuration", "G"],
      ["colServiceType", "E"],
      ["colRate", "AK"],
      ["colRateRequest", "AK"],
      ["colRateB2C", "AL"],
      ["colRateRequestB2C", "AM"],
      ["colCancellation", "AT"],
      ["colCancellationRequest", "AV"],
      ["colRelease", "AU"],
      ["colReleaseRequest", "AW"],
      ["colExtraHour", ""],
      ["colExtraHourB2C", ""],
      ["colExtraHourRequest", ""],
      ["colExtraHourRequestB2C", ""],
      ["colMaxPax", "O"],
      ["colReadyForUpload", ""],
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
    "colExtraHour",
    "colExtraHourB2C",
    "colExtraHourRequest",
    "colExtraHourRequestB2C",
    "colMaxPax",
    "colReadyForUpload",
  ].forEach((id) => {
    const el = $(id);
    if (el) el.addEventListener("input", autoSaveConfig);
  });
  $("searchInput").addEventListener("input", (e) =>
    filterTours(e.target.value),
  );
  $("readyOnlyToggle").addEventListener("change", (e) => {
    showReadyOnly = e.target.checked;
    filterTours($("searchInput").value);
  });
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
}

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
    colTitle: col("colTitle") || "F",
    colDocUrl: col("colDocUrl") || "F",
    colCountry: col("colCountry") || "A",
    colCity: col("colCity") || "B",
    colDuration: col("colDuration") || "G",
    colServiceType: col("colServiceType") || "E",
    colRate: col("colRate") || "AK",
    colRateRequest: col("colRateRequest") || "AK",
    colRateB2C: col("colRateB2C") || "AL",
    colRateRequestB2C: col("colRateRequestB2C") || "AM",
    colCancellation: col("colCancellation") || "AT",
    colCancellationRequest: col("colCancellationRequest") || "AV",
    colRelease: col("colRelease") || "AU",
    colReleaseRequest: col("colReleaseRequest") || "AW",
    colExtraHour: col("colExtraHour") || "",
    colExtraHourB2C: col("colExtraHourB2C") || "",
    colExtraHourRequest: col("colExtraHourRequest") || "",
    colExtraHourRequestB2C: col("colExtraHourRequestB2C") || "",
    colMaxPax: col("colMaxPax") || "O",
    colReadyForUpload: col("colReadyForUpload") || "",
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
const CACHE_TTL_MS = 12 * 60 * 60 * 1000; // 12 hours

function saveCachedTours(data) {
  return new Promise((resolve) =>
    chrome.storage.local.set({ tourExtCache: { data, savedAt: Date.now() } }, () => {
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
        const entry = r.tourExtCache;
        if (!entry) { resolve(null); return; }
        // Support legacy cache entries that stored the array directly
        if (Array.isArray(entry)) { resolve(null); return; }
        if (Date.now() - entry.savedAt > CACHE_TTL_MS) { resolve(null); return; }
        resolve(entry.data || null);
      }
    }),
  );
}

// ── Skeleton loader ────────────────────────────────────
function showSkeletonLoader(count = 6) {
  const container = $("tourList");
  container.onscroll = null;
  container.innerHTML = Array.from({ length: count }, () => `
    <div class="skeleton-card">
      <div class="skeleton-line skeleton-title"></div>
      <div class="skeleton-line skeleton-meta"></div>
    </div>`).join("");
}

// ── Google Sheets fetch ────────────────────────────────
// Fetches all columns, maps by header name using COLUMN_MAP
async function loadTours(forceFresh = false) {
  setStatus("busy");

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

  showSkeletonLoader();
  $("tourCountLabel").innerHTML =
    `<span style="display:inline-flex;align-items:center;gap:5px">` +
    `<span style="display:inline-block;width:9px;height:9px;border:1.5px solid var(--border-light);border-top-color:var(--muted-fg);border-radius:50%;animation:spin 0.6s linear infinite;flex-shrink:0"></span>` +
    `Fetching from Google Sheets…</span>`;

  try {
    const range = `'${config.sheetTab}'`; // no column bounds — fetch entire sheet
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

    // Fetch hyperlinks from the title column (URLs embedded as hyperlinks, not plain text)
    const titleColLetter = config.colTitle || "A";
    const hlRange = `'${config.sheetTab}'!${titleColLetter}:${titleColLetter}`;
    const hlUrl = `https://sheets.googleapis.com/v4/spreadsheets/${config.sheetId}?includeGridData=true&ranges=${encodeURIComponent(hlRange)}&fields=sheets.data.rowData.values(hyperlink,userEnteredValue)`;
    let titleHyperlinks = [];
    try {
      const hlRes = await authedFetch(hlUrl, true);
      if (hlRes.ok) {
        const hlData = await hlRes.json();
        const rowData = hlData.sheets?.[0]?.data?.[0]?.rowData || [];
        titleHyperlinks = rowData.map((r) => {
          const cell = r.values?.[0];
          if (!cell) return "";
          if (cell.hyperlink) return cell.hyperlink;
          const formula = cell.userEnteredValue?.formulaValue || "";
          const m = formula.match(/=HYPERLINK\s*\(\s*"([^"]+)"/i);
          return m ? m[1] : "";
        });
        // Index 0 is the header row — shift so index 0 = first data row
        titleHyperlinks = titleHyperlinks.slice(1);
        console.log("[TourExt] Hyperlinks fetched:", titleHyperlinks.filter(Boolean).length);
      } else {
        const errText = await hlRes.text();
        console.warn("[TourExt] Hyperlink fetch failed:", hlRes.status, errText);
      }
    } catch (hlErr) {
      console.warn("[TourExt] Could not fetch title hyperlinks:", hlErr.message);
    }

    const headers = rows[0].map((h) => h.toLowerCase().trim());

    tours = rows
      .slice(1)
      .map((row, i) => buildTour(headers, row, i + 2, titleHyperlinks[i] || ""))
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

// Strip currency symbols, commas, and whitespace — keep only digits and decimal point
function cleanRate(val) {
  if (!val) return val;
  const cleaned = val.replace(/[^0-9.]/g, "");
  return cleaned || val;
}

// Extract the first number from a cancellation value and return the policy sentence
function formatCancellation(val) {
  if (!val) return val;
  const match = String(val).match(/\d+/);
  if (!match) return val;
  return `Cancel up to ${match[0]} office days in advance for a full refund`;
}

function buildTour(headers, row, rowNum, titleHyperlink = "") {
  // Helper: get cell value by configured column letter (e.g. "A" → index 0)
  const col = (cfgKey) => {
    const letter = config[cfgKey];
    if (!letter) return "";
    const idx = colLetterToIndex(letter);
    return (row[idx] || "").trim();
  };

  // Prefer the hyperlink embedded in the title cell; fall back to a plain-text docUrl column
  const docUrl = titleHyperlink || col("colDocUrl");

  const serviceType = col("colServiceType");

  // Parse maxPax: "1-15" → "15", plain number → use as-is,
  // empty → 7 for Day Trip, 15 otherwise
  const rawMaxPax = col("colMaxPax");
  let maxPax;
  if (rawMaxPax) {
    const dashMatch = rawMaxPax.match(/-(\d+)/);
    maxPax = dashMatch ? dashMatch[1] : rawMaxPax;
  } else {
    maxPax = /day.?trip/i.test(serviceType) ? "7" : "15";
  }

  return {
    rowNum,
    title: col("colTitle"),
    docUrl,
    country: col("colCountry"),
    city: col("colCity"),
    duration: col("colDuration"),
    serviceType,
    rate: cleanRate(col("colRate")),
    rateRequest: cleanRate(col("colRateRequest")),
    rateB2C: cleanRate(col("colRateB2C")),
    rateRequestB2C: cleanRate(col("colRateRequestB2C")),
    cancellation: col("colCancellation"),
    cancellationRequest: col("colCancellationRequest"),
    release: col("colRelease"),
    releaseRequest: col("colReleaseRequest"),
    extraHour: cleanRate(col("colExtraHour")),
    extraHourB2C: cleanRate(col("colExtraHourB2C")),
    extraHourRequest: cleanRate(col("colExtraHourRequest")),
    extraHourRequestB2C: cleanRate(col("colExtraHourRequestB2C")),
    maxPax,
    readyForUpload: col("colReadyForUpload").toUpperCase() === "TRUE",
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
  let base = showReadyOnly ? tours.filter((t) => t.readyForUpload) : [...tours];
  filteredTours = q
    ? base.filter(
        (t) =>
          (t.title || "").toLowerCase().includes(q) ||
          (t.country || "").toLowerCase().includes(q) ||
          (t.city || "").toLowerCase().includes(q),
      )
    : base;
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
  { key: "cancellation", label: "Cancellation", source: "sheet" },
  { key: "release", label: "Cut Off", source: "sheet" },
  { key: "description", label: "Description (Quill)", source: "doc" },

  { key: "activityType", label: "Activity Type", source: "default" },
  { key: "subType", label: "Sub Type", source: "default" },
  { key: "willSee", label: "You Will See", source: "doc" },
  { key: "willLearn", label: "You Will Learn", source: "doc" },
  { key: "mandatoryInfo", label: "Mandatory Information", source: "default" },
  { key: "recommendedInfo", label: "Recommended Information", source: "default" },
  { key: "included", label: "Included", source: "doc" },
  { key: "notIncluded", label: "Not Included", source: "doc" },
  { key: "activityFor", label: "Activity For", source: "default" },
  { key: "voucherType", label: "Voucher Type", source: "default" },
  { key: "noOfPax", label: "No of Pax", source: "sheet" },
  {
    key: "guideLanguageInstant",
    label: "Guide Language (Instant)",
    source: "default",
  },
  {
    key: "guideLanguageRequest",
    label: "Guide Language (Request)",
    source: "default",
  },
  { key: "longitude", label: "Longitude", source: "doc" },
  { key: "latitude", label: "Latitude", source: "doc" },
  { key: "meetingPoint", label: "Meeting Point", source: "default" },
  { key: "pickupInstructions", label: "Pickup Instructions", source: "default" },
  { key: "endPoint", label: "End Point", source: "default" },
  { key: "tags", label: "Tags", source: "default" },

  { key: "extraHour", label: "Extra Hour Supplement (Instant)", source: "sheet" },
  { key: "extraHourB2C", label: "Extra Hour Supplement B2C (Instant)", source: "sheet" },
  { key: "extraHourRequest", label: "Extra Hour Supplement (On Request)", source: "sheet" },
  { key: "extraHourRequestB2C", label: "Extra Hour Supplement B2C (On Request)", source: "sheet" },
  { key: "startDate", label: "Start Date", source: "default" },
  { key: "endDate", label: "End Date", source: "default" },
  { key: "startTime", label: "Start Time", source: "default" },
  { key: "endTime", label: "End Time", source: "default" },
];

// ── Doc tour matching ───────────────────────────────────
function findMatchingDocTour(allTours, sheetTitle) {
  const norm = (s) => s.toLowerCase().replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();
  const target = norm(sheetTitle);
  return (
    allTours.find((t) => norm(t.title) === target) ||
    allTours.find((t) => norm(t.title).includes(target) || target.includes(norm(t.title)))
  ) || null;
}

// Merges parsed doc data into an existing fillData object (mutates in place)
function mergeDocData(fillData) {
  if (!currentDocTour) return;
  const d = currentDocTour;
  if (d.description)        fillData.description  = d.description;
  if (d.youWillSee?.length) fillData.willSee       = d.youWillSee.join("\n");
  if (d.youWillLearn?.length) fillData.willLearn   = d.youWillLearn.join("\n");
  if (d.inclusions?.length) fillData.included      = d.inclusions.join(",");
  if (d.exclusions?.length) fillData.notIncluded   = d.exclusions.join(",");
  if (d.additionalInfo?.length) fillData.recommendedInfo = d.additionalInfo.join(",");
  if (d.meetingPoint)       fillData.meetingPoint  = d.meetingPoint;
  if (d.endLocation)        fillData.endPoint      = d.endLocation;
  if (d.longitude)          fillData.longitude     = d.longitude;
  if (d.latitude)           fillData.latitude      = d.latitude;
}

async function goToFillPanel(tour) {
  currentDocTour = null;

  const _now = new Date();
  const _end = new Date(_now);
  _end.setFullYear(_end.getFullYear() + 2);
  const _fmt = (d) => `${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}/${d.getFullYear()}`;
  const _durHours = parseInt((tour.duration || "").replace(/h.*/i, "")) || 0;
  const _endTime = `${String(Math.max(0, 22 - _durHours)).padStart(2, "0")}:00`;

  function buildFillData() {
    const data = {
      ...DEFAULT,
      title: tour.title,
      serviceType: "Guide",
      subType: (tour.serviceType === "Driver-Guide" || /day.?trip/i.test(tour.serviceType)) ? "Driver Guide" : "Walking Tours",
      noOfPax: tour.maxPax,
      country: tour.country || DEFAULT.country,
      city: tour.city || DEFAULT.city,
      duration: tour.duration || currentDocTour?.duration || "",
      rate: tour.rate || DEFAULT.rate,
      rateRequest: tour.rateRequest || DEFAULT.rateRequest,
      rateB2C: tour.rateB2C || DEFAULT.rateB2C,
      rateRequestB2C: tour.rateRequestB2C || DEFAULT.rateRequestB2C,
      cancellation: formatCancellation(tour.cancellationRequest || tour.cancellation || DEFAULT.cancellation),
      release: tour.releaseRequest || tour.release || DEFAULT.release,
      extraHour: tour.extraHour || null,
      extraHourB2C: tour.extraHourB2C || null,
      extraHourRequest: tour.extraHourRequest || null,
      extraHourRequestB2C: tour.extraHourRequestB2C || null,
      startDate: _fmt(_now),
      endDate: _fmt(_end),
      startTime: "08:00",
      endTime: _endTime,
    };
    mergeDocData(data);
    return data;
  }

  function renderPreview(fillData, docBadge) {
    $("previewName").textContent = tour.title;

    const raw = tour.docUrl || "";
    const url = raw && !/^https?:\/\//i.test(raw) ? "https://" + raw : raw;
    const cell = url
      ? `<a href="${url}" target="_blank" title="${url}">${url.length > 50 ? url.slice(0, 50) + "…" : url}</a>`
      : `<span style="color:var(--muted)">— No doc linked</span>`;
    const badge = docBadge
      ? `&nbsp;<span style="color:${docBadge.color};font-size:9px;font-family:var(--mono)">${docBadge.text}</span>`
      : "";
    const docUrlRow = `<tr style="border-bottom:1px solid var(--border)">
      <td>Doc URL</td>
      <td colspan="2">${cell}${badge}</td>
    </tr>`;

    $("previewFields").innerHTML =
      `<table class="preview-table">${docUrlRow}` +
      FILL_FIELDS.map((f) => {
        const val = fillData[f.key];
        const display = Array.isArray(val)
          ? val.join(", ")
          : (val != null && val !== "" ? String(val) : "—");
        const src = f.source === "sheet"
          ? "sheet"
          : (currentDocTour && DOC_FIELDS.has(f.key) ? "doc" : "default");
        const valColor = src === "sheet" ? "var(--text-dim)" : src === "doc" ? "#a5b4fc" : "var(--muted-fg)";
        const srcColor = src === "sheet" ? "var(--success)" : src === "doc" ? "#818cf8" : "var(--muted)";
        return `<tr>
          <td>${f.label}</td>
          <td style="color:${valColor}">${display}</td>
          <td style="color:${srcColor}">${src}</td>
        </tr>`;
      }).join("") +
      `</table>`;

    $("fieldsChecklist").innerHTML = FILL_FIELDS.map((f) => {
      const src = f.source === "sheet"
        ? "sheet"
        : (currentDocTour && DOC_FIELDS.has(f.key) ? "doc" : "default");
      const srcColor = src === "sheet" ? "var(--success)" : src === "doc" ? "#818cf8" : "var(--muted)";
      return `<div class="check-item" id="check-${f.key}">
        <div class="check-dot"></div>
        <span>${f.label}</span>
        <span style="margin-left:auto;font-size:9px;font-family:var(--mono);color:${srcColor}">${src}</span>
      </div>`;
    }).join("");
  }

  // Show the panel immediately with default/sheet data while doc loads
  renderPreview(buildFillData(), tour.docUrl ? { text: "fetching doc…", color: "var(--warning)" } : null);
  showPanel("panelFill");
  $("startFillBtn").disabled = true;
  $("startFillBtn").innerHTML = '<div class="spinner"></div> Loading doc…';

  // Fetch and parse the Google Doc to get real content
  if (tour.docUrl) {
    try {
      const docJson = await fetchGoogleDoc(tour.docUrl);
      if (docJson) {
        const allTours = parseParisDoc(docJson);
        currentDocTour = findMatchingDocTour(allTours, tour.title);
        if (currentDocTour) {
          renderPreview(buildFillData(), { text: "✓ doc parsed", color: "var(--success)" });
        } else {
          renderPreview(buildFillData(), { text: "title not found in doc", color: "var(--warning)" });
          showToast("Tour title not found in doc — using default data for doc fields", "info");
        }
      }
    } catch (e) {
      console.warn("[TourExt] Doc fetch/parse failed:", e.message);
      renderPreview(buildFillData(), { text: `doc fetch failed: ${e.message}`, color: "var(--danger)" });
    }
  }

  $("startFillBtn").disabled = false;
  $("startFillBtn").innerHTML =
    '<svg width="11" height="11" viewBox="0 0 24 24" fill="currentColor" stroke="none"><polygon points="5 3 19 12 5 21 5 3"/></svg> Start Autofill';
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

  // Compute end time: 22:00 minus duration hours (e.g. "2h" → 20:00)
  const _durHours = parseInt((selectedTour.duration || currentDocTour?.duration || "").replace(/h.*/i, "")) || 0;
  const _endHour = Math.max(0, 22 - _durHours);
  const _endTime = `${String(_endHour).padStart(2, "0")}:00`;

  // Merge: default base → overridden by real sheet values where available
  const fillData = {
    ...DEFAULT,
    title: selectedTour.title,
    serviceType: "Guide",
    subType: (selectedTour.serviceType === "Driver-Guide" || /day.?trip/i.test(selectedTour.serviceType)) ? "Driver Guide" : "Walking Tours",
    noOfPax: selectedTour.maxPax,
    country: selectedTour.country || DEFAULT.country,
    city: selectedTour.city || DEFAULT.city,
    duration: selectedTour.duration || currentDocTour?.duration || "",
    rate: selectedTour.rate || DEFAULT.rate,
    rateRequest: selectedTour.rateRequest || DEFAULT.rateRequest,
    rateB2C: selectedTour.rateB2C || DEFAULT.rateB2C,
    rateRequestB2C: selectedTour.rateRequestB2C || DEFAULT.rateRequestB2C,
    cancellation: formatCancellation(selectedTour.cancellationRequest || selectedTour.cancellation || DEFAULT.cancellation),
    release: selectedTour.releaseRequest || selectedTour.release || DEFAULT.release,
    extraHour: selectedTour.extraHour || null,
    extraHourB2C: selectedTour.extraHourB2C || null,
    extraHourRequest: selectedTour.extraHourRequest || null,
    extraHourRequestB2C: selectedTour.extraHourRequestB2C || null,
    startDate: _fmt(_now),
    endDate: _fmt(_end),
    startTime: "08:00",
    endTime: _endTime,
  };

  // Override default doc fields with parsed Google Doc data if available
  mergeDocData(fillData);

  try {
    const [tab] = await chrome.tabs.query({
      active: true,
      currentWindow: true,
    });
    if (!tab) {
      throw new Error("No active tab found");
    }
    const REQUIRED_URL = "https://trav-ui-admin-prod.azurewebsites.net/#/admin/product/add";
    if (tab.url !== REQUIRED_URL) {
      throw new Error(`Autofill only works on the product add page.\nPlease navigate to:\n${REQUIRED_URL}`);
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
      $("startFillBtn").disabled = true;
      $("startFillBtn").innerHTML =
        '<svg width="11" height="11" viewBox="0 0 24 24" fill="currentColor" stroke="none"><polyline points="20 6 9 17 4 12" style="fill:none;stroke:currentColor;stroke-width:3"/></svg> Autofill Completed';
      $("startFillBtn").style.opacity = "0.6";
      return;
    } else {
      throw new Error(outcome?.error || "Unknown error");
    }
  } catch (e) {
    console.error("[TourExt] startFill failed:", e.message);
    showToast("Fill failed: " + e.message, "error");
    setStatus("error");
    $("startFillBtn").disabled = false;
    $("startFillBtn").innerHTML =
      '<svg width="11" height="11" viewBox="0 0 24 24" fill="currentColor" stroke="none"><polygon points="5 3 19 12 5 21 5 3"/></svg> Start Autofill';
    $("startFillBtn").style.opacity = "";
  }
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

  async function fillByControl(controlName, value, reportKey) {
    const key = reportKey || controlName;
    if (!value && value !== 0) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el || el.disabled || el.getAttribute("readonly") === "readonly") return;
    await ensurePanelOpen(el);
    const tag = el.tagName.toLowerCase();
    if (tag === "input" || tag === "textarea") {
      if (setNativeInput(el, value)) filled.push(key);
      else failed.push(key);
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

  async function fillNgSelect(controlName, value, reportKey) {
    const key = reportKey || controlName;
    if (!value) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el) { failed.push(key); return; }
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
          filled.push(key);
        } else {
          const available = items.map((i) => i.label).join(", ") || "(no items loaded)";
          errors[key] = `No match for "${value}". Available: ${available.slice(0, 120)}`;
          comp.close();
          failed.push(key);
        }
      } else {
        // Fallback: DOM click approach (for non-Ivy or unrecognised components)
        errors[key] = "Angular context not found — used DOM fallback";
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
        if (match) { delete errors[key]; match.click(); filled.push(key); }
        else {
          errors[key] = `DOM fallback: no option matching "${value}"`;
          document.body.click(); failed.push(key);
        }
      }
      await sleep(200);
    } catch (e) {
      errors[key] = e.message;
      console.error("[TourExt] fillNgSelect failed for", controlName, ":", e.message);
      failed.push(key);
    }
  }

  async function fillNgSelectMultiple(controlName, values, reportKey) {
    const key = reportKey || controlName;
    if (!values || !values.length) return;
    const el = document.querySelector(`[formcontrolname="${controlName}"]`);
    if (!el) { failed.push(key); return; }
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
        filled.push(key);
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
        filled.push(key);
      }
    } catch (e) {
      errors[key] = e.message;
      console.error("[TourExt] fillNgSelectMultiple failed for", controlName, ":", e.message);
      failed.push(key);
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

  async function runFill() {
    // Expand all collapsed mat-expansion-panels so Angular renders their content into the DOM.
    // fillByControl uses querySelector which returns null for elements inside collapsed panels
    // (Angular lazy-renders panel content — it doesn't exist in the DOM until first opened).
    const collapsedHeaders = document.querySelectorAll(
      'mat-expansion-panel-header[aria-expanded="false"]'
    );
    collapsedHeaders.forEach((h) => h.click());
    if (collapsedHeaders.length > 0) await sleep(800);

    // Text / number inputs
    await fillByControl("tourTitle", tour.title, "title");
    await fillByControl("descriptionWillSee", tour.willSee, "willSee");
    await fillByControl("descriptionLearn", tour.willLearn, "willLearn");
    await fillByControl("mandatoryInformation", tour.mandatoryInfo, "mandatoryInfo");
    await fillByControl("recommendedInformation", tour.recommendedInfo, "recommendedInfo");
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
    await fillByControl("rate_request", tour.rateRequest, "rateRequest");
    await fillByControl("rate_requestB2C", tour.rateRequestB2C, "rateRequestB2C");
    await fillByControl("extraHourCharges", tour.extraHour, "extraHour");
    await fillByControl("extraHourChargesB2C", tour.extraHourB2C, "extraHourB2C");
    await fillByControl("extraHourCharges_request", tour.extraHourRequest, "extraHourRequest");
    await fillByControl("extraHourCharges_requestB2C", tour.extraHourRequestB2C, "extraHourRequestB2C");
    // Schedule
    await fillNgbDatepicker("startDate", tour.startDate);
    await fillNgbDatepicker("endDate", tour.endDate);
    await fillByControl("startTime", tour.startTime);
    await fillByControl("endTime", tour.endTime);
    // Cancellation & cut off — two separate fields for instant vs on-request
    await fillByControl("cancellation", tour.cancellation);
    await fillByControl("release", tour.release);

    // Quill
    fillQuill(tour.description);

    // ng-selects (sequential) — cascading: serviceType → activityType → subType
    await fillNgSelect("serviceType", tour.serviceType);
    await sleep(900); // wait for activityType options to cascade-load

    await fillNgSelect("activityType", tour.activityType);
    await sleep(900); // wait for subType options to cascade-load
    await fillNgSelect("subType", tour.subType);
    await fillNgSelect("activityFor", tour.activityFor);
    await fillNgSelect("voucherType", tour.voucherType);
    await fillNgSelect("countryId", tour.country, "country");
    await sleep(900); // wait for city options to cascade-load after country selection
    await fillNgSelect("cityId", tour.city, "city");

    await fillNgSelect("tourGuideLanguageList", tour.guideLanguageInstant, "guideLanguageInstant");
    await fillNgSelectMultiple("tourGuideLanguageList_request", tour.guideLanguageRequest, "guideLanguageRequest");
    await fillNgSelect("tagsList", tour.tags, "tags");

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
