// Background service worker
chrome.runtime.onInstalled.addListener(() => {
  console.log('Tour Admin Autofill installed');
  scheduleAlarmIfNeeded();
});

chrome.runtime.onStartup.addListener(() => {
  scheduleAlarmIfNeeded();
});

// ── License re-verification alarm ───────────────────────────────────────────

const WORKER_URL = "https://upload-assistant.fahadmapari09.workers.dev/verify";
const STORAGE_KEY = "lic_v1";
const BIND = "tour-autofill-ext-2026";
const GRACE_MS = 2 * 24 * 60 * 60 * 1000;

async function hmac(key, data) {
  const enc = new TextEncoder();
  const cryptoKey = await crypto.subtle.importKey(
    "raw",
    enc.encode(key + BIND),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", cryptoKey, enc.encode(data));
  return Array.from(new Uint8Array(sig))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

// Schedule alarm based on how much grace time is left in the stored record.
async function scheduleAlarmIfNeeded() {
  const res = await chrome.storage.local.get(STORAGE_KEY);
  const rec = res[STORAGE_KEY];
  if (!rec || !rec.verifiedAt) return;

  const msLeft = GRACE_MS - (Date.now() - rec.verifiedAt);
  const minutesLeft = Math.max(1, Math.ceil(msLeft / 60000));

  // Only create if not already scheduled
  const existing = await chrome.alarms.get("licenseReVerify");
  if (!existing) {
    chrome.alarms.create("licenseReVerify", { delayInMinutes: minutesLeft });
  }
}

// Popup sends this message after a successful activation so we schedule the alarm.
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === "LICENSE_ACTIVATED") {
    chrome.alarms.clear("licenseReVerify", () => {
      chrome.alarms.create("licenseReVerify", { delayInMinutes: 2 * 24 * 60 });
    });
  }
});

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== "licenseReVerify") return;

  const res = await chrome.storage.local.get(STORAGE_KEY);
  const rec = res[STORAGE_KEY];
  if (!rec || !rec.key) return;

  let resp;
  try {
    resp = await fetch(WORKER_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ licenseKey: rec.key }),
    });
  } catch {
    // Network failure — leave existing record intact, reschedule for 1 hour
    chrome.alarms.create("licenseReVerify", { delayInMinutes: 60 });
    return;
  }

  let data;
  try { data = await resp.json(); } catch { data = {}; }

  if (data.valid) {
    const verifiedAt = Date.now();
    const sig = await hmac(rec.key, String(verifiedAt));
    await chrome.storage.local.set({ [STORAGE_KEY]: { key: rec.key, verifiedAt, sig } });
    // Reschedule for another 2 days
    chrome.alarms.create("licenseReVerify", { delayInMinutes: 2 * 24 * 60 });
  } else {
    await chrome.storage.local.remove(STORAGE_KEY);
  }
});
