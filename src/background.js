// Background service worker
chrome.runtime.onInstalled.addListener(() => {
  console.log('Tour Admin Autofill installed');
});

// ── License re-verification alarm ───────────────────────────────────────────
// The popup schedules this alarm when the grace period has expired.
// We re-verify silently here so the result is ready by the next popup open.

const WORKER_URL = "https://upload-assistant.fahadmapari09.workers.dev/verify";
const STORAGE_KEY = "lic_v1";
const BIND = "tour-autofill-ext-2026";

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

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== "licenseReVerify") return;

  const res = await chrome.storage.local.get(STORAGE_KEY);
  const rec = res[STORAGE_KEY];
  if (!rec || !rec.key) return; // nothing to verify

  let resp;
  try {
    resp = await fetch(WORKER_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ licenseKey: rec.key }),
    });
  } catch {
    // Network failure — leave existing record intact, alarm will fire again
    return;
  }

  let data;
  try { data = await resp.json(); } catch { data = {}; }

  if (data.valid) {
    const verifiedAt = Date.now();
    const sig = await hmac(rec.key, String(verifiedAt));
    await chrome.storage.local.set({ [STORAGE_KEY]: { key: rec.key, verifiedAt, sig } });
  } else {
    // Key is no longer valid — wipe it so the gate appears on next open
    await chrome.storage.local.remove(STORAGE_KEY);
  }
});
