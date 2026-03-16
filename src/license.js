// ── License Verification Module ─────────────────────────────────────────────
// Exposed as window.LicenseManager so popup.js and background.js can both use it.
//
// Security model:
//   • The license key is verified against the remote Cloudflare Worker.
//   • On success, we persist: { key, verifiedAt, sig }
//     where sig = HMAC-SHA256(key + BIND, String(verifiedAt))
//   • On every popup open we recompute the HMAC and compare — any tampering
//     with the stored key or timestamp breaks the signature and forces re-entry.
//   • After 2 days the grace period expires; the background alarm silently
//     re-verifies and refreshes the token, or wipes it on failure.

(function () {
  const WORKER_URL = "https://upload-assistant.fahadmapari09.workers.dev/verify";
  const STORAGE_KEY = "lic_v1";
  const GRACE_MS = 2 * 24 * 60 * 60 * 1000; // 2 days
  const BIND = "tour-autofill-ext-2026";

  // ── Crypto helpers ─────────────────────────────────────────────────────────

  async function computeHmac(licenseKey, data) {
    const enc = new TextEncoder();
    const cryptoKey = await crypto.subtle.importKey(
      "raw",
      enc.encode(licenseKey + BIND),
      { name: "HMAC", hash: "SHA-256" },
      false,
      ["sign"]
    );
    const sig = await crypto.subtle.sign("HMAC", cryptoKey, enc.encode(data));
    return Array.from(new Uint8Array(sig))
      .map((b) => b.toString(16).padStart(2, "0"))
      .join("");
  }

  function timingSafeEqual(a, b) {
    if (typeof a !== "string" || typeof b !== "string") return false;
    if (a.length !== b.length) return false;
    let diff = 0;
    for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
    return diff === 0;
  }

  // ── Storage helpers ────────────────────────────────────────────────────────

  function readRecord() {
    return new Promise((resolve) =>
      chrome.storage.local.get(STORAGE_KEY, (res) => resolve(res[STORAGE_KEY] || null))
    );
  }

  function writeRecord(record) {
    return new Promise((resolve) =>
      chrome.storage.local.set({ [STORAGE_KEY]: record }, resolve)
    );
  }

  function clearRecord() {
    return new Promise((resolve) =>
      chrome.storage.local.remove(STORAGE_KEY, resolve)
    );
  }

  // ── Network helper ─────────────────────────────────────────────────────────

  async function verifyWithServer(key) {
    try {
      const resp = await fetch(WORKER_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ licenseKey: key }),
      });
      const data = await resp.json().catch(() => ({}));
      return data.valid === true;
    } catch {
      return null; // null = network error (not invalid)
    }
  }

  // ── Public API ─────────────────────────────────────────────────────────────

  /**
   * Verify a key against the remote worker and persist a signed token.
   * Returns { ok: true } or { ok: false, error: string }.
   */
  async function activateLicense(rawKey) {
    const key = (rawKey || "").trim().toUpperCase();
    if (!key) return { ok: false, error: "License key is required." };

    const valid = await verifyWithServer(key);
    if (valid === null) return { ok: false, error: "Could not reach the license server. Check your connection." };
    if (!valid) return { ok: false, error: "Invalid license key." };

    const verifiedAt = Date.now();
    const sig = await computeHmac(key, String(verifiedAt));
    await writeRecord({ key, verifiedAt, sig });
    return { ok: true };
  }

  /**
   * Check whether the extension is currently licensed without hitting the network.
   * Returns:
   *   { licensed: true }
   *   { licensed: false, reason: "no_record" | "corrupt" | "tampered" | "expired", key? }
   */
  async function checkLicense() {
    const rec = await readRecord();
    if (!rec) return { licensed: false, reason: "no_record" };

    const { key, verifiedAt, sig } = rec;
    if (!key || !verifiedAt || !sig) {
      await clearRecord();
      return { licensed: false, reason: "corrupt" };
    }

    // Recompute HMAC — detects any tampering with key or verifiedAt
    const expected = await computeHmac(key, String(verifiedAt));
    if (!timingSafeEqual(expected, sig)) {
      await clearRecord();
      return { licensed: false, reason: "tampered" };
    }

    const age = Date.now() - verifiedAt;
    if (age > GRACE_MS) {
      return { licensed: false, reason: "expired", key };
    }

    return { licensed: true };
  }

  /**
   * Silently re-verify an expired license (called from background alarm).
   * Refreshes token on success; clears storage on failure so next popup shows gate.
   * Returns nothing — fire-and-forget.
   */
  async function silentReVerify(licenseKey) {
    const key = (licenseKey || "").trim().toUpperCase();
    if (!key) { await clearRecord(); return; }

    const valid = await verifyWithServer(key);
    if (valid === null) {
      // Network failure — keep the record so the user isn't locked out offline
      return;
    }
    if (valid) {
      const verifiedAt = Date.now();
      const sig = await computeHmac(key, String(verifiedAt));
      await writeRecord({ key, verifiedAt, sig });
    } else {
      await clearRecord();
    }
  }

  // Expose globally
  window.LicenseManager = { activateLicense, checkLicense, silentReVerify };
})();
