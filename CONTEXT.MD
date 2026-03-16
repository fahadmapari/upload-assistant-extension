# Tour Admin Autofill — Chrome Extension Context

## What this is

A Chrome extension that reads tour data from a Google Sheet and Google Docs, then autofills an Angular admin panel form at `/admin/product/add`. Staff previously copied data manually — this automates it. The human reviews the filled form and clicks Save.

Access is gated behind a license key verified against a Cloudflare Worker.

---

## Current status

- License gate system fully implemented with HMAC-based local validation and remote re-verification ✓
- Google Sheets loading via OAuth (no API key) fully working ✓
- Google Docs fetching and parsing fully implemented and wired up ✓
- `parisDocParser.js` extracts 15+ fields from structured tour docs ✓
- Three-source data merge (defaults → sheet → doc) working ✓
- Autofill covering 45+ fields across all form sections ✓
- Virtual scroll for large tour lists ✓
- Background alarm for silent license re-verification every 2 days ✓
- Extension version displayed in popup header ✓

---

## File structure

```
tour-extension/
├── manifest.json           # MV3 — permissions: activeTab, scripting, storage, identity, alarms
├── popup.html              # Extension UI (380px popup, 4 panels)
├── icons/                  # icon16/32/48/128.png
├── CONTEXT.MD              # This file
└── src/
    ├── popup.js            # Main logic — config, OAuth, sheet/doc fetch, tour list, autofill (~1371 lines)
    ├── background.js       # Service worker — license re-verification alarm (~91 lines)
    ├── license.js          # HMAC-SHA256 license module (window.LicenseManager)
    ├── parisDocParser.js   # Google Docs JSON parser, extracts tour fields (~394 lines)
    └── content.js          # Minimal content script — PING handler only (10 lines)
```

---

## Architecture overview

### Panels (popup.js)

Four mutually exclusive panels toggled via `showPanel(id)`:

| Panel ID       | Purpose                                          |
| -------------- | ------------------------------------------------ |
| `panelSetup`   | First-run config (sheet ID, tab, column letters) |
| `panelList`    | Tour list with search, virtual scroll            |
| `panelFill`    | Preview merged data, field checklist, start fill |
| `panelLicense` | License gate overlay (shown when not licensed)   |

### Data sources & merge order

```
defaults (DUMMY)
  ↑ overridden by
sheet data (Google Sheets API)
  ↑ overridden by
doc data (Google Docs API → parisDocParser)
```

The merged object is passed to `injectTourData()` injected into the admin page.

---

## License system

### Flow

1. Popup opens → `LicenseManager.checkLicense()` reads `lic_v1` from `chrome.storage.local`
2. If valid HMAC → hide `panelLicense`, show app
3. If invalid/missing → show `panelLicense` gate
4. User enters key → `LicenseManager.activateLicense(key)` → POST to Cloudflare Worker → store `{key, verifiedAt, sig}`

### Background re-verification

- `background.js` schedules a Chrome alarm every 2 days
- On alarm: reads stored key, POSTs to worker, updates `verifiedAt + sig` or clears storage
- Network errors → reschedule in 1 hour (offline-friendly grace period)

### Storage key: `lic_v1`

```js
{
  key: "LICENSE_KEY_STRING",
  verifiedAt: 1234567890000,  // ms timestamp
  sig: "deadbeef..."          // HMAC-SHA256(key + BIND_SECRET, verifiedAt)
}
```

### Cloudflare Worker

`POST https://upload-assistant.fahadmapari09.workers.dev/verify`
- Body: `{ licenseKey: "string" }`
- Response: `{ valid: true | false }`
- CORS locked to `chrome-extension://lfbleghapiiefannabbdccgenaniobgj`
- Valid keys hardcoded in worker's `VALID_KEYS` Set (see `worker.js` in parent folder)

---

## Google Sheets integration

- Auth: OAuth via `chrome.identity.getAuthToken()` — no API key
- API: Sheets API v4
- Three separate calls per load:
  1. `values/{range}` — all cell values
  2. `?includeGridData=true` — hyperlink extraction from title column
  3. `spreadsheets/{id}` — sheet title (display only)
- Column mapping: user configures explicit letter (A, B, AA, etc.) in settings — NOT header-based auto-detection
- Cache: full tour array stored in `tourExtCache` in `chrome.storage.local`; replaced on manual refresh

### Config keys (`tourExtConfig`)

```js
{
  sheetId, sheetTab,
  colTitle, colDocUrl, colCountry, colCity, colDuration, colServiceType, colMaxPax,
  colRate, colRateRequest, colRateB2C, colRateRequestB2C,
  colCancellation, colCancellationRequest,
  colRelease, colReleaseRequest,
  colExtraHour, colExtraHourB2C, colExtraHourRequest, colExtraHourRequestB2C,
}
```

---

## Google Docs integration

- Auth: OAuth via `chrome.identity.getAuthToken()` — same token as Sheets
- API: Docs API v1 — `GET /v1/documents/{documentId}`
- Triggered in `goToFillPanel()` when selected tour has a `docUrl`
- Parser: `parseParisDoc(docJson)` in `parisDocParser.js`

### Parser logic (`parisDocParser.js`)

- Splits document by `HEADING_1` boundaries → one section per tour title
- Finds the correct section via fuzzy title matching (strips whitespace/punctuation)
- Extracts:
  - **List fields**: inclusions, exclusions, additionalInfo, youWillLearn, youWillSee
  - **Inline fields**: meetingPoint, duration, endLocation, coordinates (lat/lng pair)
  - **Block text**: description (accumulated until next section header)
- Handles soft line breaks (separate `textRun` elements within same paragraph)
- Skips noise: review flags (`CC OK`, `IS OK`, `SS note`), platform distribution mentions
- Switzerland variant: maps "Highlights" section header → `youWillSee`

### Doc fields returned

```js
{
  description, inclusions, exclusions, additionalInfo,
  youWillLearn, youWillSee, meetingPoint, duration,
  endLocation, latitude, longitude
}
```

---

## Admin panel tech stack

- Angular application with reactive forms
- `formcontrolname` attributes are the stable selectors
- Angular Material accordion (`mat-accordion`) — **all panels must be open before autofill starts**
- `ng-select` custom dropdown components (single and multi-select)
- Quill rich text editor for description
- `ngbdatepicker` for date fields
- Form groups: root, `formarrayname="geo"` for country/city, `formgroupname="price"` for pricing

---

## Form field → formcontrolname mapping

### Section 1 — Product Description

| Field                   | formcontrolname                 | Type              |
| ----------------------- | ------------------------------- | ----------------- |
| Tour Title              | `tourTitle`                     | text input        |
| Service Type            | `serviceType`                   | ng-select single  |
| Tour Type               | `tourType`                      | ng-select single  |
| Activity Type           | `activityType`                  | ng-select single  |
| Sub Type                | `subType`                       | ng-select single  |
| Description             | `.ql-editor`                    | Quill rich text   |
| You Will See            | `descriptionWillSee`            | textarea          |
| You Will Learn          | `descriptionLearn`              | textarea          |
| Tour Language           | `tourLanguageId`                | ng-select (SKIP)  |
| Guide Language Instant  | `tourGuideLanguageList`         | ng-select multi   |
| Guide Language Request  | `tourGuideLanguageList_request` | ng-select multi   |
| Mandatory Info          | `mandatoryInformation`          | text input        |
| Recommended Info        | `recommendedInformation`        | text input        |
| Included                | `included`                      | textarea          |
| Not Included            | `notIncluded`                   | textarea          |
| Activity For            | `activityFor`                   | ng-select single  |
| Voucher Type            | `voucherType`                   | ng-select single  |
| No of Pax               | `noOfPax`                       | number input      |
| Country                 | `countryId`                     | ng-select (in geo array) |
| City                    | `cityId`                        | ng-select (in geo array) |
| Longitude               | `longitude`                     | text input        |
| Latitude                | `latitude`                      | text input        |
| Meeting Point           | `meetingPoint`                  | text input        |
| Pickup Instructions     | `pickupInstructions`            | text input        |
| End Point               | `endPoint`                      | text input        |
| Tags                    | `tagsList`                      | ng-select multi   |

### Section 2 — Product Images

**SKIPPED — manual only**

### Section 3 — Price & Schedule (`formgroupname="price"`)

| Field                    | formcontrolname                    | Type                        |
| ------------------------ | ---------------------------------- | --------------------------- |
| Price Model              | `isFixedModel`                     | ng-select                   |
| Currency                 | `currency`                         | ng-select                   |
| Rate (Instant B2B)       | `rate`                             | number input                |
| Rate (Instant B2C)       | `rateB2C`                          | number input                |
| Rate (Request B2B)       | `rate_request`                     | number input                |
| Rate (Request B2C)       | `rate_requestB2C`                  | number input                |
| Extra Hour (Instant B2B) | `extraHourCharges`                 | number input                |
| Extra Hour (Instant B2C) | `extraHourChargesB2C`              | number input                |
| Extra Hour (Request B2B) | `extraHourCharges_request`         | number input                |
| Extra Hour (Request B2C) | `extraHourCharges_requestB2C`      | number input                |
| Holiday Supplement %     | `publicHolidaySurchargePercentage` | number input                |
| Weekend Supplement %     | `weekendSupplementPercentage`      | number input                |
| Start Date               | `startDate`                        | ngbdatepicker               |
| End Date                 | `endDate`                          | ngbdatepicker               |
| Start Time               | `startTime`                        | time input                  |
| End Time                 | `endTime`                          | time input                  |
| Duration                 | `duration`                         | text input                  |
| Cancellation (Instant)   | `cancellation`                     | textarea                    |
| Cancellation (Request)   | `cancellation_request`             | textarea                    |
| Cut Off (Instant)        | `release`                          | textarea                    |
| Cut Off (Request)        | `release_request`                  | textarea                    |

### Bottom toggles (outside accordion)

| Field       | formcontrolname | Type     |
| ----------- | --------------- | -------- |
| B2C Enabled | `isB2CEnabled`  | checkbox |
| B2B Enabled | `isB2BEnabled`  | checkbox |

---

## Field filling techniques

### Native inputs & textareas

Angular intercepts `el.value = x`. Must use property descriptor:

```js
const setter = Object.getOwnPropertyDescriptor(
  HTMLInputElement.prototype, "value"
).set;
setter.call(el, value);
el.dispatchEvent(new Event("input", { bubbles: true }));
el.dispatchEvent(new Event("change", { bubbles: true }));
```

Same pattern for textareas using `HTMLTextAreaElement.prototype`.

### ng-select dropdowns

Must be filled sequentially (one at a time) — they share a single dropdown overlay. Two strategies attempted in order:

1. **Direct component API**: `ng.getComponent(ngSelectEl).select(item)` — fast, no DOM clicks
2. **DOM fallback**: `container.click()` → type into search input → wait → click matching `.ng-option`

Cascading dropdowns (`serviceType → activityType → subType`, `country → city`) require 900ms sleeps between each to allow Angular to re-render dependent options.

Multi-select: loop and call `comp.select()` for each value.

### Quill rich text editor

```js
editor.focus();
editor.innerHTML = `<p>${value.replace(/\n/g, "</p><p>")}</p>`;
editor.dispatchEvent(new Event("input", { bubbles: true }));
editor.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
```

### ngbdatepicker

Opens the calendar popup, navigates to the target year/month via dropdowns, clicks the matching day cell. Currently implemented but may need manual review — adjacent-month cells are skipped.

---

## Data flow (end-to-end)

```
Popup opens
  → LicenseManager.checkLicense()
  ├─ Valid → initApp()
  └─ Invalid → show panelLicense gate

initApp()
  → load config from chrome.storage.local
  ├─ Configured → show panelList
  └─ Not configured → show panelSetup

panelList
  → loadTours(forceFresh?)
  ├─ Hit cache (tourExtCache) if available
  └─ Fresh: OAuth token → Sheets API (values + gridData + metadata) → build tour objects → cache

User selects tour → "Fill Selected Tour"
  → goToFillPanel(tour)
  → show panelFill with sheet data preview
  → if tour.docUrl:
      → Docs API fetch → parseParisDoc() → fuzzy title match → merge doc fields into preview
  → display color-coded field checklist (sheet / doc / default)

User ensures all accordion panels are open → "Start Autofill"
  → startFill()
  → merge: defaults ← sheet ← doc = fillData
  → chrome.scripting.executeScript({ world: "MAIN", func: injectTourData, args: [fillData] })
  → receive { success, filled[], failed[], errors{} }
  → update checklist: green / red per field
```

---

## Virtual list rendering

- Row height: **64px** (fixed, required for position calculations)
- Container height: 360px visible area
- Overscan: 4 rows above and below viewport
- Only visible + overscan rows are in the DOM — smooth performance with 100+ tours

---

## OAuth token management

- `getOAuthToken()` called lazily per API request
- If response is 401 or 403 → `chrome.identity.removeCachedAuthToken()` → re-auth
- Prevents stale or wrong-scope tokens silently failing

---

## Column letter → index conversion

```
colLetterToIndex("A")  → 0
colLetterToIndex("Z")  → 25
colLetterToIndex("AA") → 26
colLetterToIndex("AK") → 36
```

Formula: iterate chars, `index = index * 26 + (charCode - 64)`, then subtract 1 for 0-based.

---

## Known limitations

1. **Date fields** — ngbdatepicker fill is best-effort; adjacent-month dates are skipped. Recommend manual review of start/end dates after autofill.
2. **Doc parser tuned to Paris/Switzerland doc format** — other doc structures may not parse correctly without updating section header patterns in `parisDocParser.js`.
3. **ng-select multi-selects** (guide languages, tags) — currently set to single value from sheet/doc. Looping for comma-separated multi-values is possible but not yet implemented.
4. **Admin URL check** — extension only injects into tabs at `/admin/product/add`. Will silently fail or warn on other paths.
5. **Images section** — intentionally skipped, manual only.
