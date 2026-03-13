# Tour Admin Autofill — Chrome Extension Context

## What this is

A Chrome extension that reads tour data from a Google Sheet and autofills an Angular admin panel form. Staff currently copy-paste data manually — this automates it. The human reviews the filled form before submitting.

## Current status

- Extension loads and displays tours from Google Sheets ✓
- Tour selection and navigation working ✓
- Autofill logic written but not fully tested against live admin panel
- Google Docs fetching implemented but NOT wired up yet (awaiting parsing logic)
- **Blocked**: Google Sheets fetch returning `FAILED_PRECONDITION: This operation is not supported for this document` — suspected cause is the spreadsheet being an Excel file (.xlsx) in Drive rather than a native Google Sheet. Fix: File → Save as Google Sheets.

---

## File structure

```
tour-extension/
├── manifest.json         # MV3, permissions: activeTab, scripting, storage, identity
├── popup.html            # Extension UI (380px wide popup)
├── src/
│   ├── popup.js          # All main logic — config, sheet fetch, tour list, autofill
│   ├── background.js     # Minimal service worker
│   └── content.js        # Minimal content script (not actively used)
├── icons/                # icon16/48/128.png
└── CONTEXT.md            # This file
```

---

## Admin panel tech stack

- Angular application
- Angular Material accordion (`mat-accordion`) with 3 panels — only one open at a time
- `ng-select` custom dropdown components
- Quill rich text editor for description field
- `ngbdatepicker` for date fields (readonly inputs — NOT handled yet, manual)
- `formcontrolname` attributes used as stable selectors
- Images section is intentionally NOT touched — manual only

---

## Form field → formcontrolname mapping

### Section 1 — Product Description

| Field                  | formcontrolname                 | Type                                              |
| ---------------------- | ------------------------------- | ------------------------------------------------- |
| Tour Title             | `tourTitle`                     | text input                                        |
| Service Type           | `serviceType`                   | ng-select single                                  |
| Tour Type              | `tourType`                      | ng-select single                                  |
| Activity Type          | `activityType`                  | ng-select single                                  |
| Sub Type               | `subType`                       | ng-select single                                  |
| Description            | Quill editor `.ql-editor`       | rich text                                         |
| You Will See           | `descriptionWillSee`            | textarea                                          |
| You Will Learn         | `descriptionLearn`              | textarea                                          |
| Tour Language          | `tourLanguageId`                | ng-select DISABLED (skip)                         |
| Guide Language Instant | `tourGuideLanguageList`         | ng-select multi                                   |
| Guide Language Request | `tourGuideLanguageList_request` | ng-select multi                                   |
| Mandatory Info         | `mandatoryInformation`          | text input                                        |
| Recommended Info       | `recommendedInformation`        | text input                                        |
| Included               | `included`                      | textarea                                          |
| Not Included           | `notIncluded`                   | textarea                                          |
| Activity For           | `activityFor`                   | ng-select single                                  |
| Voucher Type           | `voucherType`                   | ng-select single                                  |
| No of Pax              | `noOfPax`                       | number input                                      |
| Country                | `countryId`                     | ng-select searchable (inside formarrayname="geo") |
| City                   | `cityId`                        | ng-select searchable (inside formarrayname="geo") |
| Longitude              | `longitude`                     | text input                                        |
| Latitude               | `latitude`                      | text input                                        |
| Meeting Point          | `meetingPoint`                  | text input                                        |
| Pickup Instructions    | `pickupInstructions`            | text input                                        |
| End Point              | `endPoint`                      | text input                                        |
| Tags                   | `tagsList`                      | ng-select multi                                   |

### Section 2 — Product Images

- **SKIPPED — manual only**

### Section 3 — Price & Schedule (inside formgroupname="price")

| Field                    | formcontrolname                    | Type                          |
| ------------------------ | ---------------------------------- | ----------------------------- |
| Price Model              | `isFixedModel`                     | ng-select                     |
| Currency                 | `currency`                         | ng-select                     |
| Rate (Instant B2B)       | `rate`                             | number input                  |
| Rate (Instant B2C)       | `rateB2C`                          | number input                  |
| Rate (Request B2B)       | `rate_request`                     | number input                  |
| Rate (Request B2C)       | `rate_requestB2C`                  | number input                  |
| Extra Hour (Instant)     | `extraHourCharges`                 | number input                  |
| Extra Hour B2C (Instant) | `extraHourChargesB2C`              | number input                  |
| Extra Hour (Request)     | `extraHourCharges_request`         | number input                  |
| Extra Hour B2C (Request) | `extraHourCharges_requestB2C`      | number input                  |
| Holiday Supplement %     | `publicHolidaySurchargePercentage` | number input                  |
| Weekend Supplement %     | `weekendSupplementPercentage`      | number input                  |
| Start Date               | `startDate`                        | ngbdatepicker (SKIP — manual) |
| End Date                 | `endDate`                          | ngbdatepicker (SKIP — manual) |
| Start Time               | `startTime`                        | time input                    |
| End Time                 | `endTime`                          | time input                    |
| Duration                 | `duration`                         | text input                    |
| Cancellation (Instant)   | `cancellation`                     | textarea                      |
| Cancellation (Request)   | `cancellation_request`             | textarea                      |
| Cut Off (Instant)        | `release`                          | textarea                      |
| Cut Off (Request)        | `release_request`                  | textarea                      |

### Bottom toggles (outside accordion)

| Field       | formcontrolname | Type     |
| ----------- | --------------- | -------- |
| B2C Enabled | `isB2CEnabled`  | checkbox |
| B2B Enabled | `isB2BEnabled`  | checkbox |

---

## Angular input filling technique

Angular intercepts native value changes, so direct `el.value = x` doesn't work. Must use:

```js
const setter = Object.getOwnPropertyDescriptor(
  window.HTMLInputElement.prototype,
  "value",
).set;
setter.call(el, value);
el.dispatchEvent(new Event("input", { bubbles: true }));
el.dispatchEvent(new Event("change", { bubbles: true }));
el.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
```

Same pattern for textarea using `HTMLTextAreaElement.prototype`.

## ng-select filling technique

Must be sequential (one at a time) because they share a single dropdown overlay:

```js
container.click(); // open dropdown
await sleep(250);
setNativeInput(searchInput, value); // type to filter
await sleep(350);
// find matching .ng-option and click it
```

## Quill editor filling technique

```js
editor.focus();
editor.innerHTML = `<p>${value}</p>`;
editor.dispatchEvent(new Event("input", { bubbles: true }));
editor.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
```

---

## Data sources

### Google Sheets (via API key, public sheet)

Fetches columns by explicitly configured column letters (A, B, C, etc.) — NOT by header name auto-detection. User configures each column letter in the settings panel.

Current column config fields in settings UI:

- Core: Tour Title, Doc URL, Country, City, Duration, Product Type
- Pricing: B2B Instant, B2B On Request, B2C Instant, B2C On Request
- Cancellation & Cut Off: Cancel Instant, Cancel On Request, Cut Off Instant, Cut Off On Request

Sheet fetch URL pattern:

```
https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{ENCODED_RANGE}?key={API_KEY}
```

Range encoding: tab name spaces → `%20`, single quotes wrap the tab name but must NOT be `%27` encoded — this is the current bug.

### Google Docs (via chrome.identity OAuth — implemented, not wired up)

`fetchGoogleDoc(docUrl)` exists in popup.js. It:

1. Gets OAuth token via `chrome.identity.getAuthToken({ interactive: true })`
2. Calls `https://docs.googleapis.com/v1/documents/{DOC_ID}` with Bearer token
3. Returns raw document JSON — no parsing done yet
4. Caller is responsible for extracting field values from the doc

The `manifest.json` has an `oauth2` block with `documents.readonly` scope. The `client_id` field is a placeholder (`YOUR_OAUTH_CLIENT_ID`) — needs to be filled in with a real OAuth 2.0 Client ID from Google Cloud Console (type: Chrome Extension, requires extension ID).

---

## Data flow

1. Popup opens → loads config from `chrome.storage.local`
2. If configured → loads tours from cache or fetches fresh from Sheets
3. User searches/selects a tour → goToFillPanel() shows preview with sheet data
4. User clicks "Open All Sections" → clicks all mat-expansion-panel-headers with 600ms stagger
5. User clicks "Start Autofill" → `chrome.scripting.executeScript` injects `injectTourData(fillData)` into active tab
6. `injectTourData` runs in page context, fills all fields, returns `{ success, filled[], failed[] }`
7. Checklist updates green/red per field

## Fill data merge logic

Sheet data takes priority over dummy data:

```js
const fillData = {
  ...DUMMY, // fallback for everything
  title: selectedTour.title, // always from sheet
  country: selectedTour.country || DUMMY.country,
  // ... etc for all sheet-sourced fields
};
```

## What still uses dummy data

Everything not yet in the sheet: description (Quill), tour type, activity type, sub type, you will see/learn, mandatory/recommended info, included/not included, activity for, voucher type, pax, guide languages, coordinates, meeting point, pickup instructions, end point, tags, price model, currency, extra hour charges, supplements, start/end time, B2C/B2B toggles.

These will eventually come from the Google Doc once parsing logic is added.

---

## Settings panel config keys (chrome.storage.local → tourExtConfig)

```js
{
  (sheetId,
    sheetTab,
    apiKey,
    colTitle,
    colDocUrl,
    colCountry,
    colCity,
    colDuration,
    colServiceType,
    colRate,
    colRateRequest,
    colRateB2C,
    colRateRequestB2C,
    colCancellation,
    colCancellationRequest,
    colRelease,
    colReleaseRequest);
}
```

## Cache key

`tourExtCache` in `chrome.storage.local` — array of tour objects. Cleared/replaced on refresh button click.

---

## Known issues / TODO

1. **Sheets fetch broken** — `FAILED_PRECONDITION` error. Most likely the spreadsheet is an .xlsx file not converted to native Google Sheets format. Fix: File → Save as Google Sheets in the browser.
2. **Date fields** (`startDate`, `endDate`, `closedDay`, `holiday`) use `ngbdatepicker` with readonly inputs — need calendar-click simulation or NgControl patching. Currently skipped.
3. **Google Doc parsing** — `fetchGoogleDoc()` is ready but not called anywhere. Need to define what fields live in the doc and write parsing logic, then wire it into `goToFillPanel()` or `startFill()`.
4. **OAuth client ID** — placeholder in manifest.json needs replacing with real Chrome Extension OAuth client ID from Google Cloud Console.
5. **Multi-value ng-selects** — guide languages and tags currently only set one value. Need to call fillNgSelect once per comma-separated value to add multiple selections.
