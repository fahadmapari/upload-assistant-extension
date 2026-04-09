/**
 * plainTextParser.js
 * Parser for plain-text tour content pasted directly into the extension.
 * Produces the same TourSection shape as parisDocParser.js.
 *
 * Expected input format (from Daytrip / operator plain-text docs):
 *   Line 1:   Tour title
 *   Lines 2+: Optional review flags (IS OK, SS reviewed, SS note:, …) — skipped
 *   Body:     Description paragraphs (blank-line separated)
 *   Fields:   Meeting point: …  /  Duration: …  /  End location: …
 *   Coords:   "47.049700, 8.309496"
 *   Sections: Inclusions / Exclusions / Additional info / Highlights
 *   Noise:    Source: …  /  Notes: …  /  price lines  /  Voucher: …
 */

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function ptCoords(text) {
  const m = text.trim().match(/^(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)$/);
  if (!m) return null;
  return { longitude: m[1], latitude: m[2] };
}

function ptIsSeparator(text) {
  const t = text.trim();
  return t.length > 10 && t.replace(/[-–_\s\t]/g, "").length === 0;
}

/**
 * Pure editorial / metadata noise — never belongs in any content field.
 * Mirrors parisDocParser's isNoiseLine() plus plain-text-specific patterns.
 */
function ptIsNoise(text) {
  const t = text.replace(/\u00a0/g, " ").trim();
  return (
    /^sources?\s*:/i.test(t)                         ||
    /^notes?\s*:/i.test(t)                           ||
    /^changed\s+to\s+.+$/i.test(t)                  ||
    /^BOKUN\s*$/i.test(t)                            ||
    /^(VIATOR|EXPEDIA|KLOOK)/i.test(t)               ||
    /^SS\s+note\s*:/i.test(t)                        ||
    /https?:\/\//i.test(t)                           ||
    /^www\./i.test(t)                                ||
    /^cancel\s+up\s+to\s+.+refund/i.test(t)         ||
    /^price\s+is\s+missing/i.test(t)                 ||
    // Plain-text specific
    /^voucher\s*:/i.test(t)                          ||
    /^maximum\s+\w+\s+participants/i.test(t)         ||
    /^[€$£]\d/.test(t)                               ||   // "€681 total…"
    /^≈\s*[€$£]/.test(t)                                  // "≈ €189 per person"
  );
}

/**
 * Review / approval flags — same patterns as parisDocParser's isReviewFlagText().
 */
function ptIsReviewFlag(text) {
  const t = text.trim();
  return (
    /^(CC\s+OK|IS\s+OK)/i.test(t)                              ||
    /^(IS\s+OK\s*&\s*CC\s+OK|CC\s+OK\s*&\s*IS\s+OK)/i.test(t) ||
    /^SS\s*[:\s]/i.test(t)                                      ||
    /^SS\s+(OK|CSV|re-reviewed|reviewed|REVIEWD?)/i.test(t)     ||
    /^RR\s+(revised|ok)/i.test(t)                               ||
    /^On\s+Hold/i.test(t)
  );
}

// ---------------------------------------------------------------------------
// Section + field patterns (mirrors parisDocParser)
// ---------------------------------------------------------------------------

const PT_LIST_SECTION_PATTERNS = [
  [/^inclusions?\s*:?\s*$/i,                    "inclusions"],
  [/^exclusions?\s*:?\s*$/i,                    "exclusions"],
  [/^additional\s+info/i,                       "additionalInfo"],
  [/^you\s+will\s+learn(\s+about)?\s*:?\s*$/i, "youWillLearn"],  // "You will learn about:" variant
  [/^you\s+will\s+see\s*:?\s*$/i,              "youWillSee"],
  [/^highlights?\s*:?\s*$/i,                    "youWillSee"],    // Switzerland day-trip variant
];

const PT_INLINE_FIELD_PATTERNS = [
  [/^meeting\s+point\s*:/i,   "meetingPoint"],
  [/^duration\s*:/i,           "duration"],
  [/^end\s+loc(ation)?\s*:/i,  "endLocation"],
  [/^longitude\s*:/i,          "longitude"],
  [/^latitude\s*:/i,           "latitude"],
];

// ---------------------------------------------------------------------------
// Main parser
// ---------------------------------------------------------------------------

/**
 * Parse a plain-text tour pasted by the user into a TourSection object.
 * @param {string} rawText
 * @returns {Object} TourSection — same shape as parisDocParser output
 */
function parsePlainTextTour(rawText) {
  const tour = {
    title: "", description: "",
    duration: "", meetingPoint: "", endLocation: "",
    longitude: "", latitude: "",
    inclusions: [], exclusions: [], additionalInfo: [],
    youWillLearn: [], youWillSee: [],
  };

  const lines = rawText.split("\n").map(l => l.replace(/\u00a0/g, " "));

  // ── Locate title: first non-empty, non-flag, non-noise line ──
  let titleIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    if (t && !ptIsReviewFlag(t) && !ptIsNoise(t) && !ptIsSeparator(t)) {
      tour.title = t;
      titleIdx = i;
      break;
    }
  }
  if (titleIdx < 0) return tour;

  // ── Parse body ──
  const descLines = [];
  let listField          = null;
  let inDesc             = true;
  let pendingInlineField = null;

  const addToList = (fieldName, line) => {
    // Strip leading bullet characters
    const clean = line.replace(/^\s*[•\-–]\s*/, "").replace(/\u00a0/g, " ").trim();
    // Mirror parisDocParser: skip noise and overly long prose lines
    if (clean && !ptIsNoise(clean) && clean.length <= 300) {
      tour[fieldName].push(clean);
    }
  };

  for (let i = titleIdx + 1; i < lines.length; i++) {
    const text = lines[i].trim();

    // A blank line ends the active list section (e.g. blank line after "You will see")
    if (!text && listField) { listField = null; continue; }
    // Skip blank lines while waiting for an inline field value
    if (!text && pendingInlineField) continue;

    // Capture the value for a pending inline field on the next non-empty line
    if (pendingInlineField && text) {
      if (!ptIsNoise(text) && !ptIsReviewFlag(text) && !ptIsSeparator(text)) {
        tour[pendingInlineField] = text.split("\n")[0].trim();
      }
      pendingInlineField = null;
      continue;
    }

    // Review / approval flags — skip, reset list context
    if (ptIsReviewFlag(text)) { listField = null; inDesc = true; continue; }

    // Noise — always skip; also ends any active list section
    if (ptIsNoise(text)) { listField = null; continue; }

    // Separator lines — end of meaningful content
    if (ptIsSeparator(text)) { listField = null; inDesc = false; continue; }

    // Plain-text coordinate pair e.g. "47.049700, 8.309496"
    if (!tour.latitude && !tour.longitude) {
      const coords = ptCoords(text);
      if (coords) {
        tour.latitude  = coords.latitude;
        tour.longitude = coords.longitude;
        continue;
      }
    }

    // List section headers (Inclusions, Exclusions, Highlights, etc.)
    const listMatch = PT_LIST_SECTION_PATTERNS.find(([re]) => re.test(text));
    if (listMatch) {
      listField = listMatch[1];
      inDesc    = false;
      // Handle items packed after the label: "Highlights:\nItem1"
      const afterColon = text.replace(listMatch[0], "").replace(/^[\s:]+/, "").trim();
      if (afterColon) {
        for (const l of afterColon.split("\n")) addToList(listField, l);
      }
      continue;
    }

    // Inline field labels (Meeting point:, Duration:, End location:, …)
    const inlineMatch = PT_INLINE_FIELD_PATTERNS.find(([re]) => re.test(text));
    if (inlineMatch) {
      const [pattern, fieldName] = inlineMatch;
      const inlineValue = text.replace(pattern, "").replace(/^[\s:]+/, "").trim();
      inDesc    = false;
      listField = null;
      if (inlineValue) {
        tour[fieldName] = inlineValue;
      } else {
        // Value is on the next non-empty line
        pendingInlineField = fieldName;
      }
      continue;
    }

    // Accumulate list items under the active section header
    if (listField) {
      if (text) addToList(listField, text);
      continue;
    }

    // Accumulate description paragraphs (blank lines between paragraphs are fine —
    // empty lines don't reach here since inDesc is true and text is non-empty)
    if (inDesc && text) {
      descLines.push(text);
    }
  }

  tour.description = descLines.join("\n\n");
  return tour;
}
