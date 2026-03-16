/**
 * parisDocParser.js
 * Google Docs API parser for Tour Content Documents (Paris / France / Switzerland)
 * Chrome Extension compatible — no dependencies
 *
 * Extracted fields per tour:
 *   title, description, duration, meetingPoint, endLocation,
 *   inclusions, exclusions, additionalInfo, youWillLearn, youWillSee
 */


// ---------------------------------------------------------------------------
// Element-level helpers
// ---------------------------------------------------------------------------

function paragraphElements(paragraph) {
  return (paragraph.elements ?? []).map(el => {
    const tr = el.textRun ?? {};
    return {
      bold: !!tr.textStyle?.bold,
      // Do NOT strip \n here — preserve soft line breaks between runs.
      // paragraphText() strips only the final trailing \n.
      text: tr.content ?? "",
    };
  });
}

/**
 * Detect a plain-text coordinate pair like "47.049700, 8.309496".
 * Returns { latitude, longitude } strings or null.
 */
function coordsFromText(text) {
  const m = text.trim().match(/^(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)$/);
  if (!m) return null;
  return { longitude: m[1], latitude: m[2] };
}

function paragraphText(paragraph) {
  return paragraphElements(paragraph).map(e => e.text).join("").replace(/\n$/, "");
}

function paragraphStyle(paragraph) {
  return paragraph.paragraphStyle?.namedStyleType ?? "NORMAL_TEXT";
}

function paragraphHasBold(paragraph) {
  return paragraphElements(paragraph).some(e => e.bold && e.text.trim());
}

// ---------------------------------------------------------------------------
// Structural helpers
// ---------------------------------------------------------------------------

/**
 * Separator lines — dashes OR underscores.
 */
function isSeparator(text) {
  const t = text.trim();
  return t.length > 10 && t.replace(/[-–_\s\t]/g, "").length === 0;
}

function splitSoftBreaks(text) {
  return text.split("\n").map(l => l.replace(/\u00a0/g, " ").trim()).filter(Boolean);
}

/**
 * Pure editorial noise — never belongs in any content field.
 * Covers: "Price is missing", "Sources:", "Changed to …", "BOKUN",
 * "VIATOR / EXPEDIA / KLOOK …", "SS note: …", platform distribution lists.
 */
function isNoiseLine(text) {
  const t = text.replace(/\u00a0/g, " ").trim();
  return (
    /^price\s+is\s+missing:?\s*$/i.test(t)          ||
    /^sources?\s*:/i.test(t)                           ||
    /^notes?\s*:/i.test(t)                            ||
    /^changed\s+to\s+.+$/i.test(t)                   ||
    /^BOKUN\s*$/i.test(t)                             ||
    /^(VIATOR|EXPEDIA|KLOOK)/i.test(t)               ||
    /^SS\s+note\s*:/i.test(t)
  );
}

/**
 * Review / approval flags that can appear in ANY paragraph style — headings
 * or plain normal paragraphs, bold or not.
 * Covers all known variants across Paris, France, and Switzerland docs.
 */
function isReviewFlagText(text) {
  const t = text.trim();
  return (
    /^(CC\s+OK|IS\s+OK)/i.test(t)                             ||
    /^(IS\s+OK\s*&\s*CC\s+OK|CC\s+OK\s*&\s*IS\s+OK)/i.test(t) ||
    /^SS\s*[:\s]/i.test(t)                                     ||
    /^SS\s+(OK|CSV|re-reviewed|reviewed|REVIEWD?)/i.test(t)    ||
    /^RR\s+(revised|ok)/i.test(t)                              ||
    /^On\s+Hold/i.test(t)
  );
}

// ---------------------------------------------------------------------------
// Patterns
// ---------------------------------------------------------------------------

// HEADING_1 texts that are review flags, not tour titles
const SKIP_TITLE_RE = /^(CC\s+OK|IS\s+OK|SS\s+|RR\s+(revised|ok)|On\s+Hold|\s*)$/i;

// Styles we treat as "heading-like" for section label detection
// Includes "TITLE" which is used sporadically in Switzerland doc
const HEADING_LIKE_STYLES = new Set([
  "HEADING_1","HEADING_2","HEADING_3","HEADING_4","HEADING_5","HEADING_6","TITLE"
]);

// Review flag headings (any heading style)
const REVIEW_FLAG_HEADING_STYLES = new Set([
  "HEADING_1","HEADING_2","HEADING_3","HEADING_4","HEADING_5","HEADING_6"
]);

/**
 * List section headers → field names.
 * "Highlights" is used in Switzerland day-trip tours instead of "You will see".
 */
const LIST_SECTION_PATTERNS = [
  [/^inclusions?\s*:?\s*$/i,      "inclusions"],
  [/^exclusions?\s*:?\s*$/i,      "exclusions"],
  [/^additional\s+info/i,         "additionalInfo"],
  [/^you\s+will\s+learn/i,        "youWillLearn"],
  [/^you\s+will\s+see/i,          "youWillSee"],
  [/^highlights?\s*:?\s*$/i,      "youWillSee"],   // Switzerland day-trip variant
];

const INLINE_FIELD_PATTERNS = [
  [/^meeting\s+point\s*:/i,   "meetingPoint"],
  [/^duration\s*:/i,           "duration"],
  [/^end\s+loc(ation)?\s*:/i,  "endLocation"],
  [/^longitude\s*:/i,          "longitude"],
  [/^latitude\s*:/i,           "latitude"],
];

// ---------------------------------------------------------------------------
// Multi-field paragraph parser
// Handles paragraphs where fields are packed with soft line breaks, e.g.:
//   [bold] "Meeting point:" [plain] " Terminal 2E…\n "
//   [bold] "Duration:"      [plain] " 5 hours\n "
// ---------------------------------------------------------------------------

function parseMultiFieldParagraph(paragraph) {
  const els = paragraphElements(paragraph);
  if (els.length < 2) return null;

  const hasBoldLabel = els.some(e =>
    e.bold && INLINE_FIELD_PATTERNS.some(([re]) => re.test(e.text.trim()))
  );
  if (!hasBoldLabel) return null;

  const results = [];
  let currentField = null;
  let valueChunks  = [];

  const flush = () => {
    if (currentField && valueChunks.length) {
      const value = valueChunks.join("").trim().split("\n")[0].trim();
      results.push({ field: currentField, value });
    }
  };

  for (const el of els) {
    const trimmed = el.text.trim();
    if (!trimmed) continue;

    if (el.bold) {
      const match = INLINE_FIELD_PATTERNS.find(([re]) => re.test(trimmed));
      if (match) {
        flush();
        currentField = match[1];
        valueChunks  = [];
        const afterColon = trimmed.replace(/^[^:]+:\s*/, "");
        if (afterColon) valueChunks.push(afterColon);
        continue;
      }
    }

    if (currentField) valueChunks.push(el.text);
  }
  flush();

  return results.length ? results : null;
}

// ---------------------------------------------------------------------------
// Body paragraph extractor
// ---------------------------------------------------------------------------

function getBodyParagraphs(docJson) {
  const paras = [];
  for (const el of docJson.body?.content ?? []) {
    if (el.paragraph) {
      paras.push(el.paragraph);
    } else if (el.table) {
      for (const row of el.table.tableRows ?? []) {
        for (const cell of row.tableCells ?? []) {
          for (const ce of cell.content ?? []) {
            if (ce.paragraph) paras.push(ce.paragraph);
          }
        }
      }
    }
  }
  return paras;
}

// ---------------------------------------------------------------------------
// Tour detection + section splitting
// ---------------------------------------------------------------------------

function isTourTitle(paragraph) {
  if (paragraphStyle(paragraph) !== "HEADING_1") return false;
  const text = paragraphText(paragraph).trim();
  return text.length > 0 && !SKIP_TITLE_RE.test(text) && !isReviewFlagText(text);
}

function isReviewFlagParagraph(paragraph) {
  const text  = paragraphText(paragraph).trim();
  if (!text) return false;
  const style = paragraphStyle(paragraph);

  // Heading-style review flags
  if (REVIEW_FLAG_HEADING_STYLES.has(style)) return isReviewFlagText(text);

  // Normal paragraphs that are review flags (bold or not)
  if (style === "NORMAL_TEXT") return isReviewFlagText(text);

  return false;
}

function splitIntoSections(paragraphs) {
  const sections = [];
  let current    = [];
  for (const para of paragraphs) {
    if (isTourTitle(para)) {
      if (current.length) sections.push(current);
      current = [para];
    } else if (current.length) {
      current.push(para);
    }
  }
  if (current.length) sections.push(current);
  return sections;
}

// ---------------------------------------------------------------------------
// Section parser
// ---------------------------------------------------------------------------

/**
 * @typedef {Object} TourSection
 * @property {string}   title
 * @property {string}   description
 * @property {string}   duration
 * @property {string}   meetingPoint
 * @property {string}   endLocation
 * @property {string}   longitude
 * @property {string}   latitude
 * @property {string[]} inclusions
 * @property {string[]} exclusions
 * @property {string[]} additionalInfo
 * @property {string[]} youWillLearn
 * @property {string[]} youWillSee
 */

function parseSection(paragraphs) {
  /** @type {TourSection} */
  const tour = {
    title: "", description: "",
    duration: "", meetingPoint: "", endLocation: "",
    longitude: "", latitude: "",
    inclusions: [], exclusions: [], additionalInfo: [],
    youWillLearn: [], youWillSee: [],
  };

  if (!paragraphs.length) return tour;

  tour.title = paragraphText(paragraphs[0]).trim();

  const descLines = [];
  let listField   = null;
  let inDesc      = true;

  const addToList = (fieldName, text) => {
    for (const line of splitSoftBreaks(text)) {
      // Skip long walk-narrative lines (> 300 chars) from list fields —
      // these are itinerary prose that slipped into the wrong section
      if (line && !isNoiseLine(line) && line.length <= 300) {
        tour[fieldName].push(line);
      }
    }
  };

  for (const para of paragraphs.slice(1)) {
    const text    = paragraphText(para).trim();
    const style   = paragraphStyle(para);
    const hasBold = paragraphHasBold(para);
    const isHeadingLike = HEADING_LIKE_STYLES.has(style);

    // Skip empty lines inside list sections
    if (!text && listField) continue;

    // Review flags — skip regardless of style
    if (isReviewFlagParagraph(para)) { listField = null; inDesc = true; continue; }

    // Noise lines — skip everywhere
    if (isNoiseLine(text)) continue;

    // Separator lines (dashes or underscores) — end of content
    if (isSeparator(text)) { listField = null; inDesc = false; continue; }

    // Plain-text coordinate pair e.g. "47.049700, 8.309496"
    if (!tour.latitude && !tour.longitude) {
      const coords = coordsFromText(text);
      if (coords) {
        tour.latitude  = coords.latitude;
        tour.longitude = coords.longitude;
        continue;
      }
    }

    // Multi-field paragraph (Meeting point / Duration / End location packed together)
    const multiFields = parseMultiFieldParagraph(para);
    if (multiFields) {
      for (const { field, value } of multiFields) {
        if (value) tour[field] = value;
      }
      inDesc = false; listField = null;
      continue;
    }

    // Bold / heading-like labels (includes TITLE style)
    // Match against the FIRST LINE only — some paragraphs pack label + items
    // with embedded \n, e.g. "Highlights:\nAbbey of St. Gallen"
    const firstLine = text.split("\n")[0].trim();
    const restLines = text.split("\n").slice(1).join("\n").trim();

    if (hasBold || isHeadingLike) {
      // List section headers
      const listMatch = LIST_SECTION_PATTERNS.find(([re]) => re.test(firstLine));
      if (listMatch) {
        listField = listMatch[1];
        inDesc    = false;
        // Add any items packed after the label on the same paragraph
        if (restLines) addToList(listField, restLines);
        continue;
      }

      // Inline field labels
      const inlineMatch = INLINE_FIELD_PATTERNS.find(([re]) => re.test(firstLine));
      if (inlineMatch) {
        const [pattern, fieldName] = inlineMatch;
        tour[fieldName] = firstLine.replace(pattern, "").replace(/^[\s:]+/, "").trim()
                       || firstLine.split(/:\s*/)[1]?.trim()
                       || "";
        inDesc = false; listField = null;
        continue;
      }
    }

    // Accumulate into active list field
    if (listField) { if (text && !isNoiseLine(text)) addToList(listField, text); continue; }

    // Description — skip review flag text and noise even as plain paragraphs
    if (inDesc) {
      if (text && !isReviewFlagText(text) && !isNoiseLine(text)) {
        descLines.push(text);
      }
      continue;
    }
  }

  tour.description = descLines.join("\n\n");
  return tour;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse a Google Docs API response into an array of TourSection objects.
 * @param {Object} docJson — response from GET /v1/documents/{documentId}
 * @returns {TourSection[]}
 */
function parseParisDoc(docJson) {
  return splitIntoSections(getBodyParagraphs(docJson)).map(parseSection);
}
