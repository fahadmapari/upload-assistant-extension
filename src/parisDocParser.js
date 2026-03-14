/**
 * parisDocParser.js
 * Google Docs API parser for Paris/France Tour Content Documents
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
      text: (tr.content ?? "").replace(/\n$/, ""),
    };
  });
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
 * Separator lines between tour sections — dashes OR underscores.
 */
function isSeparator(text) {
  const t = text.trim();
  return t.length > 10 && t.replace(/[-–_\s\t]/g, "").length === 0;
}

function splitSoftBreaks(text) {
  return text.split("\n").map(l => l.trim()).filter(Boolean);
}

/**
 * Lines that are editorial noise and should never appear in any field.
 * Covers: "Price is missing", "Price is missing:", "Sources:", "Source:"
 * and any review flag text that slipped through as a plain paragraph.
 */
function isNoiseLine(text) {
  return /^(price\s+is\s+missing:?\s*$|sources?\s*:?\s*$|changed\s+to\s+.+$)/i.test(text.trim());
}

/**
 * Review flag text that can appear as ANY paragraph style (heading or normal).
 */
const REVIEW_FLAG_RE = /^(CC\s+OK|SS\s+(OK.*|CSV.*|re-reviewed.*|reviewed.*)|RR\s+(revised|ok).*|On\s+Hold)/i;

function isReviewFlagText(text) {
  return REVIEW_FLAG_RE.test(text.trim());
}

// ---------------------------------------------------------------------------
// Patterns
// ---------------------------------------------------------------------------

// HEADING_1 titles that are actually review flags, not tour titles
const SKIP_TITLE_RE = /^(CC\s+OK\s*$|SS\s+(CSV.*|re-reviewed.*|reviewed.*)|RR\s+(revised|ok).*|On\s+Hold|\s*)$/i;

const REVIEW_STYLES = new Set(["HEADING_1","HEADING_2","HEADING_3","HEADING_4","HEADING_5","HEADING_6"]);

const LIST_SECTION_PATTERNS = [
  [/^inclusions?\s*:?\s*$/i,  "inclusions"],
  [/^exclusions?\s*:?\s*$/i,  "exclusions"],
  [/^additional\s+info/i,     "additionalInfo"],
  [/^you\s+will\s+learn/i,    "youWillLearn"],
  [/^you\s+will\s+see/i,      "youWillSee"],
];

const INLINE_FIELD_PATTERNS = [
  [/^meeting\s+point\s*:/i,   "meetingPoint"],
  [/^duration\s*:/i,           "duration"],
  [/^end\s+loc(ation)?\s*:/i,  "endLocation"],
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
  return text.length > 0 && !SKIP_TITLE_RE.test(text);
}

function isReviewFlagParagraph(paragraph) {
  const text = paragraphText(paragraph).trim();
  if (!text) return false;
  // Review flags can appear as any heading style OR as plain normal paragraphs
  const style = paragraphStyle(paragraph);
  const isHeading = REVIEW_STYLES.has(style);
  const isNormal  = style === "NORMAL_TEXT";
  return (isHeading || isNormal) && isReviewFlagText(text);
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

function parseSection(paragraphs) {
  const tour = {
    title: "", description: "",
    duration: "", meetingPoint: "", endLocation: "",
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
      if (line && !isNoiseLine(line)) tour[fieldName].push(line);
    }
  };

  for (const para of paragraphs.slice(1)) {
    const text    = paragraphText(para).trim();
    const style   = paragraphStyle(para);
    const hasBold = paragraphHasBold(para);

    // Always skip empty lines (when inside a list field)
    if (!text && listField) continue;

    // Review flags — skip regardless of paragraph style
    if (isReviewFlagParagraph(para)) { listField = null; inDesc = true; continue; }

    // Noise lines — skip everywhere (price, sources label, "Changed to X")
    if (isNoiseLine(text)) continue;

    // Separator lines (dashes or underscores) — mark end of content
    if (isSeparator(text)) { listField = null; inDesc = false; continue; }

    // Multi-field paragraph (Meeting point / Duration / End location packed together)
    const multiFields = parseMultiFieldParagraph(para);
    if (multiFields) {
      for (const { field, value } of multiFields) {
        if (value) tour[field] = value;
      }
      inDesc = false; listField = null;
      continue;
    }

    // Bold / heading labels
    if (hasBold || style.startsWith("HEADING_")) {
      // List section headers
      const listMatch = LIST_SECTION_PATTERNS.find(([re]) => re.test(text));
      if (listMatch) {
        listField = listMatch[1];
        inDesc    = false;
        continue;
      }

      // Inline field labels
      const inlineMatch = INLINE_FIELD_PATTERNS.find(([re]) => re.test(text));
      if (inlineMatch) {
        const [pattern, fieldName] = inlineMatch;
        tour[fieldName] = text.replace(pattern, "").replace(/^[\s:]+/, "").trim()
                       || text.split(/:\s*/)[1]?.trim()
                       || "";
        inDesc = false; listField = null;
        continue;
      }
    }

    // Accumulate into active list field
    if (listField) { if (text) addToList(listField, text); continue; }

    // Description — but skip review flag text even if not caught by heading check
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
