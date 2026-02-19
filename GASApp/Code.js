/************ CONFIG ************/
const CONFIG = {
  SPREADSHEET_ID: "1hjBr5E7eDLbPzVxCrUC9CHxDt-N1TpU0ud6lMalfo8Y",
  INDEX_SHEET_NAME: "INDEX",
  REQUESTS_SHEET_NAME: "REQUESTS",

  // Drive folder that contains all enrollment slips (.html files)
  SLIPS_FOLDER_ID: "1Kxji4ck70YKNdP12KIhQYDEDpyfL5xNa",

  // File naming strategy for Drive search (matches base name without extension)
  // "NAME_ONLY" => expects base name: "<FullName>"
  // "NAME_SLIP" => expects base name: "<FullName> - Enrollment Slip"
  FILE_MATCH_MODE: "NAME_ONLY",

  // Token validity in seconds (for opening slip)
  VIEW_TOKEN_TTL_SEC: 300,

  // Suggestions limit
  SUGGEST_LIMIT: 10,

  // Attempt limiting (per user properties)
  MAX_FAILED_ATTEMPTS: 5,
  LOCKOUT_MINUTES: 10,
};

/************ ENTRY ************/
function doGet(e) {
  const token = e?.parameter?.t;
  const view = e?.parameter?.view;

  // If view=1 and token exists, render slip HTML
  if (token && view === "1") return serveSlipHtml_(token);

  // Otherwise serve the main UI
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Enrollment Slip Access")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/************ SEARCH (Sheet) ************/
function searchStudents(query) {
  query = normalize_(query);
  if (!query) return [];

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.INDEX_SHEET_NAME);
  if (!sh) throw new Error("INDEX sheet not found.");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const col = getIndexCols_(headers);

  const tokens = query.split(" ").filter(Boolean);
  const results = [];

  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][col.fullName] || "").trim();
    const status = String(values[i][col.status] || "").trim().toUpperCase();
    if (!name) continue;

    // Optional: only suggest ACTIVE students
    // if (status !== "ACTIVE") continue;

    const nameNorm = normalize_(name);
    if (matchesTokensInOrder_(nameNorm, tokens)) {
      results.push({ name });
      if (results.length >= CONFIG.SUGGEST_LIMIT) break;
    }
  }

  return results;
}

/************ VERIFY (Sheet) + ISSUE VIEW TOKEN ************/
function verifyAndIssueViewToken(selectedName, studentId) {
  selectedName = String(selectedName || "").trim();
  studentId = String(studentId || "").trim();

  if (!selectedName || !studentId) {
    return { ok: false, code: "MISSING_INPUT", message: "Name and Student ID are required." };
  }

  // Lockout check
  const lock = checkLockout_();
  if (lock.locked) {
    return {
      ok: false,
      code: "LOCKED",
      message: `Too many failed attempts. Try again in ${lock.remainingMinutes} minutes.`
    };
  }

  // Verify record (Name exact + StudentID matches either format)
  const record = findRecordByNameAndIdFlexible_(selectedName, studentId);
  if (!record) {
    registerFailedAttempt_();
    logRequest_(selectedName, studentId, "ID_MISMATCH_OR_NOT_FOUND");
    return { ok: false, code: "NO_MATCH", message: "No matching record found. Please check your Student ID." };
  }

  if (String(record.status).toUpperCase() !== "ACTIVE") {
    logRequest_(selectedName, studentId, "STATUS_NOT_ACTIVE");
    return { ok: false, code: "INACTIVE", message: "Your record is not active. Please contact admin." };
  }

  // Reset failures on success
  resetFailedAttempts_();

  // Find slip file in Drive folder (expects .html, but we match base name)
  const file = findSlipFileInFolder_(record.fullName);
  if (!file) {
    logRequest_(selectedName, studentId, "VERIFIED_BUT_FILE_MISSING");
    return {
      ok: false,
      code: "FILE_MISSING",
      message: "Verified, but slip file is missing. Submitted for admin check."
    };
  }

  // Issue token -> store fileId in cache
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(token, file.getId(), CONFIG.VIEW_TOKEN_TTL_SEC);

  const viewUrl = ScriptApp.getService().getUrl() + "?view=1&t=" + encodeURIComponent(token);
  return { ok: true, viewUrl, filename: file.getName() };
}

/************ SERVE SLIP HTML (runs HTML in browser) ************/
function serveSlipHtml_(token) {
  const fileId = CacheService.getScriptCache().get(token);

  if (!fileId) {
    return HtmlService.createHtmlOutput(
      "<div style='font-family:system-ui;padding:24px;'>Link expired. Please verify again.</div>"
    );
  }

  const file = DriveApp.getFileById(fileId);
  const htmlText = file.getBlob().getDataAsString("UTF-8");

  // Serve the HTML content directly
  return HtmlService.createHtmlOutput(htmlText)
    .setTitle("Enrollment Slip")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/************ SHEET LOOKUP ************/
function findRecordByNameAndIdFlexible_(fullName, enteredId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.INDEX_SHEET_NAME);
  if (!sh) throw new Error("INDEX sheet not found.");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0].map(h => String(h).trim());
  const col = getIndexCols_(headers);

  const nameTarget = normalizeExact_(fullName);
  const enteredVariants = normalizeStudentIdVariants_(enteredId);

  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][col.fullName] || "").trim();
    const id = String(values[i][col.studentId] || "").trim();
    const status = String(values[i][col.status] || "").trim();

    if (!name || !id) continue;

    if (normalizeExact_(name) !== nameTarget) continue;

    const recordVariants = normalizeStudentIdVariants_(id);

    const idMatch = enteredVariants.some(v => recordVariants.includes(v));
    if (idMatch) {
      return { fullName: name, studentId: id, status };
    }
  }
  return null;
}

/************ STUDENT ID FLEXIBLE NORMALIZATION ************/
/**
 * Accepts both formats:
 * 1) YYYY.MM.###.DBA.ODL
 * 2) YYYY.MM###.DBA.ODL
 *
 * Returns an array of normalized variants to compare.
 */
function normalizeStudentIdVariants_(id) {
  id = String(id || "").trim().toUpperCase().replace(/\s+/g, "");

  // Try dotted: 2025.09.019.DBA.ODL
  let m = id.match(/^(\d{4})\.(\d{2})\.(\d{3})\.DBA\.ODL$/);
  if (m) {
    const year = m[1], month = m[2], seq = m[3];
    return [
      `${year}.${month}.${seq}.DBA.ODL`,
      `${year}.${month}${seq}.DBA.ODL`
    ];
  }

  // Try merged: 2025.09019.DBA.ODL
  m = id.match(/^(\d{4})\.(\d{2})(\d{3})\.DBA\.ODL$/);
  if (m) {
    const year = m[1], month = m[2], seq = m[3];
    return [
      `${year}.${month}.${seq}.DBA.ODL`,
      `${year}.${month}${seq}.DBA.ODL`
    ];
  }

  // Unknown format: return as-is
  return [id];
}

/************ DRIVE FILE SEARCH ************/
function findSlipFileInFolder_(fullName) {
  const folder = DriveApp.getFolderById(CONFIG.SLIPS_FOLDER_ID);

  let expectedBase = fullName;
  if (CONFIG.FILE_MATCH_MODE === "NAME_SLIP") {
    expectedBase = `${fullName} - Enrollment Slip`;
  }

  const expectedNorm = normalizeExact_(expectedBase);
  const fullNameNorm = normalizeExact_(fullName);

  const files = folder.getFiles();
  let best = null;
  let bestUpdated = 0;

  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    const base = name.replace(/\.[^/.]+$/, ""); // remove extension
    const baseNorm = normalizeExact_(base);

    // Match expected pattern OR name-only
    if (baseNorm === expectedNorm || baseNorm === fullNameNorm) {
      const upd = f.getLastUpdated().getTime();
      if (upd > bestUpdated) {
        best = f;
        bestUpdated = upd;
      }
    }
  }

  return best;
}

/************ REQUEST LOG ************/
function logRequest_(name, studentId, reason) {
  const sh = ensureRequestsSheet_();
  const cols = getRequestsCols_(sh);
  const row = makeRow_(cols);
  row[cols.timestamp] = new Date();
  row[cols.name] = name || "";
  row[cols.studentId] = studentId || "";
  row[cols.reason] = reason || "";
  sh.appendRow(row);
}

/************ ISSUE SUBMISSION ************/
function submitIssue(issueType, typedName, contact) {
  issueType = String(issueType || "").trim().toUpperCase();
  typedName = String(typedName || "").trim();
  contact = String(contact || "").trim();

  if (!typedName || !contact) {
    return { ok: false, code: "MISSING_INPUT", message: "Name and contact info are required." };
  }

  if (!["NAME_NOT_FOUND", "NO_ID", "SLIP_NOT_RECEIVED"].includes(issueType)) {
    return { ok: false, code: "BAD_TYPE", message: "Invalid issue type." };
  }

  const sh = ensureRequestsSheet_();
  const cols = getRequestsCols_(sh);

  const normalizedContact = normalizeContact_(contact);
  if (hasDuplicateIssue_(sh, cols, normalizedContact)) {
    return { ok: false, code: "DUPLICATE", message: "We already received a request from this contact." };
  }

  const row = makeRow_(cols);
  row[cols.timestamp] = new Date();
  row[cols.typedName] = typedName;
  row[cols.contact] = contact;
  row[cols.issueType] = issueType;
  sh.appendRow(row);

  return { ok: true };
}

function hasDuplicateIssue_(sh, cols, normalizedContact) {
  if (!normalizedContact) return false;

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const contact = normalizeContact_(row[cols.contact] || "");
    const issueType = String(row[cols.issueType] || "").trim().toUpperCase();
    const typedName = normalizeExact_(row[cols.typedName] || "");

    if (!contact || !issueType) continue;
    if (contact === normalizedContact) {
      return true;
    }
  }
  return false;
}

function ensureRequestsSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.REQUESTS_SHEET_NAME);

  const required = [
    "Timestamp",
    "TypedName",
    "Email/Phone",
    "IssueType",
    "Name",
    "StudentID",
    "Reason"
  ];

  if (!sh) {
    sh = ss.insertSheet(CONFIG.REQUESTS_SHEET_NAME);
    sh.appendRow(required);
    return sh;
  }

  const values = sh.getDataRange().getValues();
  const headers = values.length > 0 ? values[0].map(h => String(h).trim()) : [];
  const headerMap = {};
  headers.forEach((h, i) => headerMap[normalizeHeader_(h)] = i);

  const missing = required.filter(h => headerMap[normalizeHeader_(h)] == null);
  if (missing.length > 0) {
    const startCol = headers.length + 1;
    sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }
  return sh;
}

function getRequestsCols_(sh) {
  const headers = sh.getDataRange().getValues()[0].map(h => String(h).trim());
  const map = {};
  headers.forEach((h, i) => map[normalizeHeader_(h)] = i);

  const required = {
    timestamp: "timestamp",
    typedName: "typedname",
    contact: "email/phone",
    issueType: "issuetype",
    name: "name",
    studentId: "studentid",
    reason: "reason"
  };

  const cols = {};
  Object.keys(required).forEach(key => {
    const idx = map[required[key]];
    cols[key] = idx == null ? headers.length : idx;
  });
  return cols;
}

function makeRow_(cols) {
  const max = Math.max.apply(null, Object.values(cols));
  return Array(max + 1).fill("");
}

/************ TEXT HELPERS ************/
function normalize_(s) {
  return String(s || "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeExact_(s) {
  return String(s || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function normalizeHeader_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

function normalizeContact_(s) {
  const raw = String(s || "").trim();
  if (!raw) return "";
  if (raw.indexOf("@") !== -1) {
    return raw.toLowerCase();
  }
  return raw.replace(/[^\d+]/g, "");
}

function matchesTokensInOrder_(text, tokens) {
  let idx = 0;
  for (const t of tokens) {
    const pos = text.indexOf(t, idx);
    if (pos === -1) return false;
    idx = pos + t.length;
  }
  return true;
}

function getIndexCols_(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h).toLowerCase()] = i);

  const fullName = map["fullname"];
  const studentId = map["studentid"];
  const status = map["status"];

  if (fullName == null || studentId == null || status == null) {
    throw new Error("INDEX sheet must have headers: FullName, StudentID, Status");
  }
  return { fullName, studentId, status };
}

/************ BASIC ATTEMPT LIMITING ************/
function checkLockout_() {
  const props = PropertiesService.getUserProperties();
  const lockedUntil = Number(props.getProperty("lockedUntil") || "0");
  const now = Date.now();

  if (lockedUntil > now) {
    const remainingMs = lockedUntil - now;
    return { locked: true, remainingMinutes: Math.ceil(remainingMs / 60000) };
  }
  return { locked: false, remainingMinutes: 0 };
}

function registerFailedAttempt_() {
  const props = PropertiesService.getUserProperties();
  const fails = Number(props.getProperty("fails") || "0") + 1;
  props.setProperty("fails", String(fails));

  if (fails >= CONFIG.MAX_FAILED_ATTEMPTS) {
    const lockedUntil = Date.now() + CONFIG.LOCKOUT_MINUTES * 60000;
    props.setProperty("lockedUntil", String(lockedUntil));
  }
}

function resetFailedAttempts_() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty("fails");
  props.deleteProperty("lockedUntil");
}
