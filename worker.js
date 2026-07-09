// ═══════════════════════════════════════════════════════════════════════
// HiFive Recruitment Portal — Cloudflare Worker v1.0
// Endpoints: verify-pin, get-candidates, get-candidate, add-candidate,
//            update-candidate, add-interview-round, update-interview-round,
//            upload-cv, get-dashboard
// Backend: OneDrive Excel via MS Graph
// ═══════════════════════════════════════════════════════════════════════

const TENANT_ID   = "f83af262-079f-4e52-a1e1-721873074bad";
const CLIENT_ID   = "2e594c92-1631-4cdb-ac20-355b939bcc60";
const SHEET_CANDIDATES     = "Candidates";
const SHEET_ROUNDS         = "InterviewRounds";
const SHEET_CONFIG         = "Config";
const SHEET_REQUISITIONS   = "Requisitions";
const SHEET_PIPELINE       = "PipelineStages";
const SHEET_CVDB           = "CVDatabase";
const SHEET_ACTIVITY       = "ActivityLog";
const SHEET_USERS          = "Users";
const SHEET_PIPELINE_HISTORY = "Pipeline_History"; // timestamped log of every stage transition
const SHEET_QUOTES           = "Quotes";           // witty wishes and good quotes, managed in Excel
const SHEET_NEWS             = "News";             // blog/news posts for careers page

// Pipeline_History columns (A–H, 0-based)
const PH = {
  TIMESTAMP:    0, // A — ISO datetime of the transition
  CAND_ID:      1, // B — REC-XXXX
  CAND_NAME:    2, // C — candidate name
  POSITION:     3, // D — position applied for
  ENTITY:       4, // E — entity
  FROM_STAGE:   5, // F — previous status before this transition
  TO_STAGE:     6, // G — new stage / status
  OUTCOME:      7, // H — Pass / Fail / Pending / Rejected
  DAYS_IN_PREV: 8, // I — days spent in the FROM_STAGE before moving
  ADVANCED_BY:  9, // J — user who triggered the transition
};

// Users sheet columns (0-based)
// UserKey | DisplayName | Role | PIN | EntityFilter | NavPages | Active
const U = {
  KEY:0, NAME:1, ROLE:2, PIN:3, ENTITY_FILTER:4, NAV_PAGES:5, ACTIVE:6,
  EMAIL:7, WHATSAPP:8, NOTIFY_APPS:9, NOTIFY_EVENTS:10,
};

// ── Column indices (0-based) ─────────────────────────────────────────────
const C = {
  // Candidates sheet
  ID:0,DATE:1,NAME:2,PHONE:3,EMAIL:4,POSITION:5,ENTITY:6,DEGREE:7,
  STATUS:8,DRIVERS_LICENSE:9,EMIRATE:10,NOTICE_PERIOD:11,UAE_EXP:12,
  VISA_TYPE:13,VISA_EXPIRY:14,EXPECTED_SALARY:15,SOURCE:16,SOURCE_COST:17,
  AVAILABILITY:18,INTERVIEW_STATUS:19,INTERVIEW_DATE:20,REMARKS:21,
  PREV_EMPLOYER:22,CURR_EMPLOYER:23,RELATION:24,NATIONALITY:25,
  LOCATION:26,YEAR_MONTH:27,CV_LINK:28,DUPLICATE_FLAG:29,
  REQUISITION_ID:30,REQUESTED_BY:31,APPROVED_BY:32,OFFER_STAGE:33,
  OFFERED_SALARY:34,DECLINE_REASON:35,EXPECTED_JOINING:36,
  ACTUAL_JOINING:37,TIME_TO_HIRE:38,CVDB_ID:39,  // AN — linked CVDB-XXXX
  EXCLUDED_ROUNDS:40, // AO — comma-separated round labels to exclude from score
  // EXCLUDED_RATERS at 41 (AP) — defined as EXCLUDED_RATERS_COL constant below
  NOTICE_DAYS_SERVED:42, // AQ — actual notice days served by candidate (entered when marking Joined)
  SHARE_TOKEN:43,        // AR — JSON { token, expires, visibility } for management share links
  AI_RATING:44,          // AS — AI-generated match score 0-100 against job description
  AI_RATING_REASON:45,   // AT — AI-generated short rationale for the score
};

// CVDatabase columns (0-based)
const CVDB = {
  ID:0,            // CVDB-0001
  NAME:1,
  POSITION:2,      // mandatory dropdown selection on careers page
  FACILITY:3,      // mandatory dropdown selection on careers page (= entity)
  SUBMITTED_AT:4,  // ISO datetime
  CV_LINK:5,       // OneDrive share link
  PHONE:6,
  EMAIL:7,
  SOURCE:8,        // always "Careers Page CV Drop"
  LINKED_CAND_ID:9,// J — linked REC-XXXX candidate ID
};

// PipelineStages columns (0-based)
const PS = {
  ID:0,           // CandidateID
  STAGE:1,        // HR Screen | Operations | Management | Verbal Offer | Written Offer | Accepted | Declined | Onboarding | Joined
  DATE:2,         // Stage date
  INTERVIEWER:3,  // Interviewer/reviewer name
  OUTCOME:4,      // Pass | Fail | Hold | Pending | Declined | Joined
  NOTES:5,        // Free text notes
  SCORE:6,        // /10 (for interview stages)
  SALARY:7,       // Offered salary (offer stages)
  JOINING_DATE:8, // Expected joining (offer stage)
  VISA:9,         // Company visa Y/N (offer stage)
  SENIOR:10,      // Senior/Junior flag for management stage
  DECLINE_REASON:11, // If declined
  ONBOARDING_JSON:12, // JSON checklist state
  CREATED_AT:13,  // Row created timestamp
};
const CV_FOLDER_PATH       = "/Recruitment/CVs";   // OneDrive root-relative
const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, X-App-Key",
};

const R = {
  // InterviewRounds sheet
  CAND_ID:0,ROUND:1,INTERVIEWER:2,SCHED_DATE:3,ACTUAL_DATE:4,
  OUTCOME:5,RESCHEDULE_COUNT:6,SCORE:7,FEEDBACK:8,
  SCHED_TIME:9,       // HH:MM, separate from SCHED_DATE
  TEAM:10,            // "Facility" | "Management" | "Both"
  INTERVIEWER_ROLE:11,// e.g. "Lab Manager", "Operations Director" — free text
  STATUS:12,          // "Scheduled" | "Completed" | "Skipped" | "Cancelled"
  // External interviewer rating link fields
  RATING_TOKEN:13,          // N — secure random token
  RATING_TOKEN_EXPIRES:14,  // O — ISO datetime, 7 days after interview date
  RATING_SUBMITTED_AT:15,   // P — ISO datetime when interviewer submitted
  // 10 criteria scores (1–10 each), omission tracked separately
  SCORE_EDUCATION:16,       // Q — Educational Background
  SCORE_WORK_EXP:17,        // R — Prior Work Experience
  SCORE_TECHNICAL:18,       // S — Technical Qualifications/Experience
  SCORE_COMMUNICATION:19,   // T — Verbal Communication
  SCORE_ENTHUSIASM:20,      // U — Candidate Enthusiasm
  SCORE_PRODUCT:21,         // V — Product/Service Knowledge
  SCORE_TEAMWORK:22,        // W — Teambuilding/Interpersonal Skills
  SCORE_INITIATIVE:23,      // X — Initiative
  SCORE_TIME_MGMT:24,       // Y — Time Management
  SCORE_COMPANY_KNOWLEDGE:25,// Z — Knowledge of the Company
  OMITTED_CRITERIA:26,      // AA — comma-separated omitted criteria keys
  RATING_PASS_FAIL:27,      // AB — "Pass" | "Fail"
  RATING_REMARKS:28,        // AC — free text
  SUBMITTED_BY_NAME:29,     // AD — name entered by interviewer on confirmation
  SHORT_CODE:30,            // AE — 6-char alphanumeric short code
  // Pre-outcome screening questions
  Q_SHIFT_DUTIES:31,        // AF — "Yes" | "No" | "N/A"
  Q_DISCIPLINARY:32,        // AG — "Yes" | "No"
  Q_RELATIVE:33,            // AH — "Yes" | "No"
  Q_RELATIVE_DETAIL:34,     // AI — name + relation if yes
  Q_COMPETITOR:35,          // AJ — "Yes" | "No"
  Q_COMPETITOR_DETAIL:36,   // AK — which competitor or client
  RATER_ID:37,              // AL — unique ID per rater row (for multi-rater support)
  SCORE_EMOTIONAL_SENSITIVITY:38, // AM — Emotional Sensitivity (new 11th criterion)
  INTERVIEW_DURATION:39,  // AN — duration in minutes
};

// Candidates sheet: add excluded raters column
// C.EXCLUDED_ROUNDS already at 40 — add EXCLUDED_RATERS at 41
const EXCLUDED_RATERS_COL = 41; // AP — comma-separated rater IDs to exclude from score

// Requisitions sheet columns (0-based)
const REQ = {
  ID:0, DATE:1, ENTITY:2, POSITION:3, DEPARTMENT:4,
  HEADCOUNT:5, REASON:6, BUDGETED_SALARY_MIN:7, BUDGETED_SALARY_MAX:8,
  REQUESTED_BY:9, REQUEST_DATE:10,
  HR_STATUS:11, HR_REVIEWED_BY:12, HR_REVIEW_DATE:13, HR_REMARKS:14,
  CEO_STATUS:15, CEO_APPROVED_BY:16, CEO_APPROVAL_DATE:17, CEO_REMARKS:18,
  APPROVED_VIA:34, // "CEO" (direct) or "HR on behalf of <CEO name>"
  CHANGE_LOG:35,   // AJ — pipe-separated audit trail of post-approval edits
  OVERALL_STATUS:19, FILLED_COUNT:20, CLOSED_DATE:21, YEAR_MONTH:22,
  // New fields
  TYPE:23,              // New Position / Replacement
  REPLACING_EMPLOYEE:24,// Name of employee being replaced
  JOB_DESCRIPTION:25,  // Free text or link
  COMPANY_VISA:26,      // Yes / No
  EXPECTED_JOINING:27,  // Target joining date
  CUSTOM_POSITION:28,   // If position not in master list
  BRANCH:29,            // Branch / Location
  REPORTING_MANAGER:30, // Reporting manager name
  GENDER_PREFERENCE:31, // Male / Female / Both / None
  EXPERIENCE_REQUIRED:32, // 0–10+ years
  NEW_POSITION_JUSTIFICATION:33, // Why new position needed
  SHOW_ON_CAREERS:36,   // AK — Yes/No, default Yes
  // New fields added this session
  DEPT_STRENGTH:37,     // AL — current headcount in department
  DEPARTURE_REASON:38,  // AM — why previous employee left (for Replacement type)
  NATIONALITY_PREF:39,  // AN — nationality preference for the role
  EMERGENCY_HIRING:40,  // AO — "Yes" if urgent/emergency hire
};

// ── Helpers ───────────────────────────────────────────────────────────────
function toLetter(i) {
  // 0-indexed → Excel column letter
  let s = "";
  i++;
  while (i > 0) {
    i--;
    s = String.fromCharCode(65 + (i % 26)) + s;
    i = Math.floor(i / 26);
  }
  return s;
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS, "Content-Type": "application/json" },
  });
}

function err(msg, status = 400) {
  return json({ error: msg }, status);
}

function sanitize(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim().replace(/[\x00-\x1F]/g, "");
}

// Convert Excel serial date (e.g. 45678) to ISO string YYYY-MM-DD
function excelDateToISO(v) {
  if (!v && v !== 0) return "";
  const n = parseFloat(v);
  if (isNaN(n)) return String(v).trim(); // already a string date
  if (n < 1) return "";                  // pure time fraction, not a date
  if (n < 1000) return "";               // not a date
  // Excel epoch: Dec 30 1899; JS epoch: Jan 1 1970
  const MS_PER_DAY = 86400000;
  const EXCEL_EPOCH = new Date(1899, 11, 30).getTime();
  const date = new Date(EXCEL_EPOCH + n * MS_PER_DAY);
  return date.toISOString().split("T")[0];
}

// Excel stores times as a fraction of a day (14:00 = 14/24 ≈ 0.5833...).
// This converts that fraction back to "HH:MM" for display.
function excelTimeToHHMM(v) {
  if (!v && v !== 0) return "";
  const n = parseFloat(v);
  if (isNaN(n)) return String(v).trim(); // already a string like "14:00"
  if (n >= 1) return "";                 // this is a date, not a time
  const totalMins = Math.round(n * 24 * 60);
  const h = Math.floor(totalMins / 60) % 24;
  const m = totalMins % 60;
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
}

function yearMonth() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function isoDate() {
  return new Date().toISOString().split("T")[0];
}

// ── MS Graph token ────────────────────────────────────────────────────────
// ── In-memory token cache ──────────────────────────────────────────────────
// Workers are ephemeral but a single isolate handles many requests in its lifetime.
// Caching the token in module scope avoids a round-trip to login.microsoft.com
// on every API call. Token lifetime is 3600s; we refresh 5 minutes early.
let _tokenCache = null;
let _tokenExpiry = 0;

async function getToken(env) {
  const now = Date.now();
  if (_tokenCache && now < _tokenExpiry) return _tokenCache;

  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     CLIENT_ID,
        client_secret: env.MS_CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default",
      }),
    }
  );
  const d = await res.json();
  if (!d.access_token) throw new Error("Token fetch failed: " + JSON.stringify(d));
  _tokenCache = d.access_token;
  // Cache for (expires_in - 300) seconds; default 3600s → 3300s
  _tokenExpiry = now + ((d.expires_in || 3600) - 300) * 1000;
  return _tokenCache;
}

// ── Excel helpers ─────────────────────────────────────────────────────────
// ── Graph API fetch with exponential backoff + jitter ─────────────────────
// Retries on 429 (rate-limit) and 503/504 (transient Graph errors).
// Jitter prevents thundering-herd when multiple Workers retry simultaneously.
// Max 4 attempts: ~0ms, ~600ms, ~1.8s, ~4.2s — stays within Worker CPU limits.
async function graphFetch(url, options = {}) {
  const MAX_ATTEMPTS = 4;
  const BASE_DELAY   = 500; // ms

  for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
    const res = await fetch(url, options);

    // Success or a client error we shouldn't retry (4xx except 429)
    if (res.ok) return res;
    if (res.status !== 429 && res.status < 500) return res; // let caller handle

    if (attempt === MAX_ATTEMPTS - 1) return res; // exhausted — return last response

    // Respect Retry-After header if present (Graph API sets this on 429)
    const retryAfter = parseInt(res.headers.get("Retry-After") || "0", 10);
    const backoff    = retryAfter > 0
      ? retryAfter * 1000
      : BASE_DELAY * Math.pow(2, attempt) + Math.random() * 300; // jitter ±300ms

    await new Promise(r => setTimeout(r, Math.min(backoff, 8000))); // cap at 8s
  }
}

async function getRows(token, env, sheet) {
  const encoded = encodeURIComponent(sheet);
  const url = `https://graph.microsoft.com/v1.0/users/${env.ONEDRIVE_USER}/drive/items/${env.EXCEL_FILE_ID}/workbook/worksheets/${encoded}/usedRange`;
  const res = await graphFetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    cf: { cacheTtl: 10, cacheEverything: false },
  });
  if (!res.ok) throw new Error(`getRows(${sheet}) failed: ${res.status}`);
  const d = await res.json();
  const values = d.values || [];
  if (!values.length) return values;
  const width = values[0].length;
  return values.map(row => {
    if (row.length >= width) return row;
    const padded = [...row];
    while (padded.length < width) padded.push("");
    return padded;
  });
}

async function updateCell(token, env, sheet, row, col, value) {
  const colLetter = toLetter(col);
  const addr      = `${colLetter}${row}`;
  const encoded   = encodeURIComponent(sheet);
  const url = `https://graph.microsoft.com/v1.0/users/${env.ONEDRIVE_USER}/drive/items/${env.EXCEL_FILE_ID}/workbook/worksheets/${encoded}/range(address='${addr}')`;
  await graphFetch(url, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ values: [[value]] }),
  });
}

async function updateRow(token, env, sheet, excelRow, values) {
  const startCol = toLetter(0);
  const endCol   = toLetter(values.length - 1);
  const addr     = `${startCol}${excelRow}:${endCol}${excelRow}`;
  const encoded  = encodeURIComponent(sheet);
  const url = `https://graph.microsoft.com/v1.0/users/${env.ONEDRIVE_USER}/drive/items/${env.EXCEL_FILE_ID}/workbook/worksheets/${encoded}/range(address='${addr}')`;
  await graphFetch(url, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ values: [values] }),
  });
}

async function appendRow(token, env, sheet, values) {
  const rows = await getRows(token, env, sheet);
  const nextRow = rows.length + 1;
  await updateRow(token, env, sheet, nextRow, values);
  return nextRow;
}

// ── Pipeline History writer ───────────────────────────────────────────────
// Called after every stage advance. Writes one row per transition so we can
// compute avg time per stage and identify bottlenecks in analytics.
async function appendPipelineHistory(token, env, { candID, candName, position, entity, fromStage, toStage, outcome, advancedBy }) {
  try {
    // Compute days spent in the previous stage by looking at the last
    // Pipeline_History entry for this candidate with the same fromStage.
    // If not found, fall back to candidate entry date.
    let daysInPrev = null;
    try {
      const histRows = await getRows(token, env, SHEET_PIPELINE_HISTORY);
      // Find most recent row for this candidate entering fromStage
      for (let i = histRows.length - 1; i >= 1; i--) {
        if (sanitize(histRows[i][PH.CAND_ID]) === candID &&
            sanitize(histRows[i][PH.TO_STAGE]) === fromStage) {
          const enteredAt = new Date(sanitize(histRows[i][PH.TIMESTAMP]));
          if (!isNaN(enteredAt)) {
            daysInPrev = Math.round((Date.now() - enteredAt.getTime()) / 86400000);
          }
          break;
        }
      }
    } catch {}

    const row = new Array(10).fill("");
    row[PH.TIMESTAMP]    = new Date().toISOString();
    row[PH.CAND_ID]      = candID    || "";
    row[PH.CAND_NAME]    = candName  || "";
    row[PH.POSITION]     = position  || "";
    row[PH.ENTITY]       = entity    || "";
    row[PH.FROM_STAGE]   = fromStage || "";
    row[PH.TO_STAGE]     = toStage   || "";
    row[PH.OUTCOME]      = outcome   || "";
    row[PH.DAYS_IN_PREV] = daysInPrev != null ? daysInPrev : "";
    row[PH.ADVANCED_BY]  = advancedBy || "";

    await appendRow(token, env, SHEET_PIPELINE_HISTORY, row);
  } catch(e) {
    // Never let history write failure block the main operation
    console.error("Pipeline history write failed:", e.message);
  }
}

// Deletes a single row from a sheet via the Graph API's range delete,
// which shifts all rows below it up by one — a true removal, not a soft-delete.
async function deleteSheetRow(token, env, sheet, excelRow) {
  const encoded = encodeURIComponent(sheet);
  const addr    = `A${excelRow}:Z${excelRow}`;
  const url = `https://graph.microsoft.com/v1.0/users/${env.ONEDRIVE_USER}/drive/items/${env.EXCEL_FILE_ID}/workbook/worksheets/${encoded}/range(address='${addr}')/delete`;
  await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ shift: "Up" }),
  });
}

// POST /update-cv-database — edit a CV Database entry's details (name, position,
// facility, phone, email). Does not touch the uploaded file itself.
async function handleUpdateCVDB(body, env) {
  const { cvID, name, position, facility, phone, email } = body;
  if (!cvID) return err("cvID required");

  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CVDB);
  let excelRow = -1, existing = null;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][CVDB.ID]) === cvID) { excelRow = i + 1; existing = rows[i]; break; }
  }
  if (excelRow === -1) return err("CV entry not found", 404);

  const newRow = [...existing];
  while (newRow.length < 9) newRow.push("");
  if (name     !== undefined) newRow[CVDB.NAME]     = sanitize(name);
  if (position !== undefined) newRow[CVDB.POSITION] = sanitize(position);
  if (facility !== undefined) newRow[CVDB.FACILITY] = sanitize(facility);
  if (phone    !== undefined) newRow[CVDB.PHONE]    = sanitize(phone);
  if (email    !== undefined) newRow[CVDB.EMAIL]    = sanitize(email);

  await updateRow(token, env, SHEET_CVDB, excelRow, newRow);
  return json({ success: true, cvID });
}

// POST /delete-cv-database — permanently removes a CV Database entry's row.
// Does not delete the underlying OneDrive file, only the tracking row.
async function handleDeleteCVDB(body, env) {
  const { cvID } = body;
  if (!cvID) return err("cvID required");

  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CVDB);
  let excelRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][CVDB.ID]) === cvID) { excelRow = i + 1; break; }
  }
  if (excelRow === -1) return err("CV entry not found", 404);

  await deleteSheetRow(token, env, SHEET_CVDB, excelRow);
  return json({ success: true, cvID });
}

// POST /replace-cv-database-file — uploads a new CV file for an existing entry
// and updates the CVLink to point at it. The old file is left in place in
// OneDrive (same candidate folder) rather than deleted, so nothing is lost
// even if the wrong file gets uploaded by mistake.
async function handleReplaceCVDBFile(request, env) {
  const formData = await request.formData();
  const cvID = sanitize(formData.get("cvID") || "");
  const file = formData.get("file");
  if (!cvID) return err("cvID required");
  if (!file) return err("file required");

  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CVDB);
  let excelRow = -1, existing = null;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][CVDB.ID]) === cvID) { excelRow = i + 1; existing = rows[i]; break; }
  }
  if (excelRow === -1) return err("CV entry not found", 404);

  const fileName = file.name || "cv.pdf";
  const bytes    = await file.arrayBuffer();
  const cvUrl    = await uploadCV(token, env, cvID, fileName, bytes);

  const newRow = [...existing];
  while (newRow.length < 9) newRow.push("");
  newRow[CVDB.CV_LINK] = cvUrl;

  await updateRow(token, env, SHEET_CVDB, excelRow, newRow);
  return json({ success: true, cvID, cvLink: cvUrl });
}

// ── Config helpers ────────────────────────────────────────────────────────
async function getConfig(token, env) {
  const rows = await getRows(token, env, SHEET_CONFIG);
  const map  = {};
  for (const row of rows) {
    const k = sanitize(row[0]);
    const v = sanitize(row[1]);
    if (k) map[k] = v;
  }
  return map;
}

// ── Candidate ID generator ────────────────────────────────────────────────
async function nextCandidateID(token, env) {
  const rows = await getRows(token, env, SHEET_CANDIDATES);
  let max = 0;
  for (let i = rows.length - 1; i >= 1; i--) {
    const id = sanitize(rows[i][C.ID]);
    if (id.startsWith("REC-")) {
      const n = parseInt(id.replace("REC-", ""), 10);
      if (!isNaN(n) && n > max) { max = n; break; }
    }
  }
  return `REC-${String(max + 1).padStart(4, "0")}`;
}

// ── Duplicate check ───────────────────────────────────────────────────────
async function checkDuplicate(token, env, name, phone, email, excludeID = null) {
  const rows = await getRows(token, env, SHEET_CANDIDATES);
  const nName  = name.toLowerCase().trim();
  const nPhone = sanitize(phone).replace(/\s/g, "");
  const nEmail = sanitize(email).toLowerCase().trim();
  const matches = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const rowID = sanitize(row[C.ID]);
    if (excludeID && rowID === excludeID) continue;
    const rName   = sanitize(row[C.NAME]).toLowerCase();
    const rPhone  = sanitize(row[C.PHONE]).replace(/\s/g, "");
    const rEmail  = sanitize(row[C.EMAIL]).toLowerCase();
    const nameMatch  = rName === nName;
    const phoneMatch = nPhone && rPhone === nPhone;
    const emailMatch = nEmail && rEmail === nEmail;
    if (nameMatch && (phoneMatch || emailMatch)) {
      matches.push({
        existingID:  rowID,
        name:        sanitize(row[C.NAME]),
        entity:      sanitize(row[C.ENTITY]),
        position:    sanitize(row[C.POSITION]),
        status:      sanitize(row[C.STATUS]),
        matchedOn:   phoneMatch && emailMatch ? "name + phone + email" : phoneMatch ? "name + phone" : "name + email",
      });
    }
  }
  if (matches.length > 0) {
    return { duplicate: true, existingID: matches[0].existingID, matches };
  }
  return { duplicate: false, matches: [] };
}

// Duplicate check for standalone CV Database — same name + (phone or email)
// + same position + facility, submitted within the last 10 minutes.
// Prevents accidental double-clicks / resubmits from creating duplicate
// OneDrive uploads and CVDB rows, without blocking a genuinely new application
// for a different role weeks later.
async function checkCVDBDuplicate(token, env, name, phone, email, position, facility) {
  const rows = await getRows(token, env, SHEET_CVDB).catch(() => [[]]);
  const nName  = name.toLowerCase().trim();
  const nPhone = sanitize(phone).replace(/\s/g, "");
  const nEmail = sanitize(email).toLowerCase().trim();
  const nPos   = position.toLowerCase().trim();
  const nFac   = facility.toLowerCase().trim();
  const tenMinAgo = Date.now() - 10 * 60 * 1000;

  for (let i = rows.length - 1; i >= 1; i--) {
    const row = rows[i];
    if (!row || !row[CVDB.ID]) continue;
    const rName  = sanitize(row[CVDB.NAME]).toLowerCase();
    const rPhone = sanitize(row[CVDB.PHONE]).replace(/\s/g, "");
    const rEmail = sanitize(row[CVDB.EMAIL]).toLowerCase();
    const rPos   = sanitize(row[CVDB.POSITION]).toLowerCase();
    const rFac   = sanitize(row[CVDB.FACILITY]).toLowerCase();
    const rTime  = Date.parse(sanitize(row[CVDB.SUBMITTED_AT])) || 0;

    const sameIdentity = rName === nName && (nPhone && rPhone === nPhone || nEmail && rEmail === nEmail);
    const sameRole     = rPos === nPos && rFac === nFac;

    if (sameIdentity && sameRole && rTime >= tenMinAgo) {
      return { duplicate: true, existingID: sanitize(row[CVDB.ID]) };
    }
  }
  return { duplicate: false };
}

// Permanent duplicate check for the Add Candidate → CV Database auto-link.
// Unlike checkCVDBDuplicate (which only blocks rapid double-submits on the
// careers page), this checks the ENTIRE database with no time window —
// used when a candidate added directly in the portal should also be filed
// into the CV Database, but only if they're not already in there.
async function checkCVDBPermanentDuplicate(token, env, name, email, fileName) {
  const rows = await getRows(token, env, SHEET_CVDB).catch(() => [[]]);
  const nName = (name || "").toLowerCase().trim();
  const nEmail = sanitize(email || "").toLowerCase().trim();
  // Loose filename match — strip extension and non-alphanumerics so
  // "John_Doe_CV.pdf" and "John Doe CV (1).pdf" still compare equal.
  const normalizeFile = (f) => sanitize(f || "").toLowerCase().replace(/\.[a-z0-9]+$/i, "").replace(/[^a-z0-9]/g, "");
  const nFile = normalizeFile(fileName);

  for (let i = rows.length - 1; i >= 1; i--) {
    const row = rows[i];
    if (!row || !row[CVDB.ID]) continue;
    const rName  = sanitize(row[CVDB.NAME]).toLowerCase().trim();
    const rEmail = sanitize(row[CVDB.EMAIL]).toLowerCase().trim();
    const rLink  = sanitize(row[CVDB.CV_LINK]);
    const rFileFromLink = normalizeFile(decodeURIComponent(rLink.split("/").pop() || ""));

    const nameMatch  = nName && rName === nName;
    const emailMatch = nEmail && rEmail === nEmail;
    const fileMatch  = nFile && rFileFromLink && rFileFromLink === nFile;

    if (nameMatch && emailMatch) return { duplicate: true, existingID: sanitize(row[CVDB.ID]), reason: "name+email" };
    if (nameMatch && fileMatch)  return { duplicate: true, existingID: sanitize(row[CVDB.ID]), reason: "name+cv" };
  }
  return { duplicate: false };
}

// ── CV Upload to OneDrive ─────────────────────────────────────────────────
async function uploadCV(token, env, candidateID, fileName, fileBytes) {
  const safeName = fileName.replace(/[^a-zA-Z0-9._-]/g, "_");
  const path     = `${CV_FOLDER_PATH}/${candidateID}/${safeName}`;
  const encoded  = encodeURIComponent(path).replace(/%2F/g, "/");
  const url = `https://graph.microsoft.com/v1.0/users/${env.ONEDRIVE_USER}/drive/root:/${encoded}:/content`;
  const res = await fetch(url, {
    method:  "PUT",
    headers: {
      Authorization:  `Bearer ${token}`,
      "Content-Type": "application/octet-stream",
    },
    body: fileBytes,
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`CV upload failed: ${res.status} ${t}`);
  }
  const d = await res.json();
  const itemId = d.id;

  // Create anonymous share link (anyone with link can view — no login required)
  try {
    const shareRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${env.ONEDRIVE_USER}/drive/items/${itemId}/createLink`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          type: "view",        // view-only
          scope: "anonymous",  // no login required
        }),
      }
    );
    if (shareRes.ok) {
      const shareData = await shareRes.json();
      return shareData.link?.webUrl || d.webUrl || "";
    }
  } catch (e) {
    console.warn("Share link creation failed:", e.message);
  }

  return d.webUrl || "";
}

// ── WhatsApp notify ───────────────────────────────────────────────────────
async function notifyWhatsApp(env, message) {
  if (!env.WHATSAPP_TOKEN || !env.WHATSAPP_PHONE_ID) return;
  const to = "971542346392";
  await fetch(
    `https://graph.facebook.com/v19.0/${env.WHATSAPP_PHONE_ID}/messages`,
    {
      method: "POST",
      headers: {
        Authorization:  `Bearer ${env.WHATSAPP_TOKEN}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        messaging_product: "whatsapp",
        to,
        type: "text",
        text: { body: message },
      }),
    }
  );
}

// ════════════════════════════════════════════════════════════════════════
// ROUTE HANDLERS
// ════════════════════════════════════════════════════════════════════════

// POST /verify-pin
// Simple brute-force protection for PIN login using Cloudflare's built-in Cache API
// (no KV namespace setup required). Locks out an IP for 5 minutes after 5 failed attempts.
const PIN_MAX_ATTEMPTS = 5;
const PIN_LOCKOUT_SECONDS = 300; // 5 minutes

async function getPinAttemptState(ip) {
  const cache = caches.default;
  const key = new Request(`https://hf-pin-guard.internal/${ip}`);
  const cached = await cache.match(key);
  if (!cached) return { count: 0 };
  return await cached.json();
}
async function setPinAttemptState(ip, state) {
  const cache = caches.default;
  const key = new Request(`https://hf-pin-guard.internal/${ip}`);
  const res = new Response(JSON.stringify(state), {
    headers: { "Cache-Control": `max-age=${PIN_LOCKOUT_SECONDS}`, "Content-Type": "application/json" },
  });
  await cache.put(key, res);
}

async function handleVerifyPin(body, env, ip) {
  const { pin, role: submittedRole, userKey } = body;
  if (!pin) return err("pin required");

  const attemptState = await getPinAttemptState(ip);
  if (attemptState.count >= PIN_MAX_ATTEMPTS) {
    return json({ success: false, role: null, lockedOut: true, message: "Too many incorrect attempts. Please wait a few minutes and try again." }, 429);
  }

  const token = await getToken(env);

  // --- 1. Try Users sheet first (individual named users) ---
  // SECURITY: Match on BOTH userKey AND pin together.
  // Matching pin alone means two users who coincidentally get the same 4-digit
  // PIN after nightly rotation can log in as each other (confirmed issue: Jafar/Jithesh).
  // If a userKey is supplied (named login), it MUST also match before the PIN is checked.
  const userRows = await getRows(token, env, SHEET_USERS).catch(() => [[]]);
  for (let i = 1; i < userRows.length; i++) {
    const r = userRows[i];
    if (!r[U.KEY]) continue;
    if (sanitize(r[U.ACTIVE]).toLowerCase() === "false") continue;
    // If caller identified themselves, enforce that identity check first.
    if (userKey && sanitize(r[U.KEY]) !== userKey) continue;
    // Now check PIN for this specific user only.
    if (sanitize(r[U.PIN]) !== String(pin)) continue;

    const userRole = sanitize(r[U.ROLE]);
    if (submittedRole && submittedRole !== userRole) {
      await setPinAttemptState(ip, { count: (attemptState.count || 0) + 1 });
      return json({ success: false, role: null, wrongRole: true, attemptsRemaining: Math.max(0, PIN_MAX_ATTEMPTS - (attemptState.count + 1)) }, 401);
    }
    await setPinAttemptState(ip, { count: 0 });
    return json({
      success: true,
      role: userRole,
      userName: sanitize(r[U.NAME]),
      entityFilter: sanitize(r[U.ENTITY_FILTER]),
      navPages: sanitize(r[U.NAV_PAGES]),
    });
  }

  // --- 2. Fallback: Config sheet (legacy shared role PINs) ---
  // Only used when no userKey supplied — shared role logins without individual keys.
  // Unit heads excluded — must use Users sheet so their entity filter is enforced.
  if (!userKey) {
    const config = await getConfig(token, env);
    const roles = { hr_manager:"hr_manager", ceo:"ceo", recruiter:"recruiter" };
    let matchedRole = null;
    for (const [role] of Object.entries(roles)) {
      if (config[role] && config[role] === String(pin)) { matchedRole = role; break; }
    }
    if (matchedRole) {
      if (submittedRole && submittedRole !== matchedRole) {
        await setPinAttemptState(ip, { count: (attemptState.count || 0) + 1 });
        return json({ success: false, role: null, wrongRole: true, attemptsRemaining: Math.max(0, PIN_MAX_ATTEMPTS - (attemptState.count + 1)) }, 401);
      }
      await setPinAttemptState(ip, { count: 0 });
      return json({ success: true, role: matchedRole, userName: null, entityFilter: "", navPages: "" });
    }
  }

  await setPinAttemptState(ip, { count: (attemptState.count || 0) + 1 });
  return json({ success: false, role: null, attemptsRemaining: Math.max(0, PIN_MAX_ATTEMPTS - (attemptState.count + 1)) }, 401);
}

// GET /get-candidates  ?entity=&status=&position=&page=1&limit=50
async function handleGetCandidates(url, env, request) {
  const token  = await getToken(env);
  const rows   = await getRows(token, env, SHEET_CANDIDATES);
  const params = url.searchParams;
  const filterEntity   = params.get("entity")   || "";
  const filterStatus   = params.get("status")   || "";
  const filterPosition = params.get("position") || "";
  const filterSource   = params.get("source")   || "";
  const page   = parseInt(params.get("page")  || "1", 10);
  const limit  = parseInt(params.get("limit") || "50", 10);

  // SECURITY: Server-side entity restriction from the X-Entity-Filter header.
  // This header is set by the portal from the user's entityFilter field in the
  // Users sheet. Even if a unit head manually calls the API without ?entity=,
  // they will only ever get data for their allowed entities.
  // Format: comma-separated entity names, e.g. "APML" or "APML,Pine Pharmacy"
  // Empty = no restriction (HR Manager, Director).
  const serverEntityFilter = request?.headers?.get("X-Entity-Filter") || "";
  const allowedEntities = serverEntityFilter
    ? serverEntityFilter.split(",").map(s => s.trim()).filter(Boolean)
    : [];

  let data = rows.slice(1).filter(r => r[C.ID]).map(r => ({
    id:               sanitize(r[C.ID]),
    date:             excelDateToISO(r[C.DATE]),
    name:             sanitize(r[C.NAME]),
    phone:            sanitize(r[C.PHONE]),
    email:            sanitize(r[C.EMAIL]),
    position:         sanitize(r[C.POSITION]),
    entity:           sanitize(r[C.ENTITY]),
    degree:           sanitize(r[C.DEGREE]),
    status:           sanitize(r[C.STATUS]),
    driversLicense:   sanitize(r[C.DRIVERS_LICENSE]),
    emirate:          sanitize(r[C.EMIRATE]),
    noticePeriod:     sanitize(r[C.NOTICE_PERIOD]),
    uaeExp:           sanitize(r[C.UAE_EXP]),
    visaType:         sanitize(r[C.VISA_TYPE]),
    visaExpiry:       excelDateToISO(r[C.VISA_EXPIRY]),
    expectedSalary:   sanitize(r[C.EXPECTED_SALARY]),
    source:           sanitize(r[C.SOURCE]),
    sourceCost:       sanitize(r[C.SOURCE_COST]),
    availability:     sanitize(r[C.AVAILABILITY]),
    interviewStatus:  sanitize(r[C.INTERVIEW_STATUS]),
    interviewDate:    excelDateToISO(r[C.INTERVIEW_DATE]),
    remarks:          sanitize(r[C.REMARKS]),
    prevEmployer:     sanitize(r[C.PREV_EMPLOYER]),
    currEmployer:     sanitize(r[C.CURR_EMPLOYER]),
    relation:         sanitize(r[C.RELATION]),
    nationality:      sanitize(r[C.NATIONALITY]),
    location:         sanitize(r[C.LOCATION]),
    yearMonth:        sanitize(r[C.YEAR_MONTH]),
    cvLink:           sanitize(r[C.CV_LINK]),
    duplicateFlag:    sanitize(r[C.DUPLICATE_FLAG]),
    requisitionID:    sanitize(r[C.REQUISITION_ID]),
    requestedBy:      sanitize(r[C.REQUESTED_BY]),
    approvedBy:       sanitize(r[C.APPROVED_BY]),
    offerStage:       sanitize(r[C.OFFER_STAGE]),
    offeredSalary:    sanitize(r[C.OFFERED_SALARY]),
    declineReason:    sanitize(r[C.DECLINE_REASON]),
    expectedJoining:  excelDateToISO(r[C.EXPECTED_JOINING]),
    actualJoining:    excelDateToISO(r[C.ACTUAL_JOINING]),
    timeToHire:       sanitize(r[C.TIME_TO_HIRE]),
    cvdbId:           sanitize(r[C.CVDB_ID]),
    excludedRounds:     sanitize(r[C.EXCLUDED_ROUNDS]),
    excludedRaters:     sanitize(r[EXCLUDED_RATERS_COL] || ""),
    noticeDaysServed:   sanitize(r[C.NOTICE_DAYS_SERVED] || ""),
    shareToken:         sanitize(r[C.SHARE_TOKEN] || ""),
    aiRating:           sanitize(r[C.AI_RATING] || ""),
    aiRatingReason:     sanitize(r[C.AI_RATING_REASON] || ""),
  }));

  // Server-side entity enforcement — applied before any client-side filters
  if (allowedEntities.length > 0) data = data.filter(r => allowedEntities.includes(r.entity));

  if (filterEntity)   data = data.filter(r => r.entity   === filterEntity);
  if (filterStatus)   data = data.filter(r => r.status   === filterStatus);
  if (filterPosition) data = data.filter(r => r.position === filterPosition);
  if (filterSource)   data = data.filter(r => r.source   === filterSource);

  const total = data.length;
  const start = (page - 1) * limit;
  const paged = data.slice(start, start + limit);
  return json({ total, page, limit, candidates: paged });
}

// GET /get-candidate/:id
async function handleGetCandidate(candidateID, env) {
  const token = await getToken(env);
  const [candRows, roundRows] = await Promise.all([
    getRows(token, env, SHEET_CANDIDATES),
    getRows(token, env, SHEET_ROUNDS),
  ]);

  const candRow = candRows.slice(1).find(r => sanitize(r[C.ID]) === candidateID);
  if (!candRow) return err("Candidate not found", 404);

  const candidate = {
    id:               sanitize(candRow[C.ID]),
    date:             excelDateToISO(candRow[C.DATE]),
    name:             sanitize(candRow[C.NAME]),
    phone:            sanitize(candRow[C.PHONE]),
    email:            sanitize(candRow[C.EMAIL]),
    position:         sanitize(candRow[C.POSITION]),
    entity:           sanitize(candRow[C.ENTITY]),
    degree:           sanitize(candRow[C.DEGREE]),
    status:           sanitize(candRow[C.STATUS]),
    driversLicense:   sanitize(candRow[C.DRIVERS_LICENSE]),
    emirate:          sanitize(candRow[C.EMIRATE]),
    noticePeriod:     sanitize(candRow[C.NOTICE_PERIOD]),
    uaeExp:           sanitize(candRow[C.UAE_EXP]),
    visaType:         sanitize(candRow[C.VISA_TYPE]),
    visaExpiry:       excelDateToISO(candRow[C.VISA_EXPIRY]),
    expectedSalary:   sanitize(candRow[C.EXPECTED_SALARY]),
    source:           sanitize(candRow[C.SOURCE]),
    sourceCost:       sanitize(candRow[C.SOURCE_COST]),
    availability:     sanitize(candRow[C.AVAILABILITY]),
    interviewStatus:  sanitize(candRow[C.INTERVIEW_STATUS]),
    interviewDate:    excelDateToISO(candRow[C.INTERVIEW_DATE]),
    remarks:          sanitize(candRow[C.REMARKS]),
    prevEmployer:     sanitize(candRow[C.PREV_EMPLOYER]),
    currEmployer:     sanitize(candRow[C.CURR_EMPLOYER]),
    relation:         sanitize(candRow[C.RELATION]),
    nationality:      sanitize(candRow[C.NATIONALITY]),
    location:         sanitize(candRow[C.LOCATION]),
    yearMonth:        sanitize(candRow[C.YEAR_MONTH]),
    cvLink:           sanitize(candRow[C.CV_LINK]),
    duplicateFlag:    sanitize(candRow[C.DUPLICATE_FLAG]),
    requisitionID:    sanitize(candRow[C.REQUISITION_ID]),
    requestedBy:      sanitize(candRow[C.REQUESTED_BY]),
    approvedBy:       sanitize(candRow[C.APPROVED_BY]),
    offerStage:       sanitize(candRow[C.OFFER_STAGE]),
    offeredSalary:    sanitize(candRow[C.OFFERED_SALARY]),
    declineReason:    sanitize(candRow[C.DECLINE_REASON]),
    expectedJoining:  excelDateToISO(candRow[C.EXPECTED_JOINING]),
    actualJoining:    excelDateToISO(candRow[C.ACTUAL_JOINING]),
    timeToHire:       sanitize(candRow[C.TIME_TO_HIRE]),
    cvdbId:           sanitize(candRow[C.CVDB_ID]),
    excludedRounds:   sanitize(candRow[C.EXCLUDED_ROUNDS]),
  };

  const rounds = roundRows.slice(1)
    .filter(r => sanitize(r[R.CAND_ID]) === candidateID)
    .map(r => ({
      candidateID:      sanitize(r[R.CAND_ID]),
      round:            sanitize(r[R.ROUND]),
      interviewer:      sanitize(r[R.INTERVIEWER]),
      scheduledDate:    excelDateToISO(r[R.SCHED_DATE]),
      actualDate:       excelDateToISO(r[R.ACTUAL_DATE]),
      outcome:          sanitize(r[R.OUTCOME]),
      rescheduleCount:  sanitize(r[R.RESCHEDULE_COUNT]),
      score:            sanitize(r[R.SCORE]),
      feedback:         sanitize(r[R.FEEDBACK]),
      ratingToken:      sanitize(r[R.RATING_TOKEN]),
      ratingExpires:    sanitize(r[R.RATING_TOKEN_EXPIRES]),
      ratingSubmittedAt: sanitize(r[R.RATING_SUBMITTED_AT]),
      education:        sanitize(r[R.SCORE_EDUCATION]),
      workExp:          sanitize(r[R.SCORE_WORK_EXP]),
      technical:        sanitize(r[R.SCORE_TECHNICAL]),
      communication:    sanitize(r[R.SCORE_COMMUNICATION]),
      enthusiasm:       sanitize(r[R.SCORE_ENTHUSIASM]),
      product:          sanitize(r[R.SCORE_PRODUCT]),
      teamwork:         sanitize(r[R.SCORE_TEAMWORK]),
      initiative:       sanitize(r[R.SCORE_INITIATIVE]),
      timeMgmt:         sanitize(r[R.SCORE_TIME_MGMT]),
      companyKnowledge: sanitize(r[R.SCORE_COMPANY_KNOWLEDGE]),
      emotionalSensitivity: sanitize(r[R.SCORE_EMOTIONAL_SENSITIVITY]),
      duration: sanitize(r[R.INTERVIEW_DURATION]),
      omittedCriteria:  sanitize(r[R.OMITTED_CRITERIA]),
      ratingPassFail:   sanitize(r[R.RATING_PASS_FAIL]),
      ratingRemarks:    sanitize(r[R.RATING_REMARKS]),
      submittedByName:  sanitize(r[R.SUBMITTED_BY_NAME]),
      qShiftDuties:     sanitize(r[R.Q_SHIFT_DUTIES]),
      qDisciplinary:    sanitize(r[R.Q_DISCIPLINARY]),
      qRelative:        sanitize(r[R.Q_RELATIVE]),
      qRelativeDetail:  sanitize(r[R.Q_RELATIVE_DETAIL]),
      qCompetitor:      sanitize(r[R.Q_COMPETITOR]),
      qCompetitorDetail: sanitize(r[R.Q_COMPETITOR_DETAIL]),
    }));

  return json({ candidate, rounds });
}

// POST /add-candidate
async function handleAddCandidate(body, env) {
  const { name, phone, email, position, entity, role } = body;
  const cvdbSourceID = sanitize(body.cvdbSourceID || "");
  if (!name || !phone || !email || !position || !entity) {
    return err("name, phone, email, position, entity are required");
  }

  const token = await getToken(env);

  // Duplicate check
  const dupCheck = await checkDuplicate(token, env, name, phone, email);
  const dupFlag  = dupCheck.duplicate ? "Yes" : "No";

  const candID   = await nextCandidateID(token, env);
  const today    = isoDate();
  const ym       = yearMonth();
  const sourceCV = sanitize(body.source || "");

  const row = new Array(39).fill("");
  row[C.ID]               = candID;
  row[C.DATE]             = today;
  row[C.NAME]             = sanitize(name);
  row[C.PHONE]            = sanitize(phone);
  row[C.EMAIL]            = sanitize(email);
  row[C.POSITION]         = sanitize(position);
  row[C.ENTITY]           = sanitize(entity);
  row[C.DEGREE]           = sanitize(body.degree || "");
  row[C.STATUS]           = "Active";
  row[C.DRIVERS_LICENSE]  = sanitize(body.driversLicense || "");
  row[C.EMIRATE]          = sanitize(body.emirate || "");
  row[C.NOTICE_PERIOD]    = sanitize(body.noticePeriod || "");
  row[C.UAE_EXP]          = sanitize(body.uaeExp || "");
  row[C.VISA_TYPE]        = sanitize(body.visaType || "");
  row[C.VISA_EXPIRY]      = sanitize(body.visaExpiry || "");
  row[C.EXPECTED_SALARY]  = sanitize(body.expectedSalary || "");
  row[C.SOURCE]           = sourceCV;
  row[C.SOURCE_COST]      = sanitize(body.sourceCost || "");
  row[C.AVAILABILITY]     = sanitize(body.availability || "");
  row[C.INTERVIEW_STATUS] = "";
  row[C.INTERVIEW_DATE]   = "";
  row[C.REMARKS]          = sanitize(body.remarks || "");
  row[C.PREV_EMPLOYER]    = sanitize(body.prevEmployer || "");
  row[C.CURR_EMPLOYER]    = sanitize(body.currEmployer || "");
  row[C.RELATION]         = sanitize(body.relation || "");
  row[C.NATIONALITY]      = sanitize(body.nationality || "");
  row[C.LOCATION]         = sanitize(body.location || "");
  row[C.YEAR_MONTH]       = ym;
  row[C.CV_LINK]          = sanitize(body.cvLink || "");
  row[C.DUPLICATE_FLAG]   = dupFlag;
  row[C.REQUISITION_ID]   = sanitize(body.requisitionID || "");
  row[C.REQUESTED_BY]     = sanitize(body.requestedBy || "");
  row[C.APPROVED_BY]      = sanitize(body.approvedBy || "");
  row[C.OFFER_STAGE]      = "";
  row[C.OFFERED_SALARY]   = "";
  row[C.DECLINE_REASON]   = "";
  row[C.EXPECTED_JOINING] = sanitize(body.expectedJoining || "");
  row[C.ACTUAL_JOINING]   = "";
  row[C.TIME_TO_HIRE]     = "";

  await appendRow(token, env, SHEET_CANDIDATES, row);

  // If this candidate was sourced from the CV Database (cvdbSourceID present),
  // sync the corrected/supplemented data back to the CVDB entry and record the
  // new candidateID as the linked record.
  if (cvdbSourceID) {
    try {
      const cvdbRows = await getRows(token, env, SHEET_CVDB);
      for (let i = 1; i < cvdbRows.length; i++) {
        if (sanitize(cvdbRows[i][CVDB.ID]) === cvdbSourceID) {
          const updRow = [...cvdbRows[i]];
          // Sync any corrections to core fields
          updRow[CVDB.NAME]          = sanitize(name);
          updRow[CVDB.PHONE]         = sanitize(phone);
          updRow[CVDB.EMAIL]         = sanitize(email);
          updRow[CVDB.FACILITY]      = sanitize(entity);
          // Link CVDB entry to the new candidate record
          while (updRow.length <= CVDB.LINKED_CAND_ID) updRow.push("");
          updRow[CVDB.LINKED_CAND_ID] = candID;
          await updateRow(token, env, SHEET_CVDB, i + 1, updRow);
          break;
        }
      }
    } catch(e) {
      console.warn("CVDB sync-back failed (non-fatal):", e.message);
    }
  }

  // WhatsApp notify HR on new submission
  const msg = `📋 New CV Submitted\nRef: ${candID}\nName: ${name}\nPosition: ${position}\nEntity: ${entity}\nSource: ${sourceCV}${dupFlag === "Yes" ? "\n⚠️ DUPLICATE FLAG" : ""}`;
  await notifyWhatsApp(env, msg);

  return json({
    success:       true,
    candidateID:   candID,
    duplicateFlag: dupFlag,
    existingID:    dupCheck.existingID || null,
    duplicateMatches: dupCheck.matches || [],
    message:       `Submission successful. Your reference: ${candID}`,
  });
}

// ═══════════════════════════════════════════════════════════════════════
// PIPELINE STAGE HANDLERS
// ═══════════════════════════════════════════════════════════════════════

const STAGE_STATUS_MAP = {
  "HR Screen":     { Pass:"Shortlisted",   Fail:"Rejected" },
  "Operations":    { Pass:"Ops Cleared",   Fail:"Rejected" },
  "Management":    { Pass:"Mgmt Cleared",  Fail:"Rejected" },
  "Verbal Offer":  { Pass:"Verbal Offered",Fail:null },
  "Written Offer": { Pass:"Offered",       Fail:null },
  "Accepted":      { Pass:"Accepted",      Fail:null },
  "Declined":      { Pass:null,            Fail:"Rejected" },
  "Onboarding":    { Pass:"Onboarding",    Fail:null },
  "Joined":        { Pass:"Joined",        Fail:null },
};

// GET /get-pipeline?reqID=&entity=&status=
async function handleGetPipeline(url, env) {
  const token      = await getToken(env);
  const reqFilter  = url.searchParams.get("reqID")  || "";
  const entFilter  = url.searchParams.get("entity") || "";

  const [candRows, psRows, roundRows] = await Promise.all([
    getRows(token, env, SHEET_CANDIDATES),
    getRows(token, env, SHEET_PIPELINE).catch(() => []),
    getRows(token, env, SHEET_ROUNDS).catch(() => []),
  ]);

  // Build interview history per candidate
  const interviewMap = {};
  for (const r of roundRows.slice(1)) {
    const cid = sanitize(r[R.CAND_ID]);
    if (!cid) continue;
    if (!interviewMap[cid]) interviewMap[cid] = [];
    interviewMap[cid].push({
      round:           sanitize(r[R.ROUND]),
      interviewer:     sanitize(r[R.INTERVIEWER]),
      interviewerRole: sanitize(r[R.INTERVIEWER_ROLE]),
      team:            sanitize(r[R.TEAM]),
      scheduledDate:   excelDateToISO(r[R.SCHED_DATE]),
      scheduledTime:   excelTimeToHHMM(r[R.SCHED_TIME]),
      actualDate:      excelDateToISO(r[R.ACTUAL_DATE]),
      outcome:         sanitize(r[R.OUTCOME]),
      status:          sanitize(r[R.STATUS]),
      rescheduleCount: sanitize(r[R.RESCHEDULE_COUNT]),
      score:           sanitize(r[R.SCORE]),
      feedback:        sanitize(r[R.FEEDBACK]),
      ratingToken:     sanitize(r[R.RATING_TOKEN]),
      ratingExpires:   sanitize(r[R.RATING_TOKEN_EXPIRES]),
      ratingSubmittedAt: sanitize(r[R.RATING_SUBMITTED_AT]),
      education:        sanitize(r[R.SCORE_EDUCATION]),
      workExp:          sanitize(r[R.SCORE_WORK_EXP]),
      technical:        sanitize(r[R.SCORE_TECHNICAL]),
      communication:    sanitize(r[R.SCORE_COMMUNICATION]),
      enthusiasm:       sanitize(r[R.SCORE_ENTHUSIASM]),
      product:          sanitize(r[R.SCORE_PRODUCT]),
      teamwork:         sanitize(r[R.SCORE_TEAMWORK]),
      initiative:       sanitize(r[R.SCORE_INITIATIVE]),
      timeMgmt:         sanitize(r[R.SCORE_TIME_MGMT]),
      companyKnowledge: sanitize(r[R.SCORE_COMPANY_KNOWLEDGE]),
      emotionalSensitivity: sanitize(r[R.SCORE_EMOTIONAL_SENSITIVITY]),
      duration: sanitize(r[R.INTERVIEW_DURATION]),
      omittedCriteria:  sanitize(r[R.OMITTED_CRITERIA]),
      ratingPassFail:   sanitize(r[R.RATING_PASS_FAIL]),
      ratingRemarks:    sanitize(r[R.RATING_REMARKS]),
      submittedByName:  sanitize(r[R.SUBMITTED_BY_NAME]),
      qShiftDuties:     sanitize(r[R.Q_SHIFT_DUTIES]),
      qDisciplinary:    sanitize(r[R.Q_DISCIPLINARY]),
      qRelative:        sanitize(r[R.Q_RELATIVE]),
      qRelativeDetail:  sanitize(r[R.Q_RELATIVE_DETAIL]),
      qCompetitor:      sanitize(r[R.Q_COMPETITOR]),
      qCompetitorDetail: sanitize(r[R.Q_COMPETITOR_DETAIL]),
      raterID:          sanitize(r[R.RATER_ID]),
      shortCode:        sanitize(r[R.SHORT_CODE]),
    });
  }

  // Build stage history per candidate
  const stageMap = {};
  for (const r of psRows.slice(1)) {
    const cid = sanitize(r[PS.ID]);
    if (!cid) continue;
    if (!stageMap[cid]) stageMap[cid] = [];
    stageMap[cid].push({
      stage:          sanitize(r[PS.STAGE]),
      date:           excelDateToISO(r[PS.DATE]),
      interviewer:    sanitize(r[PS.INTERVIEWER]),
      outcome:        sanitize(r[PS.OUTCOME]),
      notes:          sanitize(r[PS.NOTES]),
      score:          sanitize(r[PS.SCORE]),
      salary:         sanitize(r[PS.SALARY]),
      joiningDate:    excelDateToISO(r[PS.JOINING_DATE]),
      visa:           sanitize(r[PS.VISA]),
      senior:         sanitize(r[PS.SENIOR]),
      declineReason:  sanitize(r[PS.DECLINE_REASON]),
      onboardingJson: sanitize(r[PS.ONBOARDING_JSON]),
    });
  }

  const PIPELINE_STATUSES = ["Active","Shortlisted","Ops Cleared","Mgmt Cleared",
    "Verbal Offered","Offered","Accepted","Onboarding","Joined","Rejected"];

  let candidates = candRows.slice(1).filter(r => r[C.ID]).map(r => {
    const id = sanitize(r[C.ID]);
    const candStages = stageMap[id] || [];
    // For rejected candidates, surface the reason from the stage that failed/declined them
    const rejectStage = [...candStages].reverse().find(s => s.outcome === "Fail" || s.outcome === "Declined" || s.outcome === "Rejected");
    return {
      id, name: sanitize(r[C.NAME]), position: sanitize(r[C.POSITION]),
      entity: sanitize(r[C.ENTITY]), status: sanitize(r[C.STATUS]),
      requisitionID: sanitize(r[C.REQUISITION_ID]),
      expectedSalary: sanitize(r[C.EXPECTED_SALARY]),
      nationality: sanitize(r[C.NATIONALITY]),
      date: excelDateToISO(r[C.DATE]),
      cvLink: sanitize(r[C.CV_LINK]),
      stages: candStages,
      interviews: interviewMap[id] || [],
      rejectReason: rejectStage ? (rejectStage.declineReason || rejectStage.notes || "") : "",
      rejectedAtStage: rejectStage ? rejectStage.stage : "",
      rejectedDate: rejectStage ? rejectStage.date : "",
      // Full profile fields — used by the candidate detail view and WhatsApp export
      phone:           sanitize(r[C.PHONE]),
      email:           sanitize(r[C.EMAIL]),
      degree:          sanitize(r[C.DEGREE]),
      driversLicense:  sanitize(r[C.DRIVERS_LICENSE]),
      emirate:         sanitize(r[C.EMIRATE]),
      noticePeriod:    sanitize(r[C.NOTICE_PERIOD]),
      uaeExp:          sanitize(r[C.UAE_EXP]),
      visaType:        sanitize(r[C.VISA_TYPE]),
      visaExpiry:      excelDateToISO(r[C.VISA_EXPIRY]),
      source:          sanitize(r[C.SOURCE]),
      sourceCost:      sanitize(r[C.SOURCE_COST]),
      availability:    sanitize(r[C.AVAILABILITY]),
      interviewStatus: sanitize(r[C.INTERVIEW_STATUS]),
      interviewDate:   excelDateToISO(r[C.INTERVIEW_DATE]),
      remarks:         sanitize(r[C.REMARKS]),
      prevEmployer:    sanitize(r[C.PREV_EMPLOYER]),
      currEmployer:    sanitize(r[C.CURR_EMPLOYER]),
      relation:        sanitize(r[C.RELATION]),
      location:        sanitize(r[C.LOCATION]),
      yearMonth:       sanitize(r[C.YEAR_MONTH]),
    };
  }).filter(c => PIPELINE_STATUSES.includes(c.status));

  if (reqFilter) candidates = candidates.filter(c => c.requisitionID === reqFilter);
  if (entFilter) candidates = candidates.filter(c => c.entity === entFilter);

  const STAGE_COLUMN = {
    "Active":"hr_screen","Shortlisted":"shortlisted","Ops Cleared":"ops_cleared",
    "Mgmt Cleared":"mgmt_cleared","Verbal Offered":"offer","Offered":"offer",
    "Accepted":"offer","Onboarding":"onboarding","Joined":"joined",
  };
  const COLUMNS = [
    { id:"hr_screen",   label:"HR Screen" },
    { id:"shortlisted", label:"Shortlisted" },
    { id:"ops_cleared", label:"Ops Cleared" },
    { id:"mgmt_cleared",label:"Mgmt Cleared" },
    { id:"offer",       label:"Offer" },
    { id:"onboarding",  label:"Onboarding" },
    { id:"joined",      label:"Joined ✓" },
  ];

  const kanban = {};
  COLUMNS.forEach(c => { kanban[c.id] = []; });
  candidates.forEach(c => {
    const col = STAGE_COLUMN[c.status];
    if (col) kanban[col].push(c);
  });

  return json({ candidates, kanban, columns: COLUMNS });
}

// POST /add-pipeline-stage
async function handleAddPipelineStage(body, env) {
  const { candidateID, stage, date, interviewer, outcome, notes,
          score, salary, joiningDate, visa, senior, declineReason, onboardingJson } = body;
  if (!candidateID || !stage) return err("candidateID and stage required");

  const token = await getToken(env);
  const row   = new Array(14).fill("");
  row[PS.ID]              = candidateID;
  row[PS.STAGE]           = sanitize(stage);
  row[PS.DATE]            = sanitize(date || isoDate());
  row[PS.INTERVIEWER]     = sanitize(interviewer || "");
  row[PS.OUTCOME]         = sanitize(outcome || "Pending");
  row[PS.NOTES]           = sanitize(notes || "");
  row[PS.SCORE]           = sanitize(score || "");
  row[PS.SALARY]          = sanitize(salary || "");
  row[PS.JOINING_DATE]    = sanitize(joiningDate || "");
  row[PS.VISA]            = sanitize(visa || "");
  row[PS.SENIOR]          = sanitize(senior || "");
  row[PS.DECLINE_REASON]  = sanitize(declineReason || "");
  row[PS.ONBOARDING_JSON] = sanitize(onboardingJson || "");
  row[PS.CREATED_AT]      = isoDate();
  await appendRow(token, env, SHEET_PIPELINE, row);

  // ── Record in Pipeline_History for bottleneck analytics ──────────────
  // We need the candidate's name/position/entity — fetch lazily only if needed.
  // Use fire-and-forget (no await) so it never delays the response.
  getRows(token, env, SHEET_CANDIDATES).then(candRows => {
    const cr = candRows.slice(1).find(r => sanitize(r[C.ID]) === candidateID);
    const fromStage = cr ? sanitize(cr[C.STATUS]) : "";
    appendPipelineHistory(token, env, {
      candID:      candidateID,
      candName:    cr ? sanitize(cr[C.NAME])     : "",
      position:    cr ? (sanitize(cr[C.CUSTOM_POSITION]) || sanitize(cr[C.POSITION])) : "",
      entity:      cr ? sanitize(cr[C.ENTITY])   : "",
      fromStage,
      toStage:     stage,
      outcome:     outcome || "Pending",
      advancedBy:  body.advancedBy || "",
    });
    // Fire notification (fire-and-forget)
    fireNotify(env, "stage_advance", {
      candName: cr ? sanitize(cr[C.NAME]) : candidateID,
      position: cr ? (sanitize(cr[C.CUSTOM_POSITION]) || sanitize(cr[C.POSITION])) : "",
      entity:   cr ? sanitize(cr[C.ENTITY]) : "",
      stage,
      outcome:  outcome || "Pending",
    });
  }).catch(() => {});

  // Derive new candidate status
  const sm = STAGE_STATUS_MAP[stage];
  let newStatus = null;
  if (sm) {
    if (outcome === "Pass" && sm.Pass) newStatus = sm.Pass;
    if (outcome === "Fail" && sm.Fail) newStatus = sm.Fail;
  }
  if (stage === "Verbal Offer"  && outcome === "Pass") newStatus = "Verbal Offered";
  if (stage === "Written Offer" && outcome === "Pass") newStatus = "Offered";
  if (stage === "Accepted")                            newStatus = "Accepted";
  if (stage === "Onboarding")                          newStatus = "Onboarding";
  if (stage === "Joined")                              newStatus = "Joined";
  if (outcome === "Declined")                          newStatus = "Rejected";

  if (newStatus) {
    const candRows = await getRows(token, env, SHEET_CANDIDATES);
    for (let i = 1; i < candRows.length; i++) {
      if (sanitize(candRows[i][C.ID]) !== candidateID) continue;
      await updateCell(token, env, SHEET_CANDIDATES, i+1, C.STATUS, newStatus);
      if (newStatus === "Joined") {
        const finalJoiningDate = joiningDate || new Date().toISOString().split("T")[0];
        await updateCell(token, env, SHEET_CANDIDATES, i+1, C.ACTUAL_JOINING, finalJoiningDate);
        // excelDateToISO handles both string dates and Excel serial numbers,
        // so this stays correct even for older rows where C.DATE might not
        // be a clean "YYYY-MM-DD" string.
        const entryDateISO = excelDateToISO(candRows[i][C.DATE]) || sanitize(candRows[i][C.DATE]);
        const diff = Math.round((new Date(finalJoiningDate) - new Date(entryDateISO)) / 86400000);
        if (!isNaN(diff) && diff >= 0) await updateCell(token, env, SHEET_CANDIDATES, i+1, C.TIME_TO_HIRE, diff);
        if (body.noticeDaysServed) {
          const nd = parseInt(body.noticeDaysServed);
          if (!isNaN(nd) && nd > 0) await updateCell(token, env, SHEET_CANDIDATES, i+1, C.NOTICE_DAYS_SERVED, nd);
        }
        // Final confirmed salary at joining — this is the figure that actually
        // feeds budget-vs-actual tracking in Hiring Analytics. Writing it here
        // guarantees it's captured even if an earlier offer stage was skipped
        // or its value never made it into OFFERED_SALARY.
        if (salary) await updateCell(token, env, SHEET_CANDIDATES, i+1, C.OFFERED_SALARY, salary);
        const reqID = sanitize(candRows[i][C.REQUISITION_ID]);
        if (reqID) await handleCloseRequisition({ requisitionID: reqID }, env);
      }
      if (salary && (stage === "Mgmt Cleared" || stage === "Written Offer" || stage === "Verbal Offer")) {
        await updateCell(token, env, SHEET_CANDIDATES, i+1, C.OFFERED_SALARY, salary);
      }
      break;
    }
    // Sync Joined/Rejected status back to the linked CVDB entry so the
    // CV Database shows the correct badge without needing a separate update.
    if (newStatus === "Joined" || newStatus === "Rejected") {
      try {
        const candRow = candRows.slice(1).find(r => sanitize(r[C.ID]) === candidateID);
        const cvdbID  = candRow ? sanitize(candRow[C.CVDB_ID]) : "";
        if (cvdbID) {
          const cvdbRows = await getRows(token, env, SHEET_CVDB);
          for (let i = 1; i < cvdbRows.length; i++) {
            if (sanitize(cvdbRows[i][CVDB.ID]) === cvdbID) {
              const updRow = [...cvdbRows[i]];
              while (updRow.length <= CVDB.LINKED_CAND_ID) updRow.push("");
              updRow[CVDB.LINKED_CAND_ID] = candidateID; // ensure link is set
              await updateRow(token, env, SHEET_CVDB, i + 1, updRow);
              break;
            }
          }
        }
      } catch(e) { console.warn("CVDB status sync failed (non-fatal):", e.message); }
    }
    const NOTIFY = ["Mgmt Cleared","Offered","Accepted","Joined","Rejected"];
    if (NOTIFY.includes(newStatus)) {
      const cr   = candRows.slice(1).find(r => sanitize(r[C.ID]) === candidateID);
      const name = cr ? sanitize(cr[C.NAME]) : candidateID;
      const emo  = {"Mgmt Cleared":"✅","Offered":"📄","Accepted":"🎉","Joined":"🟢","Rejected":"❌"}[newStatus]||"📋";
      await notifyWhatsApp(env, `${emo} ${name} — ${newStatus}\n${stage} | ${outcome}`);
    }
  }
  return json({ success:true, candidateID, stage, outcome, newStatus });
}

// POST /update-pipeline-stage
async function handleUpdatePipelineStage(body, env) {
  const { candidateID, stage, updates } = body;
  if (!candidateID || !stage) return err("candidateID and stage required");
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_PIPELINE);
  let excelRow = -1;
  for (let i = rows.length-1; i >= 1; i--) {
    if (sanitize(rows[i][PS.ID]) === candidateID && sanitize(rows[i][PS.STAGE]) === stage) {
      excelRow = i+1; break;
    }
  }
  if (excelRow === -1) return err("Stage not found",404);
  const fm = { outcome:PS.OUTCOME, notes:PS.NOTES, score:PS.SCORE, salary:PS.SALARY,
    joiningDate:PS.JOINING_DATE, declineReason:PS.DECLINE_REASON, onboardingJson:PS.ONBOARDING_JSON,
    date:PS.DATE, interviewer:PS.INTERVIEWER };
  for (const [f,ci] of Object.entries(fm)) {
    if (updates[f] !== undefined) await updateCell(token, env, SHEET_PIPELINE, excelRow, ci, String(updates[f]));
  }
  return json({ success:true, candidateID, stage });
}

// POST /handle-decline
// POST /reject-candidate — HR/Recruiter rejects a candidate at any stage, reason required.
// Distinct from handleDecline (candidate declining HiFive's offer).
async function handleRejectCandidate(body, env) {
  const { candidateID, reason, rejectedBy } = body;
  if (!candidateID) return err("candidateID required");
  if (!reason || !reason.trim()) return err("A reason is required to reject a candidate");

  const token = await getToken(env);
  const candRows = await getRows(token, env, SHEET_CANDIDATES);
  let excelRow = -1, existingRow = null, prevStatus = "";
  for (let i = 1; i < candRows.length; i++) {
    if (sanitize(candRows[i][C.ID]) === candidateID) {
      excelRow = i + 1;
      existingRow = candRows[i];
      prevStatus = sanitize(candRows[i][C.STATUS]);
      break;
    }
  }
  if (excelRow === -1) return err("Candidate not found", 404);

  await updateCell(token, env, SHEET_CANDIDATES, excelRow, C.STATUS, "Rejected");

  // Log a PipelineStages entry capturing the reason and the stage they were rejected from,
  // so the rejected-candidates list and revert-to-stage feature have a full record.
  const psRow = new Array(14).fill("");
  psRow[PS.ID]             = candidateID;
  psRow[PS.STAGE]          = "Rejected";
  psRow[PS.DATE]           = isoDate();
  psRow[PS.OUTCOME]        = "Rejected";
  psRow[PS.NOTES]          = `Rejected from "${prevStatus}" by ${sanitize(rejectedBy || "HR")}: ${sanitize(reason)}`;
  psRow[PS.DECLINE_REASON] = sanitize(reason);
  await appendRow(token, env, SHEET_PIPELINE, psRow);

  // Record in Pipeline_History
  appendPipelineHistory(token, env, {
    candID:    candidateID,
    candName:  sanitize(existingRow[C.NAME]),
    position:  sanitize(existingRow[C.CUSTOM_POSITION]) || sanitize(existingRow[C.POSITION]),
    entity:    sanitize(existingRow[C.ENTITY]),
    fromStage: prevStatus,
    toStage:   "Rejected",
    outcome:   "Rejected",
    advancedBy: sanitize(rejectedBy || "HR"),
  }).catch(() => {});

  const msg = `❌ Candidate Rejected\nName: ${sanitize(existingRow[C.NAME])}\nRejected from: ${prevStatus}\nReason: ${reason}\nBy: ${rejectedBy || "HR"}`;
  await notifyWhatsApp(env, msg);

  return json({ success: true, candidateID, previousStatus: prevStatus });
}

// POST /revert-candidate — restores a rejected/declined candidate to the stage
// they were rejected from, so they can continue through the journey normally.
async function handleRevertCandidate(body, env) {
  const { candidateID, revertedBy } = body;
  if (!candidateID) return err("candidateID required");

  const token = await getToken(env);
  const [candRows, psRows] = await Promise.all([
    getRows(token, env, SHEET_CANDIDATES),
    getRows(token, env, SHEET_PIPELINE).catch(() => []),
  ]);

  let excelRow = -1;
  for (let i = 1; i < candRows.length; i++) {
    if (sanitize(candRows[i][C.ID]) === candidateID) { excelRow = i + 1; break; }
  }
  if (excelRow === -1) return err("Candidate not found", 404);

  // Find the most recent rejection entry for this candidate and look at the
  // stage they were in immediately before it, so they resume from there.
  const candStages = psRows.slice(1).filter(r => sanitize(r[PS.ID]) === candidateID);
  const rejectIdxFromEnd = [...candStages].reverse().findIndex(r => sanitize(r[PS.OUTCOME]) === "Rejected" || sanitize(r[PS.OUTCOME]) === "Declined");
  let restoreStatus = "Active"; // safe fallback: back to HR Screen
  if (rejectIdxFromEnd !== -1) {
    const rejectIdx = candStages.length - 1 - rejectIdxFromEnd;
    const priorEntry = candStages[rejectIdx - 1];
    const STAGE_TO_STATUS = {
      "HR Screen":"Active", "Operations":"Ops Cleared", "Management":"Mgmt Cleared",
      "Verbal Offer":"Verbal Offered", "Written Offer":"Offered", "Accepted":"Accepted",
      "Onboarding":"Onboarding",
    };
    restoreStatus = priorEntry ? (STAGE_TO_STATUS[sanitize(priorEntry[PS.STAGE])] || "Active") : "Active";
  }

  await updateCell(token, env, SHEET_CANDIDATES, excelRow, C.STATUS, restoreStatus);

  const msg = `↩️ Candidate Reverted\nID: ${candidateID}\nRestored to: ${restoreStatus}\nBy: ${revertedBy || "HR"}`;
  await notifyWhatsApp(env, msg);

  return json({ success: true, candidateID, restoredStatus: restoreStatus });
}

async function handleDecline(body, env) {
  const { candidateID, action, requisitionID } = body;
  if (!candidateID || !action) return err("candidateID and action required");
  const token = await getToken(env);
  const candRows = await getRows(token, env, SHEET_CANDIDATES);
  for (let i=1; i<candRows.length; i++) {
    if (sanitize(candRows[i][C.ID]) === candidateID) {
      await updateCell(token, env, SHEET_CANDIDATES, i+1, C.STATUS, "Rejected");
      await updateCell(token, env, SHEET_CANDIDATES, i+1, C.OFFER_STAGE, "Declined");
      break;
    }
  }
  if (requisitionID) {
    const reqRows = await getRows(token, env, SHEET_REQUISITIONS);
    for (let i=1; i<reqRows.length; i++) {
      if (sanitize(reqRows[i][REQ.ID]) !== requisitionID) continue;
      if (action === "reopen") {
        const cur = parseInt(sanitize(reqRows[i][REQ.FILLED_COUNT])||"0",10);
        await updateCell(token, env, SHEET_REQUISITIONS, i+1, REQ.FILLED_COUNT, Math.max(0,cur-1));
        await updateCell(token, env, SHEET_REQUISITIONS, i+1, REQ.OVERALL_STATUS, "Approved");
        await updateCell(token, env, SHEET_REQUISITIONS, i+1, REQ.CLOSED_DATE, "");
        await notifyWhatsApp(env, `🔄 Offer declined — ${requisitionID} reopened`);
      } else if (action === "close") {
        await updateCell(token, env, SHEET_REQUISITIONS, i+1, REQ.OVERALL_STATUS, "Closed — Cancelled");
      }
      break;
    }
  }
  return json({ success:true, candidateID, action });
}

// POST /close-requisition — defined early so handleUpdateCandidate can call it
async function handleCloseRequisition(body, env) {
  const { requisitionID } = body;
  if (!requisitionID) return err("requisitionID required");
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_REQUISITIONS);
  let excelRow = -1;
  let existing = null;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][REQ.ID]) === requisitionID) {
      excelRow = i + 1; existing = rows[i]; break;
    }
  }
  if (excelRow === -1) return err("Requisition not found", 404);
  const candRows = await getRows(token, env, SHEET_CANDIDATES);
  const joined   = candRows.slice(1).filter(r =>
    sanitize(r[C.REQUISITION_ID]) === requisitionID &&
    sanitize(r[C.STATUS]) === "Joined"
  ).length;
  const headcount = parseInt(sanitize(existing[REQ.HEADCOUNT]) || "1", 10);
  const isFilled  = joined >= headcount;
  const newRow = [...existing];
  while (newRow.length < 35) newRow.push("");
  newRow[REQ.FILLED_COUNT]   = joined;
  newRow[REQ.OVERALL_STATUS] = isFilled ? "Closed — Filled" : sanitize(existing[REQ.OVERALL_STATUS]);
  newRow[REQ.CLOSED_DATE]    = isFilled ? isoDate() : "";
  await updateRow(token, env, SHEET_REQUISITIONS, excelRow, newRow);
  return json({ success: true, requisitionID, filledCount: joined, isFilled });
}

// POST /withdraw-requisition — soft-delete with mandatory justification.
// Does NOT remove the row — sets status to "Closed — Withdrawn" and records
// who withdrew it and why, preserving a full audit trail.
async function handleWithdrawRequisition(body, env) {
  const { requisitionID, withdrawnBy, justification } = body;
  if (!requisitionID)  return err("requisitionID is required");
  if (!withdrawnBy)    return err("withdrawnBy is required");
  if (!justification || !justification.trim()) return err("A justification is required to withdraw a requisition");

  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_REQUISITIONS);
  let excelRow = -1;
  let existing = null;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][REQ.ID]) === requisitionID) {
      excelRow = i + 1; existing = rows[i]; break;
    }
  }
  if (excelRow === -1) return err("Requisition not found", 404);

  const newRow = [...existing];
  while (newRow.length < 35) newRow.push("");
  newRow[REQ.OVERALL_STATUS] = "Closed — Withdrawn";
  newRow[REQ.CLOSED_DATE]    = isoDate();
  newRow[REQ.HR_REMARKS]     = `${sanitize(existing[REQ.HR_REMARKS]) ? sanitize(existing[REQ.HR_REMARKS]) + " | " : ""}Withdrawn by ${sanitize(withdrawnBy)} on ${isoDate()}: ${sanitize(justification)}`;
  await updateRow(token, env, SHEET_REQUISITIONS, excelRow, newRow);

  const msg = `🗑️ Requisition Withdrawn\nRef: ${requisitionID}\nPosition: ${sanitize(existing[REQ.POSITION])}\nEntity: ${sanitize(existing[REQ.ENTITY])}\nBy: ${withdrawnBy}\nReason: ${justification}`;
  await notifyWhatsApp(env, msg);

  return json({ success: true, requisitionID });
}

// POST /update-candidate
async function handleUpdateCandidate(body, env) {
  const { candidateID, updates } = body;
  if (!candidateID || !updates) return err("candidateID and updates required");

  const token    = await getToken(env);
  const rows     = await getRows(token, env, SHEET_CANDIDATES);
  let excelRow   = -1;
  let existingRow = null;

  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][C.ID]) === candidateID) {
      excelRow    = i + 1; // Excel is 1-indexed, header is row 1
      existingRow = rows[i];
      break;
    }
  }
  if (excelRow === -1) return err("Candidate not found", 404);

  // Merge updates onto existing row
  const newRow = [...existingRow];
  while (newRow.length < 39) newRow.push("");

  const fieldMap = {
    name: C.NAME, phone: C.PHONE, email: C.EMAIL, position: C.POSITION,
    entity: C.ENTITY, degree: C.DEGREE, status: C.STATUS,
    driversLicense: C.DRIVERS_LICENSE, emirate: C.EMIRATE,
    noticePeriod: C.NOTICE_PERIOD, uaeExp: C.UAE_EXP,
    visaType: C.VISA_TYPE, visaExpiry: C.VISA_EXPIRY,
    expectedSalary: C.EXPECTED_SALARY, source: C.SOURCE,
    sourceCost: C.SOURCE_COST, availability: C.AVAILABILITY,
    interviewStatus: C.INTERVIEW_STATUS, interviewDate: C.INTERVIEW_DATE,
    remarks: C.REMARKS, prevEmployer: C.PREV_EMPLOYER,
    currEmployer: C.CURR_EMPLOYER, relation: C.RELATION,
    nationality: C.NATIONALITY, location: C.LOCATION,
    cvLink: C.CV_LINK, duplicateFlag: C.DUPLICATE_FLAG,
    requisitionID: C.REQUISITION_ID, requestedBy: C.REQUESTED_BY,
    approvedBy: C.APPROVED_BY, offerStage: C.OFFER_STAGE,
    offeredSalary: C.OFFERED_SALARY, declineReason: C.DECLINE_REASON,
    expectedJoining: C.EXPECTED_JOINING, actualJoining: C.ACTUAL_JOINING,
    noticeDaysServed: C.NOTICE_DAYS_SERVED,
  };

  for (const [field, colIdx] of Object.entries(fieldMap)) {
    if (updates[field] !== undefined) {
      newRow[colIdx] = sanitize(updates[field]);
    }
  }

  // Auto-set status + time-to-hire on actual joining date
  if (updates.actualJoining && updates.actualJoining !== "") {
    newRow[C.STATUS] = "Joined";
    // existingRow[C.DATE] may be a raw Excel serial number (e.g. "46195") rather
    // than an ISO string — excelDateToISO() normalizes either format. Using
    // new Date() directly on a raw serial produces a garbage huge timestamp,
    // which is what caused TIME_TO_HIRE to show as -16132217 instead of a
    // sane day count.
    const entryISO    = excelDateToISO(existingRow[C.DATE]) || sanitize(existingRow[C.DATE]);
    const dateEntry   = new Date(entryISO);
    const dateJoined  = new Date(updates.actualJoining);
    if (!isNaN(dateEntry) && !isNaN(dateJoined)) {
      const totalDays   = Math.round((dateJoined - dateEntry) / 86400000);
      const noticeDays  = parseInt(sanitize(updates.noticeDaysServed || "")) || 0;
      // Store total elapsed days; noticeDaysServed stored separately in AQ
      if (totalDays >= 0) newRow[C.TIME_TO_HIRE] = totalDays;
      if (noticeDays > 0) {
        while (newRow.length <= C.NOTICE_DAYS_SERVED) newRow.push("");
        newRow[C.NOTICE_DAYS_SERVED] = noticeDays;
      }
    }
    // Auto-close linked requisition
    const reqID = sanitize(existingRow[C.REQUISITION_ID]);
    if (reqID) {
      await handleCloseRequisition({ requisitionID: reqID }, env);
    }
  } else if (sanitize(updates.status) === "Joined" && !sanitize(existingRow[C.ACTUAL_JOINING])) {
    // Safety net: status was set to "Joined" through the generic field-map path
    // (e.g. a direct status dropdown) without going through the proper "Mark
    // as Joined" flow that captures the joining date. Rather than silently
    // leaving ACTUAL_JOINING/TIME_TO_HIRE blank — which breaks Avg Time to
    // Hire and budget tracking — default the joining date to today so the
    // data is never incomplete.
    const todayISO = new Date().toISOString().split("T")[0];
    newRow[C.ACTUAL_JOINING] = todayISO;
    const entryISO  = excelDateToISO(existingRow[C.DATE]) || sanitize(existingRow[C.DATE]);
    const dateEntry = new Date(entryISO);
    const dateJoined = new Date(todayISO);
    if (!isNaN(dateEntry) && !isNaN(dateJoined)) {
      const totalDays = Math.round((dateJoined - dateEntry) / 86400000);
      if (totalDays >= 0) newRow[C.TIME_TO_HIRE] = totalDays;
    }
    const reqID = sanitize(existingRow[C.REQUISITION_ID]);
    if (reqID) await handleCloseRequisition({ requisitionID: reqID }, env);
  }

  await updateRow(token, env, SHEET_CANDIDATES, excelRow, newRow);
  return json({ success: true, candidateID });
}

// POST /add-interview-round
async function handleAddInterviewRound(body, env) {
  const { candidateID, round, interviewer, interviewerRole, team, scheduledDate, scheduledTime, score, feedback, outcome, status } = body;
  if (!candidateID || !round) return err("candidateID and round required");

  const token = await getToken(env);

  // Generate a unique raterID for this row — enables multi-rater per round
  const raterID = `R-${Date.now().toString(36)}-${Math.random().toString(36).slice(2,6)}`.toUpperCase();

  const row = new Array(38).fill("");
  row[R.CAND_ID]          = candidateID;
  row[R.ROUND]            = sanitize(round);
  row[R.INTERVIEWER]      = sanitize(interviewer || "");
  row[R.SCHED_DATE]       = sanitize(scheduledDate || "");
  row[R.ACTUAL_DATE]      = sanitize(body.actualDate || "");
  row[R.OUTCOME]          = sanitize(outcome || "");
  row[R.RESCHEDULE_COUNT] = "0";
  row[R.SCORE]            = sanitize(score || "");
  row[R.FEEDBACK]         = sanitize(feedback || "");
  row[R.SCHED_TIME]       = sanitize(scheduledTime || "");
  row[R.TEAM]             = sanitize(team || "");
  row[R.INTERVIEWER_ROLE] = sanitize(interviewerRole || "");
  row[R.STATUS]           = sanitize(status || (scheduledDate ? "Scheduled" : "Skipped"));
  row[R.RATER_ID]         = raterID;

  await appendRow(token, env, SHEET_ROUNDS, row);

  // Update candidate interview status / date summary fields for quick reference
  const statusLabel = sanitize(status || (scheduledDate ? "Scheduled" : "Skipped"));
  const candRows   = await getRows(token, env, SHEET_CANDIDATES);
  for (let i = 1; i < candRows.length; i++) {
    if (sanitize(candRows[i][C.ID]) === candidateID) {
      await updateCell(token, env, SHEET_CANDIDATES, i + 1, C.INTERVIEW_STATUS, statusLabel);
      if (scheduledDate) {
        await updateCell(token, env, SHEET_CANDIDATES, i + 1, C.INTERVIEW_DATE, scheduledDate);
      }
      break;
    }
  }

  return json({ success: true, candidateID, round, raterID });
}

// POST /update-interview-round
async function handleUpdateInterviewRound(body, env) {
  const { candidateID, round, updates } = body;
  if (!candidateID || !round) return err("candidateID and round required");

  const token    = await getToken(env);
  const rows     = await getRows(token, env, SHEET_ROUNDS);
  let excelRow   = -1;
  let existingRow = null;

  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][R.CAND_ID]) === candidateID &&
        sanitize(rows[i][R.ROUND])   === round) {
      excelRow    = i + 1;
      existingRow = rows[i];
      break;
    }
  }
  if (excelRow === -1) return err("Interview round not found", 404);

  const newRow = [...existingRow];
  while (newRow.length <= R.Q_COMPETITOR_DETAIL) newRow.push("");

  const fieldMap = {
    interviewer: R.INTERVIEWER, scheduledDate: R.SCHED_DATE,
    actualDate:  R.ACTUAL_DATE, outcome: R.OUTCOME,
    score:       R.SCORE,       feedback: R.FEEDBACK,
    scheduledTime: R.SCHED_TIME, team: R.TEAM,
    interviewerRole: R.INTERVIEWER_ROLE, status: R.STATUS,
  };
  for (const [field, idx] of Object.entries(fieldMap)) {
    if (updates[field] !== undefined) newRow[idx] = sanitize(updates[field]);
  }

  // Auto-increment reschedule count
  if (updates.outcome === "Rescheduled") {
    const current = parseInt(sanitize(existingRow[R.RESCHEDULE_COUNT]) || "0", 10);
    newRow[R.RESCHEDULE_COUNT] = current + 1;
  }

  await updateRow(token, env, SHEET_ROUNDS, excelRow, newRow);
  return json({ success: true, candidateID, round });
}

// POST /upload-cv  — multipart form: candidateID + file
async function handleUploadCV(request, env) {
  const formData = await request.formData();
  const candID   = formData.get("candidateID");
  const file     = formData.get("file");
  if (!candID || !file) return err("candidateID and file required");

  const token    = await getToken(env);
  const fileName = file.name || "cv.pdf";
  const bytes    = await file.arrayBuffer();
  const cvUrl    = await uploadCV(token, env, candID, fileName, bytes);

  // Update CVLink in Candidates sheet
  const rows = await getRows(token, env, SHEET_CANDIDATES);
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][C.ID]) === candID) {
      await updateCell(token, env, SHEET_CANDIDATES, i + 1, C.CV_LINK, cvUrl);
      break;
    }
  }

  return json({ success: true, candidateID: candID, cvLink: cvUrl });
}

// GET /get-dashboard  ?period=YYYYMM  (optional — blank = all time)
async function handleGetDashboard(url, env) {
  const token     = await getToken(env);
  const [candRows, roundRows] = await Promise.all([
    getRows(token, env, SHEET_CANDIDATES),
    getRows(token, env, SHEET_ROUNDS),
  ]);
  const period    = url.searchParams.get("period") || "";

  const data = candRows.slice(1).filter(r => {
    if (!r[C.ID]) return false;
    if (period) return sanitize(r[C.YEAR_MONTH]) === period;
    return true;
  });

  // Pipeline funnel
  const STATUS_ORDER = ["Active","Shortlisted","Offered","Joined","Rejected","On Hold"];
  const funnel = {};
  STATUS_ORDER.forEach(s => { funnel[s] = 0; });
  data.forEach(r => {
    const s = sanitize(r[C.STATUS]);
    if (funnel[s] !== undefined) funnel[s]++;
    else funnel["Active"]++;
  });

  // Entity breakdown
  const entities = {};
  data.forEach(r => {
    const e = sanitize(r[C.ENTITY]) || "Unknown";
    if (!entities[e]) entities[e] = { total:0, joined:0 };
    entities[e].total++;
    if (sanitize(r[C.STATUS]) === "Joined") entities[e].joined++;
  });

  // Nationality breakdown
  const natBreakdown = {};
  data.forEach(r => {
    const n = sanitize(r[C.NATIONALITY]) || "Unknown";
    natBreakdown[n] = (natBreakdown[n] || 0) + 1;
  });

  // Visa type breakdown
  const visaBreakdown = {};
  data.forEach(r => {
    const v = sanitize(r[C.VISA_TYPE]) || "Unknown";
    visaBreakdown[v] = (visaBreakdown[v] || 0) + 1;
  });

  // Time to hire average (per entity)
  const tthByEntity = {};
  data.forEach(r => {
    const tth = parseFloat(sanitize(r[C.TIME_TO_HIRE]));
    const e   = sanitize(r[C.ENTITY]) || "Unknown";
    if (!isNaN(tth) && tth > 0) {
      if (!tthByEntity[e]) tthByEntity[e] = { total: 0, count: 0 };
      tthByEntity[e].total += tth;
      tthByEntity[e].count++;
    }
  });
  const tthAvg = {};
  for (const [e, v] of Object.entries(tthByEntity)) {
    tthAvg[e] = Math.round(v.total / v.count);
  }

  // Visa expiry alerts (< 90 days, status Active or Shortlisted)
  const today      = new Date();
  const threshold  = new Date();
  threshold.setDate(today.getDate() + 90);
  const visaAlerts = data.filter(r => {
    const exp    = new Date(excelDateToISO(r[C.VISA_EXPIRY]));
    const status = sanitize(r[C.STATUS]);
    return !isNaN(exp) && exp <= threshold && exp >= today &&
           ["Active","Shortlisted"].includes(status);
  }).map(r => ({
    id:         sanitize(r[C.ID]),
    name:       sanitize(r[C.NAME]),
    visaExpiry: excelDateToISO(r[C.VISA_EXPIRY]),
    status:     sanitize(r[C.STATUS]),
    entity:     sanitize(r[C.ENTITY]),
  }));

  // Decline reasons
  const declineReasons = {};
  data.forEach(r => {
    if (sanitize(r[C.STATUS]) === "Rejected") {
      const d = sanitize(r[C.DECLINE_REASON]) || "Not Specified";
      declineReasons[d] = (declineReasons[d] || 0) + 1;
    }
  });

  // Source breakdown
  const sourceBreakdown = {};
  data.forEach(r => {
    const s = sanitize(r[C.SOURCE]) || "Unknown";
    sourceBreakdown[s] = (sourceBreakdown[s] || 0) + 1;
  });

  // Duplicate count
  const duplicateCount = data.filter(r => sanitize(r[C.DUPLICATE_FLAG]) === "Yes").length;

  return json({
    period:          period || "all",
    totalCandidates: data.length,
    funnel,
    entities,
    nationalityBreakdown: natBreakdown,
    visaBreakdown,
    timeToHireAvg:   tthAvg,
    visaExpiryAlerts: visaAlerts,
    declineReasons,
    sourceBreakdown,
    duplicateCount,
  });
}

// ════════════════════════════════════════════════════════════════════════
// REQUISITION HANDLERS
// ════════════════════════════════════════════════════════════════════════

// Requisition ID generator — REQ-0001 format
async function nextReqID(token, env) {
  const rows = await getRows(token, env, SHEET_REQUISITIONS);
  let max = 0;
  for (let i = rows.length - 1; i >= 1; i--) {
    const id = sanitize(rows[i][REQ.ID]);
    if (id.startsWith("REQ-")) {
      const n = parseInt(id.replace("REQ-", ""), 10);
      if (!isNaN(n) && n > max) { max = n; break; }
    }
  }
  return `REQ-${String(max + 1).padStart(4, "0")}`;
}

function rowToReq(r) {
  return {
    id:                 sanitize(r[REQ.ID]),
    date:               excelDateToISO(r[REQ.DATE]),
    entity:             sanitize(r[REQ.ENTITY]),
    position:           sanitize(r[REQ.POSITION]),
    customPosition:     sanitize(r[REQ.CUSTOM_POSITION]),
    department:         sanitize(r[REQ.DEPARTMENT]),
    headcount:          sanitize(r[REQ.HEADCOUNT]),
    reason:             sanitize(r[REQ.REASON]),
    budgetedSalaryMin:  sanitize(r[REQ.BUDGETED_SALARY_MIN]),
    budgetedSalaryMax:  sanitize(r[REQ.BUDGETED_SALARY_MAX]),
    requestedBy:        sanitize(r[REQ.REQUESTED_BY]),
    requestDate:        excelDateToISO(r[REQ.REQUEST_DATE]),
    hrStatus:           sanitize(r[REQ.HR_STATUS]),
    hrReviewedBy:       sanitize(r[REQ.HR_REVIEWED_BY]),
    hrReviewDate:       excelDateToISO(r[REQ.HR_REVIEW_DATE]),
    hrRemarks:          sanitize(r[REQ.HR_REMARKS]),
    ceoStatus:          sanitize(r[REQ.CEO_STATUS]),
    ceoApprovedBy:      sanitize(r[REQ.CEO_APPROVED_BY]),
    ceoApprovalDate:    excelDateToISO(r[REQ.CEO_APPROVAL_DATE]),
    ceoRemarks:         sanitize(r[REQ.CEO_REMARKS]),
    approvedVia:        sanitize(r[REQ.APPROVED_VIA]),
    changeLog:          sanitize(r[REQ.CHANGE_LOG] || ""),
    overallStatus:      sanitize(r[REQ.OVERALL_STATUS]),
    filledCount:        sanitize(r[REQ.FILLED_COUNT]),
    closedDate:         excelDateToISO(r[REQ.CLOSED_DATE]),
    yearMonth:          sanitize(r[REQ.YEAR_MONTH]),
    type:               sanitize(r[REQ.TYPE]),
    replacingEmployee:  sanitize(r[REQ.REPLACING_EMPLOYEE]),
    jobDescription:     sanitize(r[REQ.JOB_DESCRIPTION]),
    showOnCareers:      sanitize(r[REQ.SHOW_ON_CAREERS] || "Yes"),
    deptStrength:       sanitize(r[REQ.DEPT_STRENGTH]),
    departureReason:    sanitize(r[REQ.DEPARTURE_REASON]),
    nationalityPref:    sanitize(r[REQ.NATIONALITY_PREF]),
    emergencyHiring:    sanitize(r[REQ.EMERGENCY_HIRING]),
    companyVisa:        sanitize(r[REQ.COMPANY_VISA]),
    expectedJoining:    excelDateToISO(r[REQ.EXPECTED_JOINING]),
    customPosition:     sanitize(r[REQ.CUSTOM_POSITION]),
    branch:             sanitize(r[REQ.BRANCH]),
    reportingManager:   sanitize(r[REQ.REPORTING_MANAGER]),
    genderPreference:   sanitize(r[REQ.GENDER_PREFERENCE]),
    experienceRequired: sanitize(r[REQ.EXPERIENCE_REQUIRED]),
    newPositionJustification: sanitize(r[REQ.NEW_POSITION_JUSTIFICATION]),
  };
}

// Derive overall status from HR + CEO decisions
function deriveOverallStatus(hrStatus, ceoStatus) {
  if (hrStatus === "Rejected") return "Rejected";
  if (ceoStatus === "Rejected") return "Rejected";
  if (ceoStatus === "Approved") return "Approved";
  if (hrStatus === "Approved") return "Pending CEO Approval";
  if (hrStatus === "Pending") return "Pending HR Review";
  return "Draft";
}

// GET /get-requisitions?entity=&status=&position=
async function handleGetRequisitions(url, env) {
  const token  = await getToken(env);
  const rows   = await getRows(token, env, SHEET_REQUISITIONS);
  const params = url.searchParams;
  const fEntity   = params.get("entity")   || "";
  const fStatus   = params.get("status")   || "";
  const fPosition = params.get("position") || "";
  const includeClosed = params.get("includeClosed") === "true";

  let data = rows.slice(1).filter(r => r[REQ.ID]).map(rowToReq);
  if (fEntity)   data = data.filter(r => r.entity   === fEntity);
  if (fStatus)   data = data.filter(r => r.overallStatus === fStatus);
  if (fPosition) data = data.filter(r => r.position === fPosition);

  // Active-list default: hide withdrawn/cancelled/closed requisitions unless explicitly requested.
  // Used to drive the dashboard and the default Requisitions view; "Past Requisitions" passes includeClosed=true.
  if (!includeClosed) {
    data = data.filter(r => !(r.overallStatus || "").startsWith("Closed"));
  }

  // Enrich with filled count and candidate journey metrics from Candidates + InterviewRounds
  const [candRows, roundRows] = await Promise.all([
    getRows(token, env, SHEET_CANDIDATES),
    getRows(token, env, SHEET_ROUNDS).catch(() => []),
  ]);

  const interviewedIDs = new Set(roundRows.slice(1).map(r => sanitize(r[R.CAND_ID])).filter(Boolean));

  const fillMap = {};
  const metricsMap = {}; // reqID -> { contacted, interviewed, selected }
  const SELECTED_STATUSES = ["Accepted", "Onboarding", "Joined"];
  const INTERVIEWED_STATUSES = ["Ops Cleared", "Mgmt Cleared", "Verbal Offered", "Offered", "Accepted", "Onboarding", "Joined"];

  for (const cr of candRows.slice(1)) {
    const reqID  = sanitize(cr[C.REQUISITION_ID]);
    if (!reqID) continue;
    const candID = sanitize(cr[C.ID]);
    const status = sanitize(cr[C.STATUS]);

    if (status === "Joined") fillMap[reqID] = (fillMap[reqID] || 0) + 1;

    if (!metricsMap[reqID]) metricsMap[reqID] = { contacted: 0, interviewed: 0, selected: 0 };
    metricsMap[reqID].contacted++;
    if (interviewedIDs.has(candID) || INTERVIEWED_STATUSES.includes(status)) metricsMap[reqID].interviewed++;
    if (SELECTED_STATUSES.includes(status)) metricsMap[reqID].selected++;
  }

  data = data.map(r => ({
    ...r,
    filledCount: fillMap[r.id] || 0,
    candidatesContacted: metricsMap[r.id]?.contacted || 0,
    candidatesInterviewed: metricsMap[r.id]?.interviewed || 0,
    candidatesSelected: metricsMap[r.id]?.selected || 0,
  }));

  return json({ total: data.length, requisitions: data });
}

// GET /get-requisition/:id
async function handleGetRequisition(reqID, env) {
  const token = await getToken(env);
  const [rows, candRows] = await Promise.all([
    getRows(token, env, SHEET_REQUISITIONS),
    getRows(token, env, SHEET_CANDIDATES),
  ]);
  const row = rows.slice(1).find(r => sanitize(r[REQ.ID]) === reqID);
  if (!row) return err("Requisition not found", 404);

  const req = rowToReq(row);

  const candidates = candRows.slice(1)
    .filter(r => sanitize(r[C.REQUISITION_ID]) === reqID)
    .map(r => ({
      id:       sanitize(r[C.ID]),
      name:     sanitize(r[C.NAME]),
      position: sanitize(r[C.POSITION]),
      status:   sanitize(r[C.STATUS]),
      entity:   sanitize(r[C.ENTITY]),
    }));

  // Budget vs offered analysis
  const offered = candRows.slice(1)
    .filter(r => sanitize(r[C.REQUISITION_ID]) === reqID && sanitize(r[C.OFFERED_SALARY]))
    .map(r => parseFloat(sanitize(r[C.OFFERED_SALARY])) || 0);

  const budgetAnalysis = offered.length ? {
    budgetedMin: parseFloat(req.budgetedSalaryMin) || 0,
    budgetedMax: parseFloat(req.budgetedSalaryMax) || 0,
    avgOffered:  Math.round(offered.reduce((a,b) => a+b, 0) / offered.length),
    offersCount: offered.length,
  } : null;

  return json({ requisition: req, candidates, budgetAnalysis });
}

// POST /add-requisition  (unit head / hr_manager)
async function handleAddRequisition(body, env) {
  const { entity, position, customPosition, headcount, reason, requestedBy,
          department, budgetedSalaryMin, budgetedSalaryMax,
          type, replacingEmployee, jobDescription, companyVisa, expectedJoining } = body;

  const finalPosition = customPosition?.trim() || position;
  if (!entity || !finalPosition || !requestedBy) {
    return err("entity, position, requestedBy are required");
  }
  const token = await getToken(env);
  const reqID = await nextReqID(token, env);
  const today = isoDate();

  const row = new Array(29).fill("");
  row[REQ.ID]                  = reqID;
  row[REQ.DATE]                = today;
  row[REQ.ENTITY]              = sanitize(entity);
  row[REQ.POSITION]            = sanitize(finalPosition);
  row[REQ.DEPARTMENT]          = sanitize(department || "");
  row[REQ.HEADCOUNT]           = sanitize(headcount || "1");
  row[REQ.REASON]              = sanitize(reason || "");
  row[REQ.BUDGETED_SALARY_MIN] = sanitize(budgetedSalaryMin || "");
  row[REQ.BUDGETED_SALARY_MAX] = sanitize(budgetedSalaryMax || "");
  row[REQ.REQUESTED_BY]        = sanitize(requestedBy);
  row[REQ.REQUEST_DATE]        = today;
  row[REQ.HR_STATUS]           = "Pending";
  row[REQ.HR_REVIEWED_BY]      = "";
  row[REQ.HR_REVIEW_DATE]      = "";
  row[REQ.HR_REMARKS]          = "";
  row[REQ.CEO_STATUS]          = "Pending";
  row[REQ.CEO_APPROVED_BY]     = "";
  row[REQ.CEO_APPROVAL_DATE]   = "";
  row[REQ.CEO_REMARKS]         = "";
  row[REQ.OVERALL_STATUS]      = "Pending HR Review";
  row[REQ.FILLED_COUNT]        = "0";
  row[REQ.CLOSED_DATE]         = "";
  row[REQ.YEAR_MONTH]          = yearMonth();
  row[REQ.TYPE]                = sanitize(type || "New Position");
  row[REQ.REPLACING_EMPLOYEE]  = sanitize(replacingEmployee || "");
  row[REQ.JOB_DESCRIPTION]     = sanitize(jobDescription || "");
  row[REQ.COMPANY_VISA]        = sanitize(companyVisa || "");
  row[REQ.EXPECTED_JOINING]    = sanitize(expectedJoining || "");
  row[REQ.CUSTOM_POSITION]              = sanitize(customPosition || "");
  row[REQ.BRANCH]                       = sanitize(body.branch || "");
  row[REQ.REPORTING_MANAGER]            = sanitize(body.reportingManager || "");
  row[REQ.GENDER_PREFERENCE]            = sanitize(body.genderPreference || "");
  row[REQ.EXPERIENCE_REQUIRED]          = sanitize(body.experienceRequired || "");
  row[REQ.NEW_POSITION_JUSTIFICATION]   = sanitize(body.newPositionJustification || "");
  row[REQ.DEPT_STRENGTH]                = sanitize(body.deptStrength || "");
  row[REQ.DEPARTURE_REASON]             = sanitize(body.departureReason || "");
  row[REQ.NATIONALITY_PREF]             = sanitize(body.nationalityPref || "");
  row[REQ.EMERGENCY_HIRING]             = sanitize(body.emergencyHiring || "No");

  await appendRow(token, env, SHEET_REQUISITIONS, row);

  const typeLabel = type === "Replacement" ? `Replacement for: ${replacingEmployee||"TBD"}` : "New Position";
  const msg = `📋 New Requisition\nRef: ${reqID}\nPosition: ${finalPosition}\nEntity: ${entity}\nType: ${typeLabel}\nHeadcount: ${headcount||1}\nRequested By: ${requestedBy}`;
  await notifyWhatsApp(env, msg);

  return json({ success: true, requisitionID: reqID });
}

// POST /review-requisition  (hr_manager)
async function handleReviewRequisition(body, env) {
  const { requisitionID, hrStatus, hrReviewedBy, hrRemarks } = body;
  if (!requisitionID || !hrStatus || !hrReviewedBy) {
    return err("requisitionID, hrStatus, hrReviewedBy required");
  }
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_REQUISITIONS);
  let excelRow = -1;
  let existing = null;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][REQ.ID]) === requisitionID) {
      excelRow = i + 1; existing = rows[i]; break;
    }
  }
  if (excelRow === -1) return err("Requisition not found", 404);

  const newRow = [...existing];
  while (newRow.length < 35) newRow.push("");
  newRow[REQ.HR_STATUS]       = sanitize(hrStatus);
  newRow[REQ.HR_REVIEWED_BY]  = sanitize(hrReviewedBy);
  newRow[REQ.HR_REVIEW_DATE]  = isoDate();
  newRow[REQ.HR_REMARKS]      = sanitize(hrRemarks || "");
  newRow[REQ.OVERALL_STATUS]  = deriveOverallStatus(hrStatus, sanitize(existing[REQ.CEO_STATUS]));

  await updateRow(token, env, SHEET_REQUISITIONS, excelRow, newRow);

  if (hrStatus === "Approved") {
    const pos    = sanitize(existing[REQ.POSITION]);
    const entity = sanitize(existing[REQ.ENTITY]);
    const msg    = `✅ Requisition HR Approved\nRef: ${requisitionID}\nPosition: ${pos}\nEntity: ${entity}\nNow pending CEO approval.`;
    await notifyWhatsApp(env, msg);
  }

  return json({ success: true, requisitionID });
}

// POST /approve-requisition  (ceo)
async function handleApproveRequisition(body, env) {
  const { requisitionID, ceoStatus, ceoApprovedBy, ceoRemarks, approvedVia } = body;
  if (!requisitionID || !ceoStatus) {
    return err("requisitionID and ceoStatus are required");
  }
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_REQUISITIONS);
  let excelRow = -1;
  let existing = null;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][REQ.ID]) === requisitionID) {
      excelRow = i + 1; existing = rows[i]; break;
    }
  }
  if (excelRow === -1) return err("Requisition not found", 404);

  const newRow = [...existing];
  while (newRow.length < 35) newRow.push("");
  newRow[REQ.CEO_STATUS]        = sanitize(ceoStatus);
  newRow[REQ.CEO_APPROVED_BY]   = sanitize(ceoApprovedBy || "");
  newRow[REQ.CEO_APPROVAL_DATE] = isoDate();
  newRow[REQ.CEO_REMARKS]       = sanitize(ceoRemarks || "");
  newRow[REQ.APPROVED_VIA]      = sanitize(approvedVia || "Director");
  newRow[REQ.OVERALL_STATUS]    = deriveOverallStatus(sanitize(existing[REQ.HR_STATUS]), ceoStatus);

  await updateRow(token, env, SHEET_REQUISITIONS, excelRow, newRow);

  const viaLine = approvedVia && approvedVia !== "Director" ? `\nApproved via: ${approvedVia}` : "";
  const msg = `${ceoStatus === "Approved" ? "✅" : "❌"} Requisition CEO ${ceoStatus}\nRef: ${requisitionID}\nPosition: ${sanitize(existing[REQ.POSITION])}\nEntity: ${sanitize(existing[REQ.ENTITY])}${viaLine}`;
  await notifyWhatsApp(env, msg);

  return json({ success: true, requisitionID });
}

// POST /close-requisition  (hr_manager — called when candidate joins)
// GET /get-config — returns all master lists from Config sheet
// Dynamic: any section header in the sheet (except PIN/NOTIFICATIONS) is returned.
// New rows or sections added to the sheet automatically appear — no code change needed.
async function handleGetConfig(env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CONFIG);

  // Canonical section keys — map header text → camelCase key used by the portal
  const SECTION_KEYS = {
    "POSITIONS (Master List)": "positions",
    "ENTITIES":                "entities",
    "SOURCE OF CV":            "sources",
    "VISA TYPES":              "visaTypes",
    "EMIRATES":                "emirates",
    "DECLINE REASONS":         "declineReasons",
  };

  // These sections are never returned to the client (security)
  const SKIP_SECTIONS = new Set([
    "PINS — Role : PIN", "PINS - Role : PIN", "PINS",
    "NOTIFICATIONS", "USERS",
  ]);

  // Convert any unknown header to a camelCase key automatically
  const toCamel = (s) => s.toLowerCase()
    .replace(/[^a-z0-9 ]/g, "")
    .replace(/(?:^\w| \w)/g, (m, i) => i === 0 ? m.trim().toLowerCase() : m.trim().toUpperCase())
    .replace(/\s+/g, "");

  let currentSection = null;
  const result = {
    positions: [], entities: [], sources: [], visaTypes: [],
    emirates: [], declineReasons: [],
  };

  for (const row of rows) {
    const key = sanitize(row[0]);
    const val = sanitize(row[1]);

    // Blank key = end of section
    if (!key) { currentSection = null; continue; }

    // Check for PIN/skip sections first (case-insensitive prefix match)
    const keyUp = key.toUpperCase();
    if ([...SKIP_SECTIONS].some(s => keyUp.startsWith(s.toUpperCase()))) {
      currentSection = null; continue;
    }

    // Check for section headers — either known or all-caps heuristic (header rows tend to be ALL CAPS)
    const isKnownHeader = Object.prototype.hasOwnProperty.call(SECTION_KEYS, key);
    const isAllCaps     = key === key.toUpperCase() && key.length > 3 && !/^\d/.test(key);

    if (isKnownHeader || isAllCaps) {
      const sectionKey = SECTION_KEYS[key] || toCamel(key);
      currentSection   = sectionKey;
      if (!result[currentSection]) result[currentSection] = [];
      continue;
    }

    if (!currentSection) continue;

    // Entities section — short code in col A, full name in col B
    if (currentSection === "entities") {
      const name = val || key;
      if (name && !result.entities.includes(name)) result.entities.push(name);
      continue;
    }

    // All other sections — col A is the value
    if (!Array.isArray(result[currentSection])) result[currentSection] = [];
    if (key && !result[currentSection].includes(key)) result[currentSection].push(key);
  }

  return json(result);
}

// POST /upload-cv-public — no auth, careers page CV drop
// Writes to a STANDALONE CV Database — does NOT create a recruitment pipeline candidate.
// Position and facility are mandatory.
// POST /link-candidate-to-cvdb — after a candidate is added directly in the
// portal (not via the careers page), files their CV into the standalone
// CV Database too, so it's searchable there for future roles. Reuses the
// CV link already created by /upload-cv rather than re-uploading the file.
// Skips silently (no error) if a duplicate already exists in the database,
// checked by name+email or name+CV filename — no time window, unlike the
// careers-page double-submit guard.
async function handleLinkCandidateToCVDB(body, env) {
  const { name, email, phone, position, facility, cvLink, fileName, source } = body;
  if (!name)     return err("name required");
  if (!cvLink)   return err("cvLink required");
  if (!position) return err("position required");
  if (!facility) return err("facility required");

  const token = await getToken(env);

  const dup = await checkCVDBPermanentDuplicate(token, env, name, email, fileName);
  if (dup.duplicate) {
    return json({ success: true, linked: false, duplicate: true, existingID: dup.existingID, reason: dup.reason });
  }

  const rows = await getRows(token, env, SHEET_CVDB).catch(() => [[]]);
  let max = 0;
  for (let i = rows.length - 1; i >= 1; i--) {
    const id = sanitize(rows[i]?.[CVDB.ID]);
    if (id.startsWith("CVDB-")) {
      const n = parseInt(id.replace("CVDB-", ""), 10);
      if (!isNaN(n) && n > max) { max = n; break; }
    }
  }
  const cvID = `CVDB-${String(max + 1).padStart(4, "0")}`;

  const row = new Array(9).fill("");
  row[CVDB.ID]           = cvID;
  row[CVDB.NAME]         = sanitize(name);
  row[CVDB.POSITION]     = sanitize(position);
  row[CVDB.FACILITY]     = sanitize(facility);
  row[CVDB.SUBMITTED_AT] = new Date().toISOString();
  row[CVDB.CV_LINK]      = sanitize(cvLink);
  row[CVDB.PHONE]        = sanitize(phone || "");
  row[CVDB.EMAIL]        = sanitize(email || "");
  row[CVDB.SOURCE]       = sanitize(source || "Recruitment Portal — Add Candidate");
  await appendRow(token, env, SHEET_CVDB, row);

  return json({ success: true, linked: true, cvID });
}

async function handlePublicCVUpload(request, env) {
  const formData = await request.formData();
  const name     = sanitize(formData.get("name") || "");
  const position = sanitize(formData.get("position") || "");
  const facility = sanitize(formData.get("facility") || formData.get("entity") || "");
  const phone    = sanitize(formData.get("phone") || "");
  const email    = sanitize(formData.get("email") || "");
  const source   = sanitize(formData.get("source") || "Careers Page CV Drop");
  const file     = formData.get("file");

  if (!file)     return err("file required");
  if (!name)     return err("name required");
  if (!position) return err("position required");
  if (!facility) return err("facility required");

  const token = await getToken(env);

  // Prevent accidental double-submission (double-click, retry after slow network, etc.)
  const dup = await checkCVDBDuplicate(token, env, name, phone, email, position, facility);
  if (dup.duplicate) {
    return json({
      success: true,
      cvID: dup.existingID,
      duplicate: true,
      message: `You've already submitted a CV for this position recently. Reference: ${dup.existingID}`,
    });
  }

  // Generate CVDB ID
  const rows = await getRows(token, env, SHEET_CVDB).catch(() => [[]]);
  let max = 0;
  for (let i = rows.length - 1; i >= 1; i--) {
    const id = sanitize(rows[i]?.[CVDB.ID]);
    if (id.startsWith("CVDB-")) {
      const n = parseInt(id.replace("CVDB-", ""), 10);
      if (!isNaN(n) && n > max) { max = n; break; }
    }
  }
  const cvID = `CVDB-${String(max + 1).padStart(4, "0")}`;

  // Upload CV file to OneDrive, get public share link
  const fileName = file.name || "cv.pdf";
  const bytes    = await file.arrayBuffer();
  const cvUrl    = await uploadCV(token, env, cvID, fileName, bytes);

  // Write to CVDatabase sheet (separate from recruitment pipeline)
  const row = new Array(9).fill("");
  row[CVDB.ID]            = cvID;
  row[CVDB.NAME]          = name;
  row[CVDB.POSITION]      = position;
  row[CVDB.FACILITY]      = facility;
  row[CVDB.SUBMITTED_AT]  = new Date().toISOString();
  row[CVDB.CV_LINK]       = cvUrl;
  row[CVDB.PHONE]         = phone;
  row[CVDB.EMAIL]         = email;
  row[CVDB.SOURCE]        = source;
  await appendRow(token, env, SHEET_CVDB, row);

  // WhatsApp notify HR — informational only, not a pipeline action
  const msg = `📥 New CV in database\nRef: ${cvID}\nName: ${name}\nPosition: ${position}\nFacility: ${facility}`;
  await notifyWhatsApp(env, msg);

  return json({
    success: true,
    cvID,
    cvLink: cvUrl,
    message: `CV received. Reference: ${cvID}`,
  });
}

// GET /search-cv-database?position=&facility=&name=&dateFrom=&dateTo=
// HR-authenticated browse/search of the standalone CV database
async function handleSearchCVDatabase(url, env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CVDB).catch(() => [[]]);

  const qPosition = (url.searchParams.get("position") || "").toLowerCase();
  const qFacility = (url.searchParams.get("facility") || "").toLowerCase();
  const qName     = (url.searchParams.get("name") || "").toLowerCase();
  const qID       = (url.searchParams.get("id") || "").trim().toUpperCase();
  const dateFrom  = url.searchParams.get("dateFrom") || "";
  const dateTo    = url.searchParams.get("dateTo") || "";

  let entries = rows.slice(1).filter(r => r[CVDB.ID]).map(r => ({
    id:          sanitize(r[CVDB.ID]),
    name:        sanitize(r[CVDB.NAME]),
    position:    sanitize(r[CVDB.POSITION]),
    facility:    sanitize(r[CVDB.FACILITY]),
    submittedAt: sanitize(r[CVDB.SUBMITTED_AT]),
    cvLink:      sanitize(r[CVDB.CV_LINK]),
    phone:       sanitize(r[CVDB.PHONE]),
    email:       sanitize(r[CVDB.EMAIL]),
    linkedCandId: sanitize(r[CVDB.LINKED_CAND_ID]),
  }));

  if (qID)       entries = entries.filter(e => e.id.toUpperCase() === qID);
  if (qPosition) entries = entries.filter(e => e.position.toLowerCase().includes(qPosition));
  if (qFacility) entries = entries.filter(e => e.facility.toLowerCase().includes(qFacility));
  if (qName)     entries = entries.filter(e => e.name.toLowerCase().includes(qName));
  if (dateFrom)  entries = entries.filter(e => e.submittedAt >= dateFrom);
  if (dateTo)    entries = entries.filter(e => e.submittedAt <= dateTo + "T23:59:59");

  const page      = Math.max(1, parseInt(url.searchParams.get("page") || "1", 10));
  const limit     = Math.min(100, Math.max(1, parseInt(url.searchParams.get("limit") || "50", 10)));

  entries.sort((a, b) => (b.submittedAt || "").localeCompare(a.submittedAt || ""));

  const total = entries.length;
  const paged = entries.slice((page - 1) * limit, page * limit);

  return json({ entries: paged, total, page, limit, pages: Math.ceil(total / limit) });
}

// GET /get-open-positions — public, no auth
// Only shows positions backed by an APPROVED requisition with remaining headcount.
// Candidates sitting in the pipeline do NOT create public listings on their own.
async function handleGetOpenPositions(env) {
  const token   = await getToken(env);
  const reqRows = await getRows(token, env, SHEET_REQUISITIONS).catch(() => []);

  const reqMap = {};
  for (const r of reqRows.slice(1)) {
    const status = sanitize(r[REQ.OVERALL_STATUS]);
    if (status !== "Approved") continue; // strictly approved only — no other status creates a listing

    const showOnCareers = sanitize(r[REQ.SHOW_ON_CAREERS] || "Yes").toLowerCase();
    if (showOnCareers === "no") continue; // explicitly hidden from public careers page

    const headcount = parseInt(sanitize(r[REQ.HEADCOUNT])||"1",10);
    const filled     = parseInt(sanitize(r[REQ.FILLED_COUNT])||"0",10);
    if (filled >= headcount) continue; // fully filled — don't list

    const pos    = sanitize(r[REQ.CUSTOM_POSITION]) || sanitize(r[REQ.POSITION]);
    const entity = sanitize(r[REQ.ENTITY]);
    if (!pos || !entity) continue;

    const key = `${pos}||${entity}`;
    if (!reqMap[key]) reqMap[key] = {
      position:           pos,
      entity,
      headcount,
      filledCount:        filled,
      count:              headcount - filled,
      expectedJoining:    excelDateToISO(r[REQ.EXPECTED_JOINING]),
      experienceRequired: sanitize(r[REQ.EXPERIENCE_REQUIRED]),
      genderPreference:   sanitize(r[REQ.GENDER_PREFERENCE]),
      branch:             sanitize(r[REQ.BRANCH]),
      customPosition:     sanitize(r[REQ.CUSTOM_POSITION]),
      jobDescription:     sanitize(r[REQ.JOB_DESCRIPTION]),
    };
  }

  const positions = Object.values(reqMap)
    .sort((a, b) => a.entity.localeCompare(b.entity));

  return json({ positions });
}

// GET /track-status?id=REC-0001 — public, no auth
async function handleTrackStatus(url, env) {
  const id        = (url.searchParams.get("id") || "").toUpperCase().trim();
  const firstName = (url.searchParams.get("firstName") || "").trim().toLowerCase();
  if (!id.startsWith("REC-")) return err("Invalid reference ID. Format: REC-XXXX", 400);
  if (!firstName) return err("First name is required to check status", 400);

  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CANDIDATES);
  const row   = rows.slice(1).find(r => sanitize(r[C.ID]) === id);
  if (!row) return json({ found: false });

  // Verify the first name matches — prevents looking up someone else's status via guessed ID
  const storedName = sanitize(row[C.NAME]).trim().toLowerCase();
  const storedFirst = storedName.split(/\s+/)[0] || "";
  if (storedFirst !== firstName) {
    return json({ found: false, mismatch: true });
  }

  return json({
    found:           true,
    id,
    name:            sanitize(row[C.NAME]),
    position:        sanitize(row[C.POSITION]),
    entity:          sanitize(row[C.ENTITY]),
    status:          sanitize(row[C.STATUS]),
    interviewStatus: sanitize(row[C.INTERVIEW_STATUS]),
    interviewDate:   excelDateToISO(row[C.INTERVIEW_DATE]),
    offerStage:      sanitize(row[C.OFFER_STAGE]),
    expectedJoining: excelDateToISO(row[C.EXPECTED_JOINING]),
    submittedDate:   excelDateToISO(row[C.DATE]),
  });
}

// ════════════════════════════════════════════════════════════════════════
// MAIN FETCH HANDLER
// ════════════════════════════════════════════════════════════════════════
export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: CORS });
    }

    const url      = new URL(request.url);
    const path     = url.pathname.replace(/\/$/, "");
    const appKey   = request.headers.get("X-App-Key");
    const publicPaths = ["/verify-pin", "/get-open-positions", "/track-status", "/get-config", "/upload-cv-public", "/get-rating", "/submit-rating", "/rate-og", "/og-image", "/get-share", "/get-quotes", "/get-user-list", "/share-preview", "/get-news"];
    const isPublic = publicPaths.includes(path) || path.startsWith("/r/");

    // App key guard — skip for public paths
    if (!isPublic && appKey !== env.APP_KEY) {
      return json({ error: "Unauthorized" }, 401);
    }

    try {
      if (path === "/verify-pin" && request.method === "POST") {
        const clientIP = request.headers.get("CF-Connecting-IP") || "unknown";
        return await handleVerifyPin(await request.json(), env, clientIP);
      }
      if (path === "/get-candidates" && request.method === "GET") {
        return await handleGetCandidates(url, env, request);
      }
      if (path.startsWith("/get-candidate/") && request.method === "GET") {
        const id = path.replace("/get-candidate/", "");
        return await handleGetCandidate(id, env);
      }
      if (path === "/add-candidate" && request.method === "POST") {
        return await handleAddCandidate(await request.json(), env);
      }
      if (path === "/update-candidate" && request.method === "POST") {
        return await handleUpdateCandidate(await request.json(), env);
      }
      if (path === "/add-interview-round" && request.method === "POST") {
        return await handleAddInterviewRound(await request.json(), env);
      }
      if (path === "/update-interview-round" && request.method === "POST") {
        return await handleUpdateInterviewRound(await request.json(), env);
      }
      if (path === "/upload-cv" && request.method === "POST") {
        return await handleUploadCV(request, env);
      }
      if (path === "/link-candidate-to-cvdb" && request.method === "POST") {
        return await handleLinkCandidateToCVDB(await request.json(), env);
      }
      if (path === "/get-dashboard" && request.method === "GET") {
        return await handleGetDashboard(url, env);
      }

      if (path === "/upload-cv-public" && request.method === "POST") {
        return await handlePublicCVUpload(request, env);
      }
      if (path === "/search-cv-database" && request.method === "GET") {
        return await handleSearchCVDatabase(url, env);
      }
      if (path === "/update-cv-database" && request.method === "POST") {
        return await handleUpdateCVDB(await request.json(), env);
      }
      if (path === "/delete-cv-database" && request.method === "POST") {
        return await handleDeleteCVDB(await request.json(), env);
      }
      if (path === "/replace-cv-database-file" && request.method === "POST") {
        return await handleReplaceCVDBFile(request, env);
      }
      if (path === "/get-config" && request.method === "GET") {
        return await handleGetConfig(env);
      }
      if (path === "/get-open-positions" && request.method === "GET") {
        return await handleGetOpenPositions(env);
      }
      if (path === "/track-status" && request.method === "GET") {
        return await handleTrackStatus(url, env);
      }
      if (path === "/get-requisitions" && request.method === "GET") {
        return await handleGetRequisitions(url, env);
      }
      if (path.startsWith("/get-requisition/") && request.method === "GET") {
        return await handleGetRequisition(path.replace("/get-requisition/",""), env);
      }
      if (path === "/add-requisition" && request.method === "POST") {
        return await handleAddRequisition(await request.json(), env);
      }
      if (path === "/review-requisition" && request.method === "POST") {
        return await handleReviewRequisition(await request.json(), env);
      }
      if (path === "/approve-requisition" && request.method === "POST") {
        return await handleApproveRequisition(await request.json(), env);
      }
      if (path === "/close-requisition" && request.method === "POST") {
        return await handleCloseRequisition(await request.json(), env);
      }
      if (path === "/withdraw-requisition" && request.method === "POST") {
        return await handleWithdrawRequisition(await request.json(), env);
      }

      if (path === "/get-pipeline" && request.method === "GET") {
        return await handleGetPipeline(url, env);
      }
      if (path === "/add-pipeline-stage" && request.method === "POST") {
        return await handleAddPipelineStage(await request.json(), env);
      }
      if (path === "/update-pipeline-stage" && request.method === "POST") {
        return await handleUpdatePipelineStage(await request.json(), env);
      }
      if (path === "/handle-decline" && request.method === "POST") {
        return await handleDecline(await request.json(), env);
      }
      if (path === "/reject-candidate" && request.method === "POST") {
        return await handleRejectCandidate(await request.json(), env);
      }
      if (path === "/revert-candidate" && request.method === "POST") {
        return await handleRevertCandidate(await request.json(), env);
      }
      if (path === "/log-activity" && request.method === "POST") {
        return await handleLogActivity(await request.json(), env);
      }
      if (path === "/get-activity" && request.method === "GET") {
        return await handleGetActivity(url, env);
      }
      if (path === "/link-cvdb-candidate" && request.method === "POST") {
        return await handleLinkCVDBCandidate(await request.json(), env);
      }
      if (path === "/get-hiring-analytics" && request.method === "GET") {
        return await handleGetHiringAnalytics(env);
      }
      if (path === "/toggle-round-exclusion" && request.method === "POST") {
        return await handleToggleRoundExclusion(await request.json(), env);
      }
      if (path === "/toggle-rater-exclusion" && request.method === "POST") {
        return await handleToggleRaterExclusion(await request.json(), env);
      }
      if (path === "/send-rejection-whatsapp" && request.method === "POST") {
        return await handleSendRejectionWhatsApp(await request.json(), env);
      }
      if (path === "/patch-requisition-jd" && request.method === "POST") {
        return await handlePatchRequisitionJD(await request.json(), env);
      }
      if (path === "/edit-requisition" && request.method === "POST") {
        return await handleEditRequisition(await request.json(), env);
      }
      if (path === "/get-pipeline-history" && request.method === "GET") {
        return await handleGetPipelineHistory(env);
      }
      if (path === "/backfill-time-to-hire" && request.method === "POST") {
        return await handleBackfillTimeToHire(env);
      }
      if (path === "/get-quotes" && request.method === "GET") {
        return await handleGetQuotes(env);
      }
      // GET /get-news — public, returns published news posts from News sheet
      // News sheet columns: A=ID, B=TITLE, C=BODY, D=DATE, E=TAG, F=STATUS (Published/Draft), G=IMAGE_URL
      if (path === "/get-news" && request.method === "GET") {
        try {
          const token = await getToken(env);
          const rows  = await getRows(token, env, SHEET_NEWS);
          const posts = rows.slice(1)
            .filter(r => sanitize(r[5] || "").toLowerCase() === "published")
            .map(r => ({
              id:       sanitize(r[0]),
              title:    sanitize(r[1]),
              body:     sanitize(r[2]),
              date:     sanitize(r[3]),
              tag:      sanitize(r[4]),
              imageUrl: sanitize(r[6] || ""),
            }))
            .sort((a, b) => (b.date || "").localeCompare(a.date || ""));
          return json({ posts });
        } catch(e) {
          return json({ posts: [] }); // graceful fallback — show static posts if sheet missing
        }
      }
      if (path === "/rate-candidate-ai" && request.method === "POST") {
        return await handleRateCandidateAI(await request.json(), env);
      }
      if (path === "/rate-candidates-ai-bulk" && request.method === "POST") {
        return await handleRateCandidatesAIBulk(await request.json(), env);
      }
      if (path === "/generate-share-link" && request.method === "POST") {
        return await handleGenerateShareLink(await request.json(), env);
      }
      if (path === "/share-preview" && request.method === "GET") {
        return await handleSharePreview(url.searchParams.get("t"), env);
      }
      if (path === "/get-share" && request.method === "GET") {
        return await handleGetShare(url.searchParams.get("t"), env);
      }
      if (path === "/generate-rating-link" && request.method === "POST") {
        return await handleGenerateRatingLink(await request.json(), env);
      }
      // Public endpoints — no APP_KEY check (token is the auth)
      // WhatsApp/browser OG preview — serve HTML with candidate-specific meta tags
      // Short URL redirect: /r/XXXXXX → full rating page
      if (path.startsWith("/r/") && path.length > 3) {
        const code = path.slice(3).toUpperCase();
        return await handleShortRedirect(code, env);
      }
      if (path === "/og-image" && request.method === "GET") {
        // Minimal red PNG for WhatsApp OG preview (1×1 pixel, scales fine for previews)
        const png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
        const buf = Uint8Array.from(atob(png), c => c.charCodeAt(0));
        return new Response(buf, { headers: { "Content-Type":"image/png", "Cache-Control":"public,max-age=86400" } });
      }
      if (path === "/rate-og" && request.method === "GET") {
        return await handleRateOG(url, env, request);
      }
      if (path === "/get-rating" && request.method === "GET") {
        return await handleGetRating(url, env);
      }
      if (path === "/submit-rating" && request.method === "POST") {
        return await handleSubmitRating(await request.json(), env);
      }
      if (path === "/get-user-list" && request.method === "GET") {
        return await handleGetUserList(env);
      }
      if (path === "/check-whatsapp-setup" && request.method === "GET") {
        return await handleCheckWhatsappSetup(env);
      }
      if (path === "/get-users" && request.method === "GET") {
        return await handleGetUsers(env);
      }
      if (path === "/save-user" && request.method === "POST") {
        return await handleSaveUser(await request.json(), env);
      }
      if (path === "/delete-user" && request.method === "POST") {
        return await handleDeleteUser(await request.json(), env);
      }
      return json({ error: "Not found" }, 404);

    } catch (e) {
      console.error(e);
      return json({ error: e.message || "Internal error" }, 500);
    }
  },
};

// ── Activity Log ──────────────────────────────────────────────────────────────
// Appends a row to the ActivityLog sheet: timestamp, who, action text, role.
// Reads back the last N rows for the dashboard activity feed.
// The sheet is created automatically if missing (first append creates it).
async function handleLogActivity(body, env) {
  const { who, txt, role, ts } = body;
  if (!who || !txt) return err("Missing who/txt", 400);
  const token = await getToken(env);
  const timestamp = ts ? new Date(ts).toISOString() : new Date().toISOString();
  // Append a single row: [Timestamp, Who, Action, Role]
  await appendRow(token, env, SHEET_ACTIVITY, [timestamp, sanitize(who), sanitize(txt), sanitize(role||"")]);
  return json({ success: true });
}

async function handleGetActivity(url, env) {
  const token = await getToken(env);
  const limit = Math.min(50, parseInt(url.searchParams.get("limit") || "20", 10));
  const rows = await getRows(token, env, SHEET_ACTIVITY).catch(() => []);
  // Skip header row if present, return newest first
  const data = rows.slice(1).filter(r => r[0]).map(r => ({
    ts:   sanitize(r[0]),
    who:  sanitize(r[1]),
    txt:  sanitize(r[2]),
    role: sanitize(r[3]),
  })).reverse().slice(0, limit);
  return json({ activity: data });
}

// ── CV ↔ Candidate bidirectional link ────────────────────────────────────
// Writes CVDB_ID → Candidates sheet (col AN) and LINKED_CAND_ID → CVDatabase (col J)
async function handleLinkCVDBCandidate(body, env) {
  const { candidateID, cvdbID } = body;
  if (!candidateID || !cvdbID) return err("candidateID and cvdbID required");
  const token = await getToken(env);

  // 1. Write cvdbID into Candidates row
  const candRows = await getRows(token, env, SHEET_CANDIDATES);
  let candExcelRow = -1;
  for (let i = 1; i < candRows.length; i++) {
    if (sanitize(candRows[i][C.ID]) === candidateID) { candExcelRow = i + 1; break; }
  }
  if (candExcelRow === -1) return err("Candidate not found", 404);
  const candRow = [...candRows[candExcelRow - 1]];
  while (candRow.length <= C.CVDB_ID) candRow.push("");
  candRow[C.CVDB_ID] = cvdbID;
  await updateRow(token, env, SHEET_CANDIDATES, candExcelRow, candRow);

  // 2. Write candidateID into CVDatabase row
  const cvdbRows = await getRows(token, env, SHEET_CVDB);
  let cvdbExcelRow = -1;
  for (let i = 1; i < cvdbRows.length; i++) {
    if (sanitize(cvdbRows[i][CVDB.ID]) === cvdbID) { cvdbExcelRow = i + 1; break; }
  }
  if (cvdbExcelRow === -1) return err("CVDB entry not found", 404);
  const cvdbRow = [...cvdbRows[cvdbExcelRow - 1]];
  while (cvdbRow.length <= CVDB.LINKED_CAND_ID) cvdbRow.push("");
  cvdbRow[CVDB.LINKED_CAND_ID] = candidateID;
  await updateRow(token, env, SHEET_CVDB, cvdbExcelRow, cvdbRow);

  return json({ success: true, candidateID, cvdbID });
}

// ── Users sheet CRUD ─────────────────────────────────────────────────────
// GET /get-user-list — public, no auth, returns only name+key+role for login name picker
// Never returns PINs, entity filters, or nav restrictions.
async function handleGetUserList(env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_USERS).catch(() => [[]]);
  const users = rows.slice(1)
    .filter(r => r[U.KEY] && sanitize(r[U.ACTIVE]).toLowerCase() !== "false")
    .map(r => ({
      key:  sanitize(r[U.KEY]),
      name: sanitize(r[U.NAME]),
      role: sanitize(r[U.ROLE]),
    }));
  return json({ users });
}

async function handleGetUsers(env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_USERS).catch(() => [[]]);
  const users = rows.slice(1).filter(r => r[U.KEY]).map(r => ({
    key:          sanitize(r[U.KEY]),
    name:         sanitize(r[U.NAME]),
    role:         sanitize(r[U.ROLE]),
    pin:          sanitize(r[U.PIN]),
    entityFilter: sanitize(r[U.ENTITY_FILTER]),
    navPages:     sanitize(r[U.NAV_PAGES]),
    active:       sanitize(r[U.ACTIVE]).toLowerCase() !== "false",
  }));
  return json({ users });
}

async function handleSaveUser(body, env) {
  const { key, name, role, pin, entityFilter, navPages, active } = body;
  if (!key || !role || !pin) return err("key, role, pin required");
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_USERS).catch(() => [[]]);

  let excelRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][U.KEY]) === key) { excelRow = i + 1; break; }
  }

  const row = new Array(7).fill("");
  row[U.KEY]           = sanitize(key);
  row[U.NAME]          = sanitize(name || key);
  row[U.ROLE]          = sanitize(role);
  row[U.PIN]           = sanitize(String(pin));
  row[U.ENTITY_FILTER] = sanitize(entityFilter || "");
  row[U.NAV_PAGES]     = sanitize(navPages || "");
  row[U.ACTIVE]        = active === false ? "FALSE" : "TRUE";

  if (excelRow > 0) {
    await updateRow(token, env, SHEET_USERS, excelRow, row);
  } else {
    await appendRow(token, env, SHEET_USERS, row);
  }
  return json({ success: true, key });
}

async function handleDeleteUser(body, env) {
  const { key } = body;
  if (!key) return err("key required");
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_USERS).catch(() => [[]]);
  let excelRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][U.KEY]) === key) { excelRow = i + 1; break; }
  }
  if (excelRow === -1) return err("User not found", 404);
  await deleteSheetRow(token, env, SHEET_USERS, excelRow);
  return json({ success: true });
}

// ── Interviewer Rating Link ───────────────────────────────────────────────
// Generates a secure token for an external interviewer to submit structured
// ratings without needing portal access. Token expires 48h after interview date.
// Can be regenerated — overwrites the previous token for that round.

function generateToken() {
  const bytes = new Uint8Array(24);
  crypto.getRandomValues(bytes);
  return [...bytes].map(b => b.toString(16).padStart(2,"0")).join("");
}

function generateShortCode() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // no I,O,0,1 to avoid confusion
  const bytes = new Uint8Array(6);
  crypto.getRandomValues(bytes);
  return [...bytes].map(b => chars[b % chars.length]).join("");
}

async function handleShortRedirect(code, env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_ROUNDS).catch(() => [[]]);
  for (const row of rows.slice(1)) {
    if (sanitize(row[R.SHORT_CODE]).toUpperCase() !== code.toUpperCase()) continue;
    const t = sanitize(row[R.RATING_TOKEN]);
    const RATE_URL = `https://hifive-recruitment.pages.dev/rate?t=${encodeURIComponent(t)}`;

    // Check expiry
    const expires = sanitize(row[R.RATING_TOKEN_EXPIRES]);
    if (expires && Date.now() > new Date(expires).getTime()) {
      return ogHTML("Rating Link Expired", "This interview rating link has expired. Please ask HR to generate a new one.", RATE_URL);
    }

    // Look up candidate for rich OG preview
    const candID   = sanitize(row[R.CAND_ID]);
    const candRows = await getRows(token, env, SHEET_CANDIDATES).catch(() => [[]]);
    let name = "Candidate", position = "", entity = "";
    for (const cr of candRows.slice(1)) {
      if (sanitize(cr[C.ID]) === candID) {
        name     = sanitize(cr[C.NAME]);
        position = sanitize(cr[C.POSITION]);
        entity   = sanitize(cr[C.ENTITY]);
        break;
      }
    }
    const round    = sanitize(row[R.ROUND]);
    const date     = excelDateToISO(row[R.SCHED_DATE]);
    const time     = excelTimeToHHMM(row[R.SCHED_TIME]);
    const dateStr  = date ? new Date(date).toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"}) : "";
    const timeStr  = time ? ` at ${time}` : "";
    const title    = `Rate Interview: ${name}`;
    const desc     = `${position}${entity?" · "+entity:""}${round?" · "+round:""}${dateStr?" · "+dateStr+timeStr:""}. Tap to open and submit your rating.`;
    return ogHTML(title, desc, RATE_URL);
  }
  return new Response("Link not found or expired.", { status: 404, headers: {"Content-Type":"text/plain"} });
}

async function handleGenerateRatingLink(body, env) {
  const { candidateID, round, addRater, interviewer, interviewerRole, team } = body;
  if (!candidateID || !round) return err("candidateID and round required");

  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_ROUNDS);

  let excelRow = -1;
  let existing = null;

  if (addRater) {
    // Multi-rater mode: find the FIRST row for this round to copy schedule details,
    // then create a NEW row for this additional rater
    let templateRow = null;
    for (let i = 1; i < rows.length; i++) {
      if (sanitize(rows[i][R.CAND_ID]) === candidateID &&
          sanitize(rows[i][R.ROUND])   === round) {
        templateRow = rows[i]; break;
      }
    }
    if (!templateRow) return err("Interview round not found — add the round first", 404);

    // Generate unique raterID for this new rater row
    const raterID = `R-${Date.now().toString(36)}-${Math.random().toString(36).slice(2,6)}`.toUpperCase();
    const ratingToken = generateToken();
    const shortCode   = generateShortCode();
    const expires     = new Date(Date.now() + 7 * 24 * 3600000).toISOString(); // 7 days from now // 7 days

    const newRow = new Array(R.RATER_ID + 1).fill("");
    newRow[R.CAND_ID]              = candidateID;
    newRow[R.ROUND]                = round;
    newRow[R.INTERVIEWER]          = sanitize(interviewer || "");
    newRow[R.INTERVIEWER_ROLE]     = sanitize(interviewerRole || "");
    newRow[R.TEAM]                 = sanitize(team || sanitize(templateRow[R.TEAM]));
    newRow[R.SCHED_DATE]           = schedDate;
    newRow[R.SCHED_TIME]           = sanitize(templateRow[R.SCHED_TIME]);
    newRow[R.STATUS]               = "Scheduled";
    newRow[R.RATING_TOKEN]         = ratingToken;
    newRow[R.RATING_TOKEN_EXPIRES] = expires;
    newRow[R.SHORT_CODE]           = shortCode;
    newRow[R.RATER_ID]             = raterID;

    await appendRow(token, env, SHEET_ROUNDS, newRow);

    const shortUrl = `https://recruitment.sinusuresh.workers.dev/r/${shortCode}`;
    return json({ success: true, ratingToken, shortCode, shortUrl, raterID, isNewRow: true });
  }

  // Original mode: update the first existing row for this round
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][R.CAND_ID]) === candidateID &&
        sanitize(rows[i][R.ROUND])   === round) {
      excelRow = i + 1;
      existing = [...rows[i]];
      break;
    }
  }
  if (excelRow === -1) return err("Interview round not found", 404);

  // Expire 7 days from NOW (not from interview date which may be an Excel serial)
  const expires   = new Date(Date.now() + 7 * 24 * 3600000).toISOString(); // 7 days from generation
  const ratingToken = generateToken();
  const shortCode   = generateShortCode();

  while (existing.length <= R.SHORT_CODE) existing.push("");
  existing[R.RATING_TOKEN]         = ratingToken;
  existing[R.RATING_TOKEN_EXPIRES] = expires;
  existing[R.RATING_SUBMITTED_AT]  = "";
  existing[R.SHORT_CODE]           = shortCode;

  await updateRow(token, env, SHEET_ROUNDS, excelRow, existing);

  // Short URL for sharing — much cleaner than the 48-char token
  const shortUrl  = `https://recruitment.sinusuresh.workers.dev/r/${shortCode}`;
  // OG preview URL for WhatsApp rich card
  const ratingUrl = `https://recruitment.sinusuresh.workers.dev/rate-og?t=${ratingToken}`;
  return json({ success: true, ratingUrl, shortUrl, token: ratingToken, shortCode, expires });
}

async function handleGetRating(url, env) {
  const t = url.searchParams.get("t");
  if (!t) return json({ error: "Missing token" }, 400);

  const token = await getToken(env);

  // Fetch both sheets in parallel — eliminates 1 sequential round-trip
  const [roundRows, candRowsAll] = await Promise.all([
    getRows(token, env, SHEET_ROUNDS).catch(() => [[]]),
    getRows(token, env, SHEET_CANDIDATES).catch(() => [[]]),
  ]);

  let tokenRow = null;
  for (const row of roundRows.slice(1)) {
    if (sanitize(row[R.RATING_TOKEN]) === t) { tokenRow = row; break; }
  }
  if (!tokenRow) return json({ expired: false, valid: false, error: "Token not found" }, 404);

  const expires = sanitize(tokenRow[R.RATING_TOKEN_EXPIRES]);
  if (expires && Date.now() > new Date(expires).getTime()) {
    return json({ expired: true, message: "This rating link has expired. Please ask HR to generate a new link." }, 410);
  }

  const candID = sanitize(tokenRow[R.CAND_ID]);
  let cand = {};
  for (const cr of candRowsAll.slice(1)) {
    if (sanitize(cr[C.ID]) === candID) {
      cand = {
        id:             sanitize(cr[C.ID]),
        name:           sanitize(cr[C.NAME]),
        position:       sanitize(cr[C.POSITION]),
        entity:         sanitize(cr[C.ENTITY]),
        degree:         sanitize(cr[C.DEGREE]),
        uaeExp:         sanitize(cr[C.UAE_EXP]),
        nationality:    sanitize(cr[C.NATIONALITY]),
        expectedSalary: sanitize(cr[C.EXPECTED_SALARY]),
        noticePeriod:   sanitize(cr[C.NOTICE_PERIOD]),
        currEmployer:   sanitize(cr[C.CURR_EMPLOYER]),
        remarks:        sanitize(cr[C.REMARKS]),
        cvLink:         sanitize(cr[C.CV_LINK]),
        visaType:       sanitize(cr[C.VISA_TYPE]),
        emirate:        sanitize(cr[C.EMIRATE]),
      };
      break;
    }
  }

  // Reuse the already-fetched roundRows — no second fetch needed
  const prevRounds = roundRows.slice(1)
    .filter(r => sanitize(r[R.CAND_ID]) === candID && sanitize(r[R.RATING_TOKEN]) !== t)
    .map(r => ({
      round:          sanitize(r[R.ROUND]),
      status:         sanitize(r[R.STATUS]),
      outcome:        sanitize(r[R.OUTCOME]),
      interviewer:    sanitize(r[R.INTERVIEWER]),
      scheduledDate:  excelDateToISO(r[R.SCHED_DATE]),
      score:          sanitize(r[R.SCORE]),
      feedback:       sanitize(r[R.FEEDBACK]),
      ratingPassFail: sanitize(r[R.RATING_PASS_FAIL]),
      ratingRemarks:  sanitize(r[R.RATING_REMARKS]),
    }));

  const thisRound        = sanitize(tokenRow[R.ROUND]);
  const scheduledDate    = excelDateToISO(tokenRow[R.SCHED_DATE]);
  const scheduledTime    = excelTimeToHHMM(tokenRow[R.SCHED_TIME]);
  const alreadySubmitted = !!sanitize(tokenRow[R.RATING_SUBMITTED_AT]);

  return json({
    valid:           true,
    round:           thisRound,
    interviewer:     sanitize(tokenRow[R.INTERVIEWER]),
    interviewerRole: sanitize(tokenRow[R.INTERVIEWER_ROLE]),
    team:            sanitize(tokenRow[R.TEAM]),
    scheduledDate,
    scheduledTime,
    cand,
    prevRounds,
    alreadySubmitted,
    submittedByName: sanitize(tokenRow[R.SUBMITTED_BY_NAME]),
    scores: alreadySubmitted ? {
      education:        sanitize(tokenRow[R.SCORE_EDUCATION]),
      workExp:          sanitize(tokenRow[R.SCORE_WORK_EXP]),
      technical:        sanitize(tokenRow[R.SCORE_TECHNICAL]),
      communication:    sanitize(tokenRow[R.SCORE_COMMUNICATION]),
      enthusiasm:       sanitize(tokenRow[R.SCORE_ENTHUSIASM]),
      product:          sanitize(tokenRow[R.SCORE_PRODUCT]),
      teamwork:         sanitize(tokenRow[R.SCORE_TEAMWORK]),
      initiative:       sanitize(tokenRow[R.SCORE_INITIATIVE]),
      timeMgmt:         sanitize(tokenRow[R.SCORE_TIME_MGMT]),
      companyKnowledge: sanitize(tokenRow[R.SCORE_COMPANY_KNOWLEDGE]),
      omittedCriteria:  sanitize(tokenRow[R.OMITTED_CRITERIA]),
      passFail:         sanitize(tokenRow[R.RATING_PASS_FAIL]),
      remarks:          sanitize(tokenRow[R.RATING_REMARKS]),
      qShiftDuties:     sanitize(tokenRow[R.Q_SHIFT_DUTIES]),
      qDisciplinary:    sanitize(tokenRow[R.Q_DISCIPLINARY]),
      qRelative:        sanitize(tokenRow[R.Q_RELATIVE]),
      qRelativeDetail:  sanitize(tokenRow[R.Q_RELATIVE_DETAIL]),
      qCompetitor:      sanitize(tokenRow[R.Q_COMPETITOR]),
      qCompetitorDetail: sanitize(tokenRow[R.Q_COMPETITOR_DETAIL]),
    } : null,
  });
}


async function handleSubmitRating(body, env) {
  const { token: t, duration, education, workExp, technical, communication, enthusiasm, product, teamwork, initiative, timeMgmt, companyKnowledge, emotionalSensitivity, omittedCriteria, passFail, remarks, submittedByName, qShiftDuties, qDisciplinary, qRelative, qRelativeDetail, qCompetitor, qCompetitorDetail } = body;
  if (!t) return json({ error: "Missing token" }, 400);
  if (!passFail || !["Pass","Fail"].includes(passFail)) return json({ error: "passFail must be Pass or Fail" }, 400);
  if (!remarks || !remarks.trim()) return json({ error: "Remarks are required. Please share your observations about the candidate." }, 400);

  const msToken = await getToken(env);
  const rows    = await getRows(msToken, env, SHEET_ROUNDS).catch(() => [[]]);

  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][R.RATING_TOKEN]) !== t) continue;

    const expires = sanitize(rows[i][R.RATING_TOKEN_EXPIRES]);
    if (expires && Date.now() > new Date(expires).getTime()) {
      return json({ expired: true, message: "This rating link has expired. Please ask HR to generate a new link." }, 410);
    }

    const newRow = [...rows[i]];
    while (newRow.length <= R.RATER_ID) newRow.push("");

    newRow[R.RATING_SUBMITTED_AT]     = new Date().toISOString();
    newRow[R.SCORE_EDUCATION]         = sanitize(String(education        || ""));
    newRow[R.SCORE_WORK_EXP]          = sanitize(String(workExp          || ""));
    newRow[R.SCORE_TECHNICAL]         = sanitize(String(technical        || ""));
    newRow[R.SCORE_COMMUNICATION]     = sanitize(String(communication    || ""));
    newRow[R.SCORE_ENTHUSIASM]        = sanitize(String(enthusiasm       || ""));
    newRow[R.SCORE_PRODUCT]           = sanitize(String(product          || ""));
    newRow[R.SCORE_TEAMWORK]          = sanitize(String(teamwork         || ""));
    newRow[R.SCORE_INITIATIVE]        = sanitize(String(initiative       || ""));
    newRow[R.SCORE_TIME_MGMT]         = sanitize(String(timeMgmt        || ""));
    newRow[R.SCORE_COMPANY_KNOWLEDGE] = sanitize(String(companyKnowledge || ""));
    newRow[R.SCORE_EMOTIONAL_SENSITIVITY] = sanitize(String(emotionalSensitivity || ""));
    newRow[R.INTERVIEW_DURATION]          = sanitize(String(body.duration || ""));
    newRow[R.OMITTED_CRITERIA]        = sanitize(omittedCriteria || "");
    newRow[R.RATING_PASS_FAIL]        = sanitize(passFail);
    newRow[R.RATING_REMARKS]          = sanitize(remarks || "");
    newRow[R.SUBMITTED_BY_NAME]       = sanitize(submittedByName || "");
    newRow[R.Q_SHIFT_DUTIES]          = sanitize(qShiftDuties    || "");
    newRow[R.Q_DISCIPLINARY]          = sanitize(qDisciplinary   || "");
    newRow[R.Q_RELATIVE]              = sanitize(qRelative       || "");
    newRow[R.Q_RELATIVE_DETAIL]       = sanitize(qRelativeDetail || "");
    newRow[R.Q_COMPETITOR]            = sanitize(qCompetitor     || "");
    newRow[R.Q_COMPETITOR_DETAIL]     = sanitize(qCompetitorDetail || "");

    // Aggregate score = average of active (non-omitted, scored) criteria
    const omittedList = (omittedCriteria || "").split(",").map(s => s.trim()).filter(Boolean);
    const allCriteria = [
      ["education", education], ["workExp", workExp], ["technical", technical],
      ["communication", communication], ["enthusiasm", enthusiasm], ["product", product],
      ["teamwork", teamwork], ["initiative", initiative], ["timeMgmt", timeMgmt],
      ["companyKnowledge", companyKnowledge], ["emotionalSensitivity", emotionalSensitivity],
    ];
    const active = allCriteria.filter(([k, v]) => !omittedList.includes(k) && v && !isNaN(Number(v)) && Number(v) > 0);
    const avgScore = active.length
      ? Math.round((active.reduce((s, [, v]) => s + Number(v), 0) / active.length) * 10) / 10
      : "";
    newRow[R.SCORE]   = String(avgScore);
    newRow[R.OUTCOME] = passFail;
    newRow[R.STATUS]  = "Completed";

    await updateRow(msToken, env, SHEET_ROUNDS, i + 1, newRow);

    // Log activity
    await appendRow(msToken, env, SHEET_ACTIVITY, [
      new Date().toISOString(),
      sanitize(rows[i][R.INTERVIEWER]) || "Interviewer",
      `submitted rating for ${sanitize(rows[i][R.CAND_ID])} — ${sanitize(rows[i][R.ROUND])} — ${passFail}`,
      "external",
    ]).catch(()=>{});

    return json({ success: true, message: "Thank you! Your rating has been recorded." });
  }
  return json({ error: "Invalid token", expired: true, message: "This link is invalid or has expired." }, 404);
}

// ── OG Preview for WhatsApp ───────────────────────────────────────────────
// Returns HTML with candidate-specific OG meta tags so WhatsApp shows
// a rich preview (candidate name + position + round) when the link is shared.
// The actual rating page (rate.html) loads the full interactive UI.
async function handleRateOG(url, env, request) {
  const t = url.searchParams.get("t");
  const RATE_URL = `https://hifive-recruitment.pages.dev/rate?t=${encodeURIComponent(t||"")}`;

  if (!t) {
    return ogHTML("Interview Rating — HiFive Holdings", "Rate a candidate for HiFive Holdings.", RATE_URL);
  }

  try {
    const token = await getToken(env);
    const [rows, candRows] = await Promise.all([
      getRows(token, env, SHEET_ROUNDS).catch(() => [[]]),
      getRows(token, env, SHEET_CANDIDATES).catch(() => [[]]),
    ]);
    for (const row of rows.slice(1)) {
      if (sanitize(row[R.RATING_TOKEN]) !== t) continue;
      const candID = sanitize(row[R.CAND_ID]);
      let name = candID, position = "", entity = "";
      for (const cr of candRows.slice(1)) {
        if (sanitize(cr[C.ID]) === candID) {
          name = sanitize(cr[C.NAME]); position = sanitize(cr[C.POSITION]); entity = sanitize(cr[C.ENTITY]); break;
        }
      }
      const round = sanitize(row[R.ROUND]);
      const date  = excelDateToISO(row[R.SCHED_DATE]);
      const time  = excelTimeToHHMM(row[R.SCHED_TIME]);
      const dateStr = date ? new Date(date).toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"}) : "";
      const timeStr = time ? ` at ${time}` : "";
      const title = `Rate Interview: ${name}`;
      const desc  = `${position}${entity?" · "+entity:""} · ${round}${dateStr?" · "+dateStr+timeStr:""}. Tap to open and submit your interview rating for HiFive Holdings.`;
      return ogHTML(title, desc, RATE_URL);
    }
  } catch(e) {}

  return ogHTML("Interview Rating — HiFive Holdings", "Tap to open and submit your interview rating.", RATE_URL);
}

function ogHTML(title, description, url) {
  const esc = s => (s||"").replace(/&/g,"&amp;").replace(/"/g,"&quot;");
  // og:image must be PNG/JPG — WhatsApp ignores SVG
  // Serve a pre-built PNG or use the worker's own /og-image endpoint
  const ogImg = `https://recruitment.sinusuresh.workers.dev/og-image`;
  const html = `<!DOCTYPE html>
<html prefix="og: https://ogp.me/ns#">
<head>
<meta charset="UTF-8"/>
<title>${esc(title)}</title>
<meta property="og:title" content="${esc(title)}"/>
<meta property="og:description" content="${esc(description)}"/>
<meta property="og:url" content="${esc(url)}"/>
<meta property="og:type" content="website"/>
<meta property="og:site_name" content="HiFive Recruitment Portal"/>
<meta property="og:image" content="${esc(ogImg)}"/>
<meta property="og:image:width" content="1200"/>
<meta property="og:image:height" content="630"/>
<meta name="twitter:card" content="summary_large_image"/>
<meta name="twitter:title" content="${esc(title)}"/>
<meta name="twitter:description" content="${esc(description)}"/>
<meta name="twitter:image" content="${esc(ogImg)}"/>
<meta http-equiv="refresh" content="0;url=${esc(url)}"/>
<style>body{font-family:sans-serif;background:#1a0610;color:#fff;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center;padding:24px}h1{font-size:22px;margin-bottom:10px}p{font-size:14px;opacity:.7}</style>
</head>
<body>
<div><h1>${esc(title)}</h1><p>${esc(description)}</p><p style="margin-top:20px"><a href="${esc(url)}" style="color:#f5a623;font-weight:700">Open Rating Form →</a></p></div>
</body>
</html>`;
  return new Response(html, { headers: { "Content-Type":"text/html;charset=UTF-8", "Cache-Control":"no-store", "X-Robots-Tag":"noindex" } });
}

// ── Permanent round score exclusion ─────────────────────────────────────
// Stores comma-separated excluded round labels in C.EXCLUDED_ROUNDS (col AO)
// on the Candidates sheet. Toggles a round in or out of that list.
async function handleToggleRoundExclusion(body, env) {
  const { candidateID, round } = body;
  if (!candidateID || !round) return err("candidateID and round required");
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CANDIDATES);
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][C.ID]) !== candidateID) continue;
    const current  = sanitize(rows[i][C.EXCLUDED_ROUNDS] || "");
    const excluded = current.split(",").map(s => s.trim()).filter(Boolean);
    const idx      = excluded.indexOf(round);
    if (idx > -1) excluded.splice(idx, 1);
    else          excluded.push(round);
    // Use updateRow (not updateCell) so Excel's usedRange expands to include col AO
    const row = [...rows[i]];
    while (row.length <= EXCLUDED_RATERS_COL) row.push("");
    row[C.EXCLUDED_ROUNDS] = excluded.join(",");
    await updateRow(token, env, SHEET_CANDIDATES, i + 1, row);
    return json({ success: true, excludedRounds: excluded.join(",") });
  }
  return err("Candidate not found", 404);
}

// ── Patch Requisition JD ─────────────────────────────────────────────────
async function handlePatchRequisitionJD(body, env) {
  const { requisitionID, jobDescription, requestDate, budgetedSalaryMin, budgetedSalaryMax } = body;
  if (!requisitionID) return err("requisitionID required");
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_REQUISITIONS);
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][REQ.ID]) !== requisitionID) continue;
    const row = [...rows[i]];
    while (row.length <= REQ.APPROVED_VIA) row.push("");
    if (jobDescription    !== undefined) row[REQ.JOB_DESCRIPTION]    = sanitize(jobDescription || "");
    if (requestDate       !== undefined) row[REQ.REQUEST_DATE]        = sanitize(requestDate || "");
    if (budgetedSalaryMin !== undefined) row[REQ.BUDGETED_SALARY_MIN] = sanitize(String(budgetedSalaryMin || ""));
    if (budgetedSalaryMax !== undefined) row[REQ.BUDGETED_SALARY_MAX] = sanitize(String(budgetedSalaryMax || ""));
    await updateRow(token, env, SHEET_REQUISITIONS, i + 1, row);
    return json({ success: true });
  }
  return err("Requisition not found", 404);
}

// ── Send WhatsApp message to a specific phone number ─────────────────────
// Used for candidate-facing messages (rejection, interview confirmation, etc.)
async function sendWhatsAppToNumber(env, toPhone, message) {
  if (!env.WHATSAPP_TOKEN || !env.WHATSAPP_PHONE_ID) {
    return { sent: false, reason: "WhatsApp not configured" };
  }
  // Normalise phone: strip non-digits, ensure starts with country code
  let phone = toPhone.replace(/\D/g, "");
  if (phone.startsWith("0")) phone = "971" + phone.slice(1); // UAE local
  if (!phone.startsWith("971") && phone.length === 9) phone = "971" + phone; // bare UAE
  if (phone.length < 10) return { sent: false, reason: "Invalid phone number" };

  try {
    const res = await fetch(
      `https://graph.facebook.com/v19.0/${env.WHATSAPP_PHONE_ID}/messages`,
      {
        method: "POST",
        headers: {
          Authorization:  `Bearer ${env.WHATSAPP_TOKEN}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          messaging_product: "whatsapp",
          to: phone,
          type: "text",
          text: { body: message },
        }),
      }
    );
    const data = await res.json().catch(() => ({}));
    return { sent: res.ok, messageID: data?.messages?.[0]?.id, phone };
  } catch(e) {
    return { sent: false, reason: e.message };
  }
}

// ── Rejection WhatsApp to candidate ──────────────────────────────────────
// Generates a professional, warm rejection message and sends it to the candidate's
// registered phone. Called from the portal after confirming rejection.
async function handleSendRejectionWhatsApp(body, env) {
  const { candidateID, messageBody } = body;
  if (!candidateID) return err("candidateID required");

  const token    = await getToken(env);
  const candRows = await getRows(token, env, SHEET_CANDIDATES);
  let cand = null;
  for (const row of candRows.slice(1)) {
    if (sanitize(row[C.ID]) === candidateID) {
      cand = row; break;
    }
  }
  if (!cand) return err("Candidate not found", 404);

  const phone    = sanitize(cand[C.PHONE]);
  const name     = sanitize(cand[C.NAME]);
  const position = sanitize(cand[C.POSITION]);
  const entity   = sanitize(cand[C.ENTITY]);

  // Use custom message if provided, otherwise the default template
  const msg = messageBody?.trim() || [
    `Dear ${name.split(" ")[0]},`,
    ``,
    `Thank you for your interest in the ${position} role at ${entity} – HiFive Holdings, and for the time you invested in our selection process.`,
    ``,
    `After careful consideration, we regret to inform you that we will not be proceeding with your application at this time. This was a difficult decision as we had strong candidates for this role.`,
    ``,
    `We have retained your details and will reach out if a suitable opportunity arises in the future. We encourage you to apply again for roles that match your profile at hifive-careers.pages.dev.`,
    ``,
    `We wish you every success in your career journey.`,
    ``,
    `Warm regards,`,
    `HiFive Holdings — HR Team`,
  ].join("\n");

  const result = await sendWhatsAppToNumber(env, phone, msg);

  // Log the communication
  await appendRow(token, env, SHEET_ACTIVITY, [
    new Date().toISOString(),
    "System",
    `Rejection WhatsApp ${result.sent ? "sent" : "failed"} to ${name} (${candidateID}) — ${phone}`,
    "hr",
  ]).catch(() => {});

  return json({ success: true, sent: result.sent, phone: result.phone, reason: result.reason });
}

// ── Toggle rater-level exclusion ──────────────────────────────────────────
// raterID: specific rater row to toggle (preferred, per-rater)
// round: legacy round-label toggle (kept for backward compat)
async function handleToggleRaterExclusion(body, env) {
  const { candidateID, raterID, round } = body;
  if (!candidateID || (!raterID && !round)) return err("candidateID + raterID or round required");

  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CANDIDATES);

  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][C.ID]) !== candidateID) continue;

    if (raterID) {
      // Rater-level exclusion (column AP)
      const current  = sanitize(rows[i][EXCLUDED_RATERS_COL] || "");
      const excluded = current ? current.split(",").map(s => s.trim()).filter(Boolean) : [];
      const idx      = excluded.indexOf(raterID);
      if (idx === -1) excluded.push(raterID);
      else            excluded.splice(idx, 1);
      const newVal = excluded.join(",");
      const row = [...rows[i]];
      while (row.length <= EXCLUDED_RATERS_COL) row.push("");
      row[EXCLUDED_RATERS_COL] = newVal;
      await updateRow(token, env, SHEET_CANDIDATES, i + 1, row);
      return json({ success: true, excludedRaters: newVal });
    } else {
      // Round-level exclusion (column AO) — use updateRow so usedRange expands
      const current  = sanitize(rows[i][C.EXCLUDED_ROUNDS] || "");
      const excluded = current ? current.split(",").map(s => s.trim()).filter(Boolean) : [];
      const idx      = excluded.indexOf(round);
      if (idx === -1) excluded.push(round);
      else            excluded.splice(idx, 1);
      const row2 = [...rows[i]];
      while (row2.length <= EXCLUDED_RATERS_COL) row2.push("");
      row2[C.EXCLUDED_ROUNDS] = excluded.join(",");
      await updateRow(token, env, SHEET_CANDIDATES, i + 1, row2);
      return json({ success: true, excludedRounds: excluded.join(",") });
    }
  }
  return err("Candidate not found", 404);
}

// ── GET /get-hiring-analytics ────────────────────────────────────────────
// Joins Requisitions + Candidates to produce per-requisition hiring metrics:
// salary vs budget, time-to-hire, funnel depth, source cost
async function handleGetHiringAnalytics(env) {
  const token = await getToken(env);
  const [reqRows, candRows, roundRows] = await Promise.all([
    getRows(token, env, SHEET_REQUISITIONS),
    getRows(token, env, SHEET_CANDIDATES),
    getRows(token, env, SHEET_ROUNDS),
  ]);

  // Build interview count map keyed by candidateID
  const interviewsByCand = {};
  for (const r of roundRows.slice(1)) {
    const cid = sanitize(r[R.CAND_ID]);
    if (cid) interviewsByCand[cid] = (interviewsByCand[cid] || 0) + 1;
  }

  // Build candidate map keyed by requisitionID
  const candsByReq = {};
  const allCands   = candRows.slice(1).filter(r => r[C.ID]);
  for (const r of allCands) {
    const rid = sanitize(r[C.REQUISITION_ID]);
    if (!rid) continue;
    if (!candsByReq[rid]) candsByReq[rid] = [];
    candsByReq[rid].push(r);
  }

  const metrics = reqRows.slice(1).filter(r => r[REQ.ID]).map(r => {
    const reqID      = sanitize(r[REQ.ID]);
    const position   = sanitize(r[REQ.CUSTOM_POSITION]) || sanitize(r[REQ.POSITION]);
    const entity     = sanitize(r[REQ.ENTITY]);
    const reqDate    = excelDateToISO(r[REQ.REQUEST_DATE]) || excelDateToISO(r[REQ.DATE]);
    const budMin     = parseFloat(sanitize(r[REQ.BUDGETED_SALARY_MIN])) || 0;
    const budMax     = parseFloat(sanitize(r[REQ.BUDGETED_SALARY_MAX])) || 0;
    const status     = sanitize(r[REQ.OVERALL_STATUS]);
    const headcount  = parseInt(sanitize(r[REQ.HEADCOUNT])) || 1;
    const cands      = candsByReq[reqID] || [];

    // Candidates who reached offer stage or beyond
    const offered    = cands.filter(c => ["Verbal Offered","Offered","Accepted","Onboarding","Joined"]
                         .includes(sanitize(c[C.STATUS])));
    const joined     = cands.filter(c => sanitize(c[C.STATUS]) === "Joined");
    const rejected   = cands.filter(c => sanitize(c[C.STATUS]) === "Rejected");

    // Offered salaries
    const salaries   = offered
      .map(c => parseFloat(sanitize(c[C.OFFERED_SALARY])))
      .filter(v => !isNaN(v) && v > 0);
    const offeredSalaryMin = salaries.length ? Math.min(...salaries) : null;
    const offeredSalaryMax = salaries.length ? Math.max(...salaries) : null;
    const offeredSalaryAvg = salaries.length
      ? Math.round(salaries.reduce((a,b)=>a+b,0) / salaries.length)
      : null;

    // Budget variance: positive = over budget, negative = under
    // Fall back to budMin if budMax wasn't set on the requisition (common when
    // only a single budget figure was entered instead of a min-max range)
    const effectiveBudMax = budMax || budMin || null;
    const budgetVariance = (offeredSalaryAvg && effectiveBudMax)
      ? Math.round(offeredSalaryAvg - effectiveBudMax)
      : null;
    const withinBudget = offeredSalaryAvg && effectiveBudMax
      ? offeredSalaryAvg <= effectiveBudMax
      : null;

    // Time to hire per joined candidate (days from req date to acceptance)
    const tthValues  = cands
      .map(c => parseFloat(sanitize(c[C.TIME_TO_HIRE])))
      .filter(v => !isNaN(v) && v >= 0); // v>=0 — a same-day hire (0 days) is valid and must count
    const timeToHireAvg = tthValues.length
      ? Math.round(tthValues.reduce((a,b)=>a+b,0) / tthValues.length)
      : null;

    // Funnel depth: total CVs screened per hire
    const cvPerHire = joined.length > 0
      ? Math.round(cands.length / joined.length)
      : cands.length > 0 ? null : null;

    // Interview counts for this requisition's candidates
    const totalInterviews = cands.reduce((sum, c) => sum + (interviewsByCand[sanitize(c[C.ID])] || 0), 0);
    const interviewsPerHire = joined.length > 0 && totalInterviews > 0
      ? Math.round(totalInterviews / joined.length)
      : totalInterviews > 0 ? null : null;

    // Source breakdown for this req
    const sources = {};
    cands.forEach(c => {
      // Normalize: trim + collapse whitespace + title-case for consistent grouping
      // (fixes "Emirati Talent" vs "emirati talent " being counted as separate sources)
      const raw = sanitize(c[C.SOURCE]).replace(/\s+/g, " ").trim();
      const s = raw ? raw.replace(/\w\S*/g, t => t[0].toUpperCase() + t.slice(1).toLowerCase()) : "Unknown";
      sources[s] = (sources[s] || 0) + 1;
    });
    const topSource = Object.entries(sources).sort((a,b)=>b[1]-a[1])[0]?.[0] || null;

    // Notice days actually served — from the NOTICE_DAYS_SERVED field (AQ)
    // This is entered by HR when marking the candidate as Joined, not estimated from notice period text
    const noticeDaysActual = joined.map(c => {
      const nd = parseFloat(sanitize(c[C.NOTICE_DAYS_SERVED] || ""));
      return !isNaN(nd) && nd > 0 ? nd : null;
    }).filter(v => v != null);
    const avgNoticeDaysServed = noticeDaysActual.length
      ? Math.round(noticeDaysActual.reduce((a,b)=>a+b,0)/noticeDaysActual.length)
      : null;
    // Net time to hire = total TTH minus actual notice days served
    const netTTHValues = joined.map(c => {
      const total  = parseFloat(sanitize(c[C.TIME_TO_HIRE] || ""));
      const notice = parseFloat(sanitize(c[C.NOTICE_DAYS_SERVED] || ""));
      if (!isNaN(total) && total > 0) {
        return !isNaN(notice) && notice > 0 ? total - notice : total;
      }
      return null;
    }).filter(v => v != null && v > 0);
    const netTimeToHireAvg = netTTHValues.length
      ? Math.round(netTTHValues.reduce((a,b)=>a+b,0)/netTTHValues.length)
      : timeToHireAvg; // fallback to total if no notice data

    return {
      reqID, position, entity, reqDate, status, headcount,
      budMin, budMax,
      totalCVs:       cands.length,
      shortlisted:    cands.filter(c => !["Active"].includes(sanitize(c[C.STATUS]))).length,
      offered:        offered.length,
      joined:         joined.length,
      rejected:       rejected.length,
      offeredSalaryMin, offeredSalaryMax, offeredSalaryAvg,
      budgetVariance, withinBudget,
      timeToHireAvg,          // total: joining date - entry date
      netTimeToHireAvg,       // total minus actual notice served (recruitment-only time)
      avgNoticeDaysServed,
      cvPerHire,
      totalInterviews: totalInterviews || null,
      interviewsPerHire: interviewsPerHire || null,
      topSource, sources,
    };
  });

  // Group summary stats
  const filled = metrics.filter(m => m.joined > 0);
  const avgTTH = filled.length
    ? Math.round(filled.filter(m=>m.timeToHireAvg!=null).reduce((s,m)=>s+m.timeToHireAvg,0)
        / filled.filter(m=>m.timeToHireAvg!=null).length)
    : 0;
  const netFilled = filled.filter(m => m.netTimeToHireAvg != null);
  const avgNetTTH = netFilled.length
    ? Math.round(netFilled.reduce((s,m)=>s+m.netTimeToHireAvg,0) / netFilled.length)
    : null;
  // ── Source ROI ──────────────────────────────────────────────────────────
  // Group all candidates by source channel and compute:
  // - total CVs submitted from that source
  // - how many reached offer stage (acceptance rate numerator)
  // - how many accepted / joined (quality metric)
  // - average interview score across all scored candidates from that source
  // - total source cost spent on that channel
  // Normalize source names so "Emirati Talent" / "emirati talent " / "EMIRATI TALENT"
  // all group together instead of appearing as separate rows.
  const normSource = (raw) => {
    const t = sanitize(raw).replace(/\s+/g, " ").trim();
    return t ? t.replace(/\w\S*/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase()) : "Unknown";
  };

  const sourceMap = {};
  for (const r of allCands) {
    const src   = normSource(r[C.SOURCE]);
    const stat  = sanitize(r[C.STATUS] || "");
    const score = parseFloat(sanitize(r[C.TIME_TO_HIRE] || "")); // reuse existing field
    const cost  = parseFloat(sanitize(r[C.SOURCE_COST] || ""));

    if (!sourceMap[src]) sourceMap[src] = {
      source: src,
      totalCVs: 0, offered: 0, accepted: 0, joined: 0,
      scores: [], totalCost: 0,
    };

    const g = sourceMap[src];
    g.totalCVs++;
    if (["Verbal Offered","Offered","Accepted","Onboarding","Joined"].includes(stat)) g.offered++;
    if (["Accepted","Onboarding","Joined"].includes(stat)) g.accepted++;
    if (stat === "Joined") g.joined++;
    if (!isNaN(cost) && cost > 0) g.totalCost += cost;
  }

  // Compute interview scores per source from InterviewRounds
  // We need to cross-reference candidates by their source
  const candSourceMap = {};
  for (const r of allCands) {
    const id  = sanitize(r[C.ID]);
    const src = normSource(r[C.SOURCE]);
    if (id) candSourceMap[id] = src;
  }

  // Get InterviewRounds for score aggregation per source
  const roundRowsForROI = await getRows(token, env, SHEET_ROUNDS).catch(() => [[]]);
  for (const r of roundRowsForROI.slice(1)) {
    const candID = sanitize(r[R.CAND_ID]);
    const score  = parseFloat(sanitize(r[R.SCORE] || ""));
    const src    = candSourceMap[candID];
    if (src && sourceMap[src] && !isNaN(score) && score > 0) {
      sourceMap[src].scores.push(score);
    }
  }

  const sourceROI = Object.values(sourceMap)
    .map(g => ({
      source:           g.source,
      totalCVs:         g.totalCVs,
      offered:          g.offered,
      accepted:         g.accepted,
      joined:           g.joined,
      offerRate:        g.totalCVs > 0 ? Math.round((g.offered / g.totalCVs) * 100) : 0,
      acceptanceRate:   g.offered > 0  ? Math.round((g.accepted / g.offered) * 100) : 0,
      joinRate:         g.totalCVs > 0 ? Math.round((g.joined / g.totalCVs) * 100) : 0,
      avgScore:         g.scores.length > 0
        ? Math.round((g.scores.reduce((s,v)=>s+v,0) / g.scores.length) * 10) / 10
        : null,
      totalCost:        g.totalCost || 0,
      costPerHire:      g.joined > 0 && g.totalCost > 0
        ? Math.round(g.totalCost / g.joined)
        : null,
    }))
    .sort((a, b) => b.totalCVs - a.totalCVs); // most-used source first

  const overBudget  = metrics.filter(m => m.withinBudget === false).length;
  const underBudget = metrics.filter(m => m.withinBudget === true).length;

  return json({
    metrics,
    sourceROI,
    summary: {
      totalReqs: metrics.length,
      filled: filled.length,
      avgTimeToHire: avgTTH,
      avgNetTimeToHire: avgNetTTH,
      overBudget, underBudget,
    },
  });
}

// ── POST /edit-requisition — post-approval edits with audit trail ─────────
// Allows HR Manager to edit approved requisitions. All changes are logged
// with a timestamp, editor name, and note in the CHANGE_LOG column (AJ).
async function handleEditRequisition(body, env) {
  const { requisitionID, changes, editedBy, changeNote } = body;
  if (!requisitionID) return err("requisitionID required");
  if (!changes || Object.keys(changes).length === 0) return err("No changes provided");
  if (!changeNote || !changeNote.trim()) return err("Change note is required");

  const editableFields = {
    headcount:          REQ.HEADCOUNT,
    budgetedSalaryMin:  REQ.BUDGETED_SALARY_MIN,
    budgetedSalaryMax:  REQ.BUDGETED_SALARY_MAX,
    expectedJoining:    REQ.EXPECTED_JOINING,
    reportingManager:   REQ.REPORTING_MANAGER,
    branch:             REQ.BRANCH,
    genderPreference:   REQ.GENDER_PREFERENCE,
    experienceRequired: REQ.EXPERIENCE_REQUIRED,
    companyVisa:        REQ.COMPANY_VISA,
    requestDate:        REQ.REQUEST_DATE,
    jobDescription:     REQ.JOB_DESCRIPTION,
    showOnCareers:      REQ.SHOW_ON_CAREERS,
  };

  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_REQUISITIONS);

  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][REQ.ID]) !== requisitionID) continue;

    const row = [...rows[i]];
    while (row.length <= REQ.CHANGE_LOG) row.push("");

    // Apply each change
    const changed = [];
    for (const [field, colIdx] of Object.entries(editableFields)) {
      if (changes[field] !== undefined) {
        const oldVal = sanitize(row[colIdx]);
        const newVal = sanitize(String(changes[field] || ""));
        if (oldVal !== newVal) {
          changed.push(`${field}: "${oldVal}" → "${newVal}"`);
          row[colIdx] = newVal;
        }
      }
    }

    if (changed.length === 0) return json({ success: true, message: "No changes detected" });

    // Append to change log
    const logEntry = `[${new Date().toISOString().slice(0,16).replace("T"," ")} by ${sanitize(editedBy||"HR")}] ${sanitize(changeNote.trim())} | ${changed.join(" ; ")}`;
    const existing = sanitize(row[REQ.CHANGE_LOG] || "");
    row[REQ.CHANGE_LOG] = existing ? existing + " || " + logEntry : logEntry;

    await updateRow(token, env, SHEET_REQUISITIONS, i + 1, row);

    // Log activity
    const msToken = await getToken(env);
    await appendRow(msToken, env, SHEET_ACTIVITY, [
      new Date().toISOString(), sanitize(editedBy||"HR"),
      `Edited requisition ${requisitionID}: ${sanitize(changeNote.trim())}`, "hr",
    ]).catch(() => {});

    return json({ success: true, changeLog: row[REQ.CHANGE_LOG], changed });
  }
  return err("Requisition not found", 404);
}

// ── Fire notification to notify worker (fire-and-forget) ──────────────────
async function fireNotify(env, event, data) {
  if (!env.NOTIFY_WORKER_URL || !env.NOTIFY_SECRET) return;
  fetch(`${env.NOTIFY_WORKER_URL}/notify`, {
    method: "POST",
    headers: { "Content-Type":"application/json", "X-App-Key": env.APP_KEY },
    body: JSON.stringify({ secret: env.NOTIFY_SECRET, event, data }),
  }).catch(e => console.error("fireNotify failed:", e.message));
}

// ── POST /generate-share-link ─────────────────────────────────────────────
// Generates a tokenised management share link for a candidate.
// Stores token + expiry + visibility flags as JSON in C.SHARE_TOKEN (col AR).
// Token is valid for 30 days. Re-generating replaces the previous token.
async function handleGenerateShareLink(body, env) {
  const { candidateID, visibility } = body;
  if (!candidateID) return err("candidateID required");

  const token     = generateToken();
  const shortCode = generateShortCode();
  const expires   = new Date(Date.now() + 30 * 24 * 3600 * 1000).toISOString();

  const shareData = JSON.stringify({
    token,
    shortCode,
    expires,
    visibility: {
      salary:      visibility?.salary      !== false,
      visaDetails: visibility?.visaDetails !== false,
      scores:      visibility?.scores      !== false,
      feedback:    visibility?.feedback    !== false,
      mgmtRound:   visibility?.mgmtRound   !== false,
      cvLink:      visibility?.cvLink      !== false,
      journey:     visibility?.journey     !== false,
    },
  });

  const msToken = await getToken(env);
  const rows    = await getRows(msToken, env, SHEET_CANDIDATES);
  for (let i = 1; i < rows.length; i++) {
    if (sanitize(rows[i][C.ID]) !== candidateID) continue;
    const row = [...rows[i]];
    while (row.length <= C.SHARE_TOKEN) row.push("");
    row[C.SHARE_TOKEN] = shareData;
    await updateRow(msToken, env, SHEET_CANDIDATES, i + 1, row);
    return json({ success: true, token, shortCode, expires });
  }
  return err("Candidate not found", 404);
}

// ── GET /share-preview?t=TOKEN ────────────────────────────────────────────
// Server-side OG tag renderer for WhatsApp/Telegram/iMessage link previews.
// Messaging app crawlers fetch URLs before JS runs so client-side OG tag
// updates are invisible to them. This endpoint returns a minimal server-rendered
// HTML page with the correct og:title/og:description/og:url tags populated
// from the candidate's actual data, then immediately redirects the real user
// to the full share.html page via both meta-refresh and JS.
//
// Usage: share the URL https://hifive-recruitment.pages.dev/share-preview?t=TOKEN
// instead of /share?t=TOKEN. Previews will show the candidate's name and position.
async function handleSharePreview(token, env) {
  if (!token) return new Response("Token required", { status: 400 });

  const msToken = await getToken(env);
  const candRows = await getRows(msToken, env, SHEET_CANDIDATES).catch(() => []);

  let candRow = null, shareData = null;
  for (const r of candRows.slice(1)) {
    const raw = sanitize(r[C.SHARE_TOKEN] || "");
    if (!raw) continue;
    try {
      const parsed = JSON.parse(raw);
      if (parsed.token === token || parsed.shortCode === token) {
        candRow = r; shareData = parsed; break;
      }
    } catch {}
  }

  // Build the redirect URL — the full share page
  const shareUrl = `https://hifive-recruitment.pages.dev/share?t=${encodeURIComponent(token)}`;

  // If not found or expired, redirect straight to share.html which will show its own error
  if (!candRow || new Date(shareData?.expires) < new Date()) {
    return new Response(
      `<!DOCTYPE html><html><head><meta http-equiv="refresh" content="0;url=${shareUrl}"/></head><body><script>location.replace(${JSON.stringify(shareUrl)})</script></body></html>`,
      { status: 200, headers: { "Content-Type": "text/html;charset=utf-8" } }
    );
  }

  const name     = sanitize(candRow[C.NAME]);
  const position = sanitize(candRow[C.CUSTOM_POSITION]) || sanitize(candRow[C.POSITION]);
  const entity   = sanitize(candRow[C.ENTITY]);
  const status   = sanitize(candRow[C.STATUS]);

  const ogTitle = `${name} — ${position}`;
  const ogDesc  = `${entity} · ${status} · View full candidate profile on HiFive Recruitment`;
  const ogImage = `https://hifive-recruitment.pages.dev/icons/icon-512.png`; // portal PWA icon as fallback

  const html = `<!DOCTYPE html>
<html prefix="og: https://ogp.me/ns#">
<head>
  <meta charset="UTF-8"/>
  <title>${ogTitle} · HiFive Recruitment</title>

  <!-- OpenGraph — used by WhatsApp, LinkedIn, Facebook, Telegram -->
  <meta property="og:type"        content="profile"/>
  <meta property="og:title"       content="${ogTitle}"/>
  <meta property="og:description" content="${ogDesc}"/>
  <meta property="og:url"         content="${shareUrl}"/>
  <meta property="og:image"       content="${ogImage}"/>
  <meta property="og:site_name"   content="HiFive Recruitment"/>

  <!-- Twitter / iMessage -->
  <meta name="twitter:card"        content="summary"/>
  <meta name="twitter:title"       content="${ogTitle}"/>
  <meta name="twitter:description" content="${ogDesc}"/>
  <meta name="twitter:image"       content="${ogImage}"/>

  <!-- Redirect real users to the full share page immediately -->
  <meta http-equiv="refresh" content="0;url=${shareUrl}"/>
  <style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f7f3ef;color:#1a1218}</style>
</head>
<body>
  <div style="text-align:center;padding:40px">
    <div style="font-size:24px;font-weight:700;margin-bottom:8px">${ogTitle}</div>
    <div style="color:#8a8590;margin-bottom:24px">${ogDesc}</div>
    <a href="${shareUrl}" style="display:inline-block;padding:12px 28px;border-radius:99px;background:linear-gradient(135deg,#6e0b2a,#9b0d3a);color:#fff;font-weight:700;text-decoration:none">View Profile →</a>
  </div>
  <script>location.replace(${JSON.stringify(shareUrl)});</script>
</body>
</html>`;

  return new Response(html, {
    status: 200,
    headers: {
      "Content-Type": "text/html;charset=utf-8",
      "Cache-Control": "no-cache, no-store", // prevent stale preview cache
    },
  });
}

// ── GET /get-share?t=TOKEN ────────────────────────────────────────────────
// Public endpoint — no auth required. Returns candidate profile + pipeline
// + interview rounds filtered by the visibility flags stored at generation time.
async function handleGetShare(token, env) {
  if (!token) return err("Token required", 400);

  const msToken = await getToken(env);
  const [candRows, psRows, roundRows] = await Promise.all([
    getRows(msToken, env, SHEET_CANDIDATES),
    getRows(msToken, env, SHEET_PIPELINE),
    getRows(msToken, env, SHEET_ROUNDS),
  ]);

  // Find the candidate whose SHARE_TOKEN JSON contains this token
  let candRow = null, shareData = null;
  for (const r of candRows.slice(1)) {
    const raw = sanitize(r[C.SHARE_TOKEN] || "");
    if (!raw) continue;
    try {
      const parsed = JSON.parse(raw);
      if (parsed.token === token || parsed.shortCode === token) {
        candRow   = r;
        shareData = parsed;
        break;
      }
    } catch {}
  }

  if (!candRow) return err("Share link not found or has been revoked", 404);
  if (new Date(shareData.expires) < new Date()) return err("This share link has expired", 410);

  const v   = shareData.visibility;
  const cid = sanitize(candRow[C.ID]);

  // ── Candidate profile ──────────────────────────────────────────────────
  const profile = {
    id:              cid,
    name:            sanitize(candRow[C.NAME]),
    position:        sanitize(candRow[C.CUSTOM_POSITION]) || sanitize(candRow[C.POSITION]),
    entity:          sanitize(candRow[C.ENTITY]),
    nationality:     sanitize(candRow[C.NATIONALITY]),
    degree:          sanitize(candRow[C.DEGREE]),
    status:          sanitize(candRow[C.STATUS]),
    noticePeriod:    sanitize(candRow[C.NOTICE_PERIOD]),
    uaeExp:          sanitize(candRow[C.UAE_EXP]),
    availability:    sanitize(candRow[C.AVAILABILITY]),
    date:            excelDateToISO(candRow[C.DATE]) || sanitize(candRow[C.DATE]),
    currEmployer:    sanitize(candRow[C.CURR_EMPLOYER]),
    prevEmployer:    sanitize(candRow[C.PREV_EMPLOYER]),
    relation:        sanitize(candRow[C.RELATION]),
    location:        sanitize(candRow[C.LOCATION]),
    emirate:         sanitize(candRow[C.EMIRATE]),
    source:          sanitize(candRow[C.SOURCE]),
    interviewStatus: sanitize(candRow[C.INTERVIEW_STATUS]),
    interviewDate:   excelDateToISO(candRow[C.INTERVIEW_DATE]) || sanitize(candRow[C.INTERVIEW_DATE]),
    remarks:         sanitize(candRow[C.REMARKS]),
    cvLink:       v.cvLink ? sanitize(candRow[C.CV_LINK]) : null,
    cvFileName:   v.cvLink ? sanitize(candRow[C.CV_FILE_NAME] || "") : null,
    // Conditional fields
    ...(v.salary      ? { expectedSalary: sanitize(candRow[C.EXPECTED_SALARY]), offeredSalary: sanitize(candRow[C.OFFERED_SALARY]) } : {}),
    ...(v.visaDetails ? { visaType: sanitize(candRow[C.VISA_TYPE]), visaExpiry: excelDateToISO(candRow[C.VISA_EXPIRY]) || sanitize(candRow[C.VISA_EXPIRY]) } : {}),
  };

  // ── Stage journey ──────────────────────────────────────────────────────
  const journey = v.journey
    ? (() => {
        const psEntries = psRows.slice(1)
          .filter(r => sanitize(r[PS.CAND_ID]) === cid)
          .map(r => ({
            stage:       sanitize(r[PS.STAGE]),
            outcome:     sanitize(r[PS.OUTCOME]),
            date:        excelDateToISO(r[PS.DATE]) || sanitize(r[PS.DATE]),
            notes:       sanitize(r[PS.NOTES]),
            interviewer: sanitize(r[PS.INTERVIEWER]),
          }))
          .sort((a,b) => new Date(a.date) - new Date(b.date));

        // Synthesize "Added to Pipeline" entry from candidate's date field
        // if no existing HR Screen row already covers the entry date
        const entryDate = excelDateToISO(candRow[C.DATE]) || sanitize(candRow[C.DATE]);
        const hasInitialEntry = psEntries.some(e =>
          e.stage === "HR Screen" || e.stage === "Active" || e.stage === "Added"
        );
        if (entryDate && !hasInitialEntry) {
          psEntries.unshift({
            stage:       "HR Screening",
            outcome:     "Pending",
            date:        entryDate,
            notes:       sanitize(candRow[C.SOURCE] ? `Source: ${sanitize(candRow[C.SOURCE])}` : ""),
            interviewer: "",
          });
        }
        return psEntries;
      })()
    : [];

  // ── Interview rounds ───────────────────────────────────────────────────
  const MGMT_KEYWORDS = ["management","mgmt","ceo","director","gm","general manager"];
  const isMgmtRound = (label) => MGMT_KEYWORDS.some(k => (label||"").toLowerCase().includes(k));

  const rounds = roundRows.slice(1)
    .filter(r => sanitize(r[R.CAND_ID]) === cid)
    .filter(r => {
      const label = sanitize(r[R.ROUND] || "");
      if (!v.mgmtRound && isMgmtRound(label)) return false;
      return true;
    })
    .map(r => ({
      round:           sanitize(r[R.ROUND]),
      status:          sanitize(r[R.STATUS]),
      scheduledDate:   excelDateToISO(r[R.SCHED_DATE]),   // was raw serial — now ISO
      scheduledTime:   excelTimeToHHMM(r[R.SCHED_TIME]),
      interviewer:     sanitize(r[R.INTERVIEWER]),
      raterName:       sanitize(r[R.SUBMITTED_BY_NAME] || ""),
      team:            sanitize(r[R.TEAM]),
      outcome:         sanitize(r[R.OUTCOME]),
      passFail:        sanitize(r[R.RATING_PASS_FAIL] || ""),
      score:           v.scores ? sanitize(r[R.SCORE]) : null,
      feedback:        v.feedback ? sanitize(r[R.FEEDBACK] || "") : null,
      remarks:         v.feedback ? sanitize(r[R.RATING_REMARKS] || "") : null,
      notes:           v.feedback ? sanitize(r[R.NOTES] || "") : null,
      // Individual criteria scores (only if scores visible)
      criteria: v.scores ? {
        education:    sanitize(r[R.SCORE_EDUCATION]    || ""),
        workExp:      sanitize(r[R.SCORE_WORK_EXP]     || ""),
        technical:    sanitize(r[R.SCORE_TECHNICAL]    || ""),
        communication:sanitize(r[R.SCORE_COMMUNICATION]|| ""),
        enthusiasm:   sanitize(r[R.SCORE_ENTHUSIASM]   || ""),
        product:      sanitize(r[R.SCORE_PRODUCT]      || ""),
        teamwork:     sanitize(r[R.SCORE_TEAMWORK]     || ""),
        initiative:   sanitize(r[R.SCORE_INITIATIVE]   || ""),
        timeMgmt:     sanitize(r[R.SCORE_TIME_MGMT]    || ""),
        companyKnowledge: sanitize(r[R.SCORE_COMPANY_KNOWLEDGE] || ""),
        omitted:      sanitize(r[R.OMITTED_CRITERIA]   || ""),
      } : null,
    }));

  return json({
    profile,
    journey,
    rounds,
    visibility: v,
    generatedAt: new Date().toISOString(),
    expires: shareData.expires,
  });
}

// ── GET /get-pipeline-history ─────────────────────────────────────────────
// Returns avg days per stage and bottleneck analysis from Pipeline_History.
// Used by the Analytics screen to show where candidates get stuck.
async function handleGetPipelineHistory(env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_PIPELINE_HISTORY);
  if (rows.length <= 1) return json({ stages:[], bottlenecks:[], totalTransitions:0 });

  // Group by TO_STAGE and accumulate daysInPrev values
  const stageMap = {};
  const STAGE_ORDER = [
    "HR Screen","Shortlisted","Operations","Management",
    "Verbal Offer","Written Offer","Accepted","Onboarding","Joined","Rejected"
  ];

  let totalTransitions = 0;

  for (const r of rows.slice(1)) {
    const toStage   = sanitize(r[PH.TO_STAGE]);
    const fromStage = sanitize(r[PH.FROM_STAGE]);
    const days      = parseFloat(sanitize(r[PH.DAYS_IN_PREV] || ""));
    const outcome   = sanitize(r[PH.OUTCOME]);
    const ts        = sanitize(r[PH.TIMESTAMP]);

    if (!toStage) continue;
    totalTransitions++;

    // Key by the stage candidates left (fromStage is where they spent time)
    const key = fromStage || toStage;
    if (!stageMap[key]) stageMap[key] = {
      stage: key,
      transitions: 0,
      passes: 0,
      rejections: 0,
      dayValues: [],
      lastTransition: "",
    };

    const g = stageMap[key];
    g.transitions++;
    if (outcome === "Pass" || outcome === "Pending") g.passes++;
    if (outcome === "Rejected" || outcome === "Fail") g.rejections++;
    if (!isNaN(days) && days >= 0 && days <= 365) g.dayValues.push(days);
    if (ts > g.lastTransition) g.lastTransition = ts;
  }

  const stages = Object.values(stageMap).map(g => {
    const avgDays = g.dayValues.length
      ? Math.round((g.dayValues.reduce((s,v)=>s+v,0) / g.dayValues.length) * 10) / 10
      : null;
    const passRate = g.transitions > 0
      ? Math.round((g.passes / g.transitions) * 100)
      : null;
    const order = STAGE_ORDER.indexOf(g.stage);
    return {
      stage:       g.stage,
      order:       order >= 0 ? order : 99,
      transitions: g.transitions,
      passes:      g.passes,
      rejections:  g.rejections,
      passRate,
      avgDays,
      lastTransition: g.lastTransition,
    };
  }).sort((a,b) => a.order - b.order);

  // Bottlenecks = stages with above-average days or below-average pass rate
  const avgOfAvgs = (() => {
    const vals = stages.filter(s=>s.avgDays!=null).map(s=>s.avgDays);
    return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : 0;
  })();
  const avgPassRate = (() => {
    const vals = stages.filter(s=>s.passRate!=null).map(s=>s.passRate);
    return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : 0;
  })();

  const bottlenecks = stages
    .filter(s => (s.avgDays != null && s.avgDays > avgOfAvgs * 1.3) ||
                 (s.passRate != null && s.passRate < avgPassRate * 0.7 && s.transitions >= 3))
    .map(s => ({
      ...s,
      reason: s.avgDays > avgOfAvgs * 1.3
        ? `Avg ${s.avgDays}d — ${Math.round(s.avgDays/avgOfAvgs*100-100)}% slower than average`
        : `Pass rate ${s.passRate}% — ${Math.round(avgPassRate-s.passRate)}% below average`,
    }));

  return json({ stages, bottlenecks, totalTransitions, avgOfAvgs: Math.round(avgOfAvgs*10)/10 });
}

// ── GET /get-quotes ─────────────────────────────────────────────────────────
// Public endpoint — returns all rows from the Quotes sheet.
// Excel schema (A–E): TYPE | TEXT | AUTHOR | CATEGORY | ACTIVE
// TYPE: "quote" or "wish"
// ACTIVE: "yes" / "true" / "1" to show; anything else to hide
async function handleGetQuotes(env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_QUOTES);
  if (rows.length <= 1) return json({ quotes:[], wishes:[] });

  const quotes = [], wishes = [];
  for (const r of rows.slice(1)) {
    const type     = sanitize(r[0] || "").toLowerCase();
    const text     = sanitize(r[1] || "");
    const author   = sanitize(r[2] || "");
    const category = sanitize(r[3] || "");
    const active   = sanitize(r[4] || "yes").toLowerCase();
    if (!text) continue;
    if (!["yes","true","1",""].includes(active)) continue; // skip if explicitly inactive
    const entry = { text, author, category };
    if (type === "wish" || type === "wishes") wishes.push(entry);
    else quotes.push(entry);
  }
  return json({ quotes, wishes });
}

// ── POST /rate-candidate-ai ───────────────────────────────────────────────
// Uses Cloudflare Workers AI (free, bundled with the Workers platform —
// no API key or subscription required) to score a candidate against the
// requisition's job description. Writes the score + rationale to Excel
// (cols AS/AT) so it persists and shows on the Kanban card without
// re-running the AI call every time the board loads.
//
// Requires the "ai" binding in wrangler.toml:
//   [ai]
//   binding = "AI"
async function handleRateCandidateAI(body, env) {
  const { candidateID } = body;
  if (!candidateID) return err("candidateID required");
  if (!env.AI) return err("Workers AI binding not configured. Add [ai] binding = \"AI\" to wrangler.toml", 500);

  const token = await getToken(env);
  const candRows = await getRows(token, env, SHEET_CANDIDATES);
  const reqRows  = await getRows(token, env, SHEET_REQUISITIONS);

  let candRow = null, excelRow = -1;
  for (let i = 1; i < candRows.length; i++) {
    if (sanitize(candRows[i][C.ID]) === candidateID) { candRow = candRows[i]; excelRow = i + 1; break; }
  }
  if (!candRow) return err("Candidate not found", 404);

  // Find the matching requisition for job description context
  const reqID = sanitize(candRow[C.REQUISITION_ID] || "");
  let jobDescription = "";
  let reqPosition = sanitize(candRow[C.CUSTOM_POSITION]) || sanitize(candRow[C.POSITION]);
  if (reqID) {
    const reqRow = reqRows.slice(1).find(r => sanitize(r[REQ.ID]) === reqID);
    if (reqRow) jobDescription = sanitize(reqRow[REQ.JOB_DESCRIPTION] || "");
  }
  // Fallback: match by position name if no direct requisition link
  if (!jobDescription) {
    const reqRow = reqRows.slice(1).find(r =>
      (sanitize(r[REQ.CUSTOM_POSITION]) || sanitize(r[REQ.POSITION])) === reqPosition &&
      sanitize(r[REQ.OVERALL_STATUS]) === "Approved"
    );
    if (reqRow) jobDescription = sanitize(reqRow[REQ.JOB_DESCRIPTION] || "");
  }

  // Build candidate summary for the prompt
  const candSummary = {
    name:         sanitize(candRow[C.NAME]),
    position:     reqPosition,
    degree:       sanitize(candRow[C.DEGREE]),
    uaeExp:       sanitize(candRow[C.UAE_EXP]),
    nationality:  sanitize(candRow[C.NATIONALITY]),
    noticePeriod: sanitize(candRow[C.NOTICE_PERIOD]),
    prevEmployer: sanitize(candRow[C.PREV_EMPLOYER] || ""),
    currEmployer: sanitize(candRow[C.CURR_EMPLOYER] || ""),
    remarks:      sanitize(candRow[C.REMARKS] || ""),
  };

  const stripHtml = (s) => s.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
  const jdText = jobDescription ? stripHtml(jobDescription).slice(0, 2000) : "(No job description on file — rate based on position title and general fit only.)";

  const prompt = `You are an HR recruiter assistant. Score how well this candidate matches the job description, from 0-100.

JOB: ${candSummary.position}
JOB DESCRIPTION:
${jdText}

CANDIDATE:
Name: ${candSummary.name}
Degree/Qualification: ${candSummary.degree || "Not specified"}
UAE Experience: ${candSummary.uaeExp || "Not specified"}
Nationality: ${candSummary.nationality || "Not specified"}
Notice Period: ${candSummary.noticePeriod || "Not specified"}
Previous Employer: ${candSummary.prevEmployer || "Not specified"}
Current Employer: ${candSummary.currEmployer || "Not specified"}
Additional Notes: ${candSummary.remarks || "None"}

Respond with ONLY valid JSON in this exact format, nothing else:
{"score": <number 0-100>, "reason": "<one sentence, max 25 words, explaining the score>"}`;

  try {
    const aiResponse = await env.AI.run("@cf/meta/llama-3.1-8b-instruct", {
      messages: [
        { role: "system", content: "You are a precise HR scoring assistant. You only output valid JSON, no markdown, no explanation outside the JSON." },
        { role: "user", content: prompt },
      ],
      max_tokens: 150,
    });

    const raw = aiResponse.response || "";
    // Extract JSON from response (model may wrap it in markdown fences)
    const jsonMatch = raw.match(/\{[^}]+\}/);
    if (!jsonMatch) throw new Error("AI did not return valid JSON: " + raw.slice(0, 100));

    const parsed = JSON.parse(jsonMatch[0]);
    const score  = Math.max(0, Math.min(100, parseInt(parsed.score) || 0));
    const reason = sanitize(parsed.reason || "").slice(0, 200);

    await updateCell(token, env, SHEET_CANDIDATES, excelRow, C.AI_RATING, score);
    await updateCell(token, env, SHEET_CANDIDATES, excelRow, C.AI_RATING_REASON, reason);

    return json({ success: true, score, reason });
  } catch(e) {
    console.error("AI rating failed:", e.message);
    return err("AI rating failed: " + e.message, 500);
  }
}

// ── POST /rate-candidates-ai-bulk ─────────────────────────────────────────
// Rates all un-rated candidates for a given requisition in one call.
// Useful for rating an entire Kanban column at once.
async function handleRateCandidatesAIBulk(body, env) {
  const { requisitionID, position, force } = body;
  if (!requisitionID && !position) return err("requisitionID or position required");
  if (!env.AI) return err("Workers AI binding not configured", 500);

  const token = await getToken(env);
  const candRows = await getRows(token, env, SHEET_CANDIDATES);

  const targets = [];
  for (let i = 1; i < candRows.length; i++) {
    const r = candRows[i];
    if (requisitionID && sanitize(r[C.REQUISITION_ID]) !== requisitionID) continue;
    if (position && (sanitize(r[C.CUSTOM_POSITION]) || sanitize(r[C.POSITION])) !== position) continue;
    if (!force && sanitize(r[C.AI_RATING])) continue; // skip already-rated unless force
    targets.push(sanitize(r[C.ID]));
  }

  const results = [];
  for (const candID of targets.slice(0, 20)) { // cap at 20 per call to avoid timeout
    try {
      const res = await handleRateCandidateAI({ candidateID: candID }, env);
      const data = await res.json();
      results.push({ candidateID: candID, ...data });
    } catch(e) {
      results.push({ candidateID: candID, success: false, error: e.message });
    }
  }

  return json({ rated: results.length, remaining: Math.max(0, targets.length - 20), results });
}

// ── POST /backfill-time-to-hire ───────────────────────────────────────────
// One-time fix for candidates marked Joined before the TIME_TO_HIRE bug was
// fixed. Finds every Joined candidate with a blank TIME_TO_HIRE but a valid
// ACTUAL_JOINING + entry date, and computes it retroactively.
// Safe to run multiple times — only touches blank cells.
async function handleBackfillTimeToHire(env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_CANDIDATES);
  let fixed = 0, skipped = 0, repaired = 0;
  const details = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (sanitize(r[C.STATUS]) !== "Joined") continue;

    const candID   = sanitize(r[C.ID]);
    const candName = sanitize(r[C.NAME]);
    const existingTTH = sanitize(r[C.TIME_TO_HIRE]);
    const existingTTHNum = parseFloat(existingTTH);

    // Repair corrupted values from the raw-Excel-serial-date bug: a sane
    // TIME_TO_HIRE is a small number (0-3650 days, ~10 years max). Anything
    // wildly outside that range (e.g. -16132217) came from new Date() being
    // fed a raw Excel serial number instead of an ISO string, and needs to
    // be recomputed rather than treated as "already has a value."
    const isCorrupted = existingTTH && (!isFinite(existingTTHNum) || Math.abs(existingTTHNum) > 3650);

    if (existingTTH && !isCorrupted) {
      skipped++;
      details.push({ candID, candName, reason: "already has valid TIME_TO_HIRE", value: existingTTH });
      continue;
    }

    const rawJoining = r[C.ACTUAL_JOINING];
    const rawEntry   = r[C.DATE];
    const joiningISO = excelDateToISO(rawJoining) || sanitize(rawJoining);
    const entryISO    = excelDateToISO(rawEntry) || sanitize(rawEntry);

    if (!joiningISO || !entryISO) {
      skipped++;
      details.push({ candID, candName, reason: isCorrupted ? "corrupted value, but missing date to recompute" : "missing date", rawJoining: sanitize(rawJoining), rawEntry: sanitize(rawEntry), joiningISO, entryISO, corruptedValue: isCorrupted ? existingTTH : undefined });
      continue;
    }

    const diff = Math.round((new Date(joiningISO) - new Date(entryISO)) / 86400000);
    if (isNaN(diff) || diff < 0 || diff > 3650) {
      skipped++;
      details.push({ candID, candName, reason: "invalid diff", joiningISO, entryISO, diff });
      continue;
    }

    await updateCell(token, env, SHEET_CANDIDATES, i + 1, C.TIME_TO_HIRE, diff);
    if (isCorrupted) { repaired++; details.push({ candID, candName, reason: "repaired", oldValue: existingTTH, newValue: diff }); }
    else { fixed++; details.push({ candID, candName, reason: "fixed", diff }); }
  }

  return json({ fixed, repaired, skipped, total: rows.length - 1, details });
}

// ── GET /check-whatsapp-setup ─────────────────────────────────────────────
// Diagnostic endpoint — shows which users are missing WhatsApp numbers
// in the Users sheet, so HR knows exactly which rows to fill in.
async function handleCheckWhatsappSetup(env) {
  const token = await getToken(env);
  const rows  = await getRows(token, env, SHEET_USERS).catch(() => [[]]);
  const report = rows.slice(1)
    .filter(r => r[U.KEY] && sanitize(r[U.ACTIVE]).toLowerCase() !== "false")
    .map(r => ({
      key:        sanitize(r[U.KEY]),
      name:       sanitize(r[U.NAME]),
      role:       sanitize(r[U.ROLE]),
      whatsapp:   sanitize(r[U.WHATSAPP] || ""),
      email:      sanitize(r[U.EMAIL] || ""),
      hasWA:      !!sanitize(r[U.WHATSAPP] || ""),
      hasEmail:   !!sanitize(r[U.EMAIL] || ""),
    }));

  const missing = report.filter(u => !u.hasWA);
  return json({
    total: report.length,
    missingWhatsapp: missing.length,
    users: report,
    note: missing.length > 0
      ? `${missing.length} user(s) have no WhatsApp number. Add 971XXXXXXXXX format to column I (WHATSAPP) of the Users sheet for: ${missing.map(u=>u.name).join(", ")}`
      : "All users have WhatsApp numbers configured."
  });
}
