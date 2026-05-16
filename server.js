const path = require("path");
const express = require("express");
const cookieParser = require("cookie-parser");
const multer = require("multer");
const bcrypt = require("bcryptjs");
const crypto = require("crypto");
const dotenv = require("dotenv");
const { EventEmitter } = require("events");

dotenv.config({ path: path.join(__dirname, ".env.local") });

const {
  initDatabase,
  getDb,
  checkpointDatabase,
  closeDatabase,
  createDatabaseBackup,
  databasePath,
  backupDir,
} = require("./src/db");
const { parseMetadataFile, buildSystemDataExport } = require("./src/excel");
const { classifyRecords, classifyRecordLocally } = require("./src/classifier");

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 30 * 1024 * 1024 },
});

const PORT = Number(process.env.PORT || 3000);
const SESSION_COOKIE_NAME = process.env.SESSION_COOKIE_NAME || "dcp_session";
const SESSION_DAYS = 7;
const SESSION_COOKIE_OPTIONS = {
  httpOnly: true,
  sameSite: "lax",
  path: "/",
};
const activeClassificationJobs = new Map();
const classificationJobEvents = new EventEmitter();
classificationJobEvents.setMaxListeners(200);

app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use((req, res, next) => {
  if (req.method === "GET" && (req.path === "/" || req.path.startsWith("/app"))) {
    res.setHeader("Cache-Control", "no-store");
  }
  next();
});
app.use(express.static(path.join(__dirname, "public"), { index: false }));

function nowIso() {
  return new Date().toISOString();
}

function normalize(value) {
  return String(value || "").trim();
}

function normalizeEmail(value) {
  return normalize(value).toLowerCase();
}

function hashToken(token) {
  return crypto.createHash("sha256").update(token).digest("hex");
}

function publicUser(user) {
  if (!user) return null;
  return {
    id: user.id,
    name: user.name,
    email: user.email,
    accountType: user.accountType,
    companyName: user.companyName || "",
    role: user.role || "User",
    avatarInitials: user.avatarInitials || initialsFor(user.name, user.email),
  };
}

function initialsFor(name, email) {
  const source = normalize(name) || normalize(email);
  return source
    .split(/\s+|@/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0].toUpperCase())
    .join("") || "U";
}

function createSession(res, userId) {
  const db = getDb();
  const token = crypto.randomBytes(32).toString("hex");
  const expiresAt = new Date(Date.now() + SESSION_DAYS * 24 * 60 * 60 * 1000).toISOString();
  db.prepare(
    "INSERT INTO sessions (id, userId, createdAt, expiresAt) VALUES (?, ?, ?, ?)"
  ).run(token, userId, nowIso(), expiresAt);
  checkpointDatabase("FULL");
  res.cookie(SESSION_COOKIE_NAME, token, {
    ...SESSION_COOKIE_OPTIONS,
    maxAge: SESSION_DAYS * 24 * 60 * 60 * 1000,
  });
}

function clearSession(req, res) {
  const db = getDb();
  const token = req.cookies?.[SESSION_COOKIE_NAME];
  if (token) {
    db.prepare("DELETE FROM sessions WHERE id = ?").run(token);
    checkpointDatabase("FULL");
  }
  res.clearCookie(SESSION_COOKIE_NAME, SESSION_COOKIE_OPTIONS);
  res.cookie(SESSION_COOKIE_NAME, "", {
    ...SESSION_COOKIE_OPTIONS,
    expires: new Date(0),
    maxAge: 0,
  });
}

function authOptional(req, res, next) {
  const db = getDb();
  const token = req.cookies?.[SESSION_COOKIE_NAME];
  if (!token) {
    req.user = null;
    next();
    return;
  }

  const session = db
    .prepare(
      `SELECT sessions.*, users.name, users.email, users.accountType, users.companyName,
              users.role, users.avatarInitials
         FROM sessions
         JOIN users ON users.id = sessions.userId
        WHERE sessions.id = ? AND sessions.expiresAt > ?`
    )
    .get(token, nowIso());

  if (!session) {
    db.prepare("DELETE FROM sessions WHERE id = ? OR expiresAt <= ?").run(token, nowIso());
    checkpointDatabase("FULL");
    res.clearCookie(SESSION_COOKIE_NAME, SESSION_COOKIE_OPTIONS);
    req.user = null;
    next();
    return;
  }

  req.user = {
    id: session.userId,
    name: session.name,
    email: session.email,
    accountType: session.accountType,
    companyName: session.companyName,
    role: session.role,
    avatarInitials: session.avatarInitials,
  };
  next();
}

function requireAuth(req, res, next) {
  authOptional(req, res, () => {
    if (!req.user) {
      res.status(401).json({ error: "Authentication required" });
      return;
    }
    next();
  });
}

function requirePageAuth(req, res, next) {
  authOptional(req, res, () => {
    if (!req.user) {
      res.setHeader("Cache-Control", "no-store");
      res.redirect(302, "/");
      return;
    }
    res.setHeader("Cache-Control", "no-store");
    next();
  });
}

function requireSystemAccess(req, res, next) {
  const db = getDb();
  const system = db.prepare("SELECT * FROM systems WHERE id = ?").get(req.params.systemId || req.params.id);
  if (!system) {
    res.status(404).json({ error: "System not found" });
    return;
  }
  if (!canAccessSystem(req.user, system)) {
    res.status(403).json({ error: "You do not have access to this system." });
    return;
  }
  req.system = system;
  next();
}

function canAccessSystem(user, system) {
  return Boolean(user && system && (system.userId === user.id || user.email === "admin@local"));
}

function recordToApi(row) {
  const original = safeJson(row.originalJson, {});
  return {
    id: row.id,
    systemId: row.systemId,
    rowIndex: row.rowIndex,
    original,
    tableName: row.tableName,
    columnName: row.columnName,
    dataType: row.dataType || original.DataType || original["Data Type"] || "",
    uploadedFileName: row.uploadedFileName || "",
    confidentiality: row.confidentiality || "Pending",
    confReason: row.confReason || "",
    personalData: row.personalData || "No",
    personalReason: row.personalReason || "",
    personalDataType: row.personalDataType || "",
    pseudonymizable: row.pseudonymizable || "No",
    anonymizable: row.anonymizable || "No",
    specialCategory: row.specialCategory || "No",
    confidenceScore: row.confidenceScore == null ? null : Number(row.confidenceScore),
    reviewStatus: row.reviewStatus || "Unreviewed",
    systemReviewStatus: row.systemReviewStatus || row.reviewStatus || "Unreviewed",
    personalReviewStatus: row.personalReviewStatus || "Needs Review",
    personalApprovedAt: row.personalApprovedAt || "",
    personalApprovedBy: row.personalApprovedBy || "",
    needsReview: Number(row.needsReview || 0),
    policyRecommendation: row.policyRecommendation || "",
    owner: row.owner || "",
    steward: row.steward || "",
    reviewer: row.reviewer || "",
    pushToClient: Number(row.pushToClient || 0),
    auditTrail: row.auditTrail || "",
    lastModifiedBy: row.lastModifiedBy || "",
    lastReviewedAt: row.lastReviewedAt || "",
    updatedAt: row.updatedAt,
  };
}

function safeJson(value, fallback) {
  try {
    return value ? JSON.parse(value) : fallback;
  } catch (_error) {
    return fallback;
  }
}

function getOriginalColumns(systemId) {
  const db = getDb();
  const rows = db
    .prepare("SELECT originalJson FROM data_records WHERE systemId = ? ORDER BY rowIndex LIMIT 500")
    .all(systemId);
  const columns = [];
  const seen = new Set();
  for (const row of rows) {
    const original = safeJson(row.originalJson, {});
    for (const key of Object.keys(original)) {
      if (!seen.has(key)) {
        seen.add(key);
        columns.push(key);
      }
    }
  }
  return columns;
}

// Columns that exist as real DB columns and can be used in SQL ORDER BY / WHERE
// without risk of SQL injection (we never interpolate user input directly into
// SQL — only whitelisted names reach the ORDER BY clause).
const SORTABLE_DB_COLUMNS = new Set([
  "rowIndex", "tableName", "columnName", "dataType",
  "confidentiality", "confReason", "personalData", "personalReason", "personalDataType",
  "pseudonymizable", "anonymizable", "specialCategory", "confidenceScore",
  "reviewStatus", "systemReviewStatus", "personalReviewStatus",
  "personalApprovedAt", "personalApprovedBy",
  "needsReview", "pushToClient", "policyRecommendation",
  "owner", "steward", "reviewer", "uploadedFileName",
  "auditTrail", "lastModifiedBy", "lastReviewedAt",
  "createdAt", "updatedAt",
]);

function listRecords(systemId, options = {}) {
  const db = getDb();
  const search      = normalize(options.search).toLowerCase();
  const rawSortBy   = normalize(options.sortBy) || "rowIndex";
  const sortDir     = normalize(options.sortDir).toLowerCase() === "desc" ? "DESC" : "ASC";
  const personalOnly = options.personalOnly === true || options.personalOnly === "true";
  const page         = Math.max(1, Number(options.page || 1));
  const pageSize     = Math.min(100, Math.max(1, Number(options.pageSize || 25)));
  const offset       = (page - 1) * pageSize;

  // ── Fast SQL path ──────────────────────────────────────────────────────────
  // When there is no free-text search and the sort column is a real DB column
  // we push everything — filtering, sorting, pagination — into SQLite.  For a
  // system with thousands of records this avoids loading all rows into Node.
  if (!search && SORTABLE_DB_COLUMNS.has(rawSortBy)) {
    const whereParts = ["systemId = ?"];
    const baseParams = [systemId];
    if (personalOnly) whereParts.push("personalData = 'Yes'");
    const where = `WHERE ${whereParts.join(" AND ")}`;

    const { total } = db
      .prepare(`SELECT COUNT(*) AS total FROM data_records ${where}`)
      .get(...baseParams);

    const rows = db
      .prepare(
        `SELECT * FROM data_records ${where}
         ORDER BY ${rawSortBy} ${sortDir}
         LIMIT ? OFFSET ?`
      )
      .all(...baseParams, pageSize, offset)
      .map(recordToApi);

    return {
      rows,
      // allRows is meaningful only for internal callers that pass a huge
      // pageSize (≥ total records).  For normal paginated API calls it equals
      // the current page, which is all callers outside of classificationJobTargets
      // need (see that function's refactor below).
      allRows: rows,
      total: Number(total || 0),
      page,
      pageSize,
      totalPages: Math.max(1, Math.ceil(Number(total || 0) / pageSize)),
    };
  }

  // ── JS fallback path ───────────────────────────────────────────────────────
  // Required when: (a) free-text search is active (must scan originalJson), or
  // (b) the sort column comes from the uploaded file's originalJson.
  let rows = db
    .prepare("SELECT * FROM data_records WHERE systemId = ? ORDER BY rowIndex ASC")
    .all(systemId)
    .map(recordToApi);

  if (personalOnly) rows = rows.filter((row) => row.personalData === "Yes");

  if (search) {
    rows = rows.filter((row) => {
      const haystack = [
        row.tableName, row.columnName, row.dataType,
        row.confidentiality, row.personalDataType, row.policyRecommendation,
        row.owner, row.steward, row.reviewer,
        JSON.stringify(row.original),
      ].join(" ").toLowerCase();
      return haystack.includes(search);
    });
  }

  const jsSortDir = sortDir === "DESC" ? -1 : 1;
  rows.sort((a, b) => {
    const aValue = valueForSort(a, rawSortBy);
    const bValue = valueForSort(b, rawSortBy);
    if (typeof aValue === "number" && typeof bValue === "number") return (aValue - bValue) * jsSortDir;
    return String(aValue || "").localeCompare(String(bValue || "")) * jsSortDir;
  });

  const total   = rows.length;
  const allRows = rows;
  return {
    rows: rows.slice(offset, offset + pageSize),
    allRows,
    total,
    page,
    pageSize,
    totalPages: Math.max(1, Math.ceil(total / pageSize)),
  };
}

function valueForSort(row, sortBy) {
  if (Object.prototype.hasOwnProperty.call(row, sortBy)) return row[sortBy];
  if (row.original && Object.prototype.hasOwnProperty.call(row.original, sortBy)) return row.original[sortBy];
  return "";
}

// ── getSystemSummary ─────────────────────────────────────────────────────────
// Uses SQL aggregates so it NEVER loads every record row into Node memory.
// Previously this fetched all rows for every API response, which was the main
// performance bottleneck for large systems.
function getSystemSummary(systemId) {
  const db = getDb();

  // Single-pass aggregate over data_records for the system.
  const totals = db.prepare(`
    SELECT
      COUNT(*)                                                                AS total,
      SUM(CASE WHEN confidentiality IS NOT NULL
                AND confidentiality != 'Pending'        THEN 1 ELSE 0 END)  AS classified,
      SUM(CASE WHEN personalData = 'Yes'                THEN 1 ELSE 0 END)  AS personal,
      SUM(CASE WHEN confidentiality IS NULL
                 OR confidentiality = 'Pending'         THEN 1 ELSE 0 END)  AS pending,
      SUM(CASE WHEN confidenceScore IS NOT NULL
               AND confidenceScore > 0
               AND confidenceScore < 0.75               THEN 1 ELSE 0 END)  AS lowConfidence,
      SUM(CASE WHEN policyRecommendation IS NOT NULL
               AND TRIM(policyRecommendation) != ''     THEN 1 ELSE 0 END)  AS policyRecommendationCount,
      SUM(CASE WHEN
        (COALESCE(systemReviewStatus, 'Unreviewed') != 'Approved'
          AND (needsReview = 1
               OR (confidenceScore IS NOT NULL AND confidenceScore > 0 AND confidenceScore < 0.75)))
        OR (personalData = 'Yes'
            AND COALESCE(personalReviewStatus, 'Needs Review') != 'Approved')
      THEN 1 ELSE 0 END)                                                     AS reviewQueue
    FROM data_records WHERE systemId = ?
  `).get(systemId);

  // Confidentiality distribution — one row per level.
  const confRows = db.prepare(`
    SELECT COALESCE(confidentiality, 'Pending') AS lvl, COUNT(*) AS cnt
    FROM   data_records
    WHERE  systemId = ?
    GROUP  BY COALESCE(confidentiality, 'Pending')
  `).all(systemId);

  // Count distinct tables that contain personal data.
  const { tablesWithPersonalData } = db.prepare(`
    SELECT COUNT(DISTINCT tableName) AS tablesWithPersonalData
    FROM   data_records
    WHERE  systemId = ? AND personalData = 'Yes'
  `).get(systemId);

  // Personal data type breakdown.
  const pdtRows = db.prepare(`
    SELECT COALESCE(NULLIF(TRIM(personalDataType), ''), 'Personal Data') AS type,
           COUNT(*) AS cnt
    FROM   data_records
    WHERE  systemId = ? AND personalData = 'Yes'
    GROUP  BY COALESCE(NULLIF(TRIM(personalDataType), ''), 'Personal Data')
    ORDER  BY cnt DESC
  `).all(systemId);

  const total      = Number(totals?.total                  || 0);
  const classified = Number(totals?.classified             || 0);
  const personal   = Number(totals?.personal               || 0);
  const pending    = Number(totals?.pending                || 0);

  // Build confidentiality map — always include every level so charts render.
  const confidentiality = { Confidential: 0, Secret: 0, "Top Secret": 0, Public: 0, Pending: 0 };
  for (const row of confRows) {
    const key = row.lvl || "Pending";
    if (Object.prototype.hasOwnProperty.call(confidentiality, key)) {
      confidentiality[key] = Number(row.cnt);
    }
  }

  // Personal data types — merge any duplicate labels (e.g. same type, two spellings).
  const personalDataTypes = {};
  for (const row of pdtRows) {
    const label = String(row.type || "Personal Data");
    personalDataTypes[label] = (personalDataTypes[label] || 0) + Number(row.cnt);
  }

  return {
    totalRecords: total,
    classified,
    classifiedPercentage: pct(classified, total),
    personal,
    personalPercentage: pct(personal, total),
    pending,
    pendingPercentage: pct(pending, total),
    reviewQueue:               Number(totals?.reviewQueue              || 0),
    lowConfidence:             Number(totals?.lowConfidence            || 0),
    policyRecommendationCount: Number(totals?.policyRecommendationCount || 0),
    tablesWithPersonalData:    Number(tablesWithPersonalData           || 0),
    confidentiality,
    personalDistribution: {
      "No Personal Data": total - personal,
      "Has Personal Data": personal,
    },
    personalDataTypes,
    classificationProgress: pct(classified, total),
    completedSystems: pending === 0 && total > 0 ? 1 : 0,
  };
}

function classificationNeedsReview(result) {
  return (
    Number(result.confidenceScore || 0) < 0.75 ||
    result.personalData === "Yes" ||
    /review needed|restricted access|consent required/i.test(result.policyRecommendation || "")
  );
}

function updateSystemClassificationState(systemId) {
  const db = getDb();
  const summary = getSystemSummary(systemId);
  db.prepare("UPDATE systems SET updatedAt = ?, status = ?, lastModifiedBy = ? WHERE id = ?").run(
    nowIso(),
    summary.pending === 0 && summary.totalRecords > 0 ? "Completed" : "In Progress",
    "AI Classification Service",
    systemId
  );
  return summary;
}

function saveClassificationResult(db, recordId, result) {
  const needsReview = classificationNeedsReview(result);
  const classifiedAt = nowIso();
  db.prepare(
    `UPDATE data_records SET
       confidentiality = ?, confReason = ?, personalData = ?, personalReason = ?,
       personalDataType = ?, pseudonymizable = ?, anonymizable = ?, specialCategory = ?,
       confidenceScore = ?, policyRecommendation = ?, needsReview = ?,
       reviewStatus = CASE WHEN reviewStatus = 'Approved' THEN reviewStatus ELSE 'Pending Review' END,
       auditTrail = ?, updatedAt = ?, lastModifiedBy = ?
     WHERE id = ?`
  ).run(
    result.confidentiality,
    result.reason,
    result.personalData,
    result.personalReason,
    result.personalDataType,
    result.pseudonymizable,
    result.anonymizable,
    result.specialCategory,
    result.confidenceScore,
    result.policyRecommendation,
    needsReview ? 1 : 0,
    `Classified via ${result.source || "AI"} at ${classifiedAt}`,
    classifiedAt,
    "AI Classification Service",
    recordId
  );
  return db.prepare("SELECT * FROM data_records WHERE id = ?").get(recordId);
}

function classificationJobToApi(row) {
  if (!row) return null;
  return {
    id: row.id,
    systemId: row.systemId,
    mode: row.mode,
    status: row.status,
    total: Number(row.total || 0),
    processed: Number(row.processed || 0),
    pageSize: Number(row.pageSize || 25),
    currentPage: Number(row.currentPage || 1),
    currentRecordId: row.currentRecordId || null,
    currentRowIndex: row.currentRowIndex || null,
    options: safeJson(row.optionsJson, {}),
    warning: safeJson(row.warningJson, null),
    errorMessage: row.errorMessage || "",
    createdBy: row.createdBy || "",
    startedAt: row.startedAt || "",
    completedAt: row.completedAt || "",
    failedAt: row.failedAt || "",
    createdAt: row.createdAt,
    updatedAt: row.updatedAt,
    percentage: pct(Number(row.processed || 0), Number(row.total || 0)),
  };
}

function getClassificationJob(db, jobId) {
  return db.prepare("SELECT * FROM classification_jobs WHERE id = ?").get(jobId);
}

function getClassificationJobPayload(jobId, record = null, summary = undefined) {
  const db = getDb();
  const job = getClassificationJob(db, jobId);
  if (!job) return null;
  let currentRecord = record;
  if (!currentRecord && job.currentRecordId) {
    const row = db.prepare("SELECT * FROM data_records WHERE id = ?").get(job.currentRecordId);
    currentRecord = row ? recordToApi(row) : null;
  }
  return {
    job: classificationJobToApi(job),
    record: currentRecord,
    summary: summary === undefined ? getSystemSummary(job.systemId) : summary,
  };
}

function emitClassificationJobEvent(jobId, event, payload) {
  classificationJobEvents.emit(`classification-job:${jobId}`, event, payload);
}

function classificationJobPage(job, record) {
  if (job.mode !== "all") {
    return Math.max(1, Number(safeJson(job.optionsJson, {}).page || 1));
  }
  return Math.max(1, Math.ceil(Number(record.rowIndex || 1) / Math.max(1, Number(job.pageSize || 25))));
}

function classificationJobTargets(systemId, job) {
  const options = safeJson(job.optionsJson, {});
  const mode = job.mode === "all" ? "all" : "page";

  if (mode === "all") {
    // Fetch every record for this system directly — no pagination needed.
    // This avoids the allRows / pageSize mismatch that existed when going
    // through listRecords with an arbitrarily small pageSize.
    const db = getDb();
    return db
      .prepare("SELECT * FROM data_records WHERE systemId = ? ORDER BY rowIndex ASC")
      .all(systemId)
      .map(recordToApi);
  }

  // Page mode — listRecords returns exactly the page the job was created for.
  const recordSet = listRecords(systemId, {
    page:     options.page     || 1,
    pageSize: job.pageSize     || options.pageSize || 25,
    search:   options.search   || "",
    sortBy:   options.sortBy   || "rowIndex",
    sortDir:  options.sortDir  || "asc",
  });
  return recordSet.rows.slice().sort((a, b) => a.rowIndex - b.rowIndex);
}

function classificationJobCanContinue(db, jobId, systemId) {
  const job = db.prepare("SELECT id FROM classification_jobs WHERE id = ?").get(jobId);
  const system = db.prepare("SELECT id FROM systems WHERE id = ?").get(systemId);
  return Boolean(job && system);
}

function findActiveClassificationJob(systemId) {
  const db = getDb();
  const job = db
    .prepare(
      `SELECT * FROM classification_jobs
        WHERE systemId = ? AND status IN ('Queued', 'Running')
        ORDER BY createdAt DESC
        LIMIT 1`
    )
    .get(systemId);

  if (!job) return null;
  if (activeClassificationJobs.has(job.id)) return job;

  db.prepare(
    `UPDATE classification_jobs
        SET status = 'Failed',
            errorMessage = ?,
            failedAt = ?,
            updatedAt = ?
      WHERE id = ?`
  ).run("Classification job stopped before completion.", nowIso(), nowIso(), job.id);
  checkpointDatabase("FULL");
  return null;
}

function createClassificationJob({ systemId, mode, page, pageSize, search, sortBy, sortDir, createdBy }) {
  const db = getDb();
  const safeMode = mode === "all" ? "all" : "page";
  const safePageSize = Math.min(100, Math.max(1, Number(pageSize || 25)));
  const options = {
    page: Math.max(1, Number(page || 1)),
    pageSize: safePageSize,
    search: safeMode === "all" ? "" : normalize(search),
    sortBy: safeMode === "all" ? "rowIndex" : normalize(sortBy) || "rowIndex",
    sortDir: safeMode === "all" ? "asc" : normalize(sortDir).toLowerCase() === "desc" ? "desc" : "asc",
  };
  const id = crypto.randomUUID();
  const now = nowIso();
  const draftJob = {
    mode: safeMode,
    pageSize: safePageSize,
    optionsJson: JSON.stringify(options),
  };
  const total = classificationJobTargets(systemId, draftJob).length;
  db.prepare(
    `INSERT INTO classification_jobs (
       id, systemId, mode, status, total, processed, pageSize, currentPage,
       optionsJson, createdBy, createdAt, updatedAt
     ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(
    id,
    systemId,
    safeMode,
    "Queued",
    total,
    0,
    safePageSize,
    options.page,
    JSON.stringify(options),
    createdBy,
    now,
    now
  );
  checkpointDatabase("FULL");
  return getClassificationJob(db, id);
}

function startClassificationJob(jobId) {
  if (activeClassificationJobs.has(jobId)) return activeClassificationJobs.get(jobId);
  const promise = runClassificationJob(jobId)
    .catch((error) => failClassificationJob(jobId, error))
    .finally(() => {
      activeClassificationJobs.delete(jobId);
    });
  activeClassificationJobs.set(jobId, promise);
  return promise;
}

async function runClassificationJob(jobId) {
  const db = getDb();
  let job = getClassificationJob(db, jobId);
  if (!job) return;

  const system = db.prepare("SELECT * FROM systems WHERE id = ?").get(job.systemId);
  if (!system) throw new Error("System not found for classification job.");

  const contextPoints = db.prepare("SELECT tag, content FROM context_points WHERE systemId = ?").all(job.systemId);
  const records = classificationJobTargets(job.systemId, job);
  const startedAt = nowIso();
  db.prepare(
    `UPDATE classification_jobs
        SET status = 'Running',
            total = ?,
            startedAt = COALESCE(startedAt, ?),
            updatedAt = ?
      WHERE id = ?`
  ).run(records.length, startedAt, startedAt, jobId);
  checkpointDatabase("FULL");

  job = getClassificationJob(db, jobId);
  emitClassificationJobEvent(jobId, "start", getClassificationJobPayload(jobId));

  let processed = Number(job.processed || 0);
  let fallbackWarningSent = false;
  let apiCreditWarningSent = Boolean(safeJson(job.warningJson, null)?.code === "OPENAI_NO_CREDITS");

  for (let index = 0; index < records.length;) {
    if (!classificationJobCanContinue(db, jobId, job.systemId)) return;
    const firstRecord = records[index];
    const currentPage = classificationJobPage(job, firstRecord);
    const pageRecords = [];
    while (index < records.length && classificationJobPage(job, records[index]) === currentPage) {
      pageRecords.push(records[index]);
      index += 1;
    }

    db.prepare(
      `UPDATE classification_jobs
          SET currentRecordId = ?,
              currentRowIndex = ?,
              currentPage = ?,
              updatedAt = ?
        WHERE id = ?`
    ).run(firstRecord.id, firstRecord.rowIndex, currentPage, nowIso(), jobId);
    checkpointDatabase("FULL");
    emitClassificationJobEvent(jobId, "progress", getClassificationJobPayload(jobId));

    let results;
    try {
      results = await classifyRecords(pageRecords, contextPoints);
      if (results.apiWarning?.code === "OPENAI_NO_CREDITS" && !apiCreditWarningSent) {
        apiCreditWarningSent = true;
        db.prepare("UPDATE classification_jobs SET warningJson = ?, updatedAt = ? WHERE id = ?").run(
          JSON.stringify(results.apiWarning),
          nowIso(),
          jobId
        );
        checkpointDatabase("FULL");
        emitClassificationJobEvent(jobId, "warning", results.apiWarning);
      }
    } catch (error) {
      results = new Map();
      for (const record of pageRecords) {
        results.set(record.id, classifyRecordLocally(record, contextPoints));
      }
      if (!fallbackWarningSent) {
        fallbackWarningSent = true;
        const warning = {
          severity: "warning",
          message: "AI classification failed for one row, so local rules were used and classification continued.",
          detail: error.message || "Classification fallback used",
        };
        db.prepare("UPDATE classification_jobs SET warningJson = ?, updatedAt = ? WHERE id = ?").run(
          JSON.stringify(warning),
          nowIso(),
          jobId
        );
        checkpointDatabase("FULL");
        emitClassificationJobEvent(jobId, "warning", warning);
      }
    }

    if (!classificationJobCanContinue(db, jobId, job.systemId)) return;
    for (const record of pageRecords) {
      if (!classificationJobCanContinue(db, jobId, job.systemId)) return;
      const result = results.get(record.id) || classifyRecordLocally(record, contextPoints);
      const updated = saveClassificationResult(db, record.id, result);
      if (!updated) return;
      processed += 1;
      db.prepare(
        `UPDATE classification_jobs
            SET processed = ?,
                currentRecordId = ?,
                currentRowIndex = ?,
                currentPage = ?,
                updatedAt = ?
          WHERE id = ?`
      ).run(processed, record.id, record.rowIndex, currentPage, nowIso(), jobId);
      checkpointDatabase("FULL");

      emitClassificationJobEvent(
        jobId,
        "row",
        getClassificationJobPayload(jobId, recordToApi(updated), null)
      );
    }

    const summary = updateSystemClassificationState(job.systemId);
    checkpointDatabase("FULL");
    emitClassificationJobEvent(jobId, "progress", getClassificationJobPayload(jobId, null, summary));
  }

  const summary = updateSystemClassificationState(job.systemId);
  db.prepare(
    `UPDATE classification_jobs
        SET status = 'Completed',
            processed = total,
            completedAt = ?,
            updatedAt = ?
      WHERE id = ?`
  ).run(nowIso(), nowIso(), jobId);
  checkpointDatabase("FULL");
  emitClassificationJobEvent(jobId, "done", getClassificationJobPayload(jobId, null, summary));
}

function failClassificationJob(jobId, error) {
  const db = getDb();
  const failedAt = nowIso();
  db.prepare(
    `UPDATE classification_jobs
        SET status = 'Failed',
            errorMessage = ?,
            failedAt = ?,
            updatedAt = ?
      WHERE id = ?`
  ).run(error.message || "Classification job failed", failedAt, failedAt, jobId);
  checkpointDatabase("FULL");
  emitClassificationJobEvent(jobId, "classification-error", getClassificationJobPayload(jobId));
}

function pct(part, total) {
  return total ? Math.round((part / total) * 100) : 0;
}

function exportFileName(systemName) {
  const safeName = normalize(systemName)
    .replace(/[^a-z0-9]+/gi, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 80) || "system";
  return `${safeName}-classification-report.xlsx`;
}

function contentDispositionAttachment(fileName) {
  const asciiName = fileName.replace(/[^a-zA-Z0-9._-]/g, "-");
  return `attachment; filename="${asciiName}"; filename*=UTF-8''${encodeURIComponent(fileName)}`;
}

function getDashboardSummary(userId) {
  const db = getDb();
  const systems = db.prepare("SELECT * FROM systems WHERE userId = ? ORDER BY updatedAt DESC").all(userId);
  const total = systems.length;
  let completed = 0;
  let inProgress = 0;

  for (const system of systems) {
    const summary = getSystemSummary(system.id);
    if (summary.totalRecords > 0 && summary.pending === 0) completed += 1;
    else inProgress += 1;
  }

  return {
    totalSystems: total,
    systemsInProgress: inProgress,
    completedSystems: completed,
    completionPercentage: pct(completed, total),
  };
}

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    databasePath,
    backupDir,
    startedAt: process.env.STARTED_AT || null,
  });
});

app.post("/api/auth/signup", async (req, res) => {
  const db = getDb();
  const name = normalize(req.body.name);
  const email = normalizeEmail(req.body.email);
  const password = String(req.body.password || "");
  const accountType = normalize(req.body.accountType) || "Individual";
  const companyName = normalize(req.body.companyName);

  if (!name || !email || password.length < 6) {
    res.status(400).json({ error: "Name, valid email, and a password of at least 6 characters are required." });
    return;
  }

  const existing = db.prepare("SELECT id FROM users WHERE email = ?").get(email);
  if (existing) {
    res.status(409).json({ error: "A user with this email already exists." });
    return;
  }

  const passwordHash = await bcrypt.hash(password, 12);
  const info = db
    .prepare(
      `INSERT INTO users (name, email, passwordHash, accountType, companyName, role, avatarInitials, createdAt, updatedAt)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    )
    .run(name, email, passwordHash, accountType, companyName, "Data Governance Lead", initialsFor(name, email), nowIso(), nowIso());

  createSession(res, info.lastInsertRowid);
  const user = db.prepare("SELECT * FROM users WHERE id = ?").get(info.lastInsertRowid);
  res.status(201).json({ user: publicUser(user) });
});

app.post("/api/auth/login", async (req, res) => {
  const db = getDb();
  const email = normalizeEmail(req.body.email);
  const password = String(req.body.password || "");
  const user = db.prepare("SELECT * FROM users WHERE email = ?").get(email);
  if (!user || !(await bcrypt.compare(password, user.passwordHash))) {
    res.status(401).json({ error: "Invalid email or password." });
    return;
  }

  createSession(res, user.id);
  res.json({ user: publicUser(user) });
});

app.post("/api/auth/forgot-password", (req, res) => {
  const db = getDb();
  const email = normalizeEmail(req.body.email);
  const user = db.prepare("SELECT * FROM users WHERE email = ?").get(email);

  if (!user) {
    res.json({ message: "If the account exists, a reset workflow has been prepared." });
    return;
  }

  const token = crypto.randomBytes(24).toString("hex");
  db.prepare(
    "INSERT INTO password_resets (userId, tokenHash, createdAt, expiresAt) VALUES (?, ?, ?, ?)"
  ).run(user.id, hashToken(token), nowIso(), new Date(Date.now() + 60 * 60 * 1000).toISOString());
  checkpointDatabase("FULL");

  res.json({
    message: "Reset token generated. In production this would be delivered by email.",
    resetToken: token,
  });
});

app.post("/api/auth/reset-password", async (req, res) => {
  const db = getDb();
  const token = normalize(req.body.token);
  const password = String(req.body.password || "");
  if (!token || password.length < 6) {
    res.status(400).json({ error: "Reset token and a password of at least 6 characters are required." });
    return;
  }

  const reset = db
    .prepare("SELECT * FROM password_resets WHERE tokenHash = ? AND usedAt IS NULL AND expiresAt > ?")
    .get(hashToken(token), nowIso());
  if (!reset) {
    res.status(400).json({ error: "Reset token is invalid or expired." });
    return;
  }

  const passwordHash = await bcrypt.hash(password, 12);
  db.prepare("UPDATE users SET passwordHash = ?, updatedAt = ? WHERE id = ?").run(passwordHash, nowIso(), reset.userId);
  db.prepare("UPDATE password_resets SET usedAt = ? WHERE id = ?").run(nowIso(), reset.id);
  checkpointDatabase("FULL");
  res.json({ message: "Password reset complete." });
});

app.get("/api/auth/me", requireAuth, (req, res) => {
  res.json({ user: publicUser(req.user) });
});

app.post("/api/auth/logout", authOptional, (req, res) => {
  if (req.user) {
    getDb().prepare("DELETE FROM sessions WHERE userId = ?").run(req.user.id);
    checkpointDatabase("FULL");
  }
  clearSession(req, res);
  res.json({ message: "Signed out." });
});

app.get("/api/dashboard", requireAuth, (req, res) => {
  res.json(getDashboardSummary(req.user.id));
});

app.get("/api/systems", requireAuth, (req, res) => {
  const db = getDb();
  const systems = db.prepare("SELECT * FROM systems WHERE userId = ? ORDER BY updatedAt DESC").all(req.user.id);
  res.json({
    systems: systems.map((system) => ({
      ...system,
      summary: getSystemSummary(system.id),
    })),
  });
});

app.post("/api/systems", requireAuth, (req, res) => {
  const db = getDb();
  const name = normalize(req.body.name);
  const owner = normalize(req.body.owner);
  const dba = normalize(req.body.dba);
  const ownerEmail = normalizeEmail(req.body.ownerEmail);
  const systemGroup = normalize(req.body.systemGroup);

  if (!name || !owner || !dba || !ownerEmail || !systemGroup) {
    res.status(400).json({ error: "System Name, Responsible Owner, DBA, Owner Email, and System Group are required." });
    return;
  }

  const info = db
    .prepare(
      `INSERT INTO systems (
        userId, name, owner, dba, ownerEmail, systemGroup, assignedConsultant, additionalSystem,
        sourceSystemRef, targetSystemRef, relatedSystemLinks, status, createdAt, updatedAt, lastModifiedBy
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    )
    .run(
      req.user.id,
      name,
      owner,
      dba,
      ownerEmail,
      systemGroup,
      normalize(req.body.assignedConsultant),
      req.body.additionalSystem ? 1 : 0,
      normalize(req.body.sourceSystemRef),
      normalize(req.body.targetSystemRef),
      normalize(req.body.relatedSystemLinks),
      "In Progress",
      nowIso(),
      nowIso(),
      req.user.email
    );
  checkpointDatabase("FULL");

  const system = db.prepare("SELECT * FROM systems WHERE id = ?").get(info.lastInsertRowid);
  res.status(201).json({ system: { ...system, summary: getSystemSummary(system.id) } });
});

app.get("/api/systems/:id", requireAuth, requireSystemAccess, (req, res) => {
  res.json({
    system: { ...req.system, summary: getSystemSummary(req.system.id) },
    originalColumns: getOriginalColumns(req.system.id),
  });
});

app.put("/api/systems/:id", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  db.prepare(
    `UPDATE systems
        SET owner = ?, dba = ?, ownerEmail = ?, systemGroup = ?, assignedConsultant = ?,
            sourceSystemRef = ?, targetSystemRef = ?, relatedSystemLinks = ?, updatedAt = ?, lastModifiedBy = ?
      WHERE id = ?`
  ).run(
    normalize(req.body.owner || req.system.owner),
    normalize(req.body.dba || req.system.dba),
    normalizeEmail(req.body.ownerEmail || req.system.ownerEmail),
    normalize(req.body.systemGroup || req.system.systemGroup),
    normalize(req.body.assignedConsultant || req.system.assignedConsultant),
    normalize(req.body.sourceSystemRef || req.system.sourceSystemRef),
    normalize(req.body.targetSystemRef || req.system.targetSystemRef),
    normalize(req.body.relatedSystemLinks || req.system.relatedSystemLinks),
    nowIso(),
    req.user.email,
    req.system.id
  );
  checkpointDatabase("FULL");
  const system = db.prepare("SELECT * FROM systems WHERE id = ?").get(req.system.id);
  res.json({ system: { ...system, summary: getSystemSummary(system.id) } });
});

app.delete("/api/systems/:id", requireAuth, requireSystemAccess, async (req, res) => {
  await createDatabaseBackup("before-system-delete");
  const db = getDb();
  const deletedAt = nowIso();
  db.prepare(
    `UPDATE classification_jobs
        SET status = 'Failed',
            errorMessage = ?,
            failedAt = ?,
            updatedAt = ?
      WHERE systemId = ? AND status IN ('Queued', 'Running')`
  ).run("System was deleted before classification completed.", deletedAt, deletedAt, req.system.id);
  db.prepare("DELETE FROM systems WHERE id = ?").run(req.system.id);
  checkpointDatabase("FULL");
  res.json({
    message: "System and all related records were deleted.",
    dashboard: getDashboardSummary(req.user.id),
  });
});

app.post("/api/systems/:id/upload", requireAuth, requireSystemAccess, upload.single("metadataFile"), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: "Excel or CSV file is required." });
    return;
  }

  const fileName = req.file.originalname;
  const parsed = await parseMetadataFile(req.file.buffer, fileName);
  if (!parsed.rows.length) {
    res.status(400).json({ error: "The uploaded file does not contain data rows." });
    return;
  }

  await createDatabaseBackup("before-upload-replace");
  const db = getDb();
  const insert = db.prepare(
    `INSERT INTO data_records (
      systemId, rowIndex, originalJson, tableName, columnName, dataType, reviewStatus,
      systemReviewStatus, personalReviewStatus, needsReview, owner, steward, reviewer, pushToClient,
      uploadedFileName, createdAt, updatedAt, lastModifiedBy
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  );

  try {
    db.exec("BEGIN IMMEDIATE");
    db.prepare("DELETE FROM data_records WHERE systemId = ?").run(req.system.id);
    parsed.rows.forEach((row, index) => {
      insert.run(
        req.system.id,
        index + 1,
        JSON.stringify(row.original),
        row.tableName,
        row.columnName,
        row.dataType,
        "Unreviewed",
        "Unreviewed",
        "Needs Review",
        1,
        req.system.owner,
        "",
        "",
        1,
        fileName,
        nowIso(),
        nowIso(),
        req.user.email
      );
    });
    db.prepare(
      "UPDATE systems SET lastUploadFileName = ?, status = ?, updatedAt = ?, lastModifiedBy = ? WHERE id = ?"
    ).run(fileName, "In Progress", nowIso(), req.user.email, req.system.id);
    db.exec("COMMIT");
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
  checkpointDatabase("FULL");
  res.json({
    fileName,
    rowsImported: parsed.rows.length,
    originalColumns: parsed.columns,
    summary: getSystemSummary(req.system.id),
  });
});

app.get("/api/systems/:id/records", requireAuth, requireSystemAccess, (req, res) => {
  const result = listRecords(req.system.id, {
    page: req.query.page,
    pageSize: req.query.pageSize,
    search: req.query.search,
    sortBy: req.query.sortBy,
    sortDir: req.query.sortDir,
    personalOnly: req.query.personalOnly,
  });

  res.json({
    records: result.rows,
    total: result.total,
    page: result.page,
    pageSize: result.pageSize,
    totalPages: result.totalPages,
    originalColumns: getOriginalColumns(req.system.id),
    summary: getSystemSummary(req.system.id),
  });
});

app.put("/api/records/:recordId", requireAuth, (req, res) => {
  const db = getDb();
  const record = db.prepare("SELECT * FROM data_records WHERE id = ?").get(req.params.recordId);
  if (!record) {
    res.status(404).json({ error: "Record not found" });
    return;
  }
  const system = db.prepare("SELECT * FROM systems WHERE id = ?").get(record.systemId);
  if (!canAccessSystem(req.user, system)) {
    res.status(403).json({ error: "You do not have access to this record." });
    return;
  }

  const allowed = {
    confidentiality: "confidentiality",
    confReason: "confReason",
    personalData: "personalData",
    personalReason: "personalReason",
    personalDataType: "personalDataType",
    pseudonymizable: "pseudonymizable",
    anonymizable: "anonymizable",
    specialCategory: "specialCategory",
    confidenceScore: "confidenceScore",
    reviewStatus: "reviewStatus",
    needsReview: "needsReview",
    policyRecommendation: "policyRecommendation",
    owner: "owner",
    steward: "steward",
    reviewer: "reviewer",
    pushToClient: "pushToClient",
  };

  const updates = [];
  const values = [];
  // Track whether lastReviewedAt has already been pushed to avoid the SQLite
  // "column specified more than once" error if both `reviewed` and
  // `reviewStatus=Approved` arrive in the same request body.
  let lastReviewedAtSet = false;

  if (Object.prototype.hasOwnProperty.call(req.body, "reviewed")) {
    const reviewed = Boolean(req.body.reviewed);
    updates.push("systemReviewStatus = ?");
    values.push(reviewed ? "Approved" : "Unreviewed");
    if (reviewed) {
      updates.push("lastReviewedAt = ?");
      values.push(nowIso());
      lastReviewedAtSet = true;
    }
  }

  if (Object.prototype.hasOwnProperty.call(req.body, "personalApproved")) {
    const approved = Boolean(req.body.personalApproved);
    updates.push("personalReviewStatus = ?");
    values.push(approved ? "Approved" : "Needs Review");
    updates.push("personalApprovedAt = ?");
    values.push(approved ? nowIso() : "");
    updates.push("personalApprovedBy = ?");
    values.push(approved ? req.user.email : "");
  }

  for (const [inputKey, column] of Object.entries(allowed)) {
    if (Object.prototype.hasOwnProperty.call(req.body, inputKey)) {
      updates.push(`${column} = ?`);
      values.push(inputKey === "needsReview" || inputKey === "pushToClient" ? (req.body[inputKey] ? 1 : 0) : req.body[inputKey]);
    }
  }

  if (req.body.reviewStatus === "Approved" && !lastReviewedAtSet) {
    updates.push("lastReviewedAt = ?");
    values.push(nowIso());
  }

  updates.push("updatedAt = ?");
  values.push(nowIso());
  updates.push("lastModifiedBy = ?");
  values.push(req.user.email);
  values.push(record.id);

  db.prepare(`UPDATE data_records SET ${updates.join(", ")} WHERE id = ?`).run(...values);
  checkpointDatabase("FULL");
  const updated = db.prepare("SELECT * FROM data_records WHERE id = ?").get(record.id);
  res.json({ record: recordToApi(updated), summary: getSystemSummary(updated.systemId) });
});

app.post("/api/systems/:id/personal-approvals", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  const now = nowIso();
  const recordIds = Array.isArray(req.body.recordIds)
    ? req.body.recordIds.map((id) => Number(id)).filter((id) => Number.isInteger(id) && id > 0)
    : [];

  if (!recordIds.length) {
    res.status(400).json({ error: "Select at least one personal-data record to approve." });
    return;
  }

  const placeholders = recordIds.map(() => "?").join(", ");
  db.prepare(
    `UPDATE data_records
        SET personalReviewStatus = 'Approved',
            personalApprovedAt = ?,
            personalApprovedBy = ?,
            updatedAt = ?,
            lastModifiedBy = ?
      WHERE systemId = ?
        AND personalData = 'Yes'
        AND id IN (${placeholders})`
  ).run(now, req.user.email, now, req.user.email, req.system.id, ...recordIds);
  checkpointDatabase("FULL");

  // Fetch ALL personal-data records directly — avoids the pageSize cap that
  // listRecords applies for paginated API responses.
  const records = db
    .prepare(
      "SELECT * FROM data_records WHERE systemId = ? AND personalData = 'Yes' ORDER BY rowIndex ASC"
    )
    .all(req.system.id)
    .map(recordToApi);
  res.json({
    records,
    summary: getSystemSummary(req.system.id),
  });
});

app.get("/api/systems/:id/review-queue", requireAuth, requireSystemAccess, (req, res) => {
  // Use a direct SQL query so we never pull every record into Node memory.
  const db = getDb();
  const rows = db
    .prepare(
      `SELECT * FROM data_records
       WHERE systemId = ?
         AND (
           (COALESCE(systemReviewStatus,'Unreviewed') != 'Approved'
             AND (needsReview = 1
                  OR (confidenceScore IS NOT NULL AND confidenceScore > 0 AND confidenceScore < 0.75)))
           OR (personalData = 'Yes'
               AND COALESCE(personalReviewStatus,'Needs Review') != 'Approved')
         )
       ORDER BY rowIndex ASC`
    )
    .all(req.system.id)
    .map(recordToApi);
  res.json({ records: rows, count: rows.length });
});

app.get("/api/systems/:id/context", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  const points = db.prepare("SELECT * FROM context_points WHERE systemId = ? ORDER BY createdAt DESC").all(req.system.id);
  res.json({ points });
});

app.post("/api/systems/:id/context", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  const tag = normalize(req.body.tag);
  const content = normalize(req.body.content);
  const allowed = ["System Service Brief", "System Description", "System Personas"];
  if (!allowed.includes(tag) || !content) {
    res.status(400).json({ error: "A valid context tag and content are required." });
    return;
  }

  const words = content.split(/\s+/).filter(Boolean);
  if (words.length > 200) {
    res.status(400).json({ error: "Context content must be 200 words or fewer." });
    return;
  }

  const info = db
    .prepare("INSERT INTO context_points (systemId, tag, content, createdAt, updatedAt, createdBy) VALUES (?, ?, ?, ?, ?, ?)")
    .run(req.system.id, tag, content, nowIso(), nowIso(), req.user.email);
  checkpointDatabase("FULL");
  const point = db.prepare("SELECT * FROM context_points WHERE id = ?").get(info.lastInsertRowid);
  res.status(201).json({ point });
});

app.delete("/api/context/:pointId", requireAuth, (req, res) => {
  const db = getDb();
  const point = db
    .prepare(
      `SELECT context_points.*, systems.userId
         FROM context_points
         JOIN systems ON systems.id = context_points.systemId
        WHERE context_points.id = ?`
    )
    .get(req.params.pointId);
  if (!point) {
    res.status(404).json({ error: "Context point not found." });
    return;
  }
  if (!canAccessSystem(req.user, point)) {
    res.status(403).json({ error: "You do not have access to this context point." });
    return;
  }
  db.prepare("DELETE FROM context_points WHERE id = ?").run(req.params.pointId);
  checkpointDatabase("FULL");
  res.json({ message: "Context point removed." });
});

app.get("/api/systems/:id/pdpl", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  let notes = db.prepare("SELECT * FROM pdpl_notes WHERE systemId = ?").get(req.system.id);
  if (!notes) {
    db.prepare(
      `INSERT INTO pdpl_notes (
        systemId, governanceNotes, dataSubjectRightsCoverage, consentTrackingStatus,
        crossBorderTransferFlags, createdAt, updatedAt, updatedBy
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    ).run(req.system.id, "", "Needs Review", "Needs Review", "No Flags Recorded", nowIso(), nowIso(), req.user.email);
    checkpointDatabase("FULL");
    notes = db.prepare("SELECT * FROM pdpl_notes WHERE systemId = ?").get(req.system.id);
  }

  // Fetch ALL personal-data records directly — avoids the pageSize cap.
  const personalRows = db
    .prepare(
      "SELECT * FROM data_records WHERE systemId = ? AND personalData = 'Yes' ORDER BY rowIndex ASC"
    )
    .all(req.system.id)
    .map(recordToApi);
  const compliant = personalRows.filter((row) => row.personalReviewStatus === "Approved").length;
  const nonCompliant = personalRows.filter((row) => row.specialCategory === "Yes" && row.personalReviewStatus !== "Approved").length;
  const needsReview = Math.max(0, personalRows.length - compliant - nonCompliant);
  res.json({
    notes,
    status: { compliant, needsReview, nonCompliant },
    obligations: personalRows.map((row) => ({
      id: row.id,
      tableName: row.tableName,
      columnName: row.columnName,
      personalDataType: row.personalDataType,
      policyRecommendation: row.policyRecommendation || "review needed",
      reviewStatus: row.personalReviewStatus,
      owner: row.owner,
      steward: row.steward,
    })),
  });
});

app.put("/api/systems/:id/pdpl", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  db.prepare(
    `INSERT INTO pdpl_notes (
       systemId, governanceNotes, dataSubjectRightsCoverage, consentTrackingStatus,
       crossBorderTransferFlags, createdAt, updatedAt, updatedBy
     ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
     ON CONFLICT(systemId) DO UPDATE SET
       governanceNotes = excluded.governanceNotes,
       dataSubjectRightsCoverage = excluded.dataSubjectRightsCoverage,
       consentTrackingStatus = excluded.consentTrackingStatus,
       crossBorderTransferFlags = excluded.crossBorderTransferFlags,
       updatedAt = excluded.updatedAt,
       updatedBy = excluded.updatedBy`
  ).run(
    req.system.id,
    normalize(req.body.governanceNotes),
    normalize(req.body.dataSubjectRightsCoverage),
    normalize(req.body.consentTrackingStatus),
    normalize(req.body.crossBorderTransferFlags),
    nowIso(),
    nowIso(),
    req.user.email
  );
  checkpointDatabase("FULL");
  const notes = db.prepare("SELECT * FROM pdpl_notes WHERE systemId = ?").get(req.system.id);
  res.json({ notes });
});

app.get("/api/systems/:id/links", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  const links = db.prepare("SELECT * FROM system_links WHERE systemId = ? ORDER BY category, title").all(req.system.id);
  res.json({ links });
});

app.post("/api/systems/:id/links", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  const title = normalize(req.body.title);
  const url = normalize(req.body.url);
  const category = normalize(req.body.category) || "Other";
  const description = normalize(req.body.description);
  const allowed = ["Documentation", "API", "Governance Policy", "Other"];
  if (!title || !url || !allowed.includes(category)) {
    res.status(400).json({ error: "Title, URL, and a valid category are required." });
    return;
  }
  const info = db
    .prepare("INSERT INTO system_links (systemId, title, url, category, description, createdAt, updatedAt, createdBy) VALUES (?, ?, ?, ?, ?, ?, ?, ?)")
    .run(req.system.id, title, url, category, description, nowIso(), nowIso(), req.user.email);
  checkpointDatabase("FULL");
  const link = db.prepare("SELECT * FROM system_links WHERE id = ?").get(info.lastInsertRowid);
  res.status(201).json({ link });
});

app.put("/api/links/:linkId", requireAuth, (req, res) => {
  const db = getDb();
  const link = db.prepare("SELECT * FROM system_links WHERE id = ?").get(req.params.linkId);
  if (!link) {
    res.status(404).json({ error: "Link not found" });
    return;
  }
  const system = db.prepare("SELECT * FROM systems WHERE id = ?").get(link.systemId);
  if (!canAccessSystem(req.user, system)) {
    res.status(403).json({ error: "You do not have access to this link." });
    return;
  }
  db.prepare(
    "UPDATE system_links SET title = ?, url = ?, category = ?, description = ?, updatedAt = ? WHERE id = ?"
  ).run(
    normalize(req.body.title || link.title),
    normalize(req.body.url || link.url),
    normalize(req.body.category || link.category),
    normalize(req.body.description || link.description),
    nowIso(),
    link.id
  );
  checkpointDatabase("FULL");
  const updated = db.prepare("SELECT * FROM system_links WHERE id = ?").get(link.id);
  res.json({ link: updated });
});

app.delete("/api/links/:linkId", requireAuth, (req, res) => {
  const db = getDb();
  const link = db.prepare("SELECT * FROM system_links WHERE id = ?").get(req.params.linkId);
  if (!link) {
    res.status(404).json({ error: "Link not found" });
    return;
  }
  const system = db.prepare("SELECT * FROM systems WHERE id = ?").get(link.systemId);
  if (!canAccessSystem(req.user, system)) {
    res.status(403).json({ error: "You do not have access to this link." });
    return;
  }
  db.prepare("DELETE FROM system_links WHERE id = ?").run(req.params.linkId);
  checkpointDatabase("FULL");
  res.json({ message: "Link removed." });
});

app.post("/api/systems/:id/classification-jobs", requireAuth, requireSystemAccess, (req, res) => {
  const activeJob = findActiveClassificationJob(req.system.id);
  if (activeJob) {
    res.status(202).json({ job: classificationJobToApi(activeJob), reused: true });
    return;
  }

  const job = createClassificationJob({
    systemId: req.system.id,
    mode: req.body.mode,
    page: req.body.page,
    pageSize: req.body.pageSize,
    search: req.body.search,
    sortBy: req.body.sortBy,
    sortDir: req.body.sortDir,
    createdBy: req.user.email,
  });
  setImmediate(() => startClassificationJob(job.id));
  res.status(202).json({ job: classificationJobToApi(job), reused: false });
});

app.get("/api/systems/:id/classification-jobs/:jobId", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  const job = getClassificationJob(db, req.params.jobId);
  if (!job || job.systemId !== req.system.id) {
    res.status(404).json({ error: "Classification job not found." });
    return;
  }
  res.json(getClassificationJobPayload(job.id));
});

app.get("/api/systems/:id/classification-jobs/:jobId/stream", requireAuth, requireSystemAccess, (req, res) => {
  const db = getDb();
  const job = getClassificationJob(db, req.params.jobId);
  if (!job || job.systemId !== req.system.id) {
    res.status(404).json({ error: "Classification job not found." });
    return;
  }

  res.writeHead(200, {
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache, no-transform",
    Connection: "keep-alive",
    "X-Accel-Buffering": "no",
  });

  let clientConnected = true;
  const send = (event, payload) => {
    if (!clientConnected || res.writableEnded || res.destroyed) return false;
    try {
      res.write(`event: ${event}\n`);
      res.write(`data: ${JSON.stringify(payload)}\n\n`);
      return true;
    } catch (_error) {
      clientConnected = false;
      return false;
    }
  };

  const eventKey = `classification-job:${job.id}`;
  const listener = (event, payload) => {
    send(event, payload);
    if (event === "done" || event === "classification-error") {
      setImmediate(cleanup);
    }
  };
  const cleanup = () => {
    clientConnected = false;
    clearInterval(heartbeat);
    classificationJobEvents.off(eventKey, listener);
    if (!res.writableEnded && !res.destroyed) res.end();
  };

  req.on("close", cleanup);
  classificationJobEvents.on(eventKey, listener);
  const heartbeat = setInterval(() => {
    const payload = getClassificationJobPayload(job.id);
    if (!payload) {
      cleanup();
      return;
    }
    send("heartbeat", { at: nowIso(), job: payload.job });
  }, 15000);

  const payload = getClassificationJobPayload(job.id);
  send("start", payload);
  if (payload?.job?.warning) send("warning", payload.job.warning);
  if (payload?.job?.status === "Completed") {
    send("done", payload);
    setImmediate(cleanup);
  } else if (payload?.job?.status === "Failed") {
    send("classification-error", payload);
    setImmediate(cleanup);
  }
});

// NOTE: The legacy /api/systems/:id/classify-stream endpoint has been removed.
// All classification is now handled through the job-based endpoint:
//   POST   /api/systems/:id/classification-jobs
//   GET    /api/systems/:id/classification-jobs/:jobId/stream

app.get("/api/systems/:id/export", requireAuth, requireSystemAccess, async (req, res, next) => {
  try {
    const db = getDb();
    const rows = db.prepare("SELECT * FROM data_records WHERE systemId = ? ORDER BY rowIndex ASC").all(req.system.id).map(recordToApi);
    const pdpl = db.prepare("SELECT * FROM pdpl_notes WHERE systemId = ?").get(req.system.id);
    const workbook = await buildSystemDataExport({
      system: req.system,
      rows,
      originalColumns: getOriginalColumns(req.system.id),
      summary: getSystemSummary(req.system.id),
      pdpl,
    });

    const fileName = exportFileName(req.system.name);
    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
    res.status(200);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", contentDispositionAttachment(fileName));
    res.setHeader("Content-Length", String(buffer.length));
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.send(buffer);
  } catch (error) {
    next(error);
  }
});

app.get("/", authOptional, (req, res) => {
  if (req.user) {
    res.redirect(302, "/app");
    return;
  }
  res.setHeader("Cache-Control", "no-store");
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.get(["/app", "/app/*"], requirePageAuth, (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.get("*", (req, res) => {
  if (req.accepts("html")) {
    res.redirect(302, "/");
    return;
  }
  res.status(404).json({ error: "Not found" });
});

app.use((error, _req, res, _next) => {
  console.error(error);
  res.status(500).json({ error: error.message || "Unexpected server error" });
});

// ── Graceful shutdown ────────────────────────────────────────────────────────
// On SIGTERM / SIGINT we run a FULL WAL checkpoint before exiting.  This
// ensures every committed write is merged from the WAL into the main database
// file so that the on-disk app.db is always up-to-date even when the process
// is stopped cleanly (e.g. by a process manager, Ctrl-C, or a deployment
// script that restarts the server).
function gracefulShutdown(signal) {
  console.log(`\n${signal} — flushing WAL to database before exit...`);
  try {
    closeDatabase();
    console.log("WAL checkpoint complete.");
  } catch (err) {
    console.error("WAL checkpoint failed during shutdown:", err.message);
  }
  process.exit(0);
}
process.on("SIGTERM", () => gracefulShutdown("SIGTERM"));
process.on("SIGINT",  () => gracefulShutdown("SIGINT"));
process.on("beforeExit", () => {
  try {
    closeDatabase();
  } catch (err) {
    console.error("WAL checkpoint failed before exit:", err.message);
  }
});

async function startServer() {
  await initDatabase();
  app.listen(PORT, () => {
    process.env.STARTED_AT = nowIso();
    console.log(`DATA CLASSIFICATION & GOVERNANCE PLATFORM running on http://localhost:${PORT}`);
    console.log(`SQLite database: ${databasePath}`);
    console.log(`SQLite backups: ${backupDir}`);
  });
}
startServer().catch((error) => {
  console.error("Failed to start platform:", error);
  process.exit(1);
});
