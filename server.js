const path = require("path");
const express = require("express");
const cookieParser = require("cookie-parser");
const multer = require("multer");
const bcrypt = require("bcryptjs");
const crypto = require("crypto");
const dotenv = require("dotenv");

dotenv.config({ path: path.join(__dirname, ".env.local") });

const {
  initDatabase,
  getDb,
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

function listRecords(systemId, options = {}) {
  const db = getDb();
  const search = normalize(options.search).toLowerCase();
  const sortBy = normalize(options.sortBy) || "rowIndex";
  const sortDir = normalize(options.sortDir).toLowerCase() === "desc" ? -1 : 1;
  const personalOnly = options.personalOnly === true || options.personalOnly === "true";

  let rows = db
    .prepare("SELECT * FROM data_records WHERE systemId = ? ORDER BY rowIndex ASC")
    .all(systemId)
    .map(recordToApi);

  if (personalOnly) {
    rows = rows.filter((row) => row.personalData === "Yes");
  }

  if (search) {
    rows = rows.filter((row) => {
      const haystack = [
        row.tableName,
        row.columnName,
        row.dataType,
        row.confidentiality,
        row.personalDataType,
        row.policyRecommendation,
        row.owner,
        row.steward,
        row.reviewer,
        JSON.stringify(row.original),
      ]
        .join(" ")
        .toLowerCase();
      return haystack.includes(search);
    });
  }

  rows.sort((a, b) => {
    const aValue = valueForSort(a, sortBy);
    const bValue = valueForSort(b, sortBy);
    if (typeof aValue === "number" && typeof bValue === "number") return (aValue - bValue) * sortDir;
    return String(aValue || "").localeCompare(String(bValue || "")) * sortDir;
  });

  const total = rows.length;
  const page = Math.max(1, Number(options.page || 1));
  const pageSize = Math.min(100, Math.max(1, Number(options.pageSize || 25)));
  const offset = (page - 1) * pageSize;

  return {
    rows: rows.slice(offset, offset + pageSize),
    allRows: rows,
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

function getSystemSummary(systemId) {
  const db = getDb();
  const rows = db.prepare("SELECT * FROM data_records WHERE systemId = ?").all(systemId).map(recordToApi);
  const total = rows.length;
  const classified = rows.filter((row) => row.confidentiality && row.confidentiality !== "Pending").length;
  const personal = rows.filter((row) => row.personalData === "Yes").length;
  const pending = Math.max(0, total - classified);
  const reviewQueue = rows.filter(
    (row) =>
      (row.systemReviewStatus !== "Approved" &&
        (row.needsReview || Number(row.confidenceScore || 0) < 0.75)) ||
      (row.personalData === "Yes" && row.personalReviewStatus !== "Approved")
  ).length;
  const lowConfidence = rows.filter((row) => Number(row.confidenceScore || 0) > 0 && Number(row.confidenceScore || 0) < 0.75).length;
  const policyRecommendationCount = rows.filter((row) => normalize(row.policyRecommendation)).length;
  const tablesWithPersonalData = new Set(rows.filter((row) => row.personalData === "Yes").map((row) => row.tableName)).size;
  const confidentiality = countBy(rows, "confidentiality", ["Confidential", "Secret", "Top Secret", "Public", "Pending"]);
  const personalDistribution = {
    "No Personal Data": rows.filter((row) => row.personalData !== "Yes").length,
    "Has Personal Data": personal,
  };
  const completedSystems = pending === 0 && total > 0 ? 1 : 0;

  return {
    totalRecords: total,
    classified,
    classifiedPercentage: pct(classified, total),
    personal,
    personalPercentage: pct(personal, total),
    pending,
    pendingPercentage: pct(pending, total),
    reviewQueue,
    lowConfidence,
    policyRecommendationCount,
    tablesWithPersonalData,
    confidentiality,
    personalDistribution,
    classificationProgress: pct(classified, total),
    completedSystems,
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

function countBy(rows, key, labels) {
  const result = {};
  for (const label of labels) result[label] = 0;
  for (const row of rows) {
    const value = row[key] || "Pending";
    result[value] = (result[value] || 0) + 1;
  }
  return result;
}

function pct(part, total) {
  return total ? Math.round((part / total) * 100) : 0;
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
  res.json({ message: "Password reset complete." });
});

app.get("/api/auth/me", requireAuth, (req, res) => {
  res.json({ user: publicUser(req.user) });
});

app.post("/api/auth/logout", authOptional, (req, res) => {
  if (req.user) {
    getDb().prepare("DELETE FROM sessions WHERE userId = ?").run(req.user.id);
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
  const system = db.prepare("SELECT * FROM systems WHERE id = ?").get(req.system.id);
  res.json({ system: { ...system, summary: getSystemSummary(system.id) } });
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
  if (Object.prototype.hasOwnProperty.call(req.body, "reviewed")) {
    const reviewed = Boolean(req.body.reviewed);
    updates.push("systemReviewStatus = ?");
    values.push(reviewed ? "Approved" : "Unreviewed");
    if (reviewed) {
      updates.push("lastReviewedAt = ?");
      values.push(nowIso());
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

  if (req.body.reviewStatus === "Approved") {
    updates.push("lastReviewedAt = ?");
    values.push(nowIso());
  }

  updates.push("updatedAt = ?");
  values.push(nowIso());
  updates.push("lastModifiedBy = ?");
  values.push(req.user.email);
  values.push(record.id);

  db.prepare(`UPDATE data_records SET ${updates.join(", ")} WHERE id = ?`).run(...values);
  db.exec("PRAGMA wal_checkpoint(PASSIVE)");
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
  db.exec("PRAGMA wal_checkpoint(PASSIVE)");

  const records = listRecords(req.system.id, { page: 1, pageSize: 100000, personalOnly: true }).allRows;
  res.json({
    records,
    summary: getSystemSummary(req.system.id),
  });
});

app.get("/api/systems/:id/review-queue", requireAuth, requireSystemAccess, (req, res) => {
  const rows = listRecords(req.system.id, { page: 1, pageSize: 100 }).allRows.filter(
    (row) =>
      (row.systemReviewStatus !== "Approved" &&
        (row.needsReview || Number(row.confidenceScore || 0) < 0.75)) ||
      (row.personalData === "Yes" && row.personalReviewStatus !== "Approved")
  );
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
    notes = db.prepare("SELECT * FROM pdpl_notes WHERE systemId = ?").get(req.system.id);
  }

  const personalRows = listRecords(req.system.id, { page: 1, pageSize: 100000, personalOnly: true }).allRows;
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
  res.json({ message: "Link removed." });
});

app.get("/api/systems/:id/classify-stream", requireAuth, requireSystemAccess, async (req, res) => {
  res.writeHead(200, {
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache, no-transform",
    Connection: "keep-alive",
    "X-Accel-Buffering": "no",
  });

  let clientConnected = true;
  req.on("close", () => {
    clientConnected = false;
  });

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
  const heartbeat = setInterval(() => {
    send("heartbeat", { at: nowIso() });
  }, 15000);

  try {
    const db = getDb();
    const contextPoints = db.prepare("SELECT tag, content FROM context_points WHERE systemId = ?").all(req.system.id);
    const mode = req.query.mode === "all" ? "all" : "page";
    const pageOptions = {
      page: req.query.page || 1,
      pageSize: mode === "all" ? 100000 : Math.min(50, Number(req.query.pageSize || 50)),
      search: mode === "all" ? "" : req.query.search,
      sortBy: req.query.sortBy,
      sortDir: req.query.sortDir,
    };
    const recordSet = listRecords(req.system.id, pageOptions);
    const records = mode === "all" ? recordSet.allRows : recordSet.rows.slice(0, 50);

    send("start", {
      mode,
      total: records.length,
      fileName: req.system.lastUploadFileName || records[0]?.uploadedFileName || "",
      contextPoints: contextPoints.length,
    });

    if (!records.length) {
      send("done", { total: 0, summary: getSystemSummary(req.system.id) });
      res.end();
      return;
    }

    const batchSize = 15;
    let processed = 0;
    for (let index = 0; index < records.length; index += batchSize) {
      const batch = records.slice(index, index + batchSize);
      let results;
      try {
        results = await classifyRecords(batch, contextPoints);
      } catch (error) {
        results = new Map();
        for (const record of batch) {
          results.set(record.id, classifyRecordLocally(record, contextPoints));
        }
        send("warning", {
          message: `Batch ${Math.floor(index / batchSize) + 1} used local rules after AI classification failed.`,
          detail: error.message || "Classification fallback used",
        });
      }

      for (const record of batch) {
        const result = results.get(record.id) || classifyRecordLocally(record, contextPoints);
        const needsReview = classificationNeedsReview(result);
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
          `Classified via ${result.source || "AI"} at ${nowIso()}`,
          nowIso(),
          "AI Classification Service",
          record.id
        );

        processed += 1;
        const updated = db.prepare("SELECT * FROM data_records WHERE id = ?").get(record.id);
        send("row", {
          processed,
          total: records.length,
          percentage: pct(processed, records.length),
          record: recordToApi(updated),
        });
      }

      const summary = updateSystemClassificationState(req.system.id);
      db.exec("PRAGMA wal_checkpoint(PASSIVE)");

      send("progress", {
        processed,
        total: records.length,
        percentage: pct(processed, records.length),
        summary,
      });
    }

    const summary = updateSystemClassificationState(req.system.id);
    db.exec("PRAGMA wal_checkpoint(PASSIVE)");
    send("done", { total: records.length, summary });
  } catch (error) {
    send("error", { error: error.message || "Classification failed" });
  } finally {
    clearInterval(heartbeat);
    if (!res.writableEnded && !res.destroyed) {
      res.end();
    }
  }
});

app.get("/api/systems/:id/export", requireAuth, requireSystemAccess, async (req, res) => {
  const db = getDb();
  const rows = db.prepare("SELECT * FROM data_records WHERE systemId = ? ORDER BY rowIndex").all(req.system.id).map(recordToApi);
  const pdpl = db.prepare("SELECT * FROM pdpl_notes WHERE systemId = ?").get(req.system.id);
  const links = db.prepare("SELECT * FROM system_links WHERE systemId = ? ORDER BY category, title").all(req.system.id);
  const workbook = await buildSystemDataExport({
    system: req.system,
    rows,
    originalColumns: getOriginalColumns(req.system.id),
    summary: getSystemSummary(req.system.id),
    pdpl,
    links,
  });

  const safeName = req.system.name.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "") || "system";
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${safeName}-classification-report.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
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
