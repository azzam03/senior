const fs = require("fs");
const path = require("path");
const bcrypt = require("bcryptjs");

// node:sqlite ships with Node 22.6+; the standalone `backup` helper was added
// in 22.13. We guard against it being absent so older 22.x patch releases work.
const nodeSqlite = require("node:sqlite");
const DatabaseSync = nodeSqlite.DatabaseSync;
const backup = typeof nodeSqlite.backup === "function" ? nodeSqlite.backup : null;

const projectRoot = path.resolve(__dirname, "..");

// Store the database inside the project directory.  On Render free tier the
// filesystem persists between spin-down / spin-up cycles (inactivity sleep),
// so data survives as long as no new deploy is triggered.
const databaseDir = path.join(projectRoot, "database");
const backupDir = path.join(databaseDir, "backups");
const databasePath = path.join(databaseDir, "app.db");

let connection = null;

function ensureDatabaseDirectories() {
  fs.mkdirSync(databaseDir, { recursive: true });
  fs.mkdirSync(backupDir, { recursive: true });
}

function backupFileName(reason = "startup") {
  const stamp = new Date()
    .toISOString()
    .replace("T", "-")
    .slice(0, 16)
    .replace(/:/g, "-");
  const safeReason = String(reason || "backup").replace(/[^a-z0-9-]+/gi, "-").replace(/^-|-$/g, "");
  let candidate = path.join(backupDir, `app-backup-${stamp}${safeReason ? `-${safeReason}` : ""}.db`);
  let suffix = 1;
  while (fs.existsSync(candidate)) {
    candidate = path.join(backupDir, `app-backup-${stamp}${safeReason ? `-${safeReason}` : ""}-${suffix}.db`);
    suffix += 1;
  }
  return candidate;
}

async function createDatabaseBackup(reason = "startup") {
  ensureDatabaseDirectories();
  if (!fs.existsSync(databasePath) || fs.statSync(databasePath).size === 0) {
    return null;
  }

  const destination = backupFileName(reason);
  try {
    if (connection && backup) {
      // Hot backup via node:sqlite API — consistent even while writes are in flight.
      await backup(connection, destination);
    } else {
      // Fallback: plain file copy (safe when no writes are happening).
      fs.copyFileSync(databasePath, destination);
    }
  } catch (_backupError) {
    // If the hot backup fails for any reason, fall back to file copy so startup
    // is never blocked.
    try {
      fs.copyFileSync(databasePath, destination);
    } catch (_copyError) {
      return null;
    }
  }
  return destination;
}

function getDb() {
  if (global.__DCP_SQLITE_CONNECTION) {
    connection = global.__DCP_SQLITE_CONNECTION;
    return connection;
  }

  ensureDatabaseDirectories();
  connection = new DatabaseSync(databasePath);

  // WAL mode gives better concurrency and crash safety.
  connection.exec("PRAGMA journal_mode = WAL");
  // FULL synchronous guarantees every committed write is on disk before we
  // return. This is critical: NORMAL can leave WAL frames un-flushed, so if
  // the process is killed between a write and the next auto-checkpoint those
  // frames may be lost. FULL prevents that.
  connection.exec("PRAGMA synchronous = FULL");
  // Enforce referential integrity so CASCADE deletes work correctly.
  connection.exec("PRAGMA foreign_keys = ON");
  // Wait up to 5 s before giving up on a locked database (prevents immediate
  // "SQLITE_BUSY" errors during concurrent requests).
  connection.exec("PRAGMA busy_timeout = 5000");
  // Keep temp tables and indices in RAM instead of a temp file.
  connection.exec("PRAGMA temp_store = MEMORY");
  // Use a ~16 MB page cache (negative value = KiB).
  connection.exec("PRAGMA cache_size = -16000");
  // Auto-checkpoint every 500 pages (~2 MB).  Default is 1000 which can
  // leave large amounts of data only in the WAL file.
  connection.exec("PRAGMA wal_autocheckpoint = 500");

  global.__DCP_SQLITE_CONNECTION = connection;
  return connection;
}

async function initDatabase() {
  ensureDatabaseDirectories();
  const hadExistingDatabase = fs.existsSync(databasePath) && fs.statSync(databasePath).size > 0;
  const db = getDb();
  if (hadExistingDatabase) {
    await createDatabaseBackup("startup");
  }
  createTables(db);
  runSafeMigrations(db);
  seedAdminIfNeeded(db);

  // Force a FULL checkpoint on every startup.  This merges any WAL frames
  // that survived a previous unclean shutdown into the main database file,
  // ensuring the file on disk is always up-to-date before we accept requests.
  try {
    db.exec("PRAGMA wal_checkpoint(FULL)");
  } catch (_err) {
    // Non-fatal — checkpoint failure at startup just means the WAL will be
    // merged during normal operation.
  }
}

function createTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      email TEXT NOT NULL UNIQUE,
      passwordHash TEXT NOT NULL,
      accountType TEXT NOT NULL DEFAULT 'Individual',
      companyName TEXT,
      role TEXT,
      avatarInitials TEXT,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS sessions (
      id TEXT PRIMARY KEY,
      userId INTEGER NOT NULL,
      createdAt TEXT NOT NULL,
      expiresAt TEXT NOT NULL,
      FOREIGN KEY (userId) REFERENCES users(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS password_resets (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      userId INTEGER NOT NULL,
      tokenHash TEXT NOT NULL,
      createdAt TEXT NOT NULL,
      expiresAt TEXT NOT NULL,
      usedAt TEXT,
      FOREIGN KEY (userId) REFERENCES users(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS systems (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      userId INTEGER NOT NULL,
      name TEXT NOT NULL,
      owner TEXT NOT NULL,
      dba TEXT NOT NULL,
      ownerEmail TEXT NOT NULL,
      systemGroup TEXT NOT NULL,
      assignedConsultant TEXT,
      additionalSystem INTEGER NOT NULL DEFAULT 0,
      sourceSystemRef TEXT,
      targetSystemRef TEXT,
      relatedSystemLinks TEXT,
      lastUploadFileName TEXT,
      status TEXT NOT NULL DEFAULT 'In Progress',
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL,
      lastModifiedBy TEXT,
      FOREIGN KEY (userId) REFERENCES users(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS data_records (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      systemId INTEGER NOT NULL,
      rowIndex INTEGER NOT NULL,
      originalJson TEXT NOT NULL,
      tableName TEXT NOT NULL,
      columnName TEXT NOT NULL,
      dataType TEXT,
      confidentiality TEXT DEFAULT 'Pending',
      confReason TEXT,
      personalData TEXT DEFAULT 'No',
      personalReason TEXT,
      personalDataType TEXT,
      pseudonymizable TEXT DEFAULT 'No',
      anonymizable TEXT DEFAULT 'No',
      specialCategory TEXT DEFAULT 'No',
      confidenceScore REAL,
      reviewStatus TEXT DEFAULT 'Unreviewed',
      systemReviewStatus TEXT DEFAULT 'Unreviewed',
      personalReviewStatus TEXT DEFAULT 'Needs Review',
      personalApprovedAt TEXT,
      personalApprovedBy TEXT,
      needsReview INTEGER DEFAULT 1,
      policyRecommendation TEXT,
      owner TEXT,
      steward TEXT,
      reviewer TEXT,
      pushToClient INTEGER DEFAULT 1,
      auditTrail TEXT,
      lastModifiedBy TEXT,
      lastReviewedAt TEXT,
      uploadedFileName TEXT,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL,
      FOREIGN KEY (systemId) REFERENCES systems(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS context_points (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      systemId INTEGER NOT NULL,
      tag TEXT NOT NULL,
      content TEXT NOT NULL,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL,
      createdBy TEXT,
      FOREIGN KEY (systemId) REFERENCES systems(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS pdpl_notes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      systemId INTEGER NOT NULL UNIQUE,
      governanceNotes TEXT,
      dataSubjectRightsCoverage TEXT,
      consentTrackingStatus TEXT,
      crossBorderTransferFlags TEXT,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL,
      updatedBy TEXT,
      FOREIGN KEY (systemId) REFERENCES systems(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS system_links (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      systemId INTEGER NOT NULL,
      title TEXT NOT NULL,
      url TEXT NOT NULL,
      category TEXT NOT NULL,
      description TEXT,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL,
      createdBy TEXT,
      FOREIGN KEY (systemId) REFERENCES systems(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS classification_cache (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      cacheKey TEXT NOT NULL UNIQUE,
      tableName TEXT NOT NULL,
      columnName TEXT NOT NULL,
      contextHash TEXT NOT NULL DEFAULT '',
      resultJson TEXT NOT NULL,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS classification_jobs (
      id TEXT PRIMARY KEY,
      systemId INTEGER NOT NULL,
      mode TEXT NOT NULL,
      status TEXT NOT NULL,
      total INTEGER NOT NULL DEFAULT 0,
      processed INTEGER NOT NULL DEFAULT 0,
      pageSize INTEGER NOT NULL DEFAULT 25,
      currentPage INTEGER NOT NULL DEFAULT 1,
      currentRecordId INTEGER,
      currentRowIndex INTEGER,
      optionsJson TEXT,
      warningJson TEXT,
      errorMessage TEXT,
      createdBy TEXT,
      startedAt TEXT,
      completedAt TEXT,
      failedAt TEXT,
      createdAt TEXT NOT NULL,
      updatedAt TEXT NOT NULL,
      FOREIGN KEY (systemId) REFERENCES systems(id) ON DELETE CASCADE
    );
  `);

  // ── Performance indexes ───────────────────────────────────────────────────
  // These speed up the most frequent queries: record listing, summary
  // aggregates, personal-data lookups, active-job checks, and session lookup.
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_data_records_systemId
      ON data_records(systemId);

    CREATE INDEX IF NOT EXISTS idx_data_records_systemId_personalData
      ON data_records(systemId, personalData);

    CREATE INDEX IF NOT EXISTS idx_data_records_systemId_rowIndex
      ON data_records(systemId, rowIndex);

    CREATE INDEX IF NOT EXISTS idx_classification_jobs_systemId_status
      ON classification_jobs(systemId, status);

    CREATE INDEX IF NOT EXISTS idx_sessions_userId
      ON sessions(userId);

    CREATE INDEX IF NOT EXISTS idx_sessions_expiresAt
      ON sessions(expiresAt);

    CREATE INDEX IF NOT EXISTS idx_context_points_systemId
      ON context_points(systemId);

    CREATE INDEX IF NOT EXISTS idx_system_links_systemId
      ON system_links(systemId);
  `);
}

function runSafeMigrations(db) {
  const desiredColumns = {
    systems: {
      sourceSystemRef: "TEXT",
      targetSystemRef: "TEXT",
      relatedSystemLinks: "TEXT",
      lastUploadFileName: "TEXT",
      lastModifiedBy: "TEXT",
    },
    data_records: {
      personalDataType: "TEXT",
      pseudonymizable: "TEXT DEFAULT 'No'",
      anonymizable: "TEXT DEFAULT 'No'",
      specialCategory: "TEXT DEFAULT 'No'",
      policyRecommendation: "TEXT",
      owner: "TEXT",
      steward: "TEXT",
      reviewer: "TEXT",
      pushToClient: "INTEGER DEFAULT 1",
      systemReviewStatus: "TEXT DEFAULT 'Unreviewed'",
      personalReviewStatus: "TEXT DEFAULT 'Needs Review'",
      personalApprovedAt: "TEXT",
      personalApprovedBy: "TEXT",
      auditTrail: "TEXT",
      lastModifiedBy: "TEXT",
      lastReviewedAt: "TEXT",
      uploadedFileName: "TEXT",
    },
    classification_cache: {
      contextHash: "TEXT NOT NULL DEFAULT ''",
    },
  };

  for (const [table, columns] of Object.entries(desiredColumns)) {
    const existing = new Set(db.prepare(`PRAGMA table_info(${table})`).all().map((column) => column.name));
    for (const [column, definition] of Object.entries(columns)) {
      if (!existing.has(column)) {
        db.prepare(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`).run();
      }
    }
  }
}

function seedAdminIfNeeded(db) {
  const count = db.prepare("SELECT COUNT(*) AS count FROM users").get().count;
  if (count > 0) return;

  const now = new Date().toISOString();
  const passwordHash = bcrypt.hashSync("admin123", 12);
  db.prepare(
    `INSERT INTO users (name, email, passwordHash, accountType, companyName, role, avatarInitials, createdAt, updatedAt)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(
    "Admin",
    "admin@local",
    passwordHash,
    "Company",
    "Local Governance Office",
    "Platform Administrator",
    "AD",
    now,
    now
  );
}

module.exports = {
  initDatabase,
  getDb,
  createDatabaseBackup,
  databasePath,
  backupDir,
};
