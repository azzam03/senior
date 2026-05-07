const fs = require("fs");
const path = require("path");
const bcrypt = require("bcryptjs");
const { DatabaseSync, backup } = require("node:sqlite");

const projectRoot = path.resolve(__dirname, "..");
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
  if (connection) {
    await backup(connection, destination);
  } else {
    fs.copyFileSync(databasePath, destination);
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
  connection.exec("PRAGMA journal_mode = WAL");
  connection.exec("PRAGMA foreign_keys = ON");
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
