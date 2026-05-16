const test = require("node:test");
const assert = require("node:assert/strict");

const { classifyRecordLocally } = require("../src/classifier");

const strongSystemContext = [
  {
    tag: "System Description",
    content: "HR, payroll, customer service, national ID, user account, and applicant processing platform.",
  },
];

function classify(tableName, columnName, contextPoints = strongSystemContext) {
  return classifyRecordLocally({ tableName, columnName }, contextPoints);
}

const isSecretLevel = (result) => ["Secret", "Top Secret"].includes(result.confidentiality);

test("generic technical fields are not automatically personal data", () => {
  const jobParameterName = classify("JobParameter", "Name");
  const settingsValue = classify("Settings", "Value");
  const auditDescription = classify("AuditLogs", "Description");

  assert.equal(jobParameterName.personalData, "No");
  assert.equal(settingsValue.personalData, "No");
  assert.equal(auditDescription.personalData, "No");
  assert.notEqual(jobParameterName.confidentiality, "Secret");
  assert.notEqual(settingsValue.confidentiality, "Secret");
});

test("clearly personal fields are marked as personal data", () => {
  for (const [tableName, columnName] of [
    ["Users", "Email"],
    ["Users", "PhoneNumber"],
    ["Applicants", "NationalId"],
    ["Orders", "OwnerName"],
  ]) {
    const result = classify(tableName, columnName);
    assert.equal(result.personalData, "Yes", `${tableName}.${columnName}`);
    assert.notEqual(result.personalDataType, "", `${tableName}.${columnName}`);
  }
});

test("system context does not make every field Secret or Personal Data", () => {
  const results = [
    classify("JobParameter", "Name"),
    classify("Settings", "Value"),
    classify("LookupStatus", "Code"),
    classify("Orders", "OrderStatus"),
  ];

  assert.equal(results.every((result) => result.personalData === "Yes"), false);
  assert.equal(results.every((result) => isSecretLevel(result)), false);
});

// Spec-driven validation examples — each one validates an accuracy principle of the improved classifier.

test("pure system / technical fields are System Data, not Personal Data, not Secret", () => {
  const cases = [
    ["Orders", "created_at"],
    ["Orders", "updated_at"],
    ["Orders", "status"],
    ["EventLogs", "log_id"],
    ["RequestLogs", "service_name"],
    ["RequestLogs", "error_code"],
    ["Telemetry", "event_type"],
    ["Telemetry", "status_code"],
  ];
  for (const [table, column] of cases) {
    const result = classify(table, column);
    assert.equal(result.personalData, "No", `${table}.${column} should not be Personal Data`);
    assert.equal(isSecretLevel(result), false, `${table}.${column} should not be Secret`);
  }
});

test("personal identifiers are Personal Data with a personalDataType", () => {
  const cases = [
    ["Employees", "full_name"],
    ["Employees", "email"],
    ["Employees", "phone_number"],
    ["Applicants", "national_id"],
    ["Applicants", "passport_number"],
    ["Employees", "date_of_birth"],
  ];
  for (const [table, column] of cases) {
    const result = classify(table, column);
    assert.equal(result.personalData, "Yes", `${table}.${column} should be Personal Data`);
    assert.notEqual(result.personalDataType, "", `${table}.${column} should have a personalDataType`);
  }
});

test("credentials and authentication secrets are Secret or Top Secret", () => {
  const cases = [
    ["Users", "password"],
    ["ApiClients", "api_key"],
    ["AuthSessions", "access_token"],
    ["AuthSessions", "refresh_token"],
    ["Certificates", "private_key"],
    ["RequestLogs", "authorization_header"],
    ["Secrets", "client_secret"],
    ["Vault", "encryption_key"],
  ];
  for (const [table, column] of cases) {
    const result = classify(table, column);
    assert.equal(isSecretLevel(result), true, `${table}.${column} should be Secret/Top Secret`);
  }
});

test("user_id is Personal Data only in subject tables (not in technical/log tables)", () => {
  const inUserTable = classify("Users", "user_id");
  assert.equal(inUserTable.personalData, "Yes");

  const inAuditTable = classify("AuditLogs", "user_id");
  assert.equal(inAuditTable.personalData, "No", "user_id in an audit/log table should be System Data");
  assert.equal(isSecretLevel(inAuditTable), false);
});

test("session_id and cookie_id default to System Data unless inside a subject table", () => {
  const sessionInLogs = classify("SessionLogs", "session_id");
  assert.equal(sessionInLogs.personalData, "No");
  assert.equal(isSecretLevel(sessionInLogs), false);

  const cookieInTechTable = classify("RequestLogs", "cookie_id");
  assert.equal(cookieInTechTable.personalData, "No");
});

test("ip_address is Personal Data when tied to a user subject, System Data otherwise", () => {
  const ipOnUser = classify("Users", "ip_address");
  assert.equal(ipOnUser.personalData, "Yes", "users.ip_address should be Personal Data");

  const ipOnLogs = classify("RequestLogs", "ip_address");
  assert.equal(ipOnLogs.personalData, "No", "request_logs.ip_address should default to System Data");
});

test("generic attributes do not auto-become Personal Data or Secret", () => {
  const cases = [
    ["Employees", "department"],
    ["Employees", "role"],
    ["Orders", "status"],
    ["Products", "category"],
    ["Tickets", "type"],
  ];
  for (const [table, column] of cases) {
    const result = classify(table, column);
    assert.equal(result.personalData, "No", `${table}.${column} should not be Personal Data`);
    assert.equal(isSecretLevel(result), false, `${table}.${column} should not be Secret`);
  }
});

test("each local classification carries an internal confidence score and evidence array", () => {
  const result = classify("Users", "password");
  assert.equal(typeof result.confidenceScore, "number");
  assert.ok(result.confidenceScore >= 0 && result.confidenceScore <= 1);
  assert.ok(Array.isArray(result.evidence));
  assert.ok(result.evidence.length > 0, "credentials should produce evidence");
});

test("balanced distribution: a realistic mix is not all Personal and not all Secret", () => {
  const mix = [
    classify("Users", "email"),
    classify("Users", "password"),
    classify("Orders", "created_at"),
    classify("Orders", "status"),
    classify("RequestLogs", "service_name"),
    classify("AuditLogs", "user_id"),
    classify("Employees", "department"),
  ];
  const personalCount = mix.filter((r) => r.personalData === "Yes").length;
  const secretCount = mix.filter((r) => isSecretLevel(r)).length;
  assert.ok(personalCount < mix.length, "not every record should be Personal Data");
  assert.ok(secretCount < mix.length, "not every record should be Secret");
  assert.ok(personalCount >= 1, "at least one record should be Personal Data");
  assert.ok(secretCount >= 1, "at least one record should be Secret");
});
