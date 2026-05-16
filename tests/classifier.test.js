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
  assert.equal(results.every((result) => ["Secret", "Top Secret"].includes(result.confidentiality)), false);
});
