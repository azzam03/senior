const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const BASE = process.env.PLATFORM_URL || "http://localhost:3000";
const root = path.resolve(__dirname, "..");
const tmpDir = path.join(root, "tmp");

let cookie = "";

function headers(extra = {}) {
  return cookie ? { Cookie: cookie, ...extra } : extra;
}

async function request(url, options = {}) {
  const response = await fetch(`${BASE}${url}`, {
    ...options,
    headers: headers(options.headers || {}),
  });
  const setCookie = response.headers.get("set-cookie");
  if (setCookie) cookie = setCookie.split(";")[0];
  const contentType = response.headers.get("content-type") || "";
  const body = contentType.includes("application/json") ? await response.json() : await response.arrayBuffer();
  if (!response.ok) {
    throw new Error(typeof body === "object" && body.error ? body.error : `Request failed: ${response.status}`);
  }
  return body;
}

async function main() {
  fs.mkdirSync(tmpDir, { recursive: true });
  const stamp = Date.now();
  const email = `verify-${stamp}@local`;

  const protectedResponse = await fetch(`${BASE}/app`, { redirect: "manual" });
  if (![301, 302, 303, 307, 308].includes(protectedResponse.status)) {
    throw new Error("Protected app URL did not redirect unauthenticated access.");
  }

  await request("/api/auth/signup", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: "Verification User",
      email,
      password: "verify123",
      accountType: "Company",
      companyName: "Verification Co",
    }),
  });

  const csvSystem = await request("/api/systems", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: `Verification CSV CRM ${stamp}`,
      owner: "Data Owner",
      dba: "DBA Lead",
      ownerEmail: "owner@example.com",
      systemGroup: "Customer Platforms",
      assignedConsultant: "Consultant",
      sourceSystemRef: "CRM",
      targetSystemRef: "Analytics",
      relatedSystemLinks: "ERP, DWH",
    }),
  });

  await request(`/api/systems/${csvSystem.system.id}/context`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      tag: "System Description",
      content: "Customer relationship management platform used by service agents and governance reviewers.",
    }),
  });

  const csvPath = path.join(tmpDir, "verify-metadata.csv");
  fs.writeFileSync(
    csvPath,
    "TableName,ColumnName,DataType,Source,Notes\nCustomers,EmailAddress,varchar,CRM,preserved\nCustomers,NationalId,varchar,CRM,preserved\nOrders,OrderStatus,varchar,ERP,preserved\n",
    "utf8"
  );
  await uploadFile(csvSystem.system.id, csvPath, "text/csv");
  const csvRecords = await request(`/api/systems/${csvSystem.system.id}/records?page=1&pageSize=10`);
  if (!csvRecords.originalColumns.includes("Source") || !csvRecords.originalColumns.includes("Notes")) {
    throw new Error("CSV original columns were not preserved.");
  }

  await consumeSse(`/api/systems/${csvSystem.system.id}/classify-stream?mode=all`);

  const excelSystem = await request("/api/systems", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: `Verification Excel HR ${stamp}`,
      owner: "HR Data Owner",
      dba: "HR DBA",
      ownerEmail: "hr.owner@example.com",
      systemGroup: "Human Capital",
      assignedConsultant: "Consultant",
      sourceSystemRef: "HRMS",
      targetSystemRef: "Payroll Mart",
      relatedSystemLinks: "IAM, Finance",
    }),
  });

  await request(`/api/systems/${excelSystem.system.id}/context`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      tag: "System Personas",
      content: "Human resources analysts, payroll reviewers, and data protection officers use this system.",
    }),
  });

  const excelPath = path.join(tmpDir, "verify-metadata.xlsx");
  await createExcelFixture(excelPath);
  await uploadFile(excelSystem.system.id, excelPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  const excelRecords = await request(`/api/systems/${excelSystem.system.id}/records?page=1&pageSize=10`);
  if (!excelRecords.originalColumns.includes("Source") || !excelRecords.originalColumns.includes("Retention")) {
    throw new Error("Excel original columns were not preserved.");
  }

  await consumeSse(`/api/systems/${excelSystem.system.id}/classify-stream?mode=all`);
  const classifiedExcelRecords = await request(`/api/systems/${excelSystem.system.id}/records?page=1&pageSize=10`);
  const classifiedRow = classifiedExcelRecords.records[0];
  if (!classifiedRow.confidentiality || classifiedRow.confidentiality === "Pending") {
    throw new Error("Confidentiality did not persist after classification.");
  }
  if (!/Bahrain PDPL/i.test(`${classifiedRow.confReason} ${classifiedRow.personalReason}`)) {
    throw new Error("Bahrain-law-grounded reasoning was not stored with classification results.");
  }
  await request(`/api/records/${classifiedRow.id}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      confidentiality: "Secret",
      confReason: "Manual governance edit verified against Bahrain PDPL handling expectations.",
      personalData: "Yes",
      personalReason: "Manual edit confirms this employee field can identify an individual.",
    }),
  });
  const editedRecords = await request(`/api/systems/${excelSystem.system.id}/records?page=1&pageSize=10`);
  const editedRow = editedRecords.records.find((record) => record.id === classifiedRow.id);
  if (
    editedRow.confidentiality !== "Secret" ||
    editedRow.confReason !== "Manual governance edit verified against Bahrain PDPL handling expectations." ||
    editedRow.personalData !== "Yes" ||
    editedRow.personalReason !== "Manual edit confirms this employee field can identify an individual."
  ) {
    throw new Error("Edit View classification fields did not persist through the record API.");
  }
  await request(`/api/records/${classifiedRow.id}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ reviewed: true }),
  });
  const reviewedRecords = await request(`/api/systems/${excelSystem.system.id}/records?page=1&pageSize=10`);
  if (reviewedRecords.records[0].systemReviewStatus !== "Approved") {
    throw new Error("System Data review state did not persist.");
  }
  const personalRows = await request(`/api/systems/${excelSystem.system.id}/records?personalOnly=true&page=1&pageSize=100`);
  await request(`/api/systems/${excelSystem.system.id}/personal-approvals`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ recordIds: personalRows.records.map((record) => record.id) }),
  });
  const pdpl = await request(`/api/systems/${excelSystem.system.id}/pdpl`);
  const obligation = pdpl.obligations.find((item) => item.id === classifiedRow.id);
  if (!obligation || obligation.reviewStatus !== "Approved") {
    throw new Error("PDPL obligation review status was not driven by Personal Data approval.");
  }

  const exportBuffer = await request(`/api/systems/${excelSystem.system.id}/export`);
  const exportPath = path.join(tmpDir, "verification-export.xlsx");
  fs.writeFileSync(exportPath, Buffer.from(exportBuffer));
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(exportPath);
  const sheetNames = workbook.worksheets.map((sheet) => sheet.name);
  const requiredSheets = ["Overall", "System Data", "Personal Data", "PDPL"];
  for (const sheet of requiredSheets) {
    if (!sheetNames.includes(sheet)) throw new Error(`Missing export sheet: ${sheet}`);
  }
  const systemDataHeaders = workbook.getWorksheet("System Data").getRow(1).values.join("|");
  if (!systemDataHeaders.includes("Source") || !systemDataHeaders.includes("Retention")) {
    throw new Error("Export did not preserve uploaded Excel columns.");
  }

  const health = await request("/api/health");
  const dbExists = fs.existsSync(path.join(root, "database", "app.db"));
  const backupsExist = fs.existsSync(path.join(root, "database", "backups"));
  const backupFiles = backupsExist ? fs.readdirSync(path.join(root, "database", "backups")).filter((file) => file.endsWith(".db")) : [];

  console.log(JSON.stringify({
    ok: true,
    email,
    csvSystemId: csvSystem.system.id,
    excelSystemId: excelSystem.system.id,
    sheets: sheetNames,
    databasePath: health.databasePath,
    dbExists,
    backups: backupFiles.length,
    exportPath,
  }, null, 2));
}

async function createExcelFixture(filePath) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Metadata");
  sheet.addRow(["TableName", "ColumnName", "DataType", "Source", "Retention"]);
  sheet.addRow(["Employees", "FullName", "nvarchar", "HRMS", "7 years"]);
  sheet.addRow(["Employees", "SalaryAmount", "decimal", "Payroll", "10 years"]);
  sheet.addRow(["AuditEvents", "EventTime", "datetime", "Security", "2 years"]);
  await workbook.xlsx.writeFile(filePath);
}

async function uploadFile(systemId, filePath, type) {
  const formData = new FormData();
  const blob = new Blob([fs.readFileSync(filePath)], { type });
  formData.append("metadataFile", blob, path.basename(filePath));
  await request(`/api/systems/${systemId}/upload`, {
    method: "POST",
    body: formData,
  });
}

async function consumeSse(url) {
  const response = await fetch(`${BASE}${url}`, { headers: headers() });
  if (!response.ok) throw new Error(`SSE failed: ${response.status}`);
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let doneSeen = false;
  while (!doneSeen) {
    const read = await reader.read();
    if (read.done) break;
    buffer += decoder.decode(read.value, { stream: true });
    const events = buffer.split("\n\n");
    buffer = events.pop() || "";
    for (const event of events) {
      if (event.includes("event: done")) doneSeen = true;
      if (event.includes("event: error")) throw new Error(event);
    }
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
