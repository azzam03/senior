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
  const result = await requestWithResponse(url, options);
  return result.body;
}

async function requestWithResponse(url, options = {}) {
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
  return { body, headers: response.headers, status: response.status };
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

  const csvClassificationOrder = await runClassificationJob(csvSystem.system.id, "all");
  assertAscendingRowOrder(csvClassificationOrder, "CSV");

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

  const excelClassificationOrder = await runClassificationJob(excelSystem.system.id, "all");
  assertAscendingRowOrder(excelClassificationOrder, "Excel");
  const classifiedExcelRecords = await request(`/api/systems/${excelSystem.system.id}/records?page=1&pageSize=100`);
  const classifiedRow = findRecord(classifiedExcelRecords.records, "Employees", "FullName");
  if (!classifiedRow.confidentiality || classifiedRow.confidentiality === "Pending") {
    throw new Error("Confidentiality did not persist after classification.");
  }
  if (!Object.values(classifiedExcelRecords.summary.personalDataTypes || {}).some((count) => Number(count) > 0)) {
    throw new Error("Personal data type counts were not included in the system summary.");
  }
  if (!/Bahrain PDPL/i.test(`${classifiedRow.confReason} ${classifiedRow.personalReason}`)) {
    throw new Error("Bahrain-law-grounded reasoning was not stored with classification results.");
  }
  const jobParameterName = findRecord(classifiedExcelRecords.records, "JobParameter", "Name");
  const settingsValue = findRecord(classifiedExcelRecords.records, "Settings", "Value");
  if (jobParameterName.personalData !== "No") {
    throw new Error("JobParameter.Name was incorrectly marked as personal data.");
  }
  if (settingsValue.personalData !== "No") {
    throw new Error("Settings.Value was incorrectly marked as personal data.");
  }
  for (const [tableName, columnName] of [
    ["Users", "Email"],
    ["Users", "PhoneNumber"],
    ["Applicants", "NationalId"],
    ["Orders", "OwnerName"],
  ]) {
    const record = findRecord(classifiedExcelRecords.records, tableName, columnName);
    if (record.personalData !== "Yes") {
      throw new Error(`${tableName}.${columnName} was not marked as personal data.`);
    }
  }
  const allSecret = classifiedExcelRecords.records.every((record) => ["Secret", "Top Secret"].includes(record.confidentiality));
  const allPersonal = classifiedExcelRecords.records.every((record) => record.personalData === "Yes");
  if (allSecret || allPersonal) {
    throw new Error("System Context caused all rows to become Secret or Personal Data.");
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
  const editedRecords = await request(`/api/systems/${excelSystem.system.id}/records?page=1&pageSize=100`);
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
  const reviewedRecords = await request(`/api/systems/${excelSystem.system.id}/records?page=1&pageSize=100`);
  if (findRecord(reviewedRecords.records, "Employees", "FullName").systemReviewStatus !== "Approved") {
    throw new Error("System Data review state did not persist.");
  }
  const personalRows = await request(`/api/systems/${excelSystem.system.id}/records?personalOnly=true&page=1&pageSize=100`);
  if (personalRows.records.some((record) => record.tableName === "JobParameter" && record.columnName === "Name")) {
    throw new Error("Personal Data API included generic technical JobParameter.Name.");
  }
  await request(`/api/systems/${excelSystem.system.id}/personal-approvals`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ recordIds: [classifiedRow.id] }),
  });
  const pdpl = await request(`/api/systems/${excelSystem.system.id}/pdpl`);
  const obligation = pdpl.obligations.find((item) => item.id === classifiedRow.id);
  if (!obligation || obligation.reviewStatus !== "Approved") {
    throw new Error("PDPL obligation review status was not driven by Personal Data approval.");
  }

  const exportResponse = await requestWithResponse(`/api/systems/${excelSystem.system.id}/export`);
  const exportBuffer = exportResponse.body;
  const exportContentType = exportResponse.headers.get("content-type") || "";
  const exportDisposition = exportResponse.headers.get("content-disposition") || "";
  if (!exportContentType.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
    throw new Error("Export response used the wrong workbook MIME type.");
  }
  if (!/attachment/i.test(exportDisposition) || !/\.xlsx/i.test(exportDisposition)) {
    throw new Error("Export response did not include an Excel attachment filename.");
  }
  const exportPath = path.join(tmpDir, "verification-export.xlsx");
  fs.writeFileSync(exportPath, Buffer.from(exportBuffer));
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(exportPath);
  const sheetNames = workbook.worksheets.map((sheet) => sheet.name);
  const requiredSheets = ["Overall", "System Data", "Personal Data", "PDPL"];
  for (const sheet of requiredSheets) {
    if (!sheetNames.includes(sheet)) throw new Error(`Missing export sheet: ${sheet}`);
  }
  const systemDataSheet = workbook.getWorksheet("System Data");
  const personalDataSheet = workbook.getWorksheet("Personal Data");
  const pdplSheet = workbook.getWorksheet("PDPL");
  const systemDataHeaders = systemDataSheet.getRow(1).values.join("|");
  if (!systemDataHeaders.includes("Source") || !systemDataHeaders.includes("Retention")) {
    throw new Error("Export did not preserve uploaded Excel columns.");
  }
  const systemRows = worksheetObjects(systemDataSheet);
  if (systemRows.length !== classifiedExcelRecords.records.length) {
    throw new Error("System Data export did not include all imported metadata records.");
  }
  const exportedReviewedRow = systemRows.find((row) => row.TableName === "Employees" && row.ColumnName === "FullName");
  if (
    !exportedReviewedRow ||
    exportedReviewedRow.Confidentiality !== "Secret" ||
    exportedReviewedRow["Personal Data"] !== "Yes" ||
    exportedReviewedRow["System Data Review"] !== "Approved"
  ) {
    throw new Error("System Data export did not include final reviewed classification values.");
  }
  const personalExportRows = worksheetObjects(personalDataSheet);
  if (personalExportRows.some((row) => row.TableName === "JobParameter" && row.ColumnName === "Name")) {
    throw new Error("Personal Data export included generic technical JobParameter.Name.");
  }
  const pdplRows = worksheetObjects(pdplSheet, 8);
  if (!pdplRows.some((row) => row.Table === "Employees" && row.Column === "FullName" && row["PDPL Approval Status"] === "Approved")) {
    throw new Error("PDPL export did not include the approved personal data record.");
  }
  const unapprovedPersonalIds = new Set(personalRows.records.filter((record) => record.id !== classifiedRow.id).map((record) => `${record.tableName}.${record.columnName}`));
  if (pdplRows.some((row) => unapprovedPersonalIds.has(`${row.Table}.${row.Column}`))) {
    throw new Error("PDPL export included unapproved personal data records.");
  }
  if (!worksheetContains(workbook.getWorksheet("Overall"), "Personal Data Type Counts")) {
    throw new Error("Overall export sheet did not include personal data type counts.");
  }
  await request(`/api/systems/${csvSystem.system.id}`, {
    method: "DELETE",
  });
  let deleteVerified = false;
  try {
    await request(`/api/systems/${csvSystem.system.id}`);
  } catch (error) {
    deleteVerified = /not found/i.test(error.message);
  }
  if (!deleteVerified) {
    throw new Error("Deleted system was still accessible after deletion.");
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
    deleteVerified,
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
  sheet.addRow(["JobParameter", "Name", "nvarchar", "HangFire", "30 days"]);
  sheet.addRow(["Settings", "Value", "nvarchar", "Configuration", "30 days"]);
  sheet.addRow(["Users", "Email", "nvarchar", "IAM", "7 years"]);
  sheet.addRow(["Users", "PhoneNumber", "nvarchar", "IAM", "7 years"]);
  sheet.addRow(["Applicants", "NationalId", "nvarchar", "Recruitment", "10 years"]);
  sheet.addRow(["Orders", "OwnerName", "nvarchar", "Sales", "5 years"]);
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

function worksheetContains(sheet, expected) {
  const needle = String(expected).toLowerCase();
  return sheet.getSheetValues().some((row) =>
    Array.isArray(row) && row.some((cell) => String(cell || "").toLowerCase().includes(needle))
  );
}

function worksheetObjects(sheet, headerRowNumber = 1) {
  const headers = sheet.getRow(headerRowNumber).values.slice(1).map((value) => String(value || ""));
  const rows = [];
  for (let rowNumber = headerRowNumber + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const values = row.values.slice(1);
    if (!values.some((value) => value != null && String(value) !== "")) continue;
    const item = {};
    headers.forEach((header, index) => {
      if (header) item[header] = values[index] == null ? "" : values[index];
    });
    rows.push(item);
  }
  return rows;
}

function findRecord(records, tableName, columnName) {
  const record = records.find((item) => item.tableName === tableName && item.columnName === columnName);
  if (!record) throw new Error(`Missing record ${tableName}.${columnName}`);
  return record;
}

function assertAscendingRowOrder(rowIndexes, label) {
  for (let index = 1; index < rowIndexes.length; index += 1) {
    if (rowIndexes[index] < rowIndexes[index - 1]) {
      throw new Error(`${label} classification rows were not streamed in rowIndex order.`);
    }
  }
}

async function runClassificationJob(systemId, mode) {
  const result = await request(`/api/systems/${systemId}/classification-jobs`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      mode,
      page: 1,
      pageSize: 25,
      sortBy: "rowIndex",
      sortDir: "asc",
    }),
  });
  return consumeSse(`/api/systems/${systemId}/classification-jobs/${result.job.id}/stream`);
}

async function consumeSse(url) {
  const response = await fetch(`${BASE}${url}`, { headers: headers() });
  if (!response.ok) throw new Error(`SSE failed: ${response.status}`);
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let doneSeen = false;
  const rowIndexes = [];
  while (!doneSeen) {
    const read = await reader.read();
    if (read.done) break;
    buffer += decoder.decode(read.value, { stream: true });
    const events = buffer.split("\n\n");
    buffer = events.pop() || "";
    for (const event of events) {
      const parsed = parseSseEvent(event);
      if (parsed.eventName === "row" && parsed.data?.record?.rowIndex) {
        rowIndexes.push(Number(parsed.data.record.rowIndex));
      }
      if (parsed.eventName === "done") doneSeen = true;
      if (parsed.eventName === "classification-error") throw new Error(event);
    }
  }
  return rowIndexes;
}

function parseSseEvent(raw) {
  const lines = raw.split("\n");
  const eventName = (lines.find((line) => line.startsWith("event:")) || "event: message")
    .replace(/^event:\s*/, "")
    .trim();
  const data = lines
    .filter((line) => line.startsWith("data:"))
    .map((line) => line.replace(/^data:\s*/, ""))
    .join("\n");
  return {
    eventName,
    data: data ? JSON.parse(data) : null,
  };
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
