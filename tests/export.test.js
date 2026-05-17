const test = require("node:test");
const assert = require("node:assert/strict");
const ExcelJS = require("exceljs");

const { buildSystemDataExport } = require("../src/excel");

test("export workbook contains required sheets and reviewed data", async () => {
  const workbook = await buildSystemDataExport({
    system: {
      name: "Export Test System",
      owner: "Owner",
      dba: "DBA",
      ownerEmail: "owner@example.com",
      systemGroup: "Governance",
      lastUploadFileName: "metadata.xlsx",
      status: "Completed",
    },
    originalColumns: ["TableName", "ColumnName", "DataType", "Source"],
    rows: sampleRows(),
    summary: {
      totalRecords: 4,
      classified: 4,
      classifiedPercentage: 100,
      personal: 3,
      personalPercentage: 75,
      pending: 0,
      pendingPercentage: 0,
      reviewQueue: 1,
      lowConfidence: 0,
      tablesWithPersonalData: 3,
      policyRecommendationCount: 3,
      classificationProgress: 100,
      confidentiality: { Confidential: 2, Secret: 1, "Top Secret": 0, Public: 1, Pending: 0 },
      personalDataTypes: { Name: 1, "Contact Information": 1, "Government Identifier": 1 },
    },
    pdpl: {
      dataSubjectRightsCoverage: "Compliant",
      consentTrackingStatus: "Needs Review",
      crossBorderTransferFlags: "No Flags Recorded",
      governanceNotes: "Export governance notes",
    },
    contextPoints: [
      {
        tag: "System Description",
        content: "HR and payroll platform with applicant tracking.",
        createdBy: "owner@example.com",
        createdAt: "2026-05-01T00:00:00.000Z",
        updatedAt: "2026-05-10T00:00:00.000Z",
      },
      {
        tag: "Data Sources",
        content: "Internal HR database and recruitment portal.",
        createdBy: "owner@example.com",
        createdAt: "2026-05-02T00:00:00.000Z",
        updatedAt: "2026-05-02T00:00:00.000Z",
      },
    ],
  });

  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  assert.ok(buffer.length > 0);

  const parsed = new ExcelJS.Workbook();
  await parsed.xlsx.load(buffer);
  // Exactly four sheets, in the required order. No extra sheets, no second workbook, no companion file.
  assert.deepEqual(
    parsed.worksheets.map((sheet) => sheet.name),
    ["Overall", "System Data", "Personal Data", "PDPL"]
  );

  const systemRows = worksheetObjects(parsed.getWorksheet("System Data"));
  assert.equal(systemRows.length, 4);
  const fullName = findRow(systemRows, "Employees", "FullName");
  assert.equal(fullName.Confidentiality, "Secret");
  assert.equal(fullName["Personal Data"], "Yes");
  assert.equal(fullName["System Data Review"], "Approved");

  const personalRows = worksheetObjects(parsed.getWorksheet("Personal Data"));
  assert.equal(personalRows.some((row) => row.TableName === "JobParameter" && row.ColumnName === "Name"), false);
  assert.equal(personalRows.some((row) => row.TableName === "Users" && row.ColumnName === "Email"), true);

  const pdplRows = worksheetObjects(parsed.getWorksheet("PDPL"), 8);
  assert.equal(pdplRows.some((row) => row.Table === "Employees" && row.Column === "FullName"), true);
  assert.equal(pdplRows.some((row) => row.Table === "Applicants" && row.Column === "NationalId"), true);
  assert.equal(pdplRows.some((row) => row.Table === "Users" && row.Column === "Email"), false);
  assert.equal(pdplRows.every((row) => row["PDPL Approval Status"] === "Approved"), true);
});

test("export workbook with no data still produces the four required sheets", async () => {
  const workbook = await buildSystemDataExport({
    system: { name: "No Context System" },
    originalColumns: ["TableName", "ColumnName"],
    rows: [],
    summary: { totalRecords: 0 },
    pdpl: {},
    contextPoints: [],
  });
  const parsed = new ExcelJS.Workbook();
  await parsed.xlsx.load(Buffer.from(await workbook.xlsx.writeBuffer()));
  assert.deepEqual(
    parsed.worksheets.map((sheet) => sheet.name),
    ["Overall", "System Data", "Personal Data", "PDPL"]
  );
});

function sampleRows() {
  return [
    {
      id: 1,
      original: { TableName: "Employees", ColumnName: "FullName", DataType: "nvarchar", Source: "HRMS" },
      tableName: "Employees",
      columnName: "FullName",
      confidentiality: "Secret",
      confReason: "Manual governance edit verified against Bahrain PDPL handling expectations.",
      personalData: "Yes",
      personalReason: "Reviewed as employee personal data.",
      personalDataType: "Name",
      pseudonymizable: "Yes",
      anonymizable: "Yes",
      specialCategory: "No",
      confidenceScore: 0.95,
      needsReview: 0,
      policyRecommendation: "restricted access",
      systemReviewStatus: "Approved",
      personalReviewStatus: "Approved",
      personalApprovedBy: "reviewer@example.com",
      personalApprovedAt: "2026-05-16T00:00:00.000Z",
      pushToClient: 1,
    },
    {
      id: 2,
      original: { TableName: "JobParameter", ColumnName: "Name", DataType: "nvarchar", Source: "HangFire" },
      tableName: "JobParameter",
      columnName: "Name",
      confidentiality: "Confidential",
      confReason: "Technical metadata.",
      personalData: "No",
      personalReason: "Generic technical job parameter name.",
      personalDataType: "",
      pseudonymizable: "No",
      anonymizable: "Yes",
      specialCategory: "No",
      confidenceScore: 0.7,
      needsReview: 1,
      policyRecommendation: "technical metadata review",
      systemReviewStatus: "Unreviewed",
      personalReviewStatus: "Needs Review",
      pushToClient: 1,
    },
    {
      id: 3,
      original: { TableName: "Users", ColumnName: "Email", DataType: "nvarchar", Source: "IAM" },
      tableName: "Users",
      columnName: "Email",
      confidentiality: "Confidential",
      confReason: "Contact information.",
      personalData: "Yes",
      personalReason: "Email identifies a user.",
      personalDataType: "Contact Information",
      pseudonymizable: "Yes",
      anonymizable: "Yes",
      specialCategory: "No",
      confidenceScore: 0.9,
      needsReview: 1,
      policyRecommendation: "masking recommended",
      systemReviewStatus: "Pending Review",
      personalReviewStatus: "Needs Review",
      pushToClient: 1,
    },
    {
      id: 4,
      original: { TableName: "Applicants", ColumnName: "NationalId", DataType: "nvarchar", Source: "Recruitment" },
      tableName: "Applicants",
      columnName: "NationalId",
      confidentiality: "Secret",
      confReason: "Government identifier.",
      personalData: "Yes",
      personalReason: "National ID identifies an applicant.",
      personalDataType: "Government Identifier",
      pseudonymizable: "Yes",
      anonymizable: "Yes",
      specialCategory: "No",
      confidenceScore: 0.95,
      needsReview: 0,
      policyRecommendation: "restricted access",
      systemReviewStatus: "Approved",
      personalReviewStatus: "Approved",
      personalApprovedBy: "reviewer@example.com",
      personalApprovedAt: "2026-05-16T00:00:00.000Z",
      pushToClient: 1,
    },
  ];
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

function findRow(rows, tableName, columnName) {
  const row = rows.find((item) => item.TableName === tableName && item.ColumnName === columnName);
  assert.ok(row, `Missing export row ${tableName}.${columnName}`);
  return row;
}
