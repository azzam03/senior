const path = require("path");
const ExcelJS = require("exceljs");
const { parse } = require("csv-parse/sync");

const CONFIDENTIALITY_COLORS = {
  Public: "D8F3DC",
  Confidential: "FFF3BF",
  Secret: "FFE3E3",
  "Top Secret": "E5DBFF",
  Pending: "E9ECEF",
};

async function parseMetadataFile(buffer, fileName) {
  const ext = path.extname(fileName).toLowerCase();
  const rows = ext === ".csv" ? parseCsv(buffer) : await parseWorkbook(buffer);
  if (!rows.length) return { rows: [], columns: [] };

  const columns = Object.keys(rows[0]);
  const tableKey = findColumn(columns, "TableName");
  const columnKey = findColumn(columns, "ColumnName");
  if (!tableKey || !columnKey) {
    throw new Error("Uploaded metadata must include TableName and ColumnName columns.");
  }

  return {
    columns,
    rows: rows.map((row) => {
      const normalized = normalizeObject(row);
      return {
        original: normalized,
        tableName: String(normalized[tableKey] || "").trim(),
        columnName: String(normalized[columnKey] || "").trim(),
        dataType: String(
          normalized[findColumn(columns, "DataType")] ||
            normalized[findColumn(columns, "Data Type")] ||
            normalized[findColumn(columns, "Type")] ||
            ""
        ).trim(),
      };
    }).filter((row) => row.tableName && row.columnName),
  };
}

function parseCsv(buffer) {
  return parse(buffer.toString("utf8"), {
    bom: true,
    columns: true,
    skip_empty_lines: true,
    trim: true,
  });
}

async function parseWorkbook(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) return [];

  const headerRow = worksheet.getRow(1);
  const headers = [];
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    headers[colNumber - 1] = String(cell.value || "").trim();
  });

  const rows = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const item = {};
    let hasValue = false;
    headers.forEach((header, index) => {
      if (!header) return;
      const value = cellValue(row.getCell(index + 1).value);
      if (value !== "") hasValue = true;
      item[header] = value;
    });
    if (hasValue) rows.push(item);
  });
  return rows;
}

function normalizeObject(row) {
  const normalized = {};
  for (const [key, value] of Object.entries(row)) {
    normalized[String(key).trim()] = value == null ? "" : value;
  }
  return normalized;
}

function cellValue(value) {
  if (value == null) return "";
  if (typeof value === "object") {
    if (value.text) return value.text;
    if (value.result != null) return value.result;
    if (value.richText) return value.richText.map((part) => part.text).join("");
    if (value.hyperlink) return value.text || value.hyperlink;
  }
  return value;
}

function findColumn(columns, expected) {
  const target = expected.toLowerCase().replace(/[^a-z0-9]/g, "");
  return columns.find((column) => String(column).toLowerCase().replace(/[^a-z0-9]/g, "") === target);
}

async function buildSystemDataExport({ system, rows, originalColumns, summary, pdpl }) {
  const safeSystem = normalizeSystem(system);
  const safeRows = Array.isArray(rows) ? rows : [];
  const safeColumns = normalizeOriginalColumns(originalColumns, safeRows);
  const safeSummary = normalizeSummary(summary);
  const safePdpl = normalizePdpl(pdpl);
  const personalRows = safeRows.filter(isPersonalDataRecord);
  const approvedPersonalRows = personalRows.filter(isApprovedPersonalRecord);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "DATA CLASSIFICATION & GOVERNANCE PLATFORM";
  workbook.created = new Date();

  buildOverallSheet(workbook, { system: safeSystem, summary: safeSummary, pdpl: safePdpl });
  buildDataSheet(workbook, "System Data", safeRows, safeColumns, false);
  buildDataSheet(workbook, "Personal Data", personalRows, safeColumns, true);
  buildPdplSheet(workbook, { system: safeSystem, rows: approvedPersonalRows, pdpl: safePdpl });

  return workbook;
}

function normalizeSystem(system = {}) {
  return {
    name: system.name || "System",
    owner: system.owner || "",
    dba: system.dba || "",
    ownerEmail: system.ownerEmail || "",
    systemGroup: system.systemGroup || "",
    assignedConsultant: system.assignedConsultant || "",
    sourceSystemRef: system.sourceSystemRef || "",
    targetSystemRef: system.targetSystemRef || "",
    relatedSystemLinks: system.relatedSystemLinks || "",
    lastUploadFileName: system.lastUploadFileName || "",
    status: system.status || "In Progress",
  };
}

function normalizeSummary(summary = {}) {
  const totalRecords = Number(summary.totalRecords || 0);
  const classified = Number(summary.classified || 0);
  const personal = Number(summary.personal || 0);
  const pending = Number(summary.pending || Math.max(0, totalRecords - classified));
  const confidentiality = summary.confidentiality || {};
  return {
    totalRecords,
    classified,
    classifiedPercentage: numberOrPercent(summary.classifiedPercentage, classified, totalRecords),
    personal,
    personalPercentage: numberOrPercent(summary.personalPercentage, personal, totalRecords),
    pending,
    pendingPercentage: numberOrPercent(summary.pendingPercentage, pending, totalRecords),
    reviewQueue: Number(summary.reviewQueue || 0),
    lowConfidence: Number(summary.lowConfidence || 0),
    tablesWithPersonalData: Number(summary.tablesWithPersonalData || 0),
    policyRecommendationCount: Number(summary.policyRecommendationCount || 0),
    classificationProgress: numberOrPercent(summary.classificationProgress, classified, totalRecords),
    confidentiality: {
      Confidential: Number(confidentiality.Confidential || 0),
      Secret: Number(confidentiality.Secret || 0),
      "Top Secret": Number(confidentiality["Top Secret"] || 0),
      Public: Number(confidentiality.Public || 0),
      Pending: Number(confidentiality.Pending || 0),
    },
    personalDataTypes: summary.personalDataTypes || {},
  };
}

function numberOrPercent(value, part, total) {
  if (value != null && value !== "") return Number(value || 0);
  return total ? Math.round((Number(part || 0) / Number(total || 1)) * 100) : 0;
}

function normalizePdpl(pdpl = {}) {
  return {
    dataSubjectRightsCoverage: pdpl?.dataSubjectRightsCoverage || "Needs Review",
    consentTrackingStatus: pdpl?.consentTrackingStatus || "Needs Review",
    crossBorderTransferFlags: pdpl?.crossBorderTransferFlags || "No Flags Recorded",
    governanceNotes: pdpl?.governanceNotes || "",
  };
}

function normalizeOriginalColumns(originalColumns, rows) {
  const columns = [];
  const seen = new Set();
  const add = (column) => {
    const label = String(column || "").trim();
    if (!label || seen.has(label)) return;
    seen.add(label);
    columns.push(label);
  };

  if (Array.isArray(originalColumns)) originalColumns.forEach(add);
  rows.forEach((row) => {
    Object.keys(row.original || {}).forEach(add);
  });
  if (!columns.length && rows.length) {
    ["TableName", "ColumnName", "DataType"].forEach(add);
  }
  return columns;
}

function isPersonalDataRecord(row) {
  return String(row?.personalData || "").toLowerCase() === "yes";
}

function isApprovedPersonalRecord(row) {
  return isPersonalDataRecord(row) && String(row?.personalReviewStatus || "").toLowerCase() === "approved";
}

function buildOverallSheet(workbook, { system, summary, pdpl }) {
  const sheet = workbook.addWorksheet("Overall", {
    views: [{ showGridLines: false }],
  });
  sheet.columns = [
    { width: 28 },
    { width: 20 },
    { width: 22 },
    { width: 22 },
    { width: 22 },
    { width: 26 },
  ];

  sheet.mergeCells("A1:F1");
  sheet.getCell("A1").value = "DATA CLASSIFICATION & GOVERNANCE PLATFORM";
  sheet.getCell("A1").font = { bold: true, size: 18, color: { argb: "FFFFFFFF" } };
  sheet.getCell("A1").fill = fill("1D4ED8");
  sheet.getCell("A1").alignment = { vertical: "middle" };
  sheet.getRow(1).height = 30;

  sheet.mergeCells("A3:F3");
  sheet.getCell("A3").value = `${system.name} - Client Classification Report`;
  sheet.getCell("A3").font = { bold: true, size: 16, color: { argb: "FF111827" } };

  addKeyValue(sheet, 5, "Responsible Owner", system.owner, "DBA", system.dba);
  addKeyValue(sheet, 6, "Owner Email", system.ownerEmail, "System Group", system.systemGroup);
  addKeyValue(sheet, 7, "Uploaded File", system.lastUploadFileName || "No file uploaded", "Status", system.status);
  addKeyValue(sheet, 8, "Source System", system.sourceSystemRef || "Not recorded", "Target System", system.targetSystemRef || "Not recorded");

  addSectionHeader(sheet, 10, "Executive Classification Summary");
  const cards = [
    ["Fully Classified", summary.classified, `${summary.classifiedPercentage}%`],
    ["Personal Data", summary.personal, `${summary.personalPercentage}%`],
    ["Pending Classification", summary.pending, `${summary.pendingPercentage}%`],
    ["Review Queue", summary.reviewQueue, "items"],
    ["Low Confidence", summary.lowConfidence, "items"],
    ["Tables With Personal Data", summary.tablesWithPersonalData, "tables"],
  ];
  let row = 11;
  for (let index = 0; index < cards.length; index += 3) {
    addMetricCard(sheet, row, 1, cards[index]);
    addMetricCard(sheet, row, 3, cards[index + 1]);
    addMetricCard(sheet, row, 5, cards[index + 2]);
    row += 4;
  }

  addSectionHeader(sheet, 20, "Confidentiality Distribution");
  addSummaryTable(sheet, 21, [
    ["Confidentiality", "Count"],
    ["Confidential", summary.confidentiality.Confidential || 0],
    ["Secret", summary.confidentiality.Secret || 0],
    ["Top Secret", summary.confidentiality["Top Secret"] || 0],
    ["Public", summary.confidentiality.Public || 0],
    ["Pending", summary.confidentiality.Pending || 0],
  ]);

  const personalTypeRows = personalDataTypeRows(summary);
  const personalTypesStart = 29;
  addSectionHeader(sheet, personalTypesStart, "Personal Data Type Counts");
  addSummaryTable(sheet, personalTypesStart + 1, personalTypeRows);

  const governanceStart = personalTypesStart + personalTypeRows.length + 3;
  const governanceRows = [
    ["Indicator", "Status"],
    ["Data Subject Rights Coverage", pdpl?.dataSubjectRightsCoverage || "Needs Review"],
    ["Consent Tracking Status", pdpl?.consentTrackingStatus || "Needs Review"],
    ["Cross-Border Transfer Flags", pdpl?.crossBorderTransferFlags || "No Flags Recorded"],
    ["Policy Recommendations", summary.policyRecommendationCount],
    ["Classification Progress", `${summary.classificationProgress}%`],
  ];
  addSectionHeader(sheet, governanceStart, "Governance and PDPL Summary");
  addSummaryTable(sheet, governanceStart + 1, governanceRows);

  applyOuterBorders(sheet, `A1:F${governanceStart + governanceRows.length + 1}`);
}

function addKeyValue(sheet, row, leftKey, leftValue, rightKey, rightValue) {
  sheet.getCell(row, 1).value = leftKey;
  sheet.getCell(row, 2).value = leftValue;
  sheet.getCell(row, 4).value = rightKey;
  sheet.getCell(row, 5).value = rightValue;
  [1, 4].forEach((column) => {
    sheet.getCell(row, column).font = { bold: true, color: { argb: "FF374151" } };
  });
}

function addSectionHeader(sheet, row, title) {
  sheet.mergeCells(row, 1, row, 6);
  const cell = sheet.getCell(row, 1);
  cell.value = title;
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.fill = fill("334155");
  cell.alignment = { vertical: "middle" };
}

function addMetricCard(sheet, row, col, card) {
  if (!card) return;
  sheet.mergeCells(row, col, row, col + 1);
  sheet.mergeCells(row + 1, col, row + 2, col + 1);
  sheet.getCell(row, col).value = card[0];
  sheet.getCell(row, col).font = { bold: true, color: { argb: "FF475569" } };
  sheet.getCell(row + 1, col).value = `${card[1]} ${card[2]}`;
  sheet.getCell(row + 1, col).font = { bold: true, size: 18, color: { argb: "FF1D4ED8" } };
  [row, row + 1, row + 2].forEach((r) => {
    for (let c = col; c <= col + 1; c += 1) {
      sheet.getCell(r, c).fill = fill("EFF6FF");
      sheet.getCell(r, c).border = border("BFDBFE");
    }
  });
}

function personalDataTypeRows(summary) {
  const distribution = summary?.personalDataTypes || {};
  const entries = Object.entries(distribution)
    .map(([label, count]) => [label || "Personal Data", Number(count || 0)])
    .filter(([, count]) => count > 0)
    .sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
  return [
    ["Personal Data Type", "Count"],
    ...(entries.length ? entries : [["No personal data identified", 0]]),
  ];
}

function addSummaryTable(sheet, startRow, values) {
  values.forEach((items, offset) => {
    const row = sheet.getRow(startRow + offset);
    row.values = [null, ...items];
    row.eachCell((cell) => {
      cell.border = border("CBD5E1");
      cell.alignment = { vertical: "middle" };
      if (offset === 0) {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = fill("1D4ED8");
      }
    });
  });
}

function buildDataSheet(workbook, name, rows, originalColumns, personalOnly) {
  const sheet = workbook.addWorksheet(name, {
    views: [{ state: "frozen", ySplit: 1, showGridLines: false }],
  });
  const governanceColumns = personalOnly
    ? [
        "Confidentiality",
        "Conf Reason",
        "Personal Reason",
        "Personal Data Type",
        "Can be pseudonymized",
        "Can be anonymized",
        "Special Category",
        "Audit Trail",
        "Policy Recommendation",
        "Approval Status",
        "Approved By",
        "Approved At",
      ]
    : [
        "Confidentiality",
        "Conf Reason",
        "Personal Data",
        "Personal Reason",
        "Personal Data Type",
        "Can be pseudonymized",
        "Can be anonymized",
        "Special Category",
        "Confidence Score",
        "System Data Review",
        "Needs Review",
        "Policy Recommendation",
        "Owner",
        "Steward",
        "Reviewer",
        "Push to Client",
        "Last modified by",
        "Last reviewed at",
      ];
  const headers = [...originalColumns, ...governanceColumns];
  sheet.columns = headers.map((header) => ({ header, key: header, width: Math.min(34, Math.max(14, header.length + 4)) }));

  const headerRow = sheet.getRow(1);
  headerRow.height = 24;
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = fill(personalOnly ? "0F766E" : "1D4ED8");
    cell.border = border("93C5FD");
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });

  rows.forEach((record) => {
    const row = {};
    originalColumns.forEach((column) => {
      row[column] = record.original[column] ?? "";
    });
    Object.assign(row, {
      Confidentiality: record.confidentiality,
      "Conf Reason": record.confReason,
      "Personal Data": record.personalData,
      "Personal Reason": record.personalReason,
      "Personal Data Type": record.personalDataType,
      "Can be pseudonymized": record.pseudonymizable,
      "Can be anonymized": record.anonymizable,
      "Special Category": record.specialCategory,
      "Audit Trail": record.auditTrail,
      "Needs Review": record.needsReview ? "Yes" : "No",
      "Confidence Score": record.confidenceScore == null ? "" : record.confidenceScore,
      "Policy Recommendation": record.policyRecommendation,
      Owner: record.owner,
      Steward: record.steward,
      Reviewer: record.reviewer,
      "System Data Review": record.systemReviewStatus,
      "Approval Status": record.personalReviewStatus,
      "Approved By": record.personalApprovedBy,
      "Approved At": record.personalApprovedAt,
      "Push to Client": record.pushToClient ? "Yes" : "No",
      "Last modified by": record.lastModifiedBy,
      "Last reviewed at": record.lastReviewedAt,
    });

    const added = sheet.addRow(row);
    const color = CONFIDENTIALITY_COLORS[record.confidentiality] || CONFIDENTIALITY_COLORS.Pending;
    if (added.number % 2 === 0) {
      added.eachCell((cell) => {
        cell.fill = fill("F8FAFC");
      });
    }
    added.eachCell((cell) => {
      cell.border = border("E2E8F0");
      cell.alignment = { vertical: "top", wrapText: true };
    });
    const confidentialityIndex = headers.indexOf("Confidentiality") + 1;
    if (confidentialityIndex > 0) {
      added.getCell(confidentialityIndex).fill = fill(color);
      added.getCell(confidentialityIndex).font = { bold: true, color: { argb: "FF111827" } };
    }
    if (record.needsReview) {
      const needsReviewIndex = headers.indexOf("Needs Review") + 1;
      if (needsReviewIndex > 0) added.getCell(needsReviewIndex).fill = fill("FEF3C7");
    }
    if (record.personalData === "Yes" && headers.includes("Personal Data")) {
      added.getCell(headers.indexOf("Personal Data") + 1).fill = fill("DBEAFE");
      added.getCell(headers.indexOf("Personal Data") + 1).font = { bold: true };
    }
    if (headers.includes("Approval Status")) {
      const approvalCell = added.getCell(headers.indexOf("Approval Status") + 1);
      approvalCell.fill = fill(record.personalReviewStatus === "Approved" ? "DCFCE7" : "FEF3C7");
      approvalCell.font = { bold: true, color: { argb: record.personalReviewStatus === "Approved" ? "FF15803D" : "FF92400E" } };
    }
  });

  sheet.autoFilter = {
    from: "A1",
    to: `${sheet.getColumn(headers.length).letter}1`,
  };
}

function buildLinksAndGeneralInfoSheet(workbook, { system, summary, pdpl, links }) {
  const sheet = workbook.addWorksheet("Links & General Info", {
    views: [{ showGridLines: false }],
  });
  sheet.columns = [
    { width: 28 },
    { width: 34 },
    { width: 24 },
    { width: 48 },
    { width: 22 },
    { width: 28 },
  ];

  sheet.mergeCells("A1:F1");
  sheet.getCell("A1").value = `${system.name} - Links and General Information`;
  sheet.getCell("A1").font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
  sheet.getCell("A1").fill = fill("0F766E");
  sheet.getCell("A1").alignment = { vertical: "middle" };
  sheet.getRow(1).height = 30;

  addSectionHeader(sheet, 3, "System Information");
  addSummaryTable(sheet, 4, [
    ["Field", "Value"],
    ["System Name", system.name],
    ["Responsible Owner", system.owner],
    ["DBA", system.dba],
    ["Owner Email", system.ownerEmail],
    ["System Group", system.systemGroup],
    ["Assigned Consultant", system.assignedConsultant || ""],
    ["Source System Reference", system.sourceSystemRef || ""],
    ["Target System Reference", system.targetSystemRef || ""],
    ["Related System Links", system.relatedSystemLinks || ""],
    ["Uploaded File", system.lastUploadFileName || "No file uploaded"],
    ["Status", system.status],
  ]);

  addSectionHeader(sheet, 18, "Classification Totals");
  addSummaryTable(sheet, 19, [
    ["Metric", "Value"],
    ["Total Records", summary.totalRecords],
    ["Fully Classified", summary.classified],
    ["Pending Classification", summary.pending],
    ["Personal Data Records", summary.personal],
    ["Review Queue", summary.reviewQueue],
    ["Low Confidence", summary.lowConfidence],
    ["Tables With Personal Data", summary.tablesWithPersonalData],
  ]);

  const personalTypeRows = personalDataTypeRows(summary);
  const personalTypesStart = 29;
  addSectionHeader(sheet, personalTypesStart, "Personal Data Type Counts");
  addSummaryTable(sheet, personalTypesStart + 1, personalTypeRows);

  const pdplStart = personalTypesStart + personalTypeRows.length + 3;
  const pdplRows = [
    ["Indicator", "Status"],
    ["Data Subject Rights Coverage", pdpl?.dataSubjectRightsCoverage || "Needs Review"],
    ["Consent Tracking Status", pdpl?.consentTrackingStatus || "Needs Review"],
    ["Cross-Border Transfer Flags", pdpl?.crossBorderTransferFlags || "No Flags Recorded"],
    ["Governance Notes", pdpl?.governanceNotes || ""],
  ];
  addSectionHeader(sheet, pdplStart, "PDPL General Notes");
  addSummaryTable(sheet, pdplStart + 1, pdplRows);

  const linksStart = pdplStart + pdplRows.length + 3;
  addSectionHeader(sheet, linksStart, "Reference Links");
  const headerRowNumber = linksStart + 1;
  const headers = ["Title", "URL", "Category", "Description", "Created By", "Updated At"];
  sheet.getRow(headerRowNumber).values = headers;
  sheet.getRow(headerRowNumber).height = 24;
  sheet.getRow(headerRowNumber).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = fill("1D4ED8");
    cell.border = border("93C5FD");
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });

  if (links.length) {
    links.forEach((link) => {
      const url = String(link.url || "");
      const added = sheet.addRow([
        link.title || "",
        url,               // populated but overridden below with a real hyperlink
        link.category || "",
        link.description || "",
        link.createdBy || "",
        link.updatedAt || "",
      ]);
      added.eachCell((cell) => {
        cell.border = border("E2E8F0");
        cell.alignment = { vertical: "top", wrapText: true };
      });
      // Make the URL cell a proper clickable hyperlink in Excel.
      if (url) {
        const urlCell = added.getCell(2);
        urlCell.value = { text: url, hyperlink: url };
        urlCell.font = { color: { argb: "FF1D4ED8" }, underline: true };
      }
    });
  } else {
    const added = sheet.addRow(["No links stored", "", "", "", "", ""]);
    added.eachCell((cell) => {
      cell.border = border("E2E8F0");
      cell.alignment = { vertical: "top", wrapText: true };
    });
  }

  sheet.autoFilter = {
    from: `A${headerRowNumber}`,
    to: `F${headerRowNumber}`,
  };
  applyOuterBorders(sheet, `A1:F${Math.max(headerRowNumber + Math.max(links.length, 1), headerRowNumber)}`);
}

function buildPdplSheet(workbook, { system, rows, pdpl }) {
  const sheet = workbook.addWorksheet("PDPL", {
    views: [{ state: "frozen", ySplit: 5, showGridLines: false }],
  });
  sheet.columns = [
    { width: 24 },
    { width: 24 },
    { width: 24 },
    { width: 22 },
    { width: 34 },
    { width: 24 },
    { width: 22 },
    { width: 22 },
  ];

  sheet.mergeCells("A1:H1");
  sheet.getCell("A1").value = `${system.name} - PDPL Obligations`;
  sheet.getCell("A1").font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
  sheet.getCell("A1").fill = fill("0F766E");
  sheet.getCell("A1").alignment = { vertical: "middle" };
  sheet.getRow(1).height = 30;

  addSummaryTable(sheet, 3, [
    ["Indicator", "Status"],
    ["Data Subject Rights Coverage", pdpl?.dataSubjectRightsCoverage || "Needs Review"],
    ["Consent Tracking Status", pdpl?.consentTrackingStatus || "Needs Review"],
    ["Cross-Border Transfer Flags", pdpl?.crossBorderTransferFlags || "No Flags Recorded"],
  ]);

  const headerRowNumber = 8;
  const headers = [
    "Table",
    "Column",
    "Personal Data Type",
    "PDPL Approval Status",
    "PDPL Notes / Obligations",
    "Policy Guidance",
    "Approved By",
    "Approved At",
  ];
  sheet.getRow(headerRowNumber).values = headers;
  sheet.getRow(headerRowNumber).height = 26;
  sheet.getRow(headerRowNumber).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = fill("1D4ED8");
    cell.border = border("93C5FD");
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });

  rows.forEach((record) => {
    const added = sheet.addRow([
      record.tableName,
      record.columnName,
      record.personalDataType || "Personal Data",
      record.personalReviewStatus || "Needs Review",
      pdpl?.governanceNotes || "Confirm lawful basis, access control, retention, and data subject rights handling.",
      record.policyRecommendation || "review needed",
      record.personalApprovedBy || "",
      record.personalApprovedAt || "",
    ]);
    if (added.number % 2 === 0) {
      added.eachCell((cell) => {
        cell.fill = fill("F8FAFC");
      });
    }
    added.eachCell((cell) => {
      cell.border = border("E2E8F0");
      cell.alignment = { vertical: "top", wrapText: true };
    });
    const statusCell = added.getCell(4);
    statusCell.fill = fill(record.personalReviewStatus === "Approved" ? "DCFCE7" : "FEF3C7");
    statusCell.font = { bold: true, color: { argb: record.personalReviewStatus === "Approved" ? "FF15803D" : "FF92400E" } };
  });

  sheet.autoFilter = {
    from: `A${headerRowNumber}`,
    to: `H${headerRowNumber}`,
  };
  applyOuterBorders(sheet, `A1:H${Math.max(headerRowNumber + rows.length, headerRowNumber)}`);
}

function fill(hex) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: `FF${hex}` } };
}

function border(hex) {
  return {
    top: { style: "thin", color: { argb: `FF${hex}` } },
    left: { style: "thin", color: { argb: `FF${hex}` } },
    bottom: { style: "thin", color: { argb: `FF${hex}` } },
    right: { style: "thin", color: { argb: `FF${hex}` } },
  };
}

function applyOuterBorders(sheet, range) {
  const [start, end] = range.split(":");
  const startCell = sheet.getCell(start);
  const endCell = sheet.getCell(end);
  for (let row = startCell.row; row <= endCell.row; row += 1) {
    for (let col = startCell.col; col <= endCell.col; col += 1) {
      sheet.getCell(row, col).border = border("E2E8F0");
    }
  }
}

module.exports = {
  parseMetadataFile,
  buildSystemDataExport,
  CONFIDENTIALITY_COLORS,
};
