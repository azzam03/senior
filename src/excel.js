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
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "DATA CLASSIFICATION & GOVERNANCE PLATFORM";
  workbook.created = new Date();

  buildOverallSheet(workbook, { system, summary, pdpl });
  buildDataSheet(workbook, "System Data", rows, originalColumns, false);
  buildDataSheet(workbook, "Personal Data", rows.filter((row) => row.personalData === "Yes"), originalColumns, true);
  buildPdplSheet(workbook, {
    system,
    rows: rows.filter((row) => row.personalData === "Yes"),
    pdpl,
  });

  return workbook;
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

  addSectionHeader(sheet, 29, "Governance and PDPL Summary");
  addSummaryTable(sheet, 30, [
    ["Indicator", "Status"],
    ["Data Subject Rights Coverage", pdpl?.dataSubjectRightsCoverage || "Needs Review"],
    ["Consent Tracking Status", pdpl?.consentTrackingStatus || "Needs Review"],
    ["Cross-Border Transfer Flags", pdpl?.crossBorderTransferFlags || "No Flags Recorded"],
    ["Policy Recommendations", summary.policyRecommendationCount],
    ["Classification Progress", `${summary.classificationProgress}%`],
  ]);

  applyOuterBorders(sheet, "A1:F36");
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
  sheet.getRow(headerRowNumber).values = [null, ...headers];
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
