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

async function buildSystemDataExport({ system, rows, originalColumns, summary, pdpl, contextPoints }) {
  const safeSystem = normalizeSystem(system);
  const safeRows = Array.isArray(rows) ? rows : [];
  const safeColumns = normalizeOriginalColumns(originalColumns, safeRows);
  const safeSummary = normalizeSummary(summary);
  const safePdpl = normalizePdpl(pdpl);
  const safeContextPoints = normalizeContextPoints(contextPoints);
  const personalRows = safeRows.filter(isPersonalDataRecord);
  const approvedPersonalRows = personalRows.filter(isApprovedPersonalRecord);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "DATA CLASSIFICATION & GOVERNANCE PLATFORM";
  workbook.created = new Date();

  // The workbook MUST contain exactly these five sheets in this order:
  //   1. Overall
  //   2. System Data
  //   3. Personal Data
  //   4. System Context
  //   5. PDPL
  buildOverallSheet(workbook, { system: safeSystem, summary: safeSummary, pdpl: safePdpl });
  buildDataSheet(workbook, "System Data", safeRows, safeColumns, false);
  buildDataSheet(workbook, "Personal Data", personalRows, safeColumns, true);
  buildSystemContextSheet(workbook, { system: safeSystem, contextPoints: safeContextPoints });
  buildPdplSheet(workbook, { system: safeSystem, rows: approvedPersonalRows, pdpl: safePdpl });

  return workbook;
}

function normalizeContextPoints(contextPoints) {
  if (!Array.isArray(contextPoints)) return [];
  return contextPoints
    .map((point) => ({
      tag: String(point?.tag || "").trim(),
      content: String(point?.content || "").trim(),
      createdBy: String(point?.createdBy || "").trim(),
      createdAt: String(point?.createdAt || "").trim(),
      updatedAt: String(point?.updatedAt || "").trim(),
    }))
    .filter((point) => point.tag || point.content);
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
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  sheet.columns = [
    { width: 18 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 18 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
  ];

  styleSheetBackground(sheet, 1, 48, 12, "F8FAFC");
  buildOverallHero(sheet, system);
  buildSystemInfoPanel(sheet, system);

  const cards = [
    { label: "Fully Classified", value: summary.classified, detail: `${summary.classifiedPercentage}%`, color: "2563EB" },
    { label: "Personal Data", value: summary.personal, detail: `${summary.personalPercentage}%`, color: "0F766E" },
    { label: "Pending Classification", value: summary.pending, detail: `${summary.pendingPercentage}%`, color: "D97706" },
    { label: "Review Queue", value: summary.reviewQueue, detail: "items", color: "7C3AED" },
    { label: "Low Confidence", value: summary.lowConfidence, detail: "items", color: "DC2626" },
    { label: "Tables With Personal Data", value: summary.tablesWithPersonalData, detail: "tables", color: "0891B2" },
  ];
  addKpiCard(sheet, 11, 1, cards[0]);
  addKpiCard(sheet, 11, 3, cards[1]);
  addKpiCard(sheet, 11, 5, cards[2]);
  addKpiCard(sheet, 11, 7, cards[3]);
  addKpiCard(sheet, 11, 9, cards[4]);
  addKpiCard(sheet, 11, 11, cards[5]);
  addProgressBand(sheet, 15, "Classification Progress", summary.classificationProgress, "2563EB");

  addHorizontalBarChart(sheet, 18, 1, 6, "Confidentiality Chart", [
    { label: "Confidential", value: summary.confidentiality.Confidential || 0, color: "F59E0B" },
    { label: "Secret", value: summary.confidentiality.Secret || 0, color: "DC2626" },
    { label: "Top Secret", value: summary.confidentiality["Top Secret"] || 0, color: "7C3AED" },
    { label: "Public", value: summary.confidentiality.Public || 0, color: "16A34A" },
    { label: "Pending", value: summary.confidentiality.Pending || 0, color: "64748B" },
  ]);

  addHorizontalBarChart(sheet, 18, 7, 6, "Personal Data Chart", [
    { label: "Has Personal Data", value: summary.personal || 0, color: "0F766E" },
    { label: "No Personal Data", value: Math.max(0, Number(summary.totalRecords || 0) - Number(summary.personal || 0)), color: "2563EB" },
    { label: "Needs Review", value: summary.reviewQueue || 0, color: "D97706" },
  ]);

  const personalTypeChartRows = personalDataTypeRows(summary)
    .slice(1, 7)
    .map(([label, count], index) => ({
      label,
      value: count,
      color: ["0F766E", "0891B2", "2563EB", "7C3AED", "D97706", "475569"][index] || "475569",
    }));
  addHorizontalBarChart(sheet, 27, 1, 12, "Personal Data Type Counts", personalTypeChartRows);

  const governanceStart = Math.max(37, 30 + Math.max(4, personalTypeChartRows.length));
  const governanceRows = [
    ["Indicator", "Status"],
    ["Data Subject Rights Coverage", pdpl?.dataSubjectRightsCoverage || "Needs Review"],
    ["Consent Tracking Status", pdpl?.consentTrackingStatus || "Needs Review"],
    ["Cross-Border Transfer Flags", pdpl?.crossBorderTransferFlags || "No Flags Recorded"],
    ["Policy Recommendations", summary.policyRecommendationCount],
    ["Classification Progress", `${summary.classificationProgress}%`],
  ];
  addSectionHeader(sheet, governanceStart, "Governance and PDPL Summary", 12);
  addSummaryTable(sheet, governanceStart + 1, governanceRows, { width: 12 });

  applyOuterBorders(sheet, `A1:L${governanceStart + governanceRows.length + 1}`);
}

function styleSheetBackground(sheet, startRow, endRow, endCol, color) {
  for (let row = startRow; row <= endRow; row += 1) {
    sheet.getRow(row).height = sheet.getRow(row).height || 20;
    for (let col = 1; col <= endCol; col += 1) {
      sheet.getCell(row, col).fill = fill(color);
    }
  }
}

function buildOverallHero(sheet, system) {
  sheet.mergeCells("A1:L1");
  sheet.mergeCells("A2:L2");
  sheet.mergeCells("A3:L3");
  sheet.getCell("A1").value = "DATA CLASSIFICATION & GOVERNANCE PLATFORM";
  sheet.getCell("A2").value = `${system.name} - Client Classification Report`;
  sheet.getCell("A3").value = `Generated ${new Date().toISOString().slice(0, 10)} | Excel Governance Dashboard`;
  [1, 2, 3].forEach((rowNumber) => {
    const row = sheet.getRow(rowNumber);
    row.height = rowNumber === 2 ? 32 : 24;
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = fill("0F172A");
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });
  });
  sheet.getCell("A1").font = { bold: true, size: 13, color: { argb: "FF93C5FD" } };
  sheet.getCell("A2").font = { bold: true, size: 20, color: { argb: "FFFFFFFF" } };
  sheet.getCell("A3").font = { italic: true, size: 10, color: { argb: "FFE2E8F0" } };
}

function buildSystemInfoPanel(sheet, system) {
  addSectionHeader(sheet, 5, "System Profile", 12);
  const rows = [
    ["Responsible Owner", system.owner || "Not recorded", "DBA", system.dba || "Not recorded", "Status", system.status],
    ["Owner Email", system.ownerEmail || "Not recorded", "System Group", system.systemGroup || "Not recorded", "Uploaded File", system.lastUploadFileName || "No file uploaded"],
    ["Source System", system.sourceSystemRef || "Not recorded", "Target System", system.targetSystemRef || "Not recorded", "Related Links", system.relatedSystemLinks || "Not recorded"],
  ];
  rows.forEach((items, offset) => {
    const row = sheet.getRow(6 + offset);
    row.height = 24;
    row.values = [null, ...items];
    for (let col = 1; col <= 12; col += 1) {
      const cell = row.getCell(col);
      cell.border = border("CBD5E1");
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.fill = fill(col % 2 === 1 ? "EFF6FF" : "FFFFFF");
      cell.font = col % 2 === 1 ? { bold: true, color: { argb: "FF334155" } } : { color: { argb: "FF111827" } };
    }
  });
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

function addSectionHeader(sheet, row, title, width = 6) {
  sheet.mergeCells(row, 1, row, width);
  const cell = sheet.getCell(row, 1);
  cell.value = title;
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.fill = fill("1E293B");
  cell.alignment = { vertical: "middle" };
  sheet.getRow(row).height = 24;
}

function addKpiCard(sheet, row, col, card) {
  if (!card) return;
  sheet.mergeCells(row, col, row, col + 1);
  sheet.mergeCells(row + 1, col, row + 2, col + 1);
  sheet.getCell(row, col).value = card.label;
  sheet.getCell(row, col).font = { bold: true, color: { argb: "FF475569" } };
  sheet.getCell(row + 1, col).value = Number(card.value || 0);
  sheet.getCell(row + 1, col).font = { bold: true, size: 18, color: { argb: `FF${card.color}` } };
  sheet.getCell(row + 2, col).value = card.detail;
  sheet.getCell(row + 2, col).font = { italic: true, size: 10, color: { argb: "FF64748B" } };
  [row, row + 1, row + 2].forEach((r) => {
    for (let c = col; c <= col + 1; c += 1) {
      sheet.getCell(r, c).fill = fill("FFFFFF");
      sheet.getCell(r, c).border = border("BFDBFE");
      sheet.getCell(r, c).alignment = { vertical: "middle", horizontal: "center" };
    }
  });
  sheet.getCell(row, col).fill = fill("F1F5F9");
  sheet.getCell(row, col + 1).fill = fill("F1F5F9");
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

function addSummaryTable(sheet, startRow, values, options = {}) {
  const width = Number(options.width || values[0]?.length || 2);
  values.forEach((items, offset) => {
    const row = sheet.getRow(startRow + offset);
    row.values = [null, ...items];
    for (let col = 1; col <= width; col += 1) {
      const cell = row.getCell(col);
      cell.border = border("CBD5E1");
      cell.alignment = { vertical: "middle" };
      if (offset === 0) {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = fill("1E40AF");
      } else {
        cell.fill = fill(offset % 2 === 0 ? "F8FAFC" : "FFFFFF");
      }
    }
  });
}

function addProgressBand(sheet, rowNumber, label, value, color) {
  const pct = Math.max(0, Math.min(100, Number(value || 0)));
  sheet.mergeCells(rowNumber, 1, rowNumber, 2);
  sheet.getCell(rowNumber, 1).value = label;
  sheet.getCell(rowNumber, 1).font = { bold: true, color: { argb: "FF334155" } };
  sheet.getCell(rowNumber, 1).alignment = { vertical: "middle" };
  const barStart = 3;
  const barEnd = 11;
  const filled = Math.round(((barEnd - barStart + 1) * pct) / 100);
  for (let col = barStart; col <= barEnd; col += 1) {
    const cell = sheet.getCell(rowNumber, col);
    cell.value = "";
    cell.fill = fill(col - barStart < filled ? color : "E2E8F0");
    cell.border = border("FFFFFF");
  }
  sheet.getCell(rowNumber, 12).value = `${pct}%`;
  sheet.getCell(rowNumber, 12).font = { bold: true, color: { argb: `FF${color}` } };
  sheet.getCell(rowNumber, 12).alignment = { vertical: "middle", horizontal: "right" };
  sheet.getRow(rowNumber).height = 22;
}

function addHorizontalBarChart(sheet, startRow, startCol, width, title, rows) {
  const endCol = startCol + width - 1;
  sheet.mergeCells(startRow, startCol, startRow, endCol);
  const titleCell = sheet.getCell(startRow, startCol);
  titleCell.value = title;
  titleCell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  titleCell.fill = fill("0F172A");
  titleCell.alignment = { vertical: "middle", horizontal: "center" };
  sheet.getRow(startRow).height = 24;

  const maxValue = Math.max(...rows.map((row) => Number(row.value || 0)), 1);
  const total = rows.reduce((sum, row) => sum + Number(row.value || 0), 0);
  const barStart = startCol + 3;
  const barEnd = endCol;
  const barCells = Math.max(1, barEnd - barStart + 1);

  rows.forEach((item, index) => {
    const rowNumber = startRow + index + 1;
    const value = Number(item.value || 0);
    const pct = total ? Math.round((value / total) * 100) : 0;
    const filled = Math.max(value > 0 ? 1 : 0, Math.round((value / maxValue) * barCells));
    const row = sheet.getRow(rowNumber);
    row.height = 22;

    sheet.getCell(rowNumber, startCol).value = item.label;
    sheet.getCell(rowNumber, startCol + 1).value = value;
    sheet.getCell(rowNumber, startCol + 2).value = `${pct}%`;
    [startCol, startCol + 1, startCol + 2].forEach((col) => {
      const cell = sheet.getCell(rowNumber, col);
      cell.fill = fill(index % 2 === 0 ? "FFFFFF" : "F8FAFC");
      cell.border = border("E2E8F0");
      cell.alignment = { vertical: "middle" };
      if (col === startCol) cell.font = { bold: true, color: { argb: "FF334155" } };
    });

    for (let col = barStart; col <= barEnd; col += 1) {
      const cell = sheet.getCell(rowNumber, col);
      cell.value = "";
      cell.fill = fill(col - barStart < filled ? item.color : "E2E8F0");
      cell.border = border("FFFFFF");
    }
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

// NOTE: The "Links & General Info" sheet is intentionally excluded from this
// export. The export produces exactly one .xlsx workbook with five sheets:
// Overall, System Data, Personal Data, System Context, PDPL.

function buildSystemContextSheet(workbook, { system, contextPoints }) {
  const sheet = workbook.addWorksheet("System Context", {
    views: [{ state: "frozen", ySplit: 6, showGridLines: false }],
  });
  sheet.columns = [
    { width: 28 },
    { width: 60 },
    { width: 24 },
    { width: 24 },
  ];

  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = `${system.name} - System Context`;
  sheet.getCell("A1").font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
  sheet.getCell("A1").fill = fill("0F766E");
  sheet.getCell("A1").alignment = { vertical: "middle" };
  sheet.getRow(1).height = 30;

  sheet.mergeCells("A3:D3");
  sheet.getCell("A3").value =
    "Reviewer-supplied context describing the system's purpose and data scope. The classifier uses these notes as supporting evidence, but they must not, on their own, push every field into Personal Data or Secret.";
  sheet.getCell("A3").alignment = { vertical: "middle", wrapText: true };
  sheet.getCell("A3").font = { color: { argb: "FF374151" }, italic: true };
  sheet.getRow(3).height = 36;

  const headerRowNumber = 6;
  const headers = ["Context Tag", "Content", "Recorded By", "Last Updated"];
  sheet.getRow(headerRowNumber).values = headers;
  sheet.getRow(headerRowNumber).height = 24;
  sheet.getRow(headerRowNumber).eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = fill("1D4ED8");
    cell.border = border("93C5FD");
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });

  if (contextPoints.length) {
    contextPoints.forEach((point) => {
      const added = sheet.addRow([
        point.tag || "Context",
        point.content || "",
        point.createdBy || "",
        point.updatedAt || point.createdAt || "",
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
      added.getCell(1).font = { bold: true, color: { argb: "FF1D4ED8" } };
    });
  } else {
    const added = sheet.addRow([
      "No context recorded",
      "No reviewer-supplied system context has been added for this system yet.",
      "",
      "",
    ]);
    added.eachCell((cell) => {
      cell.border = border("E2E8F0");
      cell.alignment = { vertical: "top", wrapText: true };
      cell.font = { italic: true, color: { argb: "FF6B7280" } };
    });
  }

  sheet.autoFilter = {
    from: `A${headerRowNumber}`,
    to: `D${headerRowNumber}`,
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
