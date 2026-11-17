import type { CellMatrix, CellPrimitive } from "../types/excel";
import {
  snapshotRangeBeforeMutation,
  stripSheetPrefix,
  columnIndexToLetter,
  getStartColumnIndexFromAddress
} from "./excelCore";

export interface MisActionResult {
  success: boolean;
  details: string;
  reportSheetName?: string;
  matchedCount?: number;
  unmatchedACount?: number;
  unmatchedBCount?: number;
}

export interface ColumnProfile {
  index: number;
  letter: string;
  header: string;
  type: "numeric" | "categorical" | "date" | "text";
  nonEmptyCount: number;
  uniqueCount: number;
}

/**
 * Profiles columns in a range to detect Numeric vs. Categorical vs. Text.
 */
export function profileTableColumns(
  values: CellMatrix,
  startColIndex = 0,
  hasHeaderRow = true
): ColumnProfile[] {
  if (values.length === 0 || values[0].length === 0) return [];

  const rowCount = values.length;
  const colCount = values[0].length;
  const startRow = hasHeaderRow ? 1 : 0;
  const dataRowCount = rowCount - startRow;

  const profiles: ColumnProfile[] = [];

  for (let c = 0; c < colCount; c++) {
    const colIdx = startColIndex + c;
    const letter = columnIndexToLetter(colIdx);
    const rawHeader = hasHeaderRow ? String(values[0][c] ?? "").trim() : "";
    const header = rawHeader !== "" ? rawHeader : `Column ${letter}`;

    let numericCount = 0;
    let nonEmptyCount = 0;
    const uniqueValues = new Set<string>();

    for (let r = startRow; r < rowCount; r++) {
      const val = values[r][c];
      if (val !== "" && val !== null && val !== undefined) {
        nonEmptyCount++;
        uniqueValues.add(String(val).trim().toLowerCase());
        if (typeof val === "number" || (!isNaN(Number(val)) && String(val).trim() !== "")) {
          numericCount++;
        }
      }
    }

    let type: ColumnProfile["type"] = "text";
    if (nonEmptyCount > 0 && numericCount / nonEmptyCount >= 0.7) {
      type = "numeric";
    } else if (uniqueValues.size >= 1 && uniqueValues.size <= Math.min(30, dataRowCount * 0.8)) {
      type = "categorical";
    }

    profiles.push({
      index: c,
      letter,
      header,
      type,
      nonEmptyCount,
      uniqueCount: uniqueValues.size
    });
  }

  return profiles;
}

/**
 * 1-Click Executive Summary Generator
 * Creates an executive dashboard tab with Top KPIs and Category Breakdown tables.
 */
export async function generateExecutiveSummary(options: {
  sourceSheetName: string;
  rangeAddress: string;
  hasHeaderRow: boolean;
  selectedCategoryColOffset?: number;
  selectedNumericColOffset?: number;
}): Promise<MisActionResult> {
  const {
    sourceSheetName,
    rangeAddress,
    hasHeaderRow,
    selectedCategoryColOffset,
    selectedNumericColOffset
  } = options;

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sourceSheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "rowCount", "columnCount", "address"]);
    await context.sync();

    const values = range.values as CellMatrix;
    const cleanAddr = stripSheetPrefix(range.address);
    const startColIdx = getStartColumnIndexFromAddress(cleanAddr);
    const profiles = profileTableColumns(values, startColIdx, hasHeaderRow);

    const startRow = hasHeaderRow ? 1 : 0;
    const totalRecords = Math.max(0, values.length - startRow);
    const totalCells = totalRecords * values[0].length;
    let filledCells = 0;
    for (let r = startRow; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const val = values[r][c];
        if (val !== "" && val !== null && val !== undefined) filledCells++;
      }
    }
    const completenessPct = totalCells > 0 ? (filledCells / totalCells) * 100 : 100;

    // Pick best categorical & numeric columns
    const numericCols = profiles.filter((p) => p.type === "numeric");
    const categoryCols = profiles.filter((p) => p.type === "categorical");

    const catCol =
      selectedCategoryColOffset !== undefined && profiles[selectedCategoryColOffset]
        ? profiles[selectedCategoryColOffset]
        : categoryCols[0] || profiles[0];

    const numCol =
      selectedNumericColOffset !== undefined && profiles[selectedNumericColOffset]
        ? profiles[selectedNumericColOffset]
        : numericCols[0];

    // Compute basic numeric stats
    let numSum = 0;
    let numMin = Infinity;
    let numMax = -Infinity;
    const numValues: number[] = [];

    if (numCol) {
      for (let r = startRow; r < values.length; r++) {
        const val = Number(values[r][numCol.index]);
        if (!isNaN(val) && values[r][numCol.index] !== "") {
          numSum += val;
          if (val < numMin) numMin = val;
          if (val > numMax) numMax = val;
          numValues.push(val);
        }
      }
    }
    const numMean = numValues.length > 0 ? numSum / numValues.length : 0;
    numValues.sort((a, b) => a - b);
    const numMedian =
      numValues.length > 0
        ? numValues.length % 2 === 0
          ? (numValues[numValues.length / 2 - 1] + numValues[numValues.length / 2]) / 2
          : numValues[Math.floor(numValues.length / 2)]
        : 0;

    // Unique category breakdown
    const categoryMap = new Map<string, { count: number; sum: number }>();
    if (catCol) {
      for (let r = startRow; r < values.length; r++) {
        const catKey = String(values[r][catCol.index] ?? "Unspecified").trim() || "Unspecified";
        const metricVal = numCol ? Number(values[r][numCol.index]) || 0 : 0;
        const current = categoryMap.get(catKey) || { count: 0, sum: 0 };
        current.count++;
        current.sum += metricVal;
        categoryMap.set(catKey, current);
      }
    }

    // Create New Summary Sheet with unique timestamp
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, "0");
    const timestampStr = `${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;
    const summarySheetName = `MIS_Summary_${timestampStr}`;

    const summarySheet = context.workbook.worksheets.add(summarySheetName);

    // ── Build Sheet Content ──
    const rows: Array<CellPrimitive[]> = [
      ["Executive Data Summary Dashboard", "", "", ""],
      [`Generated for sheet: "${sourceSheetName}"`, "", "", ""],
      ["", "", "", ""],
      ["OVERVIEW KPIS", "VALUE", "METRIC", "VALUE"],
      ["Total Records", totalRecords, "Total Columns", profiles.length],
      ["Data Completeness", `${completenessPct.toFixed(1)}%`, "Categorical Columns", categoryCols.length],
      ["", "", "", ""]
    ];

    if (numCol && numValues.length > 0) {
      rows.push(
        [`NUMERIC SUMMARY: ${numCol.header}`, "VALUE", "ADDITIONAL METRICS", "VALUE"],
        ["Total Sum", Math.round(numSum * 100) / 100, "Minimum", numMin === Infinity ? 0 : numMin],
        ["Average (Mean)", Math.round(numMean * 100) / 100, "Maximum", numMax === -Infinity ? 0 : numMax],
        ["Median", Math.round(numMedian * 100) / 100, "Valid Data Points", numValues.length],
        ["", "", "", ""]
      );
    }

    if (categoryMap.size > 0 && catCol) {
      rows.push([`BREAKDOWN BY: ${catCol.header}`, "RECORDS", numCol ? `SUM OF ${numCol.header}` : "% OF TOTAL", "% OF TOTAL"]);
      const sortedCats = Array.from(categoryMap.entries())
        .sort((a, b) => b[1].count - a[1].count)
        .slice(0, 30); // Top 30 categories

      for (const [catName, stats] of sortedCats) {
        const pct = totalRecords > 0 ? `${((stats.count / totalRecords) * 100).toFixed(1)}%` : "0%";
        if (numCol) {
          rows.push([catName, stats.count, Math.round(stats.sum * 100) / 100, pct]);
        } else {
          rows.push([catName, stats.count, pct, ""]);
        }
      }
    }

    // Write all rows
    const summaryRange = summarySheet.getRange(`A1:D${rows.length}`);
    summaryRange.values = rows;

    // ── Apply Styles & Formatting ──
    const titleCell = summarySheet.getRange("A1");
    titleCell.format.font.bold = true;
    titleCell.format.font.size = 14;
    titleCell.format.font.color = "#065F46"; // emerald-800

    const subTitle = summarySheet.getRange("A2");
    subTitle.format.font.size = 9;
    subTitle.format.font.color = "#64748B";

    // Style Header Rows (KPI Header & Breakdown Header)
    const kpiHeader = summarySheet.getRange("A4:D4");
    kpiHeader.format.font.bold = true;
    kpiHeader.format.fill.color = "#ECFDF5"; // emerald-50
    kpiHeader.format.font.color = "#065F46";

    // Autofit columns for professional readability
    summarySheet.getUsedRange().format.autofitColumns();
    summarySheet.activate();

    await context.sync();

    return {
      success: true,
      reportSheetName: summarySheetName,
      details: `Generated executive dashboard tab "${summarySheetName}" with overview KPIs and breakdown table.`
    };
  });
}

/**
 * Inserts adjacent Variance and Variance % (Achievement %) columns.
 */
export async function insertVarianceColumns(options: {
  sheetName: string;
  rangeAddress: string;
  targetColOffset: number;
  actualColOffset: number;
  hasHeaderRow: boolean;
}): Promise<MisActionResult> {
  const { sheetName, rangeAddress, targetColOffset, actualColOffset, hasHeaderRow } = options;

  await snapshotRangeBeforeMutation("Insert Variance Columns", sheetName, rangeAddress);

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["address", "rowCount", "columnCount"]);
    await context.sync();

    const cleanAddr = stripSheetPrefix(range.address);
    const startCell = cleanAddr.split(":")[0];
    const startColLetter = startCell.replace(/\d+/g, "");
    const startRow = parseInt(startCell.replace(/\D+/g, ""), 10);
    const rowCount = range.rowCount;
    const endRow = startRow + rowCount - 1;

    let baseColIdx = 0;
    for (let i = 0; i < startColLetter.length; i++) {
      baseColIdx = baseColIdx * 26 + (startColLetter.charCodeAt(i) - 64);
    }
    baseColIdx = Math.max(0, baseColIdx - 1);

    const targetLetter = columnIndexToLetter(baseColIdx + targetColOffset);
    const actualLetter = columnIndexToLetter(baseColIdx + actualColOffset);

    // Insert 2 new columns to the right of the table
    const lastColRange = range.getColumn(range.columnCount - 1);
    lastColRange.load(["address"]);
    await context.sync();

    const lastColLetter = stripSheetPrefix(lastColRange.address).split(":")[0].replace(/\d+/g, "");
    let lastColNumber = 0;
    for (let i = 0; i < lastColLetter.length; i++) {
      lastColNumber = lastColNumber * 26 + (lastColLetter.charCodeAt(i) - 64);
    }

    const varColLetter = columnIndexToLetter(lastColNumber); // +1 column
    const varPctColLetter = columnIndexToLetter(lastColNumber + 1); // +2 columns

    const varRange = sheet.getRange(`${varColLetter}${startRow}:${varColLetter}${endRow}`);
    const varPctRange = sheet.getRange(`${varPctColLetter}${startRow}:${varPctColLetter}${endRow}`);

    const varFormulas: string[][] = [];
    const varPctFormulas: string[][] = [];

    const dataStartRow = hasHeaderRow ? startRow + 1 : startRow;

    if (hasHeaderRow) {
      varFormulas.push(["Variance"]);
      varPctFormulas.push(["Variance %"]);
    }

    for (let r = dataStartRow; r <= endRow; r++) {
      varFormulas.push([`=${actualLetter}${r}-${targetLetter}${r}`]);
      varPctFormulas.push([`=IFERROR((${actualLetter}${r}-${targetLetter}${r})/${targetLetter}${r}, 0)`]);
    }

    varRange.formulas = varFormulas;
    varPctRange.formulas = varPctFormulas;

    // Apply percentage format on Variance % column
    const formatRange = sheet.getRange(`${varPctColLetter}${dataStartRow}:${varPctColLetter}${endRow}`);
    formatRange.numberFormat = [["0.0%"]];

    if (hasHeaderRow) {
      const headerRange = sheet.getRange(`${varColLetter}${startRow}:${varPctColLetter}${startRow}`);
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#F0FDF4"; // emerald-50
      headerRange.format.font.color = "#166534";
    }

    sheet.getUsedRange().format.autofitColumns();
    await context.sync();

    return {
      success: true,
      details: `Added "Variance" (${varColLetter}) and "Variance %" (${varPctColLetter}) columns across ${rowCount - (hasHeaderRow ? 1 : 0)} data rows.`
    };
  });
}

/**
 * Two-Sheet Reconciler
 * Compares primary keys across two sheets, flags discrepancies, and generates a reconciliation report.
 */
export async function reconcileTwoSheets(options: {
  sheetAName: string;
  keyColALetter: string;
  sheetBName: string;
  keyColBLetter: string;
  createReportSheet: boolean;
}): Promise<MisActionResult> {
  const { sheetAName, keyColALetter, sheetBName, keyColBLetter, createReportSheet } = options;

  return Excel.run(async (context) => {
    const sheetA = context.workbook.worksheets.getItem(sheetAName);
    const sheetB = context.workbook.worksheets.getItem(sheetBName);

    const usedA = sheetA.getUsedRangeOrNullObject();
    const usedB = sheetB.getUsedRangeOrNullObject();
    usedA.load(["address", "rowCount"]);
    usedB.load(["address", "rowCount"]);
    await context.sync();

    if (usedA.isNullObject || usedB.isNullObject) {
      return { success: false, details: "One or both selected sheets appear to be empty." };
    }

    const colARange = sheetA.getRange(`${keyColALetter}1:${keyColALetter}${usedA.rowCount}`);
    const colBRange = sheetB.getRange(`${keyColBLetter}1:${keyColBLetter}${usedB.rowCount}`);
    colARange.load(["values"]);
    colBRange.load(["values"]);
    await context.sync();

    const valsA = colARange.values as CellMatrix;
    const valsB = colBRange.values as CellMatrix;

    // Build key maps
    const mapA = new Map<string, number>(); // key -> row (1-indexed)
    for (let r = 1; r < valsA.length; r++) {
      const k = String(valsA[r][0] ?? "").trim();
      if (k !== "") mapA.set(k.toLowerCase(), r + 1);
    }

    const mapB = new Map<string, number>();
    for (let r = 1; r < valsB.length; r++) {
      const k = String(valsB[r][0] ?? "").trim();
      if (k !== "") mapB.set(k.toLowerCase(), r + 1);
    }

    // Compare
    const missingInB: Array<{ key: string; rowA: number }> = [];
    const missingInA: Array<{ key: string; rowB: number }> = [];
    let matchedCount = 0;

    for (const [keyA, rowA] of mapA) {
      if (mapB.has(keyA)) {
        matchedCount++;
      } else {
        missingInB.push({ key: keyA, rowA });
      }
    }

    for (const [keyB, rowB] of mapB) {
      if (!mapA.has(keyB)) {
        missingInA.push({ key: keyB, rowB });
      }
    }

    // Highlight unmatched keys in Sheet A (soft amber)
    for (const item of missingInB.slice(0, 500)) {
      sheetA.getRange(`${keyColALetter}${item.rowA}`).format.fill.color = "#FEF3C7"; // amber-100
    }

    // Highlight unmatched keys in Sheet B (soft amber)
    for (const item of missingInA.slice(0, 500)) {
      sheetB.getRange(`${keyColBLetter}${item.rowB}`).format.fill.color = "#FEF3C7";
    }

    let reportSheetName: string | undefined;

    if (createReportSheet) {
      const now = new Date();
      const pad = (n: number) => String(n).padStart(2, "0");
      reportSheetName = `Recon_${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;
      const reportSheet = context.workbook.worksheets.add(reportSheetName);

      const reportRows: Array<CellPrimitive[]> = [
        ["Two-Sheet Reconciliation Report", "", "", ""],
        [`Compared: "${sheetAName}" (${keyColALetter}) vs. "${sheetBName}" (${keyColBLetter})`, "", "", ""],
        ["", "", "", ""],
        ["RECONCILIATION SUMMARY", "COUNT", "METRIC", "VALUE"],
        ["Total Keys in Sheet A", mapA.size, "Total Keys in Sheet B", mapB.size],
        ["Matched in Both Sheets", matchedCount, "Match Rate", `${mapA.size > 0 ? ((matchedCount / mapA.size) * 100).toFixed(1) : 0}%`],
        ["Missing in Sheet B (Only in A)", missingInB.length, "Missing in Sheet A (Only in B)", missingInA.length],
        ["", "", "", ""],
        ["RECORDS ONLY IN SHEET A", "ORIGINAL ROW", "RECORDS ONLY IN SHEET B", "ORIGINAL ROW"]
      ];

      const maxDiscrepancyRows = Math.max(missingInB.length, missingInA.length);
      for (let i = 0; i < Math.min(maxDiscrepancyRows, 500); i++) {
        const itemA = missingInB[i];
        const itemB = missingInA[i];
        reportRows.push([
          itemA ? itemA.key : "",
          itemA ? itemA.rowA : "",
          itemB ? itemB.key : "",
          itemB ? itemB.rowB : ""
        ]);
      }

      const reportRange = reportSheet.getRange(`A1:D${reportRows.length}`);
      reportRange.values = reportRows;

      const title = reportSheet.getRange("A1");
      title.format.font.bold = true;
      title.format.font.size = 14;
      title.format.font.color = "#1E3A8A"; // blue-900

      const summaryHdr = reportSheet.getRange("A4:D4");
      summaryHdr.format.font.bold = true;
      summaryHdr.format.fill.color = "#EFF6FF"; // blue-50
      summaryHdr.format.font.color = "#1E40AF";

      const diffHdr = reportSheet.getRange("A9:D9");
      diffHdr.format.font.bold = true;
      diffHdr.format.fill.color = "#FEF3C7"; // amber-100
      diffHdr.format.font.color = "#92400E";

      reportSheet.getUsedRange().format.autofitColumns();
      reportSheet.activate();
    }

    await context.sync();

    return {
      success: true,
      reportSheetName,
      matchedCount,
      unmatchedACount: missingInB.length,
      unmatchedBCount: missingInA.length,
      details: `Reconciliation complete: ${matchedCount} matched, ${missingInB.length} only in ${sheetAName}, ${missingInA.length} only in ${sheetBName}.${reportSheetName ? ` Created report tab "${reportSheetName}".` : ""}`
    };
  });
}
