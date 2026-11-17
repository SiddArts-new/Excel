import type { CellMatrix } from "../types/excel";
import { snapshotRangeBeforeMutation, stripSheetPrefix, columnIndexToLetter } from "./excelCore";

export interface SheetInfo {
  id: string;
  name: string;
}

export interface SheetHeaderInfo {
  sheetName: string;
  columns: Array<{ index: number; letter: string; header: string }>;
}

export interface FormulaActionResult {
  cellsAffected: number;
  details: string;
}

export type ExtractionType = "email" | "numbers" | "before_delim" | "after_delim";

/**
 * Lists all visible worksheets in the active workbook.
 */
export async function getWorkbookSheetNames(): Promise<SheetInfo[]> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load(["items/id", "items/name", "items/visibility"]);
    await context.sync();

    return sheets.items
      .filter((s) => s.visibility === Excel.SheetVisibility.visible)
      .map((s) => ({ id: s.id, name: s.name }));
  });
}

/**
 * Reads headers (row 1 or used range first row) for a specific sheet.
 */
export async function getHeadersForSheet(sheetName: string): Promise<SheetHeaderInfo> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load(["address", "columnCount", "values"]);
    await context.sync();

    if (usedRange.isNullObject || usedRange.columnCount === 0) {
      return { sheetName, columns: [] };
    }

    const firstRow = (usedRange.values as CellMatrix)[0] ?? [];
    // Determine start column offset from usedRange address
    const cleanAddress = stripSheetPrefix(usedRange.address);
    const startColLetter = cleanAddress.split(":")[0]?.replace(/\d+/g, "") || "A";
    let startColIndex = 0;
    for (let i = 0; i < startColLetter.length; i++) {
      startColIndex = startColIndex * 26 + (startColLetter.charCodeAt(i) - 64);
    }
    startColIndex = Math.max(0, startColIndex - 1);

    const columns = firstRow.map((headerVal, i) => {
      const colIdx = startColIndex + i;
      const letter = columnIndexToLetter(colIdx);
      const text = headerVal !== null && headerVal !== undefined && headerVal !== ""
        ? String(headerVal).trim()
        : `Column ${letter}`;
      return { index: colIdx, letter, header: text };
    });

    return { sheetName, columns };
  });
}

/**
 * Builds and applies an XLOOKUP (or legacy INDEX/MATCH) formula column.
 */
export async function applyVisualLookup(options: {
  activeSheetName: string;
  lookupColLetter: string;
  targetSheetName: string;
  targetLookupColLetter: string;
  targetReturnColLetter: string;
  fallbackValue: string;
  useLegacyIndexMatch: boolean;
  outputColLetter: string;
  startRow: number;
  rowCount: number;
}): Promise<FormulaActionResult> {
  const {
    activeSheetName,
    lookupColLetter,
    targetSheetName,
    targetLookupColLetter,
    targetReturnColLetter,
    fallbackValue,
    useLegacyIndexMatch,
    outputColLetter,
    startRow,
    rowCount
  } = options;

  if (rowCount <= 0) {
    return { cellsAffected: 0, details: "No target rows to apply lookup." };
  }

  const endRow = startRow + rowCount - 1;
  const targetRangeAddress = `${outputColLetter}${startRow}:${outputColLetter}${endRow}`;

  await snapshotRangeBeforeMutation(
    useLegacyIndexMatch ? "Apply INDEX/MATCH" : "Apply XLOOKUP",
    activeSheetName,
    targetRangeAddress
  );

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(activeSheetName);
    const range = sheet.getRange(targetRangeAddress);

    // Escape quotes in fallback
    const safeFallback = isNaN(Number(fallbackValue))
      ? `"${fallbackValue.replace(/"/g, '""')}"`
      : fallbackValue;

    // Sanitize sheet name for Excel formulas (single quote if contains space or special char)
    const escapedTargetSheet = targetSheetName.includes(" ") || targetSheetName.includes("-")
      ? `'${targetSheetName.replace(/'/g, "''")}'`
      : targetSheetName;

    const formulas: string[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      let formula: string;
      if (useLegacyIndexMatch) {
        formula = `=IFERROR(INDEX(${escapedTargetSheet}!${targetReturnColLetter}:${targetReturnColLetter}, MATCH(${lookupColLetter}${r}, ${escapedTargetSheet}!${targetLookupColLetter}:${targetLookupColLetter}, 0)), ${safeFallback})`;
      } else {
        formula = `=XLOOKUP(${lookupColLetter}${r}, ${escapedTargetSheet}!${targetLookupColLetter}:${targetLookupColLetter}, ${escapedTargetSheet}!${targetReturnColLetter}:${targetReturnColLetter}, ${safeFallback})`;
      }
      formulas.push([formula]);
    }

    range.formulas = formulas;
    await context.sync();

    const modeName = useLegacyIndexMatch ? "INDEX/MATCH" : "XLOOKUP";
    return {
      cellsAffected: rowCount,
      details: `Generated ${modeName} across ${rowCount} rows in Column ${outputColLetter}.`
    };
  });
}

/**
 * Wraps unhandled formulas in IFERROR(<formula>, <fallback>).
 */
export async function wrapSelectedRangeInIfError(
  sheetName: string,
  rangeAddress: string,
  fallbackVal: string = '""'
): Promise<FormulaActionResult> {
  await snapshotRangeBeforeMutation("Wrap in IFERROR", sheetName, rangeAddress);

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["formulas", "values", "rowCount", "columnCount"]);
    await context.sync();

    const formulas = range.formulas as CellMatrix;
    let wrappedCount = 0;

    const newFormulas = formulas.map((row) =>
      row.map((cell) => {
        if (typeof cell === "string" && cell.startsWith("=")) {
          // Check if already wrapped in IFERROR
          const trimmed = cell.trim();
          if (/^=IFERROR\s*\(/i.test(trimmed)) {
            return cell;
          }
          const innerFormula = trimmed.slice(1).trim();
          wrappedCount++;
          return `=IFERROR(${innerFormula}, ${fallbackVal})`;
        }
        return cell;
      })
    );

    if (wrappedCount > 0) {
      range.formulas = newFormulas;
      await context.sync();
    }

    return {
      cellsAffected: wrappedCount,
      details: wrappedCount > 0
        ? `Protected ${wrappedCount} formula cell${wrappedCount === 1 ? "" : "s"} with IFERROR.`
        : "No unprotected formula cells found in the selection."
    };
  });
}

/**
 * Converts all formulas in the selected range to evaluated static values.
 */
export async function convertRangeFormulasToValues(
  sheetName: string,
  rangeAddress: string
): Promise<FormulaActionResult> {
  await snapshotRangeBeforeMutation("Convert Formulas to Values", sheetName, rangeAddress);

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["formulas", "values", "rowCount", "columnCount"]);
    await context.sync();

    const formulas = range.formulas as CellMatrix;
    const values = range.values as CellMatrix;

    let formulaCount = 0;
    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        if (typeof formulas[r][c] === "string" && (formulas[r][c] as string).startsWith("=")) {
          formulaCount++;
        }
      }
    }

    if (formulaCount === 0) {
      return {
        cellsAffected: 0,
        details: "No formula cells found in the selected range."
      };
    }

    // Assigning values back directly freezes evaluated results
    range.values = values;
    await context.sync();

    return {
      cellsAffected: formulaCount,
      details: `Converted ${formulaCount} formula${formulaCount === 1 ? "" : "s"} to static values.`
    };
  });
}

/**
 * Extracts structured substrings (Emails, Numbers, or Delimiter tokens)
 * into a new inserted column or destination.
 */
export async function extractPatternToNewColumn(options: {
  sheetName: string;
  rangeAddress: string;
  sourceColOffset: number; // 0-indexed within range
  patternType: ExtractionType;
  delimiter?: string;
  hasHeaderRow: boolean;
}): Promise<FormulaActionResult> {
  const { sheetName, rangeAddress, sourceColOffset, patternType, delimiter, hasHeaderRow } = options;

  await snapshotRangeBeforeMutation("Pattern Extraction", sheetName, rangeAddress);

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "rowCount", "columnCount", "address"]);
    await context.sync();

    const values = range.values as CellMatrix;
    const rowCount = range.rowCount;

    // Pattern extractors
    const extractEmail = (text: string): string => {
      const match = text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
      return match ? match[0] : "";
    };

    const extractNumbers = (text: string): string => {
      const match = text.match(/[-+]?\$?\d+(?:,\d{3})*(?:\.\d+)?/);
      return match ? match[0].replace(/[$,]/g, "") : "";
    };

    const extractBefore = (text: string, delim: string): string => {
      const idx = text.indexOf(delim);
      return idx >= 0 ? text.slice(0, idx).trim() : text.trim();
    };

    const extractAfter = (text: string, delim: string): string => {
      const idx = text.indexOf(delim);
      return idx >= 0 ? text.slice(idx + delim.length).trim() : "";
    };

    // Calculate source column cell coordinates
    const sourceColRange = range.getColumn(sourceColOffset);
    sourceColRange.load(["address"]);
    await context.sync();

    // Insert new column to the right
    sourceColRange.getEntireColumn().insert(Excel.InsertShiftDirection.right);
    await context.sync();

    // Form header label
    const patternLabels: Record<ExtractionType, string> = {
      email: "Extracted Email",
      numbers: "Extracted Number",
      before_delim: `Before '${delimiter || "-"}'`,
      after_delim: `After '${delimiter || "-"}'`
    };

    const extractedData: CellMatrix = [];
    let matchCount = 0;

    for (let r = 0; r < rowCount; r++) {
      if (hasHeaderRow && r === 0) {
        extractedData.push([patternLabels[patternType]]);
        continue;
      }

      const rawVal = String(values[r]?.[sourceColOffset] ?? "");
      let result = "";

      switch (patternType) {
        case "email":
          result = extractEmail(rawVal);
          break;
        case "numbers":
          result = extractNumbers(rawVal);
          break;
        case "before_delim":
          result = extractBefore(rawVal, delimiter || "-");
          break;
        case "after_delim":
          result = extractAfter(rawVal, delimiter || "-");
          break;
      }

      if (result !== "") matchCount++;
      extractedData.push([result]);
    }

    // Determine target range address for inserted column
    const sourceShort = stripSheetPrefix(sourceColRange.address);
    const sourceStartCell = sourceShort.split(":")[0];
    const sourceColLetter = sourceStartCell.replace(/\d+/g, "");
    const sourceStartRow = parseInt(sourceStartCell.replace(/\D+/g, ""), 10);
    const sourceEndRow = sourceStartRow + rowCount - 1;

    // Target is column immediately to the right
    let colNumber = 0;
    for (let i = 0; i < sourceColLetter.length; i++) {
      colNumber = colNumber * 26 + (sourceColLetter.charCodeAt(i) - 64);
    }
    const targetColLetter = columnIndexToLetter(colNumber); // colNumber is already +1 due to 1-based char math
    const targetRange = sheet.getRange(`${targetColLetter}${sourceStartRow}:${targetColLetter}${sourceEndRow}`);
    targetRange.values = extractedData;
    await context.sync();

    return {
      cellsAffected: matchCount,
      details: `Extracted ${matchCount} match${matchCount === 1 ? "" : "es"} into new Column ${targetColLetter}.`
    };
  });
}
