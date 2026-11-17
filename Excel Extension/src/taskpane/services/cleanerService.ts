import type { CellMatrix, CellPrimitive, StringMatrix } from "../types/excel";
import { snapshotRangeBeforeMutation, stripSheetPrefix } from "./excelCore";

// ────────────────────── Types ──────────────────────

export interface CleanResult {
  totalCells: number;
  modifiedCells: number;
  details: string;
}

export type CaseFormat = "title" | "upper" | "lower" | "snake";

export interface DuplicatePreview {
  totalRows: number;
  uniqueRows: number;
  duplicateRows: number;
}

// ────────────────────── Constants ──────────────────────

export const DEFAULT_SENTINELS = [
  "-1",
  "N/A",
  "n/a",
  "NULL",
  "null",
  "None",
  "#N/A",
  "-"
];

const HIGHLIGHT_COLORS = {
  duplicate: "#FEF9C3",
  blank: "#DBEAFE",
  outlier: "#FFE4E6"
} as const;

// ────────────────────── Helpers ──────────────────────

function isFormula(val: unknown): boolean {
  return typeof val === "string" && val.startsWith("=");
}

/**
 * Builds a write-back matrix: formula cells keep their original formula string;
 * non-formula cells use the cleaned value. Written to `range.formulas` so that
 * formulas are preserved and plain values are set correctly.
 */
function buildWriteMatrix(
  originalFormulas: CellMatrix,
  cleanedValues: CellMatrix
): CellMatrix {
  return originalFormulas.map((row, r) =>
    row.map((formula, c) => {
      if (isFormula(formula)) return formula;
      return cleanedValues[r]?.[c] ?? "";
    })
  );
}

function countModified(
  original: CellMatrix,
  cleaned: CellMatrix,
  formulas: CellMatrix
): { total: number; modified: number } {
  let total = 0;
  let modified = 0;
  for (let r = 0; r < original.length; r++) {
    for (let c = 0; c < (original[r]?.length ?? 0); c++) {
      total++;
      if (!isFormula(formulas[r]?.[c]) && original[r][c] !== cleaned[r][c]) {
        modified++;
      }
    }
  }
  return { total, modified };
}

// ────────────────────── 1. Whitespace & Control Char Cleaner ──────────────────────

export async function cleanWhitespace(
  sheetName: string,
  rangeAddress: string
): Promise<CleanResult> {
  await snapshotRangeBeforeMutation("Clean Whitespace", sheetName, rangeAddress);

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "formulas"]);
    await context.sync();

    const values = range.values as CellMatrix;
    const formulas = range.formulas as CellMatrix;

    const cleaned: CellMatrix = values.map((row, r) =>
      row.map((cell, c) => {
        if (isFormula(formulas[r][c])) return cell;
        if (typeof cell !== "string") return cell;
        return cell
          .replace(/[\u0000-\u001F\u007F]/g, "")
          .replace(/\u00A0/g, " ")
          .trim()
          .replace(/\s{2,}/g, " ");
      })
    );

    const { total, modified } = countModified(values, cleaned, formulas);
    range.formulas = buildWriteMatrix(formulas, cleaned);
    await context.sync();

    return {
      totalCells: total,
      modifiedCells: modified,
      details: `Trimmed whitespace & stripped control characters in ${modified} of ${total} cells.`
    };
  });
}

// ────────────────────── 2. Sentinel / Null Normalizer ──────────────────────

export async function normalizeSentinels(
  sheetName: string,
  rangeAddress: string,
  placeholders: string[] = DEFAULT_SENTINELS
): Promise<CleanResult> {
  await snapshotRangeBeforeMutation(
    "Normalize Sentinels",
    sheetName,
    rangeAddress
  );

  const placeholderSet = new Set(
    placeholders.map((p) => p.trim().toLowerCase())
  );

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "formulas"]);
    await context.sync();

    const values = range.values as CellMatrix;
    const formulas = range.formulas as CellMatrix;

    const cleaned: CellMatrix = values.map((row, r) =>
      row.map((cell, c) => {
        if (isFormula(formulas[r][c])) return cell;
        if (typeof cell === "string") {
          if (placeholderSet.has(cell.trim().toLowerCase())) return "";
        } else if (typeof cell === "number") {
          if (placeholderSet.has(String(cell))) return "";
        }
        return cell;
      })
    );

    const { total, modified } = countModified(values, cleaned, formulas);
    range.formulas = buildWriteMatrix(formulas, cleaned);
    await context.sync();

    return {
      totalCells: total,
      modifiedCells: modified,
      details: `Replaced ${modified} sentinel/placeholder values with empty cells.`
    };
  });
}

// ────────────────────── 3. Case Formatter ──────────────────────

function toTitleCase(str: string): string {
  return str
    .toLowerCase()
    .replace(/(?:^|\s|[-_])\S/g, (ch) => ch.toUpperCase());
}

function toSnakeCase(str: string): string {
  return str
    .replace(/([a-z])([A-Z])/g, "$1_$2")
    .replace(/[\s\-]+/g, "_")
    .toLowerCase();
}

export async function formatCase(
  sheetName: string,
  rangeAddress: string,
  caseFormat: CaseFormat
): Promise<CleanResult> {
  const labels: Record<CaseFormat, string> = {
    title: "Title Case",
    upper: "UPPERCASE",
    lower: "lowercase",
    snake: "snake_case"
  };
  await snapshotRangeBeforeMutation(
    `Format ${labels[caseFormat]}`,
    sheetName,
    rangeAddress
  );

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "formulas"]);
    await context.sync();

    const values = range.values as CellMatrix;
    const formulas = range.formulas as CellMatrix;

    const transform = (s: string): string => {
      switch (caseFormat) {
        case "title":
          return toTitleCase(s);
        case "upper":
          return s.toUpperCase();
        case "lower":
          return s.toLowerCase();
        case "snake":
          return toSnakeCase(s);
      }
    };

    const cleaned: CellMatrix = values.map((row, r) =>
      row.map((cell, c) => {
        if (isFormula(formulas[r][c])) return cell;
        if (typeof cell !== "string" || cell === "") return cell;
        return transform(cell);
      })
    );

    const { total, modified } = countModified(values, cleaned, formulas);
    range.formulas = buildWriteMatrix(formulas, cleaned);
    await context.sync();

    return {
      totalCells: total,
      modifiedCells: modified,
      details: `Converted ${modified} cells to ${labels[caseFormat]}.`
    };
  });
}

// ────────────────────── 4a. Number Standardizer ──────────────────────

function parseNumericString(val: string): number | null {
  let s = val.trim();
  if (s === "") return null;

  // Handle parentheses for negatives: (100) -> -100
  const isNegParen = /^\((.+)\)$/.test(s);
  if (isNegParen) {
    s = s.slice(1, -1);
  }

  // Strip currency symbols and surrounding whitespace
  s = s.replace(/^[\$€£¥₹₩₫₽R\s]+|[\$€£¥₹₩₫₽R\s]+$/g, "");

  // Remove thousand-separator commas
  s = s.replace(/,/g, "");
  s = s.trim();

  if (s === "" || !/^-?\d+\.?\d*$/.test(s)) return null;

  const num = parseFloat(s);
  if (isNaN(num) || !isFinite(num)) return null;

  return isNegParen ? -num : num;
}

export async function standardizeNumbers(
  sheetName: string,
  rangeAddress: string
): Promise<CleanResult> {
  await snapshotRangeBeforeMutation(
    "Standardize Numbers",
    sheetName,
    rangeAddress
  );

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "formulas"]);
    await context.sync();

    const values = range.values as CellMatrix;
    const formulas = range.formulas as CellMatrix;

    const cleaned: CellMatrix = values.map((row, r) =>
      row.map((cell, c) => {
        if (isFormula(formulas[r][c])) return cell;
        if (typeof cell !== "string") return cell;
        const parsed = parseNumericString(cell);
        return parsed !== null ? parsed : cell;
      })
    );

    const { total, modified } = countModified(values, cleaned, formulas);
    range.formulas = buildWriteMatrix(formulas, cleaned);
    await context.sync();

    return {
      totalCells: total,
      modifiedCells: modified,
      details: `Converted ${modified} text-as-number cells to true numeric values.`
    };
  });
}

// ────────────────────── 4b. Date Standardizer ──────────────────────

/**
 * Converts a JS Date to an Excel serial date number.
 * Accounts for the Lotus 1-2-3 leap-year bug (Feb 29, 1900 = serial 60).
 */
function dateToExcelSerial(date: Date): number {
  const epoch = Date.UTC(1899, 11, 31); // Dec 31, 1899 = serial 0
  const msPerDay = 86_400_000;
  let serial = Math.round((date.getTime() - epoch) / msPerDay);
  // Lotus 1-2-3 bug: all dates >= Mar 1, 1900 (serial >= 60) need +1
  if (serial >= 60) serial += 1;
  return serial;
}

const MONTH_MAP: Record<string, number> = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
};

const DATE_PATTERNS: Array<{
  regex: RegExp;
  parse: (m: RegExpMatchArray) => Date | null;
}> = [
  // YYYY-MM-DD or YYYY/MM/DD
  {
    regex: /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/,
    parse: (m) => {
      const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
      return isNaN(d.getTime()) ? null : d;
    }
  },
  // MM/DD/YYYY or MM-DD-YYYY
  {
    regex: /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/,
    parse: (m) => {
      const d = new Date(Date.UTC(+m[3], +m[1] - 1, +m[2]));
      return isNaN(d.getTime()) ? null : d;
    }
  },
  // DD-Mon-YYYY (e.g. 25-Dec-2024)
  {
    regex:
      /^(\d{1,2})[-\s](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[-\s](\d{4})$/i,
    parse: (m) => {
      const mon = MONTH_MAP[m[2].toLowerCase()];
      if (mon === undefined) return null;
      const d = new Date(Date.UTC(+m[3], mon, +m[1]));
      return isNaN(d.getTime()) ? null : d;
    }
  },
  // Mon DD, YYYY (e.g. Dec 25, 2024)
  {
    regex:
      /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),?\s+(\d{4})$/i,
    parse: (m) => {
      const mon = MONTH_MAP[m[1].toLowerCase()];
      if (mon === undefined) return null;
      const d = new Date(Date.UTC(+m[3], mon, +m[2]));
      return isNaN(d.getTime()) ? null : d;
    }
  }
];

function parseDateString(val: string): Date | null {
  const s = val.trim();
  for (const { regex, parse } of DATE_PATTERNS) {
    const match = s.match(regex);
    if (match) return parse(match);
  }
  return null;
}

export async function standardizeDates(
  sheetName: string,
  rangeAddress: string
): Promise<CleanResult> {
  await snapshotRangeBeforeMutation(
    "Standardize Dates",
    sheetName,
    rangeAddress
  );

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "formulas", "numberFormat"]);
    await context.sync();

    const values = range.values as CellMatrix;
    const formulas = range.formulas as CellMatrix;
    const numberFormat = range.numberFormat as StringMatrix;
    const newNumberFormat: StringMatrix = numberFormat.map((row) => [...row]);

    let modified = 0;
    const cleaned: CellMatrix = values.map((row, r) =>
      row.map((cell, c) => {
        if (isFormula(formulas[r][c])) return cell;
        if (typeof cell !== "string") return cell;
        const date = parseDateString(cell);
        if (!date) return cell;
        modified++;
        newNumberFormat[r][c] = "yyyy-mm-dd";
        return dateToExcelSerial(date);
      })
    );

    const total = values.reduce((sum, row) => sum + row.length, 0);
    range.formulas = buildWriteMatrix(formulas, cleaned);
    range.numberFormat = newNumberFormat;
    await context.sync();

    return {
      totalCells: total,
      modifiedCells: modified,
      details: `Standardized ${modified} date cells to YYYY-MM-DD format.`
    };
  });
}

// ────────────────────── 5a. Highlight Duplicates ──────────────────────

export async function highlightDuplicates(
  sheetName: string,
  rangeAddress: string
): Promise<CleanResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "rowCount", "columnCount"]);
    await context.sync();

    const values = range.values as CellMatrix;
    let highlighted = 0;

    for (let c = 0; c < range.columnCount; c++) {
      const seen = new Map<string, number[]>();
      for (let r = 0; r < range.rowCount; r++) {
        const key = String(values[r][c]).toLowerCase().trim();
        if (key === "") continue;
        if (!seen.has(key)) seen.set(key, []);
        seen.get(key)!.push(r);
      }

      for (const [, rows] of seen) {
        if (rows.length <= 1) continue;
        for (const r of rows) {
          range.getCell(r, c).format.fill.color = HIGHLIGHT_COLORS.duplicate;
          highlighted++;
        }
      }
    }

    await context.sync();

    return {
      totalCells: range.rowCount * range.columnCount,
      modifiedCells: highlighted,
      details: `Highlighted ${highlighted} duplicate cells in amber.`
    };
  });
}

// ────────────────────── 5b. Highlight Blanks ──────────────────────

export async function highlightBlanks(
  sheetName: string,
  rangeAddress: string
): Promise<CleanResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "rowCount", "columnCount"]);
    await context.sync();

    const values = range.values as CellMatrix;
    let highlighted = 0;

    for (let r = 0; r < range.rowCount; r++) {
      for (let c = 0; c < range.columnCount; c++) {
        const val = values[r][c];
        if (val === "" || val === null || val === undefined) {
          range.getCell(r, c).format.fill.color = HIGHLIGHT_COLORS.blank;
          highlighted++;
        }
      }
    }

    await context.sync();

    return {
      totalCells: range.rowCount * range.columnCount,
      modifiedCells: highlighted,
      details: `Highlighted ${highlighted} blank/empty cells in blue.`
    };
  });
}

// ────────────────────── 5c. Highlight Outliers (μ ± 2σ) ──────────────────────

export async function highlightOutliers(
  sheetName: string,
  rangeAddress: string
): Promise<CleanResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "rowCount", "columnCount"]);
    await context.sync();

    const values = range.values as CellMatrix;
    let highlighted = 0;

    for (let c = 0; c < range.columnCount; c++) {
      const nums: { row: number; value: number }[] = [];
      for (let r = 0; r < range.rowCount; r++) {
        const val = values[r][c];
        if (typeof val === "number" && isFinite(val)) {
          nums.push({ row: r, value: val });
        }
      }
      if (nums.length < 3) continue;

      const mean = nums.reduce((s, n) => s + n.value, 0) / nums.length;
      const variance =
        nums.reduce((s, n) => s + (n.value - mean) ** 2, 0) / nums.length;
      const stdDev = Math.sqrt(variance);
      if (stdDev === 0) continue;

      const lower = mean - 2 * stdDev;
      const upper = mean + 2 * stdDev;

      for (const { row, value } of nums) {
        if (value < lower || value > upper) {
          range.getCell(row, c).format.fill.color = HIGHLIGHT_COLORS.outlier;
          highlighted++;
        }
      }
    }

    await context.sync();

    return {
      totalCells: range.rowCount * range.columnCount,
      modifiedCells: highlighted,
      details: `Highlighted ${highlighted} outlier cells (beyond μ ± 2σ) in red.`
    };
  });
}

// ────────────────────── 5d. Remove Duplicate Rows ──────────────────────

function rowKey(row: CellPrimitive[], keyIndices: number[]): string {
  return keyIndices
    .map((i) => String(row[i] ?? "").toLowerCase().trim())
    .join("\0");
}

export function previewDuplicateRows(
  values: CellMatrix,
  keyColumnIndices: number[],
  hasHeaderRow: boolean
): DuplicatePreview {
  const startRow = hasHeaderRow ? 1 : 0;
  const seen = new Set<string>();
  let duplicates = 0;
  for (let r = startRow; r < values.length; r++) {
    const key = rowKey(values[r], keyColumnIndices);
    if (seen.has(key)) {
      duplicates++;
    } else {
      seen.add(key);
    }
  }
  const dataRows = values.length - startRow;
  return {
    totalRows: dataRows,
    uniqueRows: dataRows - duplicates,
    duplicateRows: duplicates
  };
}

export async function removeDuplicateRows(
  sheetName: string,
  rangeAddress: string,
  keyColumnIndices: number[],
  hasHeaderRow: boolean
): Promise<CleanResult> {
  await snapshotRangeBeforeMutation(
    "Remove Duplicate Rows",
    sheetName,
    rangeAddress
  );

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load([
      "values",
      "formulas",
      "numberFormat",
      "rowCount",
      "columnCount"
    ]);
    await context.sync();

    const values = range.values as CellMatrix;
    const formulas = range.formulas as CellMatrix;
    const numberFormat = range.numberFormat as StringMatrix;

    const startRow = hasHeaderRow ? 1 : 0;
    const seen = new Set<string>();
    const keepIndices: number[] = [];

    if (hasHeaderRow) keepIndices.push(0);

    for (let r = startRow; r < values.length; r++) {
      const key = rowKey(values[r], keyColumnIndices);
      if (!seen.has(key)) {
        seen.add(key);
        keepIndices.push(r);
      }
    }

    const removed = values.length - keepIndices.length;
    if (removed === 0) {
      return {
        totalCells: range.rowCount * range.columnCount,
        modifiedCells: 0,
        details: "No duplicate rows found."
      };
    }

    const colCount = range.columnCount;
    const newFormulas: CellMatrix = [];
    const newNumberFormat: StringMatrix = [];

    for (const idx of keepIndices) {
      newFormulas.push([...formulas[idx]]);
      newNumberFormat.push([...numberFormat[idx]]);
    }

    const emptyRow: CellPrimitive[] = new Array(colCount).fill("");
    const defaultFormat: string[] = new Array(colCount).fill("General");
    for (let r = keepIndices.length; r < values.length; r++) {
      newFormulas.push([...emptyRow]);
      newNumberFormat.push([...defaultFormat]);
    }

    range.formulas = newFormulas;
    range.numberFormat = newNumberFormat;
    await context.sync();

    return {
      totalCells: range.rowCount * range.columnCount,
      modifiedCells: removed * colCount,
      details: `Removed ${removed} duplicate rows (kept ${keepIndices.length}${hasHeaderRow ? " incl. header" : ""}).`
    };
  });
}
