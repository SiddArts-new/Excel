import type { CellMatrix } from "../types/excel";
import { getActiveDataRange } from "./excelCore";

export interface DetailedSheetInfo {
  id: string;
  name: string;
  visibility: Excel.SheetVisibility | string;
  tabColor: string | null;
  position: number;
}

export interface SheetActionResult {
  success: boolean;
  details: string;
  modifiedCount?: number;
}

/**
 * Loads detailed information about all worksheets in the active workbook.
 */
export async function getDetailedSheets(): Promise<DetailedSheetInfo[]> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load(["items/id", "items/name", "items/visibility", "items/tabColor", "items/position"]);
    await context.sync();

    return sheets.items.map((s) => ({
      id: s.id,
      name: s.name,
      visibility: s.visibility,
      tabColor: s.tabColor || null,
      position: s.position
    }));
  });
}

/**
 * Activates a worksheet by name and brings it to focus.
 */
export async function activateSheetByName(sheetName: string): Promise<void> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.activate();
    await context.sync();
  });
}

/**
 * Unhides all hidden worksheets in the workbook in a single sync.
 */
export async function unhideAllSheets(): Promise<SheetActionResult> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load(["items/name", "items/visibility"]);
    await context.sync();

    let unhiddenCount = 0;
    for (const sheet of sheets.items) {
      if (sheet.visibility !== Excel.SheetVisibility.visible) {
        sheet.visibility = Excel.SheetVisibility.visible;
        unhiddenCount++;
      }
    }

    if (unhiddenCount > 0) {
      await context.sync();
    }

    return {
      success: true,
      modifiedCount: unhiddenCount,
      details: unhiddenCount > 0
        ? `Unhid ${unhiddenCount} previously hidden worksheet${unhiddenCount === 1 ? "" : "s"}.`
        : "All worksheets in the workbook are already visible."
    };
  });
}

/**
 * Sorts all worksheet tabs alphabetically (A-Z or Z-A).
 */
export async function sortSheetsAlphabetically(direction: "asc" | "desc"): Promise<SheetActionResult> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load(["items/name", "items/id"]);
    await context.sync();

    const sortedNames = [...sheets.items.map((s) => s.name)].sort((a, b) => {
      const cmp = a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
      return direction === "asc" ? cmp : -cmp;
    });

    // Reorder tabs: moving each sheet sequentially to the end
    for (const name of sortedNames) {
      const sheet = context.workbook.worksheets.getItem(name);
      sheet.position = sheets.items.length;
    }

    await context.sync();

    return {
      success: true,
      modifiedCount: sortedNames.length,
      details: `Reordered ${sortedNames.length} tabs in ${direction === "asc" ? "A → Z" : "Z → A"} alphabetical order.`
    };
  });
}

const PALETTE = [
  "#10B981", // emerald
  "#3B82F6", // blue
  "#8B5CF6", // purple
  "#F59E0B", // amber
  "#EC4899", // pink
  "#14B8A6", // teal
  "#6366F1", // indigo
  "#F97316"  // orange
];

/**
 * Groups and color-codes sheet tabs based on common name prefixes (e.g. Sales_, MIS_, 2024_).
 */
export async function colorCodeSheetsByPrefix(): Promise<SheetActionResult> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load(["items/name", "items/id"]);
    await context.sync();

    // Detect prefix before _, -, space, or first 4 digits (year)
    const prefixMap = new Map<string, string[]>(); // prefix -> sheet names

    for (const sheet of sheets.items) {
      const name = sheet.name.trim();
      let prefix = "General";

      const delimiterMatch = name.match(/^([A-Za-z0-9]+)[_\-\s]/);
      if (delimiterMatch) {
        prefix = delimiterMatch[1].toUpperCase();
      } else if (/^\d{4}/.test(name)) {
        prefix = name.slice(0, 4);
      }

      const list = prefixMap.get(prefix) || [];
      list.push(sheet.name);
      prefixMap.set(prefix, list);
    }

    let colorIdx = 0;
    let coloredSheets = 0;

    for (const [, sheetNames] of prefixMap) {
      const color = PALETTE[colorIdx % PALETTE.length];
      colorIdx++;

      for (const sheetName of sheetNames) {
        const s = context.workbook.worksheets.getItem(sheetName);
        s.tabColor = color;
        coloredSheets++;
      }
    }

    await context.sync();

    return {
      success: true,
      modifiedCount: coloredSheets,
      details: `Color-coded ${coloredSheets} tabs across ${prefixMap.size} detected prefix groups.`
    };
  });
}

/**
 * Generates clean CSV or JSON string representing the active data table.
 */
export async function exportActiveTable(format: "csv" | "json"): Promise<{ content: string; filename: string; rows: number }> {
  const data = await getActiveDataRange(true);
  const values = data.values;
  const headers = data.headers.map((h) => h.header);
  const startRow = data.hasHeaderRow ? 1 : 0;
  const dataRows = values.slice(startRow);

  const cleanName = data.summary.sheetName.replace(/[^a-zA-Z0-9_-]/g, "_");

  if (format === "csv") {
    const escapeCsv = (val: unknown): string => {
      const str = val === null || val === undefined ? "" : String(val).trim();
      if (str.includes(",") || str.includes('"') || str.includes("\n")) {
        return `"${str.replace(/"/g, '""')}"`;
      }
      return str;
    };

    const lines: string[] = [headers.map(escapeCsv).join(",")];
    for (const row of dataRows) {
      lines.push(row.map(escapeCsv).join(","));
    }

    return {
      content: lines.join("\n"),
      filename: `${cleanName}_Export.csv`,
      rows: dataRows.length
    };
  } else {
    // JSON export
    const jsonObjects: Array<Record<string, CellPrimitive | null>> = [];

    for (const row of dataRows) {
      const obj: Record<string, CellPrimitive | null> = {};
      for (let c = 0; c < headers.length; c++) {
        const key = headers[c] || `Column_${c + 1}`;
        const val = row[c];
        obj[key] = val === "" || val === null || val === undefined ? null : val;
      }
      jsonObjects.push(obj);
    }

    return {
      content: JSON.stringify(jsonObjects, null, 2),
      filename: `${cleanName}_Export.json`,
      rows: dataRows.length
    };
  }
}
