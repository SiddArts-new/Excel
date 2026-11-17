import type {
  ActiveRangeData,
  CellMatrix,
  DetectedColumn,
  SelectionSummary,
  StringMatrix,
  UndoSnapshot
} from "../types/excel";

const MAX_UNDO_STEPS = 5;

const undoStack: UndoSnapshot[] = [];

/**
 * Converts a 0-based column index to an Excel column letter (0 -> A, 25 -> Z, 26 -> AA).
 */
export function columnIndexToLetter(index: number): string {
  let n = index;
  let result = "";
  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

/**
 * Strips the worksheet prefix from an Excel A1 address (e.g. "Sheet1!A1:D10" -> "A1:D10").
 */
export function stripSheetPrefix(address: string): string {
  const bangIdx = address.lastIndexOf("!");
  return bangIdx >= 0 ? address.slice(bangIdx + 1) : address;
}

/**
 * Parses the starting column index from an A1 address (e.g. "C4:F20" -> 2).
 */
export function getStartColumnIndexFromAddress(address: string): number {
  const clean = stripSheetPrefix(address).split(":")[0] ?? "A1";
  const match = clean.match(/^\$?([A-Za-z]+)/);
  if (!match) {
    return 0;
  }
  const letters = match[1].toUpperCase();
  let index = 0;
  for (let i = 0; i < letters.length; i += 1) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return Math.max(0, index - 1);
}

/**
 * Heuristic header detector on the loaded 2D values matrix.
 */
export function detectHeaders(
  values: CellMatrix,
  startColumnIndex = 0
): { hasHeaderRow: boolean; headers: DetectedColumn[] } {
  if (values.length === 0 || values[0].length === 0) {
    return { hasHeaderRow: false, headers: [] };
  }

  const colCount = values[0].length;
  const firstRow = values[0];

  let nonEmptyFirstRow = 0;
  let stringCellsFirstRow = 0;
  for (let c = 0; c < colCount; c += 1) {
    const val = firstRow[c];
    if (val !== "" && val !== null && val !== undefined) {
      nonEmptyFirstRow += 1;
      if (typeof val === "string") {
        stringCellsFirstRow += 1;
      }
    }
  }

  let hasNumericInBody = false;
  for (let r = 1; r < Math.min(values.length, 6); r += 1) {
    for (let c = 0; c < colCount; c += 1) {
      if (typeof values[r][c] === "number") {
        hasNumericInBody = true;
        break;
      }
    }
    if (hasNumericInBody) {
      break;
    }
  }

  const mostlyStringsInRow0 =
    nonEmptyFirstRow > 0 && stringCellsFirstRow / nonEmptyFirstRow >= 0.6;
  const hasHeaderRow =
    values.length > 1 && (mostlyStringsInRow0 || hasNumericInBody);

  const headers: DetectedColumn[] = [];
  for (let c = 0; c < colCount; c += 1) {
    const colIndex = startColumnIndex + c;
    const columnLetter = columnIndexToLetter(colIndex);
    const rawHeader = hasHeaderRow ? String(firstRow[c] ?? "").trim() : "";
    headers.push({
      index: c,
      columnLetter,
      header: rawHeader !== "" ? rawHeader : `Column ${columnLetter}`
    });
  }

  return { hasHeaderRow, headers };
}

/**
 * Resolves the target range for the active selection.
 * When a single cell is selected and autoDetectTable is true, expands via getSurroundingRegion().
 * Uses explicit `.load()` calls and batched `context.sync()`.
 */
export async function getActiveDataRange(
  autoDetectTable = true
): Promise<ActiveRangeData> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load(["id", "name"]);

    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load([
      "address",
      "rowCount",
      "columnCount",
      "values",
      "formulas",
      "numberFormat"
    ]);
    await context.sync();

    const isSingleCell =
      selectedRange.rowCount === 1 && selectedRange.columnCount === 1;

    let targetRange = selectedRange;
    let autoExpanded = false;

    if (isSingleCell && autoDetectTable) {
      const surrounding = selectedRange.getSurroundingRegion();
      surrounding.load([
        "address",
        "rowCount",
        "columnCount",
        "values",
        "formulas",
        "numberFormat"
      ]);
      await context.sync();

      if (surrounding.rowCount > 1 || surrounding.columnCount > 1) {
        targetRange = surrounding;
        autoExpanded = true;
      }
    }

    const values = targetRange.values as CellMatrix;
    const formulas = targetRange.formulas as CellMatrix;
    const numberFormat = targetRange.numberFormat as StringMatrix;
    const shortAddress = stripSheetPrefix(targetRange.address);
    const startColIdx = getStartColumnIndexFromAddress(shortAddress);
    const { hasHeaderRow, headers } = detectHeaders(values, startColIdx);

    const summary: SelectionSummary = {
      sheetId: sheet.id,
      sheetName: sheet.name,
      address: targetRange.address,
      shortAddress,
      rowCount: targetRange.rowCount,
      columnCount: targetRange.columnCount,
      isSingleCell,
      autoExpanded
    };

    return {
      summary,
      values,
      formulas,
      numberFormat,
      hasHeaderRow,
      headers
    };
  });
}

/**
 * Lightweight selection reader for HeaderBar live indicator updates.
 */
export async function getSelectionSummary(): Promise<SelectionSummary> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();

    sheet.load(["id", "name"]);
    range.load(["address", "rowCount", "columnCount"]);
    await context.sync();

    return {
      sheetId: sheet.id,
      sheetName: sheet.name,
      address: range.address,
      shortAddress: stripSheetPrefix(range.address),
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      isSingleCell: range.rowCount === 1 && range.columnCount === 1,
      autoExpanded: false
    };
  });
}

function cloneMatrix<T>(matrix: T[][]): T[][] {
  return matrix.map((row) => [...row]);
}

/**
 * Captures and stores a pre-mutation snapshot of a range into the in-memory undo stack (max 5 steps).
 */
export function pushUndoSnapshot(
  snapshot: Omit<UndoSnapshot, "id" | "timestamp">
): UndoSnapshot {
  const entry: UndoSnapshot = {
    ...snapshot,
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    timestamp: Date.now(),
    values: cloneMatrix(snapshot.values),
    formulas: cloneMatrix(snapshot.formulas),
    numberFormat: cloneMatrix(snapshot.numberFormat)
  };

  undoStack.push(entry);
  if (undoStack.length > MAX_UNDO_STEPS) {
    undoStack.shift();
  }
  return entry;
}

/**
 * Reads the live range state and pushes it to the undo stack before any mutation.
 */
export async function snapshotRangeBeforeMutation(
  actionLabel: string,
  sheetName: string,
  rangeAddress: string
): Promise<UndoSnapshot> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(stripSheetPrefix(rangeAddress));
    range.load(["values", "formulas", "numberFormat", "address"]);
    await context.sync();

    return pushUndoSnapshot({
      actionLabel,
      sheetName,
      address: stripSheetPrefix(range.address),
      values: range.values as CellMatrix,
      formulas: range.formulas as CellMatrix,
      numberFormat: range.numberFormat as StringMatrix
    });
  });
}

/**
 * Builds a restored matrix preserving formula cells (starting with "=") and static values.
 */
function buildRestoreMatrix(values: CellMatrix, formulas: CellMatrix): CellMatrix {
  return formulas.map((row, r) =>
    row.map((formulaCell, c) => {
      if (typeof formulaCell === "string" && formulaCell.startsWith("=")) {
        return formulaCell;
      }
      return values[r]?.[c] ?? "";
    })
  );
}

/**
 * Restores the most recent undo snapshot (both values/formulas and numberFormat) in a single sync.
 */
export async function undoLastAction(): Promise<UndoSnapshot | null> {
  const snapshot = undoStack.pop();
  if (!snapshot) {
    return null;
  }

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(snapshot.sheetName);
    const range = sheet.getRange(stripSheetPrefix(snapshot.address));
    range.formulas = buildRestoreMatrix(snapshot.values, snapshot.formulas);
    range.numberFormat = cloneMatrix(snapshot.numberFormat);
    sheet.activate();
    await context.sync();
  });

  return snapshot;
}

export function getUndoHistory(): readonly UndoSnapshot[] {
  return undoStack;
}

export function getUndoCount(): number {
  return undoStack.length;
}

export function clearUndoHistory(): void {
  undoStack.length = 0;
}

/**
 * Registers Excel selection & worksheet activation listeners for the live HeaderBar indicator.
 * Returns an async cleanup function to remove event handlers on unmount.
 */
export async function subscribeToSelectionChanges(
  onSelectionChange: () => void
): Promise<() => Promise<void>> {
  let selectionHandlerResult: OfficeExtension.EventHandlerResult<Excel.SelectionChangedEventArgs> | null =
    null;
  let activatedHandlerResult: OfficeExtension.EventHandlerResult<Excel.WorksheetActivatedEventArgs> | null =
    null;

  await Excel.run(async (context) => {
    const workbook = context.workbook;
    selectionHandlerResult = workbook.onSelectionChanged.add(async () => {
      onSelectionChange();
    });
    activatedHandlerResult = workbook.worksheets.onActivated.add(async () => {
      onSelectionChange();
    });
    await context.sync();
  });

  return async () => {
    await Excel.run(async (context) => {
      if (selectionHandlerResult) {
        selectionHandlerResult.remove();
      }
      if (activatedHandlerResult) {
        activatedHandlerResult.remove();
      }
      await context.sync();
    });
  };
}
