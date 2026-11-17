export type CellPrimitive = string | number | boolean;
export type CellMatrix = CellPrimitive[][];
export type StringMatrix = string[][];

export type TabId = "clean" | "formulas" | "mis" | "sheets" | "ai";

export interface SelectionSummary {
  sheetId: string;
  sheetName: string;
  address: string;
  shortAddress: string;
  rowCount: number;
  columnCount: number;
  isSingleCell: boolean;
  autoExpanded: boolean;
}

export interface DetectedColumn {
  index: number;
  columnLetter: string;
  header: string;
}

export interface ActiveRangeData {
  summary: SelectionSummary;
  values: CellMatrix;
  formulas: CellMatrix;
  numberFormat: StringMatrix;
  hasHeaderRow: boolean;
  headers: DetectedColumn[];
}

export interface UndoSnapshot {
  id: string;
  actionLabel: string;
  timestamp: number;
  sheetName: string;
  address: string;
  values: CellMatrix;
  formulas: CellMatrix;
  numberFormat: StringMatrix;
}

export interface ToastMessage {
  id: string;
  type: "success" | "info" | "warning" | "error";
  message: string;
}
