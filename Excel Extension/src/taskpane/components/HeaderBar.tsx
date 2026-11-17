import React from "react";
import { Grid3X3, RotateCcw, Settings, RefreshCw } from "lucide-react";
import type { SelectionSummary } from "../types/excel";

interface HeaderBarProps {
  selection: SelectionSummary | null;
  isLoadingSelection: boolean;
  undoCount: number;
  lastUndoLabel: string | null;
  isUndoing: boolean;
  onUndo: () => void;
  onRefreshSelection: () => void;
  onOpenSettings: () => void;
}

export const HeaderBar: React.FC<HeaderBarProps> = ({
  selection,
  isLoadingSelection,
  undoCount,
  lastUndoLabel,
  isUndoing,
  onUndo,
  onRefreshSelection,
  onOpenSettings
}) => {
  const rowText = selection
    ? `${selection.rowCount} ${selection.rowCount === 1 ? "row" : "rows"}`
    : "No selection";

  return (
    <header className="sticky top-0 z-20 border-b border-slate-200 bg-white px-3 py-2.5 shadow-sm">
      <div className="flex items-center justify-between gap-2">
        <div className="flex min-w-0 items-center gap-2">
          <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-lg bg-gradient-to-br from-emerald-500 to-teal-700 text-white shadow-sm ring-1 ring-emerald-600/20">
            <Grid3X3 className="h-4 w-4" />
          </div>
          <div className="min-w-0">
            <h1 className="truncate text-sm font-bold tracking-tight text-slate-900">
              GridPulse
            </h1>
            <p className="truncate text-[10px] font-medium text-slate-500">
              Smart Excel Assistant
            </p>
          </div>
        </div>

        <div className="flex shrink-0 items-center gap-1.5">
          <button
            type="button"
            onClick={onUndo}
            disabled={undoCount === 0 || isUndoing}
            title={
              lastUndoLabel
                ? `Undo: ${lastUndoLabel} (${undoCount}/5)`
                : "Nothing to undo"
            }
            className="inline-flex items-center gap-1 rounded-md border border-slate-300 bg-white px-2 py-1 text-xs font-semibold text-slate-700 transition hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-45"
          >
            <RotateCcw className={`h-3.5 w-3.5 ${isUndoing ? "animate-spin" : ""}`} />
            <span>Undo</span>
            {undoCount > 0 && (
              <span className="rounded bg-amber-100 px-1 text-[10px] font-bold text-amber-800">
                {undoCount}
              </span>
            )}
          </button>

          <button
            type="button"
            onClick={onOpenSettings}
            aria-label="Open Settings"
            title="Settings"
            className="rounded-md border border-slate-200 bg-slate-50 p-1.5 text-slate-600 transition hover:bg-slate-100 hover:text-slate-900"
          >
            <Settings className="h-3.5 w-3.5" />
          </button>
        </div>
      </div>

      {/* Live Selection Badge */}
      <div className="mt-2 flex items-center justify-between gap-2 rounded-lg border border-slate-200/90 bg-slate-50/80 px-2.5 py-1 text-xs">
        <div className="flex min-w-0 items-center gap-1.5">
          <span className="relative flex h-2 w-2">
            <span className="absolute inline-flex h-full w-full animate-ping rounded-full bg-emerald-400 opacity-75" />
            <span className="relative inline-flex h-2 w-2 rounded-full bg-emerald-500" />
          </span>
          <span className="truncate font-mono text-[11px] font-medium text-slate-800">
            {selection
              ? `${selection.sheetName}!${selection.shortAddress}`
              : "Click any cell in your sheet…"}
          </span>
        </div>

        <div className="flex shrink-0 items-center gap-1.5">
          <span className="text-[11px] font-medium text-slate-500">
            {rowText}
          </span>
          <button
            type="button"
            onClick={onRefreshSelection}
            aria-label="Refresh Selection"
            title="Refresh active selection"
            className="rounded p-0.5 text-slate-400 transition hover:bg-slate-200 hover:text-slate-800"
          >
            <RefreshCw
              className={`h-3 w-3 ${isLoadingSelection ? "animate-spin text-emerald-600" : ""}`}
            />
          </button>
        </div>
      </div>
    </header>
  );
};
