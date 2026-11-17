import React, { useCallback, useEffect, useState } from "react";
import {
  Search,
  ArrowUpDown,
  Eye,
  Palette,
  FileDown,
  Copy,
  ExternalLink,
  Layers
} from "lucide-react";
import {
  getDetailedSheets,
  activateSheetByName,
  unhideAllSheets,
  sortSheetsAlphabetically,
  colorCodeSheetsByPrefix,
  exportActiveTable,
  type DetailedSheetInfo
} from "../services/sheetService";
import type { ToastMessage } from "../types/excel";

interface SheetsTabProps {
  isOfficeReady: boolean;
  showToast: (type: ToastMessage["type"], message: string) => void;
  onActionComplete: () => void;
}

export const SheetsTab: React.FC<SheetsTabProps> = ({
  isOfficeReady,
  showToast,
  onActionComplete
}) => {
  const [sheets, setSheets] = useState<DetailedSheetInfo[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);

  const loadSheets = useCallback(async () => {
    if (!isOfficeReady) return;
    try {
      const list = await getDetailedSheets();
      setSheets(list);
    } catch {
      // Handled
    }
  }, [isOfficeReady]);

  useEffect(() => {
    void loadSheets();
  }, [loadSheets]);

  const handleActivate = async (name: string) => {
    if (!isOfficeReady) return;
    try {
      await activateSheetByName(name);
      showToast("info", `Switched to "${name}".`);
      onActionComplete();
      void loadSheets();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to activate sheet.");
    }
  };

  const handleSort = async (direction: "asc" | "desc") => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const res = await sortSheetsAlphabetically(direction);
      showToast("success", res.details);
      onActionComplete();
      void loadSheets();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to sort tabs.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleUnhideAll = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const res = await unhideAllSheets();
      showToast("success", res.details);
      onActionComplete();
      void loadSheets();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to unhide sheets.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleColorCode = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const res = await colorCodeSheetsByPrefix();
      showToast("success", res.details);
      onActionComplete();
      void loadSheets();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to color tabs.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleCopyCsv = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const result = await exportActiveTable("csv");
      await navigator.clipboard.writeText(result.content);
      showToast("success", `Copied clean CSV (${result.rows} rows) to clipboard.`);
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to export CSV.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownloadJson = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const result = await exportActiveTable("json");
      const blob = new Blob([result.content], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = result.filename;
      a.click();
      URL.revokeObjectURL(url);
      showToast("success", `Downloaded ${result.filename} (${result.rows} records).`);
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to export JSON.");
    } finally {
      setIsProcessing(false);
    }
  };

  const filteredSheets = sheets.filter((s) =>
    s.name.toLowerCase().includes(searchQuery.toLowerCase().trim())
  );

  const cardStyle = "rounded-xl border border-slate-200/80 bg-white p-3.5 shadow-sm transition hover:border-slate-300";
  const sectionTitle = "flex items-center gap-2 text-xs font-semibold text-slate-800";
  const btnAction = "flex-1 rounded-lg border border-slate-200 bg-slate-50/80 px-2.5 py-1.5 text-xs font-medium text-slate-700 transition hover:bg-white hover:border-slate-300 disabled:opacity-50";

  return (
    <div className="space-y-3.5">
      {/* ── 1. Live Sheet Search & Navigator ── */}
      <section className={cardStyle}>
        <div className="flex items-center justify-between">
          <h3 className={sectionTitle}>
            <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-blue-50 text-blue-600">
              <Layers className="h-3.5 w-3.5" />
            </span>
            Workbook Sheet Navigator
          </h3>
          <span className="rounded-full bg-blue-50 px-2 py-0.5 text-[10px] font-semibold text-blue-700">
            {filteredSheets.length} / {sheets.length} tabs
          </span>
        </div>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Search and jump to any worksheet in seconds, even in workbooks with 50+ tabs.
        </p>

        {/* Search input */}
        <div className="relative mt-2.5">
          <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-slate-400" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search sheets by name…"
            className="w-full rounded-lg border border-slate-200 bg-slate-50 pl-8 pr-2.5 py-1.5 text-xs text-slate-800 focus:border-blue-500 focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500"
          />
        </div>

        {/* Scrollable list of sheets */}
        <div className="mt-2.5 max-h-48 overflow-y-auto space-y-1 rounded-lg border border-slate-100 bg-slate-50/50 p-1.5">
          {filteredSheets.length === 0 ? (
            <div className="py-4 text-center text-xs text-slate-400">
              No sheets match "{searchQuery}"
            </div>
          ) : (
            filteredSheets.map((s) => {
              const isHidden = s.visibility !== "Visible" && s.visibility !== Excel.SheetVisibility.visible;
              return (
                <div
                  key={s.id}
                  className="flex items-center justify-between rounded-md bg-white px-2.5 py-1.5 shadow-2xs hover:bg-slate-50 transition"
                >
                  <div className="flex min-w-0 items-center gap-2">
                    <span
                      className="h-2.5 w-2.5 rounded-full shrink-0"
                      style={{ backgroundColor: s.tabColor || "#94A3B8" }}
                    />
                    <span className="truncate text-xs font-medium text-slate-800">
                      {s.name}
                    </span>
                    {isHidden && (
                      <span className="rounded bg-amber-100 px-1 py-0.2 text-[9px] font-semibold text-amber-800">
                        Hidden
                      </span>
                    )}
                  </div>
                  <button
                    type="button"
                    onClick={() => void handleActivate(s.name)}
                    className="inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-[11px] font-semibold text-blue-600 hover:bg-blue-50 transition"
                  >
                    <span>Jump</span>
                    <ExternalLink className="h-3 w-3" />
                  </button>
                </div>
              );
            })
          )}
        </div>
      </section>

      {/* ── 2. Bulk Sheet Actions ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-emerald-50 text-emerald-600">
            <ArrowUpDown className="h-3.5 w-3.5" />
          </span>
          Bulk Sheet Actions
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Organize messy workbooks with one-click sorting, tab coloring, or unhiding.
        </p>

        <div className="mt-2.5 grid grid-cols-2 gap-2">
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => void handleSort("asc")}
            className={btnAction}
          >
            Sort Tabs A → Z
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => void handleSort("desc")}
            className={btnAction}
          >
            Sort Tabs Z → A
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => void handleUnhideAll()}
            className={btnAction}
          >
            <Eye className="mr-1 inline h-3.5 w-3.5 text-blue-600" />
            Unhide All Tabs
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => void handleColorCode()}
            className={btnAction}
          >
            <Palette className="mr-1 inline h-3.5 w-3.5 text-purple-600" />
            Color by Prefix
          </button>
        </div>
      </section>

      {/* ── 3. Clean Table Export ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-indigo-50 text-indigo-600">
            <FileDown className="h-3.5 w-3.5" />
          </span>
          Export Clean Table Data
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Exports the active table with trimmed headers and clean nulls, ready for SQL, Python, or Power BI.
        </p>

        <div className="mt-2.5 grid grid-cols-2 gap-2">
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => void handleCopyCsv()}
            className="flex-1 rounded-lg bg-indigo-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:opacity-50"
          >
            <Copy className="mr-1.5 inline h-3.5 w-3.5" />
            Copy Clean CSV
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => void handleDownloadJson()}
            className="flex-1 rounded-lg border border-indigo-200 bg-indigo-50/70 px-3 py-2 text-xs font-semibold text-indigo-700 shadow-2xs transition hover:bg-indigo-100 disabled:opacity-50"
          >
            <FileDown className="mr-1.5 inline h-3.5 w-3.5" />
            Download JSON
          </button>
        </div>
      </section>
    </div>
  );
};
