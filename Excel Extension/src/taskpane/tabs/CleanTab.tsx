import React, { useCallback, useState } from "react";
import {
  Eraser,
  Ban,
  Type,
  Hash,
  Calendar,
  Paintbrush,
  ShieldCheck
} from "lucide-react";
import { getActiveDataRange } from "../services/excelCore";
import {
  cleanWhitespace,
  normalizeSentinels,
  formatCase,
  standardizeNumbers,
  standardizeDates,
  highlightDuplicates,
  highlightBlanks,
  highlightOutliers,
  previewDuplicateRows,
  removeDuplicateRows,
  DEFAULT_SENTINELS,
  type CaseFormat,
  type CleanResult,
  type DuplicatePreview
} from "../services/cleanerService";
import type { ActiveRangeData, ToastMessage } from "../types/excel";

interface CleanTabProps {
  isOfficeReady: boolean;
  autoDetectTable: boolean;
  showToast: (type: ToastMessage["type"], message: string) => void;
  onActionComplete: () => void;
}

export const CleanTab: React.FC<CleanTabProps> = ({
  isOfficeReady,
  autoDetectTable,
  showToast,
  onActionComplete
}) => {
  const [isProcessing, setIsProcessing] = useState(false);
  const [sentinelInput, setSentinelInput] = useState(DEFAULT_SENTINELS.join(", "));

  // Duplicate-removal flow state
  const [dupState, setDupState] = useState<{
    rangeData: ActiveRangeData;
    selectedKeys: Set<number>;
    preview: DuplicatePreview;
  } | null>(null);

  const runAction = useCallback(
    async (action: () => Promise<CleanResult>) => {
      if (!isOfficeReady) return;
      setIsProcessing(true);
      try {
        const result = await action();
        showToast(result.modifiedCells > 0 ? "success" : "info", result.details);
        onActionComplete();
      } catch (err) {
        showToast("error", err instanceof Error ? err.message : "Action failed.");
      } finally {
        setIsProcessing(false);
      }
    },
    [isOfficeReady, onActionComplete, showToast]
  );

  const handleCleanWhitespace = useCallback(
    () =>
      void runAction(async () => {
        const data = await getActiveDataRange(autoDetectTable);
        return cleanWhitespace(data.summary.sheetName, data.summary.address);
      }),
    [autoDetectTable, runAction]
  );

  const handleNormalizeSentinels = useCallback(
    () =>
      void runAction(async () => {
        const data = await getActiveDataRange(autoDetectTable);
        const placeholders = sentinelInput
          .split(",")
          .map((s) => s.trim())
          .filter((s) => s !== "");
        return normalizeSentinels(data.summary.sheetName, data.summary.address, placeholders);
      }),
    [autoDetectTable, runAction, sentinelInput]
  );

  const handleFormatCase = useCallback(
    (fmt: CaseFormat) =>
      void runAction(async () => {
        const data = await getActiveDataRange(autoDetectTable);
        return formatCase(data.summary.sheetName, data.summary.address, fmt);
      }),
    [autoDetectTable, runAction]
  );

  const handleStandardizeNumbers = useCallback(
    () =>
      void runAction(async () => {
        const data = await getActiveDataRange(autoDetectTable);
        return standardizeNumbers(data.summary.sheetName, data.summary.address);
      }),
    [autoDetectTable, runAction]
  );

  const handleStandardizeDates = useCallback(
    () =>
      void runAction(async () => {
        const data = await getActiveDataRange(autoDetectTable);
        return standardizeDates(data.summary.sheetName, data.summary.address);
      }),
    [autoDetectTable, runAction]
  );

  const handleHighlightDuplicates = useCallback(
    () =>
      void runAction(async () => {
        const data = await getActiveDataRange(autoDetectTable);
        return highlightDuplicates(data.summary.sheetName, data.summary.address);
      }),
    [autoDetectTable, runAction]
  );

  const handleHighlightBlanks = useCallback(
    () =>
      void runAction(async () => {
        const data = await getActiveDataRange(autoDetectTable);
        return highlightBlanks(data.summary.sheetName, data.summary.address);
      }),
    [autoDetectTable, runAction]
  );

  const handleHighlightOutliers = useCallback(
    () =>
      void runAction(async () => {
        const data = await getActiveDataRange(autoDetectTable);
        return highlightOutliers(data.summary.sheetName, data.summary.address);
      }),
    [autoDetectTable, runAction]
  );

  const prepareDupRemoval = useCallback(async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const data = await getActiveDataRange(autoDetectTable);
      const allIndices = data.headers.map((_, i) => i);
      const preview = previewDuplicateRows(data.values, allIndices, data.hasHeaderRow);
      setDupState({
        rangeData: data,
        selectedKeys: new Set(allIndices),
        preview
      });
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to scan range.");
    } finally {
      setIsProcessing(false);
    }
  }, [autoDetectTable, isOfficeReady, showToast]);

  const toggleDupKey = useCallback((colIndex: number) => {
    setDupState((prev) => {
      if (!prev) return null;
      const newKeys = new Set(prev.selectedKeys);
      if (newKeys.has(colIndex)) {
        newKeys.delete(colIndex);
      } else {
        newKeys.add(colIndex);
      }
      const preview = previewDuplicateRows(
        prev.rangeData.values,
        Array.from(newKeys),
        prev.rangeData.hasHeaderRow
      );
      return { ...prev, selectedKeys: newKeys, preview };
    });
  }, []);

  const confirmDupRemoval = useCallback(async () => {
    if (!dupState || !isOfficeReady) return;
    setIsProcessing(true);
    try {
      const { rangeData, selectedKeys } = dupState;
      const result = await removeDuplicateRows(
        rangeData.summary.sheetName,
        rangeData.summary.address,
        Array.from(selectedKeys),
        rangeData.hasHeaderRow
      );
      showToast(result.modifiedCells > 0 ? "success" : "info", result.details);
      onActionComplete();
      setDupState(null);
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to remove duplicates.");
    } finally {
      setIsProcessing(false);
    }
  }, [dupState, isOfficeReady, onActionComplete, showToast]);

  const cardStyle = "rounded-xl border border-slate-200/80 bg-white p-3.5 shadow-sm transition hover:border-slate-300";
  const sectionTitle = "flex items-center gap-2 text-xs font-semibold text-slate-800";
  const btnAction = "w-full rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed";
  const btnOption = "flex-1 rounded-lg border border-slate-200 bg-slate-50/70 px-2.5 py-1.5 text-xs font-medium text-slate-700 transition hover:bg-white hover:border-slate-300 disabled:opacity-50";

  return (
    <div className="space-y-3.5">
      {/* Friendly Reassurance Banner */}
      <div className="flex items-center justify-between rounded-lg border border-emerald-100 bg-emerald-50/60 px-3 py-2 text-[11px] text-emerald-900">
        <div className="flex items-center gap-1.5">
          <ShieldCheck className="h-4 w-4 text-emerald-600 shrink-0" />
          <span>Formulas are protected: cleaning never overwrites active <code className="font-mono text-emerald-800">=formulas</code>.</span>
        </div>
      </div>

      {/* ── 1. Whitespace & Invisible Chars ── */}
      <section className={cardStyle}>
        <div>
          <h3 className={sectionTitle}>
            <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-emerald-50 text-emerald-600">
              <Eraser className="h-3.5 w-3.5" />
            </span>
            Tidy Whitespace &amp; Invisible Characters
          </h3>
          <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
            Trims awkward leading/trailing spaces, collapses multiple spaces into one, and strips non-breaking or invisible unicode characters.
          </p>
        </div>
        <button
          type="button"
          disabled={!isOfficeReady || isProcessing}
          onClick={handleCleanWhitespace}
          className={`mt-2.5 ${btnAction}`}
        >
          Clean Selected Range
        </button>
      </section>

      {/* ── 2. Sentinel & Null Normalizer ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-amber-50 text-amber-600">
            <Ban className="h-3.5 w-3.5" />
          </span>
          Replace Placeholder Nulls
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Turns dummy values like <code className="font-mono text-slate-700">"N/A"</code>, <code className="font-mono text-slate-700">"NULL"</code>, or <code className="font-mono text-slate-700">"-1"</code> into clean empty cells so math functions (AVERAGE, MIN) don't get skewed.
        </p>

        <div className="mt-2.5">
          <label htmlFor="sentinel-input" className="block text-[11px] font-medium text-slate-600 mb-1">
            Values to treat as empty (comma-separated):
          </label>
          <input
            id="sentinel-input"
            type="text"
            value={sentinelInput}
            onChange={(e) => setSentinelInput(e.target.value)}
            className="w-full rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5 font-mono text-[11px] text-slate-800 focus:border-emerald-500 focus:bg-white focus:outline-none focus:ring-1 focus:ring-emerald-500"
          />
        </div>

        <button
          type="button"
          disabled={!isOfficeReady || isProcessing}
          onClick={handleNormalizeSentinels}
          className={`mt-2.5 ${btnAction}`}
        >
          Normalize to Empty Cells
        </button>
      </section>

      {/* ── 3. Smart Case Formatter ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-indigo-50 text-indigo-600">
            <Type className="h-3.5 w-3.5" />
          </span>
          Text Case Formatter
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Standardize mixed names, titles, or code columns into your preferred casing style in one click.
        </p>

        <div className="mt-2.5 grid grid-cols-2 gap-2">
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => handleFormatCase("title")}
            className={btnOption}
          >
            Title Case <span className="text-[10px] text-slate-400 block font-normal">e.g. John Doe</span>
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => handleFormatCase("upper")}
            className={btnOption}
          >
            UPPERCASE <span className="text-[10px] text-slate-400 block font-normal">e.g. USA-100</span>
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => handleFormatCase("lower")}
            className={btnOption}
          >
            lowercase <span className="text-[10px] text-slate-400 block font-normal">e.g. user@site</span>
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => handleFormatCase("snake")}
            className={btnOption}
          >
            snake_case <span className="text-[10px] text-slate-400 block font-normal">e.g. order_id</span>
          </button>
        </div>
      </section>

      {/* ── 4. Number & Date Standardizer ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-blue-50 text-blue-600">
            <Hash className="h-3.5 w-3.5" />
          </span>
          Numbers &amp; Dates Standardizer
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Fix numbers stored as text (e.g. <code className="font-mono text-slate-700">"$1,250"</code> or <code className="font-mono text-slate-700">"(500)"</code>) and convert mixed dates to standard <code className="font-mono text-slate-700">YYYY-MM-DD</code> format.
        </p>

        <div className="mt-2.5 grid grid-cols-2 gap-2">
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={handleStandardizeNumbers}
            className={btnOption}
          >
            <Hash className="mr-1 inline h-3.5 w-3.5 text-blue-600" />
            Fix Numbers
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={handleStandardizeDates}
            className={btnOption}
          >
            <Calendar className="mr-1 inline h-3.5 w-3.5 text-emerald-600" />
            Standardize Dates
          </button>
        </div>
      </section>

      {/* ── 5. Highlight & Deduplicate ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-rose-50 text-rose-600">
            <Paintbrush className="h-3.5 w-3.5" />
          </span>
          Find Anomalies &amp; Remove Duplicates
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Spot duplicates or statistical outliers visually without deleting data, or scan key columns to prune duplicate rows safely.
        </p>

        <div className="mt-2.5 grid grid-cols-3 gap-1.5">
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={handleHighlightDuplicates}
            className={btnOption}
          >
            Highlight Dupes
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={handleHighlightBlanks}
            className={btnOption}
          >
            Highlight Blanks
          </button>
          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={handleHighlightOutliers}
            className={btnOption}
          >
            Highlight Outliers
          </button>
        </div>

        {/* Deduplicate Row Scanner */}
        <div className="mt-3.5 border-t border-slate-100 pt-3">
          <div className="flex items-center justify-between">
            <span className="text-xs font-semibold text-slate-800">Deduplicate Rows</span>
            <span className="text-[10px] text-slate-400">Preview before deleting</span>
          </div>

          {!dupState ? (
            <button
              type="button"
              disabled={!isOfficeReady || isProcessing}
              onClick={() => void prepareDupRemoval()}
              className="mt-2 w-full rounded-lg border border-rose-200 bg-rose-50/70 px-3 py-1.5 text-xs font-semibold text-rose-700 transition hover:bg-rose-100 disabled:opacity-50"
            >
              Scan Selection for Duplicate Rows…
            </button>
          ) : (
            <div className="mt-2 rounded-xl border border-rose-200 bg-rose-50/60 p-3">
              <div className="flex items-center justify-between font-semibold text-rose-950 text-xs">
                <span>
                  {dupState.preview.duplicateRows} duplicate row{dupState.preview.duplicateRows === 1 ? "" : "s"} found
                </span>
                <span className="text-[10px] font-normal text-rose-700">
                  {dupState.preview.uniqueRows} unique rows kept
                </span>
              </div>
              <p className="mt-1 text-[11px] text-rose-800">
                Click headers to toggle which columns identify a duplicate:
              </p>

              <div className="mt-2 flex flex-wrap gap-1.5 max-h-28 overflow-y-auto">
                {dupState.rangeData.headers.map((col) => (
                  <button
                    key={col.index}
                    type="button"
                    onClick={() => toggleDupKey(col.index)}
                    className={`rounded-md px-2 py-1 text-[11px] font-medium transition ${
                      dupState.selectedKeys.has(col.index)
                        ? "border border-emerald-300 bg-emerald-100 text-emerald-800"
                        : "border border-slate-200 bg-white text-slate-500"
                    }`}
                  >
                    {col.header}
                  </button>
                ))}
              </div>

              <div className="mt-3 flex gap-2">
                <button
                  type="button"
                  disabled={
                    dupState.selectedKeys.size === 0 ||
                    dupState.preview.duplicateRows === 0 ||
                    isProcessing
                  }
                  onClick={() => void confirmDupRemoval()}
                  className="flex-1 rounded-lg bg-rose-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:opacity-50"
                >
                  Delete {dupState.preview.duplicateRows} Duplicate Rows
                </button>
                <button
                  type="button"
                  onClick={() => setDupState(null)}
                  className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                >
                  Cancel
                </button>
              </div>
            </div>
          )}
        </div>
      </section>
    </div>
  );
};
