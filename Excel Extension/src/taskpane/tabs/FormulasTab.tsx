import React, { useCallback, useEffect, useState } from "react";
import {
  FunctionSquare,
  ShieldCheck,
  Scissors,
  CheckCircle,
  Lock
} from "lucide-react";
import { getActiveDataRange } from "../services/excelCore";
import {
  getWorkbookSheetNames,
  getHeadersForSheet,
  applyVisualLookup,
  wrapSelectedRangeInIfError,
  convertRangeFormulasToValues,
  extractPatternToNewColumn,
  type SheetInfo,
  type ExtractionType
} from "../services/formulaService";
import type { ToastMessage } from "../types/excel";

interface FormulasTabProps {
  isOfficeReady: boolean;
  autoDetectTable: boolean;
  showToast: (type: ToastMessage["type"], message: string) => void;
  onActionComplete: () => void;
}

export const FormulasTab: React.FC<FormulasTabProps> = ({
  isOfficeReady,
  autoDetectTable,
  showToast,
  onActionComplete
}) => {
  const [isProcessing, setIsProcessing] = useState(false);

  // Sheets and columns cache
  const [sheets, setSheets] = useState<SheetInfo[]>([]);
  const [activeSheetHeaders, setActiveSheetHeaders] = useState<
    Array<{ index: number; letter: string; header: string }>
  >([]);
  const [targetSheetHeaders, setTargetSheetHeaders] = useState<
    Array<{ index: number; letter: string; header: string }>
  >([]);

  // ── XLOOKUP State ──
  const [currentSheetName, setCurrentSheetName] = useState("");
  const [lookupColLetter, setLookupColLetter] = useState("A");
  const [targetSheetName, setTargetSheetName] = useState("");
  const [targetLookupColLetter, setTargetLookupColLetter] = useState("A");
  const [targetReturnColLetter, setTargetReturnColLetter] = useState("B");
  const [fallbackValue, setFallbackValue] = useState("Not Found");
  const [useLegacyIndexMatch, setUseLegacyIndexMatch] = useState(false);
  const [outputColLetter, setOutputColLetter] = useState("C");
  const [startRow, setStartRow] = useState(2);
  const [rowCount, setRowCount] = useState(10);

  // ── Formula Guard State ──
  const [ifErrorFallback, setIfErrorFallback] = useState<'""' | "0" | '"Not Found"'>('""');
  const [showFreezeConfirm, setShowFreezeConfirm] = useState(false);

  // ── Pattern Extractor State ──
  const [extractColOffset, setExtractColOffset] = useState(0);
  const [patternType, setPatternType] = useState<ExtractionType>("email");
  const [delimiter, setDelimiter] = useState("-");

  // Load workbook sheets and current active sheet headers
  const loadContextInfo = useCallback(async () => {
    if (!isOfficeReady) return;
    try {
      const sheetList = await getWorkbookSheetNames();
      setSheets(sheetList);

      const activeData = await getActiveDataRange(autoDetectTable);
      const activeSheet = activeData.summary.sheetName;
      setCurrentSheetName(activeSheet);

      const headers = activeData.headers.map((h) => ({
        index: h.index,
        letter: h.columnLetter,
        header: h.header
      }));
      setActiveSheetHeaders(headers);

      if (headers.length > 0) {
        setLookupColLetter(headers[0].letter);
        // Default output to next column after the table
        const lastCol = headers[headers.length - 1];
        setOutputColLetter(lastCol.letter);
      }

      setStartRow(activeData.hasHeaderRow ? 2 : 1);
      setRowCount(Math.max(1, activeData.summary.rowCount - (activeData.hasHeaderRow ? 1 : 0)));

      // Default target sheet to next sheet or active sheet
      const otherSheet = sheetList.find((s) => s.name !== activeSheet) || sheetList[0];
      if (otherSheet) {
        setTargetSheetName(otherSheet.name);
        const targetHeaders = await getHeadersForSheet(otherSheet.name);
        setTargetSheetHeaders(targetHeaders.columns);
        if (targetHeaders.columns.length > 0) {
          setTargetLookupColLetter(targetHeaders.columns[0].letter);
          setTargetReturnColLetter(
            targetHeaders.columns[1]?.letter || targetHeaders.columns[0].letter
          );
        }
      }
    } catch {
      // Handled gracefully in UI
    }
  }, [autoDetectTable, isOfficeReady]);

  useEffect(() => {
    void loadContextInfo();
  }, [loadContextInfo]);

  // When target sheet changes, update its columns
  const handleTargetSheetChange = async (sheetName: string) => {
    setTargetSheetName(sheetName);
    try {
      const info = await getHeadersForSheet(sheetName);
      setTargetSheetHeaders(info.columns);
      if (info.columns.length > 0) {
        setTargetLookupColLetter(info.columns[0].letter);
        setTargetReturnColLetter(
          info.columns[1]?.letter || info.columns[0].letter
        );
      }
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to load sheet headers.");
    }
  };

  // ── Handlers ──

  const handleApplyLookup = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const result = await applyVisualLookup({
        activeSheetName: currentSheetName,
        lookupColLetter,
        targetSheetName,
        targetLookupColLetter,
        targetReturnColLetter,
        fallbackValue,
        useLegacyIndexMatch,
        outputColLetter,
        startRow,
        rowCount
      });
      showToast("success", result.details);
      onActionComplete();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Lookup failed to build.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleWrapIfError = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const data = await getActiveDataRange(autoDetectTable);
      const result = await wrapSelectedRangeInIfError(
        data.summary.sheetName,
        data.summary.address,
        ifErrorFallback
      );
      showToast(result.cellsAffected > 0 ? "success" : "info", result.details);
      onActionComplete();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to wrap formulas.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleFreezeValues = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    setShowFreezeConfirm(false);
    try {
      const data = await getActiveDataRange(autoDetectTable);
      const result = await convertRangeFormulasToValues(
        data.summary.sheetName,
        data.summary.address
      );
      showToast(result.cellsAffected > 0 ? "success" : "info", result.details);
      onActionComplete();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to convert formulas.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleExtractPattern = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const data = await getActiveDataRange(autoDetectTable);
      const result = await extractPatternToNewColumn({
        sheetName: data.summary.sheetName,
        rangeAddress: data.summary.address,
        sourceColOffset: extractColOffset,
        patternType,
        delimiter,
        hasHeaderRow: data.hasHeaderRow
      });
      showToast("success", result.details);
      onActionComplete();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Pattern extraction failed.");
    } finally {
      setIsProcessing(false);
    }
  };

  // Preview formula string for human clarity
  const previewFormula = useLegacyIndexMatch
    ? `=IFERROR(INDEX('${targetSheetName || "Target"}'!${targetReturnColLetter}:${targetReturnColLetter}, MATCH(${lookupColLetter}${startRow}, '${targetSheetName || "Target"}'!${targetLookupColLetter}:${targetLookupColLetter}, 0)), "${fallbackValue}")`
    : `=XLOOKUP(${lookupColLetter}${startRow}, '${targetSheetName || "Target"}'!${targetLookupColLetter}:${targetLookupColLetter}, '${targetSheetName || "Target"}'!${targetReturnColLetter}:${targetReturnColLetter}, "${fallbackValue}")`;

  const cardStyle = "rounded-xl border border-slate-200/80 bg-white p-3.5 shadow-sm transition-all hover:border-slate-300";
  const sectionTitle = "flex items-center gap-2 text-xs font-semibold text-slate-800";
  const selectStyle = "w-full rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs text-slate-800 focus:border-emerald-500 focus:bg-white focus:outline-none focus:ring-1 focus:ring-emerald-500";
  const labelStyle = "block text-[11px] font-medium text-slate-600 mb-1";
  const primaryBtn = "w-full rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed";

  return (
    <div className="space-y-3.5">
      {/* ── 1. Visual XLOOKUP Builder ── */}
      <section className={cardStyle}>
        <div className="flex items-center justify-between">
          <h3 className={sectionTitle}>
            <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-emerald-50 text-emerald-600">
              <FunctionSquare className="h-3.5 w-3.5" />
            </span>
            Visual Lookup Builder
          </h3>
          <span className="rounded-full bg-emerald-50 px-2 py-0.5 text-[10px] font-semibold text-emerald-700">
            {useLegacyIndexMatch ? "INDEX/MATCH" : "XLOOKUP"}
          </span>
        </div>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Match values between sheets without syntax errors or messy comma counting.
        </p>

        <div className="mt-3 space-y-2.5 rounded-lg border border-slate-100 bg-slate-50/60 p-2.5">
          {/* Lookup Value */}
          <div>
            <label className={labelStyle}>Lookup Value (Current Sheet):</label>
            <select
              value={lookupColLetter}
              onChange={(e) => setLookupColLetter(e.target.value)}
              className={selectStyle}
            >
              {activeSheetHeaders.map((col) => (
                <option key={col.letter} value={col.letter}>
                  Col {col.letter} — {col.header}
                </option>
              ))}
            </select>
          </div>

          {/* Target Sheet */}
          <div>
            <label className={labelStyle}>Reference Sheet:</label>
            <select
              value={targetSheetName}
              onChange={(e) => void handleTargetSheetChange(e.target.value)}
              className={selectStyle}
            >
              {sheets.map((s) => (
                <option key={s.id} value={s.name}>
                  {s.name} {s.name === currentSheetName ? "(current)" : ""}
                </option>
              ))}
            </select>
          </div>

          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className={labelStyle}>Match in Column:</label>
              <select
                value={targetLookupColLetter}
                onChange={(e) => setTargetLookupColLetter(e.target.value)}
                className={selectStyle}
              >
                {targetSheetHeaders.map((col) => (
                  <option key={col.letter} value={col.letter}>
                    Col {col.letter}: {col.header}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className={labelStyle}>Return Column:</label>
              <select
                value={targetReturnColLetter}
                onChange={(e) => setTargetReturnColLetter(e.target.value)}
                className={selectStyle}
              >
                {targetSheetHeaders.map((col) => (
                  <option key={col.letter} value={col.letter}>
                    Col {col.letter}: {col.header}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Fallback & Destination */}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className={labelStyle}>Write to Column:</label>
              <input
                type="text"
                value={outputColLetter}
                onChange={(e) => setOutputColLetter(e.target.value.toUpperCase())}
                placeholder="e.g. D"
                maxLength={3}
                className={selectStyle}
              />
            </div>
            <div>
              <label className={labelStyle}>If Not Found:</label>
              <input
                type="text"
                value={fallbackValue}
                onChange={(e) => setFallbackValue(e.target.value)}
                placeholder="Not Found"
                className={selectStyle}
              />
            </div>
          </div>

          {/* Compatibility Toggle */}
          <div className="flex items-center justify-between border-t border-slate-200/80 pt-2 text-[11px]">
            <span className="text-slate-600">Older Excel (2016/2019) compat:</span>
            <label className="relative inline-flex cursor-pointer items-center">
              <input
                type="checkbox"
                checked={useLegacyIndexMatch}
                onChange={(e) => setUseLegacyIndexMatch(e.target.checked)}
                className="h-3.5 w-3.5 rounded text-emerald-600 focus:ring-emerald-500"
              />
              <span className="ml-1.5 text-slate-500 text-[10px]">
                {useLegacyIndexMatch ? "INDEX/MATCH" : "XLOOKUP"}
              </span>
            </label>
          </div>
        </div>

        {/* Live Formula Preview */}
        <div className="mt-2.5 rounded-lg border border-emerald-100 bg-emerald-50/50 p-2">
          <div className="flex items-center gap-1 text-[10px] font-semibold text-emerald-800">
            <CheckCircle className="h-3 w-3" />
            Formula Preview (Row {startRow}):
          </div>
          <div className="mt-0.5 truncate font-mono text-[10px] text-emerald-950" title={previewFormula}>
            {previewFormula}
          </div>
        </div>

        <button
          type="button"
          disabled={!isOfficeReady || isProcessing}
          onClick={() => void handleApplyLookup()}
          className={`mt-3 ${primaryBtn}`}
        >
          {isProcessing ? "Building Formulas…" : `Generate & Fill ${rowCount} Rows`}
        </button>
      </section>

      {/* ── 2. Formula Guard (IFERROR & Value Freeze) ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-sky-50 text-sky-600">
            <ShieldCheck className="h-3.5 w-3.5" />
          </span>
          Formula Guard &amp; Value Freeze
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Clean up unsightly <code className="font-mono text-rose-600">#N/A</code> errors or lock formulas into static values before sharing.
        </p>

        {/* IFERROR Wrapper */}
        <div className="mt-3 rounded-lg border border-slate-100 bg-slate-50/60 p-2.5">
          <label className={labelStyle}>Fallback value when an error occurs:</label>
          <div className="grid grid-cols-3 gap-1.5">
            <button
              type="button"
              onClick={() => setIfErrorFallback('""')}
              className={`rounded-md border py-1 text-[11px] font-medium transition ${
                ifErrorFallback === '""'
                  ? "border-sky-300 bg-sky-100 text-sky-800"
                  : "border-slate-200 bg-white text-slate-600 hover:bg-slate-50"
              }`}
            >
              Empty Blank
            </button>
            <button
              type="button"
              onClick={() => setIfErrorFallback("0")}
              className={`rounded-md border py-1 text-[11px] font-medium transition ${
                ifErrorFallback === "0"
                  ? "border-sky-300 bg-sky-100 text-sky-800"
                  : "border-slate-200 bg-white text-slate-600 hover:bg-slate-50"
              }`}
            >
              Zero (0)
            </button>
            <button
              type="button"
              onClick={() => setIfErrorFallback('"Not Found"')}
              className={`rounded-md border py-1 text-[11px] font-medium transition ${
                ifErrorFallback === '"Not Found"'
                  ? "border-sky-300 bg-sky-100 text-sky-800"
                  : "border-slate-200 bg-white text-slate-600 hover:bg-slate-50"
              }`}
            >
              "Not Found"
            </button>
          </div>

          <button
            type="button"
            disabled={!isOfficeReady || isProcessing}
            onClick={() => void handleWrapIfError()}
            className="mt-2 w-full rounded-md border border-sky-300 bg-sky-50 py-1.5 text-xs font-semibold text-sky-700 transition hover:bg-sky-100 disabled:opacity-50"
          >
            Wrap Selected Range in =IFERROR(…)
          </button>
        </div>

        {/* Freeze to Values */}
        <div className="mt-2.5 rounded-lg border border-slate-100 bg-slate-50/60 p-2.5">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-slate-700">Convert Formulas to Static Values</span>
            <Lock className="h-3.5 w-3.5 text-slate-400" />
          </div>
          <p className="mt-0.5 text-[10px] text-slate-500">
            Replaces formulas in the active range with their calculated results.
          </p>

          {!showFreezeConfirm ? (
            <button
              type="button"
              disabled={!isOfficeReady || isProcessing}
              onClick={() => setShowFreezeConfirm(true)}
              className="mt-2 w-full rounded-md border border-amber-300 bg-amber-50 py-1.5 text-xs font-semibold text-amber-800 transition hover:bg-amber-100 disabled:opacity-50"
            >
              Freeze Range into Static Values…
            </button>
          ) : (
            <div className="mt-2 rounded-md border border-amber-200 bg-amber-100/60 p-2 text-xs">
              <p className="text-[11px] font-medium text-amber-900">
                Are you sure? This converts all formulas in the selection to static text/numbers. (You can still Undo).
              </p>
              <div className="mt-2 flex gap-1.5">
                <button
                  type="button"
                  onClick={() => void handleFreezeValues()}
                  className="flex-1 rounded bg-amber-600 py-1 text-[11px] font-semibold text-white hover:bg-amber-700"
                >
                  Yes, Freeze Values
                </button>
                <button
                  type="button"
                  onClick={() => setShowFreezeConfirm(false)}
                  className="rounded border border-slate-300 bg-white px-2 py-1 text-[11px] font-medium text-slate-600"
                >
                  Cancel
                </button>
              </div>
            </div>
          )}
        </div>
      </section>

      {/* ── 3. Pattern / Regex Extractor ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-purple-50 text-purple-600">
            <Scissors className="h-3.5 w-3.5" />
          </span>
          Pattern Extractor
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Extract specific pieces of information from messy text into an adjacent new column.
        </p>

        <div className="mt-3 space-y-2.5 rounded-lg border border-slate-100 bg-slate-50/60 p-2.5">
          <div>
            <label className={labelStyle}>Source Column to Extract From:</label>
            <select
              value={extractColOffset}
              onChange={(e) => setExtractColOffset(Number(e.target.value))}
              className={selectStyle}
            >
              {activeSheetHeaders.map((col, idx) => (
                <option key={col.letter} value={idx}>
                  Col {col.letter} — {col.header}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className={labelStyle}>What to Extract:</label>
            <div className="grid grid-cols-2 gap-1.5">
              <button
                type="button"
                onClick={() => setPatternType("email")}
                className={`rounded-md border p-1.5 text-left text-[11px] font-medium transition ${
                  patternType === "email"
                    ? "border-purple-300 bg-purple-50 text-purple-800"
                    : "border-slate-200 bg-white text-slate-600 hover:bg-slate-50"
                }`}
              >
                📧 Email Address
              </button>
              <button
                type="button"
                onClick={() => setPatternType("numbers")}
                className={`rounded-md border p-1.5 text-left text-[11px] font-medium transition ${
                  patternType === "numbers"
                    ? "border-purple-300 bg-purple-50 text-purple-800"
                    : "border-slate-200 bg-white text-slate-600 hover:bg-slate-50"
                }`}
              >
                🔢 Numbers Only
              </button>
              <button
                type="button"
                onClick={() => setPatternType("before_delim")}
                className={`rounded-md border p-1.5 text-left text-[11px] font-medium transition ${
                  patternType === "before_delim"
                    ? "border-purple-300 bg-purple-50 text-purple-800"
                    : "border-slate-200 bg-white text-slate-600 hover:bg-slate-50"
                }`}
              >
                ✂️ Text Before Delimiter
              </button>
              <button
                type="button"
                onClick={() => setPatternType("after_delim")}
                className={`rounded-md border p-1.5 text-left text-[11px] font-medium transition ${
                  patternType === "after_delim"
                    ? "border-purple-300 bg-purple-50 text-purple-800"
                    : "border-slate-200 bg-white text-slate-600 hover:bg-slate-50"
                }`}
              >
                ✂️ Text After Delimiter
              </button>
            </div>
          </div>

          {(patternType === "before_delim" || patternType === "after_delim") && (
            <div>
              <label className={labelStyle}>Delimiter Character(s):</label>
              <input
                type="text"
                value={delimiter}
                onChange={(e) => setDelimiter(e.target.value)}
                placeholder="e.g. - or @ or /"
                maxLength={5}
                className={selectStyle}
              />
            </div>
          )}
        </div>

        <button
          type="button"
          disabled={!isOfficeReady || isProcessing}
          onClick={() => void handleExtractPattern()}
          className="mt-3 w-full rounded-lg bg-purple-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-purple-700 disabled:opacity-50"
        >
          {isProcessing ? "Extracting…" : "Extract to New Adjacent Column"}
        </button>
      </section>
    </div>
  );
};
