import React, { useCallback, useEffect, useState } from "react";
import {
  BarChart3,
  TrendingUp,
  GitCompare
} from "lucide-react";
import { getActiveDataRange } from "../services/excelCore";
import {
  generateExecutiveSummary,
  insertVarianceColumns,
  reconcileTwoSheets,
  type MisActionResult
} from "../services/misService";
import {
  getWorkbookSheetNames,
  getHeadersForSheet,
  type SheetInfo
} from "../services/formulaService";
import type { ToastMessage } from "../types/excel";

interface MisTabProps {
  isOfficeReady: boolean;
  autoDetectTable: boolean;
  showToast: (type: ToastMessage["type"], message: string) => void;
  onActionComplete: () => void;
}

export const MisTab: React.FC<MisTabProps> = ({
  isOfficeReady,
  autoDetectTable,
  showToast,
  onActionComplete
}) => {
  const [isProcessing, setIsProcessing] = useState(false);

  // Sheets & Active Headers cache
  const [sheets, setSheets] = useState<SheetInfo[]>([]);
  const [activeSheetHeaders, setActiveSheetHeaders] = useState<
    Array<{ index: number; letter: string; header: string }>
  >([]);

  // ── Executive Summary State ──
  const [selectedCatOffset, setSelectedCatOffset] = useState<number>(0);
  const [selectedNumOffset, setSelectedNumOffset] = useState<number>(0);

  // ── Variance Calculator State ──
  const [targetColOffset, setTargetColOffset] = useState<number>(0);
  const [actualColOffset, setActualColOffset] = useState<number>(1);

  // ── Reconciler State ──
  const [sheetAName, setSheetAName] = useState("");
  const [sheetBName, setSheetBName] = useState("");
  const [keyColA, setKeyColA] = useState("A");
  const [keyColB, setKeyColB] = useState("A");
  const [sheetAHeaders, setSheetAHeaders] = useState<Array<{ letter: string; header: string }>>([]);
  const [sheetBHeaders, setSheetBHeaders] = useState<Array<{ letter: string; header: string }>>([]);
  const [createReportSheet, setCreateReportSheet] = useState(true);

  const loadData = useCallback(async () => {
    if (!isOfficeReady) return;
    try {
      const sheetList = await getWorkbookSheetNames();
      setSheets(sheetList);

      const data = await getActiveDataRange(autoDetectTable);

      const headers = data.headers.map((h) => ({
        index: h.index,
        letter: h.columnLetter,
        header: h.header
      }));
      setActiveSheetHeaders(headers);

      if (headers.length > 0) {
        setSelectedCatOffset(0);
        setSelectedNumOffset(headers.length > 1 ? 1 : 0);
        setTargetColOffset(0);
        setActualColOffset(headers.length > 1 ? 1 : 0);
      }

      // Reconciler defaults
      if (sheetList.length > 0) {
        const sA = data.summary.sheetName;
        const sB = sheetList.find((s) => s.name !== sA)?.name || sheetList[0].name;
        setSheetAName(sA);
        setSheetBName(sB);

        const hA = await getHeadersForSheet(sA);
        setSheetAHeaders(hA.columns);
        if (hA.columns.length > 0) setKeyColA(hA.columns[0].letter);

        const hB = await getHeadersForSheet(sB);
        setSheetBHeaders(hB.columns);
        if (hB.columns.length > 0) setKeyColB(hB.columns[0].letter);
      }
    } catch {
      // Graceful error fallback
    }
  }, [autoDetectTable, isOfficeReady]);

  useEffect(() => {
    void loadData();
  }, [loadData]);

  const handleSheetBChange = async (name: string) => {
    setSheetBName(name);
    try {
      const info = await getHeadersForSheet(name);
      setSheetBHeaders(info.columns);
      if (info.columns.length > 0) setKeyColB(info.columns[0].letter);
    } catch {
      // Handled
    }
  };

  // ── Actions ──

  const handleGenerateSummary = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const data = await getActiveDataRange(autoDetectTable);
      const result: MisActionResult = await generateExecutiveSummary({
        sourceSheetName: data.summary.sheetName,
        rangeAddress: data.summary.address,
        hasHeaderRow: data.hasHeaderRow,
        selectedCategoryColOffset: selectedCatOffset,
        selectedNumericColOffset: selectedNumOffset
      });
      showToast(result.success ? "success" : "error", result.details);
      onActionComplete();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Summary generation failed.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleInsertVariance = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const data = await getActiveDataRange(autoDetectTable);
      const result: MisActionResult = await insertVarianceColumns({
        sheetName: data.summary.sheetName,
        rangeAddress: data.summary.address,
        targetColOffset,
        actualColOffset,
        hasHeaderRow: data.hasHeaderRow
      });
      showToast(result.success ? "success" : "error", result.details);
      onActionComplete();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to calculate variance.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleRunReconciler = async () => {
    if (!isOfficeReady) return;
    setIsProcessing(true);
    try {
      const result: MisActionResult = await reconcileTwoSheets({
        sheetAName,
        keyColALetter: keyColA,
        sheetBName,
        keyColBLetter: keyColB,
        createReportSheet
      });
      showToast(result.success ? "success" : "error", result.details);
      onActionComplete();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Reconciliation failed.");
    } finally {
      setIsProcessing(false);
    }
  };

  const cardStyle = "rounded-xl border border-slate-200/80 bg-white p-3.5 shadow-sm transition hover:border-slate-300";
  const sectionTitle = "flex items-center gap-2 text-xs font-semibold text-slate-800";
  const selectStyle = "w-full rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs text-slate-800 focus:border-emerald-500 focus:bg-white focus:outline-none focus:ring-1 focus:ring-emerald-500";
  const labelStyle = "block text-[11px] font-medium text-slate-600 mb-1";
  const primaryBtn = "w-full rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed";

  return (
    <div className="space-y-3.5">
      {/* ── 1. Executive Summary ── */}
      <section className={cardStyle}>
        <div className="flex items-center justify-between">
          <h3 className={sectionTitle}>
            <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-emerald-50 text-emerald-600">
              <BarChart3 className="h-3.5 w-3.5" />
            </span>
            1-Click Executive Summary
          </h3>
          <span className="rounded-full bg-emerald-50 px-2 py-0.5 text-[10px] font-semibold text-emerald-700">
            Instant Tab
          </span>
        </div>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Builds a formatted KPI dashboard sheet with completeness %, totals, averages, and a breakdown table.
        </p>

        <div className="mt-3 space-y-2.5 rounded-lg border border-slate-100 bg-slate-50/60 p-2.5">
          <div>
            <label className={labelStyle}>Primary Grouping (Category Column):</label>
            <select
              value={selectedCatOffset}
              onChange={(e) => setSelectedCatOffset(Number(e.target.value))}
              className={selectStyle}
            >
              {activeSheetHeaders.map((col) => (
                <option key={col.letter} value={col.index}>
                  Col {col.letter} — {col.header}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className={labelStyle}>Metric Column to Sum / Average:</label>
            <select
              value={selectedNumOffset}
              onChange={(e) => setSelectedNumOffset(Number(e.target.value))}
              className={selectStyle}
            >
              {activeSheetHeaders.map((col) => (
                <option key={col.letter} value={col.index}>
                  Col {col.letter} — {col.header}
                </option>
              ))}
            </select>
          </div>
        </div>

        <button
          type="button"
          disabled={!isOfficeReady || isProcessing}
          onClick={() => void handleGenerateSummary()}
          className={`mt-3 ${primaryBtn}`}
        >
          {isProcessing ? "Building Dashboard…" : "Generate Executive Summary Tab"}
        </button>
      </section>

      {/* ── 2. Variance & Achievement Calculator ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-teal-50 text-teal-600">
            <TrendingUp className="h-3.5 w-3.5" />
          </span>
          Variance &amp; Achievement Calculator
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Adds <code className="font-mono text-slate-700">Variance</code> and <code className="font-mono text-slate-700">Variance %</code> columns to compare Actual vs. Target performance.
        </p>

        <div className="mt-3 grid grid-cols-2 gap-2 rounded-lg border border-slate-100 bg-slate-50/60 p-2.5">
          <div>
            <label className={labelStyle}>Target (Baseline):</label>
            <select
              value={targetColOffset}
              onChange={(e) => setTargetColOffset(Number(e.target.value))}
              className={selectStyle}
            >
              {activeSheetHeaders.map((col) => (
                <option key={col.letter} value={col.index}>
                  Col {col.letter}: {col.header}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className={labelStyle}>Actual (Current):</label>
            <select
              value={actualColOffset}
              onChange={(e) => setActualColOffset(Number(e.target.value))}
              className={selectStyle}
            >
              {activeSheetHeaders.map((col) => (
                <option key={col.letter} value={col.index}>
                  Col {col.letter}: {col.header}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="mt-2 rounded-lg border border-slate-100 bg-slate-50 p-2 text-[10px] text-slate-600 flex items-center justify-between">
          <span>Formulas: <code className="font-mono text-emerald-700">=Actual - Target</code> &amp; <code className="font-mono text-emerald-700">=(Actual-Target)/Target</code></span>
          <span className="text-emerald-700 font-semibold">Formatted as 0.0%</span>
        </div>

        <button
          type="button"
          disabled={!isOfficeReady || isProcessing}
          onClick={() => void handleInsertVariance()}
          className="mt-3 w-full rounded-lg bg-teal-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-teal-700 disabled:opacity-50"
        >
          {isProcessing ? "Inserting Columns…" : "Insert Variance Columns"}
        </button>
      </section>

      {/* ── 3. Two-Sheet Reconciler ── */}
      <section className={cardStyle}>
        <div className="flex items-center justify-between">
          <h3 className={sectionTitle}>
            <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-indigo-50 text-indigo-600">
              <GitCompare className="h-3.5 w-3.5" />
            </span>
            Two-Sheet Reconciler
          </h3>
          <span className="rounded-full bg-indigo-50 px-2 py-0.5 text-[10px] font-semibold text-indigo-700">
            Compare &amp; Diff
          </span>
        </div>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Match records across two sheets using an ID key (e.g. Order ID, Employee ID). Flags missing rows and creates a report.
        </p>

        <div className="mt-3 space-y-2.5 rounded-lg border border-slate-100 bg-slate-50/60 p-2.5">
          {/* Sheet A */}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className={labelStyle}>Sheet A:</label>
              <select
                value={sheetAName}
                onChange={(e) => setSheetAName(e.target.value)}
                className={selectStyle}
              >
                {sheets.map((s) => (
                  <option key={s.id} value={s.name}>
                    {s.name}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className={labelStyle}>Key Column (A):</label>
              <select
                value={keyColA}
                onChange={(e) => setKeyColA(e.target.value)}
                className={selectStyle}
              >
                {sheetAHeaders.map((col) => (
                  <option key={col.letter} value={col.letter}>
                    Col {col.letter}: {col.header}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Sheet B */}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className={labelStyle}>Sheet B:</label>
              <select
                value={sheetBName}
                onChange={(e) => void handleSheetBChange(e.target.value)}
                className={selectStyle}
              >
                {sheets.map((s) => (
                  <option key={s.id} value={s.name}>
                    {s.name}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className={labelStyle}>Key Column (B):</label>
              <select
                value={keyColB}
                onChange={(e) => setKeyColB(e.target.value)}
                className={selectStyle}
              >
                {sheetBHeaders.map((col) => (
                  <option key={col.letter} value={col.letter}>
                    Col {col.letter}: {col.header}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="flex items-center gap-2 border-t border-slate-200/80 pt-2 text-[11px] text-slate-700">
            <input
              id="recon-report-toggle"
              type="checkbox"
              checked={createReportSheet}
              onChange={(e) => setCreateReportSheet(e.target.checked)}
              className="h-3.5 w-3.5 rounded text-indigo-600 focus:ring-indigo-500 cursor-pointer"
            />
            <label htmlFor="recon-report-toggle" className="cursor-pointer">
              Create a dedicated <span className="font-semibold text-slate-800">Reconciliation Report</span> tab
            </label>
          </div>
        </div>

        <button
          type="button"
          disabled={!isOfficeReady || isProcessing || !sheetAName || !sheetBName}
          onClick={() => void handleRunReconciler()}
          className="mt-3 w-full rounded-lg bg-indigo-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:opacity-50"
        >
          {isProcessing ? "Reconciling Sheets…" : "Run Two-Sheet Reconciliation"}
        </button>
      </section>
    </div>
  );
};
