import React, { useCallback, useEffect, useState } from "react";
import {
  Sparkles,
  FunctionSquare,
  BarChart3,
  Layers,
  Bot,
  X,
  Key,
  TableProperties
} from "lucide-react";
import { HeaderBar } from "./components/HeaderBar";
import { ToastBanner } from "./components/ToastBanner";
import { CleanTab } from "./tabs/CleanTab";
import { FormulasTab } from "./tabs/FormulasTab";
import { MisTab } from "./tabs/MisTab";
import { SheetsTab } from "./tabs/SheetsTab";
import { AiHelperTab } from "./tabs/AiHelperTab";
import {
  getActiveDataRange,
  getSelectionSummary,
  getUndoCount,
  getUndoHistory,
  subscribeToSelectionChanges,
  undoLastAction
} from "./services/excelCore";
import type {
  ActiveRangeData,
  SelectionSummary,
  TabId,
  ToastMessage
} from "./types/excel";

interface AppProps {
  isOfficeReady: boolean;
}

const TABS: Array<{
  id: TabId;
  label: string;
  icon: React.ComponentType<{ className?: string }>;
}> = [
  { id: "clean", label: "Clean", icon: Sparkles },
  { id: "formulas", label: "Formulas", icon: FunctionSquare },
  { id: "mis", label: "MIS", icon: BarChart3 },
  { id: "sheets", label: "Sheets", icon: Layers },
  { id: "ai", label: "AI", icon: Bot }
];

export const App: React.FC<AppProps> = ({ isOfficeReady }): JSX.Element => {
  const [activeTab, setActiveTab] = useState<TabId>("clean");
  const [selection, setSelection] = useState<SelectionSummary | null>(null);
  const [activeRangePreview, setActiveRangePreview] =
    useState<ActiveRangeData | null>(null);
  const [showRangeInspector, setShowRangeInspector] = useState(false);
  const [isLoadingSelection, setIsLoadingSelection] = useState(false);
  const [autoDetectTable, setAutoDetectTable] = useState(true);

  const [undoCount, setUndoCount] = useState(0);
  const [lastUndoLabel, setLastUndoLabel] = useState<string | null>(null);
  const [isUndoing, setIsUndoing] = useState(false);

  const [toast, setToast] = useState<ToastMessage | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);

  // Settings in localStorage
  const [aiApiKey, setAiApiKey] = useState(() => localStorage.getItem("gridpulse_ai_key") || "");
  const [aiProvider, setAiProvider] = useState(() => localStorage.getItem("gridpulse_ai_provider") || "openai");

  const showToast = useCallback(
    (type: ToastMessage["type"], message: string) => {
      const id = `${Date.now()}`;
      setToast({ id, type, message });
      setTimeout(() => {
        setToast((prev) => (prev?.id === id ? null : prev));
      }, 4500);
    },
    []
  );

  const syncUndoBadge = useCallback(() => {
    const history = getUndoHistory();
    setUndoCount(getUndoCount());
    setLastUndoLabel(
      history.length > 0 ? history[history.length - 1].actionLabel : null
    );
  }, []);

  const refreshSelection = useCallback(async () => {
    if (!isOfficeReady) {
      return;
    }
    setIsLoadingSelection(true);
    try {
      const summary = await getSelectionSummary();
      setSelection(summary);
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "Failed to read Excel selection.";
      showToast("error", message);
    } finally {
      setIsLoadingSelection(false);
    }
  }, [isOfficeReady, showToast]);

  const handleInspectActiveRange = useCallback(async () => {
    if (!isOfficeReady) {
      return;
    }
    setIsLoadingSelection(true);
    try {
      const data = await getActiveDataRange(autoDetectTable);
      setActiveRangePreview(data);
      setSelection(data.summary);
      setShowRangeInspector(true);
      showToast(
        "info",
        `Selected ${data.summary.sheetName}!${data.summary.shortAddress} (${data.summary.rowCount}×${data.summary.columnCount}, ${data.headers.length} columns).`
      );
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "Unable to inspect active range.";
      showToast("error", message);
    } finally {
      setIsLoadingSelection(false);
    }
  }, [autoDetectTable, isOfficeReady, showToast]);

  const handleUndo = useCallback(async () => {
    if (!isOfficeReady) {
      return;
    }
    setIsUndoing(true);
    try {
      const restored = await undoLastAction();
      syncUndoBadge();
      if (restored) {
        showToast(
          "success",
          `Restored "${restored.actionLabel}" on ${restored.sheetName}!${restored.address}.`
        );
        await refreshSelection();
      } else {
        showToast("info", "Nothing in the undo stack.");
      }
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "Failed to undo last action.";
      showToast("error", message);
    } finally {
      setIsUndoing(false);
    }
  }, [isOfficeReady, refreshSelection, showToast, syncUndoBadge]);

  const saveSettings = () => {
    localStorage.setItem("gridpulse_ai_key", aiApiKey.trim());
    localStorage.setItem("gridpulse_ai_provider", aiProvider);
    setSettingsOpen(false);
    showToast("success", "Preferences saved.");
  };

  useEffect(() => {
    if (!isOfficeReady) {
      return;
    }

    let cleanup: (() => Promise<void>) | null = null;
    void refreshSelection();

    subscribeToSelectionChanges(() => {
      void refreshSelection();
      syncUndoBadge();
    })
      .then((unsub) => {
        cleanup = unsub;
      })
      .catch(() => {
        // Fallback: manual refresh button remains available
      });

    return () => {
      if (cleanup) {
        void cleanup();
      }
    };
  }, [isOfficeReady, refreshSelection, syncUndoBadge]);

  return (
    <div className="flex min-h-screen w-full flex-col bg-slate-50 text-slate-900 font-sans">
      <HeaderBar
        selection={selection}
        isLoadingSelection={isLoadingSelection}
        undoCount={undoCount}
        lastUndoLabel={lastUndoLabel}
        isUndoing={isUndoing}
        onUndo={() => void handleUndo()}
        onRefreshSelection={() => void refreshSelection()}
        onOpenSettings={() => setSettingsOpen(true)}
      />

      {/* Modern Compact Tab Navigation */}
      <nav
        aria-label="GridPulse Modules"
        className="grid grid-cols-5 border-b border-slate-200/90 bg-white px-1 pt-1.5 shadow-[0_1px_2px_rgba(0,0,0,0.02)]"
      >
        {TABS.map((tab) => {
          const Icon = tab.icon;
          const isActive = activeTab === tab.id;
          return (
            <button
              key={tab.id}
              type="button"
              onClick={() => setActiveTab(tab.id)}
              className={`flex flex-col items-center gap-1 rounded-t-lg border-b-2 py-1.5 px-1 text-[11px] font-semibold transition ${
                isActive
                  ? "border-emerald-600 bg-emerald-50/70 text-emerald-800"
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50/60"
              }`}
            >
              <Icon className="h-4 w-4" />
              <span className="truncate">{tab.label}</span>
            </button>
          );
        })}
      </nav>

      <ToastBanner toast={toast} onDismiss={() => setToast(null)} />

      {/* Main Content Area */}
      <main className="flex-1 overflow-y-auto p-3 space-y-3">
        {!isOfficeReady && (
          <div className="rounded-xl border border-amber-200 bg-amber-50/80 p-3 text-xs text-amber-900 shadow-sm">
            <div className="font-semibold flex items-center gap-1.5">
              <span>Preview Mode</span>
            </div>
            <p className="mt-1 text-[11px] text-amber-800 leading-relaxed">
              Running outside Excel host. Sideload inside Excel via <span className="font-mono font-semibold">manifest.xml</span> to interact directly with your workbook cells.
            </p>
          </div>
        )}

        {/* Global Range Controls & Quick Table Detector */}
        <div className="flex items-center justify-between rounded-xl border border-slate-200/80 bg-white px-3 py-2 shadow-sm text-xs">
          <div className="flex items-center gap-2">
            <input
              id="auto-detect-table-toggle"
              type="checkbox"
              checked={autoDetectTable}
              onChange={(e) => setAutoDetectTable(e.target.checked)}
              className="h-3.5 w-3.5 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer"
            />
            <label
              htmlFor="auto-detect-table-toggle"
              className="text-[11px] font-medium text-slate-700 cursor-pointer"
            >
              Auto-expand single cell to table
            </label>
          </div>

          <button
            type="button"
            onClick={() => void handleInspectActiveRange()}
            className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 hover:text-emerald-800"
          >
            <TableProperties className="h-3 w-3" />
            <span>Inspect</span>
          </button>
        </div>

        {/* Collapsible Detected Range Inspector */}
        {activeRangePreview && showRangeInspector && (
          <div className="rounded-xl border border-emerald-200 bg-emerald-50/60 p-2.5 text-xs shadow-sm">
            <div className="flex items-center justify-between font-semibold text-emerald-950">
              <span className="truncate">
                {activeRangePreview.summary.sheetName}!{activeRangePreview.summary.shortAddress}
              </span>
              <button
                type="button"
                onClick={() => setShowRangeInspector(false)}
                className="text-emerald-700 hover:text-emerald-900"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            </div>
            <p className="mt-0.5 text-[10px] text-emerald-800">
              {activeRangePreview.summary.rowCount} rows × {activeRangePreview.summary.columnCount} columns ({activeRangePreview.hasHeaderRow ? "Header detected" : "No header row"})
            </p>
            <div className="mt-2 flex flex-wrap gap-1 max-h-24 overflow-y-auto">
              {activeRangePreview.headers.map((col) => (
                <span
                  key={col.columnLetter}
                  className="rounded-md border border-emerald-200 bg-white px-1.5 py-0.5 text-[10px] text-slate-700 shadow-2xs"
                >
                  <span className="font-mono font-bold text-emerald-700">{col.columnLetter}</span>: {col.header}
                </span>
              ))}
            </div>
          </div>
        )}

        {/* ── Tab Views ── */}
        {activeTab === "clean" && (
          <CleanTab
            isOfficeReady={isOfficeReady}
            autoDetectTable={autoDetectTable}
            showToast={showToast}
            onActionComplete={() => {
              syncUndoBadge();
              void refreshSelection();
            }}
          />
        )}

        {activeTab === "formulas" && (
          <FormulasTab
            isOfficeReady={isOfficeReady}
            autoDetectTable={autoDetectTable}
            showToast={showToast}
            onActionComplete={() => {
              syncUndoBadge();
              void refreshSelection();
            }}
          />
        )}

        {activeTab === "mis" && (
          <MisTab
            isOfficeReady={isOfficeReady}
            autoDetectTable={autoDetectTable}
            showToast={showToast}
            onActionComplete={() => {
              syncUndoBadge();
              void refreshSelection();
            }}
          />
        )}

        {activeTab === "sheets" && (
          <SheetsTab
            isOfficeReady={isOfficeReady}
            showToast={showToast}
            onActionComplete={() => {
              syncUndoBadge();
              void refreshSelection();
            }}
          />
        )}

        {activeTab === "ai" && (
          <AiHelperTab
            isOfficeReady={isOfficeReady}
            onOpenSettings={() => setSettingsOpen(true)}
            showToast={showToast}
            onActionComplete={() => {
              syncUndoBadge();
              void refreshSelection();
            }}
          />
        )}
      </main>

      {/* Human-Centered Settings Drawer */}
      {settingsOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-xs p-3">
          <div
            role="dialog"
            aria-modal="true"
            className="w-full max-w-xs rounded-2xl border border-slate-200 bg-white p-4 shadow-2xl"
          >
            <div className="flex items-center justify-between pb-2 border-b border-slate-100">
              <h2 className="text-sm font-bold text-slate-900">Preferences &amp; BYOK Key</h2>
              <button
                type="button"
                onClick={() => setSettingsOpen(false)}
                aria-label="Close settings"
                className="rounded-lg p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="mt-3 space-y-3 text-xs">
              <div>
                <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                  AI Provider:
                </label>
                <select
                  value={aiProvider}
                  onChange={(e) => setAiProvider(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs text-slate-800"
                >
                  <option value="openai">OpenAI (GPT-4o)</option>
                  <option value="gemini">Google Gemini (Gemini 2.0 / 1.5)</option>
                  <option value="anthropic">Anthropic (Claude 3.5)</option>
                </select>
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                  API Key (Stored locally in browser):
                </label>
                <input
                  type="password"
                  value={aiApiKey}
                  onChange={(e) => setAiApiKey(e.target.value)}
                  placeholder="sk-..."
                  className="w-full rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5 font-mono text-[11px] text-slate-800"
                />
                <span className="mt-1 block text-[10px] text-slate-400">
                  Your key is only saved in your browser's localStorage and never sent anywhere else.
                </span>
              </div>
            </div>

            <div className="mt-4 flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setSettingsOpen(false)}
                className="rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={saveSettings}
                className="rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-emerald-700"
              >
                Save Changes
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
