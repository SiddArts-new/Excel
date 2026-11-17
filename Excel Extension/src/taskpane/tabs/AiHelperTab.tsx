import React, { useCallback, useEffect, useState } from "react";
import {
  Bot,
  Sparkles,
  Key,
  CheckCircle2,
  Copy,
  CornerDownLeft,
  Settings
} from "lucide-react";
import {
  getActiveSchemaContext,
  askAiForFormula,
  insertFormulaIntoActiveCell,
  type SchemaContext,
  type AiFormulaResponse
} from "../services/aiService";
import type { ToastMessage } from "../types/excel";

interface AiHelperTabProps {
  isOfficeReady: boolean;
  onOpenSettings: () => void;
  showToast: (type: ToastMessage["type"], message: string) => void;
  onActionComplete: () => void;
}

export const AiHelperTab: React.FC<AiHelperTabProps> = ({
  isOfficeReady,
  onOpenSettings,
  showToast,
  onActionComplete
}) => {
  const [schema, setSchema] = useState<SchemaContext | null>(null);
  const [prompt, setPrompt] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);
  const [isInserting, setIsInserting] = useState(false);
  const [result, setResult] = useState<AiFormulaResponse | null>(null);

  const apiKey = localStorage.getItem("gridpulse_ai_key") || "";
  const provider = (localStorage.getItem("gridpulse_ai_provider") || "openai") as
    | "openai"
    | "gemini"
    | "anthropic";

  const loadSchema = useCallback(async () => {
    if (!isOfficeReady) return;
    try {
      const ctx = await getActiveSchemaContext();
      setSchema(ctx);
    } catch {
      // Handled
    }
  }, [isOfficeReady]);

  useEffect(() => {
    void loadSchema();
  }, [loadSchema]);

  const handleGenerate = async () => {
    if (!prompt.trim()) {
      showToast("warning", "Please type a formula request.");
      return;
    }

    if (!apiKey) {
      showToast("warning", "Please configure your API key in Settings first.");
      onOpenSettings();
      return;
    }

    if (!schema) {
      showToast("error", "Unable to detect worksheet schema.");
      return;
    }

    setIsGenerating(true);
    setResult(null);

    try {
      const response = await askAiForFormula({
        userPrompt: prompt,
        schema,
        provider,
        apiKey
      });
      setResult(response);
      showToast("success", "Formula generated!");
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "AI generation failed.");
    } finally {
      setIsGenerating(false);
    }
  };

  const handleInsert = async () => {
    if (!result?.formula || !isOfficeReady) return;
    setIsInserting(true);
    try {
      await insertFormulaIntoActiveCell(result.formula);
      showToast("success", `Inserted formula into active cell!`);
      onActionComplete();
    } catch (err) {
      showToast("error", err instanceof Error ? err.message : "Failed to insert formula.");
    } finally {
      setIsInserting(false);
    }
  };

  const handleCopy = async () => {
    if (!result?.formula) return;
    await navigator.clipboard.writeText(result.formula);
    showToast("info", "Formula copied to clipboard.");
  };

  const QUICK_EXAMPLES = [
    "Sum column where Status is 'Completed'",
    "Calculate 10% bonus if Sales > 5000",
    "Find days between Date Col and today",
    "Categorize as High if Amount > 1000 else Low"
  ];

  const cardStyle = "rounded-xl border border-slate-200/80 bg-white p-3.5 shadow-sm transition hover:border-slate-300";
  const sectionTitle = "flex items-center gap-2 text-xs font-semibold text-slate-800";

  return (
    <div className="space-y-3.5">
      {/* ── Schema Grounding Banner ── */}
      <div className="flex items-center justify-between rounded-xl border border-purple-100 bg-purple-50/60 p-2.5 text-xs">
        <div className="flex items-center gap-2 min-w-0">
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-purple-100 text-purple-700 shrink-0">
            <Bot className="h-3.5 w-3.5" />
          </span>
          <div className="min-w-0">
            <span className="font-semibold text-purple-950 block truncate">
              {schema ? `Grounded in ${schema.sheetName}` : "Detecting worksheet…"}
            </span>
            <span className="text-[10px] text-purple-700">
              {schema ? `${schema.columns.length} columns detected • Zero hallucination` : "Waiting for active sheet"}
            </span>
          </div>
        </div>

        <button
          type="button"
          onClick={onOpenSettings}
          title="Configure API Key"
          className="rounded-lg border border-purple-200 bg-white p-1 text-purple-700 hover:bg-purple-50"
        >
          <Settings className="h-3.5 w-3.5" />
        </button>
      </div>

      {/* ── Missing Key Warning ── */}
      {!apiKey && (
        <div className="rounded-xl border border-amber-200 bg-amber-50/70 p-3 text-xs">
          <div className="flex items-start gap-2">
            <Key className="h-4 w-4 text-amber-700 mt-0.5 shrink-0" />
            <div>
              <p className="font-semibold text-amber-900">Bring Your Own Key (BYOK)</p>
              <p className="mt-0.5 text-[11px] text-amber-800 leading-relaxed">
                Add your OpenAI, Gemini, or Anthropic API key to activate instant formula generation. Your key is kept only in your local browser storage.
              </p>
              <button
                type="button"
                onClick={onOpenSettings}
                className="mt-2 rounded-lg bg-amber-700 px-2.5 py-1 text-[11px] font-semibold text-white hover:bg-amber-800 transition"
              >
                Configure API Key in Settings
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Main Prompt Input ── */}
      <section className={cardStyle}>
        <h3 className={sectionTitle}>
          <span className="flex h-6 w-6 items-center justify-center rounded-lg bg-purple-50 text-purple-600">
            <Sparkles className="h-3.5 w-3.5" />
          </span>
          Natural Language Formula Generator
        </h3>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
          Describe the logic in plain English. The AI will write the exact formula using your real sheet column letters.
        </p>

        <div className="mt-2.5">
          <textarea
            rows={3}
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            placeholder="e.g. If Sales in Column C is over 5000 and Region in Column D is 'North', return 'Bonus', otherwise 'Standard'…"
            className="w-full rounded-lg border border-slate-200 bg-slate-50 p-2.5 text-xs text-slate-800 placeholder-slate-400 focus:border-purple-500 focus:bg-white focus:outline-none focus:ring-1 focus:ring-purple-500"
          />
        </div>

        {/* Quick Example Pills */}
        <div className="mt-2">
          <span className="block text-[10px] font-medium text-slate-400 mb-1">Quick Prompts:</span>
          <div className="flex flex-wrap gap-1">
            {QUICK_EXAMPLES.map((ex) => (
              <button
                key={ex}
                type="button"
                onClick={() => setPrompt(ex)}
                className="rounded-md border border-slate-200 bg-slate-50/70 px-1.5 py-0.5 text-[10px] text-slate-600 hover:bg-purple-50 hover:text-purple-700 hover:border-purple-200 transition"
              >
                {ex}
              </button>
            ))}
          </div>
        </div>

        <button
          type="button"
          disabled={!isOfficeReady || isGenerating || !prompt.trim()}
          onClick={() => void handleGenerate()}
          className="mt-3 w-full rounded-lg bg-purple-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-purple-700 disabled:opacity-50"
        >
          {isGenerating ? "Generating Formula…" : "Generate Formula with AI"}
        </button>
      </section>

      {/* ── Result Card ── */}
      {result && (
        <section className="rounded-xl border border-emerald-200 bg-emerald-50/50 p-3.5 shadow-sm space-y-2.5">
          <div className="flex items-center justify-between text-xs">
            <span className="font-semibold text-emerald-950 flex items-center gap-1.5">
              <CheckCircle2 className="h-4 w-4 text-emerald-600" />
              Generated Formula:
            </span>
            <button
              type="button"
              onClick={() => void handleCopy()}
              className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 hover:text-emerald-900"
            >
              <Copy className="h-3 w-3" />
              Copy
            </button>
          </div>

          <div className="rounded-lg border border-emerald-200 bg-white p-2.5 font-mono text-xs font-bold text-emerald-900 break-all select-all">
            {result.formula}
          </div>

          <p className="text-[11px] leading-relaxed text-slate-600">
            {result.explanation}
          </p>

          <button
            type="button"
            disabled={!isOfficeReady || isInserting}
            onClick={() => void handleInsert()}
            className="w-full rounded-lg bg-emerald-600 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:bg-emerald-700 disabled:opacity-50 flex items-center justify-center gap-1.5"
          >
            <CornerDownLeft className="h-3.5 w-3.5" />
            <span>{isInserting ? "Inserting…" : "Insert into Selected Cell"}</span>
          </button>
        </section>
      )}
    </div>
  );
};
