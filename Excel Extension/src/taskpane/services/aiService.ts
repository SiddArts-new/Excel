import type { CellMatrix } from "../types/excel";
import { getActiveDataRange } from "./excelCore";

export interface SchemaContext {
  sheetName: string;
  columns: Array<{ letter: string; header: string }>;
  sampleRows: CellMatrix;
}

export interface AiFormulaResponse {
  formula: string;
  explanation: string;
}

/**
 * Extracts compact, zero-hallucination schema context (headers + 3 sample rows)
 * from the active worksheet to ground the LLM prompt.
 */
export async function getActiveSchemaContext(): Promise<SchemaContext> {
  const data = await getActiveDataRange(true);
  const headers = data.headers.map((h) => ({
    letter: h.columnLetter,
    header: h.header
  }));

  const startRow = data.hasHeaderRow ? 1 : 0;
  const sampleRows = data.values.slice(startRow, startRow + 3);

  return {
    sheetName: data.summary.sheetName,
    columns: headers,
    sampleRows
  };
}

/**
 * Calls the selected LLM provider using the user's BYOK API key.
 * Strictly forces valid Excel formula syntax referencing only real schema columns.
 */
export async function askAiForFormula(options: {
  userPrompt: string;
  schema: SchemaContext;
  provider: "openai" | "gemini" | "anthropic";
  apiKey: string;
}): Promise<AiFormulaResponse> {
  const { userPrompt, schema, provider, apiKey } = options;

  if (!apiKey || apiKey.trim() === "") {
    throw new Error("No API key provided. Please enter your API key in Settings.");
  }

  const columnDescriptions = schema.columns
    .map((c) => `Column ${c.letter}: "${c.header}"`)
    .join(", ");

  const sampleDataStr = schema.sampleRows
    .map((row, i) => `Row ${i + 2}: [${row.map((cell) => String(cell ?? "")).join(", ")}]`)
    .join("\n");

  const systemInstruction = `You are a Microsoft Excel Formula Generator.
You must generate a valid, highly efficient Excel formula that fulfills the user's request.
CRITICAL CONSTRAINTS:
1. ONLY reference the exact column letters provided in the schema: ${columnDescriptions}.
2. Data starts on Row 2 (headers are on Row 1).
3. Always start the formula with '='.
4. Output MUST be valid JSON with two fields: "formula" (string) and "explanation" (short 1-2 sentence plain-English explanation). Do not include markdown code fences or other text.`;

  const userContent = `Worksheet: "${schema.sheetName}"
Columns:
${columnDescriptions}

Sample Data:
${sampleDataStr}

User Request: "${userPrompt}"`;

  if (provider === "openai") {
    const res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey.trim()}`
      },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [
          { role: "system", content: systemInstruction },
          { role: "user", content: userContent }
        ],
        response_format: { type: "json_object" },
        temperature: 0.1
      })
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData?.error?.message || `OpenAI request failed (HTTP ${res.status})`);
    }

    const data = await res.json();
    const parsed = JSON.parse(data.choices[0].message.content);
    return {
      formula: parsed.formula || "",
      explanation: parsed.explanation || "Formula generated successfully."
    };
  } else if (provider === "gemini") {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey.trim()}`;
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [
          {
            parts: [
              { text: `${systemInstruction}\n\n${userContent}` }
            ]
          }
        ],
        generationConfig: {
          responseMimeType: "application/json",
          temperature: 0.1
        }
      })
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData?.error?.message || `Gemini request failed (HTTP ${res.status})`);
    }

    const data = await res.json();
    const rawText = data?.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
    const parsed = JSON.parse(rawText);
    return {
      formula: parsed.formula || "",
      explanation: parsed.explanation || "Formula generated successfully."
    };
  } else {
    // Anthropic
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey.trim(),
        "anthropic-version": "2023-06-01",
        "dangerously-allow-browser": "true"
      },
      body: JSON.stringify({
        model: "claude-3-5-haiku-20241022",
        system: systemInstruction,
        messages: [{ role: "user", content: userContent }],
        max_tokens: 500,
        temperature: 0.1
      })
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData?.error?.message || `Anthropic request failed (HTTP ${res.status})`);
    }

    const data = await res.json();
    const text = data?.content?.[0]?.text || "{}";
    const cleaned = text.replace(/```json|```/g, "").trim();
    const parsed = JSON.parse(cleaned);
    return {
      formula: parsed.formula || "",
      explanation: parsed.explanation || "Formula generated successfully."
    };
  }
}

/**
 * Inserts a generated formula into the currently active selected cell.
 */
export async function insertFormulaIntoActiveCell(formula: string): Promise<void> {
  return Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.formulas = [[formula]];
    await context.sync();
  });
}
