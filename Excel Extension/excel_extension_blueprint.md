# Project Blueprint: GridPulse — Smart Excel Productivity & MIS Add-in

## 1. Product Vision & Objective
Build **GridPulse**, a fast, modern, cross-platform Microsoft Excel Taskpane Add-in powered by **Office.js**, **React**, and **TypeScript**. The extension eliminates repetitive manual spreadsheet tasks for everyday employees, MIS executives, and data analysts—turning multi-step cleaning, lookup, reporting, and auditing workflows into single-click actions.

---

## 2. Tech Stack & Architecture Constraints

* **Platform:** Microsoft Office Add-in (Excel Taskpane) using `@microsoft/office-js` (`ExcelApi` requirement set `1.14+`)
* **Frontend Framework:** React 18+ with TypeScript (strict mode enabled)
* **Styling:** Tailwind CSS (configured for compact taskpane widths: `300px` to `420px`)
* **Icons:** `lucide-react`
* **Build Tooling:** Webpack / Vite (via standard `yo office` React-TypeScript template)
* **Storage:** Browser `localStorage` for user preferences, saved cleaning presets, and optional BYOK (Bring Your Own Key) AI API settings

### Critical Office.js Engineering Rules (Mandatory for Agent)
1. **Batch Operations Only:** NEVER call `await context.sync()` inside a `for`, `while`, or `.map()` loop. Queue all range reads/writes first, then execute a single `await context.sync()`.
2. **Explicit `.load()`:** Always call `range.load(["values", "formulas", "numberFormat", "address"])` before reading properties, followed by `await context.sync()`.
3. **Preserve Formulas When Cleaning:** When running text/whitespace/null cleanup on a range, inspect `range.formulas`. Skip cells starting with `=` unless the user explicitly chooses "Convert Formulas to Values".
4. **Built-in Safety Undo Stack:** Because Office.js programmatic edits clear Excel's native `Ctrl+Z` undo history, the app **MUST** snapshot the target worksheet/range values and formulas into an in-memory history stack (`max 5 steps`) before executing any mutation, exposing a prominent **"Undo Last Action"** button in the header.

---

## 3. Directory Structure

```text
gridpulse-excel-addin/
├── manifest.xml                  # Office Add-in manifest (ribbon button + taskpane config)
├── package.json
├── tsconfig.json
├── tailwind.config.js
└── src/
    ├── taskpane/
    │   ├── index.tsx             # Office.onReady() bootstrap
    │   ├── App.tsx               # Main layout, header, Undo button, tab router
    │   ├── types/
    │   │   └── excel.ts          # Shared TypeScript interfaces for ranges, snapshots, rules
    │   ├── services/
    │   │   ├── excelCore.ts      # Range readers, header detectors, snapshot/undo engine
    │   │   ├── cleanerService.ts # Data hygiene & sanitization logic
    │   │   ├── formulaService.ts # XLOOKUP builder, IFERROR wrapper, Regex extractor
    │   │   ├── misService.ts     # Auto-KPI summary generator, variance calculator, reconciler
    │   │   ├── sheetService.ts   # Multi-sheet search, sort, hide/unhide utilities
    │   │   └── aiService.ts      # Optional BYOK natural-language-to-formula translator
    │   ├── components/
    │   │   ├── HeaderBar.tsx     # Active selection badge + Undo button + Settings modal trigger
    │   │   ├── ToastBanner.tsx   # Non-intrusive status/row-count feedback notifications
    │   │   └── ConfirmModal.tsx  # Preview & confirmation dialog for destructive actions
    │   └── tabs/
    │       ├── CleanTab.tsx      # 1-Click Data Hygiene & Sanitizer UI
    │       ├── FormulasTab.tsx   # Visual XLOOKUP, Regex Extract & Formula Guard UI
    │       ├── MisTab.tsx        # Instant Report, Variance & Sheet Compare UI
    │       ├── SheetsTab.tsx     # Workbook Navigation & Bulk Sheet Manager UI
    │       └── AiHelperTab.tsx   # Context-aware Natural Language Formula & Insight assistant
```

---

## 4. Core Feature Specifications

### Module 1: Data Hygiene & Sanitizer (`CleanTab.tsx` & `cleanerService.ts`)
Operates on the currently selected range (or auto-expands to the active contiguous table if a single cell is selected and "Auto-detect table" is toggled on).

1. **Whitespace & Invisible Character Stripper:**
   * Trims leading/trailing spaces, collapses multiple internal spaces into one, and strips non-printable ASCII/Unicode control characters (`\u0000`–`\u001F`, `\u00A0` non-breaking spaces).
2. **Sentinel Null & Placeholder Normalizer:**
   * Replaces user-configurable placeholder strings (`-1`, `N/A`, `n/a`, `NULL`, `null`, `None`, `#N/A`, `-`) with true empty cells (`""`) so `AVERAGE`, `COUNT`, and `MIN` calculations are never skewed.
3. **Smart Case Formatter:**
   * One-click conversion to `Title Case`, `UPPERCASE`, `lowercase`, or `snake_case`.
4. **Number & Date Standardizer:**
   * Converts numeric strings stored as text (e.g., `" 1,250.50 "`, `"$400"`) into true numeric values.
   * Standardizes mixed date strings into ISO `YYYY-MM-DD` format with proper Excel date number formatting.
5. **Deduplicate & Anomaly Highlighter:**
   * **Highlight Duplicates / Blanks / Outliers:** Colors duplicate values, empty cells, or numeric outliers exceeding $\mu \pm 2\sigma$ in soft amber/red fill without deleting data.
   * **Remove Duplicate Rows:** Shows a preview count of duplicate rows based on selected key columns before removing them.

---

### Module 2: Smart Formula & Lookup Builder (`FormulasTab.tsx` & `formulaService.ts`)
Eliminates syntax errors and saves employees from writing complex nested formulas manually.

1. **Visual `XLOOKUP` Builder:**
   * Dropdown selectors populated dynamically from sheet headers:
     * *Lookup Value Column* (Active Sheet)
     * *Target Sheet & Lookup Array Column*
     * *Return Column(s)*
     * *If Not Found Fallback* (default: `"Not Found"` or `0`)
   * Generates and fills down the exact `=XLOOKUP(...)` formula (with optional legacy `=IFERROR(INDEX(..., MATCH(...)), ...)` toggle for older Excel compatibility).
2. **Formula Guard (`IFERROR` Wrapper & Value Freeze):**
   * **Wrap in `IFERROR`:** Scans selected formula cells and wraps any formula not already wrapped in `IFERROR` into `=IFERROR(<original_formula>, 0)` or `""`.
   * **Convert Formulas to Values:** Replaces formulas in the selected range with their evaluated static values.
3. **No-Code Pattern / Regex Extractor:**
   * Extracts structured substrings from a messy text column into a newly inserted adjacent column using native Excel formulas or batch JS processing:
     * Extract Email Address
     * Extract Numbers / Currency Only
     * Extract Text Before / After Delimiter (e.g., `-`, `@`, `/`)

---

### Module 3: Instant MIS & Analyst Suite (`MisTab.tsx` & `misService.ts`)
Automates daily reporting tasks that normally take 20–40 minutes of manual formatting.

1. **1-Click Executive Summary Generator:**
   * Scans the active table, detects column data types (Categorical vs. Numeric vs. Date), and creates a cleanly formatted worksheet named `MIS_Summary_<Timestamp>`.
   * Populates:
     * **Top KPI Block:** Total Records, Completeness %, and Sum / Mean / Median / Min / Max for all numeric columns.
     * **Category Breakdown Table:** Uses dynamic `COUNTIFS` and `SUMIFS` formulas grouped by the primary categorical column, styled with bold headers, currency/number formats, and auto-fitted column widths.
2. **Variance & Achievement Calculator:**
   * User selects a **Baseline/Target Column** and an **Actual/Current Column**.
   * Automatically inserts two adjacent columns:
     * `Variance` (`=Actual - Target`)
     * `Variance % / Achievement %` (`=IFERROR((Actual - Target)/Target, 0)`)
   * Applies clean percentage formatting (`0.0%`) and green/red conditional text coloring.
3. **Two-Sheet Reconciler (Compare & Diff):**
   * User picks **Sheet A (Key Column)** and **Sheet B (Key Column)**.
   * Highlights rows present in Sheet A but missing in Sheet B (and vice versa), optionally generating a `Reconciliation_Report` tab listing all unmatched IDs.

---

### Module 4: Workbook & Sheet Navigator (`SheetsTab.tsx` & `sheetService.ts`)
Designed for employees working inside massive, multi-tab corporate workbooks.

1. **Instant Sheet Search & Jump:** Live search bar filtering all visible and hidden worksheets; clicking a sheet activates it immediately.
2. **Bulk Sheet Actions:**
   * Sort all tabs alphabetically (`A → Z` or `Z → A`).
   * Unhide all hidden sheets in one click.
   * Color-code or group tabs by prefix.
3. **Export Active Table as Clean JSON / CSV:** Copies or downloads the active table with trimmed headers and normalized nulls ready for SQL, Python, or Power BI import.

---

### Module 5: Context-Aware AI Formula & Data Assistant (`AiHelperTab.tsx` & `aiService.ts`)
Optional module activated only if the user enters their own API key (OpenAI, Gemini, or Anthropic) in Settings.

1. **Zero-Hallucination Schema Grounding:**
   * Before sending any user prompt to the LLM, the add-in reads the active worksheet's **exact column letters, header names, and 3 sample rows** (never sending the full dataset unless explicitly permitted).
   * System prompt strictly forces the LLM to return valid Excel formula syntax referencing only the real column letters/headers provided.
2. **1-Click Insert:** Displays the generated formula, a plain-English explanation, and an **"Insert into Selected Cell"** button.

---

## 5. UI / UX Design Standards

* **Compact & Scannable:** Designed specifically for a `340px`-wide Excel side pane. No horizontal scrolling.
* **Live Selection Indicator:** Top bar displays the currently active sheet and cell range (e.g., `Sales_Data!A1:G450 • 450 rows`) updated via `Worksheet.onSelectionChanged`.
* **Dry-Run Metrics:** Every action displays exact feedback via toast notification after execution (e.g., *"Cleaned 1,420 cells: trimmed 312 spaces, replaced 48 '-1' placeholders."*).
* **Safe Defaults:** Destructive actions (removing duplicates, overwriting columns) show a lightweight confirmation step with affected row counts.

---

## 6. Phased Execution Plan for the Coding Agent

Follow this exact build order. Verify TypeScript compilation (`npm run build`) at the end of each phase before moving to the next:

* **Phase 1 — Scaffold & Core Engine:** Initialize the React + TypeScript Office Add-in project, configure Tailwind CSS, build `excelCore.ts` (including `getActiveDataRange()`, header detector, and the 5-step `UndoSnapshot` manager), and create `App.tsx` + `HeaderBar.tsx`.
* **Phase 2 — Data Hygiene Engine:** Implement `cleanerService.ts` and `CleanTab.tsx` (whitespace/control char cleaner, sentinel `-1`/`N/A` normalizer, case formatter, text-to-number/date fixer, duplicate/outlier highlighter).
* **Phase 3 — Formula & Lookup Suite:** Implement `formulaService.ts` and `FormulasTab.tsx` (Visual `XLOOKUP` generator, `IFERROR` wrapper, formula-to-value converter, delimiter/pattern extractor).
* **Phase 4 — MIS & Reporting Automations:** Implement `misService.ts` and `MisTab.tsx` (1-click Executive Summary sheet builder, Variance/Achievement column generator, Two-Sheet Reconciler).
* **Phase 5 — Sheet Navigator & AI Assistant:** Implement `SheetsTab.tsx` (instant search, sort, unhide all, clean CSV/JSON export) and `AiHelperTab.tsx` (BYOK schema-grounded formula generator).
* **Phase 6 — Final Audit:** Verify zero `context.sync()` calls inside loops, verify Undo restores both values and formulas accurately, and ensure `manifest.xml` validates cleanly with `npx office-addin-manifest validate manifest.xml`.