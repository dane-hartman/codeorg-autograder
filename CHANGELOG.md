# Changelog

All notable changes to the Game Lab Autograder.

## [2.2.0] — 2026-02-18

### Added
- **Rate limit retry with backoff** — all LLM API calls now auto-retry on 429 (Too Many Requests) and 503 errors. Exponential backoff: ~2s → 4s → 8s → 16s with jitter, up to 4 retries. Handles bursts of simultaneous student submissions on the free Gemini tier.
- **Email guard for Error rows** — automatic email flows (`onFormSubmit`, `Grade & Email All New`) skip rows with `Status = Error` to prevent sending confusing API error text to students. Manual **Email Selected Rows** still sends to all statuses.
- **"Sync Levels from Criteria"** menu item — reads the Criteria sheet and rebuilds the Levels sheet from the unique LevelIDs found. Preserves existing Enabled/Model settings. Use after importing a new criteria CSV.
- **`criteria/` folder** — criteria CSV files now live in their own folder, named by curriculum (e.g., `CSD-Unit3-Interactive-Animations-and-Games.csv`). Teachers and contributors can share rubric CSVs without touching `Code.gs`.

### Changed
- **Criteria are no longer embedded in `Code.gs`** — removed `getCriteriaTableCsvText_()`, `parseCriteriaTableCsv_()`, and `parseCsvText_()` (~170 lines). Teachers import criteria via Google Sheets' File → Import. This decouples rubric content from script code and eliminates the need to keep two copies in sync.
- **Initial Setup** now creates empty Criteria and Levels sheets with headers only. The setup completion message guides teachers to import a criteria CSV and run Sync Levels.
- **Simplified grading engine** — removed unused local check types (`code_nonempty`, `contains`, `regex_present`, `regex_absent`) from `runCriteria_()`. All criteria now go through the LLM. The `Type` column still exists on the Criteria sheet but only `llm_check` is used.
- **Menu reorganized** — "Sync Levels from Criteria" placed in the maintenance group (next to "Test API Connection").
- **Dedup fix for form import** — `normalizeTimestamp_()` now rounds to the nearest minute, fixing false mismatches between `e.values` strings and `getValues()` Date objects that caused "Grade New Submissions" to re-import already-graded rows.
- **Help dialog updated** — now includes criteria import instructions and describes all menu items with "Use this when…" guidance.
- **README overhauled** — new Step 3 "Import Criteria" in Quick Setup, Menu Reference table now includes a "Use this when…" column, new "Rate Limit Handling" section, expanded Troubleshooting, updated project structure.
- **CONTRIBUTING.md rewritten** — criteria workflow section explains the CSV-import approach, no more dual-maintenance instructions. Contributing criteria is now as simple as editing a CSV file.

## [2.1.0] — 2026-02-17

### Changed
- **"Grade New Submissions"** now auto-imports from Form Responses before grading — no separate sync step needed
- **"Sync from Form Responses"** removed; functionality merged into Grade New Submissions
- **"Sync, Grade & Email"** renamed to **"Grade & Email All New"** for clarity
- **"Grade Selected Rows"** renamed to **"Re-grade Selected Rows"** (rows in Submissions are already graded)
- **"Test API Connection"** and **"Test Structured JSON"** merged into a single combined test
- **LevelName** column on the Levels sheet replaced with **LevelURL** (clickable Code.org links)
- Setup dialog checkboxes now use CSS grid (3 columns) — no more awkward singleton rows
- Setup dialog buttons disable with "Working…" status while server-side code runs
- Help dialog and all menu descriptions now explicitly reference the Submissions sheet

### Added
- Toast notifications on all long-running operations ("Grading 5 row(s)…", "Sending emails…", etc.)
- Active-sheet validation: Re-grade Selected Rows and Email Selected Rows now prompt the user to switch to the Submissions sheet if they're on the wrong tab
- `importFormResponses_()` — reusable silent import helper (no alerts, returns count)
- `gradeAndEmailAllNew()` — one-click import + grade + email workflow
- `testAPIConnection()` — combined basic connectivity + structured JSON grading test
- `levelIdToUrl_()` — converts LevelIDs to Code.org level URLs

### Removed
- `syncFromFormResponses()` / `syncGradeAndEmail()` / `syncCore_()` — replaced by the simpler import + grade pattern
- `diagnosticsTestLLM()` / `diagnosticsTestLLMStructured()` and all four `ping*` functions — replaced by unified `testAPIConnection()`

## [2.0.0] — 2026-02-17

### Added
- Complete rewrite of the autograder from legacy v1
- **Setup wizard** with HTML dialog and period picker (replaces flat `setupSheets()`)
- **Grade View P# sheets** — formula-driven read-only views per period, auto-sorted by level and student name
- **Single Submissions sheet** for all periods (replaces per-period data duplication)
- **Long-form LevelIDs** (`Lesson-03-Level-08` instead of `L3-08`) — unambiguous and self-documenting
- **In-app Help dialog** with complete setup instructions, form field reference, and menu guide
- **Dual LLM support** — Gemini (default) and OpenAI with automatic fallback strategies
- **Grade caching** via `CacheService` with 6-hour TTL (SHA-256 key of LevelID + source)
- **Conditional formatting** on the Status column (green = OK, red = Error, yellow = warnings)
- **LevelID dropdown validation** in the Submissions sheet
- **Progress toasts** during grading ("Grading row 3 of 12…")
- **`onFormSubmit` trigger** for real-time grading and emailing on form submission
- **Embedded criteria CSV** — rubric data built into Code.gs so no external files are needed at runtime
- **Column width presets** for all sheets

### Changed from v1
- Cleaned rubric — removed legacy non-standard levels (L5-06c, L6-07adv, L6-08a, L9-05adv, L12-07adv)
- Flat menu structure (no nested sub-menus)
- OpenAI integration uses the Responses API with `json_schema` → `json_object` → plain JSON fallback
- Robust `headersSmart_()` mapping handles verbose Google Form column headers

## [1.0.0] — 2025-09-21

### Added
- Original legacy autograder (v1)
- Basic grading pipeline: fetch Code.org source → LLM evaluation → write scores
- Per-period submission management
- Short-form LevelIDs (L3-08, L5-06c, etc.)
- Nested sub-menu structure (Emails >, Admin >, Diagnostics >)
