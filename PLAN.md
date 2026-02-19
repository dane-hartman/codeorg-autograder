# Game Lab Autograder — Architecture & Design

> Internal reference document. For setup instructions, see [README.md](README.md).

## 1. Project Overview

**Goal:** A Google Apps Script (`.gs`) that lives inside a Google Sheets spreadsheet. Teachers connect a Google Form where students submit their Code.org Game Lab share links. The script fetches the student source code, sends it to an LLM (Gemini or OpenAI) along with rubric criteria, and writes the score + notes back to the spreadsheet. Students optionally receive an automated email with their results.

### Core pipeline
- Extract channel ID from share URL → fetch source from Code.org → build rubric prompt → call LLM → parse structured JSON → write score/notes → email student
- Dual LLM support (Gemini default, OpenAI optional)
- Robust JSON normalization with multiple fallback strategies
- Grade caching via `CacheService` (6-hour TTL, keyed by SHA-256 of LevelID + source)

---

## 2. Architecture: Sheets & Data Flow

### 2.1 — Sheet Inventory

| Sheet | Created by | Purpose |
|---|---|---|
| `Form Responses 1` | Google Forms (automatic) | Raw form submissions. **Never edited by the script.** |
| `Submissions` | Setup wizard | Normalized copy of every submission. All grading reads/writes happen here. |
| `Grade View P1` … `P8` | Setup wizard (per checked period) | **Read-only formula views.** Auto-populated from Submissions. Sorted by LevelID then Last name. |
| `Levels` | Setup wizard | Master level list — 16 rows. Enabled checkbox, LevelURL to Code.org, optional per-level Model override. |
| `Criteria` | Setup wizard | Master rubric — all criterion rows from the embedded CSV. |

### 2.2 — Data Flow

```
Student submits Google Form
        │
        ▼
  ┌─────────────────────┐
  │  Form Responses 1   │   (raw, untouched by script)
  └────────┬────────────┘
           │  onFormSubmit trigger  ─OR─  "Grade New Submissions" imports
           ▼
  ┌─────────────────────┐
  │    Submissions       │   (normalized row; grading writes Score/Notes here)
  └────────┬────────────┘
           │  Grade View sheets pull via SORT(FILTER(...)) formulas
           ▼
  ┌─────────────────────┐
  │  Grade View P3       │   (period 3, sorted by LevelID → Last)
  │  Grade View P4       │   (period 4, ...)
  │  …                   │
  └─────────────────────┘
           │
           ▼
  Student receives email with score + criterion breakdown
```

### 2.3 — Submissions Sheet Columns

```
Timestamp | First | Last | Period | Email | LevelID | ShareURL | ChannelID | Score | MaxScore | Status | Notes | EmailedAt
```

### 2.4 — Grade View Formula

Each `Grade View P#` sheet has a single `SORT(FILTER({...}))` array formula in A2 that:
1. Filters Submissions where Period matches
2. Reorders columns: LevelID, First, Last, Score, MaxScore, Status, Email, ShareURL, Timestamp, Notes
3. Sorts by LevelID ascending, then Last ascending

Sheets are **protected** with `setWarningOnly(true)`.

---

## 3. Code.gs Module Layout

```
 1. CONFIG & CONSTANTS
 2. MENU (onOpen)
 3. SETUP WIZARD
    - showSetupDialog / buildSetupHtml_ (HTML dialog with period picker)
    - createSheetsFromSetup (builds Submissions, Levels, Criteria, Grade Views)
    - resetEverything (wipe all sheets)
    - buildGradeViewFormula_ (SORT/FILTER formula builder)
 4. GRADING ENGINE
    - gradeNewRows (imports from form + grades ungraded)
    - gradeSelectedRows / gradeAllRows
    - gradeRows_ (core loop with progress toasts)
    - runCriteria_ (LLM checks)
 5. LLM ENGINE
    - buildRubricPrompt_
    - geminiGrade_ / openaiGrade_
    - callResponsesStructured_ (OpenAI with 3-tier fallback)
    - extractGeminiText_ / extractResponsesText_
 6. CODE.ORG FETCH
    - extractChannelId_ / fetchGameLabSource_
 7. EMAIL
    - emailSelectedRows / sendEmailForRow_
 8. FORM INTEGRATION
    - onFormSubmit (trigger: append + grade + email)
    - importFormResponses_ (silent import helper)
    - gradeAndEmailAllNew (one-click workflow)
 9. DIAGNOSTICS
    - testAPIConnection (combined basic + structured test)
10. UTILITIES
    - Sheet helpers, CSV parser, JSON normalization
    - Cache helpers, header mapping, levelIdToUrl_
11. HELP DIALOG
12. EMBEDDED CRITERIA CSV
```

---

## 4. Key Design Decisions

| # | Decision | Rationale |
|---|---|---|
| 1 | Single Submissions sheet, formula-driven Grade Views | Avoids data duplication; Grade Views update automatically |
| 2 | LevelIDs are long-form (`Lesson-03-Level-08`) | Unambiguous, self-documenting, matches criteria CSV |
| 3 | `onFormSubmit` for real-time grading | Immediate feedback to students; "Grade New Submissions" as catch-up |
| 4 | Form import built into "Grade New Submissions" | One button does everything; eliminates "sync" confusion |
| 5 | Selection actions require Submissions sheet active | Prevents confusing "no rows selected" when on wrong sheet |
| 6 | Setup dialog is additive (never overwrites) | Safe to re-run; Reset Everything is separate and explicit |
| 7 | LevelURL column on Levels sheet | Teacher convenience — click to jump to the Code.org level |
| 8 | Combined API test (basic + structured) | Fewer menu items; catches both connection and JSON issues |
| 9 | Button feedback in setup dialog | Buttons disable + show status text while server-side code runs |
| 10 | CSS grid for period checkboxes | Clean layout regardless of how many periods (no singleton rows) |

---

## 5. Changes from Legacy (v1 → v2)

### Structural
- **Single `Submissions` sheet** replaces per-period data sheets
- **Formula-driven Grade View P#** sheets replace manual copy/paste workflows
- **Setup wizard** with HTML dialog and period picker (was a flat `setupSheets()` call)
- **LevelIDs cleaned** — removed legacy short IDs (L3-08) and non-standard levels (adv, a, c variants)
- **LevelURL column** replaces LevelName on Levels sheet (clickable Code.org links)

### Menu
- **"Grade New Submissions"** now auto-imports from form first (merged old "Sync from Form Responses")
- **"Sync, Grade & Email"** → **"Grade & Email All New"** (clearer name, same function)
- **"Grade Selected Rows"** → **"Re-grade Selected Rows"** (accurate — rows in Submissions were already graded)
- **"Test API Connection"** and **"Test Structured JSON"** merged into one combined test
- **"Sync from Form Responses"** removed (functionality absorbed into Grade New Submissions)

### UX
- **Toast notifications** on all long-running operations ("Grading 5 row(s)…", "Sending emails…")
- **Setup dialog buttons disable** while server-side code runs, with status message
- **Active-sheet validation** on selection-based actions (prompts to switch to Submissions if on wrong sheet)
- **In-app Help dialog** with complete setup instructions and menu reference

### Technical
- Grade caching via `CacheService` (6-hour TTL)
- Conditional formatting on Status column (green/red/yellow)
- LevelID dropdown validation in Submissions
- `importFormResponses_()` as a reusable silent helper (returns count, no alerts)
