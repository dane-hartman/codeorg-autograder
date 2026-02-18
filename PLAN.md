# Game Lab Autograder v2 — Detailed Plan

## 1. Project Overview

**Goal:** A Google Apps Script (`.gs`) that lives inside a Google Sheets spreadsheet. Teachers connect a Google Form where students submit their Code.org Game Lab share links. The script fetches the student source code, sends it to an LLM (Gemini or OpenAI) along with rubric criteria, and writes the score + notes back to the spreadsheet. Students optionally receive an automated email with their results.

### What stays the same from v1
- Core grading pipeline: extract channel ID → fetch source from Code.org → build rubric prompt → call LLM → parse structured JSON → write score/notes
- Dual LLM provider support (Gemini default, OpenAI optional)
- Robust JSON normalization with multiple fallback strategies
- Email workflow via `GmailApp`
- `onFormSubmit` trigger for automatic grading on new form submissions
- Criteria stored in a **Criteria** sheet (populated from the embedded CSV)
- Levels stored in a **Levels** sheet (enable/disable, per-level model override)

---

## 2. Architecture: Sheets & Data Flow

### 2.1 — Sheet Inventory

| Sheet | Created by | Purpose |
|---|---|---|
| `Form Responses 1` | Google Forms (automatic) | Raw form submissions land here. **Never edited by the script.** |
| `Submissions` | Setup wizard | Normalized copy of every submission. All grading happens here. Has a `Period` column. |
| `Grade View P1` … `Grade View P8` | Setup wizard (one per checked period) | **Read-only views.** Auto-populated from `Submissions` using formulas. Sorted by LevelID then Last name. |
| `Levels` | Setup wizard | Master level list — 16 rows. Enabled checkbox, optional per-level Model override. |
| `Criteria` | Setup wizard | Master rubric — all criterion rows from the embedded CSV. |

### 2.2 — Data Flow

```
Student submits Google Form
        │
        ▼
  ┌─────────────────────┐
  │  Form Responses 1   │   (raw, untouched)
  └────────┬────────────┘
           │  onFormSubmit trigger OR manual "Import from Form Responses"
           ▼
  ┌─────────────────────┐
  │    Submissions       │   (normalized row appended; grading writes Score/Notes here)
  └────────┬────────────┘
           │  Grade View sheets pull from Submissions via formulas
           ▼
  ┌─────────────────────┐
  │  Grade View P3       │   (formula-driven: filters Period=3, sorts by LevelID → Last)
  │  Grade View P4       │   (formula-driven: filters Period=4, sorts by LevelID → Last)
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

All submissions live here regardless of period. The `Period` column is populated from the form response.

### 2.4 — Grade View Sheets (Formula-Driven, Protected)

Each `Grade View P#` sheet uses a single `SORT(FILTER(...))` formula in cell A2 that:
1. Filters `Submissions` rows where `Period` matches that sheet's period number
2. Selects columns in this display order (optimized for scanning: assignment → student → grade → details):
3. Sorts by `LevelID` ascending, then `Last` ascending

Header row (row 1) is static and bold:
```
LevelID | First | Last | Score | MaxScore | Status | Email | ShareURL | Timestamp | Notes
```

Column order rationale: The teacher's eye naturally scans left-to-right. The most important info (which level? whose? what score?) is leftmost. Lower-priority reference info (email, link, timestamp, detailed notes) is on the right and can be scrolled to when needed.

These sheets are **protected** (locked via `sheet.protect()`) so the teacher can't accidentally edit formulas. The protection warning says: *"This sheet is auto-generated. To change grades, edit the Submissions sheet."*

They update automatically as `Submissions` gets graded — no action needed.

---

## 3. Changes from Legacy (v1 → v2)

### 3.1 — Setup Wizard with Period Picker

**Legacy:** `setupSheets()` creates three fixed tabs: `Submissions`, `Levels`, `Criteria`.

**v2:** When the teacher first runs **Autograder → Initial Setup…**, an HTML dialog prompts:
> "Which class periods do you teach? (Check all that apply)"
>
> ☑ Period 1 &nbsp; ☑ Period 2 &nbsp; ☑ Period 3 &nbsp; ☐ Period 4 &nbsp; ☐ Period 5 &nbsp; ☑ Period 6 &nbsp; ☐ Period 7 &nbsp; ☐ Period 8

The script then creates:
- `Submissions` sheet (single, all periods)
- `Levels` sheet (16 levels from CSV)
- `Criteria` sheet (all rubric rows from CSV)
- `Grade View P1`, `Grade View P2`, `Grade View P3`, `Grade View P6` (one per checked period)

**Re-running setup:**
- If `Levels`/`Criteria`/`Submissions` already exist, they are **not touched**.
- The dialog shows which period Grade View sheets already exist (greyed out with ✓).
- The teacher can check additional periods to add, or click a separate **"Reset Everything"** button that wipes all data and starts fresh (with confirmation).

### 3.2 — Levels Sheet Matches CSV (Cleaned)

**Legacy embedded CSV** used short IDs (`L3-08`, `L5-06c`, `L6-07adv`, etc.).

**v2** uses the IDs from `criteria-table.csv` — the authoritative, already-cleaned version:

| LevelID | LevelName | Enabled | Model |
|---|---|---|---|
| Lesson-03-Level-08 | Lesson 3 Level 8 | TRUE | |
| Lesson-04-Level-08 | Lesson 4 Level 8 | TRUE | |
| Lesson-05-Level-07 | Lesson 5 Level 7 | TRUE | |
| Lesson-06-Level-07 | Lesson 6 Level 7 | TRUE | |
| Lesson-08-Level-10 | Lesson 8 Level 10 | TRUE | |
| Lesson-09-Level-05 | Lesson 9 Level 5 | TRUE | |
| Lesson-10-Level-05 | Lesson 10 Level 5 | TRUE | |
| Lesson-12-Level-07 | Lesson 12 Level 7 | TRUE | |
| Lesson-13-Level-07 | Lesson 13 Level 7 | TRUE | |
| Lesson-15-Level-07 | Lesson 15 Level 7 | TRUE | |
| Lesson-16-Level-06 | Lesson 16 Level 6 | TRUE | |
| Lesson-17-Level-07 | Lesson 17 Level 7 | TRUE | |
| Lesson-19-Level-09 | Lesson 19 Level 9 | TRUE | |
| Lesson-20-Level-07 | Lesson 20 Level 7 | TRUE | |
| Lesson-21-Side-Scroller | Lesson 21 Side Scroller | TRUE | |
| Lesson-22-Level-06 | Lesson 22 Level 6 | TRUE | |

**Removed levels** (were in legacy embedded CSV only):
- `L5-06c` (bear variables — "c" level)
- `L6-07adv` (caterpillar advanced)
- `L6-08a` (concentric circles — "a" level)
- `L9-05adv` (food tray advanced)
- `L12-07adv` (salt shaker advanced)

The Google Form's "Which assessment level are you submitting?" dropdown should list the exact LevelIDs (e.g., `Lesson-03-Level-08`). No mapping needed — the IDs flow through the system unchanged.

### 3.3 — Cleaner Autograder Menu

**Legacy menu** had nested sub-menus (`Emails >`, `Admin >`, `Diagnostics >`) that were hard to navigate.

**v2 menu** — flat, clearly labeled, grouped with separators:

```
📋 Autograder
─────────────────────────────────────
  ▶ Initial Setup…                    (first-time wizard with period picker)
─────────────────────────────────────
  ▶ Grade New Submissions             (rows where Score is blank)
  ▶ Grade Selected Rows               (highlight rows first)
  ▶ Re-grade All Rows…                (confirmation prompt — slow)
─────────────────────────────────────
  ▶ Email Selected Rows               (sends result email)
  ▶ Sync from Form Responses          (safety net — imports any missed submissions)
  ▶ Sync, Grade & Email               (sync + grade + email in one step)
─────────────────────────────────────
  ▶ Test API Connection               (ping current LLM provider)
  ▶ Test Structured JSON              (structured output test)
─────────────────────────────────────
  ▶ Help / Setup Guide                (modal with full instructions)
```

All items are top-level — no hunting through sub-menus. Names are action-oriented and jargon-free.

### 3.4 — In-App Help & Setup Guide

The **Help / Setup Guide** modal will contain complete, step-by-step instructions:

1. **Run "Initial Setup…"** from the Autograder menu to create sheets. Pick your class periods.

2. **Set your API key:**
   - Go to **Extensions → Apps Script → ⚙️ Project Settings → Script Properties**
   - Add property: `GEMINI_API_KEY` = your key (get one free at [aistudio.google.com](https://aistudio.google.com))
   - *(Optional)* Add: `LLM_PROVIDER` = `openai` and `OPENAI_API_KEY` = your key
   - Click **"Test API Connection"** from the Autograder menu to verify.

3. **Create a Google Form** with these fields:
   - Email Address (Settings → "Collect email addresses": Responder input)
   - First Name (Short answer)
   - Last Name (Short answer)
   - Class Period (Dropdown: 1, 2, 3, … matching the periods you chose in setup)
   - Which assessment level are you submitting? (Dropdown: `Lesson-03-Level-08`, `Lesson-04-Level-08`, … matching the LevelIDs in the Levels sheet)
   - Paste the share URL to your completed assessment level (Short answer)

4. **Link the form to this spreadsheet:**
   - In the Google Form editor → **Responses** tab → click the green Sheets icon
   - Choose **"Select existing spreadsheet"** → pick this spreadsheet
   - This creates a `Form Responses 1` tab automatically

5. **Set up automatic grading on form submit:**
   - In the Apps Script editor (Extensions → Apps Script), click the ⏰ **Triggers** icon (left sidebar)
   - Click **+ Add Trigger**
   - Choose function: `onFormSubmit`
   - Event source: **From spreadsheet**
   - Event type: **On form submit**
   - Click **Save** and authorize when prompted

6. **Automatic emails:**
   - Emails are sent automatically when a submission is graded (if the student provided an email).
   - The first time emails are sent, Google will ask you to authorize Gmail access — click Allow.
   - Students receive their score, a ✅/❌ breakdown of each criterion, and a link to their project.

### 3.5 — `onFormSubmit` Flow

When a form is submitted:
1. Read the raw form response values from `Form Responses 1`
2. Map verbose column headers to internal names via `headersSmart_()`
3. Build a normalized row (Timestamp, First, Last, Period, Email, LevelID, ShareURL)
4. Append the row to the `Submissions` sheet
5. Grade the new row (fetch Code.org source → LLM → write Score/Notes)
6. Send email to the student (if Email column has a value)
7. The appropriate `Grade View P#` sheet updates automatically via its formula

### 3.6 — Backfill ("Sync from Form Responses")

**When is this needed?** Only as a safety net. The `onFormSubmit` trigger handles every incoming submission automatically. But if the trigger temporarily fails (Apps Script quota hit, API key expired, network blip), some form responses may sit in `Form Responses 1` without ever making it into `Submissions`.

**What it does:**
1. Reads all rows from `Form Responses 1`
2. Compares against `Submissions` using a dedup key: `Timestamp + Email + LevelID`
3. Any rows in `Form Responses 1` that are NOT already in `Submissions` get appended and graded
4. Reports how many new rows were synced

**Menu label:** `Sync from Form Responses` (clearer than "backfill" for a teacher audience). There's also a `Sync, Grade & Email` variant.

**Philosophy:** This should rarely be needed. The menu item exists as a "just in case" recovery tool rather than a primary workflow step. The help dialog will note: *"You normally don't need this — submissions are graded automatically when submitted. Use this only if you notice missing submissions."*

### 3.6 — Misc Improvements

- **LevelName auto-populated** from the LevelID (e.g., `Lesson-03-Level-08` → `Lesson 3 Level 8`)
- **Grade cache** via `CacheService` with 6-hour TTL. Cache key = SHA-256 of (LevelID + source code). If the student's code hasn't changed since last grading, the cached result is used — saving API credits and time.
- **Embedded CSV updated** to match `criteria-table.csv` exactly (long IDs, no adv/a/c levels)
- **Model default:** `gemini-2.0-flash` (fast, cheap, high quality as of 2026). Teachers can override per-level in the `Levels` sheet `Model` column.
- **Progress toast** shown while grading ("Grading row 3 of 12…") so the teacher knows it's working
- **Conditional formatting** applied to Status column: green for OK, red for Error, yellow for others
- **LevelID dropdown validation** in Submissions referencing the Levels sheet

---

## 4. File Structure (Deliverables)

```
game-lab-autograder/
├── PLAN.md                          ← this file
├── criteria-table.csv               ← authoritative rubric (unchanged)
├── Code.gs                          ← the new v2 Apps Script
├── README.md                        ← teacher-facing setup & usage guide
├── Code-legacy.gs                   ← kept for reference
├── autograder-google-sheet-legacy.xlsx
└── google-form-screenshot.png
```

### `Code.gs` — Module Layout

```
 1. CONFIG & CONSTANTS
 2. MENU (onOpen)
 3. SETUP WIZARD
    a. setupSheets (entry point — shows period picker dialog)
    b. createSheetsFromSetup_ (called from dialog: builds Submissions, Levels, Criteria, Grade Views)
    c. resetEverything_ (wipe all sheets and re-run)
    d. buildGradeViewFormula_ (SORT/FILTER formula for each Grade View P# sheet)
 4. GRADING ENGINE
    a. gradeNewRows / gradeSelectedRows / gradeAllRows
    b. gradeRows_ (core loop with progress toasts)
    c. runCriteria_ (local checks + LLM checks)
 5. LLM ENGINE
    a. buildRubricPrompt_
    b. geminiGrade_ / openaiGrade_
    c. callResponsesStructured_ (OpenAI Responses API with fallback)
    d. extractGeminiText_ / extractResponsesText_
    e. normalizeAutogradeJson_
 6. CODE.ORG FETCH
    a. extractChannelId_
    b. fetchGameLabSource_
 7. EMAIL
    a. sendEmailForRow_
    b. emailSelectedRows
 8. FORM INTEGRATION
    a. onFormSubmit (appends to Submissions, grades, emails)
    b. syncFromFormResponses / syncGradeAndEmail (safety-net recovery)
 9. DIAGNOSTICS
    a. pingGemini_ / pingGPT
    b. pingGeminiStructured_ / pingGPTStructured
10. UTILITIES
    a. getSheet_, headers_, writeRow_, headersSmart_
    b. parseCsvText_, getCriteriaTableCsvText_
    c. esc_, stripCodeFences_, toBool_, normalizeAutogradeJson_
    d. Cache helpers (CacheService wrappers)
11. HELP DIALOG (showAutograderHelp)
```

---

## 5. Detailed Behavior Specifications

### 5.1 — Setup Wizard Flow

```
Teacher clicks: Autograder → Initial Setup…
  │
  ├─ Show HTML dialog:
  │   ┌───────────────────────────────────────────┐
  │   │  Game Lab Autograder — Initial Setup       │
  │   │                                            │
  │   │  Select class periods to create            │
  │   │  Grade View sheets for:                    │
  │   │                                            │
  │   │  ☑ Period 1   ☑ Period 2   ☐ Period 3     │
  │   │  ☐ Period 4   ☐ Period 5   ☐ Period 6     │
  │   │  ☐ Period 7   ☐ Period 8                   │
  │   │                                            │
  │   │  (Periods with ✓ already exist and will    │
  │   │   not be modified.)                        │
  │   │                                            │
  │   │  [ Create Sheets ]  [ Reset Everything ]   │
  │   └───────────────────────────────────────────┘
  │
  ├─ [Create Sheets] clicked:
  │   ├─ If Submissions/Levels/Criteria don't exist → create them
  │   ├─ For each newly-checked period → create "Grade View P#" with formula
  │   ├─ Already-existing Grade View sheets → skip (not touched)
  │   ├─ Apply formatting (bold headers, frozen rows, column widths)
  │   └─ Show success alert with next-steps checklist
  │
  └─ [Reset Everything] clicked:
      ├─ Confirmation: "This will delete ALL autograder sheets and data. Continue?"
      ├─ Deletes: Submissions, Levels, Criteria, all Grade View P# sheets
      └─ Re-opens the setup dialog (fresh state)
```

### 5.2 — Grade View Formula (per sheet)

Each `Grade View P#` sheet has a single array formula in cell **A2**:

We select & reorder columns so the Grade View shows:

| Grade View Col | Source (Submissions col) |
|---|---|
| LevelID | `Submissions!F` (LevelID) |
| First | `Submissions!B` (First) |
| Last | `Submissions!C` (Last) |
| Score | `Submissions!H` (Score) |
| MaxScore | `Submissions!I` (MaxScore) |
| Status | `Submissions!J` (Status) |
| Email | `Submissions!E` (Email) |
| ShareURL | `Submissions!G` (ShareURL) |
| Timestamp | `Submissions!A` (Timestamp) |
| Notes | `Submissions!K` (Notes) |

Filter: `Submissions!D = <period number>`
Sort: LevelID ascending (col 1 of result), then Last ascending (col 3 of result)

The formula will be built in code as a `SORT(FILTER({col,col,...}, condition))` using column references.

The sheet is **protected** — the teacher sees a warning if they try to edit: *"This sheet is auto-generated from the Submissions sheet."*

### 5.3 — Grading Flow (per row)

```
1. Read LevelID and ShareURL from the Submissions row
2. Validate both are present → else write Status="No URL/LevelID"
3. Check level is enabled in Levels sheet → else write Status="Level disabled"
4. Load criteria for that LevelID from Criteria sheet
5. Extract channel ID from share URL via regex
6. Check cache: SHA-256(LevelID + source). If hit → write cached result, skip LLM call
7. Fetch source code from Code.org API
8. Run local checks (code_nonempty, contains, regex_present, regex_absent)
9. Build LLM prompt with remaining llm_check criteria
10. Call LLM (Gemini or OpenAI) with structured output request
11. Parse JSON response, normalize to {checks:[{id, pass, reason}]}
12. Calculate score = sum of passed criteria points
13. Write ChannelID, Score, MaxScore, Status, Notes to the row
14. Cache the result (6-hour TTL)
15. (If auto-email enabled) Send email to student
```

### 5.4 — `onFormSubmit` Flow

```
1. Read the form response values from Form Responses 1
2. Map verbose column headers to internal names via headersSmart_()
3. Build a normalized row: Timestamp, First, Last, Period, Email, LevelID, ShareURL
4. Append the row to the Submissions sheet
5. Grade the new row
6. Send email to student (if Email present)
7. Grade View P# sheet auto-updates via formula (no action needed)
```

### 5.5 — Sync from Form Responses (Safety Net)

```
1. Read all rows from Form Responses 1
2. Read all rows from Submissions
3. Build dedup key set from Submissions: Timestamp + Email + LevelID
4. For each Form Response row NOT in Submissions → append normalized row
5. Grade all newly-appended rows
6. (If "Sync, Grade & Email" variant) → also send emails
7. Report: "Synced X new submissions."
```

This is a recovery tool, not a primary workflow. The Help dialog explains when to use it.

### 5.5 — Email Content

Same as legacy — HTML + plain text fallback with:
- Greeting with student name
- Level and score
- Link to their project
- Per-criterion ✅/❌ breakdown
- Footer note that it was auto-generated

---

## 6. Decision Log

All clarifying questions have been resolved. Final decisions:

| # | Question | Decision |
|---|---|---|
| Q1 | Period sheets vs. single Submissions | **Single `Submissions` sheet** for all data. `Grade View P#` sheets are formula-driven read-only views. |
| Q2 | LevelID format | Use exact LevelID everywhere (`Lesson-03-Level-08`). No mapping needed. |
| Q3 | Period range | **Periods 1–8** (numbered). |
| Q4 | Default model | **`gemini-2.0-flash`** (fast, cheap, high quality). |
| Q5 | Re-running setup | **Additive** (add missing Grade View sheets without touching existing ones) + option to **Reset Everything**. |
| Q6 | Caching | **Yes**, via `CacheService` with 6-hour TTL. Key = SHA-256(LevelID + source code). |
| Q7 | Form fields | Timestamp, Email, First, Last, Period, LevelID, ShareURL — complete. |
| Q8 | Grade View protection | **Yes**, sheets are protected with a descriptive warning message. |
| Q9 | Grade View columns | `LevelID, First, Last, Score, MaxScore, Status, Email, ShareURL, Timestamp, Notes` (name+grade leftmost). |
| Q10 | Backfill | Renamed to **"Sync from Form Responses"**. Safety-net only — rarely needed since `onFormSubmit` handles the normal flow. |

---

## 7. Next Steps

Ready to build. Deliverables:

1. **`Code.gs`** — the complete v2 Apps Script (~1200–1500 lines)
2. **`README.md`** — teacher-facing setup & usage guide
3. **`criteria-table.csv`** — stays as-is (already correct)
