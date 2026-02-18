# 🎮 Game Lab Autograder v2

Automatically grades Code.org Game Lab student projects using AI (Gemini or OpenAI).

Students submit their share links via a Google Form. The autograder fetches their code, evaluates it against rubric criteria using an LLM, writes the score to a spreadsheet, and emails the student their results — all automatically.

---

## 🚀 Quick Setup (5 steps)

### 1. Create the Google Sheet

- Create a new Google Sheet (or open an existing one)
- Go to **Extensions → Apps Script**
- Delete any existing code in `Code.gs`
- Paste the entire contents of `Code.gs` from this repo
- Click **Save** (💾)
- Close the Apps Script editor and **reload the spreadsheet**

### 2. Run Initial Setup

- In the spreadsheet, click **Autograder → Initial Setup…**
- Check the class periods you teach (1–8)
- Click **Create Sheets**

This creates:
| Sheet | Purpose |
|---|---|
| **Submissions** | All student submissions and grades |
| **Levels** | 16 Game Lab levels — enable/disable, set model overrides |
| **Criteria** | 61 rubric criteria (auto-populated from the built-in CSV) |
| **Grade View P#** | One per period — read-only, sorted by level then last name |

### 3. Set your API Key

- Go to **Extensions → Apps Script → ⚙️ Project Settings → Script Properties**
- Click **Add script property**
- Name: `GEMINI_API_KEY`
- Value: your Gemini API key (free at [aistudio.google.com](https://aistudio.google.com))

> **Optional:** To use OpenAI instead, add two properties:
> - `LLM_PROVIDER` = `openai`
> - `OPENAI_API_KEY` = your OpenAI key

### 4. Test your connection

- Click **Autograder → Test API Connection**
- You should see ✅ and a "pong" response
- Also try **Test Structured JSON** to verify JSON grading works

### 5. Create & Link a Google Form

Create a Google Form with these fields:

| Field | Type | Notes |
|---|---|---|
| Email Address | Built-in setting | Form Settings → Collect email addresses |
| First Name | Short answer | |
| Last Name | Short answer | |
| Class Period | Dropdown | 1, 2, 3, 4, 5, 6, 7, 8 (match your periods) |
| Assessment Level | Dropdown | Copy level IDs from the Levels sheet |
| Share URL | Short answer | Students paste their Code.org share link |

Then link it to your spreadsheet:
1. In the Form editor → **Responses** tab → click the green **Sheets** icon
2. Choose **Select existing spreadsheet** → pick your autograder spreadsheet
3. This creates a "Form Responses 1" sheet

Finally, set up the auto-grade trigger:
1. In **Extensions → Apps Script**, click the ⏰ **Triggers** icon (left sidebar)
2. Click **+ Add Trigger**
3. Function: `onFormSubmit` | Source: **From spreadsheet** | Event: **On form submit**
4. Leave "Which deployment should run" set to **Head**
5. Click **Save** and authorize when prompted

**Done!** When a student submits, their code is automatically graded and emailed.

> **📧 Note on emails:** Student result emails work automatically — no extensions or extra setup needed. The script uses Google's built-in `GmailApp` service, which is authorized when you approve the trigger. Gmail limits: ~100 emails/day (consumer) or ~1500/day (Google Workspace).

---

## 📋 Menu Reference

| Menu Item | What it does |
|---|---|
| **Initial Setup…** | Creates/configures all sheets (additive — won't overwrite existing) |
| **Grade New Submissions** | Grades all rows where Score is blank |
| **Grade Selected Rows** | Re-grades only the rows you've highlighted |
| **Re-grade All Rows…** | Re-grades every submission (slow, uses API credits) |
| **Email Selected Rows** | Sends result emails for highlighted rows |
| **Sync from Form Responses** | Imports any submissions that didn't auto-import (safety net) |
| **Sync, Grade & Email** | Sync + grade + email in one step |
| **Test API Connection** | Sends a "ping" to verify your API key works |
| **Test Structured JSON** | Verifies the LLM can return structured grading JSON |
| **Help / Setup Guide** | Opens the in-app help dialog |

---

## 📄 Sheet Reference

### Submissions
The main data sheet. One row per submission.

| Column | Description |
|---|---|
| Timestamp | When the form was submitted |
| First | Student's first name |
| Last | Student's last name |
| Period | Class period (1–8) |
| Email | Student's email address |
| LevelID | Which level they're submitting (e.g., `Lesson-03-Level-08`) |
| ShareURL | Their Code.org share link |
| ChannelID | Auto-extracted from the share link |
| Score | Points earned |
| MaxScore | Total possible points for that level |
| Status | `OK`, `Error`, `Invalid share link`, etc. |
| Notes | Per-criterion ✅/❌ breakdown |
| EmailedAt | Timestamp when the result email was sent |

### Grade View P#
Read-only formula sheets (one per period). Automatically filters and sorts Submissions by period, then by level and last name. **Protected** — don't edit these directly.

### Levels
Lists all 16 Game Lab levels. You can:
- **Uncheck Enabled** to skip grading for a level
- **Set a Model override** (e.g., `gemini-2.0-flash-lite`) for a specific level

### Criteria
The rubric. 61 rows across 16 levels. Auto-populated from the built-in CSV. You can edit descriptions or point values to customize grading.

---

## 🧪 Supported Levels

| LevelID | Criteria Count |
|---|---|
| Lesson-03-Level-08 | 4 |
| Lesson-04-Level-08 | 2 |
| Lesson-05-Level-07 | 5 |
| Lesson-06-Level-07 | 4 |
| Lesson-08-Level-10 | 3 |
| Lesson-09-Level-05 | 3 |
| Lesson-10-Level-05 | 3 |
| Lesson-12-Level-07 | 4 |
| Lesson-13-Level-07 | 4 |
| Lesson-15-Level-07 | 1 |
| Lesson-16-Level-06 | 4 |
| Lesson-17-Level-07 | 3 |
| Lesson-19-Level-09 | 4 |
| Lesson-20-Level-07 | 1 |
| Lesson-21-Side-Scroller | 16 |
| Lesson-22-Level-06 | 1 |

---

## ⚙️ Script Properties Reference

Set these in **Extensions → Apps Script → ⚙️ Project Settings → Script Properties**:

| Property | Required | Description |
|---|---|---|
| `GEMINI_API_KEY` | Yes (default) | Your Gemini API key |
| `OPENAI_API_KEY` | If using OpenAI | Your OpenAI API key |
| `LLM_PROVIDER` | No | `gemini` (default) or `openai` |

---

## 🔄 How Grading Works

1. Student submits a Google Form with their share link and level
2. `onFormSubmit` trigger fires → copies data to Submissions sheet
3. Channel ID is extracted from the share URL
4. Student source code is fetched from `studio.code.org`
5. A cache key (SHA-256 of LevelID + source) is checked
6. If not cached: rubric criteria are sent to the LLM with the source code
7. LLM returns JSON with pass/fail for each criterion
8. Score is calculated and written to the row
9. Result is cached for 6 hours
10. A results email is sent to the student

---

## 🛠 Troubleshooting

**"Missing GEMINI_API_KEY"** — You haven't set the script property. Go to Extensions → Apps Script → ⚙️ Project Settings → Script Properties.

**"Invalid share link"** — The student's URL doesn't match the expected `studio.code.org/projects/gamelab/...` pattern. Have them re-copy the share link.

**"Level disabled/unknown"** — The LevelID in the submission doesn't match any enabled level in the Levels sheet. Check for typos.

**Submissions aren't auto-importing** — Make sure the `onFormSubmit` trigger is set up (Extensions → Apps Script → Triggers). Use "Sync from Form Responses" as a fallback.

**Students not receiving emails** — Check that the Email column has valid addresses and that EmailedAt is blank. Gmail has a daily sending limit (~100/day for consumer, ~1500/day for Workspace).

---

## 📝 License

This project is provided as-is for educational use.
