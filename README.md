# 🎮 Game Lab Autograder

Automatically grades [Code.org Game Lab](https://code.org/educate/gamelab) student projects using AI (Gemini or OpenAI).

Students submit their share links via a Google Form. The autograder fetches their code, evaluates it against rubric criteria using an LLM, writes the score to a spreadsheet, and emails students their results — all automatically.

Built for the **CSD Unit 3 (Interactive Animations and Games)** curriculum.

---

## 🚀 Quick Setup

### 1. Create the Google Sheet

- Create a new Google Sheet (or open an existing one)
- Go to **Extensions → Apps Script**
- Delete any existing code in `Code.gs`
- Paste the entire contents of [`Code.gs`](Code.gs) from this repo
- Click **Save** (💾)
- Close the Apps Script editor and **reload the spreadsheet**

### 2. Run Initial Setup

- In the spreadsheet, click **Autograder → Initial Setup…**
- Check the class periods you teach (1–8)
- Click **Create Sheets**

This creates:

| Sheet | Purpose |
|---|---|
| **Submissions** | All student submissions and grades (one row per submission) |
| **Levels** | 16 Game Lab levels with Code.org links — enable/disable, set model overrides |
| **Criteria** | 61 rubric criteria (auto-populated from the built-in CSV) |
| **Grade View P#** | One per period — read-only views sorted by level then last name |

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
- This runs two tests: a basic connectivity check and a structured JSON grading test
- You should see ✅ for both

### 5. Create & Link a Google Form

Create a Google Form with these fields:

| Field | Type | Notes |
|---|---|---|
| Email Address | Built-in setting | Form Settings → Collect email addresses |
| First Name | Short answer | |
| Last Name | Short answer | |
| Class Period | Dropdown | 1, 2, 3, 4, 5, 6, 7, 8 (match your periods) |
| Assessment Level | Dropdown | Copy LevelIDs from the Levels sheet |
| Share URL | Short answer | Students paste their Code.org share link |

Then link it to your spreadsheet:

1. In the Form editor → **Responses** tab → click the green **Sheets** icon
2. Choose **Select existing spreadsheet** → pick your autograder spreadsheet
3. This creates a "Form Responses 1" sheet

Finally, set up the auto-grade trigger:

1. In **Extensions → Apps Script**, click the ⏰ **Triggers** icon (left sidebar)
2. Click **+ Add Trigger**
3. Function: `onFormSubmit` | Source: **From spreadsheet** | Event: **On form submit**
4. Leave deployment set to **Head**
5. Click **Save** and authorize when prompted

**Done!** When a student submits, their code is automatically graded and emailed.

> **📧 Note on emails:** Student result emails use Google's built-in `GmailApp` service, authorized when you approve the trigger. Gmail limits: ~100 emails/day (consumer) or ~1,500/day (Google Workspace).

---

## 📋 Menu Reference

All menu actions operate on the **Submissions** sheet — never on Form Responses directly.

| Menu Item | What it does |
|---|---|
| **Initial Setup…** | Creates/configures all sheets (additive — won't overwrite existing) |
| **Grade New Submissions** | Imports any new form responses into Submissions, then grades all ungraded rows |
| **Re-grade Selected Rows** | Re-grades only the rows you highlight in Submissions (e.g., after editing criteria) |
| **Re-grade All Rows…** | Re-grades every submission (slow, uses API credits) |
| **Grade & Email All New** | Imports, grades, and emails results for all un-emailed students in one step |
| **Email Selected Rows** | Sends result emails for rows you highlight in Submissions |
| **Test API Connection** | Verifies API key and structured JSON grading in one combined test |
| **Help / Setup Guide** | Opens the in-app help dialog |

---

## 📄 Sheet Reference

### Submissions

The main data sheet. One row per submission. All grading reads/writes happen here.

| Column | Description |
|---|---|
| Timestamp | When the form was submitted |
| First | Student's first name |
| Last | Student's last name |
| Period | Class period (1–8) |
| Email | Student's email address |
| LevelID | Which level they submitted (e.g., `Lesson-03-Level-08`) |
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

Lists all 16 Game Lab levels with direct links to Code.org. You can:
- **Uncheck Enabled** to skip grading for a level
- **Set a Model override** (e.g., `gemini-2.0-flash-lite`) for a specific level

### Criteria

The rubric. 61 rows across 16 levels. Auto-populated from the built-in CSV. You can edit descriptions or point values to customize grading.

---

## 🧪 Supported Levels

| LevelID | Criteria | Description |
|---|---|---|
| Lesson-03-Level-08 | 4 | Purple rect on top of ellipses |
| Lesson-04-Level-08 | 2 | Cloud wider than tall |
| Lesson-05-Level-07 | 5 | Both eyes use eyeSize variable |
| Lesson-06-Level-07 | 4 | Complete the caterpillar |
| Lesson-08-Level-10 | 3 | Sprite animations |
| Lesson-09-Level-05 | 3 | Shrink the food |
| Lesson-10-Level-05 | 3 | Adding text |
| Lesson-12-Level-07 | 4 | The draw loop |
| Lesson-13-Level-07 | 4 | Swimming fish |
| Lesson-15-Level-07 | 1 | Transforming dinosaur |
| Lesson-16-Level-06 | 4 | Flyer movement controls |
| Lesson-17-Level-07 | 3 | Shake the creature |
| Lesson-19-Level-09 | 4 | Fish with velocity |
| Lesson-20-Level-07 | 1 | Horse to unicorn |
| Lesson-21-Side-Scroller | 16 | Side-scroller game project |
| Lesson-22-Level-06 | 1 | Rock falls back down |

---

## ⚙️ Configuration

### Script Properties

Set in **Extensions → Apps Script → ⚙️ Project Settings → Script Properties**:

| Property | Required | Description |
|---|---|---|
| `GEMINI_API_KEY` | Yes (default) | Your Gemini API key ([get one free](https://aistudio.google.com)) |
| `OPENAI_API_KEY` | If using OpenAI | Your OpenAI API key |
| `LLM_PROVIDER` | No | `gemini` (default) or `openai` |

### Per-Level Model Overrides

In the **Levels** sheet, the **Model** column lets you override the default model for specific levels. For example, you could use a more capable model for the complex Side Scroller project while using a faster/cheaper model for simpler levels. Leave blank to use the default.

---

## 🔄 How Grading Works

1. Student submits a Google Form with their share link and level
2. `onFormSubmit` trigger fires → copies data to the Submissions sheet
3. Channel ID is extracted from the share URL
4. Student source code is fetched from `studio.code.org`
5. A cache key (SHA-256 of LevelID + source) is checked
6. If not cached: rubric criteria are sent to the LLM with the source code
7. LLM returns JSON with pass/fail for each criterion
8. Score is calculated and written to the row
9. Result is cached for 6 hours
10. A results email is sent to the student

**Grade New Submissions** automatically imports any missed form responses before grading, so it works as both a catch-up tool and a manual grade trigger.

### Criterion Types

| Type | How it works |
|---|---|
| `llm_check` | Sent to the LLM for evaluation (the majority of criteria) |
| `code_nonempty` | Local check: source code is ≥ 10 characters |
| `contains` | Local check: source contains a specific string |
| `regex_present` | Local check: source matches a regex pattern |
| `regex_absent` | Local check: source does NOT match a regex pattern |

---

## 🛠 Troubleshooting

| Problem | Solution |
|---|---|
| **"Missing GEMINI_API_KEY"** | Set the script property in Extensions → Apps Script → ⚙️ Project Settings → Script Properties |
| **"Invalid share link"** | Student's URL doesn't match `studio.code.org/projects/gamelab/...` — have them re-copy the share link |
| **"Level disabled/unknown"** | The LevelID doesn't match any enabled level in the Levels sheet — check for typos |
| **Submissions aren't auto-importing** | Verify the `onFormSubmit` trigger is set up (Extensions → Apps Script → Triggers). Use **Grade New Submissions** to catch up. |
| **Students not receiving emails** | Check that the Email column has valid addresses and EmailedAt is blank. Gmail has daily sending limits (~100/day consumer, ~1,500/day Workspace). |
| **"Please switch to the Submissions sheet"** | Selection-based actions (Re-grade Selected, Email Selected) require you to be on the Submissions sheet with rows highlighted |

---

## 🏗 Project Structure

```
game-lab-autograder/
├── Code.gs                  # The complete Apps Script — paste into your spreadsheet
├── criteria-table.csv       # Authoritative rubric data (16 levels, 61 criteria)
├── README.md                # This file
├── CHANGELOG.md             # Version history
├── PLAN.md                  # Architecture & design decisions (internal)
├── LICENSE                  # MIT License
├── .gitignore
└── legacy/                  # Original v1 files (kept for reference)
    ├── Code-legacy.gs
    ├── autograder-google-sheet-legacy.xlsx
    └── google-form-screenshot.png
```

---

## 🤝 Contributing

Contributions welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

If you're a teacher using this and have feedback, bug reports, or ideas for new levels/criteria, please [open an issue](../../issues).

---

## 📝 License

MIT — see [LICENSE](LICENSE). Built for teachers, by a teacher.
