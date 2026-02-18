# Contributing to Game Lab Autograder

Thanks for your interest in improving the autograder! This project is built by teachers for teachers.

## Ways to Contribute

### Report Issues
- **Bug reports:** Something not grading correctly? Share the LevelID and describe what happened.
- **Grading accuracy:** If a criterion is consistently too strict or too lenient, open an issue with examples.
- **Setup problems:** Describe your environment (school Google Workspace vs. personal Gmail, etc.).
- **Note:** This project is maintained in my spare time. I may not respond quickly to issues or review all pull requests. Feel free to fork and adapt for your own needs!

### Suggest Improvements
- New levels or criteria for existing CSD Unit 3 lessons
- UX improvements to the menu, dialogs, or email templates
- Documentation improvements

### Submit Code
1. Fork the repository
2. Make your changes to `Code.gs` (for script changes) or add/edit CSV files in `criteria/` (for rubric changes)
3. Test in a real Google Sheet (paste into Apps Script, run Initial Setup, import criteria CSV, grade a few submissions)
4. Open a pull request with a clear description of what changed and why

## Understanding the Criteria Workflow

Criteria live in **CSV files** in the `criteria/` folder. Teachers import them into the **Criteria sheet** via Google Sheets' built-in File → Import feature.

| Location | Purpose |
|---|---|
| `criteria/*.csv` (files in repo) | Shareable rubric definitions — one CSV per curriculum unit |
| **Criteria sheet** (in Google Sheets) | **Runtime source of truth** — the grading engine reads from here |
| **Levels sheet** (in Google Sheets) | Auto-generated from Criteria via "Sync Levels from Criteria" |

**How it works:**
- Teachers import a CSV into the Criteria sheet (File → Import → Upload → "Replace current sheet").
- They run **Sync Levels from Criteria** to rebuild the Levels sheet from the LevelIDs found in the Criteria sheet.
- The grading engine reads criteria exclusively from the **Criteria sheet** at runtime.
- Teachers can edit the Criteria sheet directly (tweak descriptions, adjust points) and those changes take effect immediately.
- There is no embedded CSV in `Code.gs` — criteria and code are completely decoupled.

**For developers adding/changing criteria:**
1. Edit the appropriate CSV file in `criteria/` (or create a new one for a different curriculum).
2. Import it into the Criteria sheet in your test spreadsheet (File → Import → "Replace current sheet").
3. Run **Sync Levels from Criteria** to update the Levels sheet.
4. Grade a known submission to verify the new criteria work as expected.

> **Tip:** Since criteria live in a plain CSV file, anyone can contribute new rubrics for different courses or units without touching `Code.gs` at all.

## Adding a New Level

1. **Add criteria rows** to the appropriate CSV in `criteria/`:
   ```
   LevelID,CriterionID,Points,Type,Description,Notes,Teacher Notes
   Lesson-XX-Level-YY,criterion_name,3,llm_check,"Description of what to check",,
   ```
   Use zero-padded numbers (e.g., `Lesson-03-Level-08`) so levels sort correctly.

2. **Import the CSV** into the Criteria sheet (File → Import → "Replace current sheet").

3. **Run "Sync Levels from Criteria"** from the Autograder menu — this adds the new level to the Levels sheet.

4. **Test** by grading a known submission for that level.

### Criterion Types

All criteria use `llm_check` — the LLM evaluates each criterion against the student's source code and decides pass/fail with a reason. There are no local/regex check types; the LLM handles everything.

## Code Style

- This is Google Apps Script (ES5-compatible JavaScript) — no `let`/`const`, no arrow functions, no template literals
- Use `var` for all variable declarations
- Functions intended as internal helpers use a trailing underscore: `myHelper_()`
- Keep the section numbering and separator comments consistent

## Testing

There's no automated test suite (it's Apps Script). Manual testing workflow:

1. Paste `Code.gs` into a Google Sheet's Apps Script editor
2. Run **Initial Setup** → verify all sheets created correctly
3. Import a criteria CSV into the Criteria sheet → run **Sync Levels from Criteria**
4. Run **Test API Connection** → verify both checks pass
5. Grade a known submission → verify score matches expectations
6. Test the email flow on a test row

## Questions?

Open an issue — happy to help when I have time!
