# Contributing to Game Lab Autograder

Thanks for your interest in improving the autograder! This project is built by teachers for teachers.

## Ways to Contribute

### Report Issues
- **Bug reports:** Something not grading correctly? Share the LevelID and describe what happened.
- **Grading accuracy:** If a criterion is consistently too strict or too lenient, open an issue with examples.
- **Setup problems:** Describe your environment (school Google Workspace vs. personal Gmail, etc.).
- **Note:** This project is maintained in my spare time. I may not respond quickly to issues or review all pull requests. Feel free to fork and adapt for your own needs.

### Suggest Improvements
- New levels or criteria for existing CSD Unit 3 lessons
- UX improvements to the menu, dialogs, or email templates
- Documentation improvements

### Submit Code
1. Fork the repository
2. Make your changes to `Code.gs` (or `criteria-table.csv` for rubric changes)
3. Test in a real Google Sheet (paste into Apps Script, run Initial Setup, grade a few submissions)
4. Open a pull request with a clear description of what changed and why

## Adding a New Level

1. **Add criteria rows** to `criteria-table.csv`:
   ```
   LevelID,CriterionID,Points,Type,Description,Notes,Teacher Notes
   Lesson-XX-Level-YY,criterion_name,3,llm_check,"Description of what to check",,
   ```

2. **Add the same rows** to the `getCriteriaTableCsvText_()` function in `Code.gs` (the embedded CSV must stay in sync with the file).

3. **Test** by running Initial Setup (Reset Everything first), then grading a known submission for that level.

### Criterion Types

| Type | Use when... |
|---|---|
| `llm_check` | The check requires understanding code logic (most criteria) |
| `code_nonempty` | You just need to verify the student wrote something |
| `contains` | Checking for a specific string (e.g., a function name) |
| `regex_present` | Pattern matching (more flexible than `contains`) |
| `regex_absent` | Ensuring something is NOT in the code |

Prefer `llm_check` for anything that requires judgment. Use local checks (`contains`, `regex_*`) only for simple, unambiguous requirements — they're faster and don't use API credits.

## Code Style

- This is Google Apps Script (ES5-compatible JavaScript) — no `let`/`const`, no arrow functions, no template literals
- Use `var` for all variable declarations
- Functions intended as internal helpers use a trailing underscore: `myHelper_()`
- Keep the section numbering and separator comments consistent

## Testing

There's no automated test suite (it's Apps Script). Manual testing workflow:

1. Paste `Code.gs` into a Google Sheet's Apps Script editor
2. Run **Initial Setup** → verify all sheets created correctly
3. Run **Test API Connection** → verify both checks pass
4. Grade a known submission → verify score matches expectations
5. Test the email flow on a test row

## Questions?

Open an issue — happy to help when I have time!
