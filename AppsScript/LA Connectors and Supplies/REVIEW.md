# LA Connectors and Supplies — Code Review

Review date: 2025-02-25 (after syncing from Google Apps Script, timezone set to America/Los_Angeles).

---

## 1. appsscript.json

**Summary:** Correct and aligned with what the scripts need.

- **Timezone:** `America/Los_Angeles` — matches your intent.
- **Scopes:** Appropriate for Forms, Spreadsheets, container UI, userinfo.email, and Gmail send. No obvious missing or unnecessary scopes.
- **Suggestion:** You can pretty-print the JSON (indentation, newlines) for easier diffs in the repo; behavior is unchanged.

---

## 2. Inventory QR.js

**Summary:** Solid structure with a real bug in the form-submit data range (fixed in this repo).

### Strengths

- **@OnlyCurrentDoc** — Limits script access to the current spreadsheet; good for security and review.
- **CONFIGS** — Single place for sheet names, form IDs, column indices, and response sheet names; easy to add another form/sheet later.
- **Shared logic** — `updateFormDropdownForConfig()` keeps Connectors and Supplies behavior consistent and avoids duplication.
- **Error handling** — Try/catch and clear errors; form open failures and missing questions throw with actionable messages.
- **Debug helpers** — `debugListSuppliesFormQuestions()` and `debugOnSubmit()` are useful for support and onboarding.

### Bug fixed in repo

- **Data range in `onFormSubmit()`**  
  Previously:
  - `numRows = getLastRow() - config.headersRow`
  - `getRange(headersRow + 1, 1, numRows, ...)`  
  So the third argument was used as *end row*, but it was the *number of data rows*, not the last row index. That caused the last several rows of inventory to be skipped when matching the submitted bin.  
  **Fix:** Use `inventorySheet.getLastRow()` as the end row so the range is `(headersRow + 1, 1, getLastRow(), getLastColumn())`.

### Minor notes

- **Hard-coded form question titles** — `'Are you taking or adding items?'` and `'Quantity'` in `onFormSubmit` will break if you rename those questions. Consider adding them to `CONFIGS` (e.g. `actionQuestionTitle`, `quantityQuestionTitle`) so one change updates both form and script.
- **Empty helper column** — If the helper column (P) has no values, the form dropdown is left unchanged and you only get a log. That’s reasonable; you could optionally show a toast or alert for visibility.

---

## 3. Timestamp Button.js

**Summary:** Simple, clear, and correct.

### Strengths

- **Single config constant** — `TARGET_COLUMN = 12` (Column L) makes it easy to change later.
- **Strict validation** — Ensures selection is in column L only and shows a clear alert if not.
- **User identity** — Uses `Session.getActiveUser().getEmail()` and handles missing email by skipping the signature step.
- **Signature formatting** — Capitalizing the local part of the email is a nice touch for Column M.

### Minor notes

- **Logger volume** — Many `Logger.log()` calls are helpful for debugging. If the script is stable and logs get noisy, you could reduce them or guard with a debug flag.
- **Multi-cell selection** — If the user selects multiple cells in column L, `range.setValue(new Date())` sets the same timestamp in all, and `range.offset(0, 1)` is only the cell next to the top-left of the range. If you ever want “one timestamp per row” for multi-cell selection, you’d need to loop over each row and set time + signature per row.

---

## 4. Weekly Report.js

**Summary:** Does the job; same range bug as Inventory QR (fixed in repo). Some duplication and small consistency tweaks suggested.

### Strengths

- **Clear constants** — Sheet name, recipients, start row, and column indices are defined at the top.
- **Low-stock logic** — Correct use of threshold vs current quantity and `parseFloat` with `isNaN` checks.
- **HTML email** — Table and inline styles are readable and should render consistently.
- **Test report** — “Run Test Report (Ryan Only)” is a safe way to verify the pipeline without spamming the full list.

### Bugs fixed in repo

- **Data range in both report functions**  
  Same idea as in Inventory QR: the end row was computed as `sheet.getLastRow() - START_ROW + 1` (a row count) but passed into `getRange(..., endRow, ...)`, so the last rows of the sheet were never scanned.  
  **Fix:** Use `sheet.getLastRow()` as the end row in both `sendWeeklyLowStockReport()` and `sendTestInventoryReport()`.

### Suggestions

- **DRY** — `sendWeeklyLowStockReport` and `sendTestInventoryReport` share almost all logic (scan, threshold check, table build). Consider a helper, e.g. `getLowStockItems(sheet, startRow, columnDefs)` that returns `itemsToOrder`, and separate functions that build the subject/body and call `GmailApp.sendEmail` with either production or test recipients. That would make future column or threshold changes easier and reduce the chance of fixing a bug in one place and forgetting the other.
- **Recipient list** — `RECIPIENT_EMAILs` has mixed casing (e.g. `ryan@...` vs `Ryan@...`). Email addresses are case-insensitive, but for consistency and to avoid “did I add this person?” confusion, you could normalize to one style (e.g. lowercase) or pull from a sheet/Named Range later if the list grows.
- **Trigger** — The weekly report is only useful if a time-driven trigger runs `sendWeeklyLowStockReport` (e.g. weekly). If that’s already set in the Apps Script project, no change needed; if not, add a trigger in the editor (Triggers → Add Trigger).

---

## Summary of changes made in this repo

| File            | Change                                                                 |
|-----------------|------------------------------------------------------------------------|
| Inventory QR.js | Fixed `onFormSubmit` data range to use `getLastRow()` as end row.      |
| Weekly Report.js| Fixed data range in both report functions to use `getLastRow()` as end row. |

No edits were made to `appsscript.json` or `Timestamp Button.js`; only the range bug fixes and this review were added.

If you deploy from this repo (e.g. copy/paste or clasp), pull these fixes into your bound Apps Script project so form submissions and weekly reports use the full inventory sheet.
