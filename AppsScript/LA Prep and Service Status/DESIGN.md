# LA Prep Floor & Camera Service Status — Design

## Purpose

- **Combined view**: Prep bay assignments + F2/camera service status + (later) Marketing, Shipping, Sub-rental in one place.
- **Audience**: LA team first; architecture must support all offices (LA, Vancouver, Toronto, Chicago, New Orleans, Atlanta, Albuquerque).
- **F2**: All offices use F2 and can export F2 Service Board Excel files; only LA has a Prep Bay Assignment sheet today. Equipment Scheduling Chart is already multi-office.

## Workbook

- **Name**: LA Prep Floor & Camera Service Status  
- **ID**: `1j_slMWpLIbjqbvGdAurozTh_1vv17SASshCZSkTUNw0`  
- **Sheets**: TODAY, Tomorrow, Prep Two Days Out, Three Days Out, Four Days Out (forecast 4+ days; may extend to 5–6).  
- **Test sheet**: "Prep Bay Schema" — used to validate v2 layout and initial output.

## Layout: v1 vs v2

### v1 (current)

- Static, horizontal: 3 bays per block (e.g. PREP BAY 1–3), each with Job name, Order #, Prep Tech, then fixed rows for Cameras (Barcode, Pulled?, Serviced?).
- One job occupying multiple bays is repeated in each bay block → duplicate info and many empty rows.
- Cameras only; limited vertical space per bay.

### v2 (target)

- **Vertical-first**: Optimized for single-mouse/phone scrolling; minimal horizontal scroll.
- **Job-centric**: One block per job (not per bay). Jobs ordered top-to-bottom by prep bay number (meeting order).
- **Prominence**: Job Name → Order # → Prep Bay(s) (e.g. "1, 2, & 3") → Marketing Agent → Prep Tech → Prep Notes.
- **Equipment**: Single table per job with categories (Cameras, Lenses, Heads, Focus, Matte Boxes, Monitors, Media, Wireless Video, Dir. Viewfinder, Ungrouped). Columns: Equipment Name, Barcode, Pulled?, RTR?, Serviced for Order?, Completion Timestamp. Rows grow dynamically as F2 processes more items.
- **Sub-rental**: Per-job section: Subbed Equipment, Quantity, Located, Locating Agent, Quote Rec., Run Sheet Out, Packing Slip, Notes.

## Data sources

| Source | ID / location | Use |
|--------|----------------|-----|
| Prep Bay Assignment | `1erp3GVvekFXUVzC4OJsTrLBgqL4d0s-HillOwyJZOTQ` | LA only; date-named sheets (e.g. "Tues 2/20"); bay, job, order, prep tech, notes. |
| Equipment Scheduling Chart | `1uECRfnLO1LoDaGZaHTHS3EaUdf8tte5kiR6JNWAeOiw` | Camera (and Consignor Use Only) by date column; LOS ANGELES rows; order → cameras. |
| F2 Imports (destination) | Same as Process F2 Import destination | Serviced gear by order; populated by Process F2 Import.js. |
| The SUB Sheet (sub-rental) | `1UUAwABLOAQLt9M4uTa8E6DkdLJSKHIC7s_xZDzLtBQw` | Quote #, Job Name, Pre-Prep/Prep dates, Marketing Agent, Prep Tech, Puller, Locating Agent; equipment lines with Called/Located/Quote Received/Run Sheet/Packing Slip/Notes. Scan whole workbook for matching order/quote and extract. TBD. |

## SUB Sheet (sub-rental)

- **Layout**: Repeating blocks per job: header row (Quote #, Production Company, Job Name, Pre-Prep Date, Prep/Ship Date, Return Date, Marketing Agent, Prep Tech, Puller, Locating Agent, Suggested Alternative); then equipment rows (Time Needed By, QTY, Requested Equipment, Billing Days, Added By, Agent Signoff, Called, Located, Quote Received, Run Sheet w/ Shipping, Packing Slip, Notes).
- **Tabs**: Template + job-specific tabs (e.g. "B.S.O.W Season 1 Main Unit 3/2 - 3/26 Prep") and/or day-based sheets (e.g. "Monday 2/23/26").
- **Strategy**: Scan entire workbook (all sheets) for rows where Quote # (or Order #) matches our orders; aggregate sub-rental lines per order for the v2 "Subbed Equipment" section.

## Forecast horizon

- **Current v1**: Same day + one day ahead.
- **v2**: At least 4 days out (TODAY, Tomorrow, Prep Two Days Out, Three Days Out, Four Days Out); possibly 5–6 days. Each forecast sheet gets the same v2 schema, filtered by the appropriate date.

## Script location

- Scripts live in the **Apps Script project bound to the workbook** `1j_slMWpLIbjqbvGdAurozTh_1vv17SASshCZSkTUNw0`.
- This Cursor folder **"LA Prep and Service Status"** holds the source used to copy/paste or deploy into that Apps Script project (no direct file binding; user deploys manually or via clasp if desired).

## Schema (Prep Bay Schema test sheet)

- **Row layout per job block** (aligned with v2 CSV):
  1. Job Name: [value]
  2. Order #: [value]
  3. Prep Bay(s): [e.g. "1, 2, & 3"]
  4. Marketing Agent: [value]
  5. Prep Tech: [value]
  6. Prep Notes: [value]
  7. (blank)
  8. Equipment table header: Equipment Name | Barcode | Pulled? | RTR? | Serviced for Order? | Completion Timestamp
  9. Category rows: Cameras, Lenses, Heads, Focus, Matte Boxes, Monitors, Media, Wireless Video, Dir. Viewfinder, Ungrouped (each can have 0+ data rows below).
  10. (blank)
  11. Subbed Equipment header: Subbed Equipment | Quantity | Located | Locating Agent | Quote Rec. | Run Sheet Out | Packing Slip | Notes
  12. Sub-rental data rows (dynamic).
  13. (blank before next job block)

- Equipment and sub-rental rows grow dynamically; no fixed slot count per job.

## Order of operations: how each job block gets updated

So styling is predictable and nothing from an early step is left visible where it shouldn’t be.

### Phase 1: Sheet-level (once per sheet)

1. **Clear** – `sheet.clear()` removes all content and formatting.
2. **Write data** – `setValues(allRows)` for the full data range. No borders or wrap are set here.
3. **Wrap** – `setWrap(true)` is applied to the **entire written range**. So every cell starts with wrap. Later, per-block formatting overrides this for Job Name (B1) and Prep Notes (B6) to Overflow.
4. **Column widths** – `applySchemaColumnWidths()` (no borders).

### Phase 2: Per-block formatting – `applyJobBlockFormatting(sheet, startRow, fmt, jobHeaderBgOverride, blockRowCount)`

For each block, `startRow` = `r` (1-based). Operations run in this order:

1. **Job header band (rows r–r+5)** – Set background `jobBg` on `(r, 1, r+5, numCols)`. Then set fonts/overflow on row r (Job Name) and rows r+1…r+5 (Order #, Prep Bay, Marketing Agent, Prep Tech, Prep Notes).
2. **Borders** – Clear all borders in the block so no default grid or previous borders remain. Then set **only** the horizontal divider under Prep Notes: bottom border on row `r+5` (grey, medium). That is the only border we want in the block.
3. **Equipment header (row r+6)** – Background, font, row height.
4. **Equipment data rows (r+7 up to Locating Agent row)** – Background `jobBg`, black font, row height; column A bold only for category labels (e.g. `"Cameras:"`). Then checkboxes and font color again on the same range (redundant but harmless).
5. **Subbed Equipment header** – Row found by scanning for "Locating Agent"; same style as Equipment header.
6. **Subbed data row** – Row height, checkboxes.
7. **Black bar** – Last row of block: full row background `#000000`.

### Styling that can look like “early vs later”

- **Wrap vs overflow** – Phase 1 sets wrap on the whole range; Phase 2 sets Overflow on Job Name and Prep Notes. So those two cells are intentionally overwritten.
- **Border between rows 11 and 12** – The only border we set is under row 6 (Prep Notes). If a line appears between 11 and 12, it was either default grid or a border we didn’t clear. Clearing all borders in the block first, then setting only the row‑6 divider, removes stray borders.
