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
