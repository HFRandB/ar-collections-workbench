---
name: collection-tool-design
description: Design spec for a standalone HTML/JS AR Collection Tool with Work tab (editable list) and Dashboard tab (view/tracker) for Huang Fan May 2026
metadata:
  type: project
---

# AR Collection Tool — Design Spec

**Date:** 2026-05-15
**Collector:** Huang Fan
**Source:** `collector work ar report.xlsx`
**Output:** Single self-contained `collection-tool.html` file (no server, opens in any browser)

---

## 1. Overview

A single `collection-tool.html` file with two modes:

- **Work tab** — editable working list for updating invoice statuses, flagging follow-up items, and exporting changes back to Excel
- **Dashboard tab** — visual tracker with KPI cards, charts, and aging table driven dynamically from the loaded file

---

## 2. Data Loading

- User drags and drops `collector work ar report.xlsx` onto the page (or clicks to browse)
- SheetJS parses the file; data headers are read from **row 2** (row 1 is metadata, row 3+ is data)
- Once loaded, the two tabs and Export button appear
- Edits live in memory; clicking **Export to Excel** downloads the full updated XLSX with edits applied to the relevant columns
- No auto-save — export is the explicit save action

**Key column indices (0-based from row 2 headers):**

| Index | Column Name |
|-------|-------------|
| 13 | Bill-to Number |
| 14 | Name of Bill-to Party |
| 18 | Sold-To Party Name |
| 19 | Collector Name |
| 20 | Document Number |
| 28 | Net Due Date |
| 37 | Dispute Case ID |
| 50 | Last Internal Note |
| 51 | Last External Note |
| 55 | OD Days |
| 56 | OD Category |
| 66 | Open Balance in EUR |
| 71 | Contact Person |
| 72 | FC Month |
| 73 | Customer Payment Process |
| 74 | Current Status |
| 75 | AR Plan |
| 76 | Pending Reason Category |
| 81 | New Item? |
| 82 | Unique ID |

---

## 3. Layout & Structure

```
┌────────────────────────────────────────────────────────┐
│  AR Collection Tool — Huang Fan                        │
│  [Drop XLSX here or click to browse]  [Export to Excel]│
│  ┌──────────┬─────────────┐                            │
│  │  Work ✎  │  Dashboard  │                            │
│  └──────────┴─────────────┘                            │
├────────────────────────────────────────────────────────┤
│  [tab content]                                         │
└────────────────────────────────────────────────────────┘
```

---

## 4. Work Tab

### 4.1 Filter Bar

```
Customer ▾  │  FC Month ▾  │  OD Category ▾  │  ☐ Flagged only
```

- **Customer** — dropdown of all unique `Name of Bill-to Party` values
- **FC Month** — dropdown: All / May / May-Risk / Jun / Jun-Risk / Jul-Risk / Offset/Concession/Cancel/Cleared
- **OD Category** — dropdown of all unique `OD Category` values
- **Flagged only** — checkbox; when checked, shows only rows with the flag set

### 4.2 Table

Displayed columns (non-editable, display only):

| Column | Source |
|--------|--------|
| ⚑ Flag | In-memory flag per row |
| Customer | Name of Bill-to Party |
| Document # | Document Number |
| Net Due Date | Net Due Date (formatted DD-Mon-YYYY) |
| OD Days | OD Days |
| OD Category | OD Category |
| FC Month | FC Month |
| Open Balance (EUR) | Open Balance in EUR (formatted €X,XXX) |

### 4.3 Expandable Rows

Clicking a row expands it to reveal inline editable fields:

| Field | Type | Writes to column |
|-------|------|-----------------|
| Current Status | Free text input | Current Status (74) |
| Customer Payment Process | Dropdown | Customer Payment Process (73) |
| FC Month | Dropdown | FC Month (72) |
| AR Plan | Free text input | AR Plan (75) |
| Pending Reason Category | Free text input | Pending Reason Category (76) |
| Contact Person | Free text input | Contact Person (71) |
| Last Internal Note | Textarea | Last Internal Note (50) |

Dropdown options for **Customer Payment Process:**
- Standard (invoice- internal payment process)
- Acceptance before FI
- PO/PR process
- Paid

Dropdown options for **FC Month:**
- May / May-Risk / Jun / Jun-Risk / Jul-Risk / Offset/Concession/Cancel/Cleared

### 4.4 Row Color Coding

Applied to the entire row background. Priority order (highest first):

1. **Green** `#dcfce7` — `Current Status` contains "paid" or "Paid" (case-insensitive)
2. **Red** `#fee2e2` — `OD Days` > 365
3. **Yellow** `#fef9c3` — `OD Days` > 90 (and ≤ 365)
4. **White** — all other rows

### 4.5 Flag Toggle

- Each row has a ⚑ button in the first column
- Clicking toggles the flag on/off (in-memory)
- Flagged rows show a bold left border (`#f97316`) regardless of row color
- Flagged state is **in-memory only** — it does not export to Excel. It is a session-only visual aid; clear when the page is reloaded.

---

## 5. Dashboard Tab

### 5.1 KPI Cards (4 cards)

```
┌──────────┬──────────────────┬──────────────────┬──────────────────┐
│  Total   │      May         │      Jun         │      Jul         │
│  Open    │  May:   €9.4M    │  Jun:   €3.0M    │  Jul:    —       │
│  €14.8M  │  Upside: €171K   │  Upside: €3.6M   │  Upside: €444K   │
│          │  Total: €9.6M    │  Total: €6.6M    │  Total: €444K    │
└──────────┴──────────────────┴──────────────────┴──────────────────┘
```

Computed from loaded data:

| Card | Calculation |
|------|-------------|
| Total Open | Sum of `Open Balance in EUR` for all rows |
| May — Normal | Sum where `FC Month = "May"` |
| May — Upside | Sum where `FC Month = "May-Risk"` |
| May — Total | Normal + Upside |
| Jun — Normal | Sum where `FC Month = "Jun"` |
| Jun — Upside | Sum where `FC Month = "Jun-Risk"` |
| Jun — Total | Normal + Upside |
| Jul — Normal | Sum where `FC Month = "Jul"` (may be 0 → shown as "—") |
| Jul — Upside | Sum where `FC Month = "Jul-Risk"` |
| Jul — Total | Normal + Upside |

Normal line: blue `#2563eb`. Upside line: orange `#f97316`. Total line: bold dark.

### 5.2 Monthly Balance Stacked Bar Chart

- X-axis labels: May / Jun / Jul / Offset/Cancel/Cleared
- Blue dataset: sum of normal FC Month rows per label
- Orange dataset: sum of Risk FC Month rows per label
- Stacked on both axes
- Y-axis ticks formatted as `€X.XM`
- Tooltip shows formatted EUR values

### 5.3 Top Customers Horizontal Bar Chart

- One bar per unique `Name of Bill-to Party`
- Value: sum of `Open Balance in EUR` per customer
- Sorted descending
- `indexAxis: 'y'`, no legend, blue `#2563eb`
- X-axis ticks `€X.XM`

### 5.4 OD Aging Breakdown Chart

- Horizontal bar chart
- One bar per unique `OD Category` value
- Value: count of invoice rows per category
- Sorted by OD category order (a → j)
- No legend, blue `#2563eb`
- X-axis shows item count (integer)

### 5.5 Details Table

Columns: Month | Normal Balance (EUR) | Risk/Upside Balance (EUR) | % Risk

- Rows: May / Jun / Jul / Offset/Concession/Cancel/Cleared
- % Risk = Risk / (Normal + Risk) × 100, highlighted orange bold when > 20%
- Negative normal balances shown in parentheses: `(€1,852,088)`
- Zero risk shown as "—"

---

## 6. Technology

| Concern | Choice |
|---------|--------|
| File | Single `collection-tool.html`, no build step |
| Excel read/write | SheetJS (`xlsx`) via CDN `cdn.jsdelivr.net` |
| Charts | Chart.js 4.x via CDN `cdn.jsdelivr.net` |
| Font | System sans-serif (no external CDN) |
| Responsive | Charts scale to window width |

---

## 7. Styling

| Element | Color |
|---------|-------|
| Normal balance / bars | Blue `#2563eb` |
| Risk/Upside | Orange `#f97316` |
| Risk KPI upside text | Orange `#f97316` |
| Row: Green (Paid) | `#dcfce7` background |
| Row: Red (OD > 365) | `#fee2e2` background |
| Row: Yellow (OD > 90) | `#fef9c3` background |
| Flagged row border | Orange `#f97316` left border 4px |
| Table % Risk highlight | Orange `#f97316` bold when > 20% |
| Background | Light gray `#f1f5f9` |
| Card background | White `#ffffff` |

---

## 8. Interactivity

- Drag-and-drop XLSX loading (click-to-browse fallback)
- Filter bar (customer, FC month, OD category, flagged-only)
- Expandable rows with inline field editing
- Flag toggle per row (persists to export via `AR Plan` prefix)
- Export to Excel (full XLSX download with edits applied)
- Chart hover tooltips showing formatted EUR values or counts
- Tab switching (Work / Dashboard)

---

## 9. Output Location

File saved to: `/Users/I306662/claude code/collection-tool.html`
(same directory as the source XLSX)
