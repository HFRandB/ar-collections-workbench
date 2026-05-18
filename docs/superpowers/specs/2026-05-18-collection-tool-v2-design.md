---
name: collection-tool-v2-design
description: Design spec for v2 enhancements to collection-tool.html — Document Number grouping, dunning email (mailto), and contact log written back to Current Status column — May 2026
metadata:
  type: project
---

# Collection Tool v2 — Design Spec

**Date:** 2026-05-18  
**File:** `/Users/I306662/claude code/collection-tool.html` (enhance existing)  
**Source:** `collector work ar report.xlsx`

---

## 1. Overview

Three enhancements to the existing `collection-tool.html`:

1. **Document Number grouping** — Work tab list collapses rows sharing the same Document Number into one entry, summing balances
2. **Dunning email (A)** — Per-row ✉ button generates a pre-filled `mailto:` link in the correct language based on Company Code
3. **Contact log (B)** — Expanded row gains a "Log Contact" section; saved entries are prepended to the Current Status field and exported with the Excel

---

## 2. New Column Indices Used

In addition to the columns already used by v1, v2 uses:

| Index | Column Name |
|-------|-------------|
| 2 | Company Code |
| 23 | Local Currency Amt |
| 24 | Local Currency |

Existing indices already in use: 13 (Bill-to Number), 14 (Name of Bill-to Party), 20 (Document Number), 28 (Net Due Date), 55 (OD Days), 56 (OD Category), 66 (Open Balance in EUR), 71 (Contact Person), 72 (FC Month), 73 (Customer Payment Process), 74 (Current Status), 75 (AR Plan), 76 (Pending Reason Category).

---

## 3. Document Number Grouping

### 3.1 Grouping Logic

After `parseWorkbook()` loads `allRows`, the Work tab groups rows by the value in `COL.DOC_NUM` (index 20) at render time. Each unique Document Number becomes one display row (`docGroup`):

```
docGroup = {
  docNum:      string,
  rows:        Row[],        // all allRows entries for this Document Number
  localAmt:    number,       // sum of Local Currency Amt across rows
  localCurr:   string,       // Local Currency from first row (all same)
  openEur:     number,       // sum of Open Balance in EUR across rows
  odDays:      number,       // max OD Days across rows
  companyCode: string,       // from first row
  // display fields from first row:
  customer, netDueDate, odCategory, fcMonth, contactPerson, currentStatus
}
```

Filters (Customer, FC Month, OD Category, Flagged) operate on `docGroup` level:
- **Customer** filter: matches `docGroup.customer`
- **FC Month** filter: matches FC Month of **any** row in the group
- **OD Category** filter: matches OD Category of **any** row in the group
- **Flagged** filter: group is flagged if **any** row in the group is flagged

Row color coding uses `docGroup.odDays` and `docGroup.currentStatus` (same priority rules as v1).

### 3.2 Work Table Columns

| Column | Source |
|--------|--------|
| ⚑ Flag | Flagged if any row in group is flagged |
| Customer | Name of Bill-to Party (first row) |
| Document # | Document Number |
| Net Due Date | Net Due Date (DD-Mon-YYYY, first row) |
| OD Days | Max across group rows |
| OD Category | First row |
| FC Month | First row |
| Open Balance (TWD) | Sum of Local Currency Amt, formatted `TWD X,XXX` |
| Open Balance (EUR) | Sum of Open Balance in EUR, formatted `€X,XXX` |
| ✉ | Dunning email button |

### 3.3 Expandable Rows

Clicking a group row expands it. The inline editable fields (Current Status, Customer Payment Process, FC Month, AR Plan, Pending Reason Category, Contact Person, Last Internal Note) write to **all rows** in the group when saved. The Contact Log section (section 5) also appears here.

---

## 4. Dunning Email (mailto:)

### 4.1 Trigger

Each row in the Work table has a ✉ button in the last column. Clicking it constructs a `mailto:` URL and opens it (calls `window.location.href = mailtoUrl`), which launches the user's default mail client (Outlook) with fields pre-filled.

### 4.2 Language Selection

| Company Code | Language |
|---|---|
| `0073` | 繁體中文 |
| `0038` | 簡體中文 |
| `0074` | English |
| (other) | English (fallback) |

### 4.3 Email Templates

**Amount format:** `{localCurr} {localAmt}` (e.g., `TWD 283,862`)  
**Recipient (`to`):** Contact Person field (col 71); empty if not set  
**Subject and body** vary by language:

#### 繁體中文 (0073)

```
主旨：催款通知 — {customer} 文件編號 {docNum}

{customer} 您好，

謹此提醒，以下應收帳款已逾期，煩請儘快安排付款：

　文件編號：{docNum}
　到期日：{netDueDate}
　逾期天數：{odDays} 天
　逾期金額：{localCurr} {localAmt}

如有任何疑問，歡迎隨時與我聯繫。

謝謝！
Huang Fan
```

#### 简体中文 (0038)

```
主题：催款通知 — {customer} 文件编号 {docNum}

{customer} 您好，

谨此提醒，以下应收账款已逾期，烦请尽快安排付款：

　文件编号：{docNum}
　到期日：{netDueDate}
　逾期天数：{odDays} 天
　逾期金额：{localCurr} {localAmt}

如有任何疑问，欢迎随时与我联系。

谢谢！
Huang Fan
```

#### English (0074 / fallback)

```
Subject: Payment Reminder — {customer} Document {docNum}

Dear {customer},

This is a reminder that the following invoice is overdue. Please arrange payment at your earliest convenience:

  Document Number: {docNum}
  Due Date:        {netDueDate}
  Days Overdue:    {odDays} days
  Amount Due:      {localCurr} {localAmt}

Please do not hesitate to contact me if you have any questions.

Best regards,
Huang Fan
```

All values are URI-encoded before being embedded in the `mailto:` URL.

---

## 5. Contact Log

### 5.1 UI (in expanded row)

Below the existing editable fields, a new "📝 記錄聯繫" section:

```
聯繫方式  [電話 ▾]   (options: 電話 / 郵件 / 線上會議)
備注      [                                    ]
                                    [儲存記錄]
```

### 5.2 Save Behaviour

On click of "儲存記錄":

1. Build entry string: `[YYYY-MM-DD 聯繫方式] 備注內容`  
   Example: `[2026-05-18 電話] 客戶表示下週付款`
2. Read current value of Current Status for the group (from first row's cells)
3. New value: `{entry} | {existing}` (prepend; if existing is empty, just `{entry}`)
4. Write new value to `cells[COL.CURR_STATUS]` of **all rows** in the group (so export captures it)
5. Re-render the expanded row to show updated Current Status
6. Clear the contact log inputs

### 5.3 Export

No change to export logic — because the log is written into `cells[COL.CURR_STATUS]`, it is automatically included when `exportToExcel()` writes back using `row.sheetRow`.

---

## 6. COL Constants to Add

```js
const COL = {
  // existing …
  COMPANY_CODE:   2,
  LOCAL_AMT:      23,
  LOCAL_CURR:     24,
  // existing columns remain unchanged
};
```

---

## 7. Styling

| Element | Style |
|---------|-------|
| ✉ button | Small blue button `#2563eb`, hover `#1d4ed8` |
| Local currency column | Right-aligned, `font-variant-numeric: tabular-nums` |
| Contact log section | Light gray background `#f8fafc`, 8px padding, rounded |
| "儲存記錄" button | Blue `#2563eb` |

---

## 8. Unchanged from v1

- Dashboard tab (KPI cards, charts, details table) — no changes
- Export to Excel logic — no structural changes, works via `row.sheetRow`
- Flag toggle behaviour
- Row color coding priority order
- Filter bar UI and logic (extended to operate on docGroups)
- All existing COL constants

---

## 9. File

Single file: `/Users/I306662/claude code/collection-tool.html`
