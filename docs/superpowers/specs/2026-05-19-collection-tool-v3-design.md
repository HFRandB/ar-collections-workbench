---
name: collection-tool-v3-design
description: Design spec for v3 enhancement to collection-tool.html — multi-select rows for batch dunning email grouped by customer, with live OD days calculation — May 2026
metadata:
  type: project
---

# Collection Tool v3 — Design Spec

**Date:** 2026-05-19  
**File:** `/Users/I306662/claude code/collection-tool.html` (enhance existing)

---

## 1. Overview

One enhancement to the existing `collection-tool.html`:

**Multi-select batch email** — Checkboxes on work table rows allow selecting multiple document groups. A toolbar shows the selection count and a "Send" button. Selected rows are grouped by customer; one combined dunning email is generated per customer, listing all selected documents for that customer. OD days in the email are calculated live as `today − net due date`.

---

## 2. Selection UI

### 2.1 Checkbox Column

A new checkbox column is prepended to the work table as the first column (before the ⚑ flag column). The table becomes 11 columns total; all `colspan` values update accordingly.

| Column | Content |
|--------|---------|
| ☐ | Per-row checkbox; `id="chk-{gIdx}"` |
| ⚑ | Flag (unchanged) |
| Customer | (unchanged) |
| … | (all remaining columns unchanged) |

The `<thead>` gains a "select all" checkbox in the first `<th>`:
```html
<th><input type="checkbox" id="chkAll" onchange="toggleSelectAll(this.checked)"></th>
```

### 2.2 Selection State

A module-level `Set` tracks selected doc numbers:
```js
const selectedDocs = new Set(); // docNum → selected
```

Per-row checkbox `onchange` calls `toggleSelect(gIdx, docNum, checked)`:
- Adds/removes `docNum` from `selectedDocs`
- Calls `updateSelectionToolbar()`

`toggleSelectAll(checked)`:
- Adds/removes ALL currently visible groups' docNums
- Re-renders checkboxes (calls `renderWorkTable()`)

### 2.3 Selection Toolbar

A `<div id="selectionToolbar">` sits between the filter bar and the work table, hidden by default (`display: none`). When `selectedDocs.size > 0` it shows:

```
3 selected   [✉ Send (3)]   [Clear]
```

`updateSelectionToolbar()` updates the count and shows/hides the toolbar.

"Clear" calls `clearSelection()`: empties `selectedDocs`, calls `renderWorkTable()`.

---

## 3. Batch Email Generation

### 3.1 Grouping Logic

When the user clicks "Send (N)":

1. Collect all groups from `_lastGroups` whose `docNum` is in `selectedDocs`
2. Group them by `group.customer` (string match)
3. For each customer group, call `buildBatchMailto(customerGroups)` and collect the resulting URLs

### 3.2 Opening the Emails

- **1 customer selected** → `window.location.href = url` (same as today's single-doc flow)
- **Multiple customers selected** → display a small panel below the toolbar with one `<a href="mailto:...">` link per customer. The user clicks each link to open them one by one in Outlook. The panel is dismissed by clicking "Clear" or deselecting all.

### 3.3 OD Days Calculation

In `buildBatchMailto` (and updated in the existing `buildMailto`), OD days are calculated live:

```js
function calcOdDays(netDueDate) {
  if (!netDueDate) return 0;
  const due = netDueDate instanceof Date ? netDueDate : new Date(netDueDate);
  if (isNaN(due)) return 0;
  return Math.max(0, Math.floor((Date.now() - due.getTime()) / 86400000));
}
```

This replaces `group.odDays` in both `buildMailto` (existing single-row) and `buildBatchMailto` (new multi-row).

---

## 4. Email Templates

### 4.1 Multi-Document Format

**Amount format per line:** `{localCurr} {localAmt}` (abs value, comma-formatted)  
**To:** `contactPerson` of the first document in the customer group (empty if not set)  
**Language:** `companyCode` of the first document in the customer group

#### 繁體中文 (0073)

```
主旨：催款通知 — {customer}

{customer} 您好，

謹此提醒，以下應收帳款已逾期，煩請盡快安排付款：

　文件編號        到期日          逾期天數    逾期金額
　{docNum}        {netDueDate}    {odDays} 天  {localCurr} {localAmt}
　{docNum}        {netDueDate}    {odDays} 天  {localCurr} {localAmt}
　...

如有任何疑問，歡迎隨時與我聯繫。

謝謝！
Huang Fan
```

#### 简体中文 (0038)

```
主题：催款通知 — {customer}

{customer} 您好，

谨此提醒，以下应收账款已逾期，烦请尽快安排付款：

　文件编号        到期日          逾期天数    逾期金额
　{docNum}        {netDueDate}    {odDays} 天  {localCurr} {localAmt}
　...

如有任何疑问，欢迎随时与我联系。

谢谢！
Huang Fan
```

#### English (0074 / fallback)

```
Subject: Payment Reminder — {customer}

Dear {customer},

This is a reminder that the following invoices are overdue. Please arrange payment at your earliest convenience:

  Document Number    Due Date        Days Overdue   Amount Due
  {docNum}           {netDueDate}    {odDays} days  {localCurr} {localAmt}
  ...

Please do not hesitate to contact me if you have any questions.

Best regards,
Huang Fan
```

All values are URI-encoded before embedding in the `mailto:` URL.

### 4.2 Single-Document Backward Compatibility

The existing `buildMailto(group)` function (used by per-row ✉ button) is updated to use `calcOdDays(group.netDueDate)` instead of `group.odDays`. No other changes to the single-row flow.

---

## 5. New Functions

| Function | Purpose |
|----------|---------|
| `toggleSelect(gIdx, docNum, checked)` | Add/remove docNum from selectedDocs; update toolbar |
| `toggleSelectAll(checked)` | Select/deselect all visible groups |
| `updateSelectionToolbar()` | Show/hide toolbar, update count |
| `clearSelection()` | Empty selectedDocs, re-render |
| `sendSelectedEmails()` | Group selected by customer, open mailto URLs |
| `buildBatchMailto(groups)` | Build single mailto: URL for multiple docs of one customer |
| `calcOdDays(netDueDate)` | Calculate live OD days from due date to today |

---

## 6. CSS Changes

| Element | Style |
|---------|-------|
| `#selectionToolbar` | Light blue background `#eff6ff`, padding 8px 16px, border-bottom `1px solid #bfdbfe`, hidden by default |
| `.btn-send-selected` | Same blue as `.email-btn` (`#2563eb`) |
| `.mailto-links-panel` | Light gray `#f8fafc`, padding, rounded — shows clickable links when multiple customers |

---

## 7. Unchanged from v2

- Document Number grouping logic
- Export to Excel
- Contact log
- Dashboard tab
- Row color coding
- Flag toggle
- Filter bar

---

## 8. File

Single file: `/Users/I306662/claude code/collection-tool.html`
