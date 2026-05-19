# Collection Tool v3 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add multi-select batch dunning email to the Work table — checkboxes on each row, a toolbar showing selection count, and a combined mailto: per customer that lists all selected documents with live OD days.

**Architecture:** All changes are to a single existing file `/Users/I306662/claude code/collection-tool.html`. Three self-contained tasks: (1) live OD days utility + update existing single-row email, (2) checkbox column + selection state + toolbar UI, (3) batch email builder + send logic.

**Tech Stack:** Vanilla HTML/CSS/JS (existing) — no new dependencies.

---

## File Map

| Action | Path |
|--------|------|
| Modify | `/Users/I306662/claude code/collection-tool.html` |

---

### Task 1: calcOdDays Utility and Update buildMailto

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

This task adds the `calcOdDays` helper and wires it into the existing single-row `buildMailto`. After this task the per-row ✉ button sends live OD days instead of the stale Excel value.

- [ ] **Step 1: Add calcOdDays function**

In the `<script>` block that contains `buildMailto` (around line 825), add the following function immediately before `buildMailto`:

```js
function calcOdDays(netDueDate) {
  if (!netDueDate) return 0;
  const due = netDueDate instanceof Date ? netDueDate : new Date(netDueDate);
  if (isNaN(due.getTime())) return 0;
  return Math.max(0, Math.floor((Date.now() - due.getTime()) / 86400000));
}
```

- [ ] **Step 2: Update buildMailto to use calcOdDays**

Inside `buildMailto`, find:
```js
  const od      = group.odDays || 0;
```

Replace with:
```js
  const od      = calcOdDays(group.netDueDate);
```

- [ ] **Step 3: Verify**

Open `/Users/I306662/claude code/collection-tool.html`, load the Excel file, and click a ✉ button. Confirm the email body's 逾期天數 / Days Overdue reflects today's date minus the due date, not the stale Excel value.

- [ ] **Step 4: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: add calcOdDays for live overdue days in email"
```

---

### Task 2: Checkbox Column, Selection State, and Toolbar UI

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

This task adds the checkbox column to the work table, tracks selection in a `selectedDocs` Set, and shows a toolbar between the filter bar and table when anything is selected. After this task the user can check/uncheck rows and see the count; the Send button calls a stub `sendSelectedEmails()` (implemented in Task 3).

- [ ] **Step 1: Add selectedDocs to global state**

Find:
```js
let _lastGroups = []; // last rendered groups; used by updateField to skip per-keystroke rebuild
```

Replace with:
```js
let _lastGroups = []; // last rendered groups; used by updateField to skip per-keystroke rebuild
const selectedDocs = new Set(); // docNum values of currently selected groups
```

- [ ] **Step 2: Add CSS for checkbox column, toolbar, and flagged row border fix**

Find:
```css
    tr.row-flagged td:first-child { border-left: 4px solid #f97316; }
```

Replace with:
```css
    tr.row-flagged td:nth-child(2) { border-left: 4px solid #f97316; }
```

(The checkbox column becomes the new first child; the flag column moves to second.)

Then find:
```css
    .num-cell { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
```

Replace with:
```css
    .num-cell { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
    .sel-cell { width: 32px; text-align: center; }
    #selectionToolbar { display: none; align-items: center; gap: 12px; padding: 8px 16px; background: #eff6ff; border-bottom: 1px solid #bfdbfe; font-size: 0.85rem; color: #1e40af; flex-wrap: wrap; }
    #selectionToolbar.visible { display: flex; }
    .btn-send-selected { background: #2563eb; color: #fff; border: none; border-radius: 6px; padding: 6px 14px; font-size: 0.85rem; font-weight: 600; cursor: pointer; white-space: nowrap; }
    .btn-send-selected:hover { background: #1d4ed8; }
    .btn-clear-sel { background: none; border: none; color: #2563eb; font-size: 0.85rem; cursor: pointer; text-decoration: underline; padding: 0; }
    .mailto-links-panel { margin-top: 8px; width: 100%; display: flex; flex-wrap: wrap; gap: 8px; }
    .mailto-link { background: #fff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 6px 12px; font-size: 0.85rem; color: #1e40af; text-decoration: none; white-space: nowrap; }
    .mailto-link:hover { background: #dbeafe; }
```

- [ ] **Step 3: Add selectionToolbar div to HTML**

Find:
```html
    <div class="table-wrapper">
```

Replace with:
```html
    <div id="selectionToolbar">
      <span id="selCount">0 selected</span>
      <button class="btn-send-selected" onclick="sendSelectedEmails()">✉ Send (<span id="selCountBtn">0</span>)</button>
      <button class="btn-clear-sel" onclick="clearSelection()">Clear</button>
      <div class="mailto-links-panel" id="mailtoLinksPanel"></div>
    </div>
    <div class="table-wrapper">
```

- [ ] **Step 4: Update thead to add checkbox column (11 columns total)**

Find:
```html
          <tr>
            <th>⚑</th>
            <th>Customer</th>
            <th>Document #</th>
            <th>Net Due Date</th>
            <th>OD Days</th>
            <th>OD Category</th>
            <th>FC Month</th>
            <th class="num-cell">Open Balance (TWD)</th>
            <th class="num-cell">Open Balance (EUR)</th>
            <th>✉</th>
          </tr>
```

Replace with:
```html
          <tr>
            <th class="sel-cell"><input type="checkbox" id="chkAll" onchange="toggleSelectAll(this.checked)"></th>
            <th>⚑</th>
            <th>Customer</th>
            <th>Document #</th>
            <th>Net Due Date</th>
            <th>OD Days</th>
            <th>OD Category</th>
            <th>FC Month</th>
            <th class="num-cell">Open Balance (TWD)</th>
            <th class="num-cell">Open Balance (EUR)</th>
            <th>✉</th>
          </tr>
```

- [ ] **Step 5: Update renderWorkTable — add checkbox td, update click guard, colspan, and chkAll state**

Find the entire `renderWorkTable` function and replace it with:

```js
function renderWorkTable() {
  const tbody  = document.getElementById('workTbody');
  tbody.innerHTML = '';
  const groups = getFilteredGroups();
  _lastGroups = groups;
  const e = v => (v === null || v === undefined ? '' : String(v)).replace(/&/g, '&amp;').replace(/</g, '&lt;');

  groups.forEach((group, gIdx) => {
    const color    = rowColorClass(group);
    const flagged  = group.rows.some(r => r.flag);
    const expanded = expandedGroups.get(group.docNum) || false;
    const checked  = selectedDocs.has(group.docNum);

    const tr = document.createElement('tr');
    tr.className = 'data-row' + (color ? ' ' + color : '') + (flagged ? ' row-flagged' : '');
    tr.dataset.gidx = gIdx;
    tr.innerHTML = `
      <td class="sel-cell"><input type="checkbox" id="chk-${gIdx}" ${checked ? 'checked' : ''} onchange="toggleSelect(${gIdx},this.checked)"></td>
      <td><button class="flag-btn${flagged ? ' active' : ''}" onclick="toggleFlag(event,${gIdx})">⚑</button></td>
      <td>${e(group.customer)}</td>
      <td>${e(group.docNum)}</td>
      <td>${e(formatDate(group.netDueDate))}</td>
      <td>${group.odDays || ''}</td>
      <td>${e(group.odCategory)}</td>
      <td>${e(group.fcMonth)}</td>
      <td class="num-cell">${formatLocalAmt(group.localAmt, group.localCurr)}</td>
      <td class="num-cell">${formatEUR(group.openEur)}</td>
      <td><button class="email-btn" onclick="openEmail(event,${gIdx})">✉</button></td>`;
    tr.addEventListener('click', ev => {
      if (ev.target.type === 'checkbox') return;
      if (ev.target.classList.contains('flag-btn') || ev.target.classList.contains('email-btn')) return;
      toggleExpand(gIdx, group.docNum);
    });
    tbody.appendChild(tr);

    const expandTr = document.createElement('tr');
    expandTr.className = 'expand-row' + (expanded ? ' open' : '');
    expandTr.id = 'expand-' + gIdx;
    expandTr.innerHTML = `<td colspan="11">${buildExpandHTML(group, gIdx)}</td>`;
    tbody.appendChild(expandTr);
  });

  // Update select-all checkbox state
  const chkAll = document.getElementById('chkAll');
  if (chkAll && groups.length > 0) {
    const allSelected  = groups.every(g => selectedDocs.has(g.docNum));
    const someSelected = groups.some(g => selectedDocs.has(g.docNum));
    chkAll.checked       = allSelected;
    chkAll.indeterminate = someSelected && !allSelected;
  } else if (chkAll) {
    chkAll.checked = false;
    chkAll.indeterminate = false;
  }
}
```

- [ ] **Step 6: Add toggleSelect, toggleSelectAll, updateSelectionToolbar, and clearSelection functions**

Add the following new `<script>` block immediately before `</body>`:

```html
<script>
function toggleSelect(gIdx, checked) {
  const group = _lastGroups[gIdx];
  if (!group) return;
  if (checked) {
    selectedDocs.add(group.docNum);
  } else {
    selectedDocs.delete(group.docNum);
  }
  updateSelectionToolbar();
  // Update select-all state without full re-render
  const chkAll = document.getElementById('chkAll');
  if (chkAll) {
    const allSelected  = _lastGroups.every(g => selectedDocs.has(g.docNum));
    const someSelected = _lastGroups.some(g => selectedDocs.has(g.docNum));
    chkAll.checked       = _lastGroups.length > 0 && allSelected;
    chkAll.indeterminate = someSelected && !allSelected;
  }
}

function toggleSelectAll(checked) {
  _lastGroups.forEach(g => {
    if (checked) {
      selectedDocs.add(g.docNum);
    } else {
      selectedDocs.delete(g.docNum);
    }
  });
  updateSelectionToolbar();
  renderWorkTable();
}

function updateSelectionToolbar() {
  const n       = selectedDocs.size;
  const toolbar = document.getElementById('selectionToolbar');
  const panel   = document.getElementById('mailtoLinksPanel');
  toolbar.classList.toggle('visible', n > 0);
  document.getElementById('selCount').textContent    = n + ' selected';
  document.getElementById('selCountBtn').textContent = n;
  if (panel) panel.innerHTML = '';
}

function clearSelection() {
  selectedDocs.clear();
  document.getElementById('mailtoLinksPanel').innerHTML = '';
  updateSelectionToolbar();
  renderWorkTable();
}
</script>
```

- [ ] **Step 7: Verify**

Load the Excel file. Confirm:
- Each row has a checkbox as its first cell
- The header row has a "select all" checkbox
- Checking a row shows the blue toolbar with count ("1 selected") and "Send (1)" button
- Checking all visible rows makes the header checkbox show as checked
- Checking some rows makes the header checkbox show as indeterminate
- "Clear" button hides the toolbar and unchecks all rows
- Flag button and row expansion still work
- The orange flag border appears on the flag column (second cell), not the checkbox cell

- [ ] **Step 8: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: add checkbox selection column and toolbar to work table"
```

---

### Task 3: Batch Email Builder and Send Logic

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

This task implements `buildBatchMailto` and `sendSelectedEmails`. After this task, clicking "Send (N)" groups selected docs by customer and opens one mailto: link per customer — directly if only one customer, or via a clickable panel if multiple.

- [ ] **Step 1: Add buildBatchMailto function**

In the `<script>` block that contains `calcOdDays` and `buildMailto`, add the following function immediately after `buildMailto`:

```js
function buildBatchMailto(groups) {
  const enc   = encodeURIComponent;
  const first = groups[0];
  const to    = first.contactPerson || '';
  const code  = first.companyCode;
  const cust  = first.customer || '';

  const lines = groups.map(g => {
    const od     = calcOdDays(g.netDueDate);
    const amt    = Math.abs(Number(g.localAmt || 0))
                     .toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    const amtStr = (g.localCurr || '') + ' ' + amt;
    const date   = formatDate(g.netDueDate);
    const docNum = (g.docNum || '').startsWith('__missing__') ? '(N/A)' : (g.docNum || '');
    if (code === '0073') {
      return '　文件編號：' + docNum + '　到期日：' + date + '　逾期天數：' + od + ' 天　逾期金額：' + amtStr;
    } else if (code === '0038') {
      return '　文件编号：' + docNum + '　到期日：' + date + '　逾期天数：' + od + ' 天　逾期金额：' + amtStr;
    } else {
      return '  Doc: ' + docNum + '  Due: ' + date + '  Overdue: ' + od + ' days  Amt: ' + amtStr;
    }
  }).join('\n');

  let subject, body;

  if (code === '0073') {
    subject = '催款通知 — ' + cust;
    body    = cust + ' 您好，\n\n謹此提醒，以下應收帳款已逾期，煩請盡快安排付款：\n\n' + lines + '\n\n如有任何疑問，歡迎隨時與我聯繫。\n\n謝謝！\nHuang Fan';
  } else if (code === '0038') {
    subject = '催款通知 — ' + cust;
    body    = cust + ' 您好，\n\n谨此提醒，以下应收账款已逾期，烦请尽快安排付款：\n\n' + lines + '\n\n如有任何疑问，欢迎随时与我联系。\n\n谢谢！\nHuang Fan';
  } else {
    subject = 'Payment Reminder — ' + cust;
    body    = 'Dear ' + cust + ',\n\nThis is a reminder that the following invoices are overdue. Please arrange payment at your earliest convenience:\n\n' + lines + '\n\nPlease do not hesitate to contact me if you have any questions.\n\nBest regards,\nHuang Fan';
  }

  return 'mailto:' + enc(to) + '?subject=' + enc(subject) + '&body=' + enc(body);
}
```

- [ ] **Step 2: Add sendSelectedEmails function**

In the same `<script>` block (the one with `toggleSelect`, `toggleSelectAll`, etc.), append after `clearSelection`:

```js
function sendSelectedEmails() {
  if (selectedDocs.size === 0) return;

  // Collect selected groups from ALL data (not just currently filtered view)
  const allGroups = buildDocGroups(allRows);
  const selected  = allGroups.filter(g => selectedDocs.has(g.docNum));
  if (selected.length === 0) return;

  // Group by customer
  const byCustomer = new Map();
  selected.forEach(g => {
    const key = g.customer || '';
    if (!byCustomer.has(key)) byCustomer.set(key, []);
    byCustomer.get(key).push(g);
  });

  const entries = [...byCustomer.entries()]; // [[customer, groups[]], ...]

  if (entries.length === 1) {
    // Single customer — open directly
    window.location.href = buildBatchMailto(entries[0][1]);
  } else {
    // Multiple customers — show clickable links panel
    const panel = document.getElementById('mailtoLinksPanel');
    panel.innerHTML = entries.map(([customer, groups]) => {
      const url = buildBatchMailto(groups);
      const label = (customer || '(unknown)') + ' (' + groups.length + ')';
      return '<a class="mailto-link" href="' + url + '">✉ ' + label.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</a>';
    }).join('');
  }
}
```

- [ ] **Step 3: Verify — single customer**

Load the Excel file. Select 2–3 rows belonging to the same customer. Click "Send (N)".

Expected:
- Outlook opens with a single email
- Subject: `催款通知 — {customer}` (no document number, unlike single-doc flow)
- Body lists each selected document on its own line with live OD days
- 逾期天數 equals today minus each row's Net Due Date

- [ ] **Step 4: Verify — multiple customers**

Select rows from two different customers. Click "Send (N)".

Expected:
- Toolbar expands to show clickable links panel: one `[✉ CustomerName (N)]` link per customer
- Clicking each link opens a separate Outlook compose window for that customer
- Each email body lists only that customer's selected documents

- [ ] **Step 5: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: add batch dunning email grouped by customer with live OD days"
```

---

## Verification Checklist

After all tasks complete, confirm end-to-end:

- [ ] Per-row ✉ button now shows live OD days (not stale Excel value)
- [ ] Each work table row has a checkbox as its first cell
- [ ] Select-all header checkbox works: checks/unchecks all visible rows; shows indeterminate when partially selected
- [ ] Selecting any row shows blue toolbar with count and "Send (N)" button
- [ ] "Clear" hides toolbar and unchecks all
- [ ] Flag toggle and row color coding still work (orange border is on the flag column, not checkbox column)
- [ ] Expanding rows still works; colspan is "11"
- [ ] Single-customer selection → Outlook opens directly with multi-doc email
- [ ] Multi-customer selection → mailto links panel appears in toolbar; each link opens correct customer email
- [ ] Contact log and Export to Excel are unchanged
- [ ] Dashboard tab renders correctly
- [ ] No JS errors in browser console (F12 → Console)
