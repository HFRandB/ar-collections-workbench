# Collection Tool v2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enhance `collection-tool.html` with three features: Document Number grouping in the Work table, a per-row dunning email button (mailto:), and an inline contact log that writes back to Current Status.

**Architecture:** All changes are to a single existing file `/Users/I306662/claude code/collection-tool.html`. The Work table rendering pipeline is refactored from per-row (allRows[idx]) to per-group (docGroups[gIdx]) using a `buildDocGroups()` function. No new files are created.

**Tech Stack:** Vanilla HTML/CSS/JS (existing), SheetJS (existing), Chart.js (existing) — no new dependencies.

---

## File Map

| Action | Path |
|--------|------|
| Modify | `/Users/I306662/claude code/collection-tool.html` |

---

### Task 1: COL Constants, buildDocGroups, getFilteredGroups

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

This task adds new COL constants and the grouping infrastructure without touching the existing render pipeline. After this task the file still renders correctly using the old system; `buildDocGroups` and `getFilteredGroups` exist but are not yet called by `renderWorkTable`.

- [ ] **Step 1: Add three new COL constants**

Find the block:
```js
const COL = {
  BILL_TO_NUM:   13,
  BILL_TO_NAME:  14,
  DOC_NUM:       20,
  NET_DUE_DATE:  28,
  LAST_INT_NOTE: 50,
  OD_DAYS:       55,
  OD_CATEGORY:   56,
  OPEN_BAL_EUR:  66,
  CONTACT:       71,
  FC_MONTH:      72,
  CUST_PP:       73,
  CURR_STATUS:   74,
  AR_PLAN:       75,
  PEND_REASON:   76
};
```

Replace with:
```js
const COL = {
  COMPANY_CODE:  2,
  LOCAL_AMT:     23,
  LOCAL_CURR:    24,
  BILL_TO_NUM:   13,
  BILL_TO_NAME:  14,
  DOC_NUM:       20,
  NET_DUE_DATE:  28,
  LAST_INT_NOTE: 50,
  OD_DAYS:       55,
  OD_CATEGORY:   56,
  OPEN_BAL_EUR:  66,
  CONTACT:       71,
  FC_MONTH:      72,
  CUST_PP:       73,
  CURR_STATUS:   74,
  AR_PLAN:       75,
  PEND_REASON:   76
};
```

- [ ] **Step 2: Add expandedGroups global after allRows declaration**

Find:
```js
let allRows = []; // [{cells: [...], flag: false, expanded: false}]
```

Replace with:
```js
let allRows = []; // [{cells: [...], flag: false, expanded: false}]
const expandedGroups = new Map(); // docNum → boolean
```

- [ ] **Step 3: Add buildDocGroups and getFilteredGroups after the existing getFilteredRows/applyFilters script block**

Find the closing tag of the script block that contains `applyFilters`:
```js
function applyFilters() {
  renderWorkTable();
}
</script>
```

Replace with:
```js
function applyFilters() {
  renderWorkTable();
}
</script>
<script>
function buildDocGroups(rows) {
  const map = new Map();
  for (const row of rows) {
    const key = String(row.cells[COL.DOC_NUM] ?? '');
    if (!map.has(key)) {
      map.set(key, {
        docNum:        key,
        rows:          [],
        localAmt:      0,
        localCurr:     String(row.cells[COL.LOCAL_CURR] ?? ''),
        openEur:       0,
        odDays:        0,
        companyCode:   String(row.cells[COL.COMPANY_CODE] ?? ''),
        customer:      row.cells[COL.BILL_TO_NAME],
        netDueDate:    row.cells[COL.NET_DUE_DATE],
        odCategory:    row.cells[COL.OD_CATEGORY],
        fcMonth:       row.cells[COL.FC_MONTH],
        contactPerson: row.cells[COL.CONTACT],
        currentStatus: row.cells[COL.CURR_STATUS],
      });
    }
    const g = map.get(key);
    g.rows.push(row);
    g.localAmt += Number(row.cells[COL.LOCAL_AMT] ?? 0);
    g.openEur  += Number(row.cells[COL.OPEN_BAL_EUR] ?? 0);
    g.odDays    = Math.max(g.odDays, Number(row.cells[COL.OD_DAYS] ?? 0));
  }
  return [...map.values()];
}

function getFilteredGroups() {
  const cust    = document.getElementById('filterCustomer').value;
  const fc      = document.getElementById('filterFCMonth').value;
  const od      = document.getElementById('filterOD').value;
  const flagged = document.getElementById('filterFlagged').checked;
  return buildDocGroups(allRows).filter(g => {
    if (cust    && String(g.customer ?? '') !== cust)               return false;
    if (fc      && !g.rows.some(r => r.cells[COL.FC_MONTH] === fc)) return false;
    if (od      && !g.rows.some(r => r.cells[COL.OD_CATEGORY] === od)) return false;
    if (flagged && !g.rows.some(r => r.flag))                       return false;
    return true;
  });
}
</script>
```

- [ ] **Step 4: Verify in browser**

Open `/Users/I306662/claude code/collection-tool.html`, load the Excel file, then in the browser console run:
```js
buildDocGroups(allRows).length
```
Expected: a number less than `allRows.length` (rows are grouped — e.g. if 498 rows have 350 unique document numbers, you'll see 350).

Also verify `getFilteredGroups().length` equals `buildDocGroups(allRows).length` when no filters are active.

- [ ] **Step 5: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: add COL constants and buildDocGroups for v2 grouping"
```

---

### Task 2: Rewrite renderWorkTable and Update Helpers

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

This task switches the Work table from per-row to per-group rendering and adds the TWD and EUR columns. After this task the table shows grouped rows with both currency columns. The ✉ column header exists but the button cell is empty (added in Task 3).

- [ ] **Step 1: Add num-cell CSS in the `<style>` block**

Find the line:
```css
thead th { position: sticky; top: 0; background: #f8fafc; color: #64748b; font-weight: 600; text-align: left; padding: 10px 12px; border-bottom: 2px solid #e2e8f0; white-space: nowrap; z-index: 1; }
```

Replace with:
```css
thead th { position: sticky; top: 0; background: #f8fafc; color: #64748b; font-weight: 600; text-align: left; padding: 10px 12px; border-bottom: 2px solid #e2e8f0; white-space: nowrap; z-index: 1; }
.num-cell { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
```

- [ ] **Step 2: Update table header to 10 columns**

Find:
```html
            <th>⚑</th>
            <th>Customer</th>
            <th>Document #</th>
            <th>Net Due Date</th>
            <th>OD Days</th>
            <th>OD Category</th>
            <th>FC Month</th>
            <th>Open Balance (EUR)</th>
```

Replace with:
```html
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
```

- [ ] **Step 3: Add formatLocalAmt helper after formatEUR**

Find:
```js
function formatEUR(val) {
  if (val === null || val === undefined || val === '') return '';
  return '€' + Math.abs(Number(val)).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}
```

Replace with:
```js
function formatEUR(val) {
  if (val === null || val === undefined || val === '') return '';
  return '€' + Math.abs(Number(val)).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function formatLocalAmt(amt, curr) {
  if (amt === null || amt === undefined) return '';
  return (curr || '') + ' ' + Math.abs(Number(amt)).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}
```

- [ ] **Step 4: Rewrite rowColorClass to accept a docGroup**

Find:
```js
function rowColorClass(row) {
  const status = (row.cells[COL.CURR_STATUS] || '').toString().toLowerCase();
  if (status.includes('paid')) return 'row-green';
  const od = Number(row.cells[COL.OD_DAYS]) || 0;
  if (od > 365) return 'row-red';
  if (od > 90)  return 'row-yellow';
  return '';
}
```

Replace with:
```js
function rowColorClass(group) {
  const status = (group.currentStatus || '').toString().toLowerCase();
  if (status.includes('paid')) return 'row-green';
  if (group.odDays > 365) return 'row-red';
  if (group.odDays > 90)  return 'row-yellow';
  return '';
}
```

- [ ] **Step 5: Rewrite renderWorkTable to iterate over docGroups**

Find the entire `renderWorkTable` function:
```js
function renderWorkTable() {
  const tbody = document.getElementById('workTbody');
  tbody.innerHTML = '';
  getFilteredRows().forEach(row => {
    const idx = allRows.indexOf(row);
    const color = rowColorClass(row);
    const flagged = row.flag ? ' row-flagged' : '';

    const tr = document.createElement('tr');
    tr.className = 'data-row' + (color ? ' ' + color : '') + flagged;
    tr.dataset.idx = idx;
    const e = v => (v === null || v === undefined ? '' : String(v)).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    tr.innerHTML = `
      <td><button class="flag-btn${row.flag ? ' active' : ''}" onclick="toggleFlag(event,${idx})">⚑</button></td>
      <td>${e(row.cells[COL.BILL_TO_NAME])}</td>
      <td>${e(row.cells[COL.DOC_NUM])}</td>
      <td>${e(formatDate(row.cells[COL.NET_DUE_DATE]))}</td>
      <td>${row.cells[COL.OD_DAYS] !== null ? row.cells[COL.OD_DAYS] : ''}</td>
      <td>${e(row.cells[COL.OD_CATEGORY])}</td>
      <td>${e(row.cells[COL.FC_MONTH])}</td>
      <td>${formatEUR(row.cells[COL.OPEN_BAL_EUR])}</td>`;
    tr.addEventListener('click', e => {
      if (e.target.classList.contains('flag-btn')) return;
      toggleExpand(idx);
    });
    tbody.appendChild(tr);

    // Expand row (hidden until toggled)
    const expandTr = document.createElement('tr');
    expandTr.className = 'expand-row' + (row.expanded ? ' open' : '');
    expandTr.id = 'expand-' + idx;
    expandTr.innerHTML = `<td colspan="8">${buildExpandHTML(row, idx)}</td>`;
    tbody.appendChild(expandTr);
  });
}
```

Replace with:
```js
function renderWorkTable() {
  const tbody  = document.getElementById('workTbody');
  tbody.innerHTML = '';
  const groups = getFilteredGroups();
  const e = v => (v === null || v === undefined ? '' : String(v)).replace(/&/g, '&amp;').replace(/</g, '&lt;');

  groups.forEach((group, gIdx) => {
    const color   = rowColorClass(group);
    const flagged = group.rows.some(r => r.flag);
    const expanded = expandedGroups.get(group.docNum) || false;

    const tr = document.createElement('tr');
    tr.className = 'data-row' + (color ? ' ' + color : '') + (flagged ? ' row-flagged' : '');
    tr.dataset.gidx = gIdx;
    tr.innerHTML = `
      <td><button class="flag-btn${flagged ? ' active' : ''}" onclick="toggleFlag(event,${gIdx})">⚑</button></td>
      <td>${e(group.customer)}</td>
      <td>${e(group.docNum)}</td>
      <td>${e(formatDate(group.netDueDate))}</td>
      <td>${group.odDays || ''}</td>
      <td>${e(group.odCategory)}</td>
      <td>${e(group.fcMonth)}</td>
      <td class="num-cell">${formatLocalAmt(group.localAmt, group.localCurr)}</td>
      <td class="num-cell">${formatEUR(group.openEur)}</td>
      <td></td>`;
    tr.addEventListener('click', ev => {
      if (ev.target.classList.contains('flag-btn') || ev.target.classList.contains('email-btn')) return;
      toggleExpand(gIdx, group.docNum);
    });
    tbody.appendChild(tr);

    const expandTr = document.createElement('tr');
    expandTr.className = 'expand-row' + (expanded ? ' open' : '');
    expandTr.id = 'expand-' + gIdx;
    expandTr.innerHTML = `<td colspan="10">${buildExpandHTML(group, gIdx)}</td>`;
    tbody.appendChild(expandTr);
  });
}
```

- [ ] **Step 6: Rewrite toggleExpand, toggleFlag, and updateField**

Find the entire block:
```js
function toggleExpand(idx) {
  allRows[idx].expanded = !allRows[idx].expanded;
  const expandTr = document.getElementById('expand-' + idx);
  if (expandTr) expandTr.classList.toggle('open', allRows[idx].expanded);
}

function updateField(idx, colIdx, value) {
  allRows[idx].cells[colIdx] = value;
  if (colIdx === COL.CURR_STATUS) {
    const tr = document.querySelector(`tr.data-row[data-idx="${idx}"]`);
    if (tr) {
      tr.className = ['data-row', rowColorClass(allRows[idx]), allRows[idx].flag ? 'row-flagged' : ''].filter(Boolean).join(' ');
    }
  }
}
```

Replace with:
```js
function toggleExpand(gIdx, docNum) {
  const current = expandedGroups.get(docNum) || false;
  expandedGroups.set(docNum, !current);
  const expandTr = document.getElementById('expand-' + gIdx);
  if (expandTr) expandTr.classList.toggle('open', !current);
}

function updateField(gIdx, colIdx, value) {
  const groups = getFilteredGroups();
  const group  = groups[gIdx];
  if (!group) return;
  group.rows.forEach(r => { r.cells[colIdx] = value; });
  if (colIdx === COL.CURR_STATUS) {
    group.currentStatus = value;
    const tr = document.querySelector(`tr.data-row[data-gidx="${gIdx}"]`);
    if (tr) {
      const color   = rowColorClass(group);
      const flagged = group.rows.some(r => r.flag) ? 'row-flagged' : '';
      tr.className  = ['data-row', color, flagged].filter(Boolean).join(' ');
    }
  }
}
```

Find and replace `toggleFlag`:
```js
function toggleFlag(event, idx) {
  event.stopPropagation();
  allRows[idx].flag = !allRows[idx].flag;
  if (document.getElementById('filterFlagged').checked) {
    renderWorkTable();
    return;
  }
  const btn = event.currentTarget;
  btn.classList.toggle('active', allRows[idx].flag);
  const tr = btn.closest('tr');
  if (tr) tr.classList.toggle('row-flagged', allRows[idx].flag);
}
```

Replace with:
```js
function toggleFlag(event, gIdx) {
  event.stopPropagation();
  const groups  = getFilteredGroups();
  const group   = groups[gIdx];
  if (!group) return;
  const newFlag = !group.rows.some(r => r.flag);
  group.rows.forEach(r => { r.flag = newFlag; });
  renderWorkTable();
}
```

- [ ] **Step 7: Update buildExpandHTML signature (read from group, write to gIdx)**

Find the start of `buildExpandHTML`:
```js
function buildExpandHTML(row, idx) {
  const c = row.cells;
```

Replace with:
```js
function buildExpandHTML(group, gIdx) {
  const c = group.rows[0].cells;
```

Then find all occurrences of `updateField(${idx},` inside `buildExpandHTML` and replace with `updateField(${gIdx},`. There are 7 occurrences — replace all of them:

Find (replace_all):
```
updateField(${idx},
```
Replace with:
```
updateField(${gIdx},
```

**Important:** only apply this replace_all within the `buildExpandHTML` function body. Verify after replacing that no other functions were affected (there should be no other `updateField(${idx},` patterns in the file after this step since `idx` no longer exists in the codebase).

- [ ] **Step 8: Verify in browser**

Open `/Users/I306662/claude code/collection-tool.html` and load the Excel file.

Expected:
- Work table shows one row per unique Document Number (fewer rows than before)
- Each row has two balance columns: "Open Balance (TWD)" and "Open Balance (EUR)"
- TWD values formatted as e.g. `TWD 283,862`
- EUR values formatted as e.g. `€7,807`
- ✉ column header exists but cells are empty
- Clicking a row expands the inline edit fields; editing Current Status updates the row color
- Flag toggle marks the row with orange left border
- Filters still work correctly

- [ ] **Step 9: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: rewrite Work table to group by Document Number with TWD+EUR columns"
```

---

### Task 3: Dunning Email Button and Templates

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add email button CSS to the `<style>` block**

Find:
```css
.num-cell { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
```

Replace with:
```css
.num-cell { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.email-btn { background: #2563eb; color: #fff; border: none; border-radius: 4px; padding: 4px 8px; font-size: 0.8rem; cursor: pointer; white-space: nowrap; }
.email-btn:hover { background: #1d4ed8; }
```

- [ ] **Step 2: Replace the empty ✉ cell in renderWorkTable with an email button**

Find (inside the `tr.innerHTML` template literal in `renderWorkTable`):
```
      <td></td>`;
```

Replace with:
```
      <td><button class="email-btn" onclick="openEmail(event,${gIdx})">✉</button></td>`;
```

- [ ] **Step 3: Add buildMailto and openEmail functions**

Add the following new `<script>` block just before `</body>`:

```html
<script>
function buildMailto(group) {
  const enc     = encodeURIComponent;
  const localAmt = Number(group.localAmt || 0)
    .toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  const amtStr  = (group.localCurr || '') + ' ' + localAmt;
  const to      = group.contactPerson || '';
  const code    = group.companyCode;
  const date    = formatDate(group.netDueDate);
  const od      = group.odDays || 0;
  const cust    = group.customer || '';
  const docNum  = group.docNum || '';

  let subject, body;

  if (code === '0073') {
    subject = '催款通知 — ' + cust + ' 文件編號 ' + docNum;
    body = cust + ' 您好，\n\n謹此提醒，以下應收帳款已逾期，煩請盡快安排付款：\n\n　文件編號：' + docNum + '\n　到期日：' + date + '\n　逾期天數：' + od + ' 天\n　逾期金額：' + amtStr + '\n\n如有任何疑問，歡迎隨時與我聯繫。\n\n謝謝！\nHuang Fan';
  } else if (code === '0038') {
    subject = '催款通知 — ' + cust + ' 文件编号 ' + docNum;
    body = cust + ' 您好，\n\n谨此提醒，以下应收账款已逾期，烦请尽快安排付款：\n\n　文件编号：' + docNum + '\n　到期日：' + date + '\n　逾期天数：' + od + ' 天\n　逾期金额：' + amtStr + '\n\n如有任何疑问，欢迎随时与我联系。\n\n谢谢！\nHuang Fan';
  } else {
    subject = 'Payment Reminder — ' + cust + ' Document ' + docNum;
    body = 'Dear ' + cust + ',\n\nThis is a reminder that the following invoice is overdue. Please arrange payment at your earliest convenience:\n\n  Document Number: ' + docNum + '\n  Due Date:        ' + date + '\n  Days Overdue:    ' + od + ' days\n  Amount Due:      ' + amtStr + '\n\nPlease do not hesitate to contact me if you have any questions.\n\nBest regards,\nHuang Fan';
  }

  return 'mailto:' + enc(to) + '?subject=' + enc(subject) + '&body=' + enc(body);
}

function openEmail(event, gIdx) {
  event.stopPropagation();
  const groups = getFilteredGroups();
  const group  = groups[gIdx];
  if (!group) return;
  window.location.href = buildMailto(group);
}
</script>
```

- [ ] **Step 4: Verify in browser**

Load the Excel file. Find a row and click the ✉ button.

Expected:
- Outlook (or default mail client) opens with pre-filled fields
- **To:** Contact Person value from that row (may be empty if not set)
- **Subject:** starts with `催款通知` (since all rows are company code 0073)
- **Body:** contains the Document Number, Due Date, OD Days, and TWD amount
- Clicking the button does not expand the row (stopPropagation works)

- [ ] **Step 5: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: add dunning email button with language templates by company code"
```

---

### Task 4: Contact Log

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add contact log CSS to the `<style>` block**

Find:
```css
.email-btn { background: #2563eb; color: #fff; border: none; border-radius: 4px; padding: 4px 8px; font-size: 0.8rem; cursor: pointer; white-space: nowrap; }
.email-btn:hover { background: #1d4ed8; }
```

Replace with:
```css
.email-btn { background: #2563eb; color: #fff; border: none; border-radius: 4px; padding: 4px 8px; font-size: 0.8rem; cursor: pointer; white-space: nowrap; }
.email-btn:hover { background: #1d4ed8; }
.log-section { background: #f8fafc; border-radius: 6px; padding: 10px 12px; margin-top: 8px; }
.log-section > label { display: block; font-size: 0.8rem; font-weight: 600; color: #64748b; margin-bottom: 8px; }
.log-row { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
.log-row select { border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px 8px; font-size: 0.875rem; font-family: inherit; }
.log-row input { flex: 1; min-width: 180px; border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px 10px; font-size: 0.875rem; font-family: inherit; }
.btn-log { background: #2563eb; color: #fff; border: none; border-radius: 6px; padding: 6px 14px; font-size: 0.875rem; font-weight: 600; cursor: pointer; white-space: nowrap; }
.btn-log:hover { background: #1d4ed8; }
```

- [ ] **Step 2: Add the contact log section to buildExpandHTML**

Find the closing of the `buildExpandHTML` return string:
```js
    <div class="field-group full-width">
      <label>Last Internal Note</label>
      <textarea oninput="updateField(${gIdx},${COL.LAST_INT_NOTE},this.value)">${esc(c[COL.LAST_INT_NOTE])}</textarea>
    </div>
  </div>`;
```

Replace with:
```js
    <div class="field-group full-width">
      <label>Last Internal Note</label>
      <textarea oninput="updateField(${gIdx},${COL.LAST_INT_NOTE},this.value)">${esc(c[COL.LAST_INT_NOTE])}</textarea>
    </div>
    <div class="field-group full-width">
      <div class="log-section">
        <label>📝 記錄聯繫</label>
        <div class="log-row">
          <select id="log-method-${gIdx}">
            <option value="電話">電話</option>
            <option value="郵件">郵件</option>
            <option value="線上會議">線上會議</option>
          </select>
          <input type="text" id="log-note-${gIdx}" placeholder="聯繫備注...">
          <button class="btn-log" onclick="saveContactLog(${gIdx})">儲存記錄</button>
        </div>
      </div>
    </div>
  </div>`;
```

- [ ] **Step 3: Add saveContactLog function**

In the `<script>` block that contains `buildMailto` and `openEmail`, append the following after `openEmail`:

```js
function saveContactLog(gIdx) {
  const groups  = getFilteredGroups();
  const group   = groups[gIdx];
  if (!group) return;
  const method  = document.getElementById('log-method-' + gIdx).value;
  const note    = document.getElementById('log-note-' + gIdx).value.trim();
  if (!note) return;
  const today   = new Date().toISOString().slice(0, 10);
  const entry   = '[' + today + ' ' + method + '] ' + note;
  const existing = String(group.rows[0].cells[COL.CURR_STATUS] || '');
  const newStatus = existing ? entry + ' | ' + existing : entry;
  group.rows.forEach(r => { r.cells[COL.CURR_STATUS] = newStatus; });
  group.currentStatus = newStatus;
  renderWorkTable();
}
```

- [ ] **Step 4: Verify in browser**

Load the Excel file. Click on any row to expand it. Scroll down to the "📝 記錄聯繫" section.

Expected:
- Three fields: method dropdown (電話/郵件/線上會議), text input, 儲存記錄 button
- Type a note and click 儲存記錄
- The row re-renders; click to expand again — Current Status field now shows `[2026-05-18 電話] your note`
- If Current Status already had content, it is appended: `[date method] note | original content`
- Export to Excel and open the file — the Current Status column contains the logged entry

- [ ] **Step 5: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: add contact log section writing back to Current Status column"
```

---

## Verification

Full end-to-end checklist after all tasks complete:

- [ ] Load `collector work ar report.xlsx` — Work table shows grouped rows (fewer rows than before)
- [ ] Both currency columns visible: `TWD X,XXX` and `€X,XXX`, right-aligned
- [ ] Customer filter, FC Month filter, OD Category filter, and Flagged filter all work correctly on groups
- [ ] Row color coding still works: green = Paid in status, red = OD > 365, yellow = OD > 90
- [ ] Flag toggle marks row with orange left border; flagged state persists through filter changes
- [ ] Expanding a row shows editable fields; editing any field updates **all** underlying Excel rows in the group
- [ ] ✉ button on each row opens a mailto: link in Outlook with pre-filled subject and body in Traditional Chinese (company code 0073)
- [ ] Email body contains: customer name, document number, due date, OD days, TWD amount
- [ ] Contact log saves `[date method] note` prepended to Current Status
- [ ] Export to Excel writes Current Status log entries back correctly
- [ ] Dashboard tab (KPI cards, charts, table) is unchanged and still renders correctly
- [ ] No JS errors in browser console (F12 → Console)
