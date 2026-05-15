# AR Collection Tool Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single self-contained `collection-tool.html` that lets Huang Fan load her AR report XLSX, manage open items on a Work tab (color coding, filters, inline editing, flagging), view KPI/chart summaries on a Dashboard tab, and export edits back to Excel.

**Architecture:** All HTML, CSS, and JS live in one file, structured as multiple `<script>` blocks (one per task) that share the same global scope. SheetJS 0.18.5 reads and writes the XLSX. Chart.js 4.4.4 renders dashboard charts. Data is parsed into an in-memory `allRows` array on file load; edits mutate that array; Export writes the modified data back to a downloadable XLSX.

**Tech Stack:** HTML5, CSS3, vanilla JS, SheetJS 0.18.5 (CDN), Chart.js 4.4.4 (CDN)

---

## File Map

| Action | Path |
|--------|------|
| Create | `/Users/I306662/claude code/collection-tool.html` |

---

### Task 1: HTML skeleton, CSS, and tab structure

**Files:**
- Create: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Create the file**

Write the following complete file. It contains all CSS and the HTML structure. A loading script block is added in Task 2.

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>AR Collection Tool — Huang Fan</title>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: #f1f5f9; color: #1e293b; }
    .page { max-width: 1400px; margin: 0 auto; padding: 16px; }

    /* Header */
    .header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; padding: 16px 20px; background: #fff; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
    .header h1 { font-size: 1.25rem; font-weight: 700; color: #0f172a; }
    .header p { font-size: 0.8rem; color: #64748b; margin-top: 2px; }
    .btn-export { padding: 8px 18px; background: #2563eb; color: #fff; border: none; border-radius: 6px; font-size: 0.875rem; font-weight: 600; cursor: pointer; display: none; }
    .btn-export:hover { background: #1d4ed8; }

    /* Drop zone */
    .drop-zone { border: 2px dashed #cbd5e1; border-radius: 10px; padding: 60px; text-align: center; background: #fff; margin-bottom: 16px; cursor: pointer; transition: border-color .2s, background .2s; }
    .drop-zone.over { border-color: #2563eb; background: #eff6ff; }
    .drop-zone .dz-icon { font-size: 2.5rem; margin-bottom: 12px; }
    .drop-zone .dz-label { font-size: 1rem; color: #475569; }
    .drop-zone .dz-sub { font-size: 0.8rem; color: #94a3b8; margin-top: 6px; }
    #fileInput { display: none; }

    /* Tabs */
    .tab-bar { display: none; margin-bottom: 16px; border-bottom: 2px solid #e2e8f0; }
    .tab-btn { padding: 10px 20px; border: none; background: none; font-size: 0.875rem; font-weight: 600; color: #64748b; cursor: pointer; border-bottom: 3px solid transparent; margin-bottom: -2px; }
    .tab-btn.active { color: #2563eb; border-bottom-color: #2563eb; }
    .tab-content { display: none; }
    .tab-content.active { display: block; }

    /* Filter bar */
    .filter-bar { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; padding: 12px 16px; background: #fff; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,.08); margin-bottom: 12px; }
    .filter-bar select { padding: 6px 10px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 0.8rem; color: #334155; background: #f8fafc; }
    .filter-bar label { font-size: 0.8rem; color: #475569; display: flex; align-items: center; gap: 6px; cursor: pointer; }

    /* Work table */
    .table-wrapper { background: #fff; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,.08); overflow: auto; max-height: 65vh; }
    table { width: 100%; border-collapse: collapse; font-size: 0.8rem; }
    thead th { position: sticky; top: 0; background: #f8fafc; color: #64748b; font-weight: 600; text-align: left; padding: 10px 12px; border-bottom: 2px solid #e2e8f0; white-space: nowrap; z-index: 1; }
    tbody tr.data-row { cursor: pointer; }
    tbody tr.data-row:hover td { filter: brightness(0.97); }
    tbody td { padding: 8px 12px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; white-space: nowrap; }

    /* Row color coding */
    tr.row-green td { background: #dcfce7; }
    tr.row-red td { background: #fee2e2; }
    tr.row-yellow td { background: #fef9c3; }
    tr.row-flagged { border-left: 4px solid #f97316; }

    /* Flag button */
    .flag-btn { background: none; border: none; cursor: pointer; font-size: 1rem; color: #cbd5e1; padding: 0; line-height: 1; }
    .flag-btn.active { color: #f97316; }

    /* Expand row */
    tr.expand-row { display: none; }
    tr.expand-row.open { display: table-row; }
    tr.expand-row td { padding: 14px 16px; background: #f8fafc; border-bottom: 2px solid #e2e8f0; }
    .expand-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
    .field-group { display: flex; flex-direction: column; gap: 4px; }
    .field-group label { font-size: 0.7rem; text-transform: uppercase; letter-spacing: .04em; color: #94a3b8; font-weight: 600; }
    .field-group input, .field-group select, .field-group textarea { padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 0.8rem; font-family: inherit; background: #fff; color: #1e293b; }
    .field-group textarea { resize: vertical; min-height: 60px; }
    .field-group.full-width { grid-column: 1 / -1; }

    /* KPI cards */
    .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 16px; }
    .kpi-card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
    .kpi-title { font-size: 0.7rem; text-transform: uppercase; letter-spacing: .05em; color: #64748b; margin-bottom: 10px; font-weight: 600; }
    .kpi-total-value { font-size: 1.8rem; font-weight: 700; color: #0f172a; }
    .kpi-row { display: flex; justify-content: space-between; align-items: center; padding: 4px 0; border-top: 1px solid #f1f5f9; margin-top: 6px; }
    .kpi-row .kpi-label { font-size: 0.75rem; color: #64748b; }
    .kpi-row .kpi-value { font-size: 0.875rem; font-weight: 600; color: #0f172a; }
    .kpi-row.upside .kpi-value { color: #f97316; }
    .kpi-row.total { border-top: 1px solid #e2e8f0; margin-top: 4px; }
    .kpi-row.total .kpi-value { font-weight: 700; }

    /* Charts */
    .chart-row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }
    .chart-full { margin-bottom: 16px; }
    .chart-card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
    .chart-card h2 { font-size: 0.875rem; font-weight: 600; color: #334155; margin-bottom: 16px; }

    /* Dashboard details table */
    .dash-table-card { background: #fff; border-radius: 10px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
    .dash-table-card h2 { font-size: 0.875rem; font-weight: 600; color: #334155; margin-bottom: 16px; }
    .dash-table { width: 100%; border-collapse: collapse; font-size: 0.875rem; }
    .dash-table th { text-align: left; padding: 8px 12px; background: #f8fafc; color: #64748b; font-weight: 600; border-bottom: 2px solid #e2e8f0; }
    .dash-table td { padding: 8px 12px; border-bottom: 1px solid #f1f5f9; }
    .dash-table tr:last-child td { border-bottom: none; }
    .risk-high { color: #f97316; font-weight: 600; }

    /* Responsive */
    @media (max-width: 900px) {
      .kpi-grid { grid-template-columns: 1fr 1fr; }
      .chart-row { grid-template-columns: 1fr; }
      .expand-grid { grid-template-columns: 1fr 1fr; }
    }
    @media (max-width: 600px) {
      .kpi-grid { grid-template-columns: 1fr; }
      .expand-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
<div class="page">

  <div class="header">
    <div>
      <h1>AR Collection Tool — Huang Fan</h1>
      <p id="fileInfo">No file loaded — drop XLSX to begin</p>
    </div>
    <button class="btn-export" id="btnExport" onclick="exportToExcel()">Export to Excel</button>
  </div>

  <div class="drop-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
    <div class="dz-icon">📂</div>
    <div class="dz-label">Drop your XLSX file here</div>
    <div class="dz-sub">or click to browse — collector work ar report.xlsx</div>
    <input type="file" id="fileInput" accept=".xlsx,.xls">
  </div>

  <div class="tab-bar" id="tabBar">
    <button class="tab-btn active" onclick="switchTab('work')">✎ Work</button>
    <button class="tab-btn" onclick="switchTab('dashboard')">Dashboard</button>
  </div>

  <div class="tab-content active" id="tab-work">
    <div class="filter-bar" id="filterBar">
      <select id="filterCustomer" onchange="applyFilters()"><option value="">All Customers</option></select>
      <select id="filterFCMonth" onchange="applyFilters()">
        <option value="">All FC Months</option>
        <option>May</option>
        <option>May-Risk</option>
        <option>Jun</option>
        <option>Jun-Risk</option>
        <option>Jul-Risk</option>
        <option>Offset/Concession/Cancel/Cleared</option>
      </select>
      <select id="filterOD" onchange="applyFilters()"><option value="">All OD Categories</option></select>
      <label><input type="checkbox" id="filterFlagged" onchange="applyFilters()"> Flagged only</label>
    </div>
    <div class="table-wrapper">
      <table id="workTable">
        <thead>
          <tr>
            <th>⚑</th>
            <th>Customer</th>
            <th>Document #</th>
            <th>Net Due Date</th>
            <th>OD Days</th>
            <th>OD Category</th>
            <th>FC Month</th>
            <th>Open Balance (EUR)</th>
          </tr>
        </thead>
        <tbody id="workTbody"></tbody>
      </table>
    </div>
  </div>

  <div class="tab-content" id="tab-dashboard">
    <div class="kpi-grid" id="kpiGrid"></div>
    <div class="chart-row">
      <div class="chart-card">
        <h2>Monthly Open Balance (EUR) — Normal vs Risk / Upside</h2>
        <canvas id="monthlyChart"></canvas>
      </div>
      <div class="chart-card">
        <h2>Top Customers by Open Balance (EUR)</h2>
        <canvas id="customerChart"></canvas>
      </div>
    </div>
    <div class="chart-full">
      <div class="chart-card">
        <h2>OD Aging Breakdown (Item Count)</h2>
        <canvas id="agingChart"></canvas>
      </div>
    </div>
    <div class="dash-table-card">
      <h2>Monthly Balance Detail</h2>
      <table class="dash-table">
        <thead>
          <tr>
            <th>Month</th>
            <th>Normal Balance (EUR)</th>
            <th>Risk / Upside Balance (EUR)</th>
            <th>% Risk</th>
          </tr>
        </thead>
        <tbody id="dashDetailTbody"></tbody>
      </table>
    </div>
  </div>

</div>
<script>
function switchTab(name) {
  document.querySelectorAll('.tab-btn').forEach((b, i) =>
    b.classList.toggle('active', ['work', 'dashboard'][i] === name));
  document.querySelectorAll('.tab-content').forEach(c =>
    c.classList.toggle('active', c.id === 'tab-' + name));
}
</script>
</body>
</html>
```

- [ ] **Step 2: Verify in browser**

Open `/Users/I306662/claude code/collection-tool.html` in a browser.  
Expected: White page with header "AR Collection Tool — Huang Fan", drop zone visible, two tabs ("✎ Work" and "Dashboard") appear but are hidden (display:none on tab-bar). No JS errors in console.

- [ ] **Step 3: Commit**

```bash
cd "/Users/I306662/claude code"
git add collection-tool.html
git commit -m "feat: add collection tool HTML skeleton and CSS"
```

---

### Task 2: XLSX data loading

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the loading script block**

Find the line `</body>` at the end of the file and insert the following `<script>` block immediately before it (after the existing `</script>` tag):

```html
<script>
// ── Column indices (0-based, from row 2 of the XLSX) ───────────────────────
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

// ── Global state ──────────────────────────────────────────────────────────────
let workbook = null;
let allRows = []; // [{cells: [...], flag: false, expanded: false}]

// ── Drop zone events ──────────────────────────────────────────────────────────
const dropZone = document.getElementById('dropZone');
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('over');
  if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
});
document.getElementById('fileInput').addEventListener('change', e => {
  if (e.target.files[0]) loadFile(e.target.files[0]);
});

function loadFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
    parseWorkbook();
    onDataLoaded(file.name);
  };
  reader.readAsArrayBuffer(file);
}

function parseWorkbook() {
  const ws = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  // Row 0 = metadata, Row 1 = headers, Row 2+ = data
  allRows = data.slice(2)
    .filter(cells => cells.some(v => v !== null && v !== ''))
    .map(cells => ({ cells, flag: false, expanded: false }));
}

function onDataLoaded(filename) {
  document.getElementById('dropZone').style.display = 'none';
  document.getElementById('tabBar').style.display = 'flex';
  document.getElementById('btnExport').style.display = 'inline-block';
  document.getElementById('fileInfo').textContent = filename + ' — ' + allRows.length + ' items loaded';
  populateFilterDropdowns();
  renderWorkTable();
  renderDashboard();
}

// Stub — replaced in Task 10
function renderDashboard() {
  if (typeof renderKPIs === 'function') renderKPIs();
  if (typeof renderCharts === 'function') renderCharts();
  if (typeof renderDetailsTable === 'function') renderDetailsTable();
}
</script>
```

- [ ] **Step 2: Verify in browser**

Reload the file. Drag `collector work ar report.xlsx` onto the drop zone.  
Expected: Drop zone disappears, tab bar appears, "Export to Excel" button appears, fileInfo text shows "collector work ar report.xlsx — 498 items loaded". No JS errors in console (there will be a console error about `renderWorkTable is not defined` — that is expected and will be fixed in Task 3).

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add XLSX loading with SheetJS"
```

---

### Task 3: Work tab — table render with row color coding

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the table rendering script block**

Find the line `</body>` and insert the following `<script>` block immediately before it (after the Task 2 `</script>` tag):

```html
<script>
// ── Formatters ────────────────────────────────────────────────────────────────
function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  return String(val);
}

function formatEUR(val) {
  if (val === null || val === undefined || val === '') return '';
  return '€' + Math.abs(Number(val)).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

// ── Row color logic ───────────────────────────────────────────────────────────
function rowColorClass(row) {
  const status = (row.cells[COL.CURR_STATUS] || '').toString().toLowerCase();
  if (status.includes('paid')) return 'row-green';
  const od = Number(row.cells[COL.OD_DAYS]) || 0;
  if (od > 365) return 'row-red';
  if (od > 90)  return 'row-yellow';
  return '';
}

// ── Table render ──────────────────────────────────────────────────────────────
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
    tr.innerHTML = `
      <td><button class="flag-btn${row.flag ? ' active' : ''}" onclick="toggleFlag(event,${idx})">⚑</button></td>
      <td>${row.cells[COL.BILL_TO_NAME] || ''}</td>
      <td>${row.cells[COL.DOC_NUM] || ''}</td>
      <td>${formatDate(row.cells[COL.NET_DUE_DATE])}</td>
      <td>${row.cells[COL.OD_DAYS] !== null ? row.cells[COL.OD_DAYS] : ''}</td>
      <td>${row.cells[COL.OD_CATEGORY] || ''}</td>
      <td>${row.cells[COL.FC_MONTH] || ''}</td>
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
</script>
```

- [ ] **Step 2: Verify in browser**

Reload the file. Drop the XLSX.  
Expected:
- Work tab shows a scrollable table with 498 rows.
- Rows where Current Status contains "paid"/"Paid" have green background.
- Rows where OD Days > 365 have red background.
- Rows where OD Days > 90 (and ≤ 365) have yellow background.
- All other rows are white.
- Console shows an error about `buildExpandHTML` not defined — expected, fixed in Task 5.

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add work tab table with row color coding"
```

---

### Task 4: Work tab — filter bar

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the filter script block**

Find `</body>` and insert immediately before it:

```html
<script>
function populateFilterDropdowns() {
  const custSel = document.getElementById('filterCustomer');
  const odSel   = document.getElementById('filterOD');

  const customers = [...new Set(allRows.map(r => r.cells[COL.BILL_TO_NAME]).filter(Boolean))].sort();
  customers.forEach(name => {
    const opt = document.createElement('option');
    opt.value = opt.textContent = name;
    custSel.appendChild(opt);
  });

  const odOrder = [
    'a.Not Due', 'b.Due 1-30 Days', 'c.Due 31-60 Days', 'd.Due 61-90 Days',
    'e.Due 91-120 Days', 'f.Due 121-180 Days', 'g.Due 181-270 Days',
    'h.Due 271-365 Days', 'i.Due 366-730 Days ', 'j.Due>730 Days'
  ];
  const present = new Set(allRows.map(r => r.cells[COL.OD_CATEGORY]).filter(Boolean));
  odOrder.filter(o => present.has(o)).forEach(o => {
    const opt = document.createElement('option');
    opt.value = opt.textContent = o;
    odSel.appendChild(opt);
  });
}

function getFilteredRows() {
  const cust    = document.getElementById('filterCustomer').value;
  const fc      = document.getElementById('filterFCMonth').value;
  const od      = document.getElementById('filterOD').value;
  const flagged = document.getElementById('filterFlagged').checked;
  return allRows.filter(row => {
    if (cust    && row.cells[COL.BILL_TO_NAME]  !== cust)    return false;
    if (fc      && row.cells[COL.FC_MONTH]       !== fc)      return false;
    if (od      && row.cells[COL.OD_CATEGORY]    !== od)      return false;
    if (flagged && !row.flag)                                  return false;
    return true;
  });
}

function applyFilters() {
  renderWorkTable();
}
</script>
```

- [ ] **Step 2: Verify in browser**

Reload and drop the XLSX.  
Expected:
- Customer dropdown is populated with 25 unique customer names.
- OD Category dropdown shows categories present in the data, in order (a→j).
- Selecting "Taiwan Semiconductor Manufacturing" from Customer shows only TSMC rows (223 rows).
- Selecting "Jun" from FC Month shows 36 rows.
- Clearing filters shows all 498 rows.

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add work tab filter bar"
```

---

### Task 5: Work tab — expandable rows with inline editing

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the expand/edit script block**

Find `</body>` and insert immediately before it:

```html
<script>
function buildExpandHTML(row, idx) {
  const c = row.cells;
  const esc = v => (v === null || v === undefined ? '' : String(v))
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

  const ppOpts = [
    'Standard (invoice- internal payment process)',
    'Acceptance before FI',
    'PO/PR process',
    'Paid'
  ].map(o => `<option value="${esc(o)}"${c[COL.CUST_PP] === o ? ' selected' : ''}>${esc(o)}</option>`).join('');

  const fcOpts = [
    'May', 'May-Risk', 'Jun', 'Jun-Risk', 'Jul-Risk',
    'Offset/Concession/Cancel/Cleared'
  ].map(o => `<option value="${esc(o)}"${c[COL.FC_MONTH] === o ? ' selected' : ''}>${esc(o)}</option>`).join('');

  return `<div class="expand-grid">
    <div class="field-group">
      <label>Current Status</label>
      <input type="text" value="${esc(c[COL.CURR_STATUS])}"
        oninput="updateField(${idx},${COL.CURR_STATUS},this.value)">
    </div>
    <div class="field-group">
      <label>Customer Payment Process</label>
      <select onchange="updateField(${idx},${COL.CUST_PP},this.value)">${ppOpts}</select>
    </div>
    <div class="field-group">
      <label>FC Month</label>
      <select onchange="updateField(${idx},${COL.FC_MONTH},this.value)">${fcOpts}</select>
    </div>
    <div class="field-group">
      <label>AR Plan</label>
      <input type="text" value="${esc(c[COL.AR_PLAN])}"
        oninput="updateField(${idx},${COL.AR_PLAN},this.value)">
    </div>
    <div class="field-group">
      <label>Pending Reason Category</label>
      <input type="text" value="${esc(c[COL.PEND_REASON])}"
        oninput="updateField(${idx},${COL.PEND_REASON},this.value)">
    </div>
    <div class="field-group">
      <label>Contact Person</label>
      <input type="text" value="${esc(c[COL.CONTACT])}"
        oninput="updateField(${idx},${COL.CONTACT},this.value)">
    </div>
    <div class="field-group full-width">
      <label>Last Internal Note</label>
      <textarea oninput="updateField(${idx},${COL.LAST_INT_NOTE},this.value)">${esc(c[COL.LAST_INT_NOTE])}</textarea>
    </div>
  </div>`;
}

function toggleExpand(idx) {
  allRows[idx].expanded = !allRows[idx].expanded;
  const expandTr = document.getElementById('expand-' + idx);
  if (expandTr) expandTr.classList.toggle('open', allRows[idx].expanded);
}

function updateField(idx, colIdx, value) {
  allRows[idx].cells[colIdx] = value;
  // Re-render only if color-affecting field changed (Current Status drives green)
  if (colIdx === COL.CURR_STATUS) {
    const tr = document.querySelector(`tr.data-row[data-idx="${idx}"]`);
    if (tr) {
      tr.className = ['data-row', rowColorClass(allRows[idx]), allRows[idx].flag ? 'row-flagged' : ''].filter(Boolean).join(' ');
    }
  }
}
</script>
```

- [ ] **Step 2: Verify in browser**

Reload and drop the XLSX.  
Expected:
- Clicking any data row reveals an expandable area with 7 editable fields.
- Clicking the same row again collapses it.
- Editing "Current Status" to "Paid" immediately changes the row background to green.
- Editing "Customer Payment Process" dropdown updates the stored value (verify by collapsing, re-expanding — value persists).

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add expandable rows with inline editing"
```

---

### Task 6: Work tab — flag toggle

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the flag script block**

Find `</body>` and insert immediately before it:

```html
<script>
function toggleFlag(event, idx) {
  event.stopPropagation();
  allRows[idx].flag = !allRows[idx].flag;
  // If "Flagged only" filter is active, re-render so unflagged rows disappear
  if (document.getElementById('filterFlagged').checked) {
    renderWorkTable();
    return;
  }
  // Otherwise update DOM directly (faster, preserves expanded states)
  const btn = event.currentTarget;
  btn.classList.toggle('active', allRows[idx].flag);
  const tr = btn.closest('tr');
  if (tr) tr.classList.toggle('row-flagged', allRows[idx].flag);
}
</script>
```

- [ ] **Step 2: Verify in browser**

Reload and drop the XLSX.  
Expected:
- Clicking ⚑ on any row turns it orange and adds an orange left border to the row.
- Clicking ⚑ again removes both.
- Checking "Flagged only" and clicking ⚑ on a row: the row appears; when unflagged it disappears from the filtered view (re-apply filters by toggling the checkbox off and on).

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add flag toggle for follow-up rows"
```

---

### Task 7: Export to Excel

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the export script block**

Find `</body>` and insert immediately before it:

```html
<script>
function exportToExcel() {
  if (!workbook) { alert('No file loaded.'); return; }
  const ws = workbook.Sheets[workbook.SheetNames[0]];

  // allRows[i] maps to 0-indexed sheet row i+2
  // (row 0 = metadata, row 1 = headers, row 2+ = data)
  allRows.forEach((row, i) => {
    const sheetRowIdx = i + 2;
    row.cells.forEach((val, colIdx) => {
      if (val === null || val === undefined) return;
      const addr = XLSX.utils.encode_cell({ r: sheetRowIdx, c: colIdx });
      if (ws[addr]) {
        ws[addr].v = val;
        if (typeof val === 'string')  ws[addr].t = 's';
        if (typeof val === 'number')  ws[addr].t = 'n';
      } else if (val !== '') {
        ws[addr] = { v: val, t: typeof val === 'number' ? 'n' : 's' };
      }
    });
  });

  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const blob  = new Blob([wbout], { type: 'application/octet-stream' });
  const url   = URL.createObjectURL(blob);
  const a     = document.createElement('a');
  a.href = url;
  a.download = 'collector-work-ar-updated.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
</script>
```

- [ ] **Step 2: Verify**

Reload and drop the XLSX. Click a row to expand it. Edit the **Current Status** field to `"Test Export"`. Click **Export to Excel**.  
Expected: Browser downloads `collector-work-ar-updated.xlsx`. Open it in Excel and confirm:
- The edited row's "Current Status" column (column CS, index 74) shows "Test Export".
- All other rows are unchanged.

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add Export to Excel with SheetJS"
```

---

### Task 8: Dashboard — KPI cards

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the KPI script block**

Find `</body>` and insert immediately before it:

```html
<script>
function renderKPIs() {
  function sumByFC(label) {
    return allRows
      .filter(r => r.cells[COL.FC_MONTH] === label)
      .reduce((s, r) => s + (Number(r.cells[COL.OPEN_BAL_EUR]) || 0), 0);
  }

  function fmtFull(v) {
    if (!v && v !== 0) return '—';
    return '€' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  }

  function fmtM(v) {
    if (!v) return '—';
    return Math.abs(v) >= 1e6
      ? '€' + (v / 1e6).toFixed(1) + 'M'
      : fmtFull(v);
  }

  const total = allRows.reduce((s, r) => s + (Number(r.cells[COL.OPEN_BAL_EUR]) || 0), 0);

  const months = [
    { label: 'May', normalFC: 'May',  riskFC: 'May-Risk'  },
    { label: 'Jun', normalFC: 'Jun',  riskFC: 'Jun-Risk'  },
    { label: 'Jul', normalFC: 'Jul',  riskFC: 'Jul-Risk'  }
  ];

  const grid = document.getElementById('kpiGrid');
  grid.innerHTML = '';

  // Total card
  grid.insertAdjacentHTML('beforeend', `
    <div class="kpi-card">
      <div class="kpi-title">Total Open Balance</div>
      <div class="kpi-total-value">${fmtM(total)}</div>
      <div class="kpi-row" style="border-top:none;margin-top:6px">
        <span class="kpi-label">EUR</span>
        <span class="kpi-value">${fmtFull(total)}</span>
      </div>
    </div>`);

  // Per-month cards
  months.forEach(m => {
    const normal  = sumByFC(m.normalFC);
    const upside  = sumByFC(m.riskFC);
    const tot     = normal + upside;
    grid.insertAdjacentHTML('beforeend', `
      <div class="kpi-card">
        <div class="kpi-title">${m.label}</div>
        <div class="kpi-row" style="border-top:none;margin-top:0">
          <span class="kpi-label">${m.normalFC}</span>
          <span class="kpi-value">${normal ? fmtFull(normal) : '—'}</span>
        </div>
        <div class="kpi-row upside">
          <span class="kpi-label">Upside</span>
          <span class="kpi-value">${upside ? fmtFull(upside) : '—'}</span>
        </div>
        <div class="kpi-row total">
          <span class="kpi-label">Total</span>
          <span class="kpi-value">${tot ? fmtFull(tot) : '—'}</span>
        </div>
      </div>`);
  });
}
</script>
```

- [ ] **Step 2: Verify in browser**

Reload and drop the XLSX. Click the **Dashboard** tab.  
Expected: Four KPI cards appear.
- Total card: "€14.8M" with "€14,822,259" below.
- May card: May ≈ €9,387,194 | Upside ≈ €170,699 | Total ≈ €9,557,893.
- Jun card: Jun ≈ €3,046,426 | Upside ≈ €3,626,056 | Total ≈ €6,672,482.
- Jul card: Jul = "—" | Upside ≈ €443,972 | Total ≈ €443,972.
- Upside row values are orange.

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add dashboard KPI cards"
```

---

### Task 9: Dashboard — three charts

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the charts script block**

Find `</body>` and insert immediately before it:

```html
<script>
let monthlyChartInst  = null;
let customerChartInst = null;
let agingChartInst    = null;

function renderCharts() {
  renderMonthlyChart();
  renderCustomerChart();
  renderAgingChart();
}

function renderMonthlyChart() {
  const labels   = ['May', 'Jun', 'Jul', 'Offset/Concession/Cancel/Cleared'];
  const riskMap  = { 'May': 'May-Risk', 'Jun': 'Jun-Risk', 'Jul': 'Jul-Risk', 'Offset/Concession/Cancel/Cleared': null };
  const sumFC    = lbl => allRows.filter(r => r.cells[COL.FC_MONTH] === lbl)
                                  .reduce((s, r) => s + (Number(r.cells[COL.OPEN_BAL_EUR]) || 0), 0);
  const normal   = labels.map(l => sumFC(l));
  const risk     = labels.map(l => riskMap[l] ? sumFC(riskMap[l]) : 0);

  if (monthlyChartInst) monthlyChartInst.destroy();
  monthlyChartInst = new Chart(document.getElementById('monthlyChart').getContext('2d'), {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Normal',        data: normal, backgroundColor: '#2563eb', borderRadius: 4 },
        { label: 'Risk / Upside', data: risk,   backgroundColor: '#f97316', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom' },
        tooltip: { callbacks: { label: ctx => ' ' + ctx.dataset.label + ': €' + Math.abs(ctx.parsed.y).toLocaleString('en-US') } }
      },
      scales: {
        x: { stacked: true },
        y: { stacked: true, ticks: { callback: v => '€' + (v / 1e6).toFixed(1) + 'M' } }
      }
    }
  });
}

function renderCustomerChart() {
  const custMap = {};
  allRows.forEach(r => {
    const name = r.cells[COL.BILL_TO_NAME];
    if (!name) return;
    custMap[name] = (custMap[name] || 0) + (Number(r.cells[COL.OPEN_BAL_EUR]) || 0);
  });
  const sorted = Object.entries(custMap).sort((a, b) => b[1] - a[1]);

  if (customerChartInst) customerChartInst.destroy();
  customerChartInst = new Chart(document.getElementById('customerChart').getContext('2d'), {
    type: 'bar',
    data: {
      labels: sorted.map(([k]) => k),
      datasets: [{ label: 'Open Balance (EUR)', data: sorted.map(([, v]) => v), backgroundColor: '#2563eb', borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' €' + ctx.parsed.x.toLocaleString('en-US') } }
      },
      scales: { x: { ticks: { callback: v => '€' + (v / 1e6).toFixed(1) + 'M' } } }
    }
  });
}

function renderAgingChart() {
  const odOrder = [
    'a.Not Due', 'b.Due 1-30 Days', 'c.Due 31-60 Days', 'd.Due 61-90 Days',
    'e.Due 91-120 Days', 'f.Due 121-180 Days', 'g.Due 181-270 Days',
    'h.Due 271-365 Days', 'i.Due 366-730 Days ', 'j.Due>730 Days'
  ];
  const counts = {};
  allRows.forEach(r => { const cat = r.cells[COL.OD_CATEGORY]; if (cat) counts[cat] = (counts[cat] || 0) + 1; });
  const labels = odOrder.filter(o => counts[o]);

  if (agingChartInst) agingChartInst.destroy();
  agingChartInst = new Chart(document.getElementById('agingChart').getContext('2d'), {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: 'Items', data: labels.map(o => counts[o]), backgroundColor: '#2563eb', borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' ' + ctx.parsed.x + ' items' } }
      },
      scales: { x: { ticks: { stepSize: 20 } } }
    }
  });
}
</script>
```

- [ ] **Step 2: Verify in browser**

Reload and drop the XLSX. Click the **Dashboard** tab.  
Expected:
- **Monthly chart** (top-left): Stacked bar with May bar mostly blue (large normal, tiny risk), Jun bar roughly equal blue and orange, Jul bar all orange.
- **Customer chart** (top-right): Horizontal bars; TSMC is the longest. Hovering shows EUR values.
- **OD Aging chart** (bottom): Horizontal bars one per OD category; "b.Due 1-30 Days" and "d.Due 61-90 Days" each show ~100+ items.

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add dashboard charts (monthly, customer, OD aging)"
```

---

### Task 10: Dashboard — details table and final wiring

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

- [ ] **Step 1: Add the details table script block**

Find `</body>` and insert immediately before it:

```html
<script>
function renderDetailsTable() {
  const rows = [
    { month: 'May',                           normalFC: 'May',                           riskFC: 'May-Risk'   },
    { month: 'Jun',                           normalFC: 'Jun',                           riskFC: 'Jun-Risk'   },
    { month: 'Jul',                           normalFC: 'Jul',                           riskFC: 'Jul-Risk'   },
    { month: 'Offset / Cancel / Cleared',     normalFC: 'Offset/Concession/Cancel/Cleared', riskFC: null      }
  ];

  function sumFC(lbl) {
    if (!lbl) return 0;
    return allRows.filter(r => r.cells[COL.FC_MONTH] === lbl)
                  .reduce((s, r) => s + (Number(r.cells[COL.OPEN_BAL_EUR]) || 0), 0);
  }

  function fmtCell(v) {
    if (!v && v !== 0) return '—';
    return v < 0 ? '(' + formatEUR(Math.abs(v)) + ')' : formatEUR(v);
  }

  const tbody = document.getElementById('dashDetailTbody');
  tbody.innerHTML = '';
  rows.forEach(row => {
    const normal  = sumFC(row.normalFC);
    const risk    = sumFC(row.riskFC);
    const total   = normal + risk;
    const pct     = total > 0 ? (risk / total) * 100 : null;
    const isHigh  = pct !== null && pct > 20;
    tbody.insertAdjacentHTML('beforeend', `
      <tr>
        <td>${row.month}</td>
        <td>${fmtCell(normal || null)}</td>
        <td>${risk ? formatEUR(risk) : '—'}</td>
        <td class="${isHigh ? 'risk-high' : ''}">${pct !== null ? pct.toFixed(1) + '%' : '—'}</td>
      </tr>`);
  });
}

// Override the Task 2 stub with the complete implementation
function renderDashboard() {
  renderKPIs();
  renderCharts();
  renderDetailsTable();
}
</script>
```

- [ ] **Step 2: Verify in browser**

Reload and drop the XLSX. Click the **Dashboard** tab.  
Expected:
- Details table shows 4 rows.
- May row: % Risk ≈ 1.8% (plain text).
- Jun row: % Risk ≈ 54.3% (orange bold).
- Jul row: % Risk = 100.0% (orange bold).
- Offset row: Normal shows "(€1,852,089)", Risk shows "—", % Risk shows "—".

- [ ] **Step 3: Commit**

```bash
git add collection-tool.html
git commit -m "feat: add dashboard details table and complete renderDashboard"
```

---

## Verification

Open `/Users/I306662/claude code/collection-tool.html` in a browser, drop `collector work ar report.xlsx` onto it, and confirm all of the following:

**Work tab:**
- [ ] Table shows 498 rows after loading
- [ ] Green rows = Current Status contains "paid/Paid", Red rows = OD > 365, Yellow rows = OD 91–365
- [ ] Customer / FC Month / OD Category filters narrow the table correctly
- [ ] "Flagged only" checkbox shows only flagged rows
- [ ] Clicking a row expands it with 7 editable fields; clicking again collapses it
- [ ] Editing Current Status to "Paid" immediately turns the row green
- [ ] ⚑ flag button turns orange and adds left border; clicking again removes both
- [ ] Flagging a row then checking "Flagged only" shows only that row

**Export:**
- [ ] Edit a field, click "Export to Excel", open downloaded file, verify edit is present

**Dashboard tab:**
- [ ] 4 KPI cards: Total | May (normal/upside/total) | Jun | Jul
- [ ] Upside values are orange; Total values are bold
- [ ] Monthly stacked bar chart renders with blue (Normal) and orange (Risk/Upside)
- [ ] Customer horizontal bar chart renders all 25 customers, TSMC longest bar
- [ ] OD Aging chart renders with items per OD category bucket
- [ ] Details table: Jun and Jul % Risk are orange; Offset normal is parenthesised

**General:**
- [ ] No errors in browser console (F12 → Console)
- [ ] Page is responsive at < 900px (KPI grid 2 columns, charts stack)
