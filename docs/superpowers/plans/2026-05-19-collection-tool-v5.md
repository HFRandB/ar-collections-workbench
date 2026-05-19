# Collection Tool v5 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the flat per-line format in `buildBatchMailto` with a plain-text table — column header, two divider lines, one data row per document, and a `總計` / `总计` / `Total` summary row showing the summed amount.

**Architecture:** Single function replacement in `/Users/I306662/claude code/collection-tool.html`. A new `pad(str, width)` helper is added immediately before `buildBatchMailto`. The existing `buildMailto` (single-document) is untouched. No other files change.

**Tech Stack:** Vanilla JS (existing) — no new dependencies.

---

## File Map

| Action | Path |
|--------|------|
| Modify | `/Users/I306662/claude code/collection-tool.html` |

---

### Task 1: Add pad helper and replace buildBatchMailto

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html` (around line 951)

- [ ] **Step 1: Add the pad helper immediately before buildBatchMailto**

Find:
```js
function buildBatchMailto(groups) {
```

Replace with:
```js
function pad(str, width) {
  const s = String(str == null ? '' : str);
  return s.length >= width ? s + '  ' : s + ' '.repeat(width - s.length);
}

function buildBatchMailto(groups) {
```

- [ ] **Step 2: Replace the entire buildBatchMailto body**

Find the entire function:
```js
function buildBatchMailto(groups) {
  if (!groups || groups.length === 0) return '';
  const enc   = encodeURIComponent;
  const first = groups[0];
  const to    = first.contactPerson || '';
  const code  = first.companyCode;
  const cust  = first.customer || '';

  const lines = groups.map(g => {
    const od     = g.netDueDate ? calcOdDays(g.netDueDate) : null;
    const odStr  = od != null ? String(od) : '(N/A)';
    const amt    = Math.abs(Number(g.localAmt || 0))
                     .toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    const amtStr = (g.localCurr || '') + ' ' + amt;
    const date   = formatDate(g.netDueDate);
    const docNum = (g.docNum || '').startsWith('__missing__') ? '(N/A)' : (g.docNum || '');
    if (code === '0073') {
      return '　文件編號：' + docNum + '　到期日：' + date + '　逾期天數：' + odStr + ' 天　逾期金額：' + amtStr;
    } else if (code === '0038') {
      return '　文件编号：' + docNum + '　到期日：' + date + '　逾期天数：' + odStr + ' 天　逾期金额：' + amtStr;
    } else {
      return '  Doc: ' + docNum + '  Due: ' + date + '  Overdue: ' + odStr + ' days  Amt: ' + amtStr;
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

Replace with:
```js
function buildBatchMailto(groups) {
  if (!groups || groups.length === 0) return '';
  const enc   = encodeURIComponent;
  const first = groups[0];
  const to    = first.contactPerson || '';
  const code  = first.companyCode;
  const cust  = first.customer || '';
  const DIV   = '─'.repeat(56);

  // Build data rows and accumulate total
  let total = 0;
  const dataRows = groups.map(g => {
    const od     = g.netDueDate ? calcOdDays(g.netDueDate) : null;
    const odStr  = od != null ? String(od) : '(N/A)';
    const amt    = Math.abs(Number(g.localAmt || 0));
    total += amt;
    const amtFmt = (g.localCurr || '') + ' ' + amt.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    const date   = formatDate(g.netDueDate);
    const docNum = (g.docNum || '').startsWith('__missing__') ? '(N/A)' : (g.docNum || '');
    if (code === '0073') {
      return pad(docNum, 22) + pad(date, 14) + pad(odStr + ' 天', 12) + amtFmt;
    } else if (code === '0038') {
      return pad(docNum, 22) + pad(date, 14) + pad(odStr + ' 天', 12) + amtFmt;
    } else {
      return pad(docNum, 22) + pad(date, 14) + pad(odStr + ' days', 12) + amtFmt;
    }
  });

  const totalFmt = (first.localCurr || '') + ' ' + total.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });

  let header, totalLabel;
  if (code === '0073') {
    header     = pad('文件編號', 22) + pad('到期日', 14) + pad('逾期天數', 12) + '逾期金額';
    totalLabel = '總計';
  } else if (code === '0038') {
    header     = pad('文件编号', 22) + pad('到期日', 14) + pad('逾期天数', 12) + '逾期金额';
    totalLabel = '总计';
  } else {
    header     = pad('Document #', 22) + pad('Due Date', 14) + pad('Days OD', 12) + 'Amount';
    totalLabel = 'Total';
  }

  const table = [header, DIV, ...dataRows, DIV, pad(totalLabel, 48) + totalFmt].join('\n');

  let subject, body;
  if (code === '0073') {
    subject = '催款通知 — ' + cust;
    body    = cust + ' 您好，\n\n謹此提醒，以下應收帳款已逾期，煩請盡快安排付款：\n\n' + table + '\n\n如有任何疑問，歡迎隨時與我聯繫。\n\n謝謝！\nHuang Fan';
  } else if (code === '0038') {
    subject = '催款通知 — ' + cust;
    body    = cust + ' 您好，\n\n谨此提醒，以下应收账款已逾期，烦请尽快安排付款：\n\n' + table + '\n\n如有任何疑问，欢迎随时与我联系。\n\n谢谢！\nHuang Fan';
  } else {
    subject = 'Payment Reminder — ' + cust;
    body    = 'Dear ' + cust + ',\n\nThis is a reminder that the following invoices are overdue. Please arrange payment at your earliest convenience:\n\n' + table + '\n\nPlease do not hesitate to contact me if you have any questions.\n\nBest regards,\nHuang Fan';
  }

  return 'mailto:' + enc(to) + '?subject=' + enc(subject) + '&body=' + enc(body);
}
```

- [ ] **Step 3: Verify**

Open `/Users/I306662/claude code/collection-tool.html` in Chrome. Load the Excel file. Select 2–3 rows from the same customer and click "Send (N)".

Expected email body (繁體中文 customer):
```
{customer} 您好，

謹此提醒，以下應收帳款已逾期，煩請盡快安排付款：

文件編號                到期日          逾期天數      逾期金額
────────────────────────────────────────────────────────
1234567890              15-Mar-2026     45 天         TWD 283,862
AB-2026-001             31-Jan-2026     75 天         TWD 286,138
────────────────────────────────────────────────────────
總計                                                  TWD 570,000

如有任何疑問，歡迎隨時與我聯繫。

謝謝！
Huang Fan
```

Confirm:
- Column header row present
- Each document on its own line with 4 columns
- Two divider lines (before data rows and before total)
- Total row shows `總計` (or `总计` / `Total` depending on company code) with summed amount
- No item count in the total row
- No JS errors in console

- [ ] **Step 4: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: table format with totals row in batch dunning email"
```

---

## Verification Checklist

- [ ] 繁體中文 (0073): header shows 文件編號/到期日/逾期天數/逾期金額; total row shows 總計 + amount
- [ ] 简体中文 (0038): header shows 文件编号/到期日/逾期天数/逾期金额; total row shows 总计 + amount
- [ ] English (fallback): header shows Document #/Due Date/Days OD/Amount; total row shows Total + amount
- [ ] Single document selection still works (table with one data row + total)
- [ ] Multi-customer selection: each customer's email has its own correct table and total
- [ ] Missing due date shows `(N/A)` in OD Days column (not a crash)
- [ ] `buildMailto` (per-row ✉ button) is unchanged
- [ ] No JS errors in browser console
