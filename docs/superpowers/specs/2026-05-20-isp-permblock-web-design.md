---
name: isp-permblock-web-design
description: Design spec for SAP WebGUI browser automation script that batch-applies NGBC PermaBlock to contract accounts via Playwright — May 2026
metadata:
  type: project
---

# ISP PermaBlock Web Automation — Design Spec

**Date:** 2026-05-20
**File:** `/Users/I306662/claude code/isp-permblock/isp_permblock_web.py`

---

## 1. Overview

A Python + Playwright script that automates the SAP WebGUI (browser-based) to batch-apply NGBC PermaBlock to contract accounts listed in `accounts.xlsx`. Replaces the previous pyautogui desktop approach with reliable browser automation.

---

## 2. URL

```
https://isp.hec.net.sap/SAP/BC/GUI/SAP/ITS/webGUI/?SAP-CLIENT=001&SAP-LANGUAGE=EN&~TRANSACTION=ZFCA0_1234#
```

No login required — the page loads directly into the transaction.

---

## 3. Input

| File | Description |
|------|-------------|
| `accounts.xlsx` | Column A: Company Code, Column B: Business Partner |

Company code is always `0038` (constant). Only the Business Partner column is used for the multiple selection input.

---

## 4. Full Workflow (per batch of 10)

### Step 1 — Open browser
Launch visible Chromium window and navigate to the SAP WebGUI URL.

### Step 2 — Enter Company Code
Find the Company Code input field and type `0038`.

### Step 3 — Enter Business Partners (multiple selection)
1. Click the multiple selection button (→ icon) next to the Business Partner field
2. In the **Multiple Selection for Business Partner** dialog, stay on **Select Single Values** tab
3. Copy the 10 BP numbers to clipboard
4. Click the **3rd icon from the right** (upload from clipboard) to paste all BP numbers at once
5. Click the **2nd icon from the left** (clock with checkmark) to transfer values and close the dialog

### Step 4 — Check 7 checkboxes
Scroll down and check only these checkboxes in the Additional Information section:
1. Include Order Block Info
2. Include Training Block Info
3. Include Payment Block Info
4. Include Vendor Block Info
5. Include "BP is Partner" Info
6. Also Incl Recs w NO AR if Blck
7. Show enterd CA/BP if NO AR/Blk

All other checkboxes remain unchecked.

### Step 5 — Execute
Click the **Execute** button at the bottom right of the screen.

### Step 6 — Apply PermaBlock (one account at a time)
For each result row:
1. Click the **checkbox** on the left of the row (turns blue when selected)
2. Click **NGBC PermaBlock** in the menu bar (6th item from left)
3. Click **Set PermaBlock** from the dropdown
4. An **Information** popup appears: "PermaBlk Action Successfully Completed" → click **Continue**
5. Log the account as `success` in `progress.csv`

### Step 7 — Pause between batches
After each batch of 10, the script pauses and asks the user to press Enter before continuing to the next batch.

---

## 5. Batch Strategy

- `BATCH_SIZE = 10`
- Process 10 accounts per run of steps 1–6
- Re-navigate to the URL for each new batch (fresh transaction screen)
- Accounts already marked `success` in `progress.csv` are skipped on re-run

---

## 6. Progress Tracking

`progress.csv` in the same folder as the script:

| Field | Value |
|-------|-------|
| `company_code` | Always `0038` |
| `business_partner` | From Excel |
| `status` | `success` / `failed` |
| `note` | Error message if failed |
| `timestamp` | ISO 8601 |

---

## 7. Implementation Order

The script is built and tested **incrementally**:

1. **Phase 1 (test first):** Steps 1–2 only — open browser, enter company code, stop. Verify it works.
2. **Phase 2:** Add steps 3–5 — multiple selection, checkboxes, execute.
3. **Phase 3:** Add step 6 — PermaBlock loop with progress logging.

---

## 8. Files

| Path | Purpose |
|------|---------|
| `/Users/I306662/claude code/isp-permblock/isp_permblock_web.py` | Main automation script |
| `/Users/I306662/claude code/isp-permblock/accounts.xlsx` | Input: company code + business partner list |
| `/Users/I306662/claude code/isp-permblock/progress.csv` | Auto-generated run log |

---

## 9. Usage

```bash
cd "/Users/I306662/claude code/isp-permblock"
pip install playwright openpyxl
playwright install chromium
python3 isp_permblock_web.py
```

---

## 10. Not in Scope

- Login automation (no login required)
- Multiple company codes (always `0038`)
- Scheduling / recurring runs
