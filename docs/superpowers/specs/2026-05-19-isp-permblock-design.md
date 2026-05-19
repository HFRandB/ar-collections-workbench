---
name: isp-permblock-design
description: Design spec for SAP GUI automation script that batch-processes 1001 rows from Excel through transaction zfca0_1234 to set NGBC PermaBlock on contract accounts — May 2026
metadata:
  type: project
---

# ISP PermaBlock Automation — Design Spec

**Date:** 2026-05-19
**File:** `/Users/I306662/claude code/isp-permblock/` (new folder)

---

## 1. Overview

A Python script that reads 1001 rows from an Excel file and automates SAP GUI for Java (Mac) to run transaction `zfca0_1234`, filling in the input grid in batches, then selecting each result account and applying NGBC PermaBlock one by one. This is a one-time batch job.

---

## 2. Input

Excel file columns (in order):
| Column | Description |
|--------|-------------|
| Company Code | SAP company code |
| Business Partner | BP number |
| Contract Account ID | Contract account identifier |

1001 data rows. The script accepts the Excel file path as a command-line argument.

---

## 3. Workflow

### Phase A — Batch input (steps 1–4, repeated per batch)

1. Navigate to transaction `zfca0_1234` (type `/nzfca0_1234` in the command field, press Enter)
2. Check both checkboxes:
   - **Also Incl Recs w NO AR if Blck**
   - **Show enterd CA/BP if No AR/Blk**
3. Fill the input grid with up to `BATCH_SIZE` rows (default: 50), one row per line: company code → Tab → business partner → Tab → contract account ID → Tab → next row
4. Press **Enter** to execute — SAP returns one result account per input row

### Phase B — PermaBlock (steps 5–6, one at a time per result)

5. For each result line item:
   a. Click/select the line item
   b. Click **NGBC PermaBlock** button
   c. In the dialog, select/confirm **PermBlock**
   d. Confirm/save
   e. Log the result to `progress.csv`
6. After all results in this batch are processed, repeat Phase A for the next batch

---

## 4. Batch Strategy

- `BATCH_SIZE = 50` (configurable constant at top of script)
- 1001 rows → 21 batches (20 × 50 + 1 × 1)
- Checkboxes are checked once per batch (they reset when the transaction is re-entered)

---

## 5. Progress & Recovery

`progress.csv` is written to the same folder as the script. Schema:

| Field | Value |
|-------|-------|
| `company_code` | From Excel |
| `business_partner` | From Excel |
| `contract_account_id` | From Excel |
| `status` | `success` / `failed` / `skipped` |
| `timestamp` | ISO 8601 |
| `note` | Error description if failed |

On each run the script loads existing `progress.csv` and **skips rows already marked `success`**. This allows re-running after a timeout without re-processing completed rows.

---

## 6. Safety & Configuration

| Feature | Detail |
|---------|--------|
| **pyautogui failsafe** | Moving mouse to top-left corner of screen immediately pauses the script |
| **Operation delay** | `DELAY = 1.0` seconds between each UI action (configurable constant) |
| **`--dry-run` flag** | Logs what it would do without clicking PermaBlock |
| **SAP window focus** | Script brings SAP GUI window to foreground before each interaction using `osascript` (AppleScript via `subprocess`) — e.g. `tell application "SAP GUI" to activate` |

---

## 7. Files

| Path | Purpose |
|------|---------|
| `/Users/I306662/claude code/isp-permblock/isp_permblock.py` | Main automation script |
| `/Users/I306662/claude code/isp-permblock/requirements.txt` | `pyautogui`, `openpyxl`, `pyperclip` |
| `/Users/I306662/claude code/isp-permblock/progress.csv` | Auto-generated run log (gitignored) |

---

## 8. Usage

```bash
cd "/Users/I306662/claude code/isp-permblock"
pip install -r requirements.txt
python isp_permblock.py path/to/input.xlsx          # normal run
python isp_permblock.py path/to/input.xlsx --dry-run # dry run
```

SAP GUI for Java must be open and logged in before running the script.

---

## 9. Not in Scope

- Scheduling / recurring runs (this is a one-time job)
- SAP login automation
- Email notification on completion
