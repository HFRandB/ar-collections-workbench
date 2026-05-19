# ISP PermaBlock Automation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Python script that reads 1001 rows from an Excel file and automates SAP GUI for Java (Mac) to run transaction `zfca0_1234`, check two checkboxes, fill the input grid in batches, then click NGBC PermaBlock on each result account one by one.

**Architecture:** Single script `isp_permblock.py` using `pyautogui` for keyboard/mouse automation on Mac, `openpyxl` for Excel reading, and a `progress.csv` file to track which rows have been processed so the script can be re-run after interruption. Image recognition (`pyautogui.locateCenterOnScreen`) is used for the NGBC PermaBlock button; keyboard Tab-navigation fills all input fields.

**Tech Stack:** Python 3.11+, pyautogui, pyperclip, openpyxl, pytest

---

## Pre-requisites (manual, before running the script)

1. SAP GUI for Java is installed and you are logged in
2. Capture a screenshot of the **NGBC PermaBlock** toolbar button (just the button, tight crop) and save it as `ngbc_permblock.png` in the `isp-permblock/` folder
3. Capture a screenshot of the **PermBlock confirm button** in the dialog that appears after clicking NGBC PermaBlock, save as `permblock_confirm.png` in the same folder

---

## File Map

| Action | Path |
|--------|------|
| Create | `/Users/I306662/claude code/isp-permblock/isp_permblock.py` |
| Create | `/Users/I306662/claude code/isp-permblock/requirements.txt` |
| Create | `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py` |
| Create | `/Users/I306662/claude code/isp-permblock/tests/fixtures/sample.xlsx` |

---

### Task 1: Project setup

**Files:**
- Create: `/Users/I306662/claude code/isp-permblock/requirements.txt`
- Create: `/Users/I306662/claude code/isp-permblock/isp_permblock.py`

- [ ] **Step 1: Create the project folder and requirements.txt**

```bash
mkdir -p "/Users/I306662/claude code/isp-permblock/tests/fixtures"
```

Write `/Users/I306662/claude code/isp-permblock/requirements.txt`:

```
pyautogui==0.9.54
pyperclip==1.9.0
openpyxl==3.1.5
pytest==8.2.0
```

- [ ] **Step 2: Create isp_permblock.py skeleton with config and imports**

Write `/Users/I306662/claude code/isp-permblock/isp_permblock.py`:

```python
import argparse
import csv
import subprocess
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import pyautogui
import pyperclip

BATCH_SIZE = 50
DELAY = 1.0
TRANSACTION = 'zfca0_1234'
SAP_APP_NAME = 'SAPGUI'
SCRIPT_DIR = Path(__file__).parent
PROGRESS_FILE = SCRIPT_DIR / 'progress.csv'
PERMBLOCK_BTN_IMG = str(SCRIPT_DIR / 'ngbc_permblock.png')
PERMBLOCK_CONFIRM_IMG = str(SCRIPT_DIR / 'permblock_confirm.png')

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1
```

- [ ] **Step 3: Install dependencies and verify**

```bash
cd "/Users/I306662/claude code/isp-permblock"
pip install -r requirements.txt
python -c "import pyautogui, pyperclip, openpyxl; print('OK')"
```

Expected: `OK`

- [ ] **Step 4: Commit**

```bash
cd "/Users/I306662/claude code"
git add isp-permblock/
git commit -m "feat: add isp-permblock project skeleton"
```

---

### Task 2: Excel reader and progress tracker

**Files:**
- Modify: `/Users/I306662/claude code/isp-permblock/isp_permblock.py`
- Create: `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`
- Create: `/Users/I306662/claude code/isp-permblock/tests/fixtures/sample.xlsx`

- [ ] **Step 1: Create the sample.xlsx fixture**

```python
# Run this once to create the fixture:
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Company Code', 'Business Partner', 'Contract Account ID'])
ws.append(['1000', 'BP001', 'CA001'])
ws.append(['1000', 'BP002', 'CA002'])
ws.append(['2000', 'BP003', 'CA003'])
wb.save('/Users/I306662/claude code/isp-permblock/tests/fixtures/sample.xlsx')
```

Run it:
```bash
cd "/Users/I306662/claude code/isp-permblock"
python -c "
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Company Code', 'Business Partner', 'Contract Account ID'])
ws.append(['1000', 'BP001', 'CA001'])
ws.append(['1000', 'BP002', 'CA002'])
ws.append(['2000', 'BP003', 'CA003'])
wb.save('tests/fixtures/sample.xlsx')
print('created')
"
```

- [ ] **Step 2: Write failing tests for read_excel and progress tracker**

Write `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`:

```python
import csv
from pathlib import Path
import pytest
import openpyxl

FIXTURE_XLSX = Path(__file__).parent / 'fixtures' / 'sample.xlsx'

# ── read_excel ────────────────────────────────────────────────────────────────

def test_read_excel_returns_list_of_dicts():
    from isp_permblock import read_excel
    rows = read_excel(str(FIXTURE_XLSX))
    assert isinstance(rows, list)
    assert len(rows) == 3

def test_read_excel_row_keys():
    from isp_permblock import read_excel
    rows = read_excel(str(FIXTURE_XLSX))
    assert rows[0] == {
        'company_code': '1000',
        'business_partner': 'BP001',
        'contract_account_id': 'CA001',
    }

def test_read_excel_all_values_are_strings():
    from isp_permblock import read_excel
    rows = read_excel(str(FIXTURE_XLSX))
    for row in rows:
        for v in row.values():
            assert isinstance(v, str)

# ── progress tracker ──────────────────────────────────────────────────────────

def test_load_progress_returns_empty_dict_when_file_missing(tmp_path):
    from isp_permblock import load_progress
    result = load_progress(str(tmp_path / 'missing.csv'))
    assert result == {}

def test_save_and_load_progress(tmp_path):
    from isp_permblock import save_progress, load_progress
    p = str(tmp_path / 'progress.csv')
    save_progress(p, '1000', 'BP001', 'CA001', 'success', '')
    save_progress(p, '1000', 'BP002', 'CA002', 'failed', 'timeout')
    loaded = load_progress(p)
    assert loaded[('1000', 'BP001', 'CA001')] == 'success'
    assert loaded[('1000', 'BP002', 'CA002')] == 'failed'

def test_filter_pending_skips_succeeded(tmp_path):
    from isp_permblock import save_progress, filter_pending
    p = str(tmp_path / 'progress.csv')
    rows = [
        {'company_code': '1000', 'business_partner': 'BP001', 'contract_account_id': 'CA001'},
        {'company_code': '1000', 'business_partner': 'BP002', 'contract_account_id': 'CA002'},
    ]
    save_progress(p, '1000', 'BP001', 'CA001', 'success', '')
    pending = filter_pending(rows, load_progress(p))
    assert len(pending) == 1
    assert pending[0]['business_partner'] == 'BP002'
```

- [ ] **Step 3: Run tests to confirm they fail**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py -v 2>&1 | head -30
```

Expected: multiple `ImportError` or `ModuleNotFoundError` failures.

- [ ] **Step 4: Implement read_excel, load_progress, save_progress, filter_pending**

Append to `/Users/I306662/claude code/isp-permblock/isp_permblock.py`:

```python

# ── Excel reader ──────────────────────────────────────────────────────────────

def read_excel(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        company_code, business_partner, contract_account_id = row[0], row[1], row[2]
        rows.append({
            'company_code': str(company_code).strip(),
            'business_partner': str(business_partner).strip(),
            'contract_account_id': str(contract_account_id).strip(),
        })
    return rows


# ── Progress tracker ──────────────────────────────────────────────────────────

def load_progress(path: str) -> dict:
    p = Path(path)
    if not p.exists():
        return {}
    result = {}
    with open(p, newline='') as f:
        for row in csv.DictReader(f):
            key = (row['company_code'], row['business_partner'], row['contract_account_id'])
            result[key] = row['status']
    return result


def save_progress(path: str, company_code: str, business_partner: str,
                  contract_account_id: str, status: str, note: str) -> None:
    p = Path(path)
    is_new = not p.exists()
    with open(p, 'a', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'company_code', 'business_partner', 'contract_account_id',
            'status', 'note', 'timestamp'
        ])
        if is_new:
            writer.writeheader()
        writer.writerow({
            'company_code': company_code,
            'business_partner': business_partner,
            'contract_account_id': contract_account_id,
            'status': status,
            'note': note,
            'timestamp': datetime.now().isoformat(timespec='seconds'),
        })


def filter_pending(rows: list[dict], progress: dict) -> list[dict]:
    return [
        r for r in rows
        if progress.get((r['company_code'], r['business_partner'], r['contract_account_id'])) != 'success'
    ]
```

- [ ] **Step 5: Run tests to confirm they pass**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py -v
```

Expected: all tests PASS.

- [ ] **Step 6: Commit**

```bash
cd "/Users/I306662/claude code"
git add isp-permblock/
git commit -m "feat: add Excel reader and progress tracker with tests"
```

---

### Task 3: SAP navigation helpers

**Files:**
- Modify: `/Users/I306662/claude code/isp-permblock/isp_permblock.py`
- Modify: `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`

- [ ] **Step 1: Write failing tests for SAP navigation (mocked)**

Append to `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`:

```python
# ── SAP navigation ────────────────────────────────────────────────────────────

from unittest.mock import patch, call

def test_focus_sap_calls_osascript():
    from isp_permblock import focus_sap, SAP_APP_NAME
    with patch('subprocess.run') as mock_run:
        with patch('time.sleep'):
            focus_sap()
    mock_run.assert_called_once_with(
        ['osascript', '-e', f'tell application "{SAP_APP_NAME}" to activate'],
        check=False
    )

def test_goto_transaction_types_command():
    from isp_permblock import goto_transaction
    with patch('pyautogui.hotkey') as mock_hotkey, \
         patch('pyautogui.press') as mock_press, \
         patch('pyperclip.copy') as mock_copy, \
         patch('time.sleep'):
        goto_transaction('zfca0_1234')
    mock_copy.assert_any_call('/nzfca0_1234')
    mock_press.assert_any_call('enter')

def test_check_checkboxes_sends_space_twice():
    from isp_permblock import check_checkboxes
    presses = []
    with patch('pyautogui.press', side_effect=lambda k: presses.append(k)), \
         patch('pyautogui.hotkey'), \
         patch('time.sleep'):
        check_checkboxes()
    assert presses.count('space') >= 2
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py::test_focus_sap_calls_osascript tests/test_isp_permblock.py::test_goto_transaction_types_command tests/test_isp_permblock.py::test_check_checkboxes_sends_space_twice -v
```

Expected: 3 failures (`ImportError` or `AttributeError`).

- [ ] **Step 3: Implement focus_sap, goto_transaction, check_checkboxes**

Append to `/Users/I306662/claude code/isp-permblock/isp_permblock.py`:

```python

# ── SAP navigation helpers ────────────────────────────────────────────────────

def focus_sap() -> None:
    subprocess.run(
        ['osascript', '-e', f'tell application "{SAP_APP_NAME}" to activate'],
        check=False
    )
    time.sleep(DELAY)


def goto_transaction(code: str) -> None:
    # Press F5 to move focus to the SAP command field, then paste /n<code>
    pyautogui.hotkey('fn', 'f5')
    time.sleep(0.3)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.2)
    pyperclip.copy(f'/n{code}')
    pyautogui.hotkey('command', 'v')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(DELAY * 2)


def check_checkboxes() -> None:
    # Tab to "Also Incl Recs w NO AR if Blck" checkbox and check it
    # Adjust tab count if needed based on the transaction's field order
    for _ in range(3):          # tab to first checkbox
        pyautogui.press('tab')
        time.sleep(0.15)
    pyautogui.press('space')    # check "Also Incl Recs w NO AR if Blck"
    time.sleep(0.2)
    pyautogui.press('tab')      # move to next checkbox
    time.sleep(0.15)
    pyautogui.press('space')    # check "Show enterd CA/BP if No AR/Blk"
    time.sleep(0.2)
```

- [ ] **Step 4: Run tests to confirm they pass**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py::test_focus_sap_calls_osascript tests/test_isp_permblock.py::test_goto_transaction_types_command tests/test_isp_permblock.py::test_check_checkboxes_sends_space_twice -v
```

Expected: 3 PASS.

- [ ] **Step 5: Commit**

```bash
cd "/Users/I306662/claude code"
git add isp-permblock/
git commit -m "feat: add SAP navigation helpers (focus, goto_transaction, check_checkboxes)"
```

---

### Task 4: Input grid filling

**Files:**
- Modify: `/Users/I306662/claude code/isp-permblock/isp_permblock.py`
- Modify: `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`

- [ ] **Step 1: Write failing tests for fill_input_grid**

Append to `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`:

```python
# ── Input grid ────────────────────────────────────────────────────────────────

def test_fill_input_grid_pastes_each_field():
    from isp_permblock import fill_input_grid
    rows = [
        {'company_code': '1000', 'business_partner': 'BP001', 'contract_account_id': 'CA001'},
        {'company_code': '2000', 'business_partner': 'BP002', 'contract_account_id': 'CA002'},
    ]
    copied = []
    with patch('pyperclip.copy', side_effect=lambda v: copied.append(v)), \
         patch('pyautogui.hotkey'), \
         patch('pyautogui.press'), \
         patch('time.sleep'):
        fill_input_grid(rows)
    assert '1000' in copied
    assert 'BP001' in copied
    assert 'CA001' in copied
    assert '2000' in copied
    assert 'BP002' in copied
    assert 'CA002' in copied

def test_fill_input_grid_tab_count():
    from isp_permblock import fill_input_grid
    rows = [{'company_code': 'X', 'business_partner': 'Y', 'contract_account_id': 'Z'}]
    tab_presses = []
    with patch('pyperclip.copy'), \
         patch('pyautogui.hotkey'), \
         patch('pyautogui.press', side_effect=lambda k: tab_presses.append(k)), \
         patch('time.sleep'):
        fill_input_grid(rows)
    # 3 tab presses per row (after each field)
    assert tab_presses.count('tab') == 3
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py::test_fill_input_grid_pastes_each_field tests/test_isp_permblock.py::test_fill_input_grid_tab_count -v
```

Expected: 2 failures.

- [ ] **Step 3: Implement fill_input_grid**

Append to `/Users/I306662/claude code/isp-permblock/isp_permblock.py`:

```python

# ── Input grid ────────────────────────────────────────────────────────────────

def fill_input_grid(rows: list[dict]) -> None:
    # Assumes focus is already in the first Company Code cell of the input grid.
    # Each row: paste company_code → Tab → business_partner → Tab → contract_account_id → Tab
    for row in rows:
        for field in ('company_code', 'business_partner', 'contract_account_id'):
            pyperclip.copy(row[field])
            pyautogui.hotkey('command', 'v')
            time.sleep(0.15)
            pyautogui.press('tab')
            time.sleep(0.15)
```

- [ ] **Step 4: Run tests to confirm they pass**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py::test_fill_input_grid_pastes_each_field tests/test_isp_permblock.py::test_fill_input_grid_tab_count -v
```

Expected: 2 PASS.

- [ ] **Step 5: Commit**

```bash
cd "/Users/I306662/claude code"
git add isp-permblock/
git commit -m "feat: add fill_input_grid with tab navigation"
```

---

### Task 5: PermaBlock action

**Files:**
- Modify: `/Users/I306662/claude code/isp-permblock/isp_permblock.py`
- Modify: `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`

- [ ] **Step 1: Write failing tests for apply_permblock**

Append to `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`:

```python
# ── PermaBlock action ─────────────────────────────────────────────────────────

def test_apply_permblock_clicks_ngbc_button(tmp_path):
    from isp_permblock import apply_permblock
    fake_btn = (100, 200)
    fake_confirm = (300, 400)
    with patch('pyautogui.locateCenterOnScreen', side_effect=[fake_btn, fake_confirm]), \
         patch('pyautogui.click') as mock_click, \
         patch('time.sleep'):
        apply_permblock(dry_run=False)
    assert mock_click.call_count >= 2

def test_apply_permblock_dry_run_does_not_click():
    from isp_permblock import apply_permblock
    with patch('pyautogui.locateCenterOnScreen') as mock_locate, \
         patch('pyautogui.click') as mock_click, \
         patch('time.sleep'):
        apply_permblock(dry_run=True)
    mock_click.assert_not_called()

def test_apply_permblock_raises_if_button_not_found():
    from isp_permblock import apply_permblock
    with patch('pyautogui.locateCenterOnScreen', return_value=None), \
         patch('time.sleep'):
        with pytest.raises(RuntimeError, match='NGBC PermaBlock button not found'):
            apply_permblock(dry_run=False)
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py::test_apply_permblock_clicks_ngbc_button tests/test_isp_permblock.py::test_apply_permblock_dry_run_does_not_click tests/test_isp_permblock.py::test_apply_permblock_raises_if_button_not_found -v
```

Expected: 3 failures.

- [ ] **Step 3: Implement apply_permblock**

Append to `/Users/I306662/claude code/isp-permblock/isp_permblock.py`:

```python

# ── PermaBlock action ─────────────────────────────────────────────────────────

def apply_permblock(dry_run: bool = False) -> None:
    if dry_run:
        print('  [dry-run] would click NGBC PermaBlock')
        return

    btn = pyautogui.locateCenterOnScreen(PERMBLOCK_BTN_IMG, confidence=0.85)
    if btn is None:
        raise RuntimeError('NGBC PermaBlock button not found on screen — '
                           'ensure SAP result screen is visible and ngbc_permblock.png is correct')
    pyautogui.click(btn)
    time.sleep(DELAY)

    confirm = pyautogui.locateCenterOnScreen(PERMBLOCK_CONFIRM_IMG, confidence=0.85)
    if confirm is None:
        raise RuntimeError('PermBlock confirm button not found — '
                           'ensure permblock_confirm.png is correct')
    pyautogui.click(confirm)
    time.sleep(DELAY)
```

- [ ] **Step 4: Run tests to confirm they pass**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py::test_apply_permblock_clicks_ngbc_button tests/test_isp_permblock.py::test_apply_permblock_dry_run_does_not_click tests/test_isp_permblock.py::test_apply_permblock_raises_if_button_not_found -v
```

Expected: 3 PASS.

- [ ] **Step 5: Commit**

```bash
cd "/Users/I306662/claude code"
git add isp-permblock/
git commit -m "feat: add apply_permblock with image recognition and dry-run support"
```

---

### Task 6: Main orchestrator and CLI

**Files:**
- Modify: `/Users/I306662/claude code/isp-permblock/isp_permblock.py`
- Modify: `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`

- [ ] **Step 1: Write failing tests for process_batch and main logic**

Append to `/Users/I306662/claude code/isp-permblock/tests/test_isp_permblock.py`:

```python
# ── Orchestrator ──────────────────────────────────────────────────────────────

def test_process_batch_logs_success(tmp_path):
    from isp_permblock import process_batch
    rows = [{'company_code': '1000', 'business_partner': 'BP001', 'contract_account_id': 'CA001'}]
    progress_file = str(tmp_path / 'progress.csv')
    with patch('isp_permblock.focus_sap'), \
         patch('isp_permblock.goto_transaction'), \
         patch('isp_permblock.check_checkboxes'), \
         patch('isp_permblock.fill_input_grid'), \
         patch('pyautogui.press'), \
         patch('isp_permblock.apply_permblock'), \
         patch('time.sleep'):
        process_batch(rows, progress_file=progress_file, dry_run=False)
    from isp_permblock import load_progress
    progress = load_progress(progress_file)
    assert progress[('1000', 'BP001', 'CA001')] == 'success'

def test_process_batch_logs_failure_on_exception(tmp_path):
    from isp_permblock import process_batch
    rows = [{'company_code': '1000', 'business_partner': 'BP001', 'contract_account_id': 'CA001'}]
    progress_file = str(tmp_path / 'progress.csv')
    with patch('isp_permblock.focus_sap'), \
         patch('isp_permblock.goto_transaction'), \
         patch('isp_permblock.check_checkboxes'), \
         patch('isp_permblock.fill_input_grid'), \
         patch('pyautogui.press'), \
         patch('isp_permblock.apply_permblock', side_effect=RuntimeError('button not found')), \
         patch('time.sleep'):
        process_batch(rows, progress_file=progress_file, dry_run=False)
    from isp_permblock import load_progress
    progress = load_progress(progress_file)
    assert progress[('1000', 'BP001', 'CA001')] == 'failed'
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/test_isp_permblock.py::test_process_batch_logs_success tests/test_isp_permblock.py::test_process_batch_logs_failure_on_exception -v
```

Expected: 2 failures.

- [ ] **Step 3: Implement process_batch and main**

Append to `/Users/I306662/claude code/isp-permblock/isp_permblock.py`:

```python

# ── Orchestrator ──────────────────────────────────────────────────────────────

def process_batch(rows: list[dict], progress_file: str, dry_run: bool) -> None:
    focus_sap()
    goto_transaction(TRANSACTION)
    check_checkboxes()
    fill_input_grid(rows)
    pyautogui.press('enter')
    time.sleep(DELAY * 2)

    for i, row in enumerate(rows):
        print(f'  [{i+1}/{len(rows)}] {row["business_partner"]} / {row["contract_account_id"]}', end=' ')
        try:
            # Select the result line item (arrow down to reach row i, then select)
            for _ in range(i):
                pyautogui.press('down')
                time.sleep(0.1)
            pyautogui.press('space')
            time.sleep(0.3)
            apply_permblock(dry_run=dry_run)
            save_progress(
                progress_file,
                row['company_code'], row['business_partner'], row['contract_account_id'],
                'success', ''
            )
            print('OK')
        except Exception as e:
            save_progress(
                progress_file,
                row['company_code'], row['business_partner'], row['contract_account_id'],
                'failed', str(e)
            )
            print(f'FAILED: {e}')


def main() -> None:
    parser = argparse.ArgumentParser(description='SAP zfca0_1234 PermaBlock automation')
    parser.add_argument('excel', help='Path to Excel file with Company Code, Business Partner, Contract Account ID')
    parser.add_argument('--dry-run', action='store_true', help='Log actions without clicking PermaBlock')
    parser.add_argument('--progress', default=str(PROGRESS_FILE), help='Path to progress CSV (default: progress.csv next to script)')
    args = parser.parse_args()

    rows = read_excel(args.excel)
    progress = load_progress(args.progress)
    pending = filter_pending(rows, progress)

    print(f'Total rows: {len(rows)}  Already done: {len(rows) - len(pending)}  Pending: {len(pending)}')
    if not pending:
        print('Nothing to do.')
        return

    if args.dry_run:
        print('[DRY RUN] No PermaBlock clicks will be made.\n')

    for batch_start in range(0, len(pending), BATCH_SIZE):
        batch = pending[batch_start:batch_start + BATCH_SIZE]
        print(f'\nBatch {batch_start // BATCH_SIZE + 1}: rows {batch_start + 1}–{batch_start + len(batch)}')
        process_batch(batch, progress_file=args.progress, dry_run=args.dry_run)

    print('\nDone.')


if __name__ == '__main__':
    main()
```

- [ ] **Step 4: Run all tests**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python -m pytest tests/ -v
```

Expected: all tests PASS.

- [ ] **Step 5: Dry-run smoke test (SAP must be open and logged in)**

```bash
cd "/Users/I306662/claude code/isp-permblock"
python isp_permblock.py tests/fixtures/sample.xlsx --dry-run
```

Expected output:
```
Total rows: 3  Already done: 0  Pending: 3

Batch 1: rows 1–3
  [1/3] BP001 / CA001 OK
  [2/3] BP002 / CA002 OK
  [3/3] BP003 / CA003 OK

Done.
```

- [ ] **Step 6: Commit**

```bash
cd "/Users/I306662/claude code"
git add isp-permblock/
git commit -m "feat: add main orchestrator with batch loop, progress tracking, and CLI"
```

---

## Verification Checklist

- [ ] `python -m pytest tests/ -v` — all tests pass
- [ ] `python isp_permblock.py <excel> --dry-run` — prints rows, no SAP clicks
- [ ] Capture `ngbc_permblock.png` and `permblock_confirm.png` reference images
- [ ] Run against 3–5 rows from real Excel, confirm PermaBlock is applied in SAP
- [ ] Run again — already-succeeded rows are skipped
- [ ] If script is interrupted mid-run, resume picks up from where it stopped
- [ ] Moving mouse to top-left corner stops the script (pyautogui failsafe)

## Calibration Notes

If SAP does not respond as expected:
- **Wrong field focus after goto_transaction:** Adjust the `fn f5` hotkey in `goto_transaction` — some Mac SAP configurations use just `F5` without `fn`
- **Checkboxes not checked:** Adjust the tab count (currently 3) in `check_checkboxes` to match how many Tab presses reach the first checkbox in `zfca0_1234`
- **Wrong line item selected:** The `process_batch` function selects results using arrow-down + space; if results use a different selection method, update those lines
- **Slow SAP:** Increase `DELAY` at the top of the script (e.g., `DELAY = 2.0`)
