# Collection Tool v4 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a one-click "↺ Reload [filename]" button to `collection-tool.html` using the File System Access API so the user can reload the OneDrive-synced SharePoint Excel without opening a file picker each time.

**Architecture:** All changes are to a single existing file `/Users/I306662/claude code/collection-tool.html`. Two self-contained tasks: (1) HTML + CSS for the reload area, (2) IndexedDB helpers + picker/reload JS functions + event wiring. The existing `loadFile()` and `parseWorkbook()` pipeline are unchanged — new code feeds into them.

**Tech Stack:** Vanilla HTML/CSS/JS (existing) — no new dependencies. File System Access API (Chrome/Edge 86+); graceful fallback to `<input type="file">` for Safari/Firefox.

---

## File Map

| Action | Path |
|--------|------|
| Modify | `/Users/I306662/claude code/collection-tool.html` |

---

### Task 1: HTML and CSS Changes

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

This task adds the reload area HTML above the drop zone and the CSS that styles it, and updates `onDataLoaded` to also hide the reload area when a file finishes loading. After this task the reload area div exists in the DOM (hidden by default); the JS to show it comes in Task 2.

- [ ] **Step 1: Add CSS for the reload area**

Find:
```css
    .drop-zone .dz-sub { font-size: 0.8rem; color: #94a3b8; margin-top: 6px; }
```

Replace with:
```css
    .drop-zone .dz-sub { font-size: 0.8rem; color: #94a3b8; margin-top: 6px; }
    #reloadArea { display: none; margin-bottom: 16px; }
    .btn-reload-file { display: block; width: 100%; background: #2563eb; color: #fff; border: none; border-radius: 8px; padding: 14px 20px; font-size: 1rem; font-weight: 600; cursor: pointer; }
    .btn-reload-file:hover { background: #1d4ed8; }
    .link-change-file { font-size: 0.8rem; color: #2563eb; text-decoration: underline; cursor: pointer; background: none; border: none; padding: 0; display: block; margin-top: 8px; text-align: center; }
```

- [ ] **Step 2: Add the reload area div and remove the inline onclick from the drop zone**

Find:
```html
  <div class="drop-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
```

Replace with:
```html
  <div id="reloadArea">
    <button class="btn-reload-file" id="btnReload" onclick="reloadLastFile()">↺ Reload <span id="reloadFilename"></span></button>
    <button class="link-change-file" onclick="openFilePicker()">Choose different file</button>
  </div>
  <div class="drop-zone" id="dropZone">
```

- [ ] **Step 3: Update onDataLoaded to hide the reload area**

Find:
```js
  document.getElementById('dropZone').style.display = 'none';
```

Replace with:
```js
  document.getElementById('dropZone').style.display = 'none';
  document.getElementById('reloadArea').style.display = 'none';
```

- [ ] **Step 4: Verify**

Open `/Users/I306662/claude code/collection-tool.html` in Chrome. Confirm:
- Page loads without JS errors (F12 → Console)
- Drop zone is visible (reload area is hidden — `#reloadArea` has `display:none` by default)
- Clicking the drop zone does nothing (onclick removed; JS wiring comes in Task 2)
- Drag-dropping a file still loads it (drag-drop handler unchanged)

- [ ] **Step 5: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: add reload area HTML/CSS for File System Access API"
```

---

### Task 2: IndexedDB Helpers, Picker Functions, and Event Wiring

**Files:**
- Modify: `/Users/I306662/claude code/collection-tool.html`

This task adds all the JS: IndexedDB persistence helpers, `openFilePicker`, `readHandleAndLoad`, `reloadLastFile`, `initFileLoader`, and updates the drop zone click and drag-drop event listeners. After this task the full feature works end-to-end.

- [ ] **Step 1: Add IndexedDB helper functions**

Find:
```js
// ── Drop zone events ──────────────────────────────────────────────────────────
```

Replace with:
```js
// ── IndexedDB helpers for file handle persistence ─────────────────────────────
function openIndexedDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('collection-tool', 1);
    req.onupgradeneeded = e => e.target.result.createObjectStore('fileHandles');
    req.onsuccess = e => resolve(e.target.result);
    req.onerror   = e => reject(e.target.error);
  });
}

async function saveFileHandle(handle) {
  const db = await openIndexedDB();
  const tx = db.transaction('fileHandles', 'readwrite');
  tx.objectStore('fileHandles').put(handle, 'lastFile');
}

async function loadFileHandle() {
  try {
    const db = await openIndexedDB();
    return await new Promise((resolve, reject) => {
      const req = db.transaction('fileHandles').objectStore('fileHandles').get('lastFile');
      req.onsuccess = e => resolve(e.target.result || null);
      req.onerror   = e => reject(e.target.error);
    });
  } catch { return null; }
}

async function clearFileHandle() {
  try {
    const db = await openIndexedDB();
    const tx = db.transaction('fileHandles', 'readwrite');
    tx.objectStore('fileHandles').delete('lastFile');
  } catch { /* ignore */ }
}

// ── Drop zone events ──────────────────────────────────────────────────────────
```

- [ ] **Step 2: Update drop zone event listeners**

Find:
```js
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
```

Replace with:
```js
const dropZone = document.getElementById('dropZone');
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('over');
  if (e.dataTransfer.files[0]) {
    clearFileHandle();
    loadFile(e.dataTransfer.files[0]);
  }
});
dropZone.addEventListener('click', () => openFilePicker());
document.getElementById('fileInput').addEventListener('change', e => {
  if (e.target.files[0]) loadFile(e.target.files[0]);
});
```

- [ ] **Step 3: Add picker and init functions**

Add the following new `<script>` block immediately before `</body>`:

```html
<script>
async function openFilePicker() {
  if (!('showOpenFilePicker' in window)) {
    document.getElementById('fileInput').click();
    return;
  }
  let handle;
  try {
    [handle] = await window.showOpenFilePicker({
      types: [{ description: 'Excel files', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xls'] } }],
      multiple: false
    });
  } catch (err) {
    if (err.name !== 'AbortError') alert('Could not open file picker: ' + err.message);
    return;
  }
  await saveFileHandle(handle);
  await readHandleAndLoad(handle);
}

async function readHandleAndLoad(handle) {
  let perm = await handle.queryPermission({ mode: 'read' });
  if (perm === 'prompt') perm = await handle.requestPermission({ mode: 'read' });
  if (perm !== 'granted') return;
  const file = await handle.getFile();
  loadFile(file);
}

async function reloadLastFile() {
  const handle = await loadFileHandle();
  if (!handle) return;
  await readHandleAndLoad(handle);
}

async function initFileLoader() {
  if (!('showOpenFilePicker' in window)) return;
  const handle = await loadFileHandle();
  if (!handle) return;
  const perm = await handle.queryPermission({ mode: 'read' });
  if (perm === 'denied') { await clearFileHandle(); return; }
  document.getElementById('reloadFilename').textContent = handle.name;
  document.getElementById('reloadArea').style.display = 'block';
  document.getElementById('dropZone').style.display = 'none';
}

initFileLoader();
</script>
```

- [ ] **Step 4: Verify — first open (no stored handle)**

Open `/Users/I306662/claude code/collection-tool.html` in Chrome (new incognito window to ensure clean IndexedDB).

Expected:
- Drop zone is visible; reload area is hidden
- Clicking the drop zone opens the native OS file picker (via `showOpenFilePicker`)
- Selecting the Excel file loads it (table renders, tabs appear)
- No JS errors in console

- [ ] **Step 5: Verify — reload on subsequent open**

Without closing the tab, reload the page (Cmd+R).

Expected:
- Reload area is visible: `↺ Reload collector work ar report.xlsx`
- Drop zone is hidden
- Clicking "↺ Reload collector work ar report.xlsx" loads the file immediately (no picker)
- "Choose different file" opens the picker and, after selecting, loads the new file

- [ ] **Step 6: Verify — drag-drop clears stored handle**

With a stored handle present (reload button visible), drag-drop a different Excel file onto the page.

Expected:
- File loads successfully
- Stored handle is cleared from IndexedDB
- Reloading the page shows the drop zone (no reload button), not the old handle

- [ ] **Step 7: Verify — fallback on unsupported browser**

Open the file in Safari (or disable `showOpenFilePicker` in DevTools by overriding: `window.showOpenFilePicker = undefined`).

Expected:
- Drop zone is visible, reload area stays hidden
- Clicking the drop zone triggers the `<input type="file">` native picker (existing behavior)
- File loads as before

- [ ] **Step 8: Commit**

```bash
git -C "/Users/I306662/claude code" add collection-tool.html
git -C "/Users/I306662/claude code" commit -m "feat: File System Access API one-click reload from OneDrive-synced SharePoint file"
```

---

## Verification Checklist

After both tasks complete, confirm end-to-end in Chrome:

- [ ] First open (clean IndexedDB): drop zone shown, file picker opens on click, file loads
- [ ] After loading: `↺ Reload [filename]` button appears on page reload
- [ ] Reload button loads the file without opening any picker
- [ ] "Choose different file" link opens the picker and updates the stored handle
- [ ] Drag-drop clears the stored handle; page reload shows drop zone again
- [ ] `onDataLoaded` hides both the drop zone and reload area after any load
- [ ] Safari / Firefox: drop zone works as before, no reload button, no JS errors
- [ ] Contact log, batch email, export, dashboard all unaffected
- [ ] No JS errors in browser console (F12 → Console)
