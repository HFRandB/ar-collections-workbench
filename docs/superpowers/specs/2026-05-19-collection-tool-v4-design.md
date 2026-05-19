---
name: collection-tool-v4-design
description: Design spec for v4 enhancement to collection-tool.html — File System Access API for one-click reload of last-opened Excel file, removing the need to manually download from SharePoint — May 2026
metadata:
  type: project
---

# Collection Tool v4 — Design Spec

**Date:** 2026-05-19  
**File:** `/Users/I306662/claude code/collection-tool.html` (enhance existing)

---

## 1. Overview

Replace the current `<input type="file">` mechanism with the browser's **File System Access API** (`showOpenFilePicker()`). The selected file handle is persisted to IndexedDB so the tool can reload the same file in one click on every subsequent open — without any Azure AD setup.

Since OneDrive syncs `sap-my.sharepoint.com` to a local folder on the user's Mac, clicking "Reload" always reads the latest version of the Excel without any manual download.

---

## 2. File Loading UI

### 2.1 No Stored Handle (first open or after clearing)

The file loading area looks and behaves the same as today, with one change: the hidden `<input type="file">` is replaced by a button that calls `showOpenFilePicker()`.

```
┌─────────────────────────────────────────────┐
│   Drag & drop Excel file here               │
│   or  [Open File]                           │
└─────────────────────────────────────────────┘
```

`[Open File]` calls `window.showOpenFilePicker({ types: [xlsx filter] })`, receives a `FileSystemFileHandle`, saves it to IndexedDB, and reads the file.

Drag-and-drop continues to work as before (it receives a `File` object directly). On a successful drag-drop load, any stored handle in IndexedDB is cleared (`clearFileHandle()`) and the reload button is hidden, so the UI stays consistent with what was actually loaded.

### 2.2 Stored Handle Present

On page load, if IndexedDB contains a stored handle, a reload button appears above the file picker area:

```
[↺ Reload  collector work ar report.xlsx]
Choose different file
```

- **"↺ Reload [filename]"** — calls `handle.getFile()` and reads the file. If permission state is `'prompt'`, the button triggers the OS one-click permission prompt first.
- **"Choose different file"** — calls `showOpenFilePicker()`, stores the new handle (replacing the old one), and reads the new file.

### 2.3 Fallback (unsupported browsers)

If `'showOpenFilePicker' in window` is `false` (Safari, older browsers), the tool falls back silently to the original `<input type="file">`. No reload button is shown. No IndexedDB operations are attempted.

---

## 3. IndexedDB Schema

| Field | Value |
|-------|-------|
| Database | `collection-tool` |
| Version | `1` |
| Object store | `fileHandles` |
| Key | `'lastFile'` |
| Value | `FileSystemFileHandle` |

The handle is stored as-is (IndexedDB natively serialises `FileSystemFileHandle`). No other fields are stored.

---

## 4. Permission Handling

On each page load, after retrieving the stored handle:

1. Call `handle.queryPermission({ mode: 'read' })`.
2. If `'granted'` → show reload button; clicking calls `handle.getFile()` directly.
3. If `'prompt'` → show reload button; clicking calls `handle.requestPermission({ mode: 'read' })` first. If the user grants, read the file. If the user denies, do nothing (leave button visible for retry).
4. If `'denied'` → clear the handle from IndexedDB and show normal file picker only.

---

## 5. New Functions

| Function | Purpose |
|----------|---------|
| `openIndexedDB()` | Opens (or creates) the `collection-tool` IndexedDB; returns a Promise resolving to the `IDBDatabase` |
| `saveFileHandle(handle)` | Stores a `FileSystemFileHandle` under key `'lastFile'` |
| `loadFileHandle()` | Returns the stored handle, or `null` if none |
| `clearFileHandle()` | Deletes the stored handle (called on `'denied'` permission) |
| `readHandleAndLoad(handle)` | Calls `handle.getFile()`, reads the `File`, passes it to the existing `parseWorkbook()` pipeline |
| `openFilePicker()` | Calls `showOpenFilePicker()`, saves the handle, reads the file |
| `initFileLoader()` | Called on `DOMContentLoaded`; checks IndexedDB, renders reload button or normal picker |

`parseWorkbook()` and `onDataLoaded()` are unchanged — they already accept a `File`/`ArrayBuffer`.

---

## 6. CSS Changes

Two new elements; minimal styling:

| Element | Style |
|---------|-------|
| `.btn-reload-file` | Same blue as `.btn-send-selected` (`#2563eb`); `font-weight: 600`; full-width on the drop zone |
| `.link-change-file` | Inline link style; `font-size: 0.8rem`; `color: #2563eb`; `text-decoration: underline`; `cursor: pointer` |

---

## 7. Unchanged from v3

- `parseWorkbook()` and all downstream logic
- Drag-and-drop zone (still works; just doesn't store a handle)
- All work table, email, selection, and contact log features

---

## 8. Browser Support

| Browser | Support |
|---------|---------|
| Chrome 86+ | Full (File System Access API) |
| Edge 86+ | Full |
| Safari (all) | Fallback (`<input type="file">`, no reload button) |
| Firefox | Fallback (API not supported as of 2026) |

---

## 9. File

Single file: `/Users/I306662/claude code/collection-tool.html`
