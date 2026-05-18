---
name: email-tool-design
description: Design spec for a local Python/Flask email sending tool with scheduled send, Outlook/Office 365 SMTP, and HTML browser interface — May 2026
metadata:
  type: project
---

# 邮件发送工具 — Design Spec

**Date:** 2026-05-18  
**Output:** `/Users/I306662/claude code/email-tool/` directory  
**Run:** `python app.py` → open `http://localhost:5000`

---

## 1. Overview

A local Python + Flask web application for composing and sending emails via Office 365 Outlook. The browser-based interface allows the user to write emails, send immediately, or schedule them for a future time. Scheduled jobs are held in memory and execute while the program is running.

---

## 2. File Structure

```
email-tool/
├── app.py              ← Flask server, SMTP logic, APScheduler jobs
├── config.json         ← Saved sender email + password (local only, gitignored)
├── requirements.txt    ← flask, apscheduler
└── templates/
    └── index.html      ← Single-page browser interface
```

---

## 3. Architecture & Data Flow

- **Flask** serves `index.html` and exposes three API endpoints:
  - `POST /send` — send immediately
  - `POST /schedule` — add a scheduled job
  - `DELETE /schedule/<job_id>` — cancel a scheduled job
  - `GET /queue` — return current scheduled job list
  - `GET /config` — return saved sender settings (email only, not password)
  - `POST /config` — save sender settings to `config.json`
- **APScheduler** (`BackgroundScheduler`) holds in-memory jobs; cleared on restart
- **smtplib** + STARTTLS connects to `smtp.office365.com:587` to send mail
- `config.json` stores sender email and password in plain text (local file, never committed)

---

## 4. Interface Layout

Single page, three sections:

### 4.1 Header
- Left: title "✉ 邮件发送工具"
- Right: ⚙ 账号设置 button — opens a modal to enter/save email credentials

### 4.2 Compose Form
| Field | Type | Notes |
|-------|------|-------|
| 收件人 | Text input | Multiple addresses separated by `,` or `;` |
| 抄送 | Text input | Multiple addresses, optional |
| 主题 | Text input | Required |
| 正文 | Textarea | Plain text, required |
| 发送时间 | Radio + datetime-local | "立即发送" (default) or "定时发送" with date/time picker |
| 发送按钮 | Button | Label: "发送" when immediate, "加入队列" when scheduled |

### 4.3 定时队列 (Scheduled Queue)
Table showing pending jobs with columns:
- 收件人 | 主题 | 发送时间 | 操作（取消）

Hidden when queue is empty.

### 4.4 账号设置 Modal
Fields: 发件邮箱 (email), 密码 (password, type=password). Save button writes to `config.json`.

---

## 5. Email Sending

- SMTP server: `smtp.office365.com`, port `587`, STARTTLS
- Sender credentials loaded from `config.json` at send time
- `To` and `CC` fields split on `,` and `;`, stripped of whitespace
- If `config.json` missing or credentials empty → return error to frontend with message "请先设置发件账号"
- On SMTP failure → return HTTP 500 with the error message displayed in the UI

---

## 6. Scheduled Sending

- Uses `APScheduler.BackgroundScheduler` with `DateTrigger` (one-shot, fires once at specified datetime)
- Job added via `POST /schedule`; stored with a UUID job ID
- `GET /queue` returns list of `{id, to, subject, run_time}` for display in the queue table
- `DELETE /schedule/<job_id>` removes the job before it fires
- If the scheduled time has already passed when submitted → return error "发送时间不能早于当前时间"
- After a scheduled job fires (success or failure), it is removed from the queue automatically

---

## 7. Settings Storage

`config.json` format:
```json
{
  "email": "user@company.com",
  "password": "app-password-here"
}
```

- File is created on first save; overwritten on subsequent saves
- Should be added to `.gitignore`
- Password stored as plain text — acceptable for a local personal tool

---

## 8. Technology Stack

| Concern | Choice |
|---------|--------|
| Web server | Flask |
| Scheduling | APScheduler 3.x |
| Email | smtplib (stdlib) + STARTTLS |
| Frontend | Single `index.html` (Jinja2 template) |
| Styling | Inline CSS, system sans-serif font |
| SMTP server | smtp.office365.com:587 |

`requirements.txt`:
```
flask>=3.0
apscheduler>=3.10
```

---

## 9. Styling

| Element | Color |
|---------|-------|
| Header background | Dark blue `#1e3a5f` |
| Send button | Blue `#2563eb` |
| Cancel button | Red `#dc2626` |
| Success toast | Green `#dcfce7` |
| Error toast | Red `#fee2e2` |
| Background | Light gray `#f1f5f9` |
| Card background | White `#ffffff` |

---

## 10. Error Handling

- No credentials configured → show inline warning, block send
- SMTP auth failure → show toast "账号或密码错误，请检查设置"
- SMTP send failure → show toast with error detail
- Invalid email address format → basic client-side validation before submit
- Scheduled time in the past → show inline error on the form
- All errors shown as dismissible toast notifications (3 seconds auto-hide)

---

## 11. Output Location

Directory: `/Users/I306662/claude code/email-tool/`
