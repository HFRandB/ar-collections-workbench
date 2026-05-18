# Email Tool Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Python/Flask web app that lets the user compose and send emails via Office 365, either immediately or on a schedule, through a browser UI at `http://localhost:5000`.

**Architecture:** Flask serves a single HTML page and exposes REST endpoints for config, send, and schedule. APScheduler runs in-process background jobs. smtplib+STARTTLS connects to `smtp.office365.com:587`. All state (credentials + scheduled jobs) lives in memory / a local `config.json`; nothing leaves the machine.

**Tech Stack:** Python 3.9+, Flask 3.x, APScheduler 3.x, smtplib (stdlib), pytest

---

## File Map

| Action | Path | Responsibility |
|--------|------|----------------|
| Create | `/Users/I306662/claude code/email-tool/requirements.txt` | Python dependencies |
| Create | `/Users/I306662/claude code/email-tool/.gitignore` | Exclude config.json |
| Create | `/Users/I306662/claude code/email-tool/app.py` | Flask app, all routes, send_email, scheduler |
| Create | `/Users/I306662/claude code/email-tool/templates/index.html` | Browser UI |
| Create | `/Users/I306662/claude code/email-tool/tests/__init__.py` | Make tests a package |
| Create | `/Users/I306662/claude code/email-tool/tests/conftest.py` | Shared fixtures |
| Create | `/Users/I306662/claude code/email-tool/tests/test_config.py` | Tests for /config endpoints |
| Create | `/Users/I306662/claude code/email-tool/tests/test_send.py` | Tests for /send endpoint |
| Create | `/Users/I306662/claude code/email-tool/tests/test_schedule.py` | Tests for /schedule + /queue endpoints |

---

### Task 1: Project Scaffold and Flask Skeleton

**Files:**
- Create: `/Users/I306662/claude code/email-tool/requirements.txt`
- Create: `/Users/I306662/claude code/email-tool/.gitignore`
- Create: `/Users/I306662/claude code/email-tool/app.py`
- Create: `/Users/I306662/claude code/email-tool/templates/index.html`
- Create: `/Users/I306662/claude code/email-tool/tests/__init__.py`
- Create: `/Users/I306662/claude code/email-tool/tests/conftest.py`

- [ ] **Step 1: Create the directory structure**

```bash
mkdir -p "/Users/I306662/claude code/email-tool/templates"
mkdir -p "/Users/I306662/claude code/email-tool/tests"
```

- [ ] **Step 2: Write requirements.txt**

```
flask>=3.0
apscheduler>=3.10
pytest>=8.0
```

Save to `/Users/I306662/claude code/email-tool/requirements.txt`.

- [ ] **Step 3: Write .gitignore**

```
config.json
__pycache__/
*.pyc
.pytest_cache/
```

Save to `/Users/I306662/claude code/email-tool/.gitignore`.

- [ ] **Step 4: Write app.py skeleton**

Write the following to `/Users/I306662/claude code/email-tool/app.py`:

```python
import json
import os
from flask import Flask, render_template

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
```

- [ ] **Step 5: Write templates/index.html placeholder**

Write the following to `/Users/I306662/claude code/email-tool/templates/index.html`:

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <title>邮件发送工具</title>
</head>
<body>
  <h1>邮件发送工具</h1>
  <p>加载中...</p>
</body>
</html>
```

- [ ] **Step 6: Write tests/__init__.py**

Create an empty file at `/Users/I306662/claude code/email-tool/tests/__init__.py`.

- [ ] **Step 7: Write tests/conftest.py**

```python
import os
import sys
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture
def client(tmp_path, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, "CONFIG_PATH", str(tmp_path / "config.json"))
    app_module.app.config["TESTING"] = True
    with app_module.app.test_client() as c:
        yield c


@pytest.fixture(scope="session", autouse=True)
def stop_scheduler():
    yield
    import app as app_module
    if hasattr(app_module, "scheduler") and app_module.scheduler.running:
        app_module.scheduler.shutdown(wait=False)
```

- [ ] **Step 8: Install dependencies**

```bash
cd "/Users/I306662/claude code/email-tool"
pip install -r requirements.txt
```

Expected: No errors. `flask`, `apscheduler`, `pytest` installed.

- [ ] **Step 9: Verify Flask runs**

```bash
cd "/Users/I306662/claude code/email-tool"
python app.py &
sleep 1
curl -s http://127.0.0.1:5000/ | grep "邮件发送工具"
kill %1
```

Expected: Output contains `邮件发送工具`.

- [ ] **Step 10: Commit**

```bash
cd "/Users/I306662/claude code/email-tool"
git -C "/Users/I306662/claude code" add email-tool/
git -C "/Users/I306662/claude code" commit -m "feat: add email tool project scaffold and Flask skeleton"
```

---

### Task 2: Config Endpoints and Tests

**Files:**
- Modify: `/Users/I306662/claude code/email-tool/app.py`
- Create: `/Users/I306662/claude code/email-tool/tests/test_config.py`

- [ ] **Step 1: Write failing tests**

Write to `/Users/I306662/claude code/email-tool/tests/test_config.py`:

```python
def test_get_config_empty(client):
    resp = client.get("/config")
    assert resp.status_code == 200
    assert resp.get_json() == {"email": ""}


def test_post_config_saves_and_get_returns_email(client):
    resp = client.post("/config", json={"email": "a@b.com", "password": "secret"})
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True

    resp2 = client.get("/config")
    assert resp2.get_json()["email"] == "a@b.com"


def test_post_config_does_not_return_password(client):
    client.post("/config", json={"email": "a@b.com", "password": "secret"})
    resp = client.get("/config")
    assert "password" not in resp.get_json()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd "/Users/I306662/claude code/email-tool"
python -m pytest tests/test_config.py -v
```

Expected: 3 FAILs (404 — routes not yet defined).

- [ ] **Step 3: Add config helpers and routes to app.py**

Replace the contents of `/Users/I306662/claude code/email-tool/app.py` with:

```python
import json
import os
from flask import Flask, jsonify, render_template, request

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

app = Flask(__name__)


def load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH) as f:
        return json.load(f)


def save_config(data):
    with open(CONFIG_PATH, "w") as f:
        json.dump(data, f)


def parse_addresses(raw):
    parts = raw.replace(";", ",").split(",")
    return [p.strip() for p in parts if p.strip()]


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/config", methods=["GET"])
def get_config():
    cfg = load_config()
    return jsonify({"email": cfg.get("email", "")})


@app.route("/config", methods=["POST"])
def post_config():
    data = request.get_json()
    save_config({"email": data.get("email", ""), "password": data.get("password", "")})
    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
cd "/Users/I306662/claude code/email-tool"
python -m pytest tests/test_config.py -v
```

Expected: 3 PASSes.

- [ ] **Step 5: Commit**

```bash
git -C "/Users/I306662/claude code" add email-tool/
git -C "/Users/I306662/claude code" commit -m "feat: add config load/save endpoints with tests"
```

---

### Task 3: Send Email Endpoint and Tests

**Files:**
- Modify: `/Users/I306662/claude code/email-tool/app.py`
- Create: `/Users/I306662/claude code/email-tool/tests/test_send.py`

- [ ] **Step 1: Write failing tests**

Write to `/Users/I306662/claude code/email-tool/tests/test_send.py`:

```python
import smtplib
from unittest.mock import MagicMock, patch


def test_send_missing_to_returns_400(client):
    resp = client.post("/send", json={"to": "", "cc": "", "subject": "Hi", "body": "Hey"})
    assert resp.status_code == 400


def test_send_no_config_returns_400(client):
    resp = client.post("/send", json={"to": "x@y.com", "cc": "", "subject": "Hi", "body": "Hey"})
    assert resp.status_code == 400
    assert "配置" in resp.get_json()["error"]


def test_send_success(client):
    client.post("/config", json={"email": "sender@b.com", "password": "pw"})
    with patch("smtplib.SMTP") as mock_smtp:
        instance = MagicMock()
        mock_smtp.return_value.__enter__ = lambda s: instance
        mock_smtp.return_value.__exit__ = MagicMock(return_value=False)
        resp = client.post("/send", json={"to": "x@y.com", "cc": "", "subject": "Hi", "body": "Hey"})
    assert resp.status_code == 200
    assert resp.get_json()["ok"] is True


def test_send_auth_failure_returns_500(client):
    client.post("/config", json={"email": "sender@b.com", "password": "wrong"})
    with patch("smtplib.SMTP") as mock_smtp:
        instance = MagicMock()
        instance.login.side_effect = smtplib.SMTPAuthenticationError(535, b"auth failed")
        mock_smtp.return_value.__enter__ = lambda s: instance
        mock_smtp.return_value.__exit__ = MagicMock(return_value=False)
        resp = client.post("/send", json={"to": "x@y.com", "cc": "", "subject": "Hi", "body": "Hey"})
    assert resp.status_code == 500
    assert "密码" in resp.get_json()["error"]
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd "/Users/I306662/claude code/email-tool"
python -m pytest tests/test_send.py -v
```

Expected: 4 FAILs (404 — `/send` not yet defined).

- [ ] **Step 3: Add send_email function and POST /send route to app.py**

Replace `/Users/I306662/claude code/email-tool/app.py` with:

```python
import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from flask import Flask, jsonify, render_template, request

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

app = Flask(__name__)


def load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH) as f:
        return json.load(f)


def save_config(data):
    with open(CONFIG_PATH, "w") as f:
        json.dump(data, f)


def parse_addresses(raw):
    parts = raw.replace(";", ",").split(",")
    return [p.strip() for p in parts if p.strip()]


def send_email(to_list, cc_list, subject, body):
    cfg = load_config()
    if not cfg.get("email") or not cfg.get("password"):
        raise ValueError("未配置发件账号，请先设置")
    msg = MIMEMultipart()
    msg["From"] = cfg["email"]
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    with smtplib.SMTP("smtp.office365.com", 587) as s:
        s.starttls()
        s.login(cfg["email"], cfg["password"])
        s.sendmail(cfg["email"], to_list + cc_list, msg.as_string())


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/config", methods=["GET"])
def get_config():
    cfg = load_config()
    return jsonify({"email": cfg.get("email", "")})


@app.route("/config", methods=["POST"])
def post_config():
    data = request.get_json()
    save_config({"email": data.get("email", ""), "password": data.get("password", "")})
    return jsonify({"ok": True})


@app.route("/send", methods=["POST"])
def post_send():
    data = request.get_json()
    to_list = parse_addresses(data.get("to", ""))
    cc_list = parse_addresses(data.get("cc", ""))
    if not to_list:
        return jsonify({"error": "收件人不能为空"}), 400
    try:
        send_email(to_list, cc_list, data["subject"], data["body"])
        return jsonify({"ok": True})
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except smtplib.SMTPAuthenticationError:
        return jsonify({"error": "账号或密码错误，请检查设置"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
```

- [ ] **Step 4: Run all tests so far**

```bash
cd "/Users/I306662/claude code/email-tool"
python -m pytest tests/test_config.py tests/test_send.py -v
```

Expected: 7 PASSes.

- [ ] **Step 5: Commit**

```bash
git -C "/Users/I306662/claude code" add email-tool/
git -C "/Users/I306662/claude code" commit -m "feat: add send_email function and POST /send endpoint with tests"
```

---

### Task 4: Scheduler Endpoints and Tests

**Files:**
- Modify: `/Users/I306662/claude code/email-tool/app.py`
- Create: `/Users/I306662/claude code/email-tool/tests/test_schedule.py`

- [ ] **Step 1: Write failing tests**

Write to `/Users/I306662/claude code/email-tool/tests/test_schedule.py`:

```python
from datetime import datetime, timedelta


def test_schedule_past_time_returns_400(client):
    resp = client.post("/schedule", json={
        "to": "x@y.com", "cc": "", "subject": "Hi", "body": "Hey",
        "run_time": "2020-01-01T10:00:00"
    })
    assert resp.status_code == 400
    assert "早于" in resp.get_json()["error"]


def test_schedule_missing_to_returns_400(client):
    future = (datetime.now() + timedelta(hours=1)).isoformat()
    resp = client.post("/schedule", json={
        "to": "", "cc": "", "subject": "Hi", "body": "Hey", "run_time": future
    })
    assert resp.status_code == 400


def test_schedule_adds_to_queue(client):
    future = (datetime.now() + timedelta(hours=1)).isoformat()
    resp = client.post("/schedule", json={
        "to": "x@y.com", "cc": "", "subject": "Sched test", "body": "Body",
        "run_time": future
    })
    assert resp.status_code == 200
    job_id = resp.get_json()["id"]

    queue = client.get("/queue").get_json()
    assert any(j["id"] == job_id for j in queue)


def test_cancel_job_removes_from_queue(client):
    future = (datetime.now() + timedelta(hours=1)).isoformat()
    resp = client.post("/schedule", json={
        "to": "x@y.com", "cc": "", "subject": "Cancel me", "body": "Body",
        "run_time": future
    })
    job_id = resp.get_json()["id"]

    del_resp = client.delete(f"/schedule/{job_id}")
    assert del_resp.status_code == 200

    queue = client.get("/queue").get_json()
    assert not any(j["id"] == job_id for j in queue)


def test_cancel_nonexistent_job_returns_404(client):
    resp = client.delete("/schedule/nonexistent-id")
    assert resp.status_code == 404
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd "/Users/I306662/claude code/email-tool"
python -m pytest tests/test_schedule.py -v
```

Expected: 5 FAILs (404 — routes not yet defined).

- [ ] **Step 3: Add APScheduler and scheduler routes to app.py**

Replace `/Users/I306662/claude code/email-tool/app.py` with the final version:

```python
import json
import os
import smtplib
import uuid
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.date import DateTrigger
from flask import Flask, jsonify, render_template, request

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

app = Flask(__name__)
scheduler = BackgroundScheduler()
scheduler.start()


def load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH) as f:
        return json.load(f)


def save_config(data):
    with open(CONFIG_PATH, "w") as f:
        json.dump(data, f)


def parse_addresses(raw):
    parts = raw.replace(";", ",").split(",")
    return [p.strip() for p in parts if p.strip()]


def send_email(to_list, cc_list, subject, body):
    cfg = load_config()
    if not cfg.get("email") or not cfg.get("password"):
        raise ValueError("未配置发件账号，请先设置")
    msg = MIMEMultipart()
    msg["From"] = cfg["email"]
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    with smtplib.SMTP("smtp.office365.com", 587) as s:
        s.starttls()
        s.login(cfg["email"], cfg["password"])
        s.sendmail(cfg["email"], to_list + cc_list, msg.as_string())


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/config", methods=["GET"])
def get_config():
    cfg = load_config()
    return jsonify({"email": cfg.get("email", "")})


@app.route("/config", methods=["POST"])
def post_config():
    data = request.get_json()
    save_config({"email": data.get("email", ""), "password": data.get("password", "")})
    return jsonify({"ok": True})


@app.route("/send", methods=["POST"])
def post_send():
    data = request.get_json()
    to_list = parse_addresses(data.get("to", ""))
    cc_list = parse_addresses(data.get("cc", ""))
    if not to_list:
        return jsonify({"error": "收件人不能为空"}), 400
    try:
        send_email(to_list, cc_list, data["subject"], data["body"])
        return jsonify({"ok": True})
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except smtplib.SMTPAuthenticationError:
        return jsonify({"error": "账号或密码错误，请检查设置"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/schedule", methods=["POST"])
def post_schedule():
    data = request.get_json()
    to_list = parse_addresses(data.get("to", ""))
    cc_list = parse_addresses(data.get("cc", ""))
    if not to_list:
        return jsonify({"error": "收件人不能为空"}), 400
    run_time = datetime.fromisoformat(data["run_time"])
    if run_time <= datetime.now():
        return jsonify({"error": "发送时间不能早于当前时间"}), 400
    job_id = str(uuid.uuid4())
    scheduler.add_job(
        send_email,
        trigger=DateTrigger(run_date=run_time),
        args=[to_list, cc_list, data["subject"], data["body"]],
        id=job_id,
        misfire_grace_time=60,
    )
    return jsonify({"ok": True, "id": job_id})


@app.route("/schedule/<job_id>", methods=["DELETE"])
def delete_schedule(job_id):
    if not scheduler.get_job(job_id):
        return jsonify({"error": "任务不存在"}), 404
    scheduler.remove_job(job_id)
    return jsonify({"ok": True})


@app.route("/queue", methods=["GET"])
def get_queue():
    jobs = []
    for job in scheduler.get_jobs():
        jobs.append({
            "id": job.id,
            "to": job.args[0],
            "subject": job.args[2],
            "run_time": job.next_run_time.isoformat() if job.next_run_time else None,
        })
    return jsonify(jobs)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
```

- [ ] **Step 4: Run all tests**

```bash
cd "/Users/I306662/claude code/email-tool"
python -m pytest tests/ -v
```

Expected: 12 PASSes (3 config + 4 send + 5 schedule).

- [ ] **Step 5: Commit**

```bash
git -C "/Users/I306662/claude code" add email-tool/
git -C "/Users/I306662/claude code" commit -m "feat: add APScheduler and schedule/queue endpoints with tests"
```

---

### Task 5: HTML Interface

**Files:**
- Modify: `/Users/I306662/claude code/email-tool/templates/index.html`

- [ ] **Step 1: Replace index.html with the complete interface**

Write the following to `/Users/I306662/claude code/email-tool/templates/index.html`:

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>邮件发送工具</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: #f1f5f9; color: #1e293b; }

    .header { background: #1e3a5f; color: #fff; padding: 16px 24px; display: flex; justify-content: space-between; align-items: center; }
    .header h1 { font-size: 1.2rem; font-weight: 600; }
    .btn-settings { background: rgba(255,255,255,.15); border: none; color: #fff; padding: 8px 16px; border-radius: 6px; cursor: pointer; font-size: 0.875rem; }
    .btn-settings:hover { background: rgba(255,255,255,.25); }

    .page { max-width: 800px; margin: 0 auto; padding: 24px 16px; }

    .card { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 1px 3px rgba(0,0,0,.08); margin-bottom: 24px; }
    .card h2 { font-size: 1rem; font-weight: 600; color: #334155; margin-bottom: 16px; }

    .field { margin-bottom: 14px; }
    .field label { display: block; font-size: 0.8rem; font-weight: 600; color: #64748b; margin-bottom: 4px; }
    .field input, .field textarea { width: 100%; border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 12px; font-size: 0.875rem; font-family: inherit; }
    .field textarea { height: 120px; resize: vertical; }
    .field input:focus, .field textarea:focus { outline: none; border-color: #2563eb; }

    .send-time { display: flex; align-items: center; gap: 16px; margin-bottom: 20px; flex-wrap: wrap; }
    .send-time label { display: flex; align-items: center; gap: 6px; font-size: 0.875rem; cursor: pointer; }
    #datetimePicker { display: none; border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 12px; font-size: 0.875rem; }

    .btn { padding: 10px 28px; border: none; border-radius: 6px; font-size: 0.875rem; font-weight: 600; cursor: pointer; }
    .btn-primary { background: #2563eb; color: #fff; }
    .btn-primary:hover { background: #1d4ed8; }
    .btn-primary:disabled { background: #93c5fd; cursor: not-allowed; }
    .btn-cancel-job { background: #dc2626; color: #fff; font-size: 0.8rem; padding: 4px 10px; border-radius: 4px; border: none; cursor: pointer; }
    .btn-cancel-job:hover { background: #b91c1c; }
    .btn-secondary { background: #f1f5f9; color: #64748b; padding: 8px 16px; border-radius: 6px; border: none; cursor: pointer; font-size: 0.875rem; }

    table { width: 100%; border-collapse: collapse; font-size: 0.875rem; }
    th { text-align: left; padding: 8px 12px; background: #f8fafc; color: #64748b; font-weight: 600; border-bottom: 2px solid #e2e8f0; }
    td { padding: 8px 12px; border-bottom: 1px solid #f1f5f9; }
    tr:last-child td { border-bottom: none; }

    .empty-queue { color: #94a3b8; font-size: 0.875rem; padding: 4px 0; }

    #toast { position: fixed; bottom: 24px; right: 24px; padding: 12px 20px; border-radius: 8px; font-size: 0.875rem; font-weight: 500; display: none; z-index: 1000; max-width: 360px; }
    #toast.success { background: #dcfce7; color: #166534; }
    #toast.error   { background: #fee2e2; color: #991b1b; }

    .modal-bg { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.4); z-index: 500; justify-content: center; align-items: center; }
    .modal-bg.open { display: flex; }
    .modal { background: #fff; border-radius: 10px; padding: 24px; width: 380px; max-width: calc(100vw - 32px); }
    .modal h3 { font-size: 1rem; font-weight: 600; margin-bottom: 16px; }
    .modal-footer { display: flex; gap: 8px; justify-content: flex-end; margin-top: 16px; }
  </style>
</head>
<body>

<div class="header">
  <h1>✉ 邮件发送工具</h1>
  <button class="btn-settings" onclick="openSettings()">⚙ 账号设置</button>
</div>

<div class="page">

  <div class="card">
    <h2>撰写邮件</h2>
    <div class="field">
      <label>收件人 <span style="color:#94a3b8;font-weight:400">（多个地址用逗号或分号分隔）</span></label>
      <input type="text" id="to" placeholder="a@company.com, b@company.com">
    </div>
    <div class="field">
      <label>抄送</label>
      <input type="text" id="cc" placeholder="可选">
    </div>
    <div class="field">
      <label>主题</label>
      <input type="text" id="subject" placeholder="邮件主题">
    </div>
    <div class="field">
      <label>正文</label>
      <textarea id="body" placeholder="邮件正文"></textarea>
    </div>
    <div class="send-time">
      <label><input type="radio" name="sendType" value="now" checked onchange="toggleDatetime(this)"> 立即发送</label>
      <label><input type="radio" name="sendType" value="scheduled" onchange="toggleDatetime(this)"> 定时发送</label>
      <input type="datetime-local" id="datetimePicker">
    </div>
    <button class="btn btn-primary" id="sendBtn" onclick="handleSend()">发 送</button>
  </div>

  <div class="card">
    <h2>定时队列</h2>
    <div id="queueContent"><p class="empty-queue">暂无待发任务</p></div>
  </div>

</div>

<!-- Settings modal -->
<div class="modal-bg" id="settingsModal">
  <div class="modal">
    <h3>⚙ 账号设置</h3>
    <div class="field">
      <label>发件邮箱</label>
      <input type="email" id="cfgEmail" placeholder="your@company.com">
    </div>
    <div class="field">
      <label>密码 / 应用专用密码</label>
      <input type="password" id="cfgPassword" placeholder="留空表示不修改">
    </div>
    <p style="font-size:0.75rem;color:#94a3b8;margin-top:8px">账号信息保存在本机 config.json，不会上传。</p>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeSettings()">取消</button>
      <button class="btn btn-primary" onclick="saveSettings()">保存</button>
    </div>
  </div>
</div>

<div id="toast"></div>

<script>
function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}

async function loadConfig() {
  const r = await fetch('/config');
  const d = await r.json();
  document.getElementById('cfgEmail').value = d.email || '';
}

function openSettings() {
  loadConfig();
  document.getElementById('settingsModal').classList.add('open');
}
function closeSettings() {
  document.getElementById('settingsModal').classList.remove('open');
}

async function saveSettings() {
  const email = document.getElementById('cfgEmail').value.trim();
  const password = document.getElementById('cfgPassword').value;
  if (!email) { showToast('请填写发件邮箱', 'error'); return; }
  const r = await fetch('/config', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password })
  });
  const d = await r.json();
  if (d.ok) { showToast('账号设置已保存', 'success'); closeSettings(); }
  else showToast(d.error || '保存失败', 'error');
}

function toggleDatetime(radio) {
  const picker = document.getElementById('datetimePicker');
  const btn = document.getElementById('sendBtn');
  if (radio.value === 'scheduled') {
    picker.style.display = 'inline-block';
    btn.textContent = '加入队列';
  } else {
    picker.style.display = 'none';
    btn.textContent = '发 送';
  }
}

async function handleSend() {
  const to      = document.getElementById('to').value.trim();
  const cc      = document.getElementById('cc').value.trim();
  const subject = document.getElementById('subject').value.trim();
  const body    = document.getElementById('body').value.trim();
  if (!to)      { showToast('请填写收件人', 'error'); return; }
  if (!subject) { showToast('请填写主题', 'error'); return; }
  if (!body)    { showToast('请填写正文', 'error'); return; }

  const isScheduled = document.querySelector('input[name="sendType"]:checked').value === 'scheduled';

  if (isScheduled) {
    const runTime = document.getElementById('datetimePicker').value;
    if (!runTime) { showToast('请选择发送时间', 'error'); return; }
    const r = await fetch('/schedule', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ to, cc, subject, body, run_time: runTime })
    });
    const d = await r.json();
    if (d.ok) { showToast('已加入定时队列', 'success'); refreshQueue(); }
    else showToast(d.error || '添加失败', 'error');
  } else {
    const btn = document.getElementById('sendBtn');
    btn.disabled = true;
    const r = await fetch('/send', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ to, cc, subject, body })
    });
    btn.disabled = false;
    const d = await r.json();
    if (d.ok) showToast('邮件已发送 ✓', 'success');
    else showToast(d.error || '发送失败', 'error');
  }
}

async function cancelJob(id) {
  const r = await fetch('/schedule/' + id, { method: 'DELETE' });
  const d = await r.json();
  if (d.ok) { showToast('已取消', 'success'); refreshQueue(); }
  else showToast(d.error || '取消失败', 'error');
}

async function refreshQueue() {
  const r = await fetch('/queue');
  const jobs = await r.json();
  const el = document.getElementById('queueContent');
  if (!jobs.length) {
    el.innerHTML = '<p class="empty-queue">暂无待发任务</p>';
    return;
  }
  const rows = jobs.map(j => {
    const dt    = new Date(j.run_time).toLocaleString('zh-CN');
    const toStr = Array.isArray(j.to) ? j.to.join(', ') : j.to;
    return `<tr>
      <td>${esc(toStr)}</td>
      <td>${esc(j.subject)}</td>
      <td>${dt}</td>
      <td><button class="btn-cancel-job" onclick="cancelJob('${esc(j.id)}')">取消</button></td>
    </tr>`;
  }).join('');
  el.innerHTML = `<table>
    <thead><tr><th>收件人</th><th>主题</th><th>发送时间</th><th>操作</th></tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function showToast(msg, type) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = type;
  t.style.display = 'block';
  clearTimeout(t._timer);
  t._timer = setTimeout(() => { t.style.display = 'none'; }, 3000);
}

refreshQueue();
setInterval(refreshQueue, 5000);
</script>
</body>
</html>
```

- [ ] **Step 2: Run all tests to ensure nothing broke**

```bash
cd "/Users/I306662/claude code/email-tool"
python -m pytest tests/ -v
```

Expected: 12 PASSes.

- [ ] **Step 3: Verify in browser**

```bash
cd "/Users/I306662/claude code/email-tool"
python app.py
```

Open `http://127.0.0.1:5000` in a browser and confirm:
- Dark blue header with title and ⚙ button
- Compose form with To, CC, Subject, Body fields
- "立即发送 / 定时发送" radio buttons; selecting "定时发送" shows datetime picker and changes button to "加入队列"
- Clicking ⚙ opens the settings modal
- "定时队列" card shows "暂无待发任务"
- No JS errors in browser console (F12 → Console)

- [ ] **Step 4: Commit**

```bash
git -C "/Users/I306662/claude code" add email-tool/
git -C "/Users/I306662/claude code" commit -m "feat: add complete HTML interface with compose form, settings modal, and queue table"
```

---

## Verification

After all tasks complete, run the full test suite and do a manual end-to-end check:

```bash
cd "/Users/I306662/claude code/email-tool"
python -m pytest tests/ -v
```

Expected: 12 PASSes, 0 failures.

Manual browser checklist:
- [ ] Page loads at `http://127.0.0.1:5000`
- [ ] ⚙ modal saves email/password to `config.json`
- [ ] Sending with empty To shows "请填写收件人" toast
- [ ] Scheduling with a past time returns error toast "发送时间不能早于当前时间"
- [ ] Scheduling with a future time adds a row to the queue table
- [ ] Cancelling a queued job removes it from the table
- [ ] No JS errors in console
