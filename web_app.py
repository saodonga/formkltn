#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
web_app.py — Web API & Server cho CheckForm KLTN
=================================================
Chạy:  python web_app.py
Mở:    http://localhost:5000
"""

import os
import sys
import json
import uuid
import tempfile
import threading
from pathlib import Path
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory, Response
import queue

# Import các thư viện (đã khai báo trong requirements.txt)

# Import checker engine
sys.path.insert(0, str(Path(__file__).parent))
from check_format_kltn import check_file, CheckResult, export_excel

app = Flask(__name__, static_folder="web_static", static_url_path="/static")
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB upload limit

# ── Cấu hình ────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config_kltn.json"
UPLOAD_DIR  = Path(tempfile.gettempdir()) / "kltn_uploads"
RESULT_DIR  = BASE_DIR / "web_results"
UPLOAD_DIR.mkdir(exist_ok=True)
RESULT_DIR.mkdir(exist_ok=True)

# In-memory job store
_jobs: dict[str, dict] = {}
_jobs_lock = threading.Lock()


# ════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════
def _load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"advisors": [], "_title_min_length": 50}


def _save_config(cfg: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _result_to_dict(res: CheckResult) -> dict:
    """Chuyển CheckResult thành dict JSON-serializable."""
    issues = []
    for iss in res.issues:
        issues.append({
            "category":   iss.category,
            "severity":   iss.severity,
            "message":    iss.message,
            "location":   iss.location,
            "suggestion": iss.suggestion,
        })
    error_count = sum(1 for i in res.issues if i.severity == "ERROR")
    warn_count  = sum(1 for i in res.issues if i.severity == "WARNING")
    info_count  = sum(1 for i in res.issues if i.severity == "INFO")
    return {
        "filename":     Path(res.filepath).name,
        "student_name": res.student_name,
        "student_id":   res.student_id,
        "title":        res.title,
        "advisor":      res.advisor,
        "score":        res.score,
        "letter_grade": res.letter_grade,
        "error_count":  error_count,
        "warn_count":   warn_count,
        "info_count":   info_count,
        "issues":       issues,
    }


# ════════════════════════════════════════════════════════════════
#  ROUTES — STATIC
# ════════════════════════════════════════════════════════════════
@app.route("/")
def index():
    return send_from_directory("web_static", "index.html")


# ════════════════════════════════════════════════════════════════
#  ROUTES — API
# ════════════════════════════════════════════════════════════════

# ── GET /api/config ──────────────────────────────────────────────
@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify(_load_config())


# ── POST /api/config ─────────────────────────────────────────────
@app.route("/api/config", methods=["POST"])
def set_config():
    data = request.get_json(force=True)
    cfg = _load_config()
    if "advisors" in data:
        cfg["advisors"] = [a.strip() for a in data["advisors"] if a.strip()]
    if "_title_min_length" in data:
        cfg["_title_min_length"] = int(data["_title_min_length"])
    _save_config(cfg)
    return jsonify({"ok": True, "count": len(cfg["advisors"])})


# ── POST /api/check ──────────────────────────────────────────────
@app.route("/api/check", methods=["POST"])
def check():
    """
    Nhận một hoặc nhiều file .docx, kiểm tra và trả kết quả JSON.
    Với file lớn / nhiều file → tạo job và stream SSE.
    """
    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "Không có file nào được gửi"}), 400

    # Tạo job ID
    job_id = str(uuid.uuid4())[:8]
    saved_paths = []
    for f in files:
        if not f.filename.lower().endswith(".docx"):
            continue
        safe_name = f"kltn_{job_id}_{uuid.uuid4().hex[:6]}_{Path(f.filename).name}"
        dest = UPLOAD_DIR / safe_name
        f.save(str(dest))
        saved_paths.append((dest, f.filename))

    if not saved_paths:
        return jsonify({"error": "Chỉ chấp nhận file .docx"}), 400

    # Nếu chỉ 1 file → xử lý đồng bộ (nhanh)
    if len(saved_paths) == 1:
        dest, orig_name = saved_paths[0]
        try:
            res = check_file(str(dest))
            res.filepath = orig_name  # Dùng tên gốc
            data = _result_to_dict(res)
            data["filename"] = orig_name
        except Exception as e:
            data = {"filename": orig_name, "error": str(e)}
        finally:
            try:
                dest.unlink()
            except Exception:
                pass
        return jsonify({"job_id": job_id, "total": 1, "results": [data]})

    # Nhiều file → chạy background, client dùng SSE
    with _jobs_lock:
        _jobs[job_id] = {
            "total":   len(saved_paths),
            "done":    0,
            "results": [],
            "finished": False,
            "queue":   queue.Queue(),
        }

    def _worker():
        job = _jobs[job_id]
        for dest, orig_name in saved_paths:
            try:
                res = check_file(str(dest))
                res.filepath = orig_name
                d = _result_to_dict(res)
                d["filename"] = orig_name
            except Exception as e:
                d = {"filename": orig_name, "error": str(e)}
            finally:
                try:
                    dest.unlink()
                except Exception:
                    pass
            with _jobs_lock:
                job["done"] += 1
                job["results"].append(d)
                job["queue"].put(d)
        with _jobs_lock:
            job["finished"] = True
            job["queue"].put(None)  # sentinel

    threading.Thread(target=_worker, daemon=True).start()
    return jsonify({"job_id": job_id, "total": len(saved_paths), "streaming": True})


# ── GET /api/stream/<job_id> ─────────────────────────────────────
@app.route("/api/stream/<job_id>")
def stream(job_id):
    """Server-Sent Events — cập nhật real-time từng file."""
    def _generate():
        job = _jobs.get(job_id)
        if not job:
            yield f"data: {json.dumps({'error': 'Job not found'})}\n\n"
            return
        q = job["queue"]
        while True:
            item = q.get()
            if item is None:
                yield f"data: {json.dumps({'done': True, 'total': job['total']})}\n\n"
                break
            yield f"data: {json.dumps({'result': item, 'done_count': job['done'], 'total': job['total']})}\n\n"

    return Response(_generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ── GET /api/results/<job_id> ────────────────────────────────────
@app.route("/api/results/<job_id>")
def get_results(job_id):
    job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "job_id":   job_id,
        "total":    job["total"],
        "done":     job["done"],
        "finished": job["finished"],
        "results":  job["results"],
    })


# ── POST /api/export ─────────────────────────────────────────────
@app.route("/api/export", methods=["POST"])
def export():
    """Nhận kết quả JSON từ client, tạo Excel và trả về file download."""
    from io import BytesIO
    import openpyxl

    # Nhận data từ form POST (field 'payload') hoặc JSON body
    if request.content_type and 'application/json' in request.content_type:
        data = request.get_json(force=True) or {}
    else:
        raw = request.form.get('payload', '')
        try:
            data = json.loads(raw) if raw else {}
        except Exception:
            data = {}
    results_data = data.get("results", [])
    if not results_data:
        return jsonify({"error": "Không có dữ liệu"}), 400

    # Rebuild CheckResult objects (light version for export)
    from check_format_kltn import CheckResult, Issue
    results = []
    for d in results_data:
        r = CheckResult(filepath=d.get("filename", "unknown"))
        r.student_name = d.get("student_name", "")
        r.student_id   = d.get("student_id", "")
        r.title        = d.get("title", "")
        r.advisor      = d.get("advisor", "")
        r.score        = d.get("score", 0)
        r.error_count  = d.get("error_count", 0)
        r.warn_count   = d.get("warn_count", 0)
        for iss_d in d.get("issues", []):
            iss = Issue(
                category=iss_d.get("category", ""),
                severity=iss_d.get("severity", "INFO"),
                message=iss_d.get("message", ""),
                location=iss_d.get("location", ""),
                suggestion=iss_d.get("suggestion", ""),
            )
            r.issues.append(iss)
        results.append(r)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Dùng temp file độc lập với RESULT_DIR để tránh lỗi đường dẫn
    tmp_path = None
    xlsx_bytes = b""
    try:
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        export_excel(results, tmp_path)
        with open(tmp_path, "rb") as f:
            xlsx_bytes = f.read()
    except Exception as e:
        return jsonify({"error": f"Lỗi tạo Excel: {str(e)}"}), 500
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass

    if not xlsx_bytes:
        return jsonify({"error": "File Excel rỗng, kiểm tra lại dữ liệu"}), 500

    return Response(
        xlsx_bytes,
        status=200,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="KIEM_TRA_KLTN_{timestamp}.xlsx"',
            "Content-Length": str(len(xlsx_bytes)),
        }
    )


# ════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    is_dev = os.environ.get("FLASK_ENV") != "production"
    print(f"\n{'='*60}")
    print(f"  CheckForm KLTN — Web Server")
    print(f"  Mở trình duyệt: http://localhost:{port}")
    print(f"{'='*60}\n")
    app.run(host="0.0.0.0", port=port, debug=is_dev, threaded=True)
