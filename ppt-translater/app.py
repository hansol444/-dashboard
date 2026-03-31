# -*- coding: utf-8 -*-
"""
app.py — PPT 장표 번역기 Web UI (Flask)
실행: python app.py → http://localhost:5000
"""
import os, sys

# PYTHONUTF8=1 없이 실행됐으면 subprocess로 재시작 (Windows cp949 완전 차단)
if os.environ.get("PYTHONUTF8") != "1":
    import subprocess
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    proc = subprocess.Popen([sys.executable] + sys.argv, env=env)
    sys.exit(proc.wait())

import json
import queue
import shutil
import threading
import time
import traceback
import uuid
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request, send_file

import config
from translate import translate_pptx
from integrations import upload_to_sharepoint, send_slack_notification

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

TEMP_DIR = Path("temp_jobs")
TEMP_DIR.mkdir(exist_ok=True)

# ─── Job 상태 관리 ─────────────────────────────────────────────────────────────

jobs: dict[str, "Job"] = {}


JOB_TTL_SECONDS = 3600  # 1시간 후 자동 정리


class Job:
    def __init__(self):
        self.q: queue.Queue = queue.Queue()
        self.pptx_path: str | None = None
        self.report_path: str | None = None
        self.created_at: float = time.time()
        self.finished: bool = False


# ─── 라우트 ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/translate", methods=["POST"])
def start_translate():
    file = request.files.get("file")
    if not file or not file.filename.endswith(".pptx"):
        return jsonify({"error": ".pptx 파일만 업로드 가능합니다."}), 400

    direction = request.form.get("direction", "ko_to_en")
    rules_json = request.form.get("rules", "{}")

    try:
        rules_map: dict = json.loads(rules_json)
    except json.JSONDecodeError:
        rules_map = {}

    enabled_rules = [rule for rule, enabled in rules_map.items() if enabled]
    if not enabled_rules:
        enabled_rules = config.get_enabled_rules()

    job_id = str(uuid.uuid4())
    job_dir = TEMP_DIR / job_id
    job_dir.mkdir()

    stem = Path(file.filename).stem
    input_path = str(job_dir / file.filename)
    suffix = "_EN" if direction == "ko_to_en" else "_KO"
    output_path = str(job_dir / f"{stem}{suffix}.pptx")

    file.save(input_path)

    job = Job()
    jobs[job_id] = job

    # 초과 항목 카운터 공유
    overflow_info = {"count": 0}

    def run():
        try:
            def callback(event_type: str, data: dict):
                if event_type == "overflow":
                    overflow_info["count"] += 1
                job.q.put({"type": event_type, **data})

            translate_pptx(
                input_path=input_path,
                output_path=output_path,
                direction=direction,
                quality="fast",
                postprocess=True,
                make_report=True,
                enabled_rules=enabled_rules,
                progress_callback=callback,
            )

            job.pptx_path = output_path
            report_path = output_path.replace(".pptx", "_번역리포트.txt")
            if os.path.exists(report_path):
                job.report_path = report_path

            filename = Path(output_path).name

            # ── SharePoint 업로드 ──────────────────────────────────────────
            job.q.put({"type": "integration", "step": "sharepoint", "status": "running"})
            sp_ok, sp_msg = upload_to_sharepoint(output_path)
            job.q.put({
                "type": "integration", "step": "sharepoint",
                "status": "done" if sp_ok else "error",
                "message": sp_msg,
            })

            # ── Slack 알림 ────────────────────────────────────────────────
            job.q.put({"type": "integration", "step": "slack", "status": "running"})
            sl_ok, sl_msg = send_slack_notification(filename, overflow_info["count"])
            job.q.put({
                "type": "integration", "step": "slack",
                "status": "done" if sl_ok else "error",
                "message": sl_msg,
            })

            job.q.put({"type": "done", "has_report": job.report_path is not None})
            job.finished = True

        except Exception as e:
            job.q.put({"type": "error", "message": str(e), "detail": traceback.format_exc()})
            job.finished = True

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/stream/<job_id>")
def stream(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404

    def generate():
        while True:
            try:
                event = job.q.get(timeout=60)
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
                if event["type"] in ("done", "error"):
                    break
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'heartbeat'})}\n\n"

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/download/pptx/<job_id>")
def download_pptx(job_id: str):
    job = jobs.get(job_id)
    if not job or not job.pptx_path or not os.path.exists(job.pptx_path):
        return "파일을 찾을 수 없습니다.", 404
    return send_file(job.pptx_path, as_attachment=True)


@app.route("/download/report/<job_id>")
def download_report(job_id: str):
    job = jobs.get(job_id)
    if not job or not job.report_path or not os.path.exists(job.report_path):
        return "리포트를 찾을 수 없습니다.", 404
    return send_file(job.report_path, as_attachment=True)


def cleanup_expired_jobs():
    """만료된 Job과 temp 파일 정리 (백그라운드 루프)."""
    while True:
        time.sleep(600)  # 10분마다
        now = time.time()
        expired = [jid for jid, j in jobs.items()
                   if j.finished and now - j.created_at > JOB_TTL_SECONDS]
        for jid in expired:
            job_dir = TEMP_DIR / jid
            if job_dir.exists():
                shutil.rmtree(job_dir, ignore_errors=True)
            jobs.pop(jid, None)


if __name__ == "__main__":
    threading.Thread(target=cleanup_expired_jobs, daemon=True).start()
    print("PPT 번역기 시작: http://localhost:5000")
    app.run(debug=False, port=5000, threaded=True)
