"""
PDF2Word Pro — Flask Web App
Run locally:  python app.py  →  http://localhost:5000
Deployed:     auto-served via gunicorn on Render / Railway / any PaaS
"""

import os
import uuid
import threading
import time
from pathlib import Path
from flask import (
    Flask, request, jsonify, send_file,
    render_template, abort
)
from werkzeug.utils import secure_filename
from src.converter import PDF2WordConverter

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB upload limit

# True when deployed to a remote server (set via env var in render.yaml)
IS_REMOTE = os.environ.get("IS_REMOTE", "false").lower() == "true"

# Temp working directories (auto-cleaned after download)
UPLOAD_DIR = Path("_uploads")
OUTPUT_DIR = Path("_output")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# In-memory job store  {job_id: {"status", "progress", "message", "output_path", "filename"}}
jobs: dict[str, dict] = {}
jobs_lock = threading.Lock()

ALLOWED_EXT = {".pdf"}


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXT


# ── Routes ──────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html", is_remote=IS_REMOTE)


@app.route("/convert", methods=["POST"])
def convert():
    """
    Accepts multipart/form-data:
      - file        : PDF file
      - output_name : desired output filename (optional)
      - ocr_lang    : tesseract lang code (default: eng)
    Returns JSON: { job_id }
    """
    if "file" not in request.files:
        return jsonify(error="No file uploaded"), 400

    f = request.files["file"]
    if not f.filename or not allowed_file(f.filename):
        return jsonify(error="Please upload a PDF file"), 400

    ocr_lang = request.form.get("ocr_lang", "eng").strip() or "eng"
    safe_name = secure_filename(f.filename)
    job_id = uuid.uuid4().hex

    # Save uploaded PDF
    upload_path = UPLOAD_DIR / f"{job_id}_{safe_name}"
    f.save(str(upload_path))

    # Output filename
    stem = Path(safe_name).stem
    output_filename = f"{stem}.docx"
    output_path = OUTPUT_DIR / f"{job_id}_{output_filename}"

    # Register job
    with jobs_lock:
        jobs[job_id] = {
            "status": "queued",
            "progress": 0,
            "message": "Queued...",
            "output_path": str(output_path),
            "output_filename": output_filename,
            "upload_path": str(upload_path),
        }

    # Run conversion in background thread
    thread = threading.Thread(
        target=_run_job,
        args=(job_id, str(upload_path), str(output_path), ocr_lang),
        daemon=True,
    )
    thread.start()

    return jsonify(job_id=job_id)


@app.route("/status/<job_id>")
def status(job_id: str):
    """Poll job status. Returns JSON with status, progress (0-100), message."""
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify(error="Job not found"), 404
    return jsonify(
        status=job["status"],
        progress=job["progress"],
        message=job["message"],
    )


@app.route("/download/<job_id>")
def download(job_id: str):
    """Stream the converted .docx to the browser for download."""
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done":
        abort(404)

    output_path = job["output_path"]
    output_filename = job["output_filename"]

    if not os.path.isfile(output_path):
        abort(404)

    response = send_file(
        output_path,
        as_attachment=True,
        download_name=output_filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    # Schedule cleanup after response is sent
    @response.call_on_close
    def cleanup():
        _cleanup_job(job_id)

    return response


@app.route("/save-to-folder", methods=["POST"])
def save_to_folder():
    """
    Save converted file to a local folder path on this machine.
    Body JSON: { job_id, folder_path }
    Only works when server and browser are on the same machine (local run).
    Returns 403 when IS_REMOTE is True.
    """
    if IS_REMOTE:
        return jsonify(error="Folder save is only available when running locally."), 403

    data = request.get_json(force=True)
    job_id = data.get("job_id", "")
    folder_path = data.get("folder_path", "").strip()

    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify(error="Job not ready"), 404

    if not folder_path:
        return jsonify(error="No folder path provided"), 400

    folder = Path(folder_path)
    try:
        folder.mkdir(parents=True, exist_ok=True)
        dest = folder / job["output_filename"]
        import shutil
        shutil.copy2(job["output_path"], str(dest))
        _cleanup_job(job_id)
        return jsonify(saved_to=str(dest))
    except Exception as exc:
        return jsonify(error=str(exc)), 500


# ── Background worker ────────────────────────────────────────────────────

def _run_job(job_id: str, upload_path: str, output_path: str, ocr_lang: str):
    def _update(progress: int, message: str, status: str = "running"):
        with jobs_lock:
            if job_id in jobs:
                jobs[job_id]["progress"] = progress
                jobs[job_id]["message"] = message
                jobs[job_id]["status"] = status

    try:
        _update(5, "Analysing PDF...")
        converter = PDF2WordConverter(ocr_lang=ocr_lang)

        _update(20, "Extracting text, tables and images...")
        result = converter.convert(upload_path, output_path)

        if result.success:
            _update(100, f"Done — {result.page_count} page(s) converted", status="done")
        else:
            _update(0, f"Conversion failed: {result.error}", status="error")

    except Exception as exc:
        _update(0, f"Error: {exc}", status="error")
    finally:
        # Clean up the uploaded PDF
        try:
            os.remove(upload_path)
        except Exception:
            pass


# ── Cleanup ──────────────────────────────────────────────────────────────

def _cleanup_job(job_id: str):
    with jobs_lock:
        job = jobs.pop(job_id, None)
    if job:
        for key in ("output_path", "upload_path"):
            try:
                os.remove(job[key])
            except Exception:
                pass


def _periodic_cleanup(max_age_seconds: int = 3600):
    """Remove output files older than max_age_seconds (runs every 10 min)."""
    while True:
        time.sleep(600)
        now = time.time()
        for path in list(OUTPUT_DIR.glob("*.docx")) + list(UPLOAD_DIR.glob("*.pdf")):
            try:
                if now - path.stat().st_mtime > max_age_seconds:
                    path.unlink()
            except Exception:
                pass


# Start background cleaner
threading.Thread(target=_periodic_cleanup, daemon=True).start()


# ── Entry point ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("\n PDF2Word Pro")
    print(f" Open in browser → http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
