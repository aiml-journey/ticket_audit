import os
import uuid
import time
import threading
import traceback
from pathlib import Path

from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

from audit_tickets import audit_tickets


app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

# Up to 1 GB per request so a full day's batch (100+ scanned PDFs) fits
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 1024

EXCEL_MIME = (
    "application/vnd.openxmlformats-"
    "officedocument.spreadsheetml.sheet"
)

# Job state lives in memory: {job_id: {...}}
JOBS = {}
JOBS_LOCK = threading.Lock()


def _set_job(job_id, **changes):
    with JOBS_LOCK:
        JOBS[job_id].update(changes)


def run_audit_job(job_id, mis_path, tickets_dir, output_path):
    """Runs in a background thread so the web request returns at once."""
    try:
        audit_tickets(
            excel_path=str(mis_path),
            tickets_dir=str(tickets_dir),
            output_path=str(output_path),
            folder_label="tickets",
        )
        _set_job(job_id, state="done", finished=time.time())
    except Exception as e:
        traceback.print_exc()
        reason = f"{type(e).__name__}: {e}" if str(e) else type(e).__name__
        _set_job(job_id, state="error", error=reason, finished=time.time())


@app.errorhandler(RequestEntityTooLarge)
def too_large(_e):
    return jsonify({"error": "Upload too large", "details": "Total upload is over 1 GB"}), 413


@app.route("/")
def index():
    return send_file(BASE_DIR / "index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/audit", methods=["POST"])
def start_audit():
    if "mis_file" not in request.files:
        return jsonify({"error": "No MIS Excel file uploaded"}), 400

    mis_file = request.files["mis_file"]
    mis_filename = secure_filename(mis_file.filename or "")
    if not mis_filename.lower().endswith(".xlsx"):
        return jsonify({"error": "MIS file must be an .xlsx file"}), 400

    pdf_files = [p for p in request.files.getlist("pdf_files") if p.filename]
    pdf_files = [p for p in pdf_files if secure_filename(p.filename).lower().endswith(".pdf")]
    if not pdf_files:
        return jsonify({"error": "No PDF tickets uploaded"}), 400

    job_id = uuid.uuid4().hex
    session_dir = UPLOAD_FOLDER / job_id
    tickets_dir = session_dir / "tickets"
    tickets_dir.mkdir(parents=True, exist_ok=True)

    mis_path = session_dir / mis_filename
    mis_file.save(str(mis_path))

    # Keep every PDF, even when two have the same name.
    # The extension is always saved in lower case (.pdf) so every file is found on Linux (Render),
    # where file search is case-sensitive: ".PDF" would otherwise be missed.
    for i, pdf in enumerate(pdf_files):
        stem = os.path.splitext(secure_filename(pdf.filename))[0] or "ticket"
        pdf.save(str(tickets_dir / f"{i:04d}_{stem}.pdf"))

    output_path = session_dir / f"Audited_{mis_filename}"

    with JOBS_LOCK:
        JOBS[job_id] = {
            "state": "running",
            "pdf_count": len(pdf_files),
            "started": time.time(),
            "output": str(output_path),
            "download_name": output_path.name,
            "error": "",
        }

    threading.Thread(
        target=run_audit_job,
        args=(job_id, mis_path, tickets_dir, output_path),
        daemon=True,
    ).start()

    return jsonify({"job_id": job_id, "pdf_count": len(pdf_files)}), 202


@app.route("/api/status/<job_id>")
def job_status(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if job is None:
        return jsonify({"state": "unknown", "error": "Job not found. Upload again."}), 404
    elapsed = int((job.get("finished") or time.time()) - job["started"])
    return jsonify({
        "state": job["state"],
        "error": job.get("error", ""),
        "pdf_count": job["pdf_count"],
        "elapsed_seconds": elapsed,
    })


@app.route("/api/download/<job_id>")
def download(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if job is None or job["state"] != "done":
        return jsonify({"error": "Result not ready"}), 409
    return send_file(
        job["output"],
        as_attachment=True,
        download_name=job["download_name"],
        mimetype=EXCEL_MIME,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, threaded=True)
