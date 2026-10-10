import os
import uuid
from pathlib import Path

from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename

from audit_tickets import audit_tickets


app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

EXCEL_MIME = (
    "application/vnd.openxmlformats-"
    "officedocument.spreadsheetml.sheet"
)


@app.route("/")
def index():
    return send_file(BASE_DIR / "index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/api/audit", methods=["POST"])
def audit_api():
    if "mis_file" not in request.files:
        return jsonify({"error": "No MIS Excel file uploaded"}), 400

    mis_file = request.files["mis_file"]
    if not mis_file.filename:
        return jsonify({"error": "MIS file has no filename"}), 400

    mis_filename = secure_filename(mis_file.filename)
    if not mis_filename.lower().endswith(".xlsx"):
        return jsonify({"error": "MIS file must be an .xlsx file"}), 400

    pdf_files = request.files.getlist("pdf_files")
    if not pdf_files:
        return jsonify({"error": "No PDF tickets uploaded"}), 400

    session_id = uuid.uuid4().hex
    session_dir = UPLOAD_FOLDER / session_id
    tickets_dir = session_dir / "tickets"
    tickets_dir.mkdir(parents=True, exist_ok=True)

    try:
        mis_path = session_dir / mis_filename
        mis_file.save(str(mis_path))

        pdf_count = 0
        for pdf in pdf_files:
            if not pdf.filename:
                continue
            pdf_filename = secure_filename(pdf.filename)
            if not pdf_filename.lower().endswith(".pdf"):
                continue
            pdf.save(str(tickets_dir / pdf_filename))
            pdf_count += 1

        if pdf_count == 0:
            return jsonify({"error": "No valid PDF files uploaded"}), 400

        output_path = session_dir / f"Audited_{mis_filename}"

        audit_tickets(
            excel_path=str(mis_path),
            tickets_dir=str(tickets_dir),
            output_path=str(output_path),
            folder_label="tickets",
        )

        return send_file(
            str(output_path),
            as_attachment=True,
            download_name=output_path.name,
            mimetype=EXCEL_MIME,
        )

    except Exception as e:
        return jsonify({"error": "Audit failed", "details": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
