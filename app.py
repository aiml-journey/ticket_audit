import os
from flask import Flask, request, send_file, render_template
from pathlib import Path
from audit_tickets import audit_tickets

app = Flask(__name__)

UPLOAD_FOLDER = Path("uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/audit", methods=["POST"])
def audit_api():
    if "mis_file" not in request.files:
        return {"error": "No MIS file uploaded"}, 400
    
    mis_file = request.files["mis_file"]
    pdf_files = request.files.getlist("pdf_files")

    session_dir = UPLOAD_FOLDER / "current_session"
    tickets_dir = session_dir / "testticket"
    tickets_dir.mkdir(parents=True, exist_ok=True)

    # Save MIS file
    mis_path = session_dir / mis_file.filename
    mis_file.save(str(mis_path))

    # Save PDFs
    for pdf in pdf_files:
        pdf_path = tickets_dir / pdf.filename
        pdf.save(str(pdf_path))

    # Output path
    output_path = session_dir / f"Audited_{mis_file.filename}"

    # Run your audit engine
    result = audit_tickets(
        excel_path=str(mis_path),
        tickets_dir=str(tickets_dir),
        output_path=str(output_path)
    )

    return send_file(result["saved_file"], as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)