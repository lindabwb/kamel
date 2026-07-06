from __future__ import annotations

import shutil
import sys
import threading
import time
import uuid
import webbrowser
from pathlib import Path

from flask import Flask, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from pcb import (
    COVER_COLUMNS,
    INSPECTION_COLUMNS,
    STANDARD_COLUMNS,
    extract_pdf_report,
    write_excel_report,
)


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    APP_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
    APP_DIR = BASE_DIR
RUNS_DIR = APP_DIR / "web_reports"

app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static"),
)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024


RUN_DOWNLOAD_NAMES: dict[str, str] = {}


def is_pdf(filename: str) -> bool:
    return filename.lower().endswith(".pdf")


def report_download_name(uploaded_names: list[str]) -> str:
    if len(uploaded_names) == 1:
        stem = Path(uploaded_names[0]).stem
        safe_stem = secure_filename(stem) or "document"
        return f"rapport_{safe_stem}.xlsx"
    return f"rapport_{len(uploaded_names)}_fichiers_pcb.xlsx"


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    files = [file for file in request.files.getlist("pdfs") if file and file.filename]
    pdf_files = [file for file in files if is_pdf(file.filename)]

    if not pdf_files:
        return render_template(
            "index.html",
            error="Ajoutez au moins un fichier PDF.",
        )

    run_id = uuid.uuid4().hex
    run_dir = RUNS_DIR / run_id
    upload_dir = run_dir / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)

    cover_rows: list[dict[str, str]] = []
    standard_rows: list[dict[str, str]] = []
    inspection_rows: list[dict[str, str]] = []
    warnings: list[str] = []

    for uploaded_file in pdf_files:
        original_name = uploaded_file.filename or "document.pdf"
        filename = secure_filename(original_name)
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        pdf_path = upload_dir / filename
        uploaded_file.save(pdf_path)

        report = extract_pdf_report(pdf_path, display_name=original_name)

        cover_rows.extend(report.cover_rows)
        standard_rows.extend(report.standard_rows)
        inspection_rows.extend(report.inspection_rows)
        warnings.extend(report.warnings)

    excel_path = run_dir / "rapport_pcb.xlsx"
    write_excel_report(cover_rows, standard_rows, inspection_rows, excel_path)

    total_checks = len(inspection_rows) + len(cover_rows) + len(standard_rows)
    non_conforme_count = (
        sum(1 for row in inspection_rows if row.get("Conformite") == "NON CONFORME")
        + sum(1 for row in cover_rows if row.get("Comparaison") == "DIFFERENT")
    )
    verify_count = (
        sum(1 for row in inspection_rows if row.get("Conformite") == "A VERIFIER")
        + sum(1 for row in cover_rows if row.get("Comparaison") == "A VERIFIER")
        + sum(1 for row in standard_rows if row.get("Norme") == "NA")
    )
    stats = {
        "pdf_count": len(pdf_files),
        "cover_count": len(cover_rows),
        "standard_count": len(standard_rows),
        "inspection_count": total_checks,
        "conforme_count": max(total_checks - non_conforme_count - verify_count, 0),
        "non_conforme_count": non_conforme_count,
        "verify_count": verify_count,
    }
    uploaded_names = [file.filename or "document.pdf" for file in pdf_files]
    RUN_DOWNLOAD_NAMES[run_id] = report_download_name(uploaded_names)

    return render_template(
        "index.html",
        processed=True,
        run_id=run_id,
        stats=stats,
        uploaded_names=uploaded_names,
        warnings=warnings,
        cover_columns=[column for column in COVER_COLUMNS if column != "FileName"],
        standard_columns=[column for column in STANDARD_COLUMNS if column != "FileName"],
        inspection_columns=[column for column in INSPECTION_COLUMNS if column != "FileName"],
        cover_rows=cover_rows,
        standard_rows=standard_rows,
        inspection_rows=inspection_rows,
    )


@app.route("/download/<run_id>", methods=["GET"])
def download(run_id: str):
    excel_path = RUNS_DIR / run_id / "rapport_pcb.xlsx"
    if not excel_path.exists():
        return "Rapport introuvable.", 404
    return send_file(
        excel_path,
        as_attachment=True,
        download_name=RUN_DOWNLOAD_NAMES.get(run_id, "rapport_pcb.xlsx"),
    )


@app.route("/clear/<run_id>", methods=["POST"])
def clear(run_id: str):
    run_dir = RUNS_DIR / run_id
    if run_dir.exists():
        shutil.rmtree(run_dir, ignore_errors=True)
    return render_template("index.html", message="Session supprimée.")


if __name__ == "__main__":
    def open_browser() -> None:
        time.sleep(1.5)
        webbrowser.open("http://127.0.0.1:5000")

    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
