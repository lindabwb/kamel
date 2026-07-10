from __future__ import annotations

import shutil
import sys
import threading
import time
import uuid
import webbrowser
import random
import zipfile
from pathlib import Path

import fitz
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
RUN_HIGHLIGHT_PATHS: dict[str, list[Path]] = {}


def is_pdf(filename: str) -> bool:
    return filename.lower().endswith(".pdf")


def report_download_name(uploaded_names: list[str]) -> str:
    if len(uploaded_names) == 1:
        stem = Path(uploaded_names[0]).stem
        safe_stem = secure_filename(stem.replace("—", "_").replace("–", "_")) or "document"
        return f"rapport_{safe_stem}.xlsx"
    return f"rapport_{len(uploaded_names)}_fichiers_pcb.xlsx"


def verified_pdf_name(original_name: str) -> str:
    path = Path(original_name)
    return f"{path.stem}V.pdf"


def split_highlight_values(value: str) -> list[str]:
    value = (value or "").strip()
    if not value or value.upper() == "NA":
        return []
    parts = [part.strip() for part in value.split("|")]
    cleaned: list[str] = []
    for part in parts:
        if not part or part.upper() == "NA":
            continue
        cleaned.append(part)
    return cleaned


def highlight_terms_in_pdf(source_pdf: Path, output_pdf: Path, terms: list[str]) -> int:
    document = fitz.open(source_pdf)
    highlighted = 0
    seen_terms: set[str] = set()

    for raw_term in terms:
        term = " ".join(str(raw_term).split())
        if not term or term.upper() in {"NA", "OK"} or term in seen_terms:
            continue
        seen_terms.add(term)

        variants = [
            term,
            term.replace(" +/- ", "±").replace("+/-", "±"),
            term.replace(" ", ""),
        ]
        if term.upper() == "UL 94V-0":
            variants.extend(["94V-0", "UL 94 Flame Class 94V-0"])
        if term.upper() == "ROHS DIRECTIVE":
            variants.extend(["RoHS", "RoHS Directive"])

        found_for_term = False
        for page in document:
            for variant in variants:
                if not variant:
                    continue
                rects = page.search_for(variant)
                if not rects:
                    continue
                for rect in rects:
                    annot = page.add_highlight_annot(rect)
                    annot.set_colors(stroke=(1, 0.86, 0.18))
                    annot.update()
                    highlighted += 1
                found_for_term = True
                break
            if found_for_term:
                break

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_pdf, garbage=4, deflate=True)
    document.close()
    return highlighted


def build_highlight_terms(
    cover_rows: list[dict[str, str]],
    standard_rows: list[dict[str, str]],
    inspection_rows: list[dict[str, str]],
    sample_size: int,
) -> list[str]:
    terms: list[str] = []

    for row in cover_rows:
        if row.get("Champ") == "QUANTITY":
            continue
        terms.extend(split_highlight_values(row.get("Valeur page de garde", "")))

    for row in standard_rows:
        terms.extend(split_highlight_values(row.get("Norme", "")))

    eligible_rows = [
        row for row in inspection_rows
        if row.get("SPEC", "NA") != "NA" or row.get("RESULTS", "NA") != "NA"
    ]
    sample_count = min(max(sample_size, 0), len(eligible_rows))
    sampled_rows = random.sample(eligible_rows, sample_count) if sample_count else []

    for row in sampled_rows:
        terms.extend(split_highlight_values(row.get("SPEC", "")))
        terms.extend(split_highlight_values(row.get("RESULTS", "")))

    return terms


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    files = [file for file in request.files.getlist("pdfs") if file and file.filename]
    pdf_files = [file for file in files if is_pdf(file.filename)]
    try:
        sample_size = int(request.form.get("sample_size", "10"))
    except ValueError:
        sample_size = 10
    sample_size = max(0, min(sample_size, 200))

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
    highlighted_paths: list[Path] = []

    for uploaded_file in pdf_files:
        original_name = uploaded_file.filename or "document.pdf"
        filename = secure_filename(original_name)
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        pdf_path = upload_dir / filename
        uploaded_file.save(pdf_path)

        report = extract_pdf_report(pdf_path, display_name=original_name)
        terms = build_highlight_terms(
            report.cover_rows,
            report.standard_rows,
            report.inspection_rows,
            sample_size,
        )
        highlighted_pdf_path = run_dir / "highlighted" / verified_pdf_name(original_name)
        highlight_terms_in_pdf(pdf_path, highlighted_pdf_path, terms)
        highlighted_paths.append(highlighted_pdf_path)

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
    RUN_HIGHLIGHT_PATHS[run_id] = highlighted_paths

    return render_template(
        "index.html",
        processed=True,
        run_id=run_id,
        stats=stats,
        uploaded_names=uploaded_names,
        sample_size=sample_size,
        highlighted_count=len(highlighted_paths),
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


@app.route("/download-highlighted/<run_id>", methods=["GET"])
def download_highlighted(run_id: str):
    highlighted_paths = [path for path in RUN_HIGHLIGHT_PATHS.get(run_id, []) if path.exists()]
    if not highlighted_paths:
        return "PDF surligné introuvable.", 404

    if len(highlighted_paths) == 1:
        return send_file(
            highlighted_paths[0],
            as_attachment=True,
            download_name=highlighted_paths[0].name,
        )

    zip_path = RUNS_DIR / run_id / "pdfs_surlignes.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in highlighted_paths:
            archive.write(path, arcname=path.name)

    return send_file(zip_path, as_attachment=True, download_name="pdfs_surlignes.zip")


@app.route("/clear/<run_id>", methods=["POST"])
def clear(run_id: str):
    run_dir = RUNS_DIR / run_id
    if run_dir.exists():
        shutil.rmtree(run_dir, ignore_errors=True)
    RUN_DOWNLOAD_NAMES.pop(run_id, None)
    RUN_HIGHLIGHT_PATHS.pop(run_id, None)
    return render_template("index.html", message="Session supprimée.")


if __name__ == "__main__":
    def open_browser() -> None:
        time.sleep(1.5)
        webbrowser.open("http://127.0.0.1:5000")

    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
