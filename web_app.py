from __future__ import annotations

import shutil
import sys
import threading
import time
import uuid
import webbrowser
import zipfile
import re
import random
from pathlib import Path

import fitz
from flask import Flask, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from pcb import (
    extract_pdf_report,
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


def add_value_highlight(page, rect, color=(1, 0.86, 0.18)) -> None:
    annot = page.add_highlight_annot(rect)
    annot.set_colors(stroke=color)
    annot.update()


def search_variants(term: str) -> list[str]:
    variants = [
        term,
        term.replace(" +/- ", "±").replace("+/-", "±"),
        term.replace(" ", ""),
    ]
    if term.upper() == "UL 94V-0":
        variants.extend(["94V-0", "UL 94 Flame Class 94V-0"])
    if term.upper() == "ROHS DIRECTIVE":
        variants.extend(["RoHS", "RoHS Directive"])
    return list(dict.fromkeys([variant for variant in variants if variant]))


def find_first_rect(document, term: str, page_number: str | None = None):
    page_indexes: list[int] = []
    if page_number and str(page_number).isdigit():
        index = int(page_number) - 1
        if 0 <= index < len(document):
            page_indexes.append(index)
    page_indexes.extend(index for index in range(len(document)) if index not in page_indexes)

    for index in page_indexes:
        page = document[index]
        for variant in search_variants(term):
            rects = page.search_for(variant)
            if rects:
                return page, rects[0]
    return None, None


def find_rects_on_page(page, term: str) -> list[fitz.Rect]:
    for variant in search_variants(term):
        rects = page.search_for(variant)
        if rects:
            return rects
    return []


def clean_row_candidates(row: dict[str, str]) -> list[str]:
    candidates: list[str] = []
    for chunk in str(row.get("TestName", "")).split(" - "):
        candidates.extend(split_highlight_values(chunk))
    candidates.extend(split_highlight_values(row.get("SPEC", "")))
    candidates.extend(split_highlight_values(row.get("RESULTS", "")))

    cleaned: list[str] = []
    for candidate in candidates:
        candidate = " ".join(candidate.split())
        if not candidate or candidate.upper() == "NA":
            continue
        if candidate not in cleaned:
            cleaned.append(candidate)
    return cleaned


def row_text(row: dict[str, str]) -> str:
    return " ".join(str(row.get(key, "")) for key in ("TestName", "SPEC", "RESULTS")).upper()


def is_cover_quantity(row: dict[str, str]) -> bool:
    return row.get("Champ") == "QUANTITY"


def is_ul94_standard(row: dict[str, str]) -> bool:
    return "94V-0" in str(row.get("Norme", "")).upper()


def first_table_required(row: dict[str, str]) -> bool:
    text = row_text(row)
    required_terms = [
        "LAMINATE MATERIAL",
        "CONDUCTOR WIDTH",
        "CONDUCTOR SPACE",
        "ANNULAR RING",
        "COPPER THICKNESS - PTH",
        "COPPER THICKNESS - VIAS FILLING",
        "SOLDERABILITY TEST",
        "ELECTRIC TEST",
        "ADHESION - FINISH",
        "ADHESION - SOLDER RESIST",
        "WARP",
        "SOLDER MASK THICKNESS",
        "GOLD THICKNESS",
        "NICKEL THICKNESS",
        "INOIC CONTAMINATION",
        "IONIC CONTAMINATION",
        "AFTER FINISH",
    ]
    if any(term in text for term in required_terms):
        return True

    impedance_markers = ["L10", "L3", "L2", "L13", "L1", "B2", "B1"]
    return "IMPEDANCE" in text and any(marker in text for marker in impedance_markers)


def is_hole_size_row(row: dict[str, str]) -> bool:
    name = str(row.get("TestName", "")).strip().upper()
    spec = str(row.get("SPEC", "")).upper()
    results = str(row.get("RESULTS", "")).upper()
    if not re.fullmatch(r"[A-Z]", name):
        return False
    return bool(re.search(r"\d", spec) and re.search(r"\d", results))


def is_dimension_row(row: dict[str, str]) -> bool:
    name = str(row.get("TestName", "")).upper()
    spec = str(row.get("SPEC", ""))
    if "ITEM" not in name:
        return False
    return bool(re.search(r"±|\+/-", spec) and re.search(r"\d", spec))


def first_number(value: str) -> float:
    match = re.search(r"(?<!\d)-?\d+(?:[.,]\d+)?", value)
    return float(match.group(0).replace(",", ".")) if match else -1.0


def is_stackup_row(row: dict[str, str]) -> bool:
    text = row_text(row)
    stackup_terms = [
        "SOLDER MASK",
        "LAYER",
        "COPPER",
        "CORE",
        " PP",
        "PP ",
        "VT-",
        "MIL",
    ]
    return "MIL" in text and any(term in text for term in stackup_terms)


def highlight_rows_in_pdf(
    source_pdf: Path,
    output_pdf: Path,
    cover_terms: list[str],
    standard_terms: list[str],
    selected_rows: list[dict[str, str]],
    extra_terms: list[str],
) -> int:
    document = fitz.open(source_pdf)
    highlighted = 0

    for term in cover_terms + standard_terms + extra_terms:
        clean_term = " ".join(str(term).split())
        if not clean_term or clean_term.upper() in {"NA", "OK"}:
            continue
        page, rect = find_first_rect(document, clean_term)
        if page and rect:
            add_value_highlight(page, rect)
            highlighted += 1

    for row in selected_rows:
        candidates = clean_row_candidates(row)
        if not candidates:
            continue

        page_index = int(row.get("Page", "0")) - 1 if str(row.get("Page", "")).isdigit() else -1
        if not 0 <= page_index < len(document):
            continue
        page = document[page_index]

        anchor_rect = None
        for candidate in sorted([c for c in candidates if c.upper() != "OK"], key=len, reverse=True):
            rects = find_rects_on_page(page, candidate)
            if rects:
                anchor_rect = rects[0]
                break

        if anchor_rect is None:
            continue

        line_highlighted = 0
        for candidate in candidates:
            rects = find_rects_on_page(page, candidate)
            if not rects:
                continue
            for rect in rects:
                if abs(rect.y0 - anchor_rect.y0) <= 8:
                    add_value_highlight(page, rect)
                    highlighted += 1
                    line_highlighted += 1

        if line_highlighted == 0:
            add_value_highlight(page, anchor_rect)
            highlighted += 1

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_pdf, garbage=4, deflate=True)
    document.close()
    return highlighted


def build_highlight_data(
    cover_rows: list[dict[str, str]],
    standard_rows: list[dict[str, str]],
    inspection_rows: list[dict[str, str]],
    sample_size: int = 0,
) -> tuple[list[str], list[str], list[dict[str, str]], list[str]]:
    cover_terms: list[str] = []
    standard_terms: list[str] = []

    for row in cover_rows:
        if is_cover_quantity(row):
            cover_terms.extend(split_highlight_values(row.get("Valeur page de garde", "")))

    for row in standard_rows:
        if is_ul94_standard(row):
            standard_terms.extend(["UL 94 Flame Class 94V-0", "94V-0"])

    selected_rows: list[dict[str, str]] = []

    def add_row(row: dict[str, str], require_conformity: bool = True) -> None:
        if require_conformity and row.get("Conformite") != "CONFORME":
            return
        if row in selected_rows:
            return
        selected_rows.append(row)

    for row in inspection_rows:
        if first_table_required(row):
            add_row(row, require_conformity=True)

    hole_rows = [row for row in inspection_rows if is_hole_size_row(row) and row.get("Conformite") == "CONFORME"]
    if hole_rows:
        add_row(random.choice(hole_rows), require_conformity=True)

    dimension_rows = [row for row in inspection_rows if is_dimension_row(row) and row.get("Conformite") == "CONFORME"]
    if dimension_rows:
        add_row(max(dimension_rows, key=lambda row: first_number(str(row.get("SPEC", "")))), require_conformity=True)

    for row in inspection_rows:
        if is_stackup_row(row):
            add_row(row, require_conformity=False)

    extra_terms = ["HOLE WALL COPPER THICKNESS"]

    return cover_terms, standard_terms, selected_rows, extra_terms


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

    warnings: list[str] = []
    highlighted_paths: list[Path] = []
    sampled_total = 0

    for uploaded_file in pdf_files:
        original_name = uploaded_file.filename or "document.pdf"
        filename = secure_filename(original_name)
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        pdf_path = upload_dir / filename
        uploaded_file.save(pdf_path)

        report = extract_pdf_report(pdf_path, display_name=original_name)
        cover_terms, standard_terms, selected_rows, extra_terms = build_highlight_data(
            report.cover_rows,
            report.standard_rows,
            report.inspection_rows,
        )
        highlighted_pdf_path = run_dir / "highlighted" / verified_pdf_name(original_name)
        highlight_rows_in_pdf(pdf_path, highlighted_pdf_path, cover_terms, standard_terms, selected_rows, extra_terms)
        highlighted_paths.append(highlighted_pdf_path)
        sampled_total += len(selected_rows)

        warnings.extend(report.warnings)

    stats = {
        "pdf_count": len(pdf_files),
        "sampled_count": sampled_total,
        "highlighted_count": len(highlighted_paths),
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
        highlighted_count=len(highlighted_paths),
        warnings=warnings,
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
