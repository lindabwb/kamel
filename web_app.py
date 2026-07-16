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


def tokenize_for_highlight(value: str) -> list[str]:
    tokens: list[str] = []
    for part in split_highlight_values(value):
        tokens.append(part)
        if " " in part:
            tokens.extend(piece.strip() for piece in part.split(" ") if piece.strip())
    return list(dict.fromkeys(token for token in tokens if token and token.upper() != "NA"))


def add_value_highlight(page, rect, color=(1, 0.86, 0.18)) -> None:
    annot = page.add_highlight_annot(rect)
    annot.set_colors(stroke=color)
    annot.update()


def page_words(page) -> list[tuple[float, float, float, float, str]]:
    return [
        (float(word[0]), float(word[1]), float(word[2]), float(word[3]), str(word[4]))
        for word in page.get_text("words")
    ]


def highlight_text_line(page, anchor_rect, words: list[tuple[float, float, float, float, str]]) -> int:
    anchor_mid = (anchor_rect.y0 + anchor_rect.y1) / 2
    highlighted = 0
    for x0, y0, x1, y1, text in words:
        if not text.strip():
            continue
        word_mid = (y0 + y1) / 2
        if abs(word_mid - anchor_mid) > 3.5:
            continue
        add_value_highlight(page, fitz.Rect(x0, y0, x1, y1))
        highlighted += 1
    return highlighted


def highlight_first_line_containing(document, terms: list[str]) -> int:
    for page in document:
        words = page_words(page)
        for term in terms:
            for variant in search_variants(term):
                rects = page.search_for(variant)
                if rects:
                    return highlight_text_line(page, rects[0], words)
    return 0


def highlight_first_value(document, term: str) -> int:
    page, rect = find_first_rect(document, term)
    if not page or not rect:
        return 0
    add_value_highlight(page, rect)
    return 1


FIRST_TABLE_TARGETS = [
    (["LAMINATE MATERIAL"], ["LAMINATE", "MATERIAL"]),
    (["CONDUCTOR WIDTH"], ["CONDUCTOR", "WIDTH"]),
    (["CONDUCTOR SPACE"], ["CONDUCTOR", "SPACE"]),
    (["ANNULAR RING"], ["ANNULAR", "RING"]),
    (["COPPER THICKNESS - PTH", "PTH"], ["COPPER", "THICKNESS", "PTH"]),
    (["COPPER THICKNESS - Vias filling", "Vias filling"], ["COPPER", "THICKNESS", "VIAS"]),
    (["SOLDERABILITY TEST"], ["SOLDERABILITY"]),
    (["ELECTRIC TEST - IPCTM650 2.5", "ELECTRIC TEST"], ["ELECTRIC", "TEST"]),
    (["Adhesion - Finish", "Finish"], ["ADHESION", "FINISH"]),
    (["Adhesion - Solder resist", "Solder resist"], ["ADHESION", "SOLDER", "RESIST"]),
    (["WARP＆TWIST", "WARP&TWIST", "WARP"], ["WARP"]),
    (["SOLDER MASK THICKNESS"], ["SOLDER", "MASK", "THICKNESS"]),
    (["GOLD THICKNESS"], ["GOLD", "THICKNESS"]),
    (["NICKEL THICKNESS"], ["NICKEL", "THICKNESS"]),
    (["INOIC CONTAMINATION - 10321310 1B2B1", "IONIC CONTAMINATION - 10321310 1B2B1", "10321310 1B2B1"], ["CONTAMINATION", "1B2B1"]),
    (["INOIC CONTAMINATION - AFTER FINISH", "IONIC CONTAMINATION - AFTER FINISH", "AFTER FINISH"], ["CONTAMINATION", "AFTER", "FINISH"]),
]


IMPEDANCE_MARKERS = ["L10", "L3", "L2", "L13", "L1", "B2", "B1"]


def highlight_first_table_by_keywords(document) -> int:
    highlighted = 0
    used_lines: set[tuple[int, int]] = set()

    def skip_page(page) -> bool:
        text = page.get_text("text").upper()
        return "CONTENTS" in text or "WIRING, PRINTED - COMPONENT" in text

    def line_text(words: list[tuple[float, float, float, float, str]], rect) -> str:
        anchor_mid = (rect.y0 + rect.y1) / 2
        return " ".join(
            text for x0, y0, x1, y1, text in words if abs(((y0 + y1) / 2) - anchor_mid) <= 3.5
        ).upper()

    def line_has_requirements(text: str, required_words: list[str]) -> bool:
        compact = compact_text(text)
        return all(compact_text(word) in compact for word in required_words)

    def add_line_once(page_index: int, rect) -> int:
        line_key = (page_index, round((rect.y0 + rect.y1) / 2))
        if line_key in used_lines:
            return 0
        used_lines.add(line_key)
        page = document[page_index]
        return highlight_text_line(page, rect, page_words(page))

    for keyword_group, required_words in FIRST_TABLE_TARGETS:
        for page_index, page in enumerate(document):
            if skip_page(page):
                continue
            words = page_words(page)
            found = False
            for keyword in keyword_group:
                for variant in search_variants(keyword):
                    rects = page.search_for(variant)
                    rect = next((rect for rect in rects if line_has_requirements(line_text(words, rect), required_words)), None)
                    if rect:
                        highlighted += add_line_once(page_index, rect)
                        found = True
                        break
                if found:
                    break
            if found:
                break

    marker_pattern = re.compile(r"\b(" + "|".join(re.escape(marker) for marker in IMPEDANCE_MARKERS) + r")\b")
    for page_index, page in enumerate(document):
        if skip_page(page):
            continue
        words = page_words(page)
        for marker in IMPEDANCE_MARKERS:
            for rect in page.search_for(marker):
                anchor_mid = (rect.y0 + rect.y1) / 2
                line_text = " ".join(
                    text for x0, y0, x1, y1, text in words if abs(((y0 + y1) / 2) - anchor_mid) <= 3.5
                )
                if "IMPEDANCE" not in line_text.upper() or not marker_pattern.search(line_text):
                    continue
                highlighted += add_line_once(page_index, rect)
                break

    return highlighted


def highlight_stackup_pages(document, page_numbers: set[str]) -> int:
    for index, page in enumerate(document):
        text = page.get_text("text").upper().replace(" ", "")
        if "STACKUP" in text and ("TYPE" in text or "VENDOR" in text or "THICKNESS" in text):
            page_numbers.add(str(index + 1))

    highlighted = 0
    for page_number in page_numbers:
        if not str(page_number).isdigit():
            continue
        page_index = int(page_number) - 1
        if not 0 <= page_index < len(document):
            continue
        page = document[page_index]
        words = page_words(page)
        stack_y = None
        for x0, y0, x1, y1, text in words:
            upper = text.strip().upper()
            if "STACK" in upper or upper == "TYPE":
                stack_y = y0 if stack_y is None else min(stack_y, y0)
        if stack_y is None:
            continue

        for x0, y0, x1, y1, text in words:
            clean = text.strip()
            if not clean:
                continue
            if y0 < stack_y or y0 > page.rect.height - 40:
                continue
            add_value_highlight(page, fitz.Rect(x0, y0, x1, y1))
            highlighted += 1
    return highlighted


def search_variants(term: str) -> list[str]:
    variants = [
        term,
        term.replace(" +/- ", "±").replace("+/-", "±"),
        term.replace(" ", ""),
        term.replace(" / ", "/"),
    ]
    if term.upper() == "UL 94V-0":
        variants.extend(["94V-0", "94V0", "UL 94 Flame Class 94V-0", "UL 94 Flame Class 94V0"])
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


def find_all_rects(document, term: str, page_number: str | None = None) -> list[tuple[int, fitz.Rect]]:
    page_indexes: list[int] = []
    if page_number and str(page_number).isdigit():
        index = int(page_number) - 1
        if 0 <= index < len(document):
            page_indexes.append(index)
    page_indexes.extend(index for index in range(len(document)) if index not in page_indexes)

    matches: list[tuple[int, fitz.Rect]] = []
    for index in page_indexes:
        page = document[index]
        for variant in search_variants(term):
            rects = page.search_for(variant)
            if rects:
                matches.extend((index, rect) for rect in rects)
                break
    return matches


def find_rects_on_page(page, term: str) -> list[fitz.Rect]:
    for variant in search_variants(term):
        rects = page.search_for(variant)
        if rects:
            return rects
    return []


def clean_row_candidates(row: dict[str, str]) -> list[str]:
    candidates: list[str] = []
    for chunk in str(row.get("TestName", "")).split(" - "):
        candidates.extend(tokenize_for_highlight(chunk))
    candidates.extend(tokenize_for_highlight(row.get("SPEC", "")))
    candidates.extend(tokenize_for_highlight(row.get("RESULTS", "")))

    cleaned: list[str] = []
    for candidate in candidates:
        candidate = " ".join(candidate.split())
        if not candidate or candidate.upper() == "NA":
            continue
        if candidate not in cleaned:
            cleaned.append(candidate)
    return cleaned


def test_name_anchor_candidates(value: str) -> list[str]:
    value = " ".join(str(value).split())
    if not value:
        return []

    candidates: list[str] = [value]
    chunks = [chunk.strip() for chunk in value.split(" - ") if chunk.strip()]
    if len(chunks) > 1:
        candidates.extend(chunks[1:])
        candidates.append(chunks[0])

    for candidate in list(candidates):
        if " " in candidate:
            candidates.extend(piece.strip() for piece in candidate.split(" ") if piece.strip())

    weak_terms = {
        "COPPER",
        "THICKNESS",
        "SOLDER",
        "MASK",
        "GOLD",
        "NICKEL",
        "TEST",
        "ADHESION",
        "CONTAMINATION",
        "INOIC",
        "IONIC",
        "CONDUCTOR",
    }
    cleaned: list[str] = []
    for candidate in candidates:
        if candidate.upper() in weak_terms:
            continue
        if candidate not in cleaned:
            cleaned.append(candidate)
    return cleaned


def anchor_row_candidates(row: dict[str, str]) -> list[str]:
    test_name = str(row.get("TestName", ""))
    compact_name = compact_text(test_name)
    prefer_test_name_order = first_table_required(row) and compact_name != "IMPEDANCE"

    if prefer_test_name_order:
        groups = [
            test_name_anchor_candidates(test_name),
            tokenize_for_highlight(row.get("SPEC", "")),
            tokenize_for_highlight(row.get("RESULTS", "")),
        ]
    else:
        groups = [
            tokenize_for_highlight(row.get("SPEC", "")),
            tokenize_for_highlight(row.get("RESULTS", "")),
            tokenize_for_highlight(test_name),
        ]

    cleaned: list[str] = []
    for index, group in enumerate(groups):
        ordered_group = group if prefer_test_name_order and index == 0 else sorted(group, key=len, reverse=True)
        for candidate in ordered_group:
            candidate = " ".join(candidate.split())
            if not candidate or candidate.upper() in {"NA", "OK"}:
                continue
            if len(candidate) <= 2 and not re.fullmatch(r"[A-Z]\d+", candidate.upper()):
                continue
            if candidate not in cleaned:
                cleaned.append(candidate)
    return cleaned


def row_text(row: dict[str, str]) -> str:
    return " ".join(str(row.get(key, "")) for key in ("TestName", "SPEC", "RESULTS")).upper()


def compact_text(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value).upper())


def is_cover_quantity(row: dict[str, str]) -> bool:
    return row.get("Champ") == "QUANTITY"


def is_ul94_standard(row: dict[str, str]) -> bool:
    return "94V-0" in str(row.get("Norme", "")).upper()


def first_table_required(row: dict[str, str]) -> bool:
    text = row_text(row)
    test_name = str(row.get("TestName", "")).upper()
    compact_name = compact_text(test_name)

    exact_tests = {
        "LAMINATEMATERIAL",
        "CONDUCTORWIDTH",
        "CONDUCTORSPACE",
        "ANNULARRING",
        "SOLDERABILITYTEST",
        "SOLDERMASKTHICKNESS",
        "GOLDTHICKNESS",
        "NICKELTHICKNESS",
    }

    if compact_name in exact_tests:
        return True
    if compact_name.startswith("COPPERTHICKNESS") and ("PTH" in compact_name or "VIASFILLING" in compact_name):
        return True
    if compact_name.startswith("ELECTRICTEST"):
        return True
    if compact_name.startswith("ADHESION") and ("FINISH" in compact_name or "SOLDERRESIST" in compact_name):
        return True
    if compact_name.startswith("WARPTWIST"):
        return True
    if (compact_name.startswith("INOICCONTAMINATION") or compact_name.startswith("IONICCONTAMINATION")) and (
        "AFTERFINISH" in compact_name or "1B2B1" in compact_name
    ):
        return True
    if compact_name == "IMPEDANCE" and re.search(r"\b(L10|L3|L2|L13|L1|B2|B1)\b", text):
        return True
    return False


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


def row_contains(row: dict[str, str], term: str) -> bool:
    return term.upper() in row_text(row)


def highlight_rows_in_pdf(
    source_pdf: Path,
    output_pdf: Path,
    cover_terms: list[str],
    standard_terms: list[str],
    selected_rows: list[dict[str, str]],
    extra_terms: list[str],
    page_terms: list[tuple[str, str]],
    stackup_pages: set[str],
) -> int:
    document = fitz.open(source_pdf)
    highlighted = 0
    words_by_page: dict[int, list[tuple[float, float, float, float, str]]] = {}

    highlighted += highlight_first_value(document, "UL 94 Flame Class 94V0")
    if not highlighted:
        highlighted += highlight_first_value(document, "UL 94 Flame Class 94V-0")
    if not highlighted:
        highlighted += highlight_first_value(document, "94V0")
    highlighted += highlight_first_table_by_keywords(document)

    for term in cover_terms:
        clean_term = " ".join(str(term).split())
        if not clean_term or clean_term.upper() in {"NA", "OK"}:
            continue
        page, rect = find_first_rect(document, clean_term)
        if page and rect:
            add_value_highlight(page, rect)
            highlighted += 1

    for term in standard_terms + extra_terms:
        clean_term = " ".join(str(term).split())
        if not clean_term or clean_term.upper() in {"NA", "OK"}:
            continue
        matches = find_all_rects(document, clean_term)
        for page_index, rect in matches[:3]:
            add_value_highlight(document[page_index], rect)
            highlighted += 1

    for term, page_number in page_terms:
        page, rect = find_first_rect(document, term, page_number)
        if page and rect:
            add_value_highlight(page, rect)
            highlighted += 1

    for row in selected_rows:
        if is_stackup_row(row):
            continue
        candidates = clean_row_candidates(row)
        if is_stackup_row(row):
            candidates = [candidate for candidate in candidates if len(candidate) > 1]
        if not candidates:
            continue

        page_index = int(row.get("Page", "0")) - 1 if str(row.get("Page", "")).isdigit() else -1
        if not 0 <= page_index < len(document):
            continue
        page = document[page_index]
        words = words_by_page.setdefault(page_index, page_words(page))

        anchor_rect = None
        for candidate in anchor_row_candidates(row):
            rects = find_rects_on_page(page, candidate)
            if rects:
                anchor_rect = rects[0]
                break

        if anchor_rect is None:
            continue

        line_highlighted = highlight_text_line(page, anchor_rect, words)
        highlighted += line_highlighted
        if line_highlighted == 0:
            add_value_highlight(page, anchor_rect)
            highlighted += 1

    highlighted += highlight_stackup_pages(document, stackup_pages)

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_pdf, garbage=1, deflate=False)
    document.close()
    return highlighted


def build_highlight_data(
    cover_rows: list[dict[str, str]],
    standard_rows: list[dict[str, str]],
    inspection_rows: list[dict[str, str]],
    sample_size: int = 0,
) -> tuple[list[str], list[str], list[dict[str, str]], list[str], list[tuple[str, str]], set[str]]:
    cover_terms: list[str] = []
    standard_terms: list[str] = []

    for row in cover_rows:
        cover_terms.extend(tokenize_for_highlight(row.get("Valeur page de garde", "")))
        if is_cover_quantity(row):
            cover_terms.append("PCS")

    for row in standard_rows:
        if is_ul94_standard(row):
            continue

    selected_rows: list[dict[str, str]] = []
    page_terms: list[tuple[str, str]] = []
    stackup_pages: set[str] = set()

    def add_row(row: dict[str, str], require_conformity: bool = True) -> None:
        if require_conformity and row.get("Conformite") != "CONFORME":
            return
        if row in selected_rows:
            return
        selected_rows.append(row)

    for row in inspection_rows:
        if first_table_required(row):
            continue
    hole_rows = [row for row in inspection_rows if is_hole_size_row(row) and row.get("Conformite") == "CONFORME"]
    if hole_rows:
        add_row(random.choice(hole_rows), require_conformity=True)

    dimension_rows = [row for row in inspection_rows if is_dimension_row(row) and row.get("Conformite") == "CONFORME"]
    if dimension_rows:
        add_row(max(dimension_rows, key=lambda row: first_number(str(row.get("SPEC", "")))), require_conformity=True)

    for row in inspection_rows:
        if is_stackup_row(row):
            add_row(row, require_conformity=False)
            stackup_pages.add(str(row.get("Page", "")))

    extra_terms = ["HOLE WALL COPPER THICKNESS"]

    return cover_terms, standard_terms, selected_rows, extra_terms, page_terms, stackup_pages


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

        try:
            report = extract_pdf_report(pdf_path, display_name=original_name)
            cover_terms, standard_terms, selected_rows, extra_terms, page_terms, stackup_pages = build_highlight_data(
                report.cover_rows,
                report.standard_rows,
                report.inspection_rows,
            )
            highlighted_pdf_path = run_dir / "highlighted" / verified_pdf_name(original_name)
            highlight_rows_in_pdf(
                pdf_path,
                highlighted_pdf_path,
                cover_terms,
                standard_terms,
                selected_rows,
                extra_terms,
                page_terms,
                stackup_pages,
            )
            highlighted_paths.append(highlighted_pdf_path)
            sampled_total += len(selected_rows)
            warnings.extend(report.warnings)
        except Exception as exc:
            shutil.rmtree(run_dir, ignore_errors=True)
            return render_template(
                "index.html",
                error=f"Erreur pendant le traitement de {original_name}: {exc}",
            )

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
