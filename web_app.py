from __future__ import annotations

import shutil
import sys
import threading
import time
import uuid
import webbrowser
import zipfile
import re
from pathlib import Path

import fitz
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

from pcb import parse_filename_values_from_name


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS)
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


def cover_terms_from_filename(filename: str) -> list[str]:
    """Extrait les termes à surligner depuis le nom du fichier."""
    values = parse_filename_values_from_name(filename)
    terms: list[str] = []
    for field in ("P.O.NO", "PART NO", "DATA CODE"):
        value = values.get(field, "NA")
        if value and value != "NA":
            terms.append(value)

    quantity = values.get("QUANTITY", "NA")
    if quantity and quantity != "NA":
        terms.append(quantity)
        quantity_number = re.search(r"\d+", quantity)
        if quantity_number:
            terms.append(quantity_number.group(0))
        terms.append("PCS")

    return list(dict.fromkeys(terms))


def add_highlight_rect(page, rect, color=(1, 0.86, 0.18)) -> None:
    """Ajoute un surlignage sur un rectangle."""
    annot = page.add_highlight_annot(rect)
    annot.set_colors(stroke=color)
    annot.update()


def get_page_blocks(page) -> list[dict]:
    """Récupère les blocs de texte d'une page avec leur structure."""
    return page.get_text("dict")["blocks"]


def get_text_lines_from_page(page) -> list[tuple[float, str, list[tuple[float, float, float, float, str]]]]:
    """
    Extrait les lignes de texte d'une page avec leurs coordonnées.
    Retourne: [(y_position, texte_complet, [(x0, y0, x1, y1, mot), ...])]
    """
    blocks = get_page_blocks(page)
    lines: dict[float, list[tuple[float, float, float, float, str]]] = {}
    
    for block in blocks:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            if "spans" not in line:
                continue
            for span in line["spans"]:
                text = span["text"].strip()
                if not text:
                    continue
                x0, y0, x1, y1 = span["bbox"]
                mid_y = (y0 + y1) / 2
                # Grouper par ligne approximative
                key = round(mid_y / 2) * 2
                if key not in lines:
                    lines[key] = []
                lines[key].append((x0, y0, x1, y1, text))
    
    # Trier par position Y et construire les lignes
    result = []
    for y_key in sorted(lines.keys()):
        words = sorted(lines[y_key], key=lambda w: w[0])  # Trier par X
        text = " ".join(w[4] for w in words)
        avg_y = sum((w[1] + w[3]) / 2 for w in words) / len(words)
        result.append((avg_y, text, words))
    
    return result


def find_item_lines(lines: list[tuple[float, str, list]], target_items: list[int]) -> dict[int, list[tuple[float, float, float, float, str]]]:
    """
    Trouve les lignes correspondant aux numéros d'items.
    """
    item_lines = {}
    
    for y_pos, text, words in lines:
        # Vérifier si la ligne commence par un numéro d'item
        match = re.match(r"^(\d+)", text.strip())
        if match:
            item_num = int(match.group(1))
            if item_num in target_items:
                item_lines[item_num] = words
                # Pour l'item 8, on cherche aussi les sous-lignes
                if item_num == 8:
                    # Chercher les lignes suivantes qui contiennent PTH, BVH, IVH, Vias
                    pass
        # Pour l'item 8, on cherche aussi les sous-lignes sur la même zone
        if "8" in text and ("PTH" in text or "BVH" in text or "IVH" in text or "Vias" in text):
            if 8 not in item_lines:
                item_lines[8] = words
    
    return item_lines


def highlight_cover_page(document, cover_terms: list[str]) -> int:
    """Surligne les termes de la page de garde."""
    highlighted = 0
    for term in cover_terms:
        if not term or term.upper() in {"NA", "OK"}:
            continue
        for page in document:
            rects = page.search_for(term)
            if rects:
                for rect in rects:
                    add_highlight_rect(page, rect)
                    highlighted += 1
                break
    return highlighted


def highlight_ul94(document) -> int:
    """Surligne UL 94 Flame Class 94V-0."""
    terms = ["UL 94 Flame Class 94V0", "UL 94 Flame Class 94V-0", "94V0", "94V-0"]
    for page in document:
        for term in terms:
            rects = page.search_for(term)
            if rects:
                for rect in rects:
                    add_highlight_rect(page, rect)
                return 1
    return 0


def highlight_inspection_report(document) -> int:
    """Surligne les lignes spécifiques du tableau INSPECTION REPORT."""
    highlighted = 0
    
    # Trouver la page du tableau INSPECTION REPORT
    inspection_page = None
    for page in document:
        text = page.get_text("text")
        if "INSPECTION REPORT" in text:
            inspection_page = page
            break

    if not inspection_page:
        return 0

    # Récupérer les lignes de la page
    lines = get_text_lines_from_page(inspection_page)
    
    # Items à surligner
    target_items = [1, 4, 5, 6, 8, 12, 13, 14, 18, 20, 21, 22, 23, 24]
    
    # Pour chaque ligne, vérifier si elle commence par un item cible
    for y_pos, text, words in lines:
        text_clean = text.strip()
        # Vérifier si la ligne commence par un numéro d'item
        match = re.match(r"^(\d+)", text_clean)
        if match:
            item_num = int(match.group(1))
            if item_num in target_items:
                # Surligner toute la ligne
                for x0, y0, x1, y1, word in words:
                    if word.strip():
                        add_highlight_rect(inspection_page, fitz.Rect(x0, y0, x1, y1))
                        highlighted += 1
    
    # Cas spéciaux pour les sous-lignes (PTH, BVH, IVH, Vias filling)
    # On parcourt toutes les lignes pour trouver celles qui contiennent ces mots
    for y_pos, text, words in lines:
        text_upper = text.upper()
        # Pour l'item 8: PTH (pas BVH, IVH) et Vias filling
        if "PTH" in text_upper and "BVH" not in text_upper and "IVH" not in text_upper:
            for x0, y0, x1, y1, word in words:
                if word.strip():
                    add_highlight_rect(inspection_page, fitz.Rect(x0, y0, x1, y1))
                    highlighted += 1
        if "VIAS" in text_upper and "FILLING" in text_upper:
            for x0, y0, x1, y1, word in words:
                if word.strip():
                    add_highlight_rect(inspection_page, fitz.Rect(x0, y0, x1, y1))
                    highlighted += 1
        
        # Pour l'item 14: Finish et Solder resist
        if "FINISH" in text_upper and "ADHESION" in text_upper:
            for x0, y0, x1, y1, word in words:
                if word.strip():
                    add_highlight_rect(inspection_page, fitz.Rect(x0, y0, x1, y1))
                    highlighted += 1
        if "SOLDER" in text_upper and "RESIST" in text_upper and "ADHESION" in text_upper:
            for x0, y0, x1, y1, word in words:
                if word.strip():
                    add_highlight_rect(inspection_page, fitz.Rect(x0, y0, x1, y1))
                    highlighted += 1
        
        # Pour l'item 23: IONIC CONTAMINATION
        if "10321310" in text or "1B2B1" in text:
            for x0, y0, x1, y1, word in words:
                if word.strip():
                    add_highlight_rect(inspection_page, fitz.Rect(x0, y0, x1, y1))
                    highlighted += 1
        if "AFTER" in text_upper and "FINISH" in text_upper and ("IONIC" in text_upper or "INOIC" in text_upper):
            for x0, y0, x1, y1, word in words:
                if word.strip():
                    add_highlight_rect(inspection_page, fitz.Rect(x0, y0, x1, y1))
                    highlighted += 1
        
        # Pour l'item 24: IMPEDANCE
        if "IMPEDANCE" in text_upper:
            for x0, y0, x1, y1, word in words:
                if word.strip():
                    add_highlight_rect(inspection_page, fitz.Rect(x0, y0, x1, y1))
                    highlighted += 1
    
    return highlighted


def highlight_hole_size(document) -> int:
    """Surligne une ligne aléatoire du tableau HOLE SIZE."""
    import random
    
    for page in document:
        text = page.get_text("text")
        if "HOLE SIZE" not in text and "DRW. DIMENSION" not in text:
            continue

        lines = get_text_lines_from_page(page)
        candidate_lines = []
        
        for y_pos, line_text, words in lines:
            # Chercher une ligne qui a une lettre majuscule suivie de chiffres
            if re.search(r"^[A-Z]\s+\d", line_text):
                candidate_lines.append(words)
            elif re.search(r"^[A-Z]\s+[0-9.]+", line_text):
                candidate_lines.append(words)
        
        if candidate_lines:
            selected = random.choice(candidate_lines)
            for x0, y0, x1, y1, word in selected:
                if word.strip():
                    add_highlight_rect(page, fitz.Rect(x0, y0, x1, y1))
            return 1

    return 0


def highlight_dimension_table(document) -> int:
    """Surligne la ligne avec la plus grande valeur DRW. DIMENSION."""
    for page in document:
        text = page.get_text("text")
        if "DRW. DIMENSION" not in text and "RESULTS" not in text:
            continue

        lines = get_text_lines_from_page(page)
        best_line = None
        best_value = -1.0

        for y_pos, line_text, words in lines:
            # Chercher une ligne avec ITEM et ±
            if "ITEM" in line_text and ("±" in line_text or "+/-" in line_text):
                match = re.search(r"(\d+[.,]\d+)", line_text)
                if match:
                    try:
                        value = float(match.group(1).replace(",", "."))
                        if value > best_value:
                            best_line = words
                            best_value = value
                    except ValueError:
                        continue

        if best_line:
            for x0, y0, x1, y1, word in best_line:
                if word.strip():
                    add_highlight_rect(page, fitz.Rect(x0, y0, x1, y1))
            return 1

    return 0


def highlight_xsection(document) -> int:
    """Surligne HOLE WALL COPPER THICKNESS."""
    for page in document:
        text = page.get_text("text")
        if "HOLE WALL COPPER THICKNESS" in text or "HOLE WALL" in text:
            rects = page.search_for("HOLE WALL COPPER THICKNESS")
            if not rects:
                rects = page.search_for("HOLE WALL")
            if rects:
                for rect in rects:
                    add_highlight_rect(page, rect)
                # Surligner aussi la valeur
                lines = get_text_lines_from_page(page)
                for y_pos, line_text, words in lines:
                    if "HOLE" in line_text and "WALL" in line_text:
                        for x0, y0, x1, y1, word in words:
                            if re.search(r"\d+[.,]\d+", word):
                                add_highlight_rect(page, fitz.Rect(x0, y0, x1, y1))
                return 1
    return 0


def highlight_stackup(document) -> int:
    """Surligne tout le tableau STACKUP."""
    highlighted = 0

    for page in document:
        text = page.get_text("text")
        if "STACKUP" not in text.upper():
            continue

        lines = get_text_lines_from_page(page)
        found_stackup = False
        
        for y_pos, line_text, words in lines:
            if "STACKUP" in line_text.upper():
                found_stackup = True
                continue
            if found_stackup:
                # Surligner toutes les lignes après STACKUP
                for x0, y0, x1, y1, word in words:
                    if word.strip():
                        add_highlight_rect(page, fitz.Rect(x0, y0, x1, y1))
                        highlighted += 1

    return highlighted


def process_pdf(source_pdf: Path, output_pdf: Path, cover_terms: list[str]) -> int:
    """Traite un PDF et applique tous les surlignages."""
    document = fitz.open(source_pdf)
    highlighted = 0

    # 1. Page de garde
    highlighted += highlight_cover_page(document, cover_terms)

    # 2. UL 94
    highlighted += highlight_ul94(document)

    # 3. Tableau INSPECTION REPORT
    highlighted += highlight_inspection_report(document)

    # 4. Tableau HOLE SIZE
    highlighted += highlight_hole_size(document)

    # 5. Tableau des dimensions
    highlighted += highlight_dimension_table(document)

    # 6. XSECTION REPORT
    highlighted += highlight_xsection(document)

    # 7. STACKUP
    highlighted += highlight_stackup(document)

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_pdf, garbage=1, deflate=False)
    document.close()
    return highlighted


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

    highlighted_paths: list[Path] = []
    processed_count = 0

    for uploaded_file in pdf_files:
        original_name = uploaded_file.filename or "document.pdf"
        filename = secure_filename(original_name)
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        pdf_path = upload_dir / filename
        uploaded_file.save(pdf_path)

        try:
            cover_terms = cover_terms_from_filename(original_name)
            highlighted_pdf_path = run_dir / "highlighted" / verified_pdf_name(original_name)
            highlighted_count = process_pdf(pdf_path, highlighted_pdf_path, cover_terms)
            highlighted_paths.append(highlighted_pdf_path)
            processed_count += 1
        except Exception as exc:
            shutil.rmtree(run_dir, ignore_errors=True)
            return render_template(
                "index.html",
                error=f"Erreur pendant le traitement de {original_name}: {exc}",
            )

    stats = {
        "pdf_count": len(pdf_files),
        "highlighted_count": len(highlighted_paths),
        "processed_count": processed_count,
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