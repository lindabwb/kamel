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
import pdfplumber
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


def get_inspection_items_with_coordinates(pdf_path: Path) -> dict:
    """
    Extrait les items du tableau INSPECTION REPORT avec leurs coordonnées précises.
    Retourne: {item_num: {"page": page_num, "y0": y0, "y1": y1, "text": text}}
    """
    items = {}
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            if "INSPECTION REPORT" not in text:
                continue
            
            # Extraire les mots avec leurs coordonnées
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            
            # Grouper les mots par ligne
            lines = {}
            for word in words:
                mid_y = (word["top"] + word["bottom"]) / 2
                key = round(mid_y / 2) * 2
                if key not in lines:
                    lines[key] = []
                lines[key].append(word)
            
            # Trier les lignes
            sorted_keys = sorted(lines.keys())
            
            # Pour chaque ligne, vérifier si elle commence par un numéro d'item
            for y_key in sorted_keys:
                line_words = sorted(lines[y_key], key=lambda w: w["x0"])
                line_text = " ".join(w["text"] for w in line_words)
                line_text_clean = re.sub(r'\s+', ' ', line_text).strip()
                
                # Vérifier si la ligne commence par un numéro
                match = re.match(r"^(\d+)", line_text_clean)
                if match:
                    item_num = int(match.group(1))
                    items[item_num] = {
                        "page": page_num,
                        "y0": min(w["top"] for w in line_words),
                        "y1": max(w["bottom"] for w in line_words),
                        "text": line_text_clean,
                        "words": line_words
                    }
                
                # Cas spécial: item 8 avec ses sous-lignes (PTH, BVH, IVH, Vias)
                if "8" in line_text_clean and ("PTH" in line_text_clean or "BVH" in line_text_clean or "IVH" in line_text_clean or "Vias" in line_text_clean):
                    if 8 not in items:
                        items[8] = {
                            "page": page_num,
                            "y0": min(w["top"] for w in line_words),
                            "y1": max(w["bottom"] for w in line_words),
                            "text": line_text_clean,
                            "words": line_words
                        }
                    else:
                        # Étendre la zone de l'item 8
                        items[8]["y0"] = min(items[8]["y0"], min(w["top"] for w in line_words))
                        items[8]["y1"] = max(items[8]["y1"], max(w["bottom"] for w in line_words))
                        items[8]["text"] += " " + line_text_clean
                        items[8]["words"].extend(line_words)
                
                # Cas spécial: item 23 avec ses sous-lignes
                if "23" in line_text_clean and ("10321310" in line_text_clean or "AFTER" in line_text_clean or "INOIC" in line_text_clean):
                    if 23 not in items:
                        items[23] = {
                            "page": page_num,
                            "y0": min(w["top"] for w in line_words),
                            "y1": max(w["bottom"] for w in line_words),
                            "text": line_text_clean,
                            "words": line_words
                        }
                    else:
                        items[23]["y0"] = min(items[23]["y0"], min(w["top"] for w in line_words))
                        items[23]["y1"] = max(items[23]["y1"], max(w["bottom"] for w in line_words))
                        items[23]["text"] += " " + line_text_clean
                        items[23]["words"].extend(line_words)
    
    return items


def highlight_inspection_report_with_coordinates(document: fitz.Document, pdf_path: Path) -> int:
    """Utilise les coordonnées extraites par pdfplumber pour surligner."""
    highlighted = 0
    
    # Extraire les items avec leurs coordonnées
    items = get_inspection_items_with_coordinates(pdf_path)
    
    if not items:
        return 0
    
    # Items cibles
    target_items = [1, 4, 5, 6, 8, 12, 13, 14, 18, 20, 21, 22, 23, 24]
    
    # Pour chaque item cible
    for item_num in target_items:
        if item_num not in items:
            continue
        
        item_data = items[item_num]
        page_num = item_data["page"]
        
        if page_num >= len(document):
            continue
        
        page = document[page_num]
        
        # Pour chaque mot de l'item, surligner
        for word in item_data["words"]:
            # Convertir les coordonnées pdfplumber en fitz
            x0 = word["x0"]
            y0 = word["top"]
            x1 = word["x1"]
            y1 = word["bottom"]
            
            # Ajouter une marge pour être sûr de surligner tout le mot
            rect = fitz.Rect(x0, y0 - 1, x1, y1 + 1)
            add_highlight_rect(page, rect)
            highlighted += 1
    
    return highlighted


def highlight_hole_size(document) -> int:
    """Surligne une ligne aléatoire du tableau HOLE SIZE."""
    import random
    
    for page in document:
        text = page.get_text("text")
        if "HOLE SIZE" not in text and "DRW. DIMENSION" not in text:
            continue

        words = page.get_text("words")
        candidate_lines = []
        used_y = set()
        
        for i, word in enumerate(words):
            word_text = word[4].strip()
            if re.match(r"^[A-Z]$", word_text) and i + 1 < len(words):
                next_text = words[i + 1][4].strip()
                if re.search(r"\d", next_text):
                    mid_y = (word[1] + word[3]) / 2
                    y_key = round(mid_y / 2) * 2
                    if y_key not in used_y:
                        line_words = []
                        for w in words:
                            w_mid = (w[1] + w[3]) / 2
                            if abs(w_mid - mid_y) <= 5:
                                line_words.append(w)
                        if line_words:
                            candidate_lines.append(line_words)
                            used_y.add(y_key)

        if candidate_lines:
            selected = random.choice(candidate_lines)
            for word in selected:
                if word[4].strip():
                    add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
            return 1

    return 0


def highlight_dimension_table(document) -> int:
    """Surligne la ligne avec la plus grande valeur DRW. DIMENSION."""
    for page in document:
        text = page.get_text("text")
        if "DRW. DIMENSION" not in text and "RESULTS" not in text:
            continue

        words = page.get_text("words")
        best_line = None
        best_value = -1.0

        for i, word in enumerate(words):
            word_text = word[4].strip()
            if "±" in word_text or "+/-" in word_text:
                match = re.search(r"(\d+[.,]\d+)", word_text)
                if match:
                    try:
                        value = float(match.group(1).replace(",", "."))
                        if value > best_value:
                            mid_y = (word[1] + word[3]) / 2
                            line_words = []
                            for w in words:
                                w_mid = (w[1] + w[3]) / 2
                                if abs(w_mid - mid_y) <= 5:
                                    line_words.append(w)
                            if line_words:
                                line_text = " ".join(w[4] for w in line_words)
                                if "ITEM" in line_text or re.search(r"\d+\s*[±]", line_text):
                                    best_line = line_words
                                    best_value = value
                    except ValueError:
                        continue

        if best_line:
            for word in best_line:
                if word[4].strip():
                    add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
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
                words = page.get_text("words")
                for word in words:
                    if re.search(r"\d+[.,]\d+\s*mil", word[4]) or re.search(r"\d+[.,]\d+", word[4]):
                        add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
                return 1
    return 0


def highlight_stackup(document) -> int:
    """Surligne tout le tableau STACKUP."""
    highlighted = 0

    for page in document:
        text = page.get_text("text")
        if "STACKUP" not in text.upper():
            continue

        words = page.get_text("words")
        stackup_start_y = None

        for word in words:
            if word[4].strip().upper() == "STACKUP":
                stackup_start_y = word[1]
                break

        if stackup_start_y is None:
            continue

        for word in words:
            if word[1] < stackup_start_y:
                continue
            if word[3] > page.rect.height - 50:
                continue
            if word[4].strip():
                add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
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

    # 3. Tableau INSPECTION REPORT (avec coordonnées pdfplumber)
    highlighted += highlight_inspection_report_with_coordinates(document, source_pdf)

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