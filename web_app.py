from __future__ import annotations

import shutil
import sys
import threading
import time
import uuid
import webbrowser
import zipfile
import re
import logging
from pathlib import Path

import fitz
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

from pcb import (
    parse_filename_values_from_name,
)

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('highlight_debug.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

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


def has_pdf_text(pdf_path: Path, expected_text: str) -> bool:
    try:
        with fitz.open(pdf_path) as document:
            expected = expected_text.upper()
            return any(expected in page.get_text("text").upper() for page in document)
    except Exception as exc:
        logger.warning(f"Recherche texte impossible pour {pdf_path.name}: {exc}")
        return False


def add_highlight_rect(page, rect, color=(1, 0.86, 0.18)) -> None:
    """Ajoute un surlignage sur un rectangle."""
    try:
        annot = page.add_highlight_annot(rect)
        annot.set_colors(stroke=color)
        annot.update()
    except Exception as e:
        logger.error(f"Erreur lors du surlignage: {e}")


def get_text_lines(page, tolerance: float = 3.0) -> dict[float, list[tuple[float, float, float, float, str]]]:
    """Récupère les lignes de texte d'une page."""
    words = page.get_text("words")
    lines = {}
    for word in words:
        mid_y = (word[1] + word[3]) / 2
        key = round(mid_y / tolerance) * tolerance
        if key not in lines:
            lines[key] = []
        lines[key].append((word[0], word[1], word[2], word[3], word[4]))
    return lines


def sorted_text_lines(page) -> list[tuple[float, list[tuple[float, float, float, float, str]], str, str]]:
    """Retourne les lignes visuelles triées: y, mots, texte, texte compact."""
    lines = []
    for y_key, words in get_text_lines(page).items():
        line_words = sorted(words, key=lambda w: w[0])
        line_text = " ".join(w[4] for w in line_words)
        compact = re.sub(r"[^A-Z0-9]+", "", line_text.upper())
        lines.append((y_key, line_words, line_text, compact))
    return sorted(lines, key=lambda item: item[0])


def highlight_line_words(page, line_words: list[tuple[float, float, float, float, str]]) -> int:
    highlighted = 0
    for word in line_words:
        if word[4].strip():
            add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
            highlighted += 1
    return highlighted


def line_item_number(line_text: str) -> int | None:
    if re.match(r"^\s*\d{1,2}\s+(ARRAY|PNL)\b", line_text, flags=re.IGNORECASE):
        return None
    match = re.match(r"^\s*(\d{1,2})(?:\s|$)", line_text)
    return int(match.group(1)) if match else None


def highlight_cover_page(document, cover_terms: list[str]) -> int:
    """Surligne les termes de la page de garde."""
    logger.info(f"Surlignage page de garde avec {len(cover_terms)} termes")
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


def highlight_table_by_item_numbers(document, page_num, target_items, page) -> int:
    """Surligne les lignes correspondant aux numéros d'items dans un tableau."""
    highlighted = 0
    lines = get_text_lines(page)
    
    for y_key in sorted(lines.keys()):
        line_words = sorted(lines[y_key], key=lambda w: w[0])
        line_text = " ".join(w[4] for w in line_words)
        
        match = re.match(r"^(\d+)", line_text.strip())
        if match:
            item_num = int(match.group(1))
            if item_num in target_items:
                logger.info(f"Item {item_num} trouvé page {page_num + 1}")
                for word in line_words:
                    if word[4].strip():
                        add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
                        highlighted += 1
    return highlighted


def is_inspection_page(text: str, inspection_started: bool) -> bool:
    upper = text.upper()
    compact = re.sub(r"[^A-Z0-9]+", "", upper)
    if "CONTENTS" in upper or "WIRING, PRINTED - COMPONENT" in upper or "CERTIFICATE" in upper:
        return False
    if "HOLE SIZE" in upper or "ARTICLE XSECTION" in upper or "STACKUP" in upper:
        return False
    first_table_terms = [
        "LAMINATEMATERIAL",
        "CONDUCTORWIDTH",
        "CONDUCTORSPACE",
        "ANNULARRING",
        "SOLDERABILITYTEST",
        "ELECTRICTEST",
        "ADHESION",
    ]
    continuation_terms = [
        "WARPTWIST",
        "IMPEDANCE",
        "INOIC",
        "IONIC",
        "SOLDERMASKTHICKNESS",
        "GOLDTHICKNESS",
        "NICKELTHICKNESS",
    ]
    if "INSPECTION" in upper and ("REPORT" in upper or "ITEM" in upper or "DESCRIPTION" in upper):
        return True
    if any(term in compact for term in first_table_terms):
        return True
    return inspection_started and (
        "ITEM DESCRIPTION" in upper
        or any(term in compact for term in continuation_terms)
    )


def highlight_inspection_item_line(page, lines, item_num: int) -> int:
    highlighted = 0
    for _y_key, line_words, line_text, _compact in lines:
        if line_item_number(line_text) == item_num:
            highlighted += highlight_line_words(page, line_words)
    return highlighted


def highlight_lines_in_item_block(page, lines, item_num: int, required_compacts: list[str]) -> int:
    highlighted = 0
    in_block = False
    for _y_key, line_words, line_text, compact in lines:
        current_item = line_item_number(line_text)
        if current_item == item_num:
            in_block = True
        elif current_item is not None and in_block:
            break

        if not in_block:
            continue
        if any(required in compact for required in required_compacts):
            highlighted += highlight_line_words(page, line_words)
    return highlighted


def highlight_ionic_contamination_block(page, lines) -> int:
    highlighted = 0
    for _y_key, line_words, line_text, compact in lines:
        is_ionic_line = any(term in compact for term in ["INOIC", "IONIC", "CONTAMINATION", "10321310", "1B2B1", "AFTERFINISH", "GR78CORE", "NACL"])
        if is_ionic_line or re.fullmatch(r"\s*23\s*", line_text):
            highlighted += highlight_line_words(page, line_words)
    return highlighted


def highlight_split_sublines(page, lines) -> int:
    """Surligne les sous-lignes que le PDF place parfois avant le numéro d'item."""
    highlighted = 0
    for _y_key, line_words, line_text, compact in lines:
        stripped = line_text.strip().upper()
        if stripped == "PTH":
            highlighted += highlight_line_words(page, line_words)
            continue
        if compact.startswith("FINISH") and "NOPEELING" in compact and "TAPE" not in compact and "COPPER" not in compact:
            highlighted += highlight_line_words(page, line_words)
            continue
    return highlighted


def highlight_impedance_lines(page, lines) -> int:
    highlighted = 0
    for _y_key, line_words, _line_text, compact in lines:
        if "IMPEDANCE" in compact:
            highlighted += highlight_line_words(page, line_words)
    return highlighted


def highlight_inspection_report(document) -> int:
    """Surligne les items du tableau INSPECTION REPORT."""
    logger.info("Surlignage INSPECTION REPORT")
    highlighted = 0

    full_item_lines = {1, 4, 5, 6, 12, 13, 18, 20, 21, 22, 24}
    inspection_started = False

    for page_num, page in enumerate(document):
        text = page.get_text("text")
        if not is_inspection_page(text, inspection_started):
            continue
        inspection_started = True
        logger.info(f"Page {page_num + 1}: INSPECTION REPORT trouvé")

        lines = sorted_text_lines(page)

        for item_num in full_item_lines:
            highlighted += highlight_inspection_item_line(page, lines, item_num)

        highlighted += highlight_lines_in_item_block(page, lines, 8, ["PTH", "VIASFILLING"])
        highlighted += highlight_lines_in_item_block(page, lines, 14, ["FINISH", "SOLDERRESIST"])
        highlighted += highlight_split_sublines(page, lines)
        highlighted += highlight_impedance_lines(page, lines)

        if any("CONTAMINATION" in compact for *_unused, compact in lines):
            highlighted += highlight_ionic_contamination_block(page, lines)

    return highlighted


def highlight_hole_size(document) -> int:
    """Surligne une ligne aléatoire du tableau HOLE SIZE."""
    import random
    logger.info("Surlignage HOLE SIZE")
    
    for page_num, page in enumerate(document):
        text = page.get_text("text")
        if "HOLE SIZE" not in text and "UNIT:INCH" not in text:
            continue

        logger.info(f"Page {page_num + 1}: HOLE SIZE trouvé")
        lines = get_text_lines(page)
        candidate_lines = []
        
        for y_key in sorted(lines.keys()):
            line_words = sorted(lines[y_key], key=lambda w: w[0])
            line_text = " ".join(w[4] for w in line_words)
            
            if re.search(r"^[A-Z]\s+[\d.]+", line_text.strip()):
                candidate_lines.append(line_words)

        if candidate_lines:
            selected = random.choice(candidate_lines)
            for word in selected:
                if word[4].strip():
                    add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
            logger.info("HOLE SIZE: 1 ligne surlignée")
            return 1

    return 0


def highlight_dimension_table(document) -> int:
    """Surligne la ligne avec la plus grande valeur dans le tableau des dimensions."""
    logger.info("Surlignage TABLEAU DES DIMENSIONS")
    
    for page_num, page in enumerate(document):
        text = page.get_text("text")
        if not ("DRW. DIMENSION" in text or "NO." in text or "STD.TOL." in text or "UNIT:MM" in text):
            continue
        
        logger.info(f"Page {page_num + 1}: Tableau des dimensions trouvé")
        
        lines = get_text_lines(page)
        best_line = None
        best_value = -1.0
        
        for y_key in sorted(lines.keys()):
            line_words = sorted(lines[y_key], key=lambda w: w[0])
            line_text = " ".join(w[4] for w in line_words)
            
            match = re.search(r"^(\d+)\s+([\d.]+)\s*[±]", line_text)
            if match:
                try:
                    value = float(match.group(2).replace(",", "."))
                    if value > best_value:
                        best_line = line_words
                        best_value = value
                except ValueError:
                    continue

        if best_line:
            logger.info(f"Meilleure ligne trouvée avec valeur {best_value}")
            for word in best_line:
                if word[4].strip():
                    add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
            return 1
    
    return 0


def highlight_xsection(document) -> int:
    """Surligne HOLE WALL COPPER THICKNESS."""
    logger.info("Surlignage XSECTION")
    for page in document:
        text = page.get_text("text")
        if "HOLE WALL COPPER THICKNESS" in text:
            rects = page.search_for("HOLE WALL COPPER THICKNESS")
            if rects:
                for rect in rects:
                    add_highlight_rect(page, rect)
                lines = get_text_lines(page)
                for y_key in lines:
                    line_words = sorted(lines[y_key], key=lambda w: w[0])
                    line_text = " ".join(w[4] for w in line_words)
                    if re.search(r"\d+[.,]\d+\s*mil", line_text):
                        for word in line_words:
                            if re.search(r"\d+[.,]\d+", word[4]):
                                add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
                return 1
    return 0


def highlight_stackup(document) -> int:
    """Surligne tout le tableau STACKUP."""
    logger.info("Surlignage STACKUP")
    highlighted = 0

    for page_num, page in enumerate(document):
        text = page.get_text("text")
        if "STACKUP" not in text.upper() and "STACK-UP" not in text.upper():
            continue

        logger.info(f"Page {page_num + 1}: STACKUP trouvé")
        lines = get_text_lines(page)
        start_highlight = False
        
        for y_key in sorted(lines.keys()):
            line_words = sorted(lines[y_key], key=lambda w: w[0])
            line_text = " ".join(w[4] for w in line_words)
            
            if "STACKUP" in line_text.upper() or "STACK-UP" in line_text.upper():
                start_highlight = True
                continue
            
            if start_highlight:
                for word in line_words:
                    if word[4].strip():
                        add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
                        highlighted += 1

    logger.info(f"STACKUP: {highlighted} surlignages")
    return highlighted


def process_pdf(source_pdf: Path, output_pdf: Path, cover_terms: list[str]) -> int:
    """Traite un PDF et applique tous les surlignages."""
    logger.info(f"=== Traitement de {source_pdf.name} ===")
    document = fitz.open(source_pdf)
    logger.info(f"PDF ouvert avec {len(document)} pages")
    
    highlighted = 0

    try:
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

        logger.info(f"Total surlignages: {highlighted}")
        
        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        document.save(output_pdf, garbage=1, deflate=False)
        logger.info(f"PDF sauvegardé: {output_pdf}")
        
    except Exception as e:
        logger.error(f"Erreur lors du traitement: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise
    finally:
        document.close()
    
    return highlighted


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    try:
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
                logger.error(f"Erreur: {exc}")
                import traceback
                logger.error(traceback.format_exc())
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
            warnings=[],
            cover_rows=[],
            standard_rows=[],
            inspection_rows=[],
        )
    except Exception as e:
        logger.error(f"Erreur générale: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return render_template(
            "index.html",
            error=f"Erreur serveur: {e}",
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