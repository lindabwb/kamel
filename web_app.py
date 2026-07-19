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

from pcb import parse_filename_values_from_name

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


def add_highlight_rect(page, rect, color=(1, 0.86, 0.18)) -> None:
    """Ajoute un surlignage sur un rectangle."""
    try:
        annot = page.add_highlight_annot(rect)
        annot.set_colors(stroke=color)
        annot.update()
    except Exception as e:
        logger.error(f"Erreur lors du surlignage: {e}")


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
                logger.debug(f"Terme '{term}' trouvé sur la page {page.number + 1}")
                for rect in rects:
                    add_highlight_rect(page, rect)
                    highlighted += 1
                break
    logger.info(f"Page de garde: {highlighted} surlignages")
    return highlighted


def highlight_ul94(document) -> int:
    """Surligne UL 94 Flame Class 94V-0."""
    logger.info("Surlignage UL 94")
    terms = ["UL 94 Flame Class 94V0", "UL 94 Flame Class 94V-0", "94V0", "94V-0"]
    for page in document:
        for term in terms:
            rects = page.search_for(term)
            if rects:
                logger.info(f"UL 94 trouvé sur la page {page.number + 1} avec '{term}'")
                for rect in rects:
                    add_highlight_rect(page, rect)
                return 1
    logger.warning("UL 94 non trouvé")
    return 0


def highlight_inspection_report(document) -> int:
    """Surligne les items du tableau INSPECTION REPORT sur toutes les pages."""
    logger.info("Surlignage du tableau INSPECTION REPORT")
    highlighted = 0
    
    # Items cibles (première et deuxième partie)
    target_items = [1, 4, 5, 6, 8, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]
    
    # Parcourir toutes les pages
    for page_num, page in enumerate(document):
        text = page.get_text("text")
        if "INSPECTION REPORT" not in text:
            continue
        
        logger.info(f"Page {page_num + 1}: contient INSPECTION REPORT")
        words = page.get_text("words")
        
        # Grouper les mots par ligne
        lines = {}
        for word in words:
            mid_y = (word[1] + word[3]) / 2
            key = round(mid_y / 2) * 2
            if key not in lines:
                lines[key] = []
            lines[key].append(word)
        
        logger.info(f"Page {page_num + 1}: {len(lines)} lignes détectées")
        
        # Pour chaque ligne, vérifier si elle commence par un numéro d'item cible
        for y_key in sorted(lines.keys()):
            line_words = sorted(lines[y_key], key=lambda w: w[0])
            line_text = " ".join(w[4] for w in line_words)
            
            # Vérifier si la ligne commence par un numéro
            match = re.match(r"^(\d+)", line_text.strip())
            if match:
                item_num = int(match.group(1))
                if item_num in target_items:
                    logger.info(f"Item {item_num} trouvé sur la page {page_num + 1}: '{line_text[:50]}...'")
                    # Surligner tous les mots de la ligne
                    for word in line_words:
                        if word[4].strip():
                            add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
                            highlighted += 1
                    continue
            
            # Cas spécial: sous-lignes de l'item 8 (PTH, BVH, IVH, Vias filling)
            if re.search(r"(PTH|BVH|IVH|Vias)", line_text, re.IGNORECASE):
                # Vérifier si c'est une sous-ligne de l'item 8
                # Chercher la ligne précédente qui contient "8"
                for prev_key in sorted(lines.keys()):
                    if prev_key >= y_key:
                        break
                    prev_line_words = sorted(lines[prev_key], key=lambda w: w[0])
                    prev_line_text = " ".join(w[4] for w in prev_line_words)
                    if re.match(r"^8\s+", prev_line_text):
                        logger.info(f"Sous-ligne de l'item 8 trouvée: '{line_text[:50]}...'")
                        for word in line_words:
                            if word[4].strip():
                                add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
                                highlighted += 1
                        break
            
            # Cas spécial: sous-lignes de l'item 23 (10321310, AFTER FINISH, IONIC)
            if re.search(r"(10321310|AFTER.*FINISH|IONIC|INOIC)", line_text, re.IGNORECASE):
                for prev_key in sorted(lines.keys()):
                    if prev_key >= y_key:
                        break
                    prev_line_words = sorted(lines[prev_key], key=lambda w: w[0])
                    prev_line_text = " ".join(w[4] for w in prev_line_words)
                    if re.match(r"^23\s+", prev_line_text):
                        logger.info(f"Sous-ligne de l'item 23 trouvée: '{line_text[:50]}...'")
                        for word in line_words:
                            if word[4].strip():
                                add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
                                highlighted += 1
                        break
            
            # Cas spécial: item 14 sous-lignes (Finish, Solder resist)
            if re.search(r"(Finish|Solder.*resist)", line_text, re.IGNORECASE):
                for prev_key in sorted(lines.keys()):
                    if prev_key >= y_key:
                        break
                    prev_line_words = sorted(lines[prev_key], key=lambda w: w[0])
                    prev_line_text = " ".join(w[4] for w in prev_line_words)
                    if re.match(r"^14\s+", prev_line_text):
                        logger.info(f"Sous-ligne de l'item 14 trouvée: '{line_text[:50]}...'")
                        for word in line_words:
                            if word[4].strip():
                                add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
                                highlighted += 1
                        break
    
    logger.info(f"INSPECTION REPORT: {highlighted} surlignages")
    return highlighted


def highlight_hole_size(document) -> int:
    """Surligne une ligne aléatoire du tableau HOLE SIZE."""
    import random
    logger.info("Surlignage du tableau HOLE SIZE")
    
    for page_num, page in enumerate(document):
        text = page.get_text("text")
        if "HOLE SIZE" not in text:
            continue

        logger.info(f"Tableau HOLE SIZE trouvé sur la page {page_num + 1}")
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
            logger.info(f"{len(candidate_lines)} lignes candidates trouvées")
            selected = random.choice(candidate_lines)
            for word in selected:
                if word[4].strip():
                    add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
            logger.info("HOLE SIZE: 1 ligne surlignée")
            return 1

    logger.warning("HOLE SIZE: aucune ligne trouvée")
    return 0


def highlight_dimension_table(document) -> int:
    """Surligne la ligne avec la plus grande valeur DRW. DIMENSION."""
    logger.info("Surlignage du tableau ITEM / DRW. DIMENSION")
    
    for page_num, page in enumerate(document):
        text = page.get_text("text")
        if "DRW. DIMENSION" not in text and "UNIT : MM" not in text:
            continue
        
        logger.info(f"Tableau ITEM / DRW. DIMENSION trouvé sur la page {page_num + 1}")
        
        words = page.get_text("words")
        logger.info(f"{len(words)} mots extraits")
        
        # Grouper par ligne
        lines = {}
        for word in words:
            mid_y = (word[1] + word[3]) / 2
            key = round(mid_y / 2) * 2
            if key not in lines:
                lines[key] = []
            lines[key].append(word)
        
        logger.info(f"{len(lines)} lignes détectées")
        
        # Chercher la ligne avec la plus grande valeur
        best_line = None
        best_value = -1.0
        best_item = None
        
        for y_key in sorted(lines.keys()):
            line_words = sorted(lines[y_key], key=lambda w: w[0])
            line_text = " ".join(w[4] for w in line_words)
            
            # Chercher une ligne qui commence par un numéro et contient "±"
            match = re.match(r"^(\d+)\s+([\d.]+)\s*[±]", line_text)
            if match:
                try:
                    item_num = int(match.group(1))
                    value = float(match.group(2).replace(",", "."))
                    logger.debug(f"Ligne {item_num}: valeur={value}")
                    if value > best_value:
                        best_line = line_words
                        best_value = value
                        best_item = item_num
                except ValueError:
                    continue
        
        if best_line:
            logger.info(f"Meilleure ligne: ITEM {best_item} avec valeur {best_value}")
            for word in best_line:
                if word[4].strip():
                    add_highlight_rect(page, fitz.Rect(word[0], word[1], word[2], word[3]))
            logger.info("ITEM / DRW. DIMENSION: 1 ligne surlignée")
            return 1
    
    logger.warning("TABLEAU ITEM / DRW. DIMENSION: aucune ligne trouvée")
    return 0


def highlight_xsection(document) -> int:
    """Surligne HOLE WALL COPPER THICKNESS."""
    logger.info("Surlignage XSECTION REPORT")
    for page in document:
        text = page.get_text("text")
        if "HOLE WALL COPPER THICKNESS" in text:
            logger.info(f"XSECTION trouvé sur la page {page.number + 1}")
            rects = page.search_for("HOLE WALL COPPER THICKNESS")
            if rects:
                for rect in rects:
                    add_highlight_rect(page, rect)
                logger.info("XSECTION: surligné")
                return 1
    logger.warning("XSECTION: non trouvé")
    return 0


def highlight_stackup(document) -> int:
    """Surligne tout le tableau STACKUP."""
    logger.info("Surlignage STACKUP")
    highlighted = 0

    for page in document:
        text = page.get_text("text")
        if "STACKUP" not in text.upper():
            continue

        logger.info(f"STACKUP trouvé sur la page {page.number + 1}")
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

        # 3. Tableau INSPECTION REPORT (toutes les pages)
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
            logger.info(f"Fichier sauvegardé: {pdf_path}")

            try:
                cover_terms = cover_terms_from_filename(original_name)
                logger.info(f"Termes de la page de garde: {cover_terms}")
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