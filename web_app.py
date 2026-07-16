from __future__ import annotations

import shutil
import sys
import threading
import time
import uuid
import webbrowser
import zipfile
import re
import random  # <--- AJOUTER CETTE LIGNE
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


def page_words(page) -> list[tuple[float, float, float, float, str]]:
    """Récupère tous les mots d'une page."""
    return [
        (float(word[0]), float(word[1]), float(word[2]), float(word[3]), str(word[4]))
        for word in page.get_text("words")
    ]


def get_line_from_rect(page, rect, tolerance: float = 3.5) -> list[tuple[float, float, float, float, str]]:
    """Récupère tous les mots sur la même ligne qu'un rectangle."""
    anchor_mid = (rect.y0 + rect.y1) / 2
    words = page_words(page)
    return [
        (x0, y0, x1, y1, text)
        for x0, y0, x1, y1, text in words
        if abs(((y0 + y1) / 2) - anchor_mid) <= tolerance
    ]


def highlight_line(page, rect) -> int:
    """Surligne toute la ligne contenant le rectangle."""
    line_words = get_line_from_rect(page, rect)
    highlighted = 0
    for x0, y0, x1, y1, text in line_words:
        if text.strip():
            add_highlight_rect(page, fitz.Rect(x0, y0, x1, y1))
            highlighted += 1
    return highlighted


def find_first_rect(document, term: str) -> tuple[fitz.Page | None, fitz.Rect | None]:
    """Trouve le premier rectangle contenant le terme recherché."""
    for page in document:
        rects = page.search_for(term)
        if rects:
            return page, rects[0]
    return None, None


def find_rects(document, term: str) -> list[tuple[int, fitz.Rect]]:
    """Trouve tous les rectangles contenant le terme recherché."""
    matches: list[tuple[int, fitz.Rect]] = []
    for page_index, page in enumerate(document):
        for variant in search_variants(term):
            rects = page.search_for(variant)
            if rects:
                matches.extend((page_index, rect) for rect in rects)
                break
    return matches


def search_variants(term: str) -> list[str]:
    """Génère des variantes d'un terme pour la recherche."""
    variants = [
        term,
        term.replace(" ", ""),
        term.replace(" / ", "/"),
        term.replace(" ", " "),
    ]
    if "94V0" in term.upper() or "94V-0" in term.upper():
        variants.extend(["94V0", "94V-0", "UL 94 Flame Class 94V0", "UL 94 Flame Class 94V-0"])
    return list(dict.fromkeys(v for v in variants if v))


def highlight_cover_page(document, cover_terms: list[str]) -> int:
    """Surligne les termes de la page de garde."""
    highlighted = 0
    for term in cover_terms:
        if not term or term.upper() in {"NA", "OK"}:
            continue
        page, rect = find_first_rect(document, term)
        if page and rect:
            add_highlight_rect(page, rect)
            highlighted += 1
    return highlighted


def highlight_ul94(document) -> int:
    """Surligne UL 94 Flame Class 94V-0."""
    terms = ["UL 94 Flame Class 94V0", "UL 94 Flame Class 94V-0", "94V0", "94V-0"]
    for term in terms:
        page, rect = find_first_rect(document, term)
        if page and rect:
            add_highlight_rect(page, rect)
            return 1
    return 0


def highlight_inspection_report(document) -> int:
    """Surligne les lignes spécifiques du tableau INSPECTION REPORT."""
    highlighted = 0
    inspection_page = None
    inspection_index = -1

    # Trouver la page du tableau INSPECTION REPORT
    for idx, page in enumerate(document):
        text = page.get_text("text")
        if "INSPECTION REPORT" in text:
            inspection_page = page
            inspection_index = idx
            break

    if not inspection_page:
        return 0

    # Items à surligner avec leurs mots-clés
    target_items = [
        # item 1
        ["LAMINATE", "MATERIAL"],
        # item 4
        ["CONDUCTOR", "WIDTH"],
        # item 5
        ["CONDUCTOR", "SPACE"],
        # item 6
        ["ANNULAR", "RING"],
        # item 8 - PTH et Vias filling (pas BVH, IVH)
        ["PTH"],
        ["Vias", "filling"],
        # item 12
        ["SOLDERABILITY", "TEST"],
        # item 13
        ["ELECTRIC", "TEST"],
        # item 14 - Finish et Solder resist
        ["Finish"],
        ["Solder", "resist"],
        # item 18
        ["WARP", "TWIST"],
        # item 20
        ["SOLDER", "MASK", "THICKNESS"],
        # item 21
        ["GOLD", "THICKNESS"],
        # item 22
        ["NICKEL", "THICKNESS"],
        # item 23 - ionic contamination
        ["10321310", "1B2B1"],
        ["AFTER", "FINISH"],
        # item 24 - IMPEDANCE
        ["IMPEDANCE"],
    ]

    # Pour chaque item, trouver la ligne et la surligner
    for keywords in target_items:
        # Chercher chaque mot-clé
        for keyword in keywords:
            rects = inspection_page.search_for(keyword)
            if rects:
                # Vérifier que c'est la bonne ligne
                for rect in rects:
                    line_text = " ".join(text for _, _, _, _, text in get_line_from_rect(inspection_page, rect))
                    # Vérifier que tous les mots-clés sont sur la ligne
                    if all(kw in line_text for kw in keywords):
                        highlighted += highlight_line(inspection_page, rect)
                        break
                break

    return highlighted


def highlight_hole_size(document) -> int:
    """Surligne une ligne aléatoire du tableau HOLE SIZE."""
    for page in document:
        text = page.get_text("text")
        if "HOLE SIZE" not in text and "DRW. DIMENSION" not in text:
            continue

        words = page_words(page)
        # Chercher les lignes avec des valeurs de trous (format: A 0.0098±0.003 ...)
        candidate_lines: list[list[tuple[float, float, float, float, str]]] = []
        used_y: set[float] = set()

        for i, (x0, y0, x1, y1, text) in enumerate(words):
            # Chercher des motifs de trous: lettre majuscule suivie de chiffres
            if re.match(r"^[A-Z]\s*$", text.strip()) and i + 1 < len(words):
                # Vérifier que la ligne suivante contient un nombre
                next_text = words[i + 1][4]
                if re.search(r"\d", next_text):
                    # Prendre toute la ligne
                    mid = (y0 + y1) / 2
                    if mid not in used_y:
                        line_words = get_line_from_rect(page, fitz.Rect(x0, y0, x1, y1))
                        if line_words:
                            candidate_lines.append(line_words)
                            used_y.add(mid)

        if candidate_lines:
            # Choisir une ligne aléatoire
            selected = random.choice(candidate_lines)
            for x0, y0, x1, y1, text in selected:
                if text.strip():
                    add_highlight_rect(page, fitz.Rect(x0, y0, x1, y1))
            return 1

    return 0


def highlight_dimension_table(document) -> int:
    """Surligne la ligne avec la plus grande valeur DRW. DIMENSION."""
    for page in document:
        text = page.get_text("text")
        if "DRW. DIMENSION" not in text and "RESULTS" not in text:
            continue

        words = page_words(page)
        # Chercher les lignes avec des dimensions (format: ITEM 146.09±0.127 ...)
        best_line = None
        best_value = -1.0

        for i, (x0, y0, x1, y1, text) in enumerate(words):
            # Chercher des nombres avec ±
            if "±" in text or "+/-" in text:
                # Extraire la valeur numérique
                match = re.search(r"(\d+[.,]\d+)", text)
                if match:
                    try:
                        value = float(match.group(1).replace(",", "."))
                        if value > best_value:
                            # Prendre toute la ligne
                            rect = fitz.Rect(x0, y0, x1, y1)
                            line_words = get_line_from_rect(page, rect)
                            if line_words:
                                # Vérifier que c'est une ligne de dimension (contient ITEM ou un numéro)
                                line_text = " ".join(t for _, _, _, _, t in line_words)
                                if "ITEM" in line_text or re.search(r"\d+\s*[±]", line_text):
                                    best_line = line_words
                                    best_value = value
                    except ValueError:
                        continue

        if best_line:
            for x0, y0, x1, y1, text in best_line:
                if text.strip():
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
                add_highlight_rect(page, rects[0])
                # Surligner aussi la valeur
                line_words = get_line_from_rect(page, rects[0])
                for x0, y0, x1, y1, t in line_words:
                    if re.search(r"\d+[.,]\d+\s*mil", t) or re.search(r"\d+[.,]\d+", t):
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

        # Trouver le début du tableau
        words = page_words(page)
        stackup_start_y = None

        for x0, y0, x1, y1, word in words:
            if word.strip().upper() == "STACKUP":
                stackup_start_y = y0
                break

        if stackup_start_y is None:
            continue

        # Surligner toutes les lignes du tableau
        for x0, y0, x1, y1, word in words:
            # Ne pas surligner au-dessus du titre STACKUP
            if y0 < stackup_start_y:
                continue
            # Ne pas surligner trop bas (pied de page)
            if y0 > page.rect.height - 50:
                continue
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
    import random

    def open_browser() -> None:
        time.sleep(1.5)
        webbrowser.open("http://127.0.0.1:5000")

    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False)