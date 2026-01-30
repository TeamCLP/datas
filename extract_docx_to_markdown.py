#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversion DOCX -> Markdown (dataset pour training LLM)
- Scanne les dossiers EDB et NDC
- Identifie les fichiers par leur code RITM (CAGIPRITMNNNNNNN)
- Utilise Mammoth pour conversion DOCX (meilleure qualité)
- Supprime page de garde, table des matières, préambule
- Parallélisé avec ProcessPoolExecutor

Dépendances:
  pip install mammoth html2text
"""

from __future__ import annotations

import re
import os
import logging
import traceback
import argparse
from pathlib import Path
from typing import List, Tuple, Optional
from concurrent.futures import ProcessPoolExecutor, as_completed

import pandas as pd
import mammoth
import html2text

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ==============================================================================
# CONFIGURATION
# ==============================================================================

# Dossiers source (DOCX classés)
DEFAULT_EDB_DIR = "classified_docx/edb"
DEFAULT_NDC_DIR = "classified_docx/ndc"

# Dossiers de sortie (Markdown)
DEFAULT_OUTPUT_DIR = "markdown"
LOG_DIRNAME = "_logs"

# Pattern pour extraire le code RITM du nom de fichier
# Format: CAGIPRITM suivi de chiffres (ex: CAGIPRITM0012345)
RITM_PATTERN = r"^(CAGIPRITM\d+)"

# Nombre de workers (0 = auto = nombre de CPU)
DEFAULT_WORKERS = 0

# Style mapping Mammoth : ignorer les styles de TOC et mapper les autres
MAMMOTH_STYLE_MAP = """
p[style-name='toc 1'] => !
p[style-name='toc 2'] => !
p[style-name='toc 3'] => !
p[style-name='toc 4'] => !
p[style-name='TM1'] => !
p[style-name='TM2'] => !
p[style-name='TM3'] => !
p[style-name='TM4'] => !
p[style-name='TOC Heading'] => !
p[style-name='TOC 1'] => !
p[style-name='TOC 2'] => !
p[style-name='TOC 3'] => !
p[style-name='Title'] => h1
p[style-name='Titre'] => h1
p[style-name='Heading 1'] => h1
p[style-name='Heading 2'] => h2
p[style-name='Heading 3'] => h3
p[style-name='Heading 4'] => h4
p[style-name='Titre 1'] => h1
p[style-name='Titre 2'] => h2
p[style-name='Titre 3'] => h3
p[style-name='Titre 4'] => h4
p[style-name='List Paragraph'] => p
p[style-name='Paragraphedeliste'] => p
p[style-name='No Spacing'] => p
p[style-name='Body Text'] => p
"""

# Patterns pour détecter la fin du préambule/TOC
TOC_END_MARKERS = [
    "Table des matières",
    "Table des matieres",
    "Sommaire",
    "TABLE DES MATIÈRES",
    "SOMMAIRE",
]

# Patterns pour nettoyer le contenu indésirable
CLEANUP_PATTERNS = [
    (r'^[IVXLCDM]+(?:\.\d+)*\.?\s+.+?\s+\d+\s*$', '', re.MULTILINE),
    (r'^\d+(?:\.\d+)*\.?\s+.+?\s+\d+\s*$', '', re.MULTILINE),
    (r'\[([^\]]+)\]\(#[^)]+\)', r'\1', 0),
    (r'!\[.*?\]\(data:image[^)]+\)', '', 0),
    (r'\n{4,}', '\n\n\n', 0),
]


# ==============================================================================
# EXTRACTION CODE RITM
# ==============================================================================

def extract_ritm(filename: str) -> Optional[str]:
    """Extrait le code RITM du nom de fichier."""
    match = re.match(RITM_PATTERN, filename, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    return None


def scan_docx_files(directory: Path) -> List[Tuple[Path, str]]:
    """
    Scanne un dossier pour trouver les fichiers DOCX avec un code RITM.
    Retourne une liste de (chemin, code_ritm).
    """
    results = []
    if not directory.exists():
        return results

    for docx_path in directory.glob("*.docx"):
        ritm = extract_ritm(docx_path.name)
        if ritm:
            results.append((docx_path, ritm))
        else:
            logger.warning(f"Pas de code RITM détecté: {docx_path.name}")

    return results


# ==============================================================================
# CONVERSION DOCX -> MARKDOWN
# ==============================================================================

def clean_html_toc(html: str) -> str:
    """Supprime les éléments de TOC du HTML."""
    html = re.sub(
        r'<p[^>]*>\s*<a\s+href="#_Toc[^"]*"[^>]*>.*?</a>\s*</p>',
        '',
        html,
        flags=re.DOTALL | re.IGNORECASE
    )

    html = re.sub(
        r'<a\s+href="#_Toc[^"]*"[^>]*>(.*?)</a>',
        r'\1',
        html,
        flags=re.DOTALL | re.IGNORECASE
    )

    def fix_toc_h1(match):
        content = match.group(1)
        toc_patterns = [
            r'Table\s+des\s+mati[eè]res',
            r'Sommaire',
            r'TABLE\s+DES\s+MATI[EÈ]RES',
            r'SOMMAIRE'
        ]
        for pattern in toc_patterns:
            if re.search(pattern, content, re.IGNORECASE):
                parts = re.split(r'<a\s+id="[^"]*"[^>]*>\s*</a>', content)
                if len(parts) > 1 and parts[-1].strip():
                    return f'<h1>{parts[-1].strip()}</h1>'
                return ''
        return match.group(0)

    html = re.sub(r'<h1[^>]*>(.*?)</h1>', fix_toc_h1, html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<a\s+id="_Toc[^"]*"[^>]*>\s*</a>', '', html, flags=re.IGNORECASE)

    return html


def docx_to_markdown(docx_path: Path) -> str:
    """Convertit un fichier DOCX en Markdown propre."""
    with open(docx_path, 'rb') as f:
        result = mammoth.convert_to_html(
            f,
            style_map=MAMMOTH_STYLE_MAP,
            include_embedded_style_map=False,
        )
        html_content = result.value

    html_content = clean_html_toc(html_content)

    h2t = html2text.HTML2Text()
    h2t.ignore_links = False
    h2t.ignore_images = True
    h2t.ignore_emphasis = False
    h2t.body_width = 0
    h2t.unicode_snob = True
    h2t.skip_internal_links = True

    markdown = h2t.handle(html_content)
    markdown = post_process_markdown(markdown)

    return markdown


def post_process_markdown(content: str) -> str:
    """Post-traite le Markdown pour supprimer préambule et TOC."""
    lines = content.split('\n')

    start_index = find_content_start(lines)
    if start_index > 0:
        lines = lines[start_index:]

    content = '\n'.join(lines)

    for pattern, replacement, flags in CLEANUP_PATTERNS:
        if flags:
            content = re.sub(pattern, replacement, content, flags=flags)
        else:
            content = re.sub(pattern, replacement, content)

    content = clean_tables(content)
    content = normalize_headings(content)
    content = final_cleanup(content)

    return content


def find_content_start(lines: List[str]) -> int:
    """Trouve l'index de la première ligne du vrai contenu."""

    def has_content_after(start_idx: int) -> bool:
        current_line = lines[start_idx].strip()
        current_level = current_line.count('#') if current_line.startswith('#') else 0
        content_lines = 0

        for j in range(start_idx + 1, min(start_idx + 20, len(lines))):
            line = lines[j].strip()
            if not line:
                continue

            if line.startswith('#'):
                line_level = line.count('#') if line.startswith('#') else 0
                if line_level > 0 and line_level <= current_level:
                    if content_lines == 0:
                        return False
                continue

            if re.match(r'^\d+\.\d*\s+[A-Z]', line) and len(line) < 50:
                continue

            if len(line) > 40:
                content_lines += 1
                if content_lines >= 1:
                    return True

        return content_lines >= 1

    toc_found = False
    toc_end_index = 0

    for i, line in enumerate(lines):
        line_stripped = line.strip()
        for marker in TOC_END_MARKERS:
            if marker.lower() in line_stripped.lower():
                toc_found = True
                toc_end_index = i
                break

    if toc_found:
        for i in range(toc_end_index + 1, len(lines)):
            line = lines[i].strip()
            if not line:
                continue
            if is_chapter_heading(line) and has_content_after(i):
                return i

    consecutive_titles = 0
    last_title_idx = -1

    for i, line in enumerate(lines):
        line_stripped = line.strip()
        if not line_stripped:
            continue

        is_title = line_stripped.startswith('#') or re.match(r'^\d+\.', line_stripped)

        if is_title:
            consecutive_titles += 1
            last_title_idx = i
        else:
            if consecutive_titles >= 5:
                for j in range(last_title_idx, len(lines)):
                    if lines[j].strip().startswith('#') and has_content_after(j):
                        return j
            consecutive_titles = 0

    for i, line in enumerate(lines):
        line_stripped = line.strip()
        if is_chapter_heading(line_stripped) and has_content_after(i):
            return i

    return 0


def is_chapter_heading(line: str) -> bool:
    """Vérifie si une ligne est un titre de chapitre principal."""
    match = re.match(r'^(#{1,3})\s+(.+)$', line)
    if not match:
        return False

    title_text = match.group(2).strip()

    if re.search(r'\s+\d+\s*$', title_text):
        return False

    title_text = re.sub(r'^\*\*(.+)\*\*$', r'\1', title_text)

    if not title_text or len(title_text) < 5:
        return False

    known_chapter_starts = [
        r'^Description\s+du\s+projet',
        r'^Introduction',
        r'^Contexte\s+(?:du|et)',
        r'^Pr[ée]sentation',
        r'^Objectifs?\s+(?:du|et)',
    ]

    for pattern in known_chapter_starts:
        if re.search(pattern, title_text, re.IGNORECASE):
            return True

    has_numbering = re.match(
        r'^[IVXLCDM]+\.?\s+[IVXLCDM]*\.?\s*\d*\.?\s*[A-ZÀ-Ý]',
        title_text
    ) or re.match(
        r'^\d+\.?\s+[A-ZÀ-Ý]',
        title_text
    )

    is_preliminary = re.match(r'^I\.\d+\.?\s+', title_text)

    return has_numbering and not is_preliminary


def clean_tables(content: str) -> str:
    """Nettoie et normalise les tableaux Markdown."""
    lines = content.split('\n')
    result = []
    in_table = False
    table_lines = []

    for line in lines:
        stripped = line.strip()

        is_table_line = False
        if '|' in stripped:
            if not stripped.startswith('#'):
                if re.match(r'^[-|\s:]+$', stripped) or '|' in stripped:
                    is_table_line = True

        if is_table_line:
            if not in_table:
                in_table = True
                table_lines = []
            table_lines.append(line)
        else:
            if in_table:
                processed_table = process_table(table_lines)
                result.extend(processed_table)
                result.append('')
                in_table = False
                table_lines = []
            result.append(line)

    if in_table and table_lines:
        processed_table = process_table(table_lines)
        result.extend(processed_table)

    return '\n'.join(result)


def process_table(table_lines: List[str]) -> List[str]:
    """Traite un tableau pour le normaliser au format Markdown standard."""
    if not table_lines:
        return []

    rows = []

    for line in table_lines:
        line = line.strip()
        if not line:
            continue

        if re.match(r'^[-|\s:]+$', line) and '-' in line:
            continue

        if line.startswith('|'):
            line = line[1:]
        if line.endswith('|'):
            line = line[:-1]

        cells = [c.strip() for c in line.split('|')]
        if cells:
            rows.append(cells)

    if not rows:
        return []

    max_cols = max(len(row) for row in rows)
    if max_cols == 0:
        return []

    for row in rows:
        while len(row) < max_cols:
            row.append('')

    result = []
    header = '| ' + ' | '.join(rows[0]) + ' |'
    result.append(header)
    separator = '| ' + ' | '.join(['---'] * max_cols) + ' |'
    result.append(separator)

    for row in rows[1:]:
        line = '| ' + ' | '.join(row) + ' |'
        result.append(line)

    return result


def normalize_headings(content: str) -> str:
    """Normalise les titres."""
    lines = content.split('\n')
    result = []

    for line in lines:
        match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if match:
            level = len(match.group(1))
            title_text = match.group(2).strip()
            title_text = re.sub(r'^\*\*(.+)\*\*$', r'\1', title_text)
            line = '#' * level + ' ' + title_text

        result.append(line)

    return '\n'.join(result)


def final_cleanup(content: str) -> str:
    """Nettoyage final du contenu."""
    content = re.sub(r'\n{3,}', '\n\n', content)
    content = re.sub(r' +$', '', content, flags=re.MULTILINE)
    content = content.lstrip('\n')
    content = content.rstrip() + '\n'

    lines = content.split('\n')
    cleaned_lines = []
    for line in lines:
        if re.match(r'^[-_=\s|]+$', line) and '|' not in line:
            continue
        cleaned_lines.append(line)

    return '\n'.join(cleaned_lines)


# ==============================================================================
# TRAITEMENT PARALLÈLE
# ==============================================================================

def process_single_file(args: Tuple[str, str, str, str, str]) -> Tuple[str, str, str, str, str, str, str]:
    """
    Traite un seul fichier DOCX -> Markdown.
    Retourne (mode, ritm, filename, src, out_path, status, error).
    """
    src_path, ritm, mode, out_dir, log_dir = args
    src_path = Path(src_path)
    out_dir = Path(out_dir)
    log_dir = Path(log_dir)

    try:
        md_content = docx_to_markdown(src_path)

        out_name = src_path.stem + ".md"
        out_path = out_dir / out_name
        out_path.write_text(md_content, encoding="utf-8")

        return (mode, ritm, src_path.name, str(src_path), str(out_path), "OK", "")

    except Exception as ex:
        err_msg = f"{type(ex).__name__}: {ex}"

        trace = traceback.format_exc()
        log_file = log_dir / (src_path.stem + f".{mode}.error.log")
        log_file.write_text(trace, encoding="utf-8")

        return (mode, ritm, src_path.name, str(src_path), "", "ERROR", err_msg)


# ==============================================================================
# PROGRAMME PRINCIPAL
# ==============================================================================

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convertit les DOCX en Markdown en scannant les dossiers EDB/NDC par code RITM."
    )
    parser.add_argument("--edb-dir", type=str, default=DEFAULT_EDB_DIR,
                        help=f"Dossier des EDB (défaut: {DEFAULT_EDB_DIR})")
    parser.add_argument("--ndc-dir", type=str, default=DEFAULT_NDC_DIR,
                        help=f"Dossier des NDC (défaut: {DEFAULT_NDC_DIR})")
    parser.add_argument("--output-dir", type=str, default=DEFAULT_OUTPUT_DIR,
                        help=f"Dossier de sortie (défaut: {DEFAULT_OUTPUT_DIR})")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS,
                        help="Nombre de workers (défaut: 0 = auto)")
    args = parser.parse_args()

    edb_dir = Path(args.edb_dir).resolve()
    ndc_dir = Path(args.ndc_dir).resolve()
    output_dir = Path(args.output_dir).resolve()

    out_edb = output_dir / "edb"
    out_ndc = output_dir / "ndc"
    log_dir = output_dir / LOG_DIRNAME

    for d in [out_edb, out_ndc, log_dir]:
        d.mkdir(parents=True, exist_ok=True)

    # Scanner les dossiers
    logger.info(f"Scan EDB: {edb_dir}")
    edb_files = scan_docx_files(edb_dir)
    logger.info(f"  -> {len(edb_files)} fichiers avec code RITM")

    logger.info(f"Scan NDC: {ndc_dir}")
    ndc_files = scan_docx_files(ndc_dir)
    logger.info(f"  -> {len(ndc_files)} fichiers avec code RITM")

    if not edb_files and not ndc_files:
        logger.info("Aucun fichier à traiter.")
        return 0

    # Préparer les tâches
    tasks = []
    for path, ritm in edb_files:
        tasks.append((str(path), ritm, "edb", str(out_edb), str(log_dir)))
    for path, ritm in ndc_files:
        tasks.append((str(path), ritm, "ndc", str(out_ndc), str(log_dir)))

    total = len(tasks)
    workers = args.workers if args.workers > 0 else os.cpu_count()
    logger.info(f"Traitement de {total} fichiers avec {workers} workers...")

    stats = {"ok": 0, "error": 0}
    results = []

    # Traitement parallèle
    with ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(process_single_file, task): task for task in tasks}

        for i, future in enumerate(as_completed(futures), 1):
            result = future.result()
            mode, ritm, filename, src, out_path, status, error = result
            results.append(result)

            if status == "OK":
                stats["ok"] += 1
                logger.info(f"[{i}/{total}] OK ({mode}) {ritm} - {filename}")
            else:
                stats["error"] += 1
                logger.error(f"[{i}/{total}] ERROR ({mode}) {ritm} - {filename}: {error}")

    # Générer le rapport Excel
    report_rows = []
    for mode, ritm, filename, src, out_path, status, error in results:
        report_rows.append({
            "Type": mode.upper(),
            "Code RITM": ritm,
            "Fichier source": filename,
            "Chemin source": src,
            "Fichier Markdown": out_path,
            "Statut": status,
            "Erreur": error,
        })

    report_path = Path.cwd() / "extract_report.xlsx"
    df = pd.DataFrame(report_rows)
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extraction Markdown")

    # Afficher les codes RITM trouvés
    edb_ritms = set(ritm for _, ritm in edb_files)
    ndc_ritms = set(ritm for _, ritm in ndc_files)
    common_ritms = edb_ritms & ndc_ritms

    # Résumé
    logger.info("")
    logger.info("=== Résumé ===")
    logger.info(f"OK: {stats['ok']}")
    logger.info(f"Erreurs: {stats['error']}")
    logger.info(f"Codes RITM EDB: {len(edb_ritms)}")
    logger.info(f"Codes RITM NDC: {len(ndc_ritms)}")
    logger.info(f"Codes RITM communs: {len(common_ritms)}")
    logger.info(f"Sortie: {output_dir}")
    logger.info(f"Rapport: {report_path}")

    return 0 if stats["error"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
