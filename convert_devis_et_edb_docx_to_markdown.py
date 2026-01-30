#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversion DOCX -> Markdown (dataset pour training LLM)
- NDC (colonne G) + EDB (colonne F)
- Filtre Excel configurable
- Utilise Mammoth pour conversion DOCX (ignore headers/footers automatiquement)
- Ignore page de garde, synthèse, table des matières
- Préserve titres, paragraphes, listes, tableaux
- Format Markdown homogène pour training

Dépendances:
  pip install pandas openpyxl mammoth html2text
"""

from __future__ import annotations

import re
import logging
import traceback
from pathlib import Path
from typing import List, Tuple

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
# CONFIGURATION - MODIFIEZ ICI
# ==============================================================================

# Fichier Excel source
EXCEL_NAME = "couverture_EDB_NDC_par_RITM.xlsx"

# Colonnes contenant les chemins des fichiers (index 0-based: A=0, B=1, etc.)
COL_EDB = 5  # Colonne F = index 5
COL_NDC = 6  # Colonne G = index 6

# ------------------------------
# FILTRES EXCEL
# ------------------------------
EXCEL_FILTERS = [
    (1, None),      # Colonne B (mettre None pour désactiver)
    (2, None),      # Colonne C (mettre None pour désactiver)
    (3, "OUI"),     # Colonne D = "OUI"
    (4, None),      # Colonne E (mettre None pour désactiver)
]

# ------------------------------
# Dossiers de sortie
# ------------------------------
OUTPUT_DIRNAME = "dataset_markdown"
LOG_DIRNAME = "_logs"
SUBDIR_NDC = "ndc"
SUBDIR_EDB = "edb"

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

# Patterns pour détecter le début du vrai contenu
CHAPTER_START_PATTERNS = [
    r'^#{1,2}\s+[IVXLCDM]+\.\s+',
    r'^#{1,2}\s+[IVXLCDM]+\s+',
    r'^#{1,2}\s+\d+\.\s+',
    r'^#{1,2}\s+\d+\s+',
    r'^#\s+Description\s+du\s+projet',
    r'^#\s+Introduction',
    r'^#\s+Contexte',
    r'^#\s+Périmètre',
    r'^#\s+Perimetre',
]

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


# ------------------------------
# Conversion DOCX -> Markdown
# ------------------------------
def clean_html_toc(html: str) -> str:
    """Supprime les éléments de TOC du HTML."""
    # Supprimer les paragraphes contenant des liens TOC
    html = re.sub(
        r'<p[^>]*>\s*<a\s+href="#_Toc[^"]*"[^>]*>.*?</a>\s*</p>',
        '',
        html,
        flags=re.DOTALL | re.IGNORECASE
    )

    # Supprimer les liens TOC restants
    html = re.sub(
        r'<a\s+href="#_Toc[^"]*"[^>]*>(.*?)</a>',
        r'\1',
        html,
        flags=re.DOTALL | re.IGNORECASE
    )

    # Nettoyer les H1 qui contiennent "Table des matières"
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

    # Supprimer les ancres TOC restantes
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

    # Chercher après un marqueur de TOC explicite
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

    # Détecter la fin de la TOC
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

    # Fallback
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


# ------------------------------
# Chargement Excel
# ------------------------------
def load_targets_from_excel(excel_path: Path) -> Tuple[List[str], List[str]]:
    """Charge les fichiers à traiter depuis Excel en appliquant les filtres configurés."""
    df = pd.read_excel(excel_path, engine="openpyxl")

    mask = pd.Series([True] * len(df))

    for col_idx, expected_value in EXCEL_FILTERS:
        if expected_value is None:
            continue

        col_data = df.iloc[:, col_idx]

        if isinstance(expected_value, str):
            col_normalized = col_data.astype(str).str.strip().str.upper()
            expected_normalized = expected_value.strip().upper()
            mask = mask & (col_normalized == expected_normalized)
        else:
            mask = mask & (col_data == expected_value)

    edb_col = df.iloc[:, COL_EDB]
    ndc_col = df.iloc[:, COL_NDC]

    edb = edb_col[mask].dropna().astype(str).tolist()
    ndc = ndc_col[mask].dropna().astype(str).tolist()

    def clean(lst):
        result = []
        for cell in lst:
            cell = cell.strip()
            if not cell:
                continue

            parts = cell.split('|')
            for part in parts:
                part = part.strip().strip('"').strip("'")
                if part and part.lower() not in ('nan', 'none', ''):
                    result.append(part)

        return result

    active_filters = [(col, val) for col, val in EXCEL_FILTERS if val is not None]
    if active_filters:
        filter_desc = ", ".join([f"col{col}={val}" for col, val in active_filters])
        logger.info(f"Filtres appliqués: {filter_desc}")
    else:
        logger.info("Aucun filtre appliqué (tous les fichiers seront traités)")

    return clean(ndc), clean(edb)


# ------------------------------
# Programme principal
# ------------------------------
def main() -> int:
    cwd = Path(".").resolve()
    excel_path = cwd / EXCEL_NAME

    base_out = cwd / OUTPUT_DIRNAME
    log_dir = base_out / LOG_DIRNAME
    out_ndc = base_out / SUBDIR_NDC
    out_edb = base_out / SUBDIR_EDB

    for d in [base_out, log_dir, out_ndc, out_edb]:
        d.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        logger.error(f"Fichier Excel introuvable: {excel_path}")
        return 2

    ndc_list, edb_list = load_targets_from_excel(excel_path)

    logger.info(f"Fichiers NDC: {len(ndc_list)}")
    logger.info(f"Fichiers EDB: {len(edb_list)}")

    if not ndc_list and not edb_list:
        logger.info("Aucun fichier à traiter.")
        return 0

    report_rows = []
    stats = {"ok": 0, "error": 0, "missing": 0}

    def process_file(name: str, mode: str):
        src = cwd / name
        ext = src.suffix.lower()

        # S'assurer que c'est un .docx
        if not ext:
            src = src.with_suffix(".docx")
        elif ext != ".docx":
            # Chercher la version .docx
            src = src.with_suffix(".docx")

        out_dir = out_ndc if mode == "ndc" else out_edb

        if not src.exists():
            stats["missing"] += 1
            report_rows.append((mode, name, str(src), "", "MISSING", "Fichier introuvable"))
            logger.warning(f"Introuvable: {src}")
            return

        try:
            md_content = docx_to_markdown(src)

            out_name = Path(name).stem + ".md"
            out_path = out_dir / out_name
            out_path.write_text(md_content, encoding="utf-8")

            stats["ok"] += 1
            report_rows.append((mode, name, str(src), str(out_path), "OK", ""))
            logger.info(f"[OK] ({mode}) {src.name} -> {out_path.name}")

        except Exception as ex:
            stats["error"] += 1
            err_msg = f"{type(ex).__name__}: {ex}"
            report_rows.append((mode, name, str(src), "", "ERROR", err_msg))

            trace = traceback.format_exc()
            log_file = log_dir / (Path(name).stem + f".{mode}.error.log")
            log_file.write_text(trace, encoding="utf-8")

            logger.error(f"[ERREUR] ({mode}) {src.name}: {err_msg}")

    for f in ndc_list:
        process_file(f, mode="ndc")

    for f in edb_list:
        process_file(f, mode="edb")

    report_df = pd.DataFrame(
        report_rows,
        columns=["type", "source_excel", "input_path", "output_md", "status", "error"]
    )
    report_path = base_out / "conversion_report.csv"
    report_df.to_csv(report_path, index=False, encoding="utf-8")

    logger.info("")
    logger.info("=== Résumé ===")
    logger.info(f"OK: {stats['ok']}")
    logger.info(f"Manquants: {stats['missing']}")
    logger.info(f"Erreurs: {stats['error']}")
    logger.info(f"Rapport: {report_path}")

    return 0 if stats["error"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
