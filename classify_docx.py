#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Classification des DOCX (première page + nom de fichier) en EDB / NDC / AUTRES.

Règles :
- EDB si l'une des expressions (insensible casse/accents) est présente dans la première page :
    • "expression de besoin"
    • "expression de besoins"
    • "expressions de besoins"
- NDC si un motif code client est détecté (prioritaire sur EDB), au choix :
    • dans la première page
    • dans le nom du fichier (avec extension)
  Motif accepté : CAPS[-_]YYYY[-_]NNN (ex : CAPS_2020-132, CAPS-2020_132, CAPS_2020_132)
  (client fixé à "CAPS", année = 4 chiffres, numéro = >= 1 chiffre)

- AUTRES sinon

Copies :
- Les fichiers sont copiés selon la classe dans :
    classified_docx/
      ├─ edb/
      ├─ ndc/
      └─ autres/

Traçabilité :
- Un rapport Excel est généré dans le dossier d'entrée (ex : docx/classify_report.xlsx).

CLI :
- --docx-dir : dossier d'entrée (défaut: docx)
- --output-dir : dossier racine de sortie (défaut: classified_docx)
- --on-exists : skip | overwrite | suffix (défaut: skip)
- --recursive : parcourir récursivement le dossier d'entrée
- --first-page-char-limit : troncature si pas de saut de page explicite (défaut: 5000)
"""

import argparse
from datetime import datetime
from pathlib import Path
import re
import shutil
import unicodedata

import pandas as pd
from docx import Document
from docx.oxml.ns import qn


# ---------- Configuration par défaut ----------
DEFAULT_INPUT_DIR = "docx"
DEFAULT_OUTPUT_DIR = "classified_docx"
DEFAULT_ON_EXISTS = "skip"      # skip | overwrite | suffix
DEFAULT_FIRST_PAGE_CHAR_LIMIT = 5000


# ---------- Utilitaires ----------
def strip_accents(s: str) -> str:
    """Supprime les accents pour comparaison diacritics-insensitive."""
    if s is None:
        return ""
    return "".join(c for c in unicodedata.normalize("NFD", s) if not unicodedata.combining(c))


def paragraph_has_page_break(p) -> bool:
    """Détecte un saut de page explicite (<w:br w:type="page"/>) dans un paragraphe DOCX."""
    for br in p._element.xpath(".//w:br"):
        br_type = br.get(qn("w:type"))
        if (br_type or "").lower() == "page":
            return True
    return False


def extract_first_page_text(docx_path: Path, char_limit: int = DEFAULT_FIRST_PAGE_CHAR_LIMIT) -> str:
    """
    Extrait le texte de la première page en s'arrêtant au premier saut de page explicite.
    Si aucun saut n'est trouvé, tronque à `char_limit` caractères.
    """
    doc = Document(str(docx_path))
    chunks = []
    total_len = 0
    for p in doc.paragraphs:
        txt = p.text or ""
        chunks.append(txt)
        total_len += len(txt) + 1  # +1 pour le saut de ligne
        if paragraph_has_page_break(p):
            break
        if total_len >= char_limit:
            break
    return "\n".join(chunks)


def ensure_dirs(base_out: Path):
    """Crée la hiérarchie de sortie : base_out/{edb,ndc,autres}."""
    (base_out / "edb").mkdir(parents=True, exist_ok=True)
    (base_out / "ndc").mkdir(parents=True, exist_ok=True)
    (base_out / "autres").mkdir(parents=True, exist_ok=True)


def safe_copy(src: Path, dst_dir: Path, on_exists: str):
    """
    Copie src vers dst_dir en respectant la politique on_exists: skip|overwrite|suffix.
    Retourne (dst_path_effectif, status).
    """
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / src.name
    if dst.exists():
        if on_exists == "skip":
            return dst, "skipped_existing"
        elif on_exists == "overwrite":
            shutil.copy2(src, dst)
            return dst, "overwritten"
        elif on_exists == "suffix":
            stem, ext = src.stem, src.suffix
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            suffixed = dst_dir / f"{stem}_{ts}{ext}"
            shutil.copy2(src, suffixed)
            return suffixed, "copied_with_suffix"
        else:
            raise ValueError(f"on_exists invalide: {on_exists}")
    else:
        shutil.copy2(src, dst)
        return dst, "copied"


# ---------- Détection métier ----------
# Variantes EDB (insensibles aux accents/majuscules)
EDB_TOKENS = [
    "expression de besoin",
    "expression de besoins",
    "expressions de besoins",
]

# NDC : autoriser '-' ou '_' entre les segments
# CAPS[-_]YYYY[-_]NNN  (insensible à la casse)
NDC_REGEX = re.compile(r"\bCAPS[-_]\d{4}[-_]\d+\b", flags=re.IGNORECASE)


def detect_edb(text: str) -> tuple[bool, str]:
    """Détection EDB dans la première page (insensible accents/casse)."""
    norm_text = strip_accents(text).lower()
    for token in EDB_TOKENS:
        if token in norm_text:
            return True, f"contains:'{token}'"
    return False, ""


def detect_ndc_in_text_or_filename(text: str, filename: str) -> tuple[bool, str]:
    """
    Détection NDC dans la première page OU le nom de fichier.
    Renvoie (is_ndc, reason) avec la source et le motif trouvé.
    """
    # Cherche dans le texte de la première page (cas-insensitive via regex)
    m_text = NDC_REGEX.search(text or "")
    if m_text:
        return True, f"pattern:{m_text.group(0)} source:text"

    # Cherche dans le nom de fichier (incl. extension)
    m_name = NDC_REGEX.search(filename or "")
    if m_name:
        return True, f"pattern:{m_name.group(0)} source:filename"

    return False, ""


def classify(first_page_text: str, filename: str) -> tuple[str, str]:
    """
    Classement avec priorité : NDC > EDB > AUTRES.
    Retourne (classe, raison).
    """
    # 1) NDC prioritaire – texte ou nom du fichier
    is_ndc, reason_ndc = detect_ndc_in_text_or_filename(first_page_text, filename)
    if is_ndc:
        return "NDC", reason_ndc

    # 2) EDB sur la première page (accents/casse insensibles)
    is_edb, reason_edb = detect_edb(first_page_text)
    if is_edb:
        return "EDB", reason_edb

    # 3) Sinon autres
    return "AUTRES", ""


# ---------- Programme principal ----------
def main():
    parser = argparse.ArgumentParser(description="Classement DOCX en EDB / NDC / AUTRES (1ère page + nom de fichier).")
    parser.add_argument("--docx-dir", default=DEFAULT_INPUT_DIR,
                        help="Dossier d'entrée contenant les .docx (défaut: docx)")
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR,
                        help="Dossier racine de sortie (défaut: classified_docx)")
    parser.add_argument("--on-exists", choices=["skip", "overwrite", "suffix"], default=DEFAULT_ON_EXISTS,
                        help="Politique en cas de collision de nom (défaut: skip)")
    parser.add_argument("--recursive", action="store_true",
                        help="Parcourir récursivement le dossier d'entrée")
    parser.add_argument("--first-page-char-limit", type=int, default=DEFAULT_FIRST_PAGE_CHAR_LIMIT,
                        help="Troncature si pas de saut de page explicite (défaut: 5000)")
    args = parser.parse_args()

    in_dir = Path(args.docx_dir).resolve()
    base_out = Path(args.output_dir).resolve()
    ensure_dirs(base_out)

    # Collecte des fichiers
    candidates = list(in_dir.rglob("*.docx")) if args.recursive else list(in_dir.glob("*.docx"))

    records = []
    total = len(candidates)
    print(f"[INFO] {total} fichier(s) .docx à traiter dans: {in_dir}")

    for i, path in enumerate(sorted(candidates, key=lambda p: str(p).lower()), start=1):
        try:
            rel = path.relative_to(in_dir)
        except Exception:
            rel = path.name

        print(f"[{i}/{total}] Traitement: {rel}")
        classification = "ERREUR"
        reason = ""
        dest_path = None
        copy_status = "not_copied"

        try:
            first_page = extract_first_page_text(path, char_limit=args.first_page_char_limit)
            classification, reason = classify(first_page, path.name)

            # Dossier cible selon la classe
            if classification == "EDB":
                target_dir = base_out / "edb"
            elif classification == "NDC":
                target_dir = base_out / "ndc"
            elif classification == "AUTRES":
                target_dir = base_out / "autres"
            else:
                target_dir = None  # ERREUR

            if target_dir is not None:
                dest_path, copy_status = safe_copy(path, target_dir, args.on_exists)

        except Exception as e:
            classification = "ERREUR"
            reason = f"exception:{type(e).__name__}: {e}"

        records.append({
            "filename": path.name,
            "original_path": str(path),
            "classification": classification,
            "reason": reason,
            "destination_path": "" if dest_path is None else str(dest_path),
            "copy_status": copy_status,
        })

    # Rapport Excel -> dans le dossier d'entrée (ex: docx/classify_report.xlsx)
    df = pd.DataFrame.from_records(
        records,
        columns=["filename", "original_path", "classification", "reason", "destination_path", "copy_status"],
    )
    report_path = in_dir / "classify_report.xlsx"
    df.to_excel(report_path, index=False)
    print(f"[OK] Rapport écrit : {report_path}")
    print("[OK] Terminé.")


if __name__ == "__main__":
    main()
