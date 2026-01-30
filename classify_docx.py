#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Classification des DOCX (première page) en EDB / NDC / AUTRES.

Règles :
- EDB si "expression de besoin" est présent (insensible casse/accents)
- NDC si motif "CAPS_YYYY-NNN" (e.g. CAPS_2020-132) détecté
- AUTRES sinon

Copie des fichiers vers :
  classified_docx/
    ├─ edb/
    ├─ ndc/
    └─ autres/

Génère un Excel récapitulatif : classified_docx/classify_report.xlsx

Limitations/choix techniques :
- La "première page" est considérée comme le contenu jusqu'au **premier saut de page explicite**
  s'il existe. À défaut, on tronque aux N premiers caractères (paramétrable).
"""

import argparse
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
import unicodedata

import pandas as pd
from docx import Document
from docx.oxml.ns import qn

# ---------- Configuration par défaut ----------
DEFAULT_INPUT_DIR = "docx"
DEFAULT_OUTPUT_DIR = "classified_docx"  # <— renommé
DEFAULT_ON_EXISTS = "skip"              # skip | overwrite | suffix
DEFAULT_FIRST_PAGE_CHAR_LIMIT = 5000


# ---------- Utilitaires ----------
def strip_accents(s: str) -> str:
    """Supprime les accents pour une comparaison diacritics-insensitive."""
    if s is None:
        return ""
    return "".join(c for c in unicodedata.normalize("NFD", s) if not unicodedata.combining(c))


def paragraph_has_page_break(p) -> bool:
    """
    Détecte un saut de page explicite (<w:br w:type="page"/>) dans un paragraphe DOCX.
    """
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
EDB_TOKEN = "expression de besoin"
NDC_REGEX = re.compile(r"\bCAPS_\d{4}-\d+\b")  # e.g. CAPS_2020-132 (client CAPS)

def classify_first_page(text: str) -> tuple[str, str]:
    """
    Retourne (classe, raison).
    Priorité NDC > EDB, puis AUTRES.
    - EDB: recherche insensible aux accents/majuscules.
    - NDC: motif CAPS_YYYY-NNN (sensible au motif, non accentué).
    """
    # Pour EDB : normalisation casse + accents
    norm_text = strip_accents(text).lower()
    is_edb = EDB_TOKEN in norm_text  # token est déjà en minuscules sans accents

    # Pour NDC : motif structuré
    ndc_match = NDC_REGEX.search(text)

    if ndc_match:
        return "NDC", f"pattern:{ndc_match.group(0)}"
    if is_edb:
        return "EDB", f"contains:'{EDB_TOKEN}'"
    return "AUTRES", ""


# ---------- Programme principal ----------
def main():
    parser = argparse.ArgumentParser(description="Classement DOCX en EDB / NDC / AUTRES (première page).")
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
    if args.recursive:
        candidates = list(in_dir.rglob("*.docx"))
    else:
        candidates = list(in_dir.glob("*.docx"))

    records = []
    total = len(candidates)
    print(f"[INFO] {total} fichier(s) .docx à traiter dans: {in_dir}")

    for i, path in enumerate(sorted(candidates), start=1):
        rel = path.relative_to(in_dir) if hasattr(path, "relative_to") and str(in_dir) in str(path) else path.name
        print(f"[{i}/{total}] Traitement: {rel}")
        classification = "ERREUR"
        reason = ""
        dest_path = None
        copy_status = "not_copied"

        try:
            first_page = extract_first_page_text(path, char_limit=args.first_page_char_limit)
            classification, reason = classify_first_page(first_page)
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

    # Rapport Excel
    df = pd.DataFrame.from_records(records,
                                   columns=["filename", "original_path", "classification",
                                            "reason", "destination_path", "copy_status"])
    report_path = "classify_report.xlsx"
    df.to_excel(report_path, index=False)
    print(f"[OK] Rapport écrit : {report_path}")
    print("[OK] Terminé.")


if __name__ == "__main__":
    main()
