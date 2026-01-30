#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
dedupe.py
- Parcourt le dossier source (par défaut ./clean_extension) qui ne contient que pdf/doc/docx
- Dédoublonne avec la règle : à nom de base identique → Word prioritaire (docx > doc) sur PDF
- Si plusieurs fichiers du même type existent, conserve le plus récent (mtime)
- Génère un Excel (avant les copies) listant tous les fichiers, l'Action (conserver/ignorer) et la Raison
- Copie ensuite les "conserver" vers un dossier 'dedupe' situé au MÊME NIVEAU que le dossier source

Usage :
    python dedupe.py
    python dedupe.py --source ./clean_extension --report dedupe_report.xlsx

Options :
    --source   : dossier source (défaut: ./clean_extension)
    --report   : nom du fichier Excel de rapport (défaut: dedupe_report.xlsx) écrit dans le dossier courant
    --dry-run  : n’effectue pas les copies, génère uniquement le rapport

Dépendances :
    - pandas
    - openpyxl
"""

import argparse
from pathlib import Path
from datetime import datetime
import shutil
import sys
import re
import pandas as pd

DEFAULT_SOURCE = "clean_extension"
DEFAULT_REPORT = "dedupe_report.xlsx"
ALLOW = {".pdf", ".doc", ".docx"}

# Suffixe anti-collision de clean_extension.py : _YYYYMMDD_HHMMSS
TS_SUFFIX_RE = re.compile(r"_(\d{8}_\d{6})$")  # appliqué au stem (sans extension)


def normalized_key(p: Path) -> str:
    """Clé de regroupement = stem en minuscules, trim, SANS suffixe horodaté."""
    stem = p.stem.strip().lower()
    stem = TS_SUFFIX_RE.sub("", stem)
    return stem


def pick_most_recent(paths: list[Path]) -> Path:
    """Retourne le fichier le plus récent (mtime) parmi la liste."""
    return max(paths, key=lambda x: x.stat().st_mtime)


def safe_copy(src: Path, dest_dir: Path) -> Path:
    """
    Copie src -> dest_dir en évitant l'écrasement (suffixe horodaté si collision).
    Retourne le chemin final créé.
    """
    dest_dir.mkdir(parents=True, exist_ok=True)
    target = dest_dir / src.name
    if target.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = target.with_name(f"{target.stem}_{ts}{target.suffix}")
    shutil.copy2(src, target)
    return target


def main():
    parser = argparse.ArgumentParser(
        description="Dédoublonne (Word > PDF) et génère un Excel avant de copier les fichiers dans un dossier 'dedupe' au même niveau."
    )
    parser.add_argument("--source", type=str, default=DEFAULT_SOURCE,
                        help="Dossier source (défaut: ./clean_extension)")
    parser.add_argument("--report", type=str, default=DEFAULT_REPORT,
                        help="Nom du fichier Excel de rapport (défaut: dedupe_report.xlsx)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Génère le rapport sans copier")
    args = parser.parse_args()

    source_dir = Path(args.source).resolve()
    if not source_dir.exists() or not source_dir.is_dir():
        print(f"❌ Dossier source invalide : {source_dir}")
        sys.exit(1)

    # Dossier 'dedupe' AU MÊME NIVEAU que source (frère de clean_extension et raw)
    dedupe_dir = (source_dir.parent / "dedupe").resolve()
    # On ne le crée pas tout de suite pour être fidèle à 'Excel avant copier'

    # Fichiers du 1er niveau uniquement, filtrés sur les extensions attendues
    files = [p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() in ALLOW]

    if not files:
        print("ℹ️ Aucun fichier pdf/doc/docx trouvé.")
        sys.exit(0)

    # Regroupement par clé normalisée
    groups: dict[str, list[Path]] = {}
    for f in files:
        groups.setdefault(normalized_key(f), []).append(f)

    rows = []
    keep_set: set[Path] = set()

    for key, paths in sorted(groups.items()):
        # Partitionner par extension
        docx_list = [p for p in paths if p.suffix.lower() == ".docx"]
        doc_list  = [p for p in paths if p.suffix.lower() == ".doc"]
        pdf_list  = [p for p in paths if p.suffix.lower() == ".pdf"]

        chosen = None
        rule_reason = ""

        if docx_list:
            chosen = pick_most_recent(docx_list)
            rule_reason = "DOCX prioritaire (Word > PDF). Autres ignorés."
        elif doc_list:
            chosen = pick_most_recent(doc_list)
            rule_reason = "DOC conservé (pas de DOCX). PDF ignorés."
        elif pdf_list:
            chosen = pick_most_recent(pdf_list)
            rule_reason = "PDF seul → conservé (aucun Word)."

        if chosen is None:
            # Ne devrait pas arriver (filtrage ALLOW)
            continue

        keep_set.add(chosen)

        # Prépare lignes Excel (Action/Raison par fichier)
        for p in sorted(paths):
            ext = p.suffix.lower().lstrip(".")
            if p == chosen:
                action = "conserver"
                # Raison spécifique si plusieurs du même type
                same_type = [x for x in paths if x.suffix.lower() == p.suffix.lower()]
                if len(same_type) > 1:
                    reason = f"Conservé (plus récent parmi les {p.suffix.lower()})"
                else:
                    reason = rule_reason
                planned_dest = str((dedupe_dir / p.name).resolve())
            else:
                action = "ignorer"
                # Raison d'ignoré
                if p.suffix.lower() == ".pdf" and (docx_list or doc_list):
                    reason = "PDF ignoré (Word présent)"
                elif p.suffix.lower() == ".doc" and docx_list:
                    reason = "DOC ignoré (DOCX présent)"
                else:
                    # même extension que le choisi → moins récent
                    reason = "Ignoré (moins récent que celui conservé)"
                planned_dest = ""

            rows.append({
                "Nom du fichier": p.name,
                "Extension": ext,
                "Groupe (stem normalisé)": key,
                "Action": action,
                "Raison": reason,
                "Chemin source": str(p),
                "Chemin destination (prévu)": planned_dest,
            })

    # 1) Écrire le rapport Excel AVANT COPIE
    report_path = (Path.cwd() / args.report).resolve()
    df = pd.DataFrame(rows, columns=[
        "Nom du fichier", "Extension", "Groupe (stem normalisé)", "Action",
        "Raison", "Chemin source", "Chemin destination (prévu)"
    ])
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dédoublonnage")

    print("🗒️  Rapport Excel généré (avant copie) :", report_path)

    # 2) Effectuer les copies des fichiers à conserver
    if not args.dry_run:
        dedupe_dir.mkdir(parents=True, exist_ok=True)
        copied = 0
        for p in sorted(keep_set):
            _ = safe_copy(p, dedupe_dir)
            copied += 1
        print(f"✅ Copie terminée dans : {dedupe_dir} (fichiers copiés : {copied})")
    else:
        print("🔎 Mode --dry-run : aucune copie effectuée.")

    print("✔️  Dédoublonnage terminé.")

if __name__ == "__main__":
    main()
