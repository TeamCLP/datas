#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
clean_extension.py
- Parcourt le sous-dossier ./raw du dossier courant (ou autre via --raw)
- Copie tous les fichiers .pdf/.doc/.docx dans le dossier frère 'clean_extension' situé à côté de 'raw'
- Génère un Excel listant tous les fichiers présents dans ./raw :
    Col A : Nom du fichier
    Col B : Extension
    Col C : Action ("conserver" ou "ignorer")

Usage :
    python clean_extension.py
    python clean_extension.py --raw ./raw --out-name inventaire_raw.xlsx

Dépendances :
    - pandas
    - openpyxl
"""

import argparse
from pathlib import Path
from datetime import datetime
import shutil
import pandas as pd

ALLOWED_EXT = {".pdf", ".doc", ".docx"}  # insensible à la casse
DEFAULT_RAW_REL = "raw"
DEFAULT_REPORT_NAME = "inventaire_raw.xlsx"

def main():
    parser = argparse.ArgumentParser(
        description="Copie PDF/DOC/DOCX de ./raw vers le dossier frère clean_extension et génère un Excel d'inventaire."
    )
    parser.add_argument("--raw", type=str, default=DEFAULT_RAW_REL,
                        help="Chemin du sous-dossier raw (défaut: ./raw)")
    parser.add_argument("--out-name", type=str, default=DEFAULT_REPORT_NAME,
                        help="Nom du fichier Excel de sortie (défaut: inventaire_raw.xlsx)")
    args = parser.parse_args()

    cwd = Path.cwd()
    raw_dir = (cwd / args.raw).resolve()

    if not raw_dir.exists() or not raw_dir.is_dir():
        print(f"❌ Le dossier source n'existe pas ou n'est pas un dossier: {raw_dir}")
        return

    # 👉 Dossier cible : à CÔTÉ de 'raw'
    # Ex : /chemin/projet/raw  -> cible : /chemin/projet/clean_extension
    target_dir = (raw_dir.parent / "clean_extension").resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    inventory_rows = []
    copied_count = 0
    ignored_count = 0

    for entry in raw_dir.iterdir():
        # Ignorer les sous-dossiers
        if not entry.is_file():
            continue

        ext = entry.suffix.lower()
        filename = entry.name

        if ext in ALLOWED_EXT:
            # Copie vers clean_extension (à côté de raw)
            dest_path = target_dir / filename
            # Éviter l'écrasement si un fichier homonyme existe déjà
            if dest_path.exists():
                stem = dest_path.stem
                new_name = f"{stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{dest_path.suffix}"
                dest_path = dest_path.with_name(new_name)

            shutil.copy2(entry, dest_path)
            action = "conserver"
            copied_count += 1
        else:
            action = "ignorer"
            ignored_count += 1

        inventory_rows.append({
            "Nom du fichier": filename,
            "Extension": ext[1:] if ext.startswith(".") else ext,
            "Action": action,
        })

    # Génération de l'Excel dans le dossier courant (où vous lancez le script)
    report_path = (cwd / args.out_name).resolve()
    df = pd.DataFrame(inventory_rows, columns=["Nom du fichier", "Extension", "Action"])
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Inventaire", index=False)

    print("✅ Traitement terminé.")
    print(f"   - Dossier source         : {raw_dir}")
    print(f"   - Dossier de destination : {target_dir}")
    print(f"   - Rapport Excel          : {report_path}")
    print(f"   - Fichiers conservés (copiés) : {copied_count}")
    print(f"   - Fichiers ignorés            : {ignored_count}")

if __name__ == "__main__":
    main()
