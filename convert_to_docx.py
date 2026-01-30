#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_to_docx.py
- Convertit tous les .doc présents dans ./dedupe en .docx via LibreOffice (soffice)
- Écrit les .docx dans un dossier 'docx' situé au MÊME NIVEAU que 'dedupe'

Exemples :
    python convert_to_docx.py
    python convert_to_docx.py --source ./dedupe --dest ./docx --overwrite
    python convert_to_docx.py --soffice /usr/bin/soffice

Prérequis :
    - LibreOffice installé (commande 'soffice' accessible)
"""

import argparse
import subprocess
import shutil
from pathlib import Path
from datetime import datetime
import sys

DEFAULT_SOURCE = "dedupe"

def find_soffice(user_path: str | None) -> Path | None:
    """Trouve le binaire soffice (via --soffice, sinon PATH)."""
    if user_path:
        p = Path(user_path)
        return p if p.exists() and p.is_file() else None
    which = shutil.which("soffice")
    return Path(which) if which else None

def convert_one_doc(soffice: Path, src_doc: Path, out_dir: Path, overwrite: bool) -> tuple[bool, str]:
    """
    Convertit DOC -> DOCX via soffice.
    - Si overwrite=False et que le .docx attendu existe, on suffixe avec un timestamp.
    Retourne (success:bool, message:str).
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    expected_docx = out_dir / (src_doc.stem + ".docx")

    # Commande LibreOffice : conversion directe vers out_dir
    cmd = [
        str(soffice),
        "--headless", "--nologo", "--nodefault", "--invisible",
        "--convert-to", "docx",
        "--outdir", str(out_dir),
        str(src_doc)
    ]

    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            check=False
        )
        output = proc.stdout.strip()

        produced = expected_docx
        if not produced.exists():
            # Rien n'a été produit → échec
            return (False, f"ÉCHEC: aucun .docx produit | sortie: {output}")

        # Gestion collision si le docx existait déjà et overwrite=False
        if expected_docx.exists() and not overwrite:
            # Le convertisseur vient d'écrire sur expected_docx si overwrite possible,
            # mais on renomme pour éviter l'écrasement logique
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            final_target = expected_docx.with_name(f"{expected_docx.stem}_{ts}{expected_docx.suffix}")
            try:
                produced.rename(final_target)
                return (True, f"OK (renommé pour éviter collision) -> {final_target.name}")
            except Exception as e:
                return (False, f"ERREUR renommage : {e}")
        else:
            return (True, f"OK -> {expected_docx.name} | {output.splitlines()[-1] if output else 'ok'}")

    except FileNotFoundError:
        return (False, "ERREUR: 'soffice' introuvable (installez LibreOffice ou utilisez --soffice).")
    except Exception as e:
        return (False, f"ERREUR inattendue: {e.__class__.__name__}: {e}")

def main():
    parser = argparse.ArgumentParser(
        description="Convertit les .doc de 'dedupe' en .docx via LibreOffice, sortie dans un dossier 'docx' au même niveau."
    )
    parser.add_argument("--source", type=str, default=DEFAULT_SOURCE,
                        help="Dossier source contenant les .doc (défaut: ./dedupe)")
    parser.add_argument("--dest", type=str, default="",
                        help="Dossier de sortie .docx (défaut: un dossier 'docx' à côté de --source)")
    parser.add_argument("--soffice", type=str, default="",
                        help="Chemin explicite vers le binaire 'soffice' (sinon détection via PATH)")
    parser.add_argument("--overwrite", action="store_true",
                        help="Autoriser l'écrasement des .docx existants (défaut: non)")
    args = parser.parse_args()

    source_dir = Path(args.source).resolve()
    if not source_dir.exists() or not source_dir.is_dir():
        print(f"❌ Dossier source invalide: {source_dir}")
        sys.exit(1)

    if args.dest:
        dest_dir = Path(args.dest).resolve()
    else:
        dest_dir = (source_dir.parent / "docx").resolve()
    dest_dir.mkdir(parents=True, exist_ok=True)

    soffice_path = find_soffice(args.soffice or None)
    if not soffice_path:
        print("❌ 'soffice' introuvable. Installez LibreOffice ou fournissez --soffice /chemin/vers/soffice")
        sys.exit(1)

    docs = sorted(p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() == ".doc")
    if not docs:
        print("ℹ️ Aucun fichier .doc à convertir dans la source.")
        sys.exit(0)

    print(f"▶️  Conversion DOC -> DOCX")
    print(f"    - Source : {source_dir}")
    print(f"    - Sortie : {dest_dir}")
    print(f"    - Binaire LibreOffice : {soffice_path}")
    print(f"    - Overwrite : {'oui' if args.overwrite else 'non'}\n")

    total = len(docs)
    converted = 0
    failed = 0

    for doc_path in docs:
        ok, msg = convert_one_doc(soffice_path, doc_path, dest_dir, args.overwrite)
        print(f"- {doc_path.name}: {msg}")
        if ok:
            converted += 1
        else:
            failed += 1

    print("\n✅ Terminé.")
    print(f"   Total .doc      : {total}")
    print(f"   Convertis       : {converted}")
    print(f"   Échecs          : {failed}")

if __name__ == "__main__":
    main()
