#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_to_docx.py (corrigé)
- Source par défaut : ./dedupe
- Destination par défaut : ./docx (à côté de la source)
- Convertit tous les .doc -> .docx via LibreOffice (soffice)
- Copie aussi tous les .docx déjà au bon format depuis dedupe vers docx
- Génère un Excel de traçabilité : convert_report.xlsx
- Gestion des collisions via --on-exists {skip|overwrite|suffix}
"""

import argparse
import subprocess
import shutil
from pathlib import Path
from datetime import datetime
import tempfile
import pandas as pd
import sys

DEFAULT_SOURCE = "dedupe"
DEFAULT_REPORT = "convert_report.xlsx"
ON_EXISTS_CHOICES = {"skip", "overwrite", "suffix"}


def find_soffice(user_path: str | None) -> Path | None:
    """Trouve le binaire 'soffice'."""
    if user_path:
        p = Path(user_path)
        return p if p.exists() and p.is_file() else None
    which = shutil.which("soffice")
    return Path(which) if which else None


def run_soffice_convert(soffice: Path, src_doc: Path, out_dir: Path) -> tuple[bool, str]:
    """Exécute LibreOffice pour convertir src_doc -> .docx dans out_dir."""
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(soffice),
        "--headless", "--nologo", "--nodefault", "--invisible",
        "--convert-to", "docx",
        "--outdir", str(out_dir),
        str(src_doc),
    ]
    proc = subprocess.run(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True
    )
    output = (proc.stdout or "").strip()
    produced = out_dir / (src_doc.stem + ".docx")
    return produced.exists(), output


def copy_with_policy(src: Path, dest: Path, on_exists: str) -> tuple[bool, str, str]:
    """
    Copie src -> dest en appliquant la politique de collision.
    Retourne (success, action_message, final_path).
    """
    dest.parent.mkdir(parents=True, exist_ok=True)

    if dest.exists():
        if on_exists == "skip":
            return True, "ignoré (existe déjà)", ""
        elif on_exists == "overwrite":
            shutil.copy2(src, dest)
            return True, "copié (écrasé)", str(dest)
        elif on_exists == "suffix":
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            final = dest.with_name(f"{dest.stem}_{ts}{dest.suffix}")
            shutil.copy2(src, final)
            return True, "copié (suffixé)", str(final)
    else:
        shutil.copy2(src, dest)
        return True, "copié", str(dest)

    return False, "inconnu", ""  # ne devrait pas arriver


def main():
    parser = argparse.ArgumentParser(
        description="Conversion .doc -> .docx via LibreOffice + copie des .docx existants, avec rapport Excel."
    )
    parser.add_argument("--source", type=str, default=DEFAULT_SOURCE,
                        help="Dossier source (défaut: ./dedupe)")
    parser.add_argument("--dest", type=str, default="",
                        help="Dossier cible (défaut: ./docx à côté de --source)")
    parser.add_argument("--soffice", type=str, default="",
                        help="Chemin vers 'soffice' (sinon détection via PATH)")
    parser.add_argument("--on-exists", type=str, default="skip", choices=sorted(ON_EXISTS_CHOICES),
                        help="Collision policy: skip | overwrite | suffix (défaut: skip)")
    parser.add_argument("--report", type=str, default=DEFAULT_REPORT,
                        help="Nom du fichier Excel de rapport (défaut: convert_report.xlsx)")
    args = parser.parse_args()

    source_dir = Path(args.source).resolve()
    if not source_dir.exists() or not source_dir.is_dir():
        print(f"❌ Dossier source invalide: {source_dir}")
        sys.exit(1)

    dest_dir = Path(args.dest).resolve() if args.dest else (source_dir.parent / "docx").resolve()
    dest_dir.mkdir(parents=True, exist_ok=True)

    soffice_path = find_soffice(args.soffice or None)
    if not soffice_path:
        print("❌ 'soffice' introuvable. Installez LibreOffice ou fournissez --soffice /chemin/vers/soffice")
        sys.exit(1)

    # Collecte
    docs = sorted([p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() == ".doc"])
    docxs = sorted([p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() == ".docx"])

    print(f"▶️  Conversion & copie")
    print(f"    - Source : {source_dir}")
    print(f"    - Sortie : {dest_dir}")
    print(f"    - soffice: {soffice_path}")
    print(f"    - Collision: {args.on_exists}\n")

    rows = []
    total_docs = len(docs)
    total_docxs = len(docxs)
    converted = overwritten = suffixed_conv = skipped_conv = failed_conv = 0
    copied = overwritten_copy = suffixed_copy = skipped_copy = 0

    # 1) Conversion des .doc -> .docx
    for doc_path in docs:
        expected = dest_dir / (doc_path.stem + ".docx")
        action = ""
        message = ""
        out_path = ""

        if expected.exists():
            if args.on_exists == "skip":
                action, message = "ignoré", "existe déjà (skip)"
                skipped_conv += 1
            elif args.on_exists == "overwrite":
                ok, out = run_soffice_convert(soffice_path, doc_path, dest_dir)
                if ok:
                    action, message, out_path = "converti (écrasé)", "overwrite", str(expected)
                    overwritten += 1
                else:
                    action, message = "échec", f"conversion échouée | {out}"
                    failed_conv += 1
            elif args.on_exists == "suffix":
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmp_out = Path(tmpdir)
                    ok, out = run_soffice_convert(soffice_path, doc_path, tmp_out)
                    if ok:
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        final = expected.with_name(f"{expected.stem}_{ts}{expected.suffix}")
                        (tmp_out / expected.name).replace(final)
                        action, message, out_path = "converti (suffixé)", "collision évitée", str(final)
                        suffixed_conv += 1
                    else:
                        action, message = "échec", f"conversion échouée | {out}"
                        failed_conv += 1
        else:
            ok, out = run_soffice_convert(soffice_path, doc_path, dest_dir)
            if ok:
                action, message, out_path = "converti", "OK", str(expected)
                converted += 1
            else:
                action, message = "échec", f"conversion échouée | {out}"
                failed_conv += 1

        rows.append({
            "Type": "DOC->DOCX",
            "Fichier source": doc_path.name,
            "Chemin source": str(doc_path),
            "Action": action,
            "Message": message,
            "Fichier généré": out_path,
        })
        print(f"- [DOC] {doc_path.name}: {action} | {message} -> {out_path if out_path else expected}")

    # 2) Copie des .docx déjà au bon format
    for src_docx in docxs:
        expected = dest_dir / src_docx.name
        ok, msg, out_path = copy_with_policy(src_docx, expected, args.on_exists)
        if "ignoré" in msg:
            skipped_copy += 1
        elif "écrasé" in msg:
            overwritten_copy += 1
        elif "suffixé" in msg:
            suffixed_copy += 1
        elif "copié" in msg:
            copied += 1

        rows.append({
            "Type": "COPIE DOCX",
            "Fichier source": src_docx.name,
            "Chemin source": str(src_docx),
            "Action": "copié" if ok else "échec",
            "Message": msg,
            "Fichier généré": out_path,
        })
        print(f"- [DOCX] {src_docx.name}: {msg} -> {out_path if out_path else expected}")

    # Rapport Excel
    report_path = Path(args.report).resolve()
    df = pd.DataFrame(rows, columns=["Type", "Fichier source", "Chemin source", "Action", "Message", "Fichier généré"])
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Conversions & Copies")

    print("\n✅ Terminé.")
    print(f"   DOC traités     : {total_docs} | convertis: {converted}, écrasés: {overwritten}, suffixés: {suffixed_conv}, ignorés: {skipped_conv}, échecs: {failed_conv}")
    print(f"   DOCX traités    : {total_docxs} | copiés: {copied}, écrasés: {overwritten_copy}, suffixés: {suffixed_copy}, ignorés: {skipped_copy}")
    print(f"📄 Rapport Excel   : {report_path}")


if __name__ == "__main__":
    main()
