#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
convert_to_docx.py (DOC + PDF -> DOCX) - VERSION PARALLÉLISÉE
- Source par défaut : ./dedupe
- Destination par défaut : ./docx (à côté de la source)
- Convertit tous les .doc -> .docx via LibreOffice (soffice)
- Convertit tous les .pdf -> .docx via pdf2docx
- Copie aussi tous les .docx déjà au bon format depuis dedupe vers docx
- Génère un Excel de traçabilité : convert_report.xlsx
- Gestion des collisions via --on-exists {skip|overwrite|suffix}
- Parallélisation configurable via --workers
"""

import argparse
import subprocess
import shutil
import os
from pathlib import Path
from datetime import datetime
import tempfile
import pandas as pd
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from typing import Tuple, Optional

# --- PDF conversion (pdf2docx) ---
try:
    from pdf2docx import Converter as PdfConverter
    PDF2DOCX_AVAILABLE = True
except Exception:
    PDF2DOCX_AVAILABLE = False

DEFAULT_SOURCE = "dedupe"
DEFAULT_REPORT = "convert_report.xlsx"
ON_EXISTS_CHOICES = {"skip", "overwrite", "suffix"}
DEFAULT_WORKERS = 0  # 0 = auto (nombre de CPU)


def find_soffice(user_path: Optional[str]) -> Optional[Path]:
    """Trouve le binaire 'soffice'."""
    if user_path:
        p = Path(user_path)
        return p if p.exists() and p.is_file() else None
    which = shutil.which("soffice")
    return Path(which) if which else None


def run_soffice_convert(soffice: Path, src_doc: Path, out_dir: Path) -> Tuple[bool, str]:
    """
    Exécute LibreOffice pour convertir src_doc -> .docx dans out_dir.
    Utilise un profil utilisateur temporaire pour permettre le parallélisme.
    Retourne (success, output_text).
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    # Créer un profil temporaire unique pour cette instance
    with tempfile.TemporaryDirectory() as user_profile:
        cmd = [
            str(soffice),
            f"-env:UserInstallation=file://{user_profile}",
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


# ------------------------------
# Fonctions de traitement unitaire (pour parallélisation)
# ------------------------------

def process_doc(args: Tuple) -> dict:
    """Traite un fichier .doc -> .docx"""
    doc_path, dest_dir, soffice_path, on_exists = args
    doc_path = Path(doc_path)
    dest_dir = Path(dest_dir)
    soffice_path = Path(soffice_path)

    expected = dest_dir / (doc_path.stem + ".docx")
    action = ""
    message = ""
    out_path = ""

    if expected.exists():
        if on_exists == "skip":
            action, message = "ignoré", "existe déjà (skip)"
        elif on_exists == "overwrite":
            ok, out = run_soffice_convert(soffice_path, doc_path, dest_dir)
            if ok:
                action, message, out_path = "converti (écrasé)", "overwrite", str(expected)
            else:
                action, message = "échec", f"conversion échouée | {out}"
        elif on_exists == "suffix":
            with tempfile.TemporaryDirectory() as tmpdir:
                tmp_out = Path(tmpdir)
                ok, out = run_soffice_convert(soffice_path, doc_path, tmp_out)
                if ok:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    final = expected.with_name(f"{expected.stem}_{ts}{expected.suffix}")
                    (tmp_out / expected.name).replace(final)
                    action, message, out_path = "converti (suffixé)", "collision évitée", str(final)
                else:
                    action, message = "échec", f"conversion échouée | {out}"
    else:
        ok, out = run_soffice_convert(soffice_path, doc_path, dest_dir)
        if ok:
            action, message, out_path = "converti", "OK", str(expected)
        else:
            action, message = "échec", f"conversion échouée | {out}"

    return {
        "Type": "DOC->DOCX",
        "Fichier source": doc_path.name,
        "Chemin source": str(doc_path),
        "Action": action,
        "Message": message,
        "Fichier généré": out_path,
    }


def process_pdf(args: Tuple) -> dict:
    """Traite un fichier .pdf -> .docx"""
    pdf_path, dest_dir, on_exists = args
    pdf_path = Path(pdf_path)
    dest_dir = Path(dest_dir)

    expected = dest_dir / (pdf_path.stem + ".docx")
    action = ""
    message = ""
    out_path = ""

    if not PDF2DOCX_AVAILABLE:
        return {
            "Type": "PDF->DOCX",
            "Fichier source": pdf_path.name,
            "Chemin source": str(pdf_path),
            "Action": "échec",
            "Message": "pdf2docx non disponible",
            "Fichier généré": "",
        }

    try:
        if expected.exists():
            if on_exists == "skip":
                action, message, out_path = "ignoré", "existe déjà (skip)", ""
            elif on_exists == "overwrite":
                try:
                    expected.unlink()
                except Exception:
                    pass
                cv = PdfConverter(str(pdf_path))
                cv.convert(str(expected), start=0, end=None)
                cv.close()
                if expected.exists():
                    action, message, out_path = "converti", "converti (écrasé)", str(expected)
                else:
                    action, message = "échec", "échec conversion (sortie manquante)"
            elif on_exists == "suffix":
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                final_path = expected.with_name(f"{expected.stem}_{ts}{expected.suffix}")
                cv = PdfConverter(str(pdf_path))
                cv.convert(str(final_path), start=0, end=None)
                cv.close()
                if final_path.exists():
                    action, message, out_path = "converti", "converti (suffixé)", str(final_path)
                else:
                    action, message = "échec", "échec conversion (sortie manquante)"
        else:
            cv = PdfConverter(str(pdf_path))
            cv.convert(str(expected), start=0, end=None)
            cv.close()
            if expected.exists():
                action, message, out_path = "converti", "converti", str(expected)
            else:
                action, message = "échec", "échec conversion (sortie manquante)"

    except Exception as e:
        action, message = "échec", f"échec conversion PDF ({e.__class__.__name__}: {e})"

    return {
        "Type": "PDF->DOCX",
        "Fichier source": pdf_path.name,
        "Chemin source": str(pdf_path),
        "Action": action,
        "Message": message,
        "Fichier généré": out_path,
    }


def process_copy(args: Tuple) -> dict:
    """Copie un fichier .docx"""
    src_docx, dest_dir, on_exists = args
    src_docx = Path(src_docx)
    dest_dir = Path(dest_dir)

    expected = dest_dir / src_docx.name
    dest_dir.mkdir(parents=True, exist_ok=True)

    action = ""
    msg = ""
    out_path = ""

    if expected.exists():
        if on_exists == "skip":
            action, msg = "copié", "ignoré (existe déjà)"
        elif on_exists == "overwrite":
            shutil.copy2(src_docx, expected)
            action, msg, out_path = "copié", "copié (écrasé)", str(expected)
        elif on_exists == "suffix":
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            final = expected.with_name(f"{expected.stem}_{ts}{expected.suffix}")
            shutil.copy2(src_docx, final)
            action, msg, out_path = "copié", "copié (suffixé)", str(final)
    else:
        shutil.copy2(src_docx, expected)
        action, msg, out_path = "copié", "copié", str(expected)

    return {
        "Type": "COPIE DOCX",
        "Fichier source": src_docx.name,
        "Chemin source": str(src_docx),
        "Action": action,
        "Message": msg,
        "Fichier généré": out_path,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Conversion DOC & PDF -> DOCX + copie des DOCX existants (parallélisé)."
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
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS,
                        help="Nombre de workers (défaut: 0 = auto)")
    args = parser.parse_args()

    source_dir = Path(args.source).resolve()
    if not source_dir.exists() or not source_dir.is_dir():
        print(f"❌ Dossier source invalide: {source_dir}")
        sys.exit(1)

    dest_dir = Path(args.dest).resolve() if args.dest else (source_dir.parent / "docx").resolve()
    dest_dir.mkdir(parents=True, exist_ok=True)

    # LibreOffice pour DOC -> DOCX
    soffice_path = find_soffice(args.soffice or None)
    if not soffice_path:
        print("❌ 'soffice' introuvable. Installez LibreOffice ou fournissez --soffice /chemin/vers/soffice")
        sys.exit(1)

    # Collecte
    docs = sorted([p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() == ".doc"])
    pdfs = sorted([p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"])
    docxs = sorted([p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() == ".docx"])

    workers = args.workers if args.workers > 0 else os.cpu_count()
    total = len(docs) + len(pdfs) + len(docxs)

    print(f"▶️  Conversion & copie (parallélisé avec {workers} workers)")
    print(f"    - Source : {source_dir}")
    print(f"    - Sortie : {dest_dir}")
    print(f"    - soffice: {soffice_path}")
    print(f"    - pdf2docx: {'OK' if PDF2DOCX_AVAILABLE else 'NON DISPONIBLE'}")
    print(f"    - Collision: {args.on_exists}")
    print(f"    - Fichiers: {len(docs)} DOC, {len(pdfs)} PDF, {len(docxs)} DOCX\n")

    rows = []
    stats = {
        "doc_ok": 0, "doc_skip": 0, "doc_fail": 0,
        "pdf_ok": 0, "pdf_skip": 0, "pdf_fail": 0,
        "copy_ok": 0, "copy_skip": 0,
    }

    # Préparer les tâches
    doc_tasks = [(str(p), str(dest_dir), str(soffice_path), args.on_exists) for p in docs]
    pdf_tasks = [(str(p), str(dest_dir), args.on_exists) for p in pdfs]
    copy_tasks = [(str(p), str(dest_dir), args.on_exists) for p in docxs]

    completed = 0

    with ProcessPoolExecutor(max_workers=workers) as executor:
        # Soumettre toutes les tâches
        futures = {}

        for task in doc_tasks:
            futures[executor.submit(process_doc, task)] = "DOC"
        for task in pdf_tasks:
            futures[executor.submit(process_pdf, task)] = "PDF"
        for task in copy_tasks:
            futures[executor.submit(process_copy, task)] = "COPY"

        # Collecter les résultats
        for future in as_completed(futures):
            completed += 1
            task_type = futures[future]
            result = future.result()
            rows.append(result)

            # Mise à jour des stats
            action = result["Action"]
            if task_type == "DOC":
                if "ignoré" in action:
                    stats["doc_skip"] += 1
                elif "échec" in action:
                    stats["doc_fail"] += 1
                else:
                    stats["doc_ok"] += 1
            elif task_type == "PDF":
                if "ignoré" in action:
                    stats["pdf_skip"] += 1
                elif "échec" in action:
                    stats["pdf_fail"] += 1
                else:
                    stats["pdf_ok"] += 1
            else:  # COPY
                if "ignoré" in result["Message"]:
                    stats["copy_skip"] += 1
                else:
                    stats["copy_ok"] += 1

            # Affichage progression
            print(f"[{completed}/{total}] [{task_type:4}] {result['Fichier source']}: {result['Action']}")

    # Rapport Excel
    report_path = Path(args.report).resolve()
    df = pd.DataFrame(rows, columns=["Type", "Fichier source", "Chemin source", "Action", "Message", "Fichier généré"])
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Conversions & Copies")

    print("\n✅ Terminé.")
    print(f"   DOC  : {len(docs):3d} | convertis: {stats['doc_ok']}, ignorés: {stats['doc_skip']}, échecs: {stats['doc_fail']}")
    print(f"   PDF  : {len(pdfs):3d} | convertis: {stats['pdf_ok']}, ignorés: {stats['pdf_skip']}, échecs: {stats['pdf_fail']}")
    print(f"   DOCX : {len(docxs):3d} | copiés: {stats['copy_ok']}, ignorés: {stats['copy_skip']}")
    print(f"📄 Rapport Excel : {report_path}")


if __name__ == "__main__":
    main()
