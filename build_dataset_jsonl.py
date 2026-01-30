#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Construit des datasets JSONL (train/val) pour fine-tuning Mistral Instruct.
Apparie les fichiers EDB (Expression de Besoin) et NDC (Note de Cadrage) par référence.

Gère les cas avec plusieurs versions:
- Plusieurs EDB et plusieurs NDC pour une même référence
- 1 EDB et plusieurs NDC
- Plusieurs EDB et 1 NDC

Dépendances:
  pip install (aucune externe, utilise stdlib)
  (Le rapport Excel utilise pandas + openpyxl, déjà référencé dans le code)

Usage:
  python build_dataset_jsonl.py --edb_dir edb --ndc_dir ndc --report
"""

import json
import re
import argparse
import random
import logging
from pathlib import Path
from collections import defaultdict
from typing import Dict, Tuple, List, Optional, NamedTuple

import pandas as pd

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ==============================================================================
# CONFIGURATION - MODIFIEZ ICI
# ==============================================================================

# ------------------------------
# DOSSIERS PAR DÉFAUT
# ------------------------------
DEFAULT_EDB_DIR = "markdown/edb"
DEFAULT_NDC_DIR = "markdown/ndc"
DEFAULT_TRAIN_OUT = "train_dataset.jsonl"
DEFAULT_VAL_OUT = "val_dataset.jsonl"

# ------------------------------
# SPLIT TRAIN/VAL
# ------------------------------
DEFAULT_TRAIN_RATIO = 0.9  # 90% train, 10% val
DEFAULT_SEED = 42  # Pour reproductibilité

# ------------------------------
# STRATÉGIE DE MAPPING MULTI-FICHIERS
# ------------------------------
# Quand il y a plusieurs fichiers EDB ou NDC pour une même référence:
#
# "version_match" : Apparie par version détectée dans le nom de fichier
#                   (v1 avec v1, Etude avec Etude, etc.)
#                   Si pas de correspondance, utilise le premier de chaque
#
# "all_combinations" : Crée toutes les combinaisons possibles (plus de données)
#                      Ex: 2 EDB x 3 NDC = 6 paires
#
# "latest_only" : Utilise seulement la version la plus récente (par nom ou date)
#
# "first_only" : Utilise seulement le premier fichier trouvé (comportement actuel)

MULTI_FILE_STRATEGY = "version_match"

# Patterns pour extraire la version depuis le nom de fichier
# L'ordre est important : le premier match est utilisé
VERSION_PATTERNS = [
    # Version numérique: v1, v2, V1.0, version2, etc.
    (r'[_\-\s]?[vV](\d+(?:\.\d+)?)', 'v'),
    # Suffixe de stade: _Etude, _Realisation, _Cadrage, etc.
    (r'[_\-]?(Etude|Realisation|Cadrage|Final|Draft|Brouillon)', 'stage'),
    # Numéro de révision: _01, _02, -1, -2
    (r'[_\-](\d{1,2})(?:[_\-\.]|$)', 'rev'),
    # Date dans le nom: 2025-01-15, 20250115
    (r'(\d{4}[-_]?\d{2}[-_]?\d{2})', 'date'),
]

# ------------------------------
# FILTRES CONTENU
# ------------------------------
MIN_CONTENT_CHARS = 100
MAX_CONTENT_CHARS = 0  # 0 = pas de limite
MIN_CONTENT_LINES = 5

# ⚠️ Corrigé : exclusions au niveau "document entier" uniquement.
# - Un document strictement vide
# - Un document qui n'est QUE "ERREUR"
# - Un document qui n'est QUE "[TEMPLATE]"
EXCLUDE_CONTENT_PATTERNS = [
    r'^\s*$',                # document vide
    r'^\s*ERREUR\s*$',       # document = ERREUR
    r'^\s*\[TEMPLATE\]\s*$', # document = [TEMPLATE]
]

REQUIRE_CONTENT_PATTERNS = []

# ------------------------------
# FILTRES NOM DE FICHIER
# ------------------------------
# Pattern: CAGIPRITM suivi de chiffres, puis n'importe quoi (ou rien)
REF_FROM_FILENAME_PATTERN = r"^(CAGIPRITM\d+)"

EXCLUDE_FILENAME_PATTERNS = [
    r'^\..*',
    r'^_.*',
    r'.*\.backup\.md$',
    r'.*\.old\.md$',
]

# ------------------------------
# FORMAT DATASET
# ------------------------------
DATASET_FORMAT = "mistral_instruct"
SYSTEM_PROMPT = ""


# ==============================================================================
# FIN CONFIGURATION
# ==============================================================================

# Compile les patterns
REF_FROM_FILENAME_RE = re.compile(REF_FROM_FILENAME_PATTERN, re.IGNORECASE)
EXCLUDE_FILENAME_RES = [re.compile(p, re.IGNORECASE) for p in EXCLUDE_FILENAME_PATTERNS]
# ⚠️ Corrigé : pas de MULTILINE ici, on fera un fullmatch() sur le doc strip()
EXCLUDE_CONTENT_RES = [re.compile(p, re.IGNORECASE) for p in EXCLUDE_CONTENT_PATTERNS]
REQUIRE_CONTENT_RES = [re.compile(p, re.MULTILINE | re.IGNORECASE) for p in REQUIRE_CONTENT_PATTERNS]
VERSION_RES = [(re.compile(p, re.IGNORECASE), vtype) for p, vtype in VERSION_PATTERNS]


class FileInfo(NamedTuple):
    """Information sur un fichier indexé."""
    path: Path
    content: str
    version: Optional[str]
    version_type: Optional[str]


def read_text(path: Path) -> str:
    """Lit un fichier en UTF-8 (fallback latin-1) et normalise les fins de ligne."""
    try:
        txt = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        txt = path.read_text(encoding="latin-1")
    return txt.replace("\r\n", "\n").replace("\r", "\n")


def extract_ref_from_filename(path: Path) -> Optional[str]:
    """Extrait la référence depuis le nom de fichier."""
    stem = path.stem.strip()
    m = REF_FROM_FILENAME_RE.match(stem)
    return m.group(1).upper() if m else None


def extract_version_from_filename(path: Path) -> Tuple[Optional[str], Optional[str]]:
    """
    Extrait la version depuis le nom de fichier.
    Retourne (version, type_version) ou (None, None).
    """
    stem = path.stem

    for pattern, vtype in VERSION_RES:
        match = pattern.search(stem)
        if match:
            version = match.group(1).lower()
            return version, vtype

    return None, None


def should_exclude_filename(path: Path) -> bool:
    """Vérifie si le fichier doit être exclu selon son nom."""
    name = path.name
    for pattern in EXCLUDE_FILENAME_RES:
        if pattern.match(name):
            return True
    return False


def validate_content(content: str) -> Tuple[bool, str]:
    """Valide le contenu selon les filtres configurés."""
    if MIN_CONTENT_CHARS > 0 and len(content) < MIN_CONTENT_CHARS:
        return False, f"trop court ({len(content)} < {MIN_CONTENT_CHARS} chars)"

    if MAX_CONTENT_CHARS > 0 and len(content) > MAX_CONTENT_CHARS:
        return False, f"trop long ({len(content)} > {MAX_CONTENT_CHARS} chars)"

    non_empty_lines = [l for l in content.split('\n') if l.strip()]
    if MIN_CONTENT_LINES > 0 and len(non_empty_lines) < MIN_CONTENT_LINES:
        return False, f"pas assez de lignes ({len(non_empty_lines)} < {MIN_CONTENT_LINES})"

    # ⚠️ Corrigé : les patterns d'exclusion doivent matcher le document ENTIER
    doc = content.strip()
    for pattern in EXCLUDE_CONTENT_RES:
        if pattern.fullmatch(doc):
            return False, "matche pattern d'exclusion (doc entier)"

    if REQUIRE_CONTENT_RES:
        found = any(pattern.search(content) for pattern in REQUIRE_CONTENT_RES)
        if not found:
            return False, "ne matche aucun pattern requis"

    return True, ""


def index_folder_all_versions(folder: Path) -> Tuple[
    Dict[str, List[FileInfo]],
    List[Path],
    List[Tuple[Path, str]]
]:
    """
    Indexe TOUS les fichiers .md par référence (pas seulement le premier).

    Retourne:
      - index: ref -> [FileInfo, ...] (tous les fichiers pour cette ref)
      - no_ref: liste des fichiers sans ref détectée
      - invalid: liste de (path, reason) pour fichiers filtrés
    """
    index: Dict[str, List[FileInfo]] = defaultdict(list)
    no_ref: List[Path] = []
    invalid: List[Tuple[Path, str]] = []

    for p in sorted(folder.rglob("*.md")):
        if should_exclude_filename(p):
            continue

        ref = extract_ref_from_filename(p)
        if not ref:
            no_ref.append(p)
            continue

        content = read_text(p).strip()

        is_valid, reason = validate_content(content)
        if not is_valid:
            invalid.append((p, reason))
            continue

        version, vtype = extract_version_from_filename(p)
        file_info = FileInfo(path=p, content=content, version=version, version_type=vtype)
        index[ref].append(file_info)

    return dict(index), no_ref, invalid


def match_versions(edb_files: List[FileInfo], ndc_files: List[FileInfo]) -> List[Tuple[FileInfo, FileInfo]]:
    """
    Apparie les fichiers EDB et NDC par version.

    Stratégie:
    1. Si les deux côtés ont des versions du même type, les matcher
    2. Si un seul côté a des versions, dupliquer l'autre
    3. Sinon, utiliser le premier de chaque
    """
    pairs = []

    # Cas simple: 1 EDB, 1 NDC
    if len(edb_files) == 1 and len(ndc_files) == 1:
        return [(edb_files[0], ndc_files[0])]

    # Grouper par version
    edb_by_version = {}
    ndc_by_version = {}

    for f in edb_files:
        key = (f.version, f.version_type) if f.version else ('_default', None)
        edb_by_version[key] = f

    for f in ndc_files:
        key = (f.version, f.version_type) if f.version else ('_default', None)
        ndc_by_version[key] = f

    # Trouver les versions communes
    common_versions = set(edb_by_version.keys()) & set(ndc_by_version.keys())

    if common_versions and ('_default', None) not in common_versions:
        # Matcher par version
        for v in common_versions:
            pairs.append((edb_by_version[v], ndc_by_version[v]))

        # Ajouter les versions non matchées avec le premier de l'autre côté
        for v, edb in edb_by_version.items():
            if v not in common_versions:
                # Utiliser le premier NDC
                pairs.append((edb, ndc_files[0]))

        for v, ndc in ndc_by_version.items():
            if v not in common_versions:
                # Utiliser le premier EDB
                pairs.append((edb_files[0], ndc))
    else:
        # Pas de versions détectées, faire du 1-to-many ou many-to-1
        if len(edb_files) == 1:
            # 1 EDB, plusieurs NDC
            for ndc in ndc_files:
                pairs.append((edb_files[0], ndc))
        elif len(ndc_files) == 1:
            # Plusieurs EDB, 1 NDC
            for edb in edb_files:
                pairs.append((edb, ndc_files[0]))
        else:
            # Plusieurs des deux côtés sans version -> premier de chaque
            pairs.append((edb_files[0], ndc_files[0]))

    return pairs


def create_all_combinations(edb_files: List[FileInfo], ndc_files: List[FileInfo]) -> List[Tuple[FileInfo, FileInfo]]:
    """Crée toutes les combinaisons possibles EDB x NDC."""
    pairs = []
    for edb in edb_files:
        for ndc in ndc_files:
            pairs.append((edb, ndc))
    return pairs


def use_latest_only(files: List[FileInfo]) -> FileInfo:
    """Retourne le fichier le plus récent (par version ou nom)."""
    if len(files) == 1:
        return files[0]

    # Trier par version décroissante si disponible (heuristique)
    versioned = [(f, f.version or '') for f in files]
    versioned.sort(key=lambda x: x[1], reverse=True)

    return versioned[0][0]


def build_pairs(
    edb_index: Dict[str, List[FileInfo]],
    ndc_index: Dict[str, List[FileInfo]],
    strategy: str
) -> List[Tuple[str, FileInfo, FileInfo]]:
    """
    Construit les paires (ref, edb, ndc) selon la stratégie choisie.
    """
    edb_refs = set(edb_index.keys())
    ndc_refs = set(ndc_index.keys())
    common_refs = edb_refs & ndc_refs

    all_pairs = []

    for ref in sorted(common_refs):
        edb_files = edb_index[ref]
        ndc_files = ndc_index[ref]

        if strategy == "all_combinations":
            pairs = create_all_combinations(edb_files, ndc_files)
        elif strategy == "version_match":
            pairs = match_versions(edb_files, ndc_files)
        elif strategy == "latest_only":
            edb = use_latest_only(edb_files)
            ndc = use_latest_only(ndc_files)
            pairs = [(edb, ndc)]
        else:  # first_only
            pairs = [(edb_files[0], ndc_files[0])]

        for edb, ndc in pairs:
            all_pairs.append((ref, edb, ndc))

    return all_pairs


def make_record(edb_text: str, ndc_text: str) -> dict:
    """Crée un enregistrement au format configuré."""
    if DATASET_FORMAT == "mistral_instruct":
        messages = []
        if SYSTEM_PROMPT:
            messages.append({"role": "system", "content": SYSTEM_PROMPT})
        messages.append({"role": "user", "content": edb_text})
        messages.append({"role": "assistant", "content": ndc_text})
        return {"messages": messages}

    elif DATASET_FORMAT == "chatml":
        messages = []
        if SYSTEM_PROMPT:
            messages.append({"role": "system", "content": SYSTEM_PROMPT})
        messages.append({"role": "user", "content": edb_text})
        messages.append({"role": "assistant", "content": ndc_text})
        return {"messages": messages}

    elif DATASET_FORMAT == "alpaca":
        record = {"instruction": edb_text, "output": ndc_text}
        if SYSTEM_PROMPT:
            record["input"] = SYSTEM_PROMPT
        return record

    else:
        return {
            "messages": [
                {"role": "user", "content": edb_text},
                {"role": "assistant", "content": ndc_text}
            ]
        }


def write_jsonl(path: Path, records: List[dict]) -> None:
    """Écrit les enregistrements dans un fichier JSONL."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")


def print_config_summary():
    """Affiche un résumé de la configuration active."""
    logger.info("=== Configuration ===")
    logger.info(f"Format dataset: {DATASET_FORMAT}")
    logger.info(f"Stratégie multi-fichiers: {MULTI_FILE_STRATEGY}")
    logger.info(f"System prompt: {'Oui' if SYSTEM_PROMPT else 'Non'}")
    logger.info(f"Min chars: {MIN_CONTENT_CHARS}")
    logger.info(f"Max chars: {MAX_CONTENT_CHARS if MAX_CONTENT_CHARS > 0 else 'illimité'}")


def main():
    parser = argparse.ArgumentParser(
        description="Construit datasets JSONL pour Mistral Instruct avec support multi-versions."
    )
    parser.add_argument("--edb_dir", default=DEFAULT_EDB_DIR)
    parser.add_argument("--ndc_dir", default=DEFAULT_NDC_DIR)
    parser.add_argument("--train_out", default=DEFAULT_TRAIN_OUT)
    parser.add_argument("--val_out", default=DEFAULT_VAL_OUT)
    parser.add_argument("--train_ratio", type=float, default=DEFAULT_TRAIN_RATIO)
    parser.add_argument("--seed", type=int, default=DEFAULT_SEED)
    parser.add_argument("--min_chars", type=int, default=None)
    parser.add_argument("--max_chars", type=int, default=None)
    parser.add_argument("--strategy", choices=["version_match", "all_combinations", "latest_only", "first_only"],
                        default=None, help="Override MULTI_FILE_STRATEGY")
    parser.add_argument("--report", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--verbose", "-v", action="store_true")

    args = parser.parse_args()

    global MIN_CONTENT_CHARS, MAX_CONTENT_CHARS, MULTI_FILE_STRATEGY
    if args.min_chars is not None:
        MIN_CONTENT_CHARS = args.min_chars
    if args.max_chars is not None:
        MAX_CONTENT_CHARS = args.max_chars
    if args.strategy:
        MULTI_FILE_STRATEGY = args.strategy

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    print_config_summary()

    edb_dir = Path(args.edb_dir).resolve()
    ndc_dir = Path(args.ndc_dir).resolve()
    train_path = Path(args.train_out).resolve()
    val_path = Path(args.val_out).resolve()

    if not edb_dir.exists():
        logger.error(f"Dossier EDB introuvable: {edb_dir}")
        return 1
    if not ndc_dir.exists():
        logger.error(f"Dossier NDC introuvable: {ndc_dir}")
        return 1

    logger.info(f"EDB dir: {edb_dir}")
    logger.info(f"NDC dir: {ndc_dir}")

    # Indexer les dossiers (toutes les versions)
    logger.info("Indexation des fichiers EDB...")
    edb_index, edb_no_ref, edb_invalid = index_folder_all_versions(edb_dir)

    logger.info("Indexation des fichiers NDC...")
    ndc_index, ndc_no_ref, ndc_invalid = index_folder_all_versions(ndc_dir)

    # Stats multi-fichiers
    edb_multi = {ref: files for ref, files in edb_index.items() if len(files) > 1}
    ndc_multi = {ref: files for ref, files in ndc_index.items() if len(files) > 1}

    edb_refs = set(edb_index.keys())
    ndc_refs = set(ndc_index.keys())
    common_refs = edb_refs & ndc_refs
    only_edb = edb_refs - ndc_refs
    only_ndc = ndc_refs - edb_refs

    # Construire les paires selon la stratégie
    logger.info(f"Construction des paires (stratégie: {MULTI_FILE_STRATEGY})...")
    all_pairs = build_pairs(edb_index, ndc_index, MULTI_FILE_STRATEGY)

    # Shuffle + split
    rng = random.Random(args.seed)
    rng.shuffle(all_pairs)

    n_total = len(all_pairs)
    n_train = int(round(n_total * args.train_ratio))
    n_train = min(max(n_train, 1), n_total - 1) if n_total >= 2 else n_total

    train_pairs = all_pairs[:n_train]
    val_pairs = all_pairs[n_train:]

    train_records = [make_record(edb.content, ndc.content) for (_, edb, ndc) in train_pairs]
    val_records = [make_record(edb.content, ndc.content) for (_, edb, ndc) in val_pairs]

    if not args.dry_run:
        write_jsonl(train_path, train_records)
        write_jsonl(val_path, val_records)

    # Générer le rapport Excel
    report_rows = []
    train_set = set(train_pairs)  # pour marquer le split
    for ref, edb, ndc in all_pairs:
        split = "train" if (ref, edb, ndc) in train_set else "val"
        report_rows.append({
            "Code RITM": ref,
            "Fichier EDB": edb.path.name,
            "Version EDB": edb.version or "",
            "Fichier NDC": ndc.path.name,
            "Version NDC": ndc.version or "",
            "Split": split,
        })

    # Ajouter les références orphelines
    for ref in only_edb:
        for f in edb_index.get(ref, []):
            report_rows.append({
                "Code RITM": ref,
                "Fichier EDB": f.path.name,
                "Version EDB": f.version or "",
                "Fichier NDC": "",
                "Version NDC": "",
                "Split": "orphelin_edb",
            })

    for ref in only_ndc:
        for f in ndc_index.get(ref, []):
            report_rows.append({
                "Code RITM": ref,
                "Fichier EDB": "",
                "Version EDB": "",
                "Fichier NDC": f.path.name,
                "Version NDC": f.version or "",
                "Split": "orphelin_ndc",
            })

    report_path = Path.cwd() / "dataset_report.xlsx"
    if not args.dry_run:
        df = pd.DataFrame(report_rows)
        with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Dataset JSONL")

    # Résumé
    print()
    print("=" * 60)
    print("RÉSUMÉ")
    print("=" * 60)
    print(f"Références EDB          : {len(edb_index)}")
    print(f"Références NDC          : {len(ndc_index)}")
    print(f"Références communes     : {len(common_refs)}")
    print()
    print(f"EDB avec multi-fichiers : {len(edb_multi)} refs ({sum(len(f) for f in edb_multi.values())} fichiers)")
    print(f"NDC avec multi-fichiers : {len(ndc_multi)} refs ({sum(len(f) for f in ndc_multi.values())} fichiers)")
    print()
    print(f"Stratégie utilisée      : {MULTI_FILE_STRATEGY}")
    print(f"Paires générées         : {n_total}")
    print()
    print(f"Train : {len(train_records):5d} ({args.train_ratio*100:.1f}%)")
    print(f"Val   : {len(val_records):5d} ({(1-args.train_ratio)*100:.1f}%)")
    print()

    if args.dry_run:
        print("[DRY-RUN] Aucun fichier écrit.")
    else:
        print(f"Train  -> {train_path}")
        print(f"Val    -> {val_path}")
        print(f"Rapport-> {report_path}")

    # Rapport détaillé
    if args.report:
        print()
        print("=" * 60)
        print("RAPPORT DÉTAILLÉ")
        print("=" * 60)

        # Multi-fichiers EDB
        if edb_multi:
            print(f"\nEDB multi-fichiers ({len(edb_multi)} refs):")
            for ref, files in list(edb_multi.items())[:20]:
                print(f"  {ref}:")
                for f in files:
                    v = f"[{f.version_type}:{f.version}]" if f.version else "[no version]"
                    print(f"    - {f.path.name} {v}")
            if len(edb_multi) > 20:
                print(f"  ... et {len(edb_multi) - 20} autres refs")

        # Multi-fichiers NDC
        if ndc_multi:
            print(f"\nNDC multi-fichiers ({len(ndc_multi)} refs):")
            for ref, files in list(ndc_multi.items())[:20]:
                print(f"  {ref}:")
                for f in files:
                    v = f"[{f.version_type}:{f.version}]" if f.version else "[no version]"
                    print(f"    - {f.path.name} {v}")
            if len(ndc_multi) > 20:
                print(f"  ... et {len(ndc_multi) - 20} autres refs")

        # Fichiers filtrés
        if edb_invalid:
            print(f"\nEDB filtrés ({len(edb_invalid)}):")
            for p, reason in edb_invalid[:20]:
                print(f"  - {p.name}: {reason}")

        if ndc_invalid:
            print(f"\nNDC filtrés ({len(ndc_invalid)}):")
            for p, reason in ndc_invalid[:20]:
                print(f"  - {p.name}: {reason}")

        # Références orphelines
        if only_edb:
            print(f"\nUniquement dans EDB ({len(only_edb)}):")
            print(f"  {', '.join(list(only_edb)[:15])}")

        if only_ndc:
            print(f"\nUniquement dans NDC ({len(only_ndc)}):")
            print(f"  {', '.join(list(only_ndc)[:15])}")

        # Aperçu des paires générées
        print(f"\nAperçu des paires générées:")
        for ref, edb, ndc in all_pairs[:10]:
            edb_v = f"[{edb.version}]" if edb.version else ""
            ndc_v = f"[{ndc.version}]" if ndc.version else ""
            print(f"  {ref}: {edb.path.name}{edb_v} <-> {ndc.path.name}{ndc_v}")
        if len(all_pairs) > 10:
            print(f"  ... et {len(all_pairs) - 10} autres paires")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
