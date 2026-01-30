# 📄 Pipeline documentaire – Nettoyage • Dédoublonnage • Conversion • Classification • Export Markdown • Dataset LLM

## 🧩 Schéma global du pipeline

```
                 ┌────────────────────┐
                 │    raw/ (brut)     │
                 └─────────┬──────────┘
                           │
                           ▼
            ┌────────────────────────────────┐
            │ 1) clean_extension.py          │
            │ - Filtrage extensions          │
            │ - Suffixes anti‑collision      │
            └───────────┬────────────────────┘
                        │
                        ▼
          ┌────────────────────────────────────┐
          │ clean_extension/                    │
          └──────────────────┬──────────────────┘
                             │
                             ▼
            ┌────────────────────────────────┐
            │ 2) dedupe.py                   │
            │ - Règles DOC/DOCX/PDF          │
            │ - Sélection fichier le + récent│
            └───────────┬────────────────────┘
                        │
                        ▼
               ┌────────────────────────┐
               │       dedupe/          │
               └──────────┬─────────────┘
                          │
                          ▼
        ┌────────────────────────────────────────┐
        │ 3) convert_to_docx.py (parallélisé)    │
        │ - DOC → DOCX (LibreOffice)             │
        │ - PDF → DOCX (pdf2docx)                │
        └────────────────┬───────────────────────┘
                         │
                         ▼
                   ┌───────────┐
                   │   docx/   │
                   └─────┬─────┘
                         │
                         ▼
       ┌────────────────────────────────────────────┐
       │ 4) classify_docx.py                        │
       │ - Analyse 1ère page + nom fichier          │
       │ - Détection EDB / NDC / AUTRES             │
       └────────────────┬───────────────────────────┘
                        │
                        ▼
         ┌──────────────────────────────────────────────┐
         │ classified_docx/                              │
         │   ├── edb/   (CAGIPRITM...)                  │
         │   ├── ndc/   (CAGIPRITM...)                  │
         │   └── autres/                                │
         └───────────────────────┬──────────────────────┘
                                 │
                                 ▼
      ┌────────────────────────────────────────────────────┐
      │ 5) extract_docx_to_markdown.py (parallélisé)       │
      │ - DOCX → Markdown (Mammoth)                        │
      │ - Suppression TOC, page de garde                   │
      └───────────────────┬────────────────────────────────┘
                          │
                          ▼
         ┌──────────────────────────────────────────┐
         │ markdown/                                 │
         │   ├── edb/*.md                           │
         │   └── ndc/*.md                           │
         └───────────────────────┬──────────────────┘
                                 │
                                 ▼
      ┌────────────────────────────────────────────────────┐
      │ 6) build_dataset_jsonl.py                          │
      │ - Appariement EDB ↔ NDC par code RITM              │
      │ - Export JSONL pour fine-tuning                    │
      └───────────────────┬────────────────────────────────┘
                          │
                          ▼
         ┌──────────────────────────────────────────┐
         │ train_dataset.jsonl                       │
         │ val_dataset.jsonl                         │
         └──────────────────────────────────────────┘
```

---

# 📘 Description générale

Ce dépôt contient un pipeline complet permettant de transformer un lot de documents bruts en un dataset prêt pour le fine-tuning LLM :

- Nettoyage et filtrage des fichiers
- Dédoublonnage intelligent
- Conversion homogène en DOCX
- Classification automatique (NDC / EDB / AUTRES)
- Export Markdown de qualité
- Constitution du dataset JSONL

Il repose sur **6 scripts Python**, exécutés dans cet ordre :

1. `clean_extension.py` — Filtrage des extensions valides
2. `dedupe.py` — Dédoublonnage intelligent
3. `convert_to_docx.py` — Conversion DOC/PDF → DOCX (parallélisé)
4. `classify_docx.py` — Classification EDB / NDC / AUTRES par code RITM
5. `extract_docx_to_markdown.py` — Export Markdown avec Mammoth (parallélisé)
6. `build_dataset_jsonl.py` — Constitution dataset JSONL pour fine-tuning

**Code RITM** : Les fichiers sont identifiés par leur code `CAGIPRITMNNNNNNN` au début du nom de fichier.

---

# 🚀 Exécution depuis un Pod JupyterLab (template *scribe*)

## ✔️ Instructions exactes

### **1) Créer un Pod**

- Template : **scribe**
- **Sans GPU**
- Ouvrir JupyterLab
- Ouvrir un Terminal

### **2) Installer l'environnement**

```bash
bash
git clone https://github.com/TeamCLP/datas.git /home/datas && source /home/datas/install.sh
```

Le script `install.sh` configure automatiquement :

- Proxy
- LibreOffice
- Miniconda + Python 3.13
- Environnement `pipeline`
- Installation du `requirements.txt`

### **3) Déposer les données sources**

Placer `raw_datas.tar` dans `/home/datas` puis extraire :

```bash
tar -xvf raw_datas.tar -C raw/
```

### **4) Lancer le pipeline complet**

```bash
python clean_extension.py
python dedupe.py
python convert_to_docx.py
python classify_docx.py
python extract_docx_to_markdown.py
python build_dataset_jsonl.py --report
```

---

# 🧱 Architecture finale

```
datas/
├── raw/                          # Fichiers bruts d'entrée
├── clean_extension/              # Fichiers filtrés
├── dedupe/                       # Fichiers dédoublonnés
├── docx/                         # Tous les fichiers en DOCX
├── classified_docx/
│   ├── edb/                      # Expressions de Besoin
│   ├── ndc/                      # Notes de Cadrage
│   └── autres/                   # Non classés
├── markdown/
│   ├── edb/                      # EDB en Markdown
│   └── ndc/                      # NDC en Markdown
├── train_dataset.jsonl           # Dataset d'entraînement
├── val_dataset.jsonl             # Dataset de validation
└── *.py                          # Scripts du pipeline
```

---

# ⚙️ 1. Nettoyage des extensions
**Script : `clean_extension.py`**

- Parcourt `raw/`
- Ne conserve que : `.pdf`, `.doc`, `.docx`
- Ajoute un suffixe `_YYYYMMDD_HHMMSS` en cas de collision
- Produit : `inventaire_raw.xlsx`

```bash
python clean_extension.py
```

---

# 🧹 2. Dédoublonnage intelligent
**Script : `dedupe.py`**

| Cas | Conserver |
|-----|-----------|
| `.docx` présent | `.docx` le plus récent |
| `.doc` sans `.docx` | `.doc` le plus récent |
| seulement PDF | PDF le plus récent |

- Produit : `dedupe_report.xlsx`

```bash
python dedupe.py
```

---

# 🔁 3. Conversion DOC/PDF → DOCX
**Script : `convert_to_docx.py`** (parallélisé)

- Conversion `.doc` via LibreOffice
- Conversion `.pdf` via `pdf2docx`
- Copie des `.docx` existants
- Produit : `convert_report.xlsx`

```bash
python convert_to_docx.py
python convert_to_docx.py --workers 4  # limiter à 4 workers
```

---

# 🔎 4. Classification EDB / NDC / AUTRES
**Script : `classify_docx.py`**

Analyse de la **première page** et du **nom de fichier** :

1. **NDC** si code client détecté en 1ère page
2. **EDB** si le nom contient "edb" ou "expression de besoin"
3. **NDC** si code client détecté dans le nom
4. **AUTRES** sinon

**Codes clients reconnus** : `CAPS`, `AVEM` (ex: `CAPS_2024_001`)

- Produit : `classify_report.xlsx`

```bash
python classify_docx.py
```

---

# 📤 5. Export Markdown
**Script : `extract_docx_to_markdown.py`** (parallélisé)

- Scanne `classified_docx/edb/` et `classified_docx/ndc/`
- Identifie les fichiers par leur code RITM (`CAGIPRITMNNNNNNN`)
- Convertit les DOCX en Markdown via **Mammoth**
- Supprime automatiquement : page de garde, table des matières, préambule
- Produit : `extract_report.xlsx`

```bash
python extract_docx_to_markdown.py
python extract_docx_to_markdown.py --workers 4
```

---

# 🤖 6. Constitution du dataset JSONL
**Script : `build_dataset_jsonl.py`**

- Scanne `markdown/edb/` et `markdown/ndc/`
- Apparie les fichiers EDB ↔ NDC par code RITM commun
- Gère les cas multi-versions
- Split train/val (90/10 par défaut)
- Format Mistral Instruct
- Produit : `dataset_report.xlsx`

| Stratégie | Description |
|-----------|-------------|
| `version_match` | Apparie par version (v1↔v1) |
| `all_combinations` | Toutes les combinaisons EDB×NDC |
| `latest_only` | Version la plus récente uniquement |
| `first_only` | Premier fichier trouvé |

```bash
python build_dataset_jsonl.py --report
python build_dataset_jsonl.py --strategy all_combinations --train_ratio 0.8
```

---

# 📊 Fichiers générés

| Étape | Fichier | Contenu |
|-------|---------|---------|
| 1 | `inventaire_raw.xlsx` | Inventaire et actions |
| 2 | `dedupe_report.xlsx` | Décisions de dédoublonnage |
| 3 | `convert_report.xlsx` | Statut des conversions |
| 4 | `classify_report.xlsx` | Classification EDB/NDC/AUTRES |
| 5 | `extract_report.xlsx` | Extraction DOCX → Markdown |
| 6 | `dataset_report.xlsx` | Appariements EDB/NDC et orphelins |
| 6 | `train_dataset.jsonl` | Dataset d'entraînement |
| 6 | `val_dataset.jsonl` | Dataset de validation |

---

# ⭐ Bonnes pratiques

- Toujours suivre le pipeline dans l'ordre
- Ne jamais modifier manuellement les dossiers intermédiaires
- Utiliser `--report` pour diagnostiquer les problèmes
- Vérifier les codes RITM communs entre EDB et NDC

---

# 🧩 Résultat attendu

À la fin du pipeline :

- Corpus nettoyé et dédoublonné
- Documents classés par type (EDB/NDC)
- Export Markdown de qualité
- Dataset JSONL prêt pour fine-tuning LLM
- Traçabilité complète via les rapports Excel
