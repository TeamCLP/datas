# 📄 Pipeline documentaire – Nettoyage • Dédoublonnage • Conversion • Classification • Export Markdown • Dataset LLM

## 🧩 Schéma global du pipeline (ASCII)

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
        │ 3) convert_to_docx.py                  │
        │ - DOC → DOCX (LibreOffice)             │
        │ - PDF → DOCX (pdf2docx)                │
        │ - Copie des DOCX                       │
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
       │ - Analyse 1ère page                        │
       │ - Détection EDB / NDC / AUTRES             │
       └────────────────┬───────────────────────────┘
                        │
                        ▼
         ┌──────────────────────────────────────────────┐
         │ classified_docx/                              │
         │   ├── edb/                                   │
         │   ├── ndc/                                   │
         │   └── autres/                                │
         └───────────────────────┬──────────────────────┘
                                 │
                                 ▼
      ┌────────────────────────────────────────────────────┐
      │ 5) convert_classified_to_md.py                     │
      │ - DOCX → Markdown                                  │
      │ - Export EDB & NDC                                 │
      └───────────────────┬────────────────────────────────┘
                          │
                          ▼
         ┌──────────────────────────────────────────┐
         │ markdown/                                 │
         │   ├── edb/                                │
         │   └── ndc/                                │
         └───────────────────────┬──────────────────┘
                                 │
                                 ▼
      ┌────────────────────────────────────────────────────┐
      │ 6) build_dataset_jsonl.py                          │
      │ - Appariement EDB ↔ NDC                            │
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

Ce dépôt contient un pipeline complet permettant de transformer un lot de documents bruts en un ensemble :

- propre
- dédoublonné
- homogène
- converti au format DOCX
- classé automatiquement (NDC / EDB / AUTRES)
- exporté en Markdown
- prêt pour fine-tuning LLM (dataset JSONL)

Il repose sur **sept scripts Python** :

**Pipeline principal (étapes 1-5) :**
1. `clean_extension.py`
2. `dedupe.py`
3. `convert_to_docx.py`
4. `classify_docx.py`
5. `convert_classified_to_md.py`

**Scripts complémentaires :**
6. `extract_docx_to_markdown.py` — Extraction DOCX → Markdown (via Excel de mapping)
7. `build_dataset_jsonl.py` — Constitution dataset JSONL pour fine-tuning  

---

# 🚀 Exécution depuis un Pod JupyterLab (template *scribe*)

## ✔️ Instructions exactes

### **1) Créer un Pod**

- Template : **scribe**
- **Sans GPU**
- Ouvrir JupyterLab
- Ouvrir un Terminal

### **2) Installer l’environnement**

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
- Activation du venv  
- Positionnement dans `/home/datas`

### **3) Déposer les données sources**

Placer `raw_datas.tar` ici :

```
/home/datas
```

Puis extraire :

```bash
tar -xvf raw_datas.tar -C raw/
```

### **4) Lancer le pipeline complet**

```bash
python clean_extension.py
python dedupe.py
python convert_to_docx.py
python classify_docx.py
python convert_classified_to_md.py
```

---

# 🧱 Architecture finale

Après exécution :

```
datas/
├── raw/
├── clean_extension/
├── dedupe/
├── docx/
├── classified_docx/
│   ├── edb/
│   ├── ndc/
│   └── autres/
├── markdown/
│   ├── edb/
│   └── ndc/
├── clean_extension.py
├── dedupe.py
├── convert_to_docx.py
├── classify_docx.py
├── convert_classified_to_md.py
├── extract_docx_to_markdown.py
├── build_dataset_jsonl.py
├── train_dataset.jsonl
├── val_dataset.jsonl
└── README.md
```

---

# ⚙️ 1. Étape 1 — Nettoyage des extensions  
**Script : `clean_extension.py`**

### Rôle

- Parcourt `raw/`
- Ne conserve que : `.pdf`, `.doc`, `.docx`
- Ajoute un suffixe `_YYYYMMDD_HHMMSS` en cas de collision
- Produit : `inventaire_raw.xlsx`
- Remplit : `clean_extension/`

### Exécution

```bash
python clean_extension.py
```

---

# 🧹 2. Étape 2 — Dédoublonnage intelligent  
**Script : `dedupe.py`**

### Règles métier

| Cas | Conserver |
|-----|-----------|
| `.docx` présent | `.docx` le plus récent |
| `.doc` sans `.docx` | `.doc` le plus récent |
| seulement PDF | PDF le plus récent |

### Sorties

- répertoire : `dedupe/`
- rapport : `dedupe_report.xlsx`

### Exécution

```bash
python dedupe.py
```

---

# 🔁 3. Étape 3 — Conversion DOC→DOCX & PDF→DOCX  
**Script : `convert_to_docx.py`**

### Rôle

- Conversion `.doc` via LibreOffice  
- Conversion `.pdf` via `pdf2docx`  
- Copie des `.docx` existants  
- Output : `docx/`
- Rapport : `convert_report.xlsx`

### Options

- `--on-exists skip` (défaut)  
- `--on-exists overwrite`  
- `--on-exists suffix`  

### Exécution

```bash
python convert_to_docx.py
```

---

# 🔎 4. Étape 4 — Classification des DOCX
**Script : `classify_docx.py`**

### Rôle

Analyse de la **première page** et du **nom de fichier** selon cet ordre :

1. **NDC** si code détecté en 1ère page
2. **EDB** si le nom contient "edb"
3. **EDB** si le nom contient "expression de besoin(s)"
4. **EDB** si le nom contient "eb" ET pas de code NDC en 1ère page
5. **NDC** si code détecté dans le nom du fichier
6. **EDB** si la 1ère page contient "expression de besoin(s)"
7. **AUTRES** sinon

### Motif NDC

Pattern reconnu : `CLIENT` + `ANNÉE` + `CODE`

- **CLIENT** : `CAPS` ou `AVEM` (tolérance aux espaces internes)
- **ANNÉE** : 4 caractères alphanumériques (ex: `2024`, `A2B3`)
- **CODE** : alphanumérique avec tirets/underscores

Exemples : `CAPS_2024_001`, `AVEM2023-42_PF`, `C A P S_A1B2_123`

### Sorties

```
classified_docx/
    edb/
    ndc/
    autres/
```

### Rapport

```
classify_report.xlsx  (dans le dossier racine datas/)
```

### Exécution

```bash
python classify_docx.py
```

---

# ✍️ 5. Étape 5 — Export Markdown  
**Script : `convert_classified_to_md.py`**

### Rôle

- Convertit en Markdown tous les fichiers de :
  - `classified_docx/ndc/`
  - `classified_docx/edb/`

- Dépose les `.md` dans :
  - `markdown/ndc/`
  - `markdown/edb/`

### Exécution

```bash
python convert_classified_to_md.py
```

---

# 📤 6. Extraction DOCX → Markdown (alternative)
**Script : `extract_docx_to_markdown.py`**

### Rôle

Script alternatif d'extraction basé sur un fichier Excel de mapping :

- Lit un fichier Excel contenant les chemins des EDB et NDC
- Convertit les DOCX en Markdown via **Mammoth** (meilleure qualité)
- Supprime automatiquement : page de garde, table des matières, préambule
- Préserve : titres, paragraphes, listes, tableaux

### Configuration

Modifier les constantes en début de fichier :

```python
EXCEL_NAME = "couverture_EDB_NDC_par_RITM.xlsx"
COL_EDB = 5  # Colonne F
COL_NDC = 6  # Colonne G
EXCEL_FILTERS = [(3, "OUI")]  # Filtre colonne D = "OUI"
```

### Sorties

```
dataset_markdown/
├── edb/
├── ndc/
├── _logs/
└── conversion_report.csv
```

### Exécution

```bash
python extract_docx_to_markdown.py
```

---

# 🤖 7. Constitution du dataset JSONL
**Script : `build_dataset_jsonl.py`**

### Rôle

Construit un dataset JSONL pour fine-tuning LLM (Mistral Instruct) :

- Apparie les fichiers EDB et NDC par référence (ex: `CAGIPRITM123456`)
- Gère les cas multi-versions (plusieurs EDB/NDC pour une même référence)
- Split train/val configurable (90/10 par défaut)
- Format compatible Mistral Instruct / ChatML / Alpaca

### Stratégies de mapping multi-fichiers

| Stratégie | Description |
|-----------|-------------|
| `version_match` | Apparie par version détectée (v1↔v1, Etude↔Etude) |
| `all_combinations` | Crée toutes les combinaisons EDB×NDC |
| `latest_only` | Utilise uniquement la version la plus récente |
| `first_only` | Utilise le premier fichier trouvé |

### Exécution

```bash
# Exécution standard
python build_dataset_jsonl.py

# Avec rapport détaillé
python build_dataset_jsonl.py --report

# Simulation sans écriture
python build_dataset_jsonl.py --dry-run --report

# Options avancées
python build_dataset_jsonl.py --strategy all_combinations --train_ratio 0.8
```

### Sorties

- `train_dataset.jsonl` — Dataset d'entraînement
- `val_dataset.jsonl` — Dataset de validation

---

# 🧭 8. Pipeline complet (ordre recommandé)

```bash
# Pipeline principal (traitement des documents bruts)
python clean_extension.py
python dedupe.py
python convert_to_docx.py
python classify_docx.py
python convert_classified_to_md.py

# Constitution du dataset LLM (après le pipeline principal)
python build_dataset_jsonl.py --report
```

---

# 📊 9. Fichiers Excel/CSV générés

| Étape | Fichier | Emplacement | Contenu |
|-------|---------|-------------|---------|
| Nettoyage | `inventaire_raw.xlsx` | `datas/` | inventaire et actions appliquées |
| Dédoublonnage | `dedupe_report.xlsx` | `datas/` | règles, décisions, justification |
| Conversion | `convert_report.xlsx` | `datas/` | conversion/copied, logs |
| Classification | `classify_report.xlsx` | `datas/` | EDB / NDC / AUTRES + destination |
| Extraction | `conversion_report.csv` | `dataset_markdown/` | statut extraction DOCX → MD |

---

# ⭐ 10. Bonnes pratiques

- Toujours suivre le pipeline dans l’ordre  
- Ne jamais modifier manuellement les dossiers intermédiaires  
- Conserver `--on-exists skip` sauf besoin explicite  
- Utiliser les rapports Excel pour audit et contrôle  

---

# 🧩 11. Résultat attendu

À la fin du pipeline :

- Fichiers nettoyés
- Doublons supprimés
- Corpus converti à 100% en `.docx`
- Documents automatiquement classés
- Export Markdown propre et structuré
- Dataset JSONL prêt pour fine-tuning
- Traçabilité complète

Le pipeline produit un corpus documentaire propre, homogène et un dataset directement exploitable pour le fine-tuning de LLM.
