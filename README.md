# 📄 Pipeline documentaire – Nettoyage • Dédoublonnage • Conversion • Classification • Export Markdown

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

Il repose sur **cinq scripts Python**, exécutés dans cet ordre :

1. `clean_extension.py`  
2. `dedupe.py`  
3. `convert_to_docx.py`  
4. `classify_docx.py`  
5. `convert_classified_to_md.py`  

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

# 🧭 6. Pipeline complet (ordre recommandé)

```bash
python clean_extension.py
python dedupe.py
python convert_to_docx.py
python classify_docx.py
python convert_classified_to_md.py
```

---

# 📊 7. Fichiers Excel générés

| Étape | Fichier | Emplacement | Contenu |
|-------|---------|-------------|---------|
| Nettoyage | `inventaire_raw.xlsx` | `datas/` | inventaire et actions appliquées |
| Dédoublonnage | `dedupe_report.xlsx` | `datas/` | règles, décisions, justification |
| Conversion | `convert_report.xlsx` | `datas/` | conversion/copied, logs |
| Classification | `classify_report.xlsx` | `datas/` | EDB / NDC / AUTRES + destination |

---

# ⭐ Bonnes pratiques

- Toujours suivre le pipeline dans l’ordre  
- Ne jamais modifier manuellement les dossiers intermédiaires  
- Conserver `--on-exists skip` sauf besoin explicite  
- Utiliser les rapports Excel pour audit et contrôle  

---

# 🧩 Résultat attendu

À la fin du pipeline :

- Fichiers nettoyés  
- Doublons supprimés  
- Corpus converti à 100% en `.docx`  
- Documents automatiquement classés  
- Export Markdown propre et structuré  
- Traçabilité complète  

Le pipeline produit un corpus documentaire propre, homogène et exploitable immédiatement.
