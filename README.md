# 📄 Pipeline documentaire – Nettoyage • Dédoublonnage • Conversion DOC→DOCX

Ce dépôt contient un pipeline complet permettant de transformer un lot de documents bruts en un ensemble propre, dédoublonné, homogène et converti au format DOCX.  
Il repose sur **trois scripts Python** travaillant de manière séquentielle :

1. `clean_extension.py` → Nettoyage & filtrage des extensions  
2. `dedupe.py` → Dédoublonnage intelligent (Word > PDF)  
3. `convert_to_docx.py` → Conversion DOC→DOCX, PDF→DOCX + copie des DOCX

L’objectif final est d’obtenir un corpus documentaire propre, cohérent et normalisé.

---

# 🚀 Exécution depuis un Pod JupyterLab (template *scribe*)

## ✔️ Instructions exactes à suivre

### **1) Créer un Pod**
- Utiliser le **template scribe**
- **Ne pas allouer de GPU**
- Ouvrir JupyterLab
- Ouvrir un Terminal

### **2) Installer l’environnement**
Dans le terminal JupyterLab :

```bash
bash
git clone https://github.com/TeamCLP/datas.git /home/datas && source /home/datas/install.sh
```

> Le script `install.sh` configure automatiquement :  
> - Proxy  
> - LibreOffice  
> - Miniconda + Python 3.13  
> - Environnement conda `pipeline`  
> - Installation du `requirements.txt`  
> - Activation automatique du venv  
> - Positionnement dans `/home/datas`

### **3) Déposer les données sources**
Déposer `raw_datas.tar` dans :

```
/home/datas
```

Puis exécuter :

```bash
mkdir raw && tar -xvf raw_datas.tar -C raw/
```

### **4) Lancer le pipeline**
Toujours depuis `/home/datas` avec conda actif :

```bash
python clean_extension.py
python dedupe.py
python convert_to_docx.py
```

---

# 🧱 Architecture finale

Après exécution :

```
datas/
├── raw/                   # Contenu brut extrait
├── clean_extension/       # Fichiers filtrés + Excel de traçabilité
├── dedupe/                # Fichiers dédoublonnés + Excel de traçabilité
├── docx/                  # Fichiers convertis + copies + Excel
│
├── clean_extension.py
├── dedupe.py
├── convert_to_docx.py
└── README.md
```

---

# ⚙️ 1. Préparation de l’environnement (si exécution hors Pod)

### Installer Python et LibreOffice

```bash
echo -e 'Acquire::http::Proxy "http://10.246.42.30:8080";\nAcquire::https::Proxy "http://10.246.42.30:8080";' > /etc/apt/apt.conf.d/95proxies
apt update
apt-get install -y python3 python3-pip
apt-get install -y libreoffice
soffice --version
```

### Installer les dépendances Python

```bash
pip install pandas openpyxl pdf2docx
```

---

# 📥 2. Récupération du dépôt & préparation des données

Cloner le repo :

```bash
git clone https://github.com/TeamCLP/datas.git
cd datas
```

Déposer `raw_datas.tar` dans ce dossier, puis :

```bash
mkdir raw
tar -xvf raw_datas.tar -C raw/
```

Vous obtenez :

```
datas/
└── raw/
    ├── fichier1.pdf
    ├── fichier2.doc
    ├── fichier3.docx
    └── ...
```

---

# 🚀 3. Étape 1 — Nettoyage des extensions  
**Script : `clean_extension.py`**

### Rôle

- Parcourt le dossier `raw/`
- Ne conserve que :
  - `.pdf`
  - `.doc`
  - `.docx`
- Ajoute un suffixe anti-collision `_YYYYMMDD_HHMMSS` si nécessaire
- Produit un rapport Excel : **`inventaire_raw.xlsx`**
- Remplit le dossier `clean_extension/`

### Exécution

```bash
python3 clean_extension.py
```

Sorties :

```
clean_extension/
inventaire_raw.xlsx
```

---

# 🧹 4. Étape 2 — Dédoublonnage intelligent  
**Script : `dedupe.py`**

### Règles métier appliquées (par nom de base, suffixe horodaté neutralisé)

| Cas | Ce qu’on garde |
|-----|----------------|
| `.docx` présent | le `.docx` **le plus récent** |
| `.doc` sans `.docx` | le `.doc` **le plus récent** |
| uniquement PDF | le PDF **le plus récent** |

Tous les autres fichiers du groupe → **ignorés**.

### Fonctionnalités

- Génère un rapport Excel **avant copie** : `dedupe_report.xlsx`
- Explique pour chaque fichier :
  - Action (conserver / ignorer)
  - Raison
  - Chemins source & destination
- Copie les fichiers “conserver” dans : **`dedupe/`**

### Exécution

```bash
python3 dedupe.py
```

Mode simulation (sans copier) :

```bash
python3 dedupe.py --dry-run
```

Sorties :

```
dedupe/
dedupe_report.xlsx
```

---

# 🔁 5. Étape 3 — Conversion DOC→DOCX, PDF→DOCX & copie des DOCX  
**Script : `convert_to_docx.py`**

## Rôle

Ce script traite **trois types d’entrées** depuis `dedupe/` :

1. **`.doc` → `.docx`** via LibreOffice (`soffice`)  
2. **`.pdf` → `.docx`** via la librairie **pdf2docx**  
3. **`.docx` → copie directe**  

Tous les fichiers sont déposés dans :

```
docx/
```

Un rapport unique assure la traçabilité :

```
convert_report.xlsx
```

---

## Règles appliquées aux PDF

- Tous les `.pdf` présents dans `dedupe/` sont convertis en `.docx`
- Conversion réalisée via **pdf2docx**
- Gestion des collisions via `--on-exists` :

| Option         | Comportement PDF → DOCX |
|----------------|--------------------------|
| `skip`         | ignore si le `.docx` existe déjà |
| `overwrite`    | remplace le `.docx` existant |
| `suffix`       | crée `nom_YYYYMMDD_HHMMSS.docx` |

---

## Dépendances PDF

La conversion PDF nécessite :

```
pdf2docx
```

Ce package est installé automatiquement via `requirements.txt`.

---

## Exécution

```bash
python3 convert_to_docx.py
```

Exemples :

```bash
python3 convert_to_docx.py --on-exists overwrite
python3 convert_to_docx.py --on-exists suffix
python3 convert_to_docx.py --soffice /usr/bin/soffice
```

Sorties :

```
docx/
convert_report.xlsx
```

---

## Récapitulatif des conversions gérées

| Format d'entrée | Traitement | Méthode | Sortie |
|------------------|------------|----------|---------|
| `.doc`           | Converti   | LibreOffice (soffice) | `.docx` |
| `.pdf`           | Converti   | pdf2docx | `.docx` |
| `.docx`          | Copié tel quel | — | `.docx` |

---

# 🧭 6. Pipeline complet (ordre recommandé)

```bash
python3 clean_extension.py
python3 dedupe.py
python3 convert_to_docx.py
```

---

# 📊 7. Fichiers Excel générés

| Étape | Fichier | Contenu |
|-------|---------|----------|
| Nettoyage | `inventaire_raw.xlsx` | action appliquée à chaque fichier brut |
| Dédoublonnage | `dedupe_report.xlsx` | décision, raison, chemin source/destination |
| Conversion | `convert_report.xlsx` | action (converti/copied/ignored), message, fichier généré |

---

# ⭐ Bonnes pratiques

- Toujours exécuter le pipeline **dans l’ordre** : Clean → Dedupe → Convert  
- Ne jamais modifier manuellement `clean_extension/` ou `dedupe/`  
- Laisser l’option `--on-exists skip` sauf besoin explicite  
- Les suffixes anti-collision garantissent **aucune perte de fichier**  
- Chaque étape laisse une **traçabilité complète en Excel**

---

# 🧩 Résultat attendu

À la fin du pipeline :

- Tous les fichiers non pertinents ont été exclus  
- Les doublons sont résolus selon les règles métier  
- Tous les documents sont au même format `.docx`  
- Vous disposez d’une traçabilité complète pour audit ou archivage  

Le pipeline produit un corpus documentaire propre, homogène et exploitable immédiatement.
