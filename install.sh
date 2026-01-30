#!/usr/bin/env bash

set -e

echo "=== Configuration des proxies APT ==="
echo -e 'Acquire::http::Proxy "http://10.246.42.30:8080";\nAcquire::https::Proxy "http://10.246.42.30:8080";' \
    | tee /etc/apt/apt.conf.d/95proxies

echo "=== Mise à jour APT ==="
apt-get update -y

echo "=== Installation de LibreOffice ==="
apt-get install -y libreoffice

echo "=== Installation de wget et bzip2 pour Miniconda ==="
apt-get install -y wget bzip2

echo "=== Téléchargement de Miniconda ==="
wget https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh -O /tmp/miniconda.sh

echo "=== Installation de Miniconda dans /opt/miniconda ==="
bash /tmp/miniconda.sh -b -p /opt/miniconda

echo "=== Ajout de conda au PATH ==="
export PATH="/opt/miniconda/bin:$PATH"

echo "=== Initialisation de conda ==="
source /opt/miniconda/etc/profile.d/conda.sh

echo "=== Mise à jour conda ==="
conda update -y conda

echo "=== Création de l'environnement Python 3.13 ==="
conda create -y -n pipeline python=3.13

echo "=== Activation de l'environnement ==="
conda activate pipeline

echo "=== Installation des dépendances Python ==="
pip install -r requirements.txt

echo "=== Installation terminée ==="
echo "Active ton environnement avec :   conda activate pipeline"
echo "Tu peux maintenant exécuter clean_extension.py, dedupe.py et convert_to_docx.py"
