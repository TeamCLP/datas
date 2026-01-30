#!/usr/bin/env bash
# ============================================================
# Script d'installation complet pour le pipeline documentaire
# Derrière proxy CA-GIP – Installation Miniconda + conda env
# Puis terminal placé dans le bon dossier + bon venv
# ============================================================

set -euo pipefail

### --- CONFIG PROXY ---
HTTP_PROXY_URL="http://10.246.42.30:8080"
HTTPS_PROXY_URL="http://10.246.42.30:8080"
NO_PROXY_LIST="localhost,127.0.0.1,::1,*.local"

### --- EXPORT PROXY ENV ---
export http_proxy="$HTTP_PROXY_URL"
export https_proxy="$HTTPS_PROXY_URL"
export no_proxy="$NO_PROXY_LIST"
export HTTP_PROXY="$HTTP_PROXY_URL"
export HTTPS_PROXY="$HTTPS_PROXY_URL"
export NO_PROXY="$NO_PROXY_LIST"

echo "=== Configuration des proxies APT ==="
echo -e "Acquire::http::Proxy \"${HTTP_PROXY_URL}\";\nAcquire::https::Proxy \"${HTTPS_PROXY_URL}\";" \
  | tee /etc/apt/apt.conf.d/95proxies

echo "=== Mise à jour APT ==="
apt-get update -y

echo "=== Installation LibreOffice ==="
DEBIAN_FRONTEND=noninteractive apt-get install -y libreoffice

echo "=== Installation outils requis (wget, bzip2) ==="
apt-get install -y wget bzip2

### --- TELECHARGEMENT MINICONDA ---
echo "=== Téléchargement Miniconda via proxy ==="
wget -O /tmp/miniconda.sh https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh \
 || wget -e use_proxy=yes -e http_proxy="$HTTP_PROXY_URL" -e https_proxy="$HTTPS_PROXY_URL" \
      -O /tmp/miniconda.sh https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh

echo "=== Installation Miniconda dans /opt/miniconda ==="
bash /tmp/miniconda.sh -b -p /opt/miniconda

### --- CHARGEMENT CONDA ---
echo "=== Ajout conda au PATH ==="
export PATH="/opt/miniconda/bin:$PATH"

echo "=== Activation du système conda ==="
source /opt/miniconda/etc/profile.d/conda.sh

### --- PROXY POUR CONDA ---
echo "=== Configuration du proxy pour conda (~/.condarc) ==="
cat > ~/.condarc <<EOF
proxy_servers:
  http: ${HTTP_PROXY_URL}
  https: ${HTTPS_PROXY_URL}
EOF

### --- MISE A JOUR CONDA ---
echo "=== Mise à jour conda ==="
conda update -n base -c defaults -y conda

### --- CREATION ENV PYTHON 3.13 ---
echo "=== Création environnement conda : pipeline (Python 3.13) ==="
conda create -y -n pipeline python=3.13

echo "=== Activation de l'environnement pipeline ==="
conda activate pipeline

### --- INSTALLATION REQUIREMENTS ---
echo "=== Installation des dépendances Python ==="
if [ -f "requirements.txt" ]; then
    pip install -r requirements.txt
else
    echo "❗ requirements.txt introuvable !"
fi

### --- FINALISATION ---
echo "=== Installation terminée avec succès ==="
echo "➡️  Terminal prêt à l’emploi : conda actif + bon dossier"

# On se place dans /home/quentin/datas (au cas où l'utilisateur a lancé depuis ailleurs)
cd /home/quentin/datas

echo "📌 Vous êtes maintenant dans : $(pwd)"
echo "📌 Environnement conda actif : $(conda env list | grep '*' | awk '{print $1}')"
echo
echo "Vous pouvez exécuter :"
echo "   python3 clean_extension.py"
echo "   python3 dedupe.py"
echo "   python3 convert_to_docx.py"
echo
echo "🎯 Votre terminal est PRÊT."
