#!/usr/bin/env bash
set -euo pipefail

### ============================
### 0. CONFIG PROXY
### ============================
HTTP_PROXY_URL="http://10.246.42.30:8080"
HTTPS_PROXY_URL="http://10.246.42.30:8080"
NO_PROXY_LIST="localhost,127.0.0.1,::1,*.local"

export http_proxy="$HTTP_PROXY_URL"
export https_proxy="$HTTPS_PROXY_URL"
export no_proxy="$NO_PROXY_LIST"
export HTTP_PROXY="$HTTP_PROXY_URL"
export HTTPS_PROXY="$HTTPS_PROXY_URL"
export NO_PROXY="$NO_PROXY_LIST"

echo "Proxy configuré."

### ============================
### 1. CONFIG APT
### ============================
echo "=== Configuration des proxies APT ==="
echo -e "Acquire::http::Proxy \"${HTTP_PROXY_URL}\";\nAcquire::https::Proxy \"${HTTPS_PROXY_URL}\";" \
  > /etc/apt/apt.conf.d/95proxies

echo "=== apt update ==="
apt-get update -y

### ============================
### 2. INSTALL LIBREOFFICE
### ============================
echo "=== Installation LibreOffice ==="
DEBIAN_FRONTEND=noninteractive apt-get install -y libreoffice

### ============================
### 3. INSTALL wget+bzip2
### ============================
echo "=== Installation wget + bzip2 ==="
apt-get install -y wget bzip2

### ============================
### 4. INSTALL MINICONDA (proxy OK)
### ============================
echo "=== Téléchargement Miniconda ==="
wget -O /tmp/miniconda.sh https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh \
 || wget -e use_proxy=yes -e http_proxy="$HTTP_PROXY_URL" -e https_proxy="$HTTPS_PROXY_URL" \
      -O /tmp/miniconda.sh https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh

echo "=== Installation Miniconda dans /opt/miniconda ==="
bash /tmp/miniconda.sh -b -p /opt/miniconda

### ============================
### 5. ACTIVER CONDA
### ============================
export PATH="/opt/miniconda/bin:$PATH"
source /opt/miniconda/etc/profile.d/conda.sh

### ============================
### 6. CONFIG PROXY POUR conda
### ============================
echo "=== Configuration proxy conda ==="
cat > ~/.condarc <<EOF
proxy_servers:
  http: ${HTTP_PROXY_URL}
  https: ${HTTPS_PROXY_URL}
EOF

### ============================
### 7. ACCEPTATION AUTOMATIQUE DES TOS ANACONDA
### ============================
echo "=== Acceptation auto des Terms of Service Anaconda ==="
conda tos accept --override-channels --channel https://repo.anaconda.com/pkgs/main || true
conda tos accept --override-channels --channel https://repo.anaconda.com/pkgs/r || true

### ============================
### 8. UPDATE CONDA
### ============================
echo "=== Mise à jour conda ==="
conda update -n base -c defaults -y conda

### ============================
### 9. CREER ENV PYTHON 3.13
### ============================
echo "=== Création environnement pipeline (Python 3.13) ==="
conda create -y -n pipeline python=3.13

echo "=== Activation environnement pipeline ==="
conda activate pipeline

### ============================
### 10. INSTALL REQUIREMENTS
### ============================
echo "=== Installation requirements.txt ==="
if [[ -f "requirements.txt" ]]; then
    pip install -r requirements.txt
else
    echo "⚠️ requirements.txt introuvable, étape ignorée"
fi

### ============================
### 11. FIN — TERMINAL PRET
### ============================
cd /home/datas

echo
echo "==========================================================="
echo "   INSTALLATION TERMINEE — ENVIRONNEMENT PRET"
echo "==========================================================="
echo "Dossier courant : $(pwd)"
echo "Environnement conda actif : pipeline"
echo "Vous pouvez exécuter :"
echo "   python3 clean_extension.py"
echo "   python3 dedupe.py"
echo "   python3 convert_to_docx.py"
echo
echo "🎯 L'environnement est opérationnel."
