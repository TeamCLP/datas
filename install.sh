#!/usr/bin/env bash
set -euo pipefail

############################################################
# 0) DÉTECTION DU DOSSIER DU SCRIPT  (important !)
############################################################
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"
echo "[OK] Script exécuté depuis : $SCRIPT_DIR"

############################################################
# 1) CONFIGURATION DES PROXIES ENVIRONNEMENT
############################################################
HTTP_PROXY_URL="http://10.246.42.30:8080"
HTTPS_PROXY_URL="http://10.246.42.30:8080"
NO_PROXY_LIST="localhost,127.0.0.1,::1,*.local"

export http_proxy="$HTTP_PROXY_URL"
export https_proxy="$HTTPS_PROXY_URL"
export no_proxy="$NO_PROXY_LIST"
export HTTP_PROXY="$HTTP_PROXY_URL"
export HTTPS_PROXY="$HTTPS_PROXY_URL"
export NO_PROXY="$NO_PROXY_LIST"
export PATH="/opt/miniconda/bin:$PATH"

echo "[OK] Proxy exports appliqués."

############################################################
# 2) PROXY APT + UPDATE
############################################################
echo "[OK] Configuration proxy APT"
echo -e "Acquire::http::Proxy \"${HTTP_PROXY_URL}\";\nAcquire::https::Proxy \"${HTTPS_PROXY_URL}\";" \
 > /etc/apt/apt.conf.d/95proxies

apt-get update -y

############################################################
# 3) INSTALL LIBREOFFICE + OUTILS
############################################################
DEBIAN_FRONTEND=noninteractive apt-get install -y libreoffice wget bzip2

############################################################
# 4) INSTALLATION MINICONDA
############################################################
echo "Téléchargement Miniconda…"
wget -O /tmp/miniconda.sh https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh \
 || wget -e use_proxy=yes -e http_proxy="$HTTP_PROXY_URL" -e https_proxy="$HTTPS_PROXY_URL" \
         -O /tmp/miniconda.sh https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh

echo "Installation Miniconda dans /opt/miniconda"
bash /tmp/miniconda.sh -b -p /opt/miniconda

############################################################
# 5) ACTIVATION CONDA
############################################################
export PATH="/opt/miniconda/bin:$PATH"
source /opt/miniconda/etc/profile.d/conda.sh

############################################################
# 6) PROXY POUR CONDA
############################################################
cat > ~/.condarc <<EOF
proxy_servers:
  http: ${HTTP_PROXY_URL}
  https: ${HTTPS_PROXY_URL}
EOF

############################################################
# 7) ACCEPTATION TOS ANACONDA AUTO
############################################################
conda tos accept --override-channels --channel https://repo.anaconda.com/pkgs/main || true
conda tos accept --override-channels --channel https://repo.anaconda.com/pkgs/r || true

############################################################
# 8) UPDATE CONDA
############################################################
conda update -n base -c defaults -y conda

############################################################
# 9) ENV PYTHON 3.13
############################################################
echo "[OK] Création environnement pipeline"
conda create -y -n pipeline python=3.13

echo "[OK] Activation environnement pipeline"
conda activate pipeline

############################################################
# 10) INSTALLATION REQUIREMENTS.TXT
############################################################
REQ_FILE="$SCRIPT_DIR/requirements.txt"

echo "[OK] Installation des dépendances depuis : $REQ_FILE"
if [[ -f "$REQ_FILE" ]]; then
    pip install -r "$REQ_FILE"
else
    echo "⚠️ requirements.txt introuvable dans $SCRIPT_DIR"
fi

############################################################
# 11) FIN – TERMINAL PRÊT DANS LE BON DOSSIER + BON CONDA
############################################################
cd "$SCRIPT_DIR"
mkdir raw

echo
echo "============================================================"
echo " INSTALLATION TERMINÉE — ENVIRONNEMENT PRÊT"
echo "============================================================"
echo "Dossier courant            : $(pwd)"
echo "Environnement conda actif  : pipeline"
echo
echo "Vous pouvez maintenant exécuter :"
echo "   python clean_extension.py"
echo "   python dedupe.py"
echo "   python convert_to_docx.py"
echo "   python classify_docx.py"
echo "   python extract_docx_to_markdown.py"
echo "   python build_dataset_jsonl.py"
echo "============================================================"
