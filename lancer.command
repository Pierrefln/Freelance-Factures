#!/bin/bash
# Double-cliquez sur ce fichier pour lancer le générateur de factures.

cd "$(dirname "$0")"

# Vérifie si reportlab est installé, sinon l'installe
python3 -c "import reportlab" 2>/dev/null || {
    echo "Installation de reportlab (une seule fois)..."
    pip3 install reportlab
}

python3 generateur_factures.py
