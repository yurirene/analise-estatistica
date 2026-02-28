#!/bin/bash
python3 -m venv venv
source venv/bin/activate

echo "Gerando base estatística institucional..."
python3 main.py
echo "Base pronta para IA em /saida/base_ia_estatistica.json"

echo "Gerando base estatística institucional..."
python3 script.py
echo "Base pronta para IA em /saida/base_ia_institucional.json"

deactivate