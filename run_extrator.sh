#!/bin/bash
# Atualiza o ponto.html com a planilha do Cartão Ponto.
# Uso: ./run_extrator.sh ponto_ate_m_s_10.xls

cd "$(dirname "$0")"

if [ -f .venv/bin/activate ]; then
    source .venv/bin/activate
fi

if [ -z "$1" ]; then
    echo "Uso: ./run_extrator.sh <planilha.xls>"
    exit 1
fi

python3 extrator.py "$1" --html ponto.html
