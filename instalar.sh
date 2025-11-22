#!/bin/bash
# Script de instalação para o Sistema de Controle de Estoque
# EXERCÍCIO 04 - Programação II

echo "=========================================="
echo "INSTALAÇÃO - Sistema de Controle de Estoque"
echo "=========================================="
echo ""

# Verificar se Python está instalado
if ! command -v python3 &> /dev/null
then
    echo "❌ Python 3 não está instalado!"
    echo "Por favor, instale Python 3.7 ou superior"
    exit 1
fi

echo "✓ Python 3 encontrado: $(python3 --version)"
echo ""

# Instalar dependências
echo "Instalando biblioteca openpyxl..."
pip3 install --break-system-packages openpyxl || pip3 install openpyxl

if [ $? -eq 0 ]; then
    echo "✓ Biblioteca instalada com sucesso!"
else
    echo "❌ Erro ao instalar a biblioteca"
    exit 1
fi

echo ""
echo "=========================================="
echo "Instalação concluída com sucesso!"
echo "=========================================="
echo ""
echo "Para executar o programa, use:"
echo "  python3 main.py"
echo ""

