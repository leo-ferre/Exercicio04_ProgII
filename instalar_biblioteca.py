#!/usr/bin/env python3
"""
Script de instalação automática da biblioteca openpyxl
EXERCÍCIO 04 - Programação II
"""

import subprocess
import sys

def instalar_openpyxl():
    """
    Instala a biblioteca openpyxl usando diferentes métodos.
    """
    print("="*60)
    print("INSTALADOR DE BIBLIOTECA - openpyxl")
    print("="*60)
    print()

    metodos = [
        # Método 1: pip install com --user (mais compatível)
        [sys.executable, "-m", "pip", "install", "--user", "openpyxl"],
        # Método 2: pip install usando sys.executable
        [sys.executable, "-m", "pip", "install", "openpyxl"],
        # Método 3: pip install com --break-system-packages (para macOS recente)
        [sys.executable, "-m", "pip", "install", "--break-system-packages", "openpyxl"],
        # Método 4: /usr/bin/python3 com --user (para PyCharm)
        ["/usr/bin/python3", "-m", "pip", "install", "--user", "openpyxl"],
    ]

    for i, metodo in enumerate(metodos, 1):
        print(f"Tentativa {i}: {' '.join(metodo[2:])}")
        try:
            resultado = subprocess.run(
                metodo,
                capture_output=True,
                text=True,
                timeout=120
            )

            if resultado.returncode == 0:
                print("✓ Instalação bem-sucedida!")
                print()
                verificar_instalacao()
                return True
            else:
                print(f"✗ Falhou. Tentando próximo método...")
                print()
        except Exception as e:
            print(f"✗ Erro: {e}")
            print()
            continue

    print("="*60)
    print("⚠️  AVISO: Não foi possível instalar automaticamente")
    print("="*60)
    print()
    print("SOLUÇÕES ALTERNATIVAS:")
    print()
    print("1. Usar ambiente virtual (RECOMENDADO):")
    print("   python3 -m venv venv")
    print("   source venv/bin/activate")
    print("   pip install openpyxl")
    print()
    print("2. Tentar manualmente:")
    print("   python3 -m pip install openpyxl")
    print()
    print("3. O código já está pronto! Você pode:")
    print("   - Usar o PyCharm para instalar (Tools > Python Packages)")
    print("   - Rodar o código assim mesmo e instalar quando solicitado")
    print()
    return False

def verificar_instalacao():
    """
    Verifica se a biblioteca foi instalada corretamente.
    """
    print("="*60)
    print("VERIFICANDO INSTALAÇÃO")
    print("="*60)
    print()

    try:
        import openpyxl
        print("✓ openpyxl instalado com sucesso!")
        print(f"✓ Versão: {openpyxl.__version__}")
        print()
        print("="*60)
        print("✓ TUDO PRONTO! Você já pode executar o programa:")
        print("  python3 main.py")
        print("="*60)
        return True
    except ImportError:
        print("✗ openpyxl ainda não está disponível")
        return False

if __name__ == "__main__":
    print()
    instalar_openpyxl()
    print()

