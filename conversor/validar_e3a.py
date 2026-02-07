"""
Wrapper - importa o validador da raiz do projecto.
O ficheiro principal é: ../validar_e3a.py
"""
import sys
import os

# Adicionar raiz do projecto ao path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from validar_e3a import validate_e3a, main

if __name__ == '__main__':
    main()
