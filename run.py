#!/usr/bin/env python3
"""
Launcher para CORREAGRO Extractor PDF
Ejecutar desde la raíz del proyecto:
    python run.py
"""
import sys
import os

# Asegurar que el directorio raíz esté en el path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from main import main

if __name__ == "__main__":
    main()
