"""
config.py

Carga la configuracion desde el archivo .env (que NO se sube al repositorio).
Todos los scripts leen de aqui usuarios, claves, URLs y rutas, para que el
codigo no contenga ningun dato sensible. Ver .env.example.

Requisitos:
    pip install python-dotenv
"""

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")


def requerido(nombre: str) -> str:
    """Devuelve la variable de entorno 'nombre' o termina con un mensaje claro."""
    valor = os.environ.get(nombre, "").strip()
    if not valor:
        raise SystemExit(
            f"Falta la variable '{nombre}'. Copia .env.example como .env "
            f"y completa su valor."
        )
    return valor


def opcional(nombre: str, por_defecto: str = "") -> str:
    return os.environ.get(nombre, "").strip() or por_defecto
