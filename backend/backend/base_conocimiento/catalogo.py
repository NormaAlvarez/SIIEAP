"""Acceso de lectura al catálogo oficial IDI (catalogo_idi.json).

Este módulo es deliberadamente simple: carga el JSON generado por
construir_catalogo.py y ofrece funciones de consulta. No duplica la lógica
de negocio de diagnóstico, que vive en backend/motores.
"""
from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path

RUTA_CATALOGO = Path(__file__).resolve().parent / "catalogo_idi.json"


@lru_cache(maxsize=1)
def cargar_catalogo() -> dict:
    if not RUTA_CATALOGO.exists():
        raise FileNotFoundError(
            f"No existe {RUTA_CATALOGO}. Ejecute construir_catalogo.py primero."
        )
    return json.loads(RUTA_CATALOGO.read_text(encoding="utf-8"))


def dimension_de_indice(codigo_indice: str) -> dict | None:
    """Dada la sigla de un índice (p.ej. 'I05'), retorna su dimensión y política."""
    codigo_indice = codigo_indice.strip().upper()
    catalogo = cargar_catalogo()
    for dim in catalogo["dimensiones"].values():
        for pol in dim["politicas"].values():
            if codigo_indice in pol["indices"]:
                return {
                    "dimension": {"codigo": dim["codigo"], "nombre": dim["nombre"]},
                    "politica": {"codigo": pol["codigo"], "nombre": pol["nombre"]},
                    "indice": pol["indices"][codigo_indice],
                }
    return None


def todos_los_indices() -> dict[str, dict]:
    """Mapa código_indice -> metadatos (incluye a qué política/dimensión pertenece)."""
    catalogo = cargar_catalogo()
    salida: dict[str, dict] = {}
    for dim in catalogo["dimensiones"].values():
        for pol in dim["politicas"].values():
            for cod, idx in pol["indices"].items():
                salida[cod] = {
                    **idx,
                    "dimension_codigo": dim["codigo"],
                    "dimension_nombre": dim["nombre"],
                    "politica_codigo": pol["codigo"],
                    "politica_nombre": pol["nombre"],
                }
    return salida


def dimensiones() -> dict[str, dict]:
    return cargar_catalogo()["dimensiones"]
