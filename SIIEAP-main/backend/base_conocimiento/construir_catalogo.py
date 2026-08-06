"""
Construye el catálogo jerárquico oficial del IDI-MIPG a partir del Excel
publicado por Función Pública (Tabla de índices MDI, vigencia 2025).

Jerarquía real que existe en la fuente (orden secuencial del archivo):
    IDI (C)
      Dimensión (D1..D7)
        Política (POL01..POL19)
          Índice (I01.. / i01..)

El archivo de origen no trae la jerarquía en columnas separadas: se infiere
por el ORDEN de aparición y el prefijo del código. Este script es la única
fuente de verdad para regenerar catalogo_idi.json si Función Pública publica
una nueva versión del Excel.

Uso:
    python construir_catalogo.py
"""
import json
import re
from pathlib import Path

import pandas as pd

RUTA_EXCEL = Path(__file__).resolve().parents[2] / "data" / "catalogo_idi_2025.xlsx"
RUTA_SALIDA = Path(__file__).resolve().parent / "catalogo_idi.json"


def clasificar_codigo(codigo: str) -> str:
    codigo = str(codigo).strip()
    if codigo.upper() == codigo and codigo.startswith("D"):
        return "dimension"
    if codigo.upper().startswith("POL"):
        return "politica"
    if codigo.upper().startswith("I"):
        return "indice"
    if codigo.upper() == "IDI":
        return "idi"
    return "desconocido"


def construir_catalogo() -> dict:
    if not RUTA_EXCEL.exists():
        raise FileNotFoundError(
            f"No se encontró el Excel fuente en {RUTA_EXCEL}. "
            "Copie allí el archivo oficial de Función Pública."
        )

    df = pd.read_excel(RUTA_EXCEL, sheet_name="Ind2025", header=None, skiprows=2)
    df.columns = ["_col0", "nombre", "codigo", "descripcion"]
    df = df[["nombre", "codigo", "descripcion"]].dropna(subset=["nombre", "codigo"])

    catalogo = {
        "idi": None,
        "dimensiones": {},   # codigo -> {...}
    }

    dimension_actual = None
    politica_actual = None

    for _, fila in df.iterrows():
        codigo = str(fila["codigo"]).strip()
        nombre = str(fila["nombre"]).strip()
        descripcion = str(fila["descripcion"]).strip() if pd.notna(fila["descripcion"]) else ""
        tipo = clasificar_codigo(codigo)

        if tipo == "idi":
            catalogo["idi"] = {"nombre": nombre, "codigo": codigo, "descripcion": descripcion}

        elif tipo == "dimension":
            dimension_actual = codigo
            politica_actual = None
            catalogo["dimensiones"][codigo] = {
                "nombre": nombre,
                "codigo": codigo,
                "descripcion": descripcion,
                "politicas": {},
            }

        elif tipo == "politica":
            if dimension_actual is None:
                continue  # dato mal formado en la fuente, se ignora explícitamente
            politica_actual = codigo
            catalogo["dimensiones"][dimension_actual]["politicas"][codigo] = {
                "nombre": nombre,
                "codigo": codigo,
                "descripcion": descripcion,
                "indices": {},
            }

        elif tipo == "indice":
            if dimension_actual is None or politica_actual is None:
                continue
            codigo_norm = codigo.upper()
            catalogo["dimensiones"][dimension_actual]["politicas"][politica_actual]["indices"][codigo_norm] = {
                "nombre": nombre,
                "codigo": codigo_norm,
                "descripcion": descripcion,
            }

    _normalizar_codigo_dimension_d4(catalogo)
    _agregar_politicas_no_decompuestas(catalogo)
    return catalogo


def _normalizar_codigo_dimension_d4(catalogo: dict) -> None:
    """
    La Tabla de Índices fuente trae la dimensión 4 con el código "D04",
    pero los archivos reales de resultados FURAG (columnas de
    Resultados_vigXXXX_nacion/territorio.xlsx) usan "D4". Se normaliza a
    "D4" para que ambas fuentes crucen por el mismo código.
    """
    dims = catalogo["dimensiones"]
    if "D04" in dims and "D4" not in dims:
        dims["D4"] = dims.pop("D04")
        dims["D4"]["codigo"] = "D4"


def _agregar_politicas_no_decompuestas(catalogo: dict) -> None:
    """
    POL14 ("Índice de Seguimiento y Evaluación del Desempeño Institucional")
    existe oficialmente y aparece con valores reales en los archivos de
    resultados de Función Pública, pero NO aparece como fila en la Tabla de
    Índices (este Excel), porque no se descompone en índices individuales
    (es un agregado propio, no una suma de i##). Sin este parche, la
    Dimensión D4 (Evaluación de Resultados) queda vacía y el catálogo
    queda con 18 políticas en vez de las 19 oficiales.

    Por su nombre y function ("Seguimiento y Evaluación del Desempeño
    Institucional"), POL14 pertenece a D4 (Evaluación de Resultados), que
    es justamente la dimensión que la Tabla de Índices fuente deja vacía.
    """
    if "D4" not in catalogo["dimensiones"]:
        return  # estructura inesperada: no forzar el parche
    politicas_d4 = catalogo["dimensiones"]["D4"]["politicas"]
    if "POL14" not in politicas_d4:
        politicas_d4["POL14"] = {
            "nombre": "Índice de Seguimiento y Evaluación del Desempeño Institucional",
            "codigo": "POL14",
            "descripcion": (
                "Política sin índices propios en la Tabla de Índices oficial; "
                "se reporta como agregado directo en los archivos de resultados FURAG."
            ),
            "indices": {},
        }


def resumen(catalogo: dict) -> str:
    n_dim = len(catalogo["dimensiones"])
    n_pol = sum(len(d["politicas"]) for d in catalogo["dimensiones"].values())
    n_ind = sum(
        len(p["indices"])
        for d in catalogo["dimensiones"].values()
        for p in d["politicas"].values()
    )
    return f"IDI construido con {n_dim} dimensiones, {n_pol} políticas, {n_ind} índices."


if __name__ == "__main__":
    catalogo = construir_catalogo()
    RUTA_SALIDA.write_text(
        json.dumps(catalogo, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(resumen(catalogo))
    print(f"Guardado en: {RUTA_SALIDA}")
