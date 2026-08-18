"""Cargador de Recomendaciones Oficiales por Entidad.

Lee el archivo de recomendaciones que Función Pública entrega a CADA
entidad después de la medición (formato verificado con el caso real de
ALCALDÍA DE SAN RAFAEL: hoja "Recomendaciones", columnas "Política",
"Recomendación", "Política relacionada" con el código POL##).

Este módulo NO inventa recomendaciones para entidades que no tengan su
archivo cargado: si no hay archivo, el diagnóstico simplemente no muestra
recomendaciones (se avisa explícitamente en la interfaz).
"""
from __future__ import annotations

import gzip
import json
import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

PATRON_CODIGO_POLITICA = re.compile(r"^\s*(POL\d+)\b", re.IGNORECASE)


def cargar_consolidado(ruta_o_archivo) -> dict:
    """
    Carga la base consolidada generada por consolidar_recomendaciones.py.
    Acepta ruta local o archivo subido por Streamlit, en .json o .json.gz.
    """
    if hasattr(ruta_o_archivo, "read"):
        datos = ruta_o_archivo.read()
        nombre = getattr(ruta_o_archivo, "name", "")
        if nombre.endswith(".gz"):
            datos = gzip.decompress(datos)
        if isinstance(datos, bytes):
            datos = datos.decode("utf-8")
        return json.loads(datos)

    ruta = Path(ruta_o_archivo)
    if ruta.suffix == ".gz":
        with gzip.open(ruta, "rt", encoding="utf-8") as f:
            return json.load(f)
    return json.loads(ruta.read_text(encoding="utf-8"))


def recomendaciones_de_entidad(consolidado: dict, nombre_entidad: str) -> list["Recomendacion"]:
    """Busca una entidad (por nombre, insensible a mayúsculas/espacios) en el consolidado."""
    clave = nombre_entidad.strip().upper()
    registro = consolidado.get(clave)
    if registro is None:
        return []
    return [
        Recomendacion(
            politica_nombre=r["politica_nombre"],
            codigo_politica=r["codigo_politica"],
            texto=r["texto"],
        )
        for r in registro["recomendaciones"]
    ]


@dataclass
class Recomendacion:
    politica_nombre: str
    codigo_politica: str
    texto: str


def cargar_recomendaciones(ruta_o_archivo, nombre_hoja: str = "Recomendaciones") -> list[Recomendacion]:
    """
    Lee un archivo de recomendaciones oficial. Detecta automáticamente en
    qué fila empieza la tabla real (el archivo trae varias filas de
    encabezado institucional antes de la tabla), buscando la fila que
    contiene literalmente "Política" y "Recomendación".
    """
    crudo = pd.read_excel(ruta_o_archivo, sheet_name=nombre_hoja, header=None)

    fila_encabezado = None
    for i in range(min(20, len(crudo))):
        valores = [str(v).strip() for v in crudo.iloc[i].tolist()]
        if "Política" in valores and "Recomendación" in valores:
            fila_encabezado = i
            break

    if fila_encabezado is None:
        raise ValueError(
            "No se encontró la fila de encabezado ('Política', 'Recomendación') "
            "en las primeras 20 filas del archivo. ¿Es el formato oficial de Función Pública?"
        )

    df = pd.read_excel(ruta_o_archivo, sheet_name=nombre_hoja, header=fila_encabezado)
    df = df.dropna(subset=["Política", "Recomendación"])

    recomendaciones = []
    for _, fila in df.iterrows():
        codigo = ""
        if "Política relacionada" in df.columns and pd.notna(fila.get("Política relacionada")):
            m = PATRON_CODIGO_POLITICA.match(str(fila["Política relacionada"]))
            if m:
                codigo = m.group(1).upper()
        recomendaciones.append(
            Recomendacion(
                politica_nombre=str(fila["Política"]).strip(),
                codigo_politica=codigo,
                texto=str(fila["Recomendación"]).strip(),
            )
        )
    return recomendaciones


def recomendaciones_por_politica(recomendaciones: list[Recomendacion]) -> dict[str, list[Recomendacion]]:
    """Agrupa recomendaciones por código de política, para cruzarlas con las Brechas del diagnóstico."""
    agrupado: dict[str, list[Recomendacion]] = {}
    for r in recomendaciones:
        if not r.codigo_politica:
            continue
        agrupado.setdefault(r.codigo_politica, []).append(r)
    return agrupado


def cruzar_brechas_con_recomendaciones(brechas, recomendaciones: list[Recomendacion]) -> dict[str, list[str]]:
    """
    Dado el listado de Brecha (del Motor de Diagnóstico) y las Recomendacion
    cargadas para esa MISMA entidad, retorna un mapa
    codigo_indice_o_politica -> [textos de recomendación relacionados].
    """
    agrupado = recomendaciones_por_politica(recomendaciones)
    resultado = {}
    for b in brechas:
        textos = [r.texto for r in agrupado.get(b.codigo_politica, [])]
        if textos:
            resultado[b.codigo_indice] = textos
    return resultado
