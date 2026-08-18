"""Cargador de Resultados Oficiales FURAG/MDI.

Lee los archivos reales que publica Función Pública cada año
(Resultados_vigXXXX_nacion.xlsx / _territorio.xlsx) y construye objetos
Entidad con sus ResultadoIndice, listos para pasar al Motor de Diagnóstico.

Estos archivos NO tienen los códigos de índice en columnas separadas: el
código (p.ej. "I01") viene como prefijo del nombre de columna
("I01 Calidad de la planeación..."). Este módulo extrae ese prefijo.

Este es el punto de validación central de SIIEAP: el IDI que calcula
nuestro Motor de Diagnóstico (backend/motores/motor_diagnostico.py) debe
coincidir con el IDI que ya publicó Función Pública para la misma entidad.
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from backend.modelos.entidades import Entidad, ResultadoIndice, ResultadoPolitica

PATRON_CODIGO_INDICE = re.compile(r"^\s*[iI](\d+)\b")
PATRON_CODIGO_POLITICA = re.compile(r"^\s*(POL\d+)\b", re.IGNORECASE)

# CORRECCIÓN (agosto 2026): antes de este cambio, el sistema nunca leía las
# columnas oficiales D1-D7 del archivo de resultados — solo el IDI general.
# Las 7 dimensiones se recalculaban internamente en motor_diagnostico.py
# promediando los índices I01-I67, un cálculo propio que casi nunca
# coincidía con el valor oficial que publica Función Pública (porque la
# metodología oficial de agregación no es un promedio simple). Este mapa
# permite leer las columnas oficiales D1-D7 tal como aparecen en los
# archivos "Resultados_vigXXXX_territorio.xlsx" / "_nacion.xlsx", para
# usarlas como el dato protagonista en los 3 informes, igual que ya se
# hacía con el IDI general.
COLUMNAS_DIMENSION_OFICIAL = {
    "D1": "D1 Talento Humano",
    "D2": "D2 Direcciona- miento Estratégico y Planeación",
    "D3": "D3  Gestión para Resultados con Valores",
    "D4": "D4 Evaluación de Resultados",
    "D5": "D5 Información y Comunicación",
    "D6": "D6 Gestión del Conocimiento",
    "D7": "D7 Control Interno",
}


def _extraer_dimensiones_oficiales(fila: "pd.Series") -> dict[str, float]:
    """Lee, para una fila de entidad, el valor OFICIAL de cada dimensión
    D1-D7 tal como lo publica Función Pública. Si la entidad no reportó
    una dimensión (p.ej. régimen MECI), simplemente no se incluye esa
    clave — nunca se inventa ni se interpola."""
    salida = {}
    for cod_dim, nombre_columna in COLUMNAS_DIMENSION_OFICIAL.items():
        if nombre_columna not in fila.index:
            continue
        valor = fila[nombre_columna]
        if pd.notna(valor):
            salida[cod_dim] = float(valor)
    return salida

COLUMNAS_IDENTIFICACION = [
    "Código Sigep",
    "Entidad",
    "Naturaleza Jurídica",
    "Departamento",
    "Municipio",
    "Grupo par",
    "Tipo Formulario",
    "Índice de Desempeño Institucional",
]


def _extraer_codigo_indice(nombre_columna: str) -> str | None:
    m = PATRON_CODIGO_INDICE.match(nombre_columna)
    if not m:
        return None
    return f"I{int(m.group(1)):02d}"  # normaliza i01 / I1 / i9 -> I01, I09


def cargar_hoja(ruta_excel: str | Path, nombre_hoja: str) -> pd.DataFrame:
    """Lee la hoja real con el encabezado en la fila correcta (fila 3 de Excel, índice 2)."""
    return pd.read_excel(ruta_excel, sheet_name=nombre_hoja, header=2)


def columnas_de_indice(df: pd.DataFrame) -> dict[str, str]:
    """Mapa nombre_columna_original -> código de índice normalizado (solo columnas I##)."""
    mapa = {}
    for col in df.columns:
        if not isinstance(col, str):
            continue
        codigo = _extraer_codigo_indice(col)
        if codigo:
            mapa[col] = codigo
    return mapa


def columnas_de_politica_directa(df: pd.DataFrame) -> dict[str, str]:
    """Mapa nombre_columna_original -> código de política, para TODAS las
    columnas de resumen por política (POL01 Índice... hasta POL19
    Índice...) presentes en el archivo — no solo las que el catálogo marca
    como "sin índices propios" (p.ej. POL14).

    Se guardan como RESPALDO: motor_diagnostico.py usa el detalle por
    índice cuando la entidad lo reportó, y solo recurre a este agregado
    por política cuando no encontró ningún índice individual reportado
    para esa política. Esto es indispensable para entidades de régimen
    especial (Concejos, Personerías, Contralorías) que solo están
    obligadas al MECI: Función Pública les publica el agregado por
    política (p.ej. "POL19 Índice de Control Interno"), pero nunca el
    desglose índice por índice que sí reporta una alcaldía o gobernación.
    Antes de este cambio, esa columna se descartaba precisamente para esas
    entidades porque el catálogo esperaba encontrar el detalle en otro
    lado — detalle que esa entidad nunca reporta."""
    mapa = {}
    for col in df.columns:
        if not isinstance(col, str):
            continue
        m = PATRON_CODIGO_POLITICA.match(col)
        if m:
            mapa[col] = m.group(1).upper()
    return mapa


def entidad_desde_fila(
    fila: pd.Series,
    mapa_columnas_indice: dict[str, str],
    vigencia: int,
    mapa_columnas_politica_directa: dict[str, str] | None = None,
) -> Entidad:
    resultados = []
    for col, codigo in mapa_columnas_indice.items():
        valor = fila[col]
        if pd.isna(valor):
            continue  # índice no aplicable a esta entidad: se omite, no se inventa
        resultados.append(ResultadoIndice(codigo_indice=codigo, puntaje=float(valor), vigencia=vigencia))

    resultados_politica = []
    for col, cod_pol in (mapa_columnas_politica_directa or {}).items():
        valor = fila[col]
        if pd.isna(valor):
            continue
        resultados_politica.append(
            ResultadoPolitica(codigo_politica=cod_pol, puntaje=float(valor), vigencia=vigencia)
        )

    return Entidad(
        nombre=str(fila.get("Entidad", "")).strip(),
        codigo_dane=str(fila.get("Código Sigep", "")).strip() or None,
        vigencia=vigencia,
        resultados=resultados,
        resultados_politica_directa=resultados_politica,
    )


def listar_entidades(df: pd.DataFrame) -> list[str]:
    """Lista ordenada de nombres de entidad presentes en la hoja, para un selector."""
    return sorted(df["Entidad"].dropna().astype(str).str.strip().unique().tolist())


def entidad_por_nombre_exacto(
    df: pd.DataFrame,
    nombre_exacto: str,
    vigencia: int = 2025,
) -> tuple[Entidad, float | None, str]:
    """
    Igual que buscar_entidad, pero recibe el DataFrame ya cargado y el
    nombre EXACTO de la entidad (tal como aparece en la columna 'Entidad'),
    pensado para usarse después de listar_entidades() en un selector.
    """
    mapa = columnas_de_indice(df)
    mapa_politica_directa = columnas_de_politica_directa(df)

    coincidencias = df[df["Entidad"].astype(str).str.strip() == nombre_exacto.strip()]
    if coincidencias.empty:
        raise ValueError(f"No se encontró la entidad '{nombre_exacto}'")

    fila = coincidencias.iloc[0]
    entidad = entidad_desde_fila(fila, mapa, vigencia, mapa_politica_directa)
    idi_oficial = (
        float(fila["Índice de Desempeño Institucional"])
        if pd.notna(fila["Índice de Desempeño Institucional"])
        else None
    )
    grupo_par = str(fila.get("Grupo par", "")).strip()
    dimensiones_oficiales = _extraer_dimensiones_oficiales(fila)
    return entidad, idi_oficial, grupo_par, dimensiones_oficiales


def buscar_entidad(
    ruta_excel: str | Path,
    nombre_hoja: str,
    texto_busqueda: str,
    vigencia: int = 2025,
) -> tuple[Entidad, float | None, str]:
    """
    Busca una entidad por coincidencia de texto en la columna 'Entidad' y
    retorna: (objeto Entidad lista para diagnosticar, IDI oficial publicado, grupo par)
    """
    df = cargar_hoja(ruta_excel, nombre_hoja)
    mapa = columnas_de_indice(df)
    mapa_politica_directa = columnas_de_politica_directa(df)

    coincidencias = df[df["Entidad"].astype(str).str.contains(texto_busqueda, case=False, na=False)]
    if coincidencias.empty:
        raise ValueError(f"No se encontró ninguna entidad que coincida con '{texto_busqueda}'")
    if len(coincidencias) > 1:
        nombres = coincidencias["Entidad"].tolist()
        raise ValueError(
            f"'{texto_busqueda}' coincide con {len(coincidencias)} entidades, sea más específico: {nombres}"
        )

    fila = coincidencias.iloc[0]
    entidad = entidad_desde_fila(fila, mapa, vigencia, mapa_politica_directa)
    idi_oficial = float(fila["Índice de Desempeño Institucional"]) if pd.notna(fila["Índice de Desempeño Institucional"]) else None
    grupo_par = str(fila.get("Grupo par", "")).strip()
    dimensiones_oficiales = _extraer_dimensiones_oficiales(fila)
    return entidad, idi_oficial, grupo_par, dimensiones_oficiales
