"""Motor de Análisis 360.

Da la "foto agregada" (el estado del arte) de un conjunto de entidades
reales — no una sola entidad, sino su grupo de comparación — cruzando:

  - Los resultados oficiales reales (resultados_territorio.xlsx /
    resultados_nacion.xlsx de Función Pública: IDI, D1-D7 por entidad).
  - El "Grupo par" YA calculado por Función Pública en esos mismos
    archivos (columna "Grupo par": p.ej. "ALCALDÍA GRUPO 4"), que es la
    forma oficial de no comparar entidades de tamaño/naturaleza distinta.
  - Opcionalmente, la subregión de Antioquia (backend.base_conocimiento.
    subregiones_antioquia), para acotar la comparación a un contexto
    geográfico cercano además del grupo par oficial.

No inventa ninguna cifra: todo sale de las columnas reales del archivo.
"""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from backend.base_conocimiento.subregiones_antioquia import subregion_de

COLUMNAS_DIMENSION = [
    "D1 Talento Humano",
    "D2 Direcciona- miento Estratégico y Planeación",
    "D3 Gestión para Resultados con Valores",
    "D4 Evaluación de Resultados",
    "D5 Información y Comunicación",
    "D6 Gestión del Conocimiento",
    "D7 Control Interno",
]
COLUMNA_IDI = "Índice de Desempeño Institucional"


@dataclass
class ResultadoAnalisis360:
    filtro_descripcion: str
    n_entidades: int
    promedio_idi: float | None
    promedio_por_dimension: dict[str, float]
    top5: list[tuple[str, float]]       # (entidad, IDI) más altas
    bottom5: list[tuple[str, float]]    # (entidad, IDI) más bajas
    percentil_entidad_referencia: float | None  # si se dio una entidad de referencia
    idi_entidad_referencia: float | None


def _normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza espacios raros en nombres de columna (el Excel oficial trae
    saltos de línea y dobles espacios en algunos encabezados de dimensión)."""
    df = df.rename(columns=lambda c: " ".join(str(c).split()))
    return df


def cargar_resultados(ruta_o_archivo, nombre_hoja: str) -> pd.DataFrame:
    """Carga resultados_territorio.xlsx o resultados_nacion.xlsx.

    El archivo oficial trae 2 filas de título antes del encabezado real,
    por eso header=2 (fila 3, índice 2).
    """
    df = pd.read_excel(ruta_o_archivo, sheet_name=nombre_hoja, header=2)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = _normalizar_columnas(df)
    return df


def analizar_360(
    df: pd.DataFrame,
    departamento: str | None = None,
    subregion: str | None = None,
    grupo_par_contiene: str | None = None,
    entidad_referencia: str | None = None,
) -> ResultadoAnalisis360:
    """
    Filtra el DataFrame de resultados oficiales y calcula el agregado.

    - departamento: p.ej. "ANTIOQUIA" (coincide con la columna Departamento).
    - subregion: p.ej. "Oriente" (se calcula a partir de la columna Municipio;
      solo tiene sentido si departamento="ANTIOQUIA").
    - grupo_par_contiene: p.ej. "ALCALDÍA" para quedarse solo con alcaldías,
      o "ALCALDÍA GRUPO 4" para una categoría de tamaño específica.
    - entidad_referencia: nombre (o parte del nombre) de una entidad, tal
      como aparece en la columna Entidad, para ubicar su percentil dentro
      del grupo filtrado.
    """
    filtrado = df.copy()
    partes_filtro = []

    if departamento:
        filtrado = filtrado[filtrado["Departamento"].astype(str).str.upper() == departamento.upper()]
        partes_filtro.append(f"Departamento={departamento}")

    if subregion:
        filtrado = filtrado[
            filtrado["Municipio"].astype(str).map(lambda m: subregion_de(m) == subregion)
        ]
        partes_filtro.append(f"Subregión={subregion}")

    if grupo_par_contiene:
        filtrado = filtrado[
            filtrado["Grupo par"].astype(str).str.upper().str.contains(grupo_par_contiene.upper(), na=False)
        ]
        partes_filtro.append(f"Grupo par contiene '{grupo_par_contiene}'")

    filtrado_con_idi = filtrado.dropna(subset=[COLUMNA_IDI])
    n_entidades = len(filtrado_con_idi)

    promedio_idi = round(filtrado_con_idi[COLUMNA_IDI].mean(), 2) if n_entidades else None

    promedio_por_dimension = {}
    for col in COLUMNAS_DIMENSION:
        if col in filtrado_con_idi.columns:
            serie = filtrado_con_idi[col].dropna()
            if len(serie):
                promedio_por_dimension[col] = round(serie.mean(), 2)

    ordenado = filtrado_con_idi.sort_values(COLUMNA_IDI, ascending=False)
    top5 = list(zip(ordenado["Entidad"].head(5), ordenado[COLUMNA_IDI].head(5)))
    bottom5 = list(zip(ordenado["Entidad"].tail(5), ordenado[COLUMNA_IDI].tail(5)))

    percentil = None
    idi_referencia = None
    if entidad_referencia and n_entidades:
        coincidencias = filtrado_con_idi[
            filtrado_con_idi["Entidad"].astype(str).str.upper().str.contains(entidad_referencia.upper(), na=False)
        ]
        if len(coincidencias):
            idi_referencia = float(coincidencias.iloc[0][COLUMNA_IDI])
            percentil = round(
                (filtrado_con_idi[COLUMNA_IDI] <= idi_referencia).mean() * 100, 1
            )

    return ResultadoAnalisis360(
        filtro_descripcion=" | ".join(partes_filtro) if partes_filtro else "Todas las entidades",
        n_entidades=n_entidades,
        promedio_idi=promedio_idi,
        promedio_por_dimension=promedio_por_dimension,
        top5=[(str(n), round(float(v), 2)) for n, v in top5],
        bottom5=[(str(n), round(float(v), 2)) for n, v in bottom5],
        percentil_entidad_referencia=percentil,
        idi_entidad_referencia=idi_referencia,
    )
