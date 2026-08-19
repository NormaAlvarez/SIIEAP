"""Motor de Análisis 360.

Da la "foto agregada" (el estado del arte) de un conjunto de entidades
reales — no una sola entidad, sino su grupo de comparación — cruzando:

  - Los resultados oficiales reales (resultados_territorio.xlsx /
    resultados_nacion.xlsx de Función Pública: IDI, D1-D7 por entidad).
  - El "Grupo par" YA calculado por Función Pública en esos mismos
    archivos (columna "Grupo par": p.ej. "ALCALDÍA GRUPO 4"), que es la
    forma oficial de no comparar entidades de tamaño/naturaleza distinta.
  - Opcionalmente, la subregión del departamento (Antioquia o Chocó, ver
    REGISTRO_SUBREGIONES_POR_DEPARTAMENTO más abajo), para acotar la
    comparación a un contexto geográfico cercano además del grupo par
    oficial. Agregar un nuevo departamento solo requiere crear su módulo
    de subregiones (mismo patrón que subregiones_antioquia.py) y
    registrarlo en ese diccionario.

No inventa ninguna cifra: todo sale de las columnas reales del archivo.
"""
from __future__ import annotations

import unicodedata
from dataclasses import dataclass

import pandas as pd

from backend.base_conocimiento.subregiones_antioquia import (
    subregion_de as _subregion_de_antioquia,
    todas_las_subregiones as _todas_las_subregiones_antioquia,
)
from backend.base_conocimiento.subregiones_choco import (
    subregion_de as _subregion_de_choco,
    todas_las_subregiones as _todas_las_subregiones_choco,
)
from backend.base_conocimiento.subregiones_santander import (
    subregion_de as _subregion_de_santander,
    todas_las_subregiones as _todas_las_subregiones_santander,
)

# Registro de departamentos con subregionalización disponible en el sistema.
# Para agregar un nuevo departamento: crear backend/base_conocimiento/
# subregiones_<departamento>.py con las mismas 4 funciones (SUBREGIONES_X,
# subregion_de, municipios_de_subregion, todas_las_subregiones) e incluirlo
# aquí con su nombre en mayúsculas tal como aparece en la columna
# "Departamento" del archivo oficial de Función Pública.
REGISTRO_SUBREGIONES_POR_DEPARTAMENTO = {
    "ANTIOQUIA": {
        "subregion_de": _subregion_de_antioquia,
        "todas_las_subregiones": _todas_las_subregiones_antioquia,
    },
    "CHOCO": {
        "subregion_de": _subregion_de_choco,
        "todas_las_subregiones": _todas_las_subregiones_choco,
    },
    "CHOCÓ": {
        "subregion_de": _subregion_de_choco,
        "todas_las_subregiones": _todas_las_subregiones_choco,
    },
    "SANTANDER": {
        "subregion_de": _subregion_de_santander,
        "todas_las_subregiones": _todas_las_subregiones_santander,
    },
}


def subregiones_disponibles_para(departamento: str | None) -> list[str]:
    """Lista de subregiones disponibles para un departamento dado, o lista
    vacía si el departamento no tiene subregionalización registrada todavía."""
    if not departamento:
        return []
    entrada = REGISTRO_SUBREGIONES_POR_DEPARTAMENTO.get(departamento.strip().upper())
    return entrada["todas_las_subregiones"]() if entrada else []


def _subregion_de_municipio(municipio: str, departamento: str | None) -> str | None:
    """Resuelve la subregión de un municipio usando el módulo del
    departamento correspondiente. Si el departamento no está registrado,
    devuelve None (no se puede filtrar por subregión, pero el resto del
    análisis 360 sigue funcionando por grupo par y por IDI)."""
    if not departamento:
        return None
    entrada = REGISTRO_SUBREGIONES_POR_DEPARTAMENTO.get(departamento.strip().upper())
    return entrada["subregion_de"](municipio) if entrada else None


COLUMNAS_DIMENSION = [
    "D1 Talento Humano",
    "D2 Direcciona- miento Estratégico y Planeación",
    # CORRECCIÓN (agosto 2026, hallazgo incidental durante el arreglo del
    # régimen especial): el nombre real de esta columna en el archivo
    # oficial de Función Pública tiene DOBLE espacio entre "D3" y
    # "Gestión" ("D3  Gestión para Resultados con Valores"). Con un solo
    # espacio, el filtro `col in df.columns` de motor_isvpt.py y de
    # analizar_360() nunca encontraba la columna, así que D3 se excluía en
    # silencio del promedio de grupo de comparación y del ISVPT — para
    # TODAS las entidades, no solo las de régimen especial. Los valores
    # D1-D7 protagonistas que se muestran en los informes NO se vieron
    # afectados (vienen de un módulo distinto, cargar_resultados_oficiales.py,
    # que ya tenía el espacio correcto) — este bug solo afectaba el
    # promedio de grupo/ISVPT, un análisis secundario.
    "D3  Gestión para Resultados con Valores",
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


def _normalizar_texto_libre(texto: str) -> str:
    """Quita tildes/mayúsculas para comparar texto libre sin que un acento
    faltante (p.ej. escribir 'ALCALDIA' en vez de 'ALCALDÍA' al filtrar por
    Grupo par) haga que el filtro no encuentre ninguna entidad en silencio."""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def filtrar_grupo(
    df: pd.DataFrame,
    departamento: str | None = None,
    subregion: str | None = None,
    grupo_par_contiene: str | None = None,
) -> tuple[pd.DataFrame, str]:
    """
    Aplica los mismos filtros de comparación (departamento, subregión, grupo
    par) que usa analizar_360(), y devuelve el DataFrame filtrado junto con
    la descripción del filtro aplicado. Reutilizable por otros motores
    (ej. motor_isvpt) que necesiten operar sobre el mismo grupo exacto.
    """
    filtrado = df.copy()
    partes_filtro = []

    if departamento:
        filtrado = filtrado[
            filtrado["Departamento"].astype(str).map(_normalizar_texto_libre) == _normalizar_texto_libre(departamento)
        ]
        partes_filtro.append(f"Departamento={departamento}")

    if subregion:
        filtrado = filtrado[
            filtrado["Municipio"].astype(str).map(lambda m: _subregion_de_municipio(m, departamento) == subregion)
        ]
        partes_filtro.append(f"Subregión={subregion}")

    if grupo_par_contiene:
        filtrado = filtrado[
            filtrado["Grupo par"].astype(str).map(_normalizar_texto_libre).str.contains(
                _normalizar_texto_libre(grupo_par_contiene), na=False
            )
        ]
        partes_filtro.append(f"Grupo par contiene '{grupo_par_contiene}'")

    descripcion = " | ".join(partes_filtro) if partes_filtro else "Todas las entidades"
    return filtrado, descripcion


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
    filtrado, _descripcion = filtrar_grupo(df, departamento, subregion, grupo_par_contiene)
    partes_filtro = [_descripcion] if _descripcion != "Todas las entidades" else []

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
