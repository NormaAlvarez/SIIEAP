"""Motor del Índice Sintético de Valor Público Territorial (ISVPT).

Novedad metodológica de este informe: aplica a las 7 dimensiones REALES del
IDI-MIPG (las mismas que ya usa motor_analisis_360) la metodología de
construcción de índices sintéticos compuestos validada académicamente por
Vélez Tamayo, Ortiz-Muñoz y Cardona Montoya (2026), "Piloto para la
construcción de un índice sintético de desarrollo económico local para dos
municipios de Antioquia" (Administración & Desarrollo, 56(1), e-1352,
https://doi.org/10.22431/25005227.1352), quienes a su vez siguen los
lineamientos de la OCDE (2008, "Handbook on Constructing Composite
Indicators") para normalización min-max y agregación aritmética simple.

Diferencia con el IDI oficial: el IDI de Función Pública ya es en sí mismo
un índice compuesto, pero se calcula con la metodología propia de Función
Pública (ponderaciones internas por índice/política). El ISVPT que aquí se
construye NO reemplaza al IDI oficial ni compite con él: es un ejercicio
complementario que normaliza las 7 dimensiones de un GRUPO DE COMPARACIÓN
(ej. la subregión o el grupo par) entre 0 y 1 relativos a ESE grupo
específico, permitiendo ver qué tan cerca o lejos está una entidad de la
mejor y la peor observada en su propio contexto — algo que el IDI absoluto,
por sí solo, no muestra.

No inventa ninguna cifra: toda la base es la misma columna real de
resultados oficiales ya cargada (resultados_territorio.xlsx /
resultados_nacion.xlsx).
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from backend.motores.motor_analisis_360 import COLUMNAS_DIMENSION


NOTA_METODOLOGICA_ISVPT = (
    "El ISVPT normaliza (min-max, entre 0 y 1) cada una de las 7 dimensiones "
    "del IDI-MIPG DENTRO del grupo de comparación seleccionado (ej. la "
    "subregión o el grupo par oficial), y luego agrega los resultados con un "
    "promedio aritmético simple, siguiendo la misma metodología validada "
    "académicamente por Vélez Tamayo et al. (2026) para el ISDEL y las "
    "directrices de la OCDE (2008) para indicadores compuestos. A diferencia "
    "del IDI oficial (absoluto, calculado por Función Pública con su propia "
    "ponderación), el ISVPT es siempre RELATIVO al grupo de comparación "
    "elegido: un valor de 1.0 significa 'el mejor de este grupo específico', "
    "y 0.0 significa 'el más rezagado de este grupo específico' — no un "
    "estándar absoluto nacional."
)


@dataclass
class ResultadoISVPT:
    n_entidades: int
    subindices_por_dimension_entidad: dict  # {dimension: valor normalizado 0-1} de la entidad de referencia
    isvpt_entidad_referencia: float | None
    ranking: list  # [(entidad, isvpt), ...] ordenado de mayor a menor
    posicion_entidad_referencia: int | None  # 1 = el mejor del grupo
    nota_metodologica: str = field(default=NOTA_METODOLOGICA_ISVPT)

    @property
    def nota_valor_extremo(self) -> str | None:
        """Aclaración específica para cuando el ISVPT (o alguno de sus
        subíndices) sale en 0.0 o 1.0 exactos — el caso que más se presta a
        malinterpretarse como "la entidad no tiene nada bien" o "la entidad
        no tiene nada mal", cuando en realidad solo describe su posición
        DENTRO del grupo de comparación elegido. Devuelve None si el valor
        no es extremo (no hace falta la aclaración) o si no hay datos.

        Se calcula aquí, una sola vez, para que los tres informes (Técnico,
        Estudio de Caso, Ejecutivo) puedan mostrar exactamente la misma
        explicación en vez de tener el texto repetido y potencialmente
        desincronizado en cada motor de generación de documentos.
        """
        if self.isvpt_entidad_referencia is None or not self.subindices_por_dimension_entidad:
            return None

        dimensiones_en_minimo = [d for d, v in self.subindices_por_dimension_entidad.items() if v == 0.0]
        dimensiones_en_maximo = [d for d, v in self.subindices_por_dimension_entidad.items() if v == 1.0]

        if self.isvpt_entidad_referencia == 0.0:
            return (
                f"⚠️ Este 0.0 NO significa que la entidad no tenga nada bien — significa que, "
                f"de las {self.n_entidades} entidades de ESTE grupo de comparación específico, "
                f"esta entidad tuvo el puntaje más bajo en {'todas las dimensiones evaluadas' if len(dimensiones_en_minimo) == len(self.subindices_por_dimension_entidad) else 'las dimensiones señaladas'} "
                "(el mínimo de cada dimensión siempre se normaliza a 0.0, sin importar qué tan alto o "
                "bajo sea ese puntaje en términos absolutos). Es una posición RELATIVA a un grupo "
                "concreto, casi siempre compuesto por entidades con desempeño destacado — no una "
                "calificación de que la gestión de la entidad esté en cero. Consulte el IDI oficial "
                "(cifra absoluta, no relativa) para la magnitud real del desempeño, y considere repetir "
                "este ejercicio con un grupo de comparación más amplio si busca una lectura menos extrema."
            )
        if self.isvpt_entidad_referencia == 1.0:
            return (
                f"✔️ Este 1.0 significa que, de las {self.n_entidades} entidades de ESTE grupo de "
                "comparación específico, esta entidad tuvo el puntaje más alto — no que su gestión "
                "sea perfecta en términos absolutos (el máximo de cada dimensión siempre se normaliza "
                "a 1.0). Es una posición RELATIVA a ese grupo puntual: un grupo de comparación más "
                "exigente podría arrojar un resultado distinto. Consulte el IDI oficial (cifra "
                "absoluta) para la magnitud real del desempeño."
            )
        if dimensiones_en_minimo or dimensiones_en_maximo:
            partes = []
            if dimensiones_en_minimo:
                partes.append(
                    f"el puntaje más bajo del grupo en {len(dimensiones_en_minimo)} de "
                    f"{len(self.subindices_por_dimension_entidad)} dimensiones ({', '.join(dimensiones_en_minimo)})"
                )
            if dimensiones_en_maximo:
                partes.append(
                    f"el puntaje más alto del grupo en {len(dimensiones_en_maximo)} de "
                    f"{len(self.subindices_por_dimension_entidad)} dimensiones ({', '.join(dimensiones_en_maximo)})"
                )
            return (
                "⚠️ Nota sobre los subíndices en 0.0 y/o 1.0 de la tabla anterior: la entidad tuvo "
                + " y ".join(partes)
                + f" DENTRO de su grupo de comparación específico ({self.n_entidades} entidades) — no "
                "un puntaje absoluto de cero o perfecto. El mínimo y el máximo del grupo siempre se "
                "normalizan a 0.0 y 1.0 respectivamente, sin importar la magnitud real del puntaje."
            )
        return None


def _normalizar_minmax(serie: pd.Series) -> pd.Series:
    """Normaliza una serie entre 0 y 1. Si no hay variación (max == min),
    devuelve 0.5 para todos (dimensión sin poder discriminante en este grupo)."""
    minimo, maximo = serie.min(), serie.max()
    if pd.isna(minimo) or pd.isna(maximo) or maximo == minimo:
        return serie.apply(lambda _: 0.5)
    return (serie - minimo) / (maximo - minimo)


def calcular_isvpt(df_grupo: pd.DataFrame, entidad_referencia: str | None = None) -> ResultadoISVPT:
    """
    Calcula el ISVPT para cada entidad del grupo ya filtrado (df_grupo debe
    venir ya recortado al grupo de comparación deseado, ej. con el mismo
    filtro que se usó en motor_analisis_360.analizar_360()).
    """
    df = df_grupo.copy()
    columnas_presentes = [c for c in COLUMNAS_DIMENSION if c in df.columns]

    normalizadas = pd.DataFrame(index=df.index)
    for col in columnas_presentes:
        normalizadas[col] = _normalizar_minmax(df[col])

    df["_ISVPT"] = normalizadas.mean(axis=1, skipna=True)
    df_valido = df.dropna(subset=["_ISVPT"])

    ranking_df = df_valido[["Entidad", "_ISVPT"]].sort_values("_ISVPT", ascending=False)
    ranking = [(str(n), round(float(v), 4)) for n, v in zip(ranking_df["Entidad"], ranking_df["_ISVPT"])]

    isvpt_referencia = None
    subindices_referencia = {}
    posicion = None
    if entidad_referencia:
        coincidencias = df_valido[
            df_valido["Entidad"].astype(str).str.upper().str.contains(entidad_referencia.upper(), na=False)
        ]
        if len(coincidencias):
            idx_referencia = coincidencias.index[0]
            isvpt_referencia = round(float(df_valido.loc[idx_referencia, "_ISVPT"]), 4)
            for col in columnas_presentes:
                subindices_referencia[col] = round(float(normalizadas.loc[idx_referencia, col]), 4)
            nombre_exacto = str(df_valido.loc[idx_referencia, "Entidad"])
            for i, (nombre_r, _valor) in enumerate(ranking, start=1):
                if nombre_r == nombre_exacto:
                    posicion = i
                    break

    return ResultadoISVPT(
        n_entidades=len(ranking),
        subindices_por_dimension_entidad=subindices_referencia,
        isvpt_entidad_referencia=isvpt_referencia,
        ranking=ranking,
        posicion_entidad_referencia=posicion,
    )

