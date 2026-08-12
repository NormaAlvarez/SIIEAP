"""Motor de Índice Sintético Comparativo.

Implementa la metodología descrita en:
Vélez Tamayo, Ortiz-Muñoz y Cardona Montoya (2026), "Piloto para la
construcción de un índice sintético de desarrollo económico local para
dos municipios de Antioquia", Administración & Desarrollo, 56(1).

Pasos (iguales a los del artículo):
  1. Normalización min-max de cada indicador a escala 0-1.
     - Indicadores de impacto positivo: (x - min) / (max - min)
     - Indicadores de impacto negativo: (max - x) / (max - min)
  2. Subíndice por dimensión: promedio aritmético simple de sus indicadores.
  3. Índice sintético final: promedio aritmético simple de los subíndices.

Esta implementación es agnóstica del dominio: sirve tanto para comparar
municipios (como en el artículo) como para comparar entidades públicas
en el IDI-MIPG, siempre que se le entreguen los datos en el formato
{entidad: {dimension: {indicador: valor}}} y se indique qué indicadores
son de impacto negativo.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass
class ResultadoIndiceSintetico:
    entidad: str
    subindices: dict[str, float]   # dimensión -> subíndice normalizado (0-1)
    indice_sintetico: float        # promedio de los subíndices (0-1)


def _normalizar_min_max(valores: dict[str, float], invertir: bool) -> dict[str, float]:
    numeros = list(valores.values())
    x_min, x_max = min(numeros), max(numeros)
    if x_max == x_min:
        # Sin variación entre entidades: no se puede discriminar, se asigna 0.5 a todas
        return {k: 0.5 for k in valores}
    normalizado = {}
    for entidad, x in valores.items():
        if invertir:
            normalizado[entidad] = (x_max - x) / (x_max - x_min)
        else:
            normalizado[entidad] = (x - x_min) / (x_max - x_min)
    return normalizado


def calcular_indice_sintetico(
    datos: dict[str, dict[str, dict[str, float]]],
    indicadores_negativos: set[str] | None = None,
) -> list[ResultadoIndiceSintetico]:
    """
    datos: {entidad: {dimension: {indicador: valor_crudo}}}
    indicadores_negativos: nombres de indicador cuyo valor alto es indeseable
                           (p.ej. tasa de desempleo, pobreza) -> se invierten.
    """
    indicadores_negativos = indicadores_negativos or set()
    entidades = list(datos.keys())
    if len(entidades) < 2:
        raise ValueError(
            "Se necesitan al menos 2 entidades para normalizar por min-max "
            "(la normalización es relativa al grupo comparado)."
        )

    dimensiones = list(next(iter(datos.values())).keys())

    # valores_normalizados[dimension][indicador][entidad] = valor 0-1
    valores_normalizados: dict[str, dict[str, dict[str, float]]] = {}

    for dimension in dimensiones:
        indicadores = list(datos[entidades[0]][dimension].keys())
        valores_normalizados[dimension] = {}
        for indicador in indicadores:
            crudos = {e: datos[e][dimension][indicador] for e in entidades}
            invertir = indicador in indicadores_negativos
            valores_normalizados[dimension][indicador] = _normalizar_min_max(crudos, invertir)

    resultados = []
    for entidad in entidades:
        subindices = {}
        for dimension in dimensiones:
            valores_dim = [
                valores_normalizados[dimension][indicador][entidad]
                for indicador in valores_normalizados[dimension]
            ]
            subindices[dimension] = round(sum(valores_dim) / len(valores_dim), 4)

        indice_final = round(sum(subindices.values()) / len(subindices), 4)
        resultados.append(
            ResultadoIndiceSintetico(
                entidad=entidad,
                subindices=subindices,
                indice_sintetico=indice_final,
            )
        )

    resultados.sort(key=lambda r: r.indice_sintetico, reverse=True)
    return resultados
