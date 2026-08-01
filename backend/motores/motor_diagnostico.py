"""Motor de Diagnóstico Institucional.

Toma los resultados oficiales de una entidad (ResultadoIndice por índice IDI)
y el catálogo oficial (dimensión -> política -> índice) y produce:

  - promedio por dimensión y por política
  - listado de brechas (índices por debajo de un umbral)
  - priorización simple (Alta/Media/Baja) por dimensión

No inventa datos: si la entidad no reportó un índice, ese índice
simplemente no participa en el promedio (se documenta como faltante).
"""
from __future__ import annotations

from dataclasses import dataclass

from backend.base_conocimiento.catalogo import dimensiones
from backend.modelos.entidades import Entidad

UMBRAL_BRECHA = 60.0  # por debajo de esto (escala 0-100 FURAG) se considera brecha


@dataclass
class ResultadoDimension:
    codigo: str
    nombre: str
    promedio: float | None
    n_indices_evaluados: int
    n_indices_esperados: int
    nivel_riesgo: str  # Alta / Media / Baja / Sin información


@dataclass
class Brecha:
    codigo_indice: str
    nombre_indice: str
    puntaje: float
    dimension: str
    politica: str
    codigo_politica: str


@dataclass
class DiagnosticoInstitucional:
    entidad: str
    vigencia: int
    resultados_por_dimension: list[ResultadoDimension]
    brechas: list[Brecha]
    idi_estimado: float | None  # promedio simple de dimensiones con información


def _nivel_riesgo(promedio: float | None) -> str:
    if promedio is None:
        return "Sin información"
    if promedio < 50:
        return "Alta"
    if promedio < 75:
        return "Media"
    return "Baja"


def diagnosticar(entidad: Entidad, umbral_brecha: float = UMBRAL_BRECHA) -> DiagnosticoInstitucional:
    catalogo_dim = dimensiones()
    resultados_dim: list[ResultadoDimension] = []
    brechas: list[Brecha] = []

    for cod_dim, dim in catalogo_dim.items():
        puntajes_dim: list[float] = []
        indices_esperados = 0

        for cod_pol, pol in dim["politicas"].items():
            if not pol["indices"]:
                # Política sin índices propios (p.ej. POL14): se usa su
                # puntaje directo, si la entidad lo reportó.
                indices_esperados += 1
                directa = entidad.resultado_politica_directa_de(cod_pol)
                if directa is None:
                    continue
                puntajes_dim.append(directa.puntaje)
                if directa.puntaje < umbral_brecha:
                    brechas.append(
                        Brecha(
                            codigo_indice=cod_pol,
                            nombre_indice=pol["nombre"],
                            puntaje=directa.puntaje,
                            dimension=dim["nombre"],
                            politica=pol["nombre"],
                            codigo_politica=cod_pol,
                        )
                    )
                continue

            for cod_idx, idx in pol["indices"].items():
                indices_esperados += 1
                resultado = entidad.resultado_de(cod_idx)
                if resultado is None:
                    continue
                puntajes_dim.append(resultado.puntaje)
                if resultado.puntaje < umbral_brecha:
                    brechas.append(
                        Brecha(
                            codigo_indice=cod_idx,
                            nombre_indice=idx["nombre"],
                            puntaje=resultado.puntaje,
                            dimension=dim["nombre"],
                            politica=pol["nombre"],
                            codigo_politica=cod_pol,
                        )
                    )

        promedio = sum(puntajes_dim) / len(puntajes_dim) if puntajes_dim else None
        resultados_dim.append(
            ResultadoDimension(
                codigo=cod_dim,
                nombre=dim["nombre"],
                promedio=round(promedio, 2) if promedio is not None else None,
                n_indices_evaluados=len(puntajes_dim),
                n_indices_esperados=indices_esperados,
                nivel_riesgo=_nivel_riesgo(promedio),
            )
        )

    promedios_validos = [r.promedio for r in resultados_dim if r.promedio is not None]
    idi_estimado = round(sum(promedios_validos) / len(promedios_validos), 2) if promedios_validos else None

    brechas.sort(key=lambda b: b.puntaje)  # las más críticas primero

    return DiagnosticoInstitucional(
        entidad=entidad.nombre,
        vigencia=entidad.vigencia,
        resultados_por_dimension=resultados_dim,
        brechas=brechas,
        idi_estimado=idi_estimado,
    )
