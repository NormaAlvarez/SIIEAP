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
    aplica_mipg_integral: bool = True
    regimen_especial: str | None = None


def _nivel_riesgo(promedio: float | None) -> str:
    if promedio is None:
        return "Sin información"
    if promedio < 50:
        return "Alta"
    if promedio < 75:
        return "Media"
    return "Baja"


def _es_politica_control_interno(nombre_politica: str) -> bool:
    """Identifica la política de Control Interno por NOMBRE, no por código
    numérico — el número de política ha cambiado entre versiones del Manual
    Operativo del MIPG (era la política 15 en la v5/2023, es la política 19
    en la v6.1/2026), así que no es un identificador estable."""
    return "control interno" in nombre_politica.strip().lower()


def diagnosticar(entidad: Entidad, umbral_brecha: float = UMBRAL_BRECHA) -> DiagnosticoInstitucional:
    catalogo_dim = dimensiones()
    resultados_dim: list[ResultadoDimension] = []
    brechas: list[Brecha] = []

    aplica_integral = entidad.aplica_mipg_integral()
    # Entidades de régimen especial (universidades autónomas, órganos de
    # control, Concejos/Asambleas, Banco de la República, Corporaciones
    # Autónomas Regionales) solo están obligadas a la política de Control
    # Interno (MECI) — art. 40 Ley 489/1998 y art. 2.2.22.3.4 Decreto
    # 1499/2017. Si reportan datos en otras políticas de forma voluntaria,
    # esos datos SÍ se muestran y SÍ entran al promedio de su dimensión
    # (es información real, no hay razón para ocultarla), pero NO se
    # marcan como "brecha" — no tiene sentido decirle a un Concejo o a una
    # Personería que "debe corregir" algo que la norma no le exige.

    for cod_dim, dim in catalogo_dim.items():
        puntajes_dim: list[float] = []
        indices_esperados = 0

        for cod_pol, pol in dim["politicas"].items():
            politica_aplica = aplica_integral or _es_politica_control_interno(pol["nombre"])

            if not pol["indices"]:
                # Política sin índices propios (p.ej. POL14): se usa su
                # puntaje directo, si la entidad lo reportó.
                indices_esperados += 1
                directa = entidad.resultado_politica_directa_de(cod_pol)
                if directa is None:
                    continue
                puntajes_dim.append(directa.puntaje)
                if politica_aplica and directa.puntaje < umbral_brecha:
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
                if politica_aplica and resultado.puntaje < umbral_brecha:
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

            # Respaldo: si el catálogo espera índices propios para esta
            # política pero la entidad no reportó NINGUNO de ellos, se usa
            # el agregado por política que sí publicó Función Pública
            # (columna "POLxx Índice..." del archivo oficial). Es el caso
            # típico de entidades MECI-only (Concejos, Personerías,
            # Contralorías): Función Pública nunca les publica el
            # desglose índice por índice, solo el agregado de la política.
            indices_reportados_en_esta_politica = any(
                entidad.resultado_de(cod_idx) is not None for cod_idx in pol["indices"]
            )
            if not indices_reportados_en_esta_politica:
                directa = entidad.resultado_politica_directa_de(cod_pol)
                if directa is not None:
                    puntajes_dim.append(directa.puntaje)
                    if politica_aplica and directa.puntaje < umbral_brecha:
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
        aplica_mipg_integral=aplica_integral,
        regimen_especial=entidad.regimen_especial,
    )
