"""Acceso de lectura a la normativa vigente por política MIPG
(backend/base_conocimiento/normativa_politicas_mipg.json).

Origen y alcance de estos datos
--------------------------------
Este archivo consolida la normatividad de las 19 políticas de Gestión y
Desempeño de MIPG, a partir de los 19 documentos de política que la
entidad entregó (uno por política, con su respectiva normatividad ya
identificada por el equipo). Sobre esos 19 documentos se hizo, el 8 de
agosto de 2026, una verificación puntual de vigencia (qué ha cambiado en
los últimos ~3 años) y se aplicaron las siguientes correcciones y
adiciones, ya incorporadas en el JSON:

  - POL04 (Gestión Presupuestal): corregido el año de la Ley 2056 (2021 -> 2020).
  - POL05 (Compras y Contratación): se agregó la Ley 2195 de 2022 y el
    Decreto 874 de 2024.
  - POL07 (Gobierno Digital): se anota que el Decreto 1008 de 2018 fue
    subrogado por el Decreto 767 de 2022 en lo relativo a esta política.
  - POL08 (Seguridad Digital): se agregaron el CONPES 3995 de 2020 y el
    Decreto 338 de 2022.
  - POL09 (Defensa Jurídica): se agregaron la Ley 2294 de 2023 (crea el
    Sistema de Defensa Jurídica del Estado) y el Decreto 104 de 2025 (lo
    reglamenta).
  - POL10 (Mejora Normativa): se agregaron el CONPES 3816 de 2014
    (documento fundacional de la política) y el Decreto 385 de 2026.
  - POL13 (Participación Ciudadana): se agregó el Decreto 1535 de 2022.
  - POL14 (Seguimiento y Evaluación): se agregaron la Ley 2195 de 2022 y
    el Decreto 1122 de 2024 (transforma el PAAC en el PTEP).
  - POL16 (Gestión Documental): se agregó el Acuerdo 001 de 2024 del AGN.
  - POL17 (Gestión Estadística): se agregó la Ley 2335 de 2023 y se anota
    la derogatoria del art. 155 de la Ley 1955/2019 que ese artículo sufrió.
  - POL01 (Talento Humano) y POL02 (Integridad): ver adiciones en el JSON
    (Ley 2466/2025, Ley 2013/2019, Circular 100-004/2025).
  - POL03 (Planeación Institucional): se agregó la Ley 2294 de 2023 (PND
    2022-2026 vigente).

Pendiente (declarado explícitamente, no se debe asumir cubierto):
  - POL06 (Fortalecimiento de Procesos), POL11 (Servicio al Ciudadano),
    POL12 (Racionalización de Trámites), POL15 (Transparencia), POL18
    (Gestión del Conocimiento) y POL19 (Control Interno) fueron revisadas
    en la misma sesión sin que se confirmara un vacío normativo específico
    con las búsquedas realizadas ese día. Esto NO equivale a una
    verificación exhaustiva línea por línea de cada norma listada en ellas
    (la mayoría de sus normas de base, anteriores a 2019, no fueron
    verificadas una por una contra el texto oficial vigente).
  - Ninguna de las 19 políticas tuvo una verificación exhaustiva artículo
    por artículo; se priorizaron los cambios normativos de 2022 en
    adelante, que es la ventana que la propia entidad pidió revisar.

Corrección adicional del 21 de agosto de 2026 (a raíz de una Fe de Erratas
externa recibida sobre el documento técnico del SIIEAP)
-----------------------------------------------------------------------
Un reporte externo señaló 11 normas potencialmente desactualizadas dentro
de este mismo catálogo. Se verificaron directamente en fuente oficial
(SUIN-Juriscol y Gestor Normativo de Función Pública) y se corrigieron/
anotaron en el JSON, EN TODAS las políticas donde aparecían:

  - Ley 872 de 2003 (aparecía en las 19 políticas): anotada como derogada
    por el Decreto 1499 de 2017, art. 5 — que ya está presente en las 19
    listas, así que no se dejó ningún vacío normativo.
  - Ley 1753 de 2015 (aparecía en las 19 políticas): anotada — el Plan
    Nacional de Desarrollo vigente es la Ley 2294 de 2023; se conserva la
    cita porque su art. 133 es la base legal del Decreto 1499/2017.
  - POL07 (Gobierno Digital): "Resolución 3564 de 2015" anotada como
    derogada por la Resolución 1519 de 2020 (ya presente en la lista);
    "Decreto 2693 de 2008" ELIMINADO — no existe un decreto con ese
    número y año (la norma real de 2008 en esta materia, Decreto 1151 de
    2008, fue la que derogó el Decreto 2693, pero este es de 2012, no de
    2008; probable error de transcripción en el reporte original).
  - POL09 (Defensa Jurídica): "Ley 640 de 2001" anotada como derogada por
    la Ley 2220 de 2022 (agregada a la lista); "Decreto 1784 de 2019"
    anotado como derogado por el Decreto 2647 de 2022, art. 43 (agregado
    a la lista) — verificado: el Decreto 1784/2019 regulaba la estructura
    del Departamento Administrativo de la Presidencia de la República.
  - POL01 (Talento Humano): "Decreto 171 de 2016" anotado como de
    objetivo cumplido (solo modificó un plazo de transición del SG-SST,
    ya vencido); "Resolución 1111 de 2017" anotada como derogada por la
    Resolución 312 de 2019 (ya presente en la lista).
  - POL04 (Gestión Presupuestal): "Ley 1530 de 2012" anotada como
    derogada parcialmente por la Ley 2056 de 2020 (ya presente en la
    lista, con excepción de los arts. 106 a 126 y 128).
  - POL05 (Compras y Contratación): "Decreto 874 de 2024" — el reporte
    externo lo daba como derogado por la Ley 2294 de 2023, orden
    cronológico imposible (norma de 2024 no puede ser derogada por una de
    2023); se anotó la inconsistencia SIN eliminar la cita, ya que su
    objeto (Asociaciones Público-Populares) ya estaba correctamente
    descrito en este mismo archivo desde la corrección de agosto 8.
  - POL13 (Participación Ciudadana): "Decreto 1535 de 2022" — el reporte
    externo lo daba como derogado por el Decreto 1458 de 2025; no se pudo
    verificar en esta corrección, se anotó como pendiente sin eliminar la
    cita.

Pendiente declarado de esta ronda: 3 de las 11 normas reportadas (Decreto
1784 sin año explícito en el reporte original, y los dos casos anotados
como "pendiente de verificación" arriba) no alcanzaron verificación
completa contra fuente oficial. Esta corrección tampoco es una auditoría
exhaustiva de las 19 políticas: se limitó a las 11 normas puntuales del
reporte recibido.
"""
from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path

RUTA_NORMATIVA = Path(__file__).resolve().parent / "normativa_politicas_mipg.json"


@lru_cache(maxsize=1)
def cargar_normativa_politicas() -> dict:
    if not RUTA_NORMATIVA.exists():
        raise FileNotFoundError(f"No existe {RUTA_NORMATIVA}.")
    return json.loads(RUTA_NORMATIVA.read_text(encoding="utf-8"))


def normativa_de_politica(codigo_politica: str) -> dict | None:
    """Dado un código de política (p.ej. 'POL05'), retorna su nombre y la
    lista de normas que la sustentan."""
    return cargar_normativa_politicas().get(codigo_politica.strip().upper())


def normativa_de_dimension(codigo_dimension: str) -> dict[str, dict]:
    """Retorna {codigo_politica: {...}} de todas las políticas que
    pertenecen a una dimensión dada (usa el catálogo oficial para conocer
    qué políticas caen en cada dimensión)."""
    from backend.base_conocimiento.catalogo import dimensiones

    dims = dimensiones()
    dim = dims.get(codigo_dimension.strip().upper())
    if not dim:
        return {}
    normativa = cargar_normativa_politicas()
    return {
        cod: normativa[cod]
        for cod in dim["politicas"].keys()
        if cod in normativa
    }


def todas_las_politicas_con_normativa() -> dict[str, dict]:
    """Mapa completo código_política -> {nombre, normas}, en el orden
    POL01..POL19."""
    datos = cargar_normativa_politicas()
    return dict(sorted(datos.items(), key=lambda kv: int(kv[0][3:])))
