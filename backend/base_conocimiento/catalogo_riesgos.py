"""Carga del catálogo de referencia del Módulo de Administración y Gestión del
Riesgo (SIIEAP).

Este archivo es NUEVO y NO modifica backend/base_conocimiento/catalogo.py ni
catalogo_idi.json (catálogo IDI-MIPG v6.1). Sigue el mismo patrón: un JSON de
datos + funciones de acceso de solo lectura.

Fuente principal: Guía para la Gestión Integral del Riesgo en Entidades
Públicas, Versión 7 (DAFP, agosto de 2025).
Fuente ambiental: Manual Operativo MIPG, Versión 7 (DAFP, agosto de 2026),
numeral 4.10 — Política de Gestión Ambiental Institucional.
"""
from __future__ import annotations

import json
from pathlib import Path
from functools import lru_cache

RUTA_CATALOGO_RIESGOS = Path(__file__).resolve().parent / "catalogo_riesgos.json"


@lru_cache(maxsize=1)
def cargar_catalogo_riesgos() -> dict:
    """Carga (una sola vez, cacheado) el catálogo de riesgos desde JSON."""
    with open(RUTA_CATALOGO_RIESGOS, encoding="utf-8") as f:
        return json.load(f)


def factores_riesgo() -> dict:
    """Devuelve el diccionario de factores de riesgo (Tabla 2, Guía v7)."""
    return cargar_catalogo_riesgos()["factores_riesgo"]


def opciones_factor_riesgo() -> list[str]:
    """Lista lista para poblar un st.selectbox: 'FR01 — Ejecución...'."""
    return [f'{cod} — {info["nombre"]}' for cod, info in factores_riesgo().items()]


def tabla_probabilidad() -> list[dict]:
    return cargar_catalogo_riesgos()["tabla_probabilidad"]["niveles"]


def tabla_impacto() -> list[dict]:
    return cargar_catalogo_riesgos()["tabla_impacto"]["niveles"]


def tabla_valoracion_controles(tipologia: str = "Gestión") -> dict:
    """Devuelve la tabla de valoración de controles que corresponde según la
    tipología del riesgo. Instrucción de la docente/experta temática
    (19-sep-2026): NO unificar los porcentajes entre fuentes — la tipología
    'Seguridad de la información' usa el peso del Anexo 5 (Manual = 10%);
    todas las demás tipologías usan el peso de la Guía v7 / Anexo 1
    (Manual = 15%)."""
    catalogo = cargar_catalogo_riesgos()
    if tipologia == "Seguridad de la información":
        return catalogo["tabla_valoracion_controles_seguridad_informacion"]
    return catalogo["tabla_valoracion_controles_general"]


def tipologias_riesgo() -> list[dict]:
    return cargar_catalogo_riesgos()["tipologias_riesgo"]


def opciones_tipologia_riesgo() -> list[str]:
    return [t["nombre"] for t in tipologias_riesgo()]


def preguntas_orientadoras_riesgo_fiscal() -> list[dict]:
    """Tabla 10 de la Guía v7 (preguntas guía, complementarias al catálogo)."""
    return cargar_catalogo_riesgos()["catalogo_puntos_riesgo_fiscal"][
        "preguntas_orientadoras_tabla_10"
    ]


def catalogo_puntos_riesgo_fiscal_disponible() -> bool:
    """True desde que se cargó el Anexo 3 real (Catálogo Indicativo de Puntos
    de Riesgo Fiscal, 50 puntos)."""
    estado = cargar_catalogo_riesgos()["catalogo_puntos_riesgo_fiscal"]["_estado"]
    return estado.startswith("DISPONIBLE")


def puntos_riesgo_fiscal() -> list[dict]:
    """Los 50 puntos de riesgo fiscal del Anexo 3 (DAFP), cada uno con su
    circunstancia inmediata asociada."""
    return cargar_catalogo_riesgos()["catalogo_puntos_riesgo_fiscal"]["puntos"]


def opciones_punto_riesgo_fiscal() -> list[str]:
    """Lista para un st.selectbox: 'Id — Punto de riesgo fiscal'."""
    return [f'{p["id"]} — {p["punto_riesgo_fiscal"]}' for p in puntos_riesgo_fiscal()]


def circunstancia_inmediata_de(id_punto) -> str | None:
    """Devuelve la circunstancia inmediata asociada a un id del catálogo
    fiscal, para autocompletar el campo una vez el usuario elige el punto."""
    for p in puntos_riesgo_fiscal():
        if str(p["id"]) == str(id_punto):
            return p["circunstancia_inmediata"]
    return None


def matriz_calor_severidad() -> dict:
    """Matriz de calor oficial (Anexo 1, DAFP): filas = probabilidad,
    columnas = impacto, valor = severidad (Bajo/Moderado/Alto/Extremo)."""
    return cargar_catalogo_riesgos()["matriz_calor_severidad"]


def tratamiento_riesgo_opciones() -> list[str]:
    return cargar_catalogo_riesgos()["tratamiento_riesgo"]["opciones"]


def campos_matriz_seguridad_informacion() -> list[str]:
    return cargar_catalogo_riesgos()["campos_matriz_seguridad_informacion"]["campos"]


def politica_ambiental_mipg_v7() -> dict:
    return cargar_catalogo_riesgos()["politica_ambiental_mipg_v7"]


def instrumentos_insumo_ambientales() -> list[str]:
    return politica_ambiental_mipg_v7()["instrumentos_insumo"]


def tipos_proceso_carepa() -> list[dict]:
    """Los 4 tipos generales de proceso (Decreto 092/2021 de Carepa), como
    referencia normativa. Para el listado nominal de los 19 procesos reales
    ya caracterizados, ver procesos_confirmados_carepa()."""
    return cargar_catalogo_riesgos()["procesos_carepa"]["tipos"]


def opciones_proceso_carepa() -> list[str]:
    return [f'{p["codigo"]} — {p["nombre"]}' for p in tipos_proceso_carepa()]


def procesos_confirmados_carepa() -> list[dict]:
    """Los 19 procesos institucionales de la Alcaldía de Carepa, con su
    caracterización completa (objetivo, alcance, entradas, salidas,
    responsable), tomados del Mapa de Procesos oficial (código C-ADM-xx-01)
    entregado por la docente/experta temática el 19-sep-2026. Reemplaza el
    marcador temporal de los 4 tipos generales para la identificación de
    riesgos por proceso."""
    return cargar_catalogo_riesgos()["procesos_carepa"].get("procesos_confirmados", [])


def procesos_confirmados_disponibles() -> bool:
    return len(procesos_confirmados_carepa()) > 0


def esquema_caracterizacion_proceso() -> dict:
    """Estructura de campos para caracterizar un proceso individual (más
    allá de los 4 tipos generales), agregada a petición de la docente/
    experta temática (validación 19-sep-2026)."""
    return cargar_catalogo_riesgos()["procesos_carepa"]["esquema_caracterizacion"]


def glosario_secciones() -> list[dict]:
    """Las 5 secciones del Anexo 2 (Glosario oficial de la Guía v7)."""
    return cargar_catalogo_riesgos()["glosario"]["secciones"]


def buscar_termino_glosario(texto: str) -> list[dict]:
    """Busca (sin distinguir mayúsculas) un texto en los términos del
    glosario oficial, incluyendo subtipos anidados."""
    texto = texto.lower().strip()
    encontrados = []
    for seccion in glosario_secciones():
        for t in seccion["terminos"]:
            if texto in t["termino"].lower() or texto in t["definicion"].lower():
                encontrados.append({**t, "_seccion": seccion["titulo"]})
            for sub in t.get("subtipos", []):
                if texto in sub["termino"].lower() or texto in sub["definicion"].lower():
                    encontrados.append({**sub, "_seccion": seccion["titulo"], "_termino_padre": t["termino"]})
    return encontrados


def modelo_madurez_componentes() -> list[dict]:
    """Los 5 componentes COSO-ERM del Anexo 4, cada uno con sus puntos de
    reflexión oficiales (77 en total)."""
    return cargar_catalogo_riesgos()["modelo_madurez_erm"]["componentes"]


def escala_madurez_erm() -> list[dict]:
    return cargar_catalogo_riesgos()["modelo_madurez_erm"]["escala_madurez"]
