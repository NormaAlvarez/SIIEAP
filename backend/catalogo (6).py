"""Subregiones del Chocó y sus municipios.

Clasificación oficial (IGAC — Instituto Geográfico Agustín Codazzi, Diccionario
Geográfico; y DANE) en 5 subregiones, usada para agrupar alcaldías del Chocó
por cercanía geográfica y contexto regional — el mismo propósito que
subregiones_antioquia.py cumple para Antioquia, para que el sistema pueda
usarse también con entidades del Chocó (departamento del cual la CETAP
Antioquia de la ESAP recibe una alta proporción de estudiantes).

Fuente: IGAC — Diccionario Geográfico; Gobernación del Departamento del Chocó;
DANE (proyecciones de población). 30 municipios agrupados en 5 subregiones:
Atrato, Darién, Pacífico Norte, Pacífico Sur y San Juan.

Nota: el municipio de Belén de Bajirá (subregión Darién) tiene un litigio
territorial histórico entre Antioquia y Chocó; se incluye aquí bajo Darién
siguiendo la fuente IGAC, sin que ello prejuzgue el litigio.

Uso típico:
    from backend.base_conocimiento.subregiones_choco import subregion_de

    subregion_de("Quibdó")   -> "Atrato"
    subregion_de("Istmina")  -> "San Juan"
    subregion_de("Nuquí")    -> "Pacífico Norte"
"""
from __future__ import annotations

import unicodedata

SUBREGIONES_CHOCO: dict[str, list[str]] = {
    "Atrato": [
        "Quibdó", "Atrato", "Bagadó", "Bojayá", "El Carmen de Atrato",
        "Lloró", "Medio Atrato", "Río Quito",
    ],
    "Darién": [
        "Acandí", "El Carmen del Darién", "Riosucio", "Unguía", "Belén de Bajirá",
    ],
    "Pacífico Norte": [
        "Bahía Solano", "Juradó", "Nuquí",
    ],
    "Pacífico Sur": [
        "Alto Baudó", "Bajo Baudó", "El Litoral de San Juan", "Medio Baudó",
    ],
    "San Juan": [
        "Cértegui", "Condoto", "El Cantón de San Pablo", "Istmina",
        "Medio San Juan", "Nóvita", "Río Iró", "San José del Palmar",
        "Sipí", "Tadó", "Unión Panamericana",
    ],
}


def _normalizar(texto: str) -> str:
    """Quita tildes/mayúsculas para poder comparar 'QUIBDÓ' con 'Quibdó'."""
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


# Índice invertido municipio (normalizado) -> subregión, construido una sola vez
_INDICE_MUNICIPIO_SUBREGION: dict[str, str] = {
    _normalizar(municipio): subregion
    for subregion, municipios in SUBREGIONES_CHOCO.items()
    for municipio in municipios
}


def subregion_de(nombre_municipio: str) -> str | None:
    """Devuelve la subregión del Chocó a la que pertenece un municipio.

    Devuelve None si el municipio no está en la lista (p.ej. porque no es
    del Chocó, o porque el nombre viene mal escrito/con otro formato).
    """
    return _INDICE_MUNICIPIO_SUBREGION.get(_normalizar(nombre_municipio))


def municipios_de_subregion(subregion: str) -> list[str]:
    """Devuelve la lista de municipios de una subregión (busca sin distinguir mayúsculas)."""
    for nombre_subregion, municipios in SUBREGIONES_CHOCO.items():
        if _normalizar(nombre_subregion) == _normalizar(subregion):
            return municipios
    return []


def todas_las_subregiones() -> list[str]:
    return list(SUBREGIONES_CHOCO.keys())
