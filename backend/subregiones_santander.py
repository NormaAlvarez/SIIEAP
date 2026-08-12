"""Subregiones (provincias) de Santander y sus municipios.

Clasificación oficial en 7 provincias (Gobernación de Santander:
santander.gov.co; ver también DANE y Wikipedia "Organización territorial
de Santander"), usada para agrupar alcaldías por cercanía geográfica y
contexto regional (complementaria a la tipología/categoría de tamaño del
DNP y al "Grupo par" de Función Pública, que se cruzan aparte).

Nota de trazabilidad: a diferencia de Antioquia (donde la Gobernación
mantiene una única fuente digital clara de las 9 subregiones), la
clasificación de Santander viene de varias fuentes públicas (Gobernación
de Santander, DANE, fuentes académicas) que no siempre coinciden en el
detalle de 2-3 municipios fronterizos entre provincias vecinas (el caso
más citado es Cepitá, ubicado aquí en García Rovira siguiendo el conteo
oficial de "13 municipios" de esa provincia). El total de 87 municipios
usado aquí sí es consistente entre fuentes. Si su Facultad/ESAP maneja
una fuente distinta para algún municipio límite, avise para corregir.

Caso especial — Tona: por decisión propia del municipio, Tona pertenece
simultáneamente a la provincia Metropolitana y a la provincia de Soto
Norte. Este módulo, para mantener una función `subregion_de` que
devuelve un único valor, lo asigna a "Metropolitana" (su vinculación
funcional más fuerte, por el Área Metropolitana de Bucaramanga).

Uso típico:
    from backend.base_conocimiento.subregiones_santander import subregion_de

    subregion_de("Sabana de Torres")  -> "Yariguíes"
    subregion_de("sabana de torres")  -> "Yariguíes"  (no distingue mayúsculas/tildes)
    subregion_de("Municipio inexistente") -> None
"""
from __future__ import annotations

import unicodedata

SUBREGIONES_SANTANDER: dict[str, list[str]] = {
    "Metropolitana": [
        "Bucaramanga", "Floridablanca", "Girón", "Piedecuesta",
        "El Playón", "Lebrija", "Los Santos", "Rionegro",
        "Santa Bárbara", "Zapatoca", "Tona",
    ],
    "Soto Norte": [
        "California", "Charta", "Matanza", "Suratá", "Vetas",
    ],
    "Guanentá": [
        "San Gil", "Aratoca", "Barichara", "Cabrera", "Charalá",
        "Coromoro", "Curití", "Encino", "Jordán Sube", "Mogotes",
        "Ocamonte", "Onzaga", "Páramo", "Pinchote", "San Joaquín",
        "Valle de San José", "Villanueva",
    ],
    "Comunera": [
        "El Socorro", "Chima", "Confines", "Contratación",
        "El Guacamayo", "Galán", "Gámbita", "Guadalupe", "Guapotá",
        "Hato", "Oiba", "Palmar", "Palmas del Socorro",
        "Santa Helena del Opón", "Simacota", "Suaita",
    ],
    "García Rovira": [
        "Málaga", "Capitanejo", "Carcasí", "Cepitá", "Cerrito",
        "Concepción", "Enciso", "Guaca", "Macaravita", "Molagavita",
        "San Andrés", "San José de Miranda", "San Miguel",
    ],
    "Vélez": [
        "Vélez", "Aguada", "Albania", "Barbosa", "Bolívar", "Chipatá",
        "Cimitarra", "El Peñón", "Florián", "Guavatá", "Güepsa",
        "Jesús María", "La Belleza", "La Paz", "Landázuri",
        "Puente Nacional", "Puerto Parra", "San Benito", "Sucre",
    ],
    "Yariguíes": [
        "Barrancabermeja", "Betulia", "El Carmen de Chucurí",
        "Puerto Wilches", "Sabana de Torres", "San Vicente de Chucurí",
    ],
}


def _normalizar(texto: str) -> str:
    """Quita tildes/mayúsculas para poder comparar 'BUCARAMANGA' con 'Bucaramanga'."""
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


# Índice invertido municipio (normalizado) -> subregión, construido una sola vez
_INDICE_MUNICIPIO_SUBREGION: dict[str, str] = {
    _normalizar(municipio): subregion
    for subregion, municipios in SUBREGIONES_SANTANDER.items()
    for municipio in municipios
}


def subregion_de(nombre_municipio: str) -> str | None:
    """Devuelve la provincia de Santander a la que pertenece un municipio.

    Devuelve None si el municipio no está en la lista (p.ej. porque no es
    de Santander, o porque el nombre viene mal escrito/con otro formato).
    """
    return _INDICE_MUNICIPIO_SUBREGION.get(_normalizar(nombre_municipio))


def municipios_de_subregion(subregion: str) -> list[str]:
    """Devuelve la lista de municipios de una provincia (busca sin distinguir mayúsculas)."""
    for nombre_subregion, municipios in SUBREGIONES_SANTANDER.items():
        if _normalizar(nombre_subregion) == _normalizar(subregion):
            return municipios
    return []


def todas_las_subregiones() -> list[str]:
    return list(SUBREGIONES_SANTANDER.keys())
