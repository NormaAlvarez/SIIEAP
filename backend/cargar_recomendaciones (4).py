"""Subregiones de Antioquia y sus municipios.

Clasificación oficial de la Gobernación de Antioquia en 9 subregiones,
usada para agrupar alcaldías por cercanía geográfica y contexto regional
(complementaria a la tipología/categoría de tamaño del DNP, que se cruza
aparte).

Uso típico:
    from backend.base_conocimiento.subregiones_antioquia import subregion_de

    subregion_de("Rionegro")  -> "Oriente"
    subregion_de("rionegro")  -> "Oriente"  (no distingue mayúsculas/tildes)
    subregion_de("Municipio inexistente") -> None
"""
from __future__ import annotations

import unicodedata

SUBREGIONES_ANTIOQUIA: dict[str, list[str]] = {
    "Valle de Aburrá": [
        "Medellín", "Bello", "Itagüí", "Envigado", "Sabaneta",
        "La Estrella", "Caldas", "Copacabana", "Girardota", "Barbosa",
    ],
    "Oriente": [
        "Rionegro", "Marinilla", "La Ceja", "El Retiro", "El Carmen de Viboral",
        "La Unión", "El Santuario", "Guarne", "Sonsón", "Abejorral",
        "Concepción", "Alejandría", "San Vicente Ferrer", "Cocorná",
        "San Rafael", "San Carlos", "San Luis", "Granada", "Guatapé",
        "El Peñol", "Nariño", "Argelia", "San Francisco",
    ],
    "Suroeste": [
        "Andes", "Jardín", "Jericó", "Titiribí", "Amagá", "Fredonia",
        "Venecia", "Ciudad Bolívar", "Betania", "Concordia", "Salgar",
        "Betulia", "Urrao", "Caramanta", "Támesis", "Valparaíso",
        "Santa Bárbara", "Montebello", "Tarso", "Pueblorrico", "Hispania",
        "La Pintada", "Angelópolis",
    ],
    "Occidente": [
        "Santa Fe de Antioquia", "San Jerónimo", "Sopetrán", "Olaya",
        "Liborina", "Sabanalarga", "Buriticá", "Ebéjico", "Anzá",
        "Armenia", "Frontino", "Abriaquí", "Cañasgordas", "Dabeiba",
        "Giraldo", "Peque", "Uramita", "Heliconia", "Caicedo",
    ],
    "Norte": [
        "Yarumal", "Santa Rosa de Osos", "Don Matías", "Entrerríos",
        "San Pedro de los Milagros", "Belmira", "Angostura", "Campamento",
        "Carolina del Príncipe", "Gómez Plata", "Guadalupe",
        "San Andrés de Cuerquia", "Ituango", "San José de la Montaña",
        "Toledo", "Valdivia", "Briceño",
    ],
    "Nordeste": [
        "Segovia", "Remedios", "Yolombó", "Vegachí", "Yalí", "Amalfi",
        "Anorí", "Cisneros", "San Roque", "Santo Domingo",
    ],
    "Bajo Cauca": [
        "Caucasia", "El Bagre", "Zaragoza", "Nechí", "Tarazá", "Cáceres",
    ],
    "Magdalena Medio": [
        "Puerto Berrío", "Puerto Nare", "Puerto Triunfo", "Yondó",
        "Maceo", "Caracolí",
    ],
    "Urabá": [
        "Apartadó", "Turbo", "Chigorodó", "Carepa", "Necoclí", "Arboletes",
        "San Pedro de Urabá", "San Juan de Urabá", "Mutatá", "Murindó",
        "Vigía del Fuerte",
    ],
}


def _normalizar(texto: str) -> str:
    """Quita tildes/mayúsculas para poder comparar 'RIONEGRO' con 'Rionegro'."""
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


# Índice invertido municipio (normalizado) -> subregión, construido una sola vez
_INDICE_MUNICIPIO_SUBREGION: dict[str, str] = {
    _normalizar(municipio): subregion
    for subregion, municipios in SUBREGIONES_ANTIOQUIA.items()
    for municipio in municipios
}


def subregion_de(nombre_municipio: str) -> str | None:
    """Devuelve la subregión de Antioquia a la que pertenece un municipio.

    Devuelve None si el municipio no está en la lista (p.ej. porque no es
    de Antioquia, o porque el nombre viene mal escrito/con otro formato).
    """
    return _INDICE_MUNICIPIO_SUBREGION.get(_normalizar(nombre_municipio))


def municipios_de_subregion(subregion: str) -> list[str]:
    """Devuelve la lista de municipios de una subregión (busca sin distinguir mayúsculas)."""
    for nombre_subregion, municipios in SUBREGIONES_ANTIOQUIA.items():
        if _normalizar(nombre_subregion) == _normalizar(subregion):
            return municipios
    return []


def todas_las_subregiones() -> list[str]:
    return list(SUBREGIONES_ANTIOQUIA.keys())
