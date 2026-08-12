"""Tipologías municipales 2025 del DNP — Departamento de Antioquia.

Fuente oficial: DNP, "Tipologías de las Entidades Territoriales para el
Reconocimiento de Capacidades. Resultados para la Vigencia 2025.
Departamento: Antioquia" (anexo con código DANE, municipio y tipología).

La tipología agrupa municipios con niveles similares de capacidad fiscal
y administrativa, y de conectividad/densidad poblacional:
  - "Ciudades grandes": la capital, Medellín (caso único).
  - "SC- Centro aglomeración": Rionegro (caso único, por el Sistema de Ciudades).
  - "1": los 13 municipios con las capacidades más altas.
  - "2", "3", "4": niveles intermedios (altos a bajos).
  - "5": los 4 municipios con menor capacidad y mayor ruralidad.

Esta tabla se usa para NO comparar, por ejemplo, Medellín con un municipio
de tipología 5: la comparación "grupo par" debe cruzar tipología + región,
tal como exige la metodología de Función Pública/DNP.

Uso típico:
    from backend.base_conocimiento.tipologias_antioquia_2025 import tipologia_de

    tipologia_de("Rionegro")  -> "SC- Centro aglomeración"
    tipologia_de("montebello") -> "3"
"""
from __future__ import annotations

import unicodedata

# (código_dane, municipio, tipología_2025)
TIPOLOGIAS_ANTIOQUIA_2025: list[tuple[str, str, str]] = [
    ("05001", "MEDELLÍN", "Ciudades grandes"),
    ("05615", "RIONEGRO", "SC- Centro aglomeración"),
    ("05360", "ITAGÜÍ", "1"), ("05631", "SABANETA", "1"), ("05088", "BELLO", "1"),
    ("05266", "ENVIGADO", "1"), ("05380", "LA ESTRELLA", "1"), ("05212", "COPACABANA", "1"),
    ("05129", "CALDAS", "1"), ("05308", "GIRARDOTA", "1"), ("05440", "MARINILLA", "1"),
    ("05376", "LA CEJA", "1"), ("05318", "GUARNE", "1"), ("05697", "EL SANTUARIO", "1"),
    ("05607", "RETIRO", "1"),
    ("05030", "AMAGÁ", "2"), ("05190", "CISNEROS", "2"), ("05148", "EL CARMEN DE VIBORAL", "2"),
    ("05541", "PEÑOL", "2"), ("05045", "APARTADÓ", "2"), ("05400", "LA UNIÓN", "2"),
    ("05679", "SANTA BÁRBARA", "2"), ("05656", "SAN JERÓNIMO", "2"),
    ("05664", "SAN PEDRO DE LOS MILAGROS", "2"), ("05674", "SAN VICENTE", "2"),
    ("05237", "DON MATÍAS", "2"), ("05576", "PUEBLORRICO", "2"), ("05147", "CAREPA", "2"),
    ("05321", "GUATAPÉ", "2"), ("05861", "VENECIA", "2"), ("05101", "CIUDAD BOLÍVAR", "2"),
    ("05809", "TITIRIBÍ", "2"), ("05761", "SOPETRÁN", "2"), ("05209", "CONCORDIA", "2"),
    ("05306", "GIRALDO", "2"), ("05368", "JERICÓ", "2"), ("05154", "CAUCASIA", "2"),
    ("05686", "SANTA ROSA DE OSOS", "2"), ("05315", "GUADALUPE", "2"),
    ("05138", "CAÑASGORDAS", "2"), ("05206", "CONCEPCIÓN", "2"), ("05113", "BURITICÁ", "2"),
    ("05250", "EL BAGRE", "2"),
    ("05079", "BARBOSA", "3"), ("05390", "LA PINTADA", "3"), ("05282", "FREDONIA", "3"),
    ("05467", "MONTEBELLO", "3"), ("05034", "ANDES", "3"), ("05353", "HISPANIA", "3"),
    ("05036", "ANGELÓPOLIS", "3"), ("05197", "COCORNÁ", "3"), ("05172", "CHIGORODÓ", "3"),
    ("05042", "SANTAFÉ DE ANTIOQUIA", "3"), ("05642", "SALGAR", "3"), ("05364", "JARDÍN", "3"),
    ("05240", "EBÉJICO", "3"), ("05264", "ENTRERRIOS", "3"), ("05313", "GRANADA", "3"),
    ("05347", "HELICONIA", "3"), ("05887", "YARUMAL", "3"), ("05690", "SANTO DOMINGO", "3"),
    ("05670", "SAN ROQUE", "3"), ("05091", "BETANIA", "3"), ("05093", "BETULIA", "3"),
    ("05411", "LIBORINA", "3"), ("05002", "ABEJORRAL", "3"), ("05134", "CAMPAMENTO", "3"),
    ("05591", "PUERTO TRIUNFO", "3"), ("05501", "OLAYA", "3"), ("05819", "TOLEDO", "3"),
    ("05044", "ANZÁ", "3"), ("05310", "GÓMEZ PLATA", "3"), ("05038", "ANGOSTURA", "3"),
    ("05647", "SAN ANDRÉS DE CUERQUÍA", "3"), ("05579", "PUERTO BERRÍO", "3"),
    ("05490", "NECOCLÍ", "3"), ("05660", "SAN LUIS", "3"), ("05628", "SABANALARGA", "3"),
    ("05483", "NARIÑO", "3"), ("05756", "SONSON", "3"), ("05150", "CAROLINA", "3"),
    ("05736", "SEGOVIA", "3"), ("05890", "YOLOMBÓ", "3"), ("05086", "BELMIRA", "3"),
    ("05854", "VALDIVIA", "3"),
    ("05659", "SAN JUAN DE URABÁ", "4"), ("05789", "TÁMESIS", "4"), ("05792", "TARSO", "4"),
    ("05059", "ARMENIA", "4"), ("05856", "VALPARAÍSO", "4"), ("05145", "CARAMANTA", "4"),
    ("05125", "CAICEDO", "4"), ("05665", "SAN PEDRO DE URABÁ", "4"),
    ("05667", "SAN RAFAEL", "4"), ("05021", "ALEJANDRÍA", "4"), ("05051", "ARBOLETES", "4"),
    ("05837", "TURBO", "4"), ("05658", "SAN JOSÉ DE LA MONTAÑA", "4"),
    ("05055", "ARGELIA", "4"), ("05842", "URAMITA", "4"), ("05649", "SAN CARLOS", "4"),
    ("05790", "TARAZÁ", "4"), ("05495", "NECHÍ", "4"), ("05858", "VEGACHÍ", "4"),
    ("05031", "AMALFI", "4"), ("05107", "BRICEÑO", "4"), ("05585", "PUERTO NARE", "4"),
    ("05652", "SAN FRANCISCO", "4"), ("05895", "ZARAGOZA", "4"), ("05142", "CARACOLÍ", "4"),
    ("05885", "YALÍ", "4"), ("05284", "FRONTINO", "4"), ("05120", "CÁCERES", "4"),
    ("05543", "PEQUE", "4"), ("05480", "MUTATÁ", "4"), ("05234", "DABEIBA", "4"),
    ("05604", "REMEDIOS", "4"), ("05040", "ANORÍ", "4"), ("05847", "URRAO", "4"),
    ("05893", "YONDÓ", "4"), ("05361", "ITUANGO", "4"),
    ("05425", "MACEO", "5"), ("05004", "ABRIAQUÍ", "5"),
    ("05873", "VIGÍA DEL FUERTE", "5"), ("05475", "MURINDÓ", "5"),
]


def _normalizar(texto: str) -> str:
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    # La fuente DNP usa formas cortas ("RETIRO") que otros archivos escriben
    # con artículo ("EL RETIRO", "LA CEJA" ya coincide, etc.). Quitamos el
    # artículo inicial para poder cruzar ambas variantes.
    for articulo in ("EL ", "LA ", "LOS ", "LAS "):
        if texto.startswith(articulo):
            texto = texto[len(articulo):]
            break
    return texto


_INDICE_MUNICIPIO_TIPOLOGIA: dict[str, str] = {
    _normalizar(municipio): tipologia
    for _cod, municipio, tipologia in TIPOLOGIAS_ANTIOQUIA_2025
}

_INDICE_DANE_TIPOLOGIA: dict[str, str] = {
    cod: tipologia for cod, _municipio, tipologia in TIPOLOGIAS_ANTIOQUIA_2025
}


def tipologia_de(nombre_municipio: str) -> str | None:
    """Tipología DNP 2025 de un municipio de Antioquia (o None si no aplica)."""
    return _INDICE_MUNICIPIO_TIPOLOGIA.get(_normalizar(nombre_municipio))


def tipologia_de_codigo_dane(codigo_dane: str) -> str | None:
    """Tipología DNP 2025 a partir del código DANE (p.ej. '05615')."""
    return _INDICE_DANE_TIPOLOGIA.get(codigo_dane.strip())


def municipios_de_tipologia(tipologia: str) -> list[str]:
    """Lista de municipios (nombre tal como en la fuente) de una tipología dada."""
    return [m for _c, m, t in TIPOLOGIAS_ANTIOQUIA_2025 if t == tipologia]
