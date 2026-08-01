"""Modelos de dominio para el diagnóstico institucional IDI-MIPG."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ResultadoIndice:
    """Resultado oficial de un índice para una entidad, en una vigencia dada."""
    codigo_indice: str      # p.ej. "I01"
    puntaje: float          # 0-100, tal como lo reporta FURAG
    vigencia: int

    def __post_init__(self):
        self.codigo_indice = self.codigo_indice.strip().upper()
        if not (0 <= self.puntaje <= 100):
            raise ValueError(
                f"Puntaje fuera de rango para {self.codigo_indice}: {self.puntaje} "
                "(se espera 0-100, escala oficial FURAG)"
            )


@dataclass
class ResultadoPolitica:
    """Resultado directo de una política SIN índices propios (p.ej. POL14).

    La mayoría de políticas se calculan agregando sus índices (I01, I02...),
    pero algunas (como POL14: Seguimiento y Evaluación del Desempeño
    Institucional) se reportan como un puntaje directo, sin descomposición.
    """
    codigo_politica: str  # p.ej. "POL14"
    puntaje: float
    vigencia: int

    def __post_init__(self):
        self.codigo_politica = self.codigo_politica.strip().upper()
        if not (0 <= self.puntaje <= 100):
            raise ValueError(
                f"Puntaje fuera de rango para {self.codigo_politica}: {self.puntaje}"
            )


@dataclass
class Entidad:
    """Una entidad pública colombiana y sus resultados IDI-MIPG de una vigencia."""
    nombre: str
    codigo_dane: str | None
    vigencia: int
    resultados: list[ResultadoIndice] = field(default_factory=list)
    resultados_politica_directa: list[ResultadoPolitica] = field(default_factory=list)

    def resultado_de(self, codigo_indice: str) -> ResultadoIndice | None:
        codigo_indice = codigo_indice.strip().upper()
        for r in self.resultados:
            if r.codigo_indice == codigo_indice:
                return r
        return None

    def resultado_politica_directa_de(self, codigo_politica: str) -> ResultadoPolitica | None:
        codigo_politica = codigo_politica.strip().upper()
        for r in self.resultados_politica_directa:
            if r.codigo_politica == codigo_politica:
                return r
        return None

    def codigos_reportados(self) -> set[str]:
        return {r.codigo_indice for r in self.resultados}
