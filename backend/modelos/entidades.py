"""Modelos de dominio para el diagnóstico institucional IDI-MIPG."""
from __future__ import annotations

from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# Régimen especial MIPG/MECI: algunas entidades no están obligadas a
# implementar el MIPG en su integralidad (7 dimensiones, 19 políticas) sino
# únicamente la política de Control Interno (MECI), en virtud del artículo 40
# de la Ley 489 de 1998 y del artículo 2.2.22.3.4 del Decreto 1499 de 2017.
#
# Estos códigos son la fuente de verdad única del sistema: motor_diagnostico.py
# los usa para no penalizar a estas entidades por políticas que no les aplican,
# y backend/motores/generador_informe.py los usa para la nota jurídica que se
# inserta en los 3 informes. Si se agrega una categoría nueva, agréguese aquí
# primero y luego en el catálogo de notas de generador_informe.py.
# ---------------------------------------------------------------------------

REGIMEN_ESPECIAL_NINGUNO = "ninguno"  # Rama Ejecutiva: MIPG íntegro aplica
REGIMEN_ESPECIAL_UNIVERSIDAD_AUTONOMA = "universidad_autonoma"
REGIMEN_ESPECIAL_ORGANO_CONTROL = "organo_control"  # Contraloría territorial
REGIMEN_ESPECIAL_PERSONERIA = "personeria"
REGIMEN_ESPECIAL_CONCEJO_ASAMBLEA = "concejo_asamblea"
REGIMEN_ESPECIAL_BANCO_REPUBLICA = "banco_republica"
REGIMEN_ESPECIAL_CORPORACION_AUTONOMA = "corporacion_autonoma_regional"
REGIMEN_ESPECIAL_RAMA_LEGISLATIVA = "rama_legislativa"  # Senado, Cámara de Representantes
REGIMEN_ESPECIAL_ORGANO_CONTROL_NACIONAL = "organo_control_nacional"  # Procuraduría, Contraloría General, Defensoría, Auditoría General
REGIMEN_ESPECIAL_RAMA_JUDICIAL = "rama_judicial"  # Fiscalía, Consejo Superior de la Judicatura, Medicina Legal
REGIMEN_ESPECIAL_ORGANIZACION_ELECTORAL = "organizacion_electoral"  # Registraduría, Consejo Nacional Electoral
# CORRECCIÓN (agosto 2026, hallazgo en Corporación Ruta N Medellín): el
# artículo 2.2.22.3.4 del Decreto 1499 de 2017 nombra EXPLÍCITAMENTE a "los
# institutos científicos y tecnológicos" en el mismo listado que las
# Contralorías/Personerías/Concejos/Asambleas/Ramas Legislativa y Judicial —
# es decir, la ley no los trata como un caso ambiguo o "parecido" a régimen
# especial: los nombra por su tipo exacto. Solo la política de Control
# Interno (Ley 87/1993) es obligatoria para ellos; las demás políticas
# "aplicarán... en la medida en que les sean aplicables" (mismo texto legal
# que rige a las Contralorías) — es decir, son voluntarias/condicionadas,
# nunca exigencia normativa plena.
REGIMEN_ESPECIAL_INSTITUTO_CIENTIFICO_TECNOLOGICO = "instituto_cientifico_tecnologico"  # Ruta N Medellín y análogos

# Categorías que NO están obligadas al MIPG íntegro (solo MECI/Control Interno)
ENTIDADES_MECI_UNICAMENTE = frozenset({
    REGIMEN_ESPECIAL_UNIVERSIDAD_AUTONOMA,
    REGIMEN_ESPECIAL_ORGANO_CONTROL,
    REGIMEN_ESPECIAL_PERSONERIA,
    REGIMEN_ESPECIAL_CONCEJO_ASAMBLEA,
    REGIMEN_ESPECIAL_BANCO_REPUBLICA,
    REGIMEN_ESPECIAL_CORPORACION_AUTONOMA,
    REGIMEN_ESPECIAL_RAMA_LEGISLATIVA,
    REGIMEN_ESPECIAL_ORGANO_CONTROL_NACIONAL,
    REGIMEN_ESPECIAL_RAMA_JUDICIAL,
    REGIMEN_ESPECIAL_ORGANIZACION_ELECTORAL,
    REGIMEN_ESPECIAL_INSTITUTO_CIENTIFICO_TECNOLOGICO,
})


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
    regimen_especial: str | None = None  # ver constantes REGIMEN_ESPECIAL_* arriba

    def aplica_mipg_integral(self) -> bool:
        """True si esta entidad está obligada al MIPG en su integralidad (7
        dimensiones, 19 políticas). False si, por su naturaleza jurídica
        (ente universitario autónomo, órgano de control, Concejo/Asamblea,
        Banco de la República, Corporación Autónoma Regional), solo está
        obligada a la política de Control Interno (MECI)."""
        return self.regimen_especial not in ENTIDADES_MECI_UNICAMENTE

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
