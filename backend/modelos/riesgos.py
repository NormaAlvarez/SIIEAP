"""Modelos de dominio para el Módulo de Administración y Gestión del Riesgo.

Archivo NUEVO y separado de backend/modelos/entidades.py (modelos del
diagnóstico IDI-MIPG v6.1), al que no se le hace ningún cambio.

Fuente: Guía para la Gestión Integral del Riesgo en Entidades Públicas,
Versión 7 (DAFP, agosto de 2025).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

TIPOLOGIA_GESTION = "Gestión"
TIPOLOGIA_FISCAL = "Fiscal"
TIPOLOGIA_SEGURIDAD_INFORMACION = "Seguridad de la información"
TIPOLOGIA_INTEGRIDAD_SIGRIP = "Integridad pública (SIGRIP)"
TIPOLOGIA_AMBIENTAL = "Ambiental"

TIPO_CONTROL_PREVENTIVO = "preventivo"
TIPO_CONTROL_DETECTIVO = "detectivo"
TIPO_CONTROL_CORRECTIVO = "correctivo"

IMPLEMENTACION_AUTOMATICO = "automatico"
IMPLEMENTACION_MANUAL = "manual"


@dataclass
class ProcesoCaracterizado:
    """Caracterización completa de un proceso institucional, más allá de su
    tipo general (Estratégico/Misional/Apoyo/Evaluación). Agregado a
    petición de la docente/experta temática (validación 19-sep-2026) para
    poder vincular cada riesgo a un proceso caracterizado con su objetivo,
    alcance, entradas, salidas, responsable e indicadores — no solo a su
    tipo general, mientras se recibe el listado nominal oficial de Carepa."""
    codigo: str
    nombre: str
    tipo: str  # Estratégico | Misional | De Apoyo | De Evaluación
    objetivo: str = ""
    alcance: str = ""
    entradas: str = ""
    salidas: str = ""
    responsable: str = ""
    indicadores: list[str] = field(default_factory=list)


@dataclass
class Control:
    """Un control de los descritos en la Guía v7, numeral 3.8-3.10."""
    responsable: str
    accion: str
    tipo: str  # preventivo | detectivo | correctivo
    implementacion: str  # automatico | manual
    documentacion: Optional[str] = None
    frecuencia: Optional[str] = None
    evidencia: Optional[str] = None
    fuente_ejecucion: Optional[str] = None  # interna | externa | mixta

    def peso_eficiencia(self, tabla_valoracion: dict) -> float:
        """Suma el peso de tipo + peso de implementación (Tabla 6, Guía v7).
        Ej.: preventivo (25%) + automático (25%) = 50% de mitigación.
        tabla_valoracion["tipo"][self.tipo] es un dict {"peso": .., "afecta": ..}
        desde que se incorporó el Anexo 1 oficial (ver motor_riesgos._eje_del_control)."""
        return (
            tabla_valoracion["tipo"][self.tipo]["peso"]
            + tabla_valoracion["implementacion"][self.implementacion]
        )


@dataclass
class ValoracionRiesgo:
    """Resultado del análisis de probabilidad/impacto/severidad de un riesgo,
    antes (inherente) o después (residual) de aplicar controles."""
    probabilidad: float  # 0.20 a 1.00
    impacto: float  # 0.20 a 1.00
    nivel_probabilidad: str = ""
    nivel_impacto: str = ""
    severidad: str = ""  # etiqueta de zona en la matriz de calor


@dataclass
class Riesgo:
    """Un riesgo identificado en un proceso, siguiendo la estructura de
    descripción de la Guía v7 (numeral 3.4): Impacto + Causa inmediata ->
    Evento no deseado; Causa raíz; Tipología; Factor de riesgo."""
    entidad: str
    proceso: str
    nombre: str
    causa_raiz: str
    causa_inmediata: str
    evento_no_deseado: str
    impacto_descripcion: str
    factor_riesgo: str  # código del catálogo (FR01..FR06)
    tipologia: str  # Gestión | Fiscal | Seguridad de la información | Integridad pública (SIGRIP) | Ambiental

    frecuencia_anual: Optional[int] = None  # base para la probabilidad inherente
    smlmv_afectacion: Optional[float] = None  # base para el impacto económico
    nivel_reputacional: Optional[str] = None  # base para el impacto reputacional

    controles: list[Control] = field(default_factory=list)

    inherente: Optional[ValoracionRiesgo] = None
    residual: Optional[ValoracionRiesgo] = None

    # Campos específicos por tipología (se dejan opcionales para no forzar
    # estructura a riesgos de tipología "Gestión")
    punto_riesgo_fiscal: Optional[str] = None
    activo_informacion: Optional[str] = None
    clasificacion_informacion: Optional[str] = None
    amenaza_integridad: Optional[str] = None
    aspecto_ambiental: Optional[str] = None
    instrumento_origen_ambiental: Optional[str] = None


@dataclass
class KRI:
    """Indicador Clave de Riesgo (Guía v7, Capítulo VIII)."""
    riesgo_asociado: str  # nombre o id del Riesgo
    nombre: str
    umbral_alerta: float
    umbral_apetito: float
    valor_actual: float
    responsable: str

    def semaforo(self) -> str:
        """Verde: dentro del apetito. Amarillo: cerca del umbral de alerta.
        Rojo: excede el apetito al riesgo. (numeral 9.3, Guía v7)"""
        if self.valor_actual >= self.umbral_apetito:
            return "rojo"
        if self.valor_actual >= self.umbral_alerta:
            return "amarillo"
        return "verde"
