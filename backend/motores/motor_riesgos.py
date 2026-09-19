"""Motor de cálculo del Módulo de Administración y Gestión del Riesgo.

Archivo NUEVO, independiente de motor_diagnostico.py (motor del IDI-MIPG
v6.1), al que no se le hace ningún cambio.

Implementa literalmente las reglas de la Guía para la Gestión Integral del
Riesgo en Entidades Públicas, Versión 7 (DAFP, agosto de 2025):
  - Tabla 4 (numeral 3.5): niveles de probabilidad según frecuencia anual.
  - Tabla 5 (numeral 3.6): niveles de impacto según afectación económica
    (SMLMV) y reputacional; se toma el nivel más alto entre ambos.
  - Figura 18 (numeral 3.7): matriz de calor de 4 zonas de severidad.
  - Tabla 6 (numeral 3.10): valoración de eficiencia de controles.
  - Tabla 8 (numeral 3.11): aplicación acumulativa de controles para
    obtener el riesgo residual.
"""
from __future__ import annotations

from backend.base_conocimiento.catalogo_riesgos import (
    tabla_probabilidad,
    tabla_impacto,
    tabla_valoracion_controles,
    matriz_calor_severidad,
)
from backend.modelos.riesgos import Control, ValoracionRiesgo, Riesgo, KRI


def calcular_probabilidad(frecuencia_anual: int) -> tuple[float, str]:
    """Devuelve (valor 0-1, nivel) según la Tabla 4 de la Guía v7."""
    for fila in tabla_probabilidad():
        minimo = fila["frecuencia_anual_min"]
        maximo = fila["frecuencia_anual_max"]
        if frecuencia_anual >= minimo and (maximo is None or frecuencia_anual <= maximo):
            return fila["valor"], fila["nivel"]
    # Si por algún motivo no cae en ningún rango, se toma el nivel más alto
    ultimo = tabla_probabilidad()[-1]
    return ultimo["valor"], ultimo["nivel"]


def calcular_impacto(smlmv: float, nivel_reputacional_elegido: str | None = None) -> tuple[float, str]:
    """Devuelve (valor 0-1, nivel) tomando el MAYOR entre el impacto
    económico (según SMLMV) y el impacto reputacional elegido por el
    usuario (Guía v7, numeral 3.6: 'se deberá tomar el nivel más alto')."""
    niveles = tabla_impacto()
    orden = {fila["nivel"]: i for i, fila in enumerate(niveles)}

    # Impacto económico según rango de SMLMV
    nivel_economico = niveles[-1]
    for fila in niveles:
        minimo = fila["smlmv_min"]
        maximo = fila["smlmv_max"]
        if smlmv >= minimo and (maximo is None or smlmv <= maximo):
            nivel_economico = fila
            break

    if nivel_reputacional_elegido and nivel_reputacional_elegido in orden:
        nivel_reputacional = niveles[orden[nivel_reputacional_elegido]]
    else:
        nivel_reputacional = niveles[0]

    # El más alto de los dos (mayor índice en la lista = mayor severidad)
    if orden[nivel_economico["nivel"]] >= orden[nivel_reputacional["nivel"]]:
        elegido = nivel_economico
    else:
        elegido = nivel_reputacional
    return elegido["valor"], elegido["nivel"]


_NIVELES_PROB_ORDEN = ["Muy Baja", "Baja", "Media", "Alta", "Muy Alta"]
_NIVELES_IMP_ORDEN = ["Leve", "Menor", "Moderado", "Mayor", "Catastrófico"]
_VALOR_A_NIVEL_PROB = {0.20: "Muy Baja", 0.40: "Baja", 0.60: "Media", 0.80: "Alta", 1.00: "Muy Alta"}
_VALOR_A_NIVEL_IMP = {0.20: "Leve", 0.40: "Menor", 0.60: "Moderado", 0.80: "Mayor", 1.00: "Catastrófico"}


def calcular_severidad(probabilidad: float, impacto: float) -> str:
    """Ubica el cruce probabilidad x impacto en la MATRIZ DE CALOR OFICIAL
    del Anexo 1 (Formato mapa de riesgos integral, DAFP) — no una
    aproximación: es la tabla de 4 zonas (Bajo/Moderado/Alto/Extremo) que
    trae el Excel oficial del DAFP para la Guía v7."""
    matriz = matriz_calor_severidad()
    nivel_prob = _VALOR_A_NIVEL_PROB.get(round(probabilidad, 2))
    nivel_imp = _VALOR_A_NIVEL_IMP.get(round(impacto, 2))

    if nivel_prob is None:
        # Probabilidad residual con decimales (p.ej. 25.2%): se redondea al
        # nivel válido más cercano de la tabla 4, para poder ubicarla en la
        # matriz de calor oficial (que solo tiene 5 niveles discretos).
        nivel_prob = min(_VALOR_A_NIVEL_PROB.items(), key=lambda kv: abs(kv[0] - probabilidad))[1]
    if nivel_imp is None:
        nivel_imp = min(_VALOR_A_NIVEL_IMP.items(), key=lambda kv: abs(kv[0] - impacto))[1]

    for fila in matriz["filas"]:
        if fila["probabilidad"] == nivel_prob:
            return fila[nivel_imp]
    return "Sin determinar"


def valorar_riesgo_inherente(riesgo: Riesgo) -> ValoracionRiesgo:
    """Calcula la valoración inherente de un riesgo a partir de su
    frecuencia anual y su afectación económica/reputacional."""
    if riesgo.frecuencia_anual is None or riesgo.smlmv_afectacion is None:
        raise ValueError(
            "El riesgo necesita frecuencia_anual y smlmv_afectacion para "
            "calcular la valoración inherente (Guía v7, Paso 2)."
        )
    prob_valor, prob_nivel = calcular_probabilidad(riesgo.frecuencia_anual)
    imp_valor, imp_nivel = calcular_impacto(riesgo.smlmv_afectacion, riesgo.nivel_reputacional)
    severidad = calcular_severidad(prob_valor, imp_valor)
    return ValoracionRiesgo(
        probabilidad=prob_valor,
        impacto=imp_valor,
        nivel_probabilidad=prob_nivel,
        nivel_impacto=imp_nivel,
        severidad=severidad,
    )


def _eje_del_control(control: Control, tabla: dict) -> str:
    """Devuelve 'probabilidad' o 'impacto' según el TIPO del control, tal
    como lo define el Anexo 1 oficial (hoja '11 FORMULAS', tabla
    B51:D53): Preventivo y Detectivo afectan la probabilidad; Correctivo
    afecta el impacto. (Corrección de la versión 1.0 del motor, que
    aplicaba todos los controles sobre la probabilidad por defecto)."""
    return tabla["tipo"][control.tipo]["afecta"]


def aplicar_controles(valor_inicial: float, controles: list[Control], eje: str, tabla: dict) -> float:
    """Aplica, de forma ACUMULATIVA, solo los controles cuyo tipo afecta el
    eje indicado ('probabilidad' o 'impacto'), tal como lo ejemplifica la
    Tabla 8 de la Guía v7 y lo confirma el Anexo 1 oficial: el resultado de
    aplicar un control se usa como base para aplicar el siguiente.

    `tabla` es la tabla de valoración de controles a usar — NO se asume una
    única tabla: la tipología 'Seguridad de la información' usa el peso del
    Anexo 5 (Manual=10%), las demás usan el de la Guía v7/Anexo 1
    (Manual=15%). Ver catalogo_riesgos.tabla_valoracion_controles().

    Ejemplo de la guía (dos controles que SÍ afectan probabilidad, tabla
    general): probabilidad inherente 60%, control preventivo (peso
    25%+15%=40%) -> 60% - (60% * 40%) = 36%; luego control detectivo (peso
    15%+15%=30%) sobre el 36% -> 36% - (36% * 30%) = 25.2% (probabilidad
    residual)."""
    valor = valor_inicial
    for control in controles:
        if _eje_del_control(control, tabla) != eje:
            continue
        peso = control.peso_eficiencia(tabla)
        valor = valor - (valor * peso)
    return max(valor, 0.0)


def calcular_riesgo_residual(riesgo: Riesgo) -> ValoracionRiesgo:
    """Calcula la valoración residual aplicando los controles preventivos y
    detectivos del riesgo (acumulativamente) sobre la probabilidad, y los
    controles correctivos (acumulativamente) sobre el impacto — regla
    confirmada contra el Anexo 1 oficial del DAFP. Si un riesgo no tiene
    controles de un tipo dado, ese eje queda igual al valor inherente
    (mismo comportamiento que el ejemplo de la Tabla 8 de la Guía v7, donde
    el impacto residual queda igual al inherente por falta de controles de
    impacto).

    La tabla de pesos usada depende de la TIPOLOGÍA del riesgo (instrucción
    de la docente/experta temática, 19-sep-2026: no unificar los pesos)."""
    if riesgo.inherente is None:
        riesgo.inherente = valorar_riesgo_inherente(riesgo)

    tabla = tabla_valoracion_controles(riesgo.tipologia)
    prob_residual = aplicar_controles(riesgo.inherente.probabilidad, riesgo.controles, "probabilidad", tabla)
    imp_residual = aplicar_controles(riesgo.inherente.impacto, riesgo.controles, "impacto", tabla)

    severidad = calcular_severidad(prob_residual, imp_residual)
    return ValoracionRiesgo(
        probabilidad=prob_residual,
        impacto=imp_residual,
        nivel_probabilidad=riesgo.inherente.nivel_probabilidad,
        nivel_impacto=riesgo.inherente.nivel_impacto,
        severidad=severidad,
    )


def consolidar_mapa(riesgos: list[Riesgo]) -> dict:
    """Agrupa los riesgos por proceso y por tipología para armar el mapa de
    riesgos integral (Guía v7, numeral 3.12)."""
    mapa: dict[str, dict[str, list[Riesgo]]] = {}
    for r in riesgos:
        mapa.setdefault(r.proceso, {}).setdefault(r.tipologia, []).append(r)
    return mapa


def evaluar_kri(kri: KRI) -> str:
    """Delegado a KRI.semaforo(); expuesto también como función del motor
    para mantener el mismo patrón `motor_*` que usa el resto de SIIEAP."""
    return kri.semaforo()
