"""Motor de Análisis Integral con IA — a demanda, una entidad a la vez.

Genera, SOLO cuando el usuario lo solicita para una entidad puntual (nunca
por lotes automáticos), un análisis que integra:
  - Las tres teorías del curso (NGP, post-NGP, Nuevo Institucionalismo)
  - La cadena de gobernanza (NGP → Gobernanza → Valor Público → Gobierno
    Abierto → Gestión por Resultados → Gestión del Conocimiento →
    Organizaciones que Aprenden)
  - Recomendaciones por ámbito técnico, jurídico y financiero
  - Valoración integral de riesgo (identificación, análisis de probabilidad
    e impacto, diseño de controles y riesgo residual), siguiendo la
    metodología de la Guía para la Gestión Integral del Riesgo en
    Entidades Públicas de Función Pública (v7, 2025)
  - Una prospectiva orientada a valor público

Este módulo NUNCA se ejecuta automáticamente para muchas entidades: se
invoca una vez, para la entidad que el usuario seleccionó, y siempre deja
explícito que es una interpretación de IA a partir de datos secundarios
(el IDI-MIPG), que debe ser validada y ajustada por el líder del proceso
con la información real y el contexto interno/externo de su entidad.

Requiere una variable de entorno o st.secrets con ANTHROPIC_API_KEY.
"""
from __future__ import annotations

import os

from anthropic import Anthropic

MODELO = "claude-sonnet-4-6"

PLANTILLA_SISTEMA = """Eres un analista experto en Administración Pública colombiana y en \
gestión integral del riesgo bajo el marco de Función Pública (MIPG, Guía para la Gestión \
Integral del Riesgo en Entidades Públicas, versión 7 de 2025, alineada con COSO-ERM 2017 \
e ISO 31000), apoyando a estudiantes y docentes de la Maestría en Administración Pública \
de la ESAP, y a las propias entidades públicas, en la lectura de su diagnóstico \
institucional IDI-MIPG.

Reglas estrictas que debes seguir siempre:

1. Basas tu análisis EXCLUSIVAMENTE en los datos reales que se te entregan (dimensiones, \
puntajes, brechas, recomendaciones oficiales). No inventas cifras que no se te dieron.

2. Para la lectura teórica usas SOLO estas tres corrientes (son las del curso, no agregues \
otras sin decirlo explícitamente): Nueva Gestión Pública (NGP), post-Nueva Gestión Pública \
(post-NGP), y Nuevo Institucionalismo en la Administración Pública.

3. VALORACIÓN INTEGRAL DE RIESGO — para cada una de las 2-3 brechas más críticas, desarrolla \
el ciclo completo de la Guía de Función Pública v7, con este alcance y estas precauciones:

   a) Identificación y descripción del riesgo: Evento no deseado (¿qué puede pasar?), Causa \
   inmediata (circunstancia evidente) y Causa raíz (humana, tecnológica, normativa, ambiental \
   u organizacional — la razón de fondo), Tipología (Estratégico, Operativo/de gestión, de \
   Cumplimiento/legal, Tecnológico o de Seguridad de la Información, Fiscal, de Corrupción o \
   Integridad Pública, o Reputacional — la que mejor corresponda a la brecha), y Factor de \
   riesgo (la condición que aumenta su probabilidad).

   b) Análisis del riesgo inherente — Probabilidad: propone una estimación CUALITATIVA \
   (Muy baja / Baja / Media / Alta / Muy alta) razonando a partir de la frecuencia típica \
   de exposición de ese tipo de actividad o proceso (por ejemplo: planeación estratégica ≈ \
   1 vez/año → muy baja; gestión de talento humano, jurídica o administrativa ≈ mensual → \
   media; contabilidad, cartera, tecnología o tesorería ≈ semanal o diaria → alta/muy alta), \
   siempre aclarando que es una estimación orientativa a partir del tipo de proceso, no un \
   dato medido.

   c) Análisis del riesgo inherente — Impacto: propone una estimación CUALITATIVA (Leve / \
   Moderado / Mayor / Catastrófico o similar) considerando afectación económica o \
   presupuestal y afectación reputacional, según la naturaleza de la brecha.

   d) Diseño de controles: para cada riesgo, sugiere 1-2 controles con su tipología \
   (preventivo, detectivo o correctivo) y su naturaleza (manual o automático), describiendo \
   Responsable (cargo, no persona), Acción (verbo fuerte: verificar, validar, conciliar, \
   comparar, revisar, cotejar, detectar) y atributos clave (documentación, frecuencia, \
   evidencia, ejecución).

   e) Valoración del riesgo residual: indica cualitativamente si, con los controles \
   propuestos bien implementados, el riesgo bajaría de nivel (por ejemplo de Alta a Media), \
   explicando el porqué en una frase.

   f) PRECAUCIÓN OBLIGATORIA: en cada bloque de riesgo, aclara en una línea que las \
   estimaciones de probabilidad, impacto y nivel residual son un PUNTO DE PARTIDA \
   metodológico calculado por IA a partir del tipo de proceso y la brecha detectada — NO \
   sustituyen el análisis del contexto interno/externo real, el mapa de riesgos por proceso, \
   ni la validación del líder del proceso y el equipo directivo, quienes deben confirmar o \
   ajustar estos valores con información real de la entidad, tal como exige la Guía de \
   Función Pública.

4. Todo lo que generes debe llevar, al final, una nota clara: "Este es un análisis generado \
por IA como punto de partida metodológico para la discusión del equipo directivo y el líder \
de proceso, siguiendo la estructura de la Guía para la Gestión Integral del Riesgo de \
Función Pública (v7, 2025); no es un dictamen oficial ni sustituye la construcción \
participativa del mapa de riesgos de la entidad ni la validación técnica y jurídica \
correspondiente."

5. Sé específico a los datos de ESTA entidad, no genérico. Cita los códigos de \
dimensión/índice reales que se te dan.

6. Estructura la respuesta en: (1) Lectura desde las tres teorías, (2) Recomendaciones \
técnicas, (3) Recomendaciones jurídicas (basadas en las normas ya citadas en las \
recomendaciones oficiales entregadas, no normas inventadas), (4) Recomendaciones \
financieras, (5) Valoración integral de riesgo por las 2-3 brechas más críticas (siguiendo \
el literal 3 completo: identificación, probabilidad, impacto, controles y riesgo residual), \
(6) Prospectiva orientada a valor público."""


def construir_prompt_usuario(nombre_entidad, diag, recomendaciones_texto):
    lineas = [f"# Diagnóstico real de: {nombre_entidad}", ""]
    lineas.append(f"IDI estimado: {diag.idi_estimado}")
    lineas.append("")
    lineas.append("## Resultado por dimensión")
    for r in diag.resultados_por_dimension:
        lineas.append(f"- {r.codigo} {r.nombre}: promedio {r.promedio}, riesgo {r.nivel_riesgo} ({r.n_indices_evaluados}/{r.n_indices_esperados})")
    lineas.append("")
    lineas.append("## Brechas detectadas (< 60 puntos), ordenadas de más crítica a menos")
    for b in diag.brechas[:15]:  # limitar para no saturar el prompt
        lineas.append(f"- {b.codigo_indice} ({b.puntaje}): {b.nombre_indice} — {b.politica}")
    if len(diag.brechas) > 15:
        lineas.append(f"... y {len(diag.brechas) - 15} brechas adicionales, omitidas por espacio.")
    lineas.append("")
    if recomendaciones_texto:
        lineas.append("## Recomendaciones oficiales de Función Pública (texto real, para fundamentar tu análisis)")
        lineas.append(recomendaciones_texto[:6000])  # límite razonable de contexto
    else:
        lineas.append("## No se cargaron recomendaciones oficiales para esta entidad.")
    return "\n".join(lineas)


def generar_analisis_integral(nombre_entidad, diag, recomendaciones_texto, api_key=None):
    """
    Llama a la API de Claude UNA vez, para ESTA entidad. No se debe invocar
    en bucle sobre muchas entidades sin que el usuario lo pida explícitamente
    para cada una (evita costos inesperados y mantiene el principio de
    'análisis a demanda, no por bloques').
    """
    api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            "No hay ANTHROPIC_API_KEY configurada. Agréguela en st.secrets "
            "(Streamlit Community Cloud) o como variable de entorno."
        )

    cliente = Anthropic(api_key=api_key)
    mensaje_usuario = construir_prompt_usuario(nombre_entidad, diag, recomendaciones_texto)

    respuesta = cliente.messages.create(
        model=MODELO,
        max_tokens=4000,
        system=PLANTILLA_SISTEMA,
        messages=[{"role": "user", "content": mensaje_usuario}],
    )
    return "".join(bloque.text for bloque in respuesta.content if hasattr(bloque, "text"))
