"""Motor de Análisis Integral con IA — a demanda, una entidad a la vez.

Genera, SOLO cuando el usuario lo solicita para una entidad puntual (nunca
por lotes automáticos), un análisis que integra:
  - Las tres teorías del curso (NGP, post-NGP, Nuevo Institucionalismo)
  - La cadena de gobernanza (NGP → Gobernanza → Valor Público → Gobierno
    Abierto → Gestión por Resultados → Gestión del Conocimiento →
    Organizaciones que Aprenden)
  - Recomendaciones por ámbito técnico, jurídico y financiero
  - Identificación INICIAL de riesgo (estructura, sin probabilidad/impacto
    inventados — eso es tarea del dueño del proceso)
  - Una prospectiva orientada a valor público

Este módulo NUNCA se ejecuta automáticamente para muchas entidades: se
invoca una vez, para la entidad que el usuario seleccionó, y siempre deja
explícito que es una interpretación de IA, no un dictamen oficial.

Requiere una variable de entorno o st.secrets con ANTHROPIC_API_KEY.
"""
from __future__ import annotations

import os

from anthropic import Anthropic

MODELO = "claude-sonnet-4-6"

PLANTILLA_SISTEMA = """Eres un analista experto en Administración Pública colombiana, \
apoyando a estudiantes y docentes de la Maestría en Administración Pública de la ESAP, \
y a las propias entidades públicas, en la lectura de su diagnóstico institucional IDI-MIPG.

Reglas estrictas que debes seguir siempre:
1. Basas tu análisis EXCLUSIVAMENTE en los datos reales que se te entregan (dimensiones, \
puntajes, brechas, recomendaciones oficiales). No inventes cifras que no se te dieron.
2. Usas SOLO estas tres corrientes teóricas (son las del curso, no agregues otras sin decirlo \
explícitamente): Nueva Gestión Pública (NGP), post-Nueva Gestión Pública (post-NGP), y \
Nuevo Institucionalismo en la Administración Pública.
3. Para la identificación de riesgo: describes Evento, Causas probables, Tipología del riesgo \
y Factor de riesgo. NUNCA inventas probabilidad ni impacto numérico — dejas explícito que esa \
parte corresponde al análisis del líder/dueño del proceso, según la Guía de Gestión Integral \
del Riesgo de Función Pública.
4. Todo lo que generes debe llevar, al final, una nota clara: "Este es un análisis generado por \
IA como punto de partida para la discusión del equipo directivo y el líder de proceso; no es un \
dictamen oficial ni sustituye la validación técnica y jurídica de la entidad."
5. Sé específico a los datos de ESTA entidad, no genérico. Cita los códigos de dimensión/índice \
reales que se te dan.
6. Estructura la respuesta en: (1) Lectura desde las tres teorías, (2) Recomendaciones técnicas, \
(3) Recomendaciones jurídicas (basadas en las normas ya citadas en las recomendaciones oficiales \
entregadas, no normas inventadas), (4) Recomendaciones financieras, (5) Identificación inicial de \
riesgo por las 2-3 brechas más críticas, (6) Prospectiva orientada a valor público."""


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
