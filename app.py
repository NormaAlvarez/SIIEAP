"""
SIIEAP — punto de entrada único de la interfaz (Streamlit).

Ejecutar con:
    streamlit run app.py

Este archivo es el ROUTER principal del sistema. Desde el 19-sep-2026 ofrece
un menú en la barra lateral para elegir entre los dos módulos disponibles:

  1. Diagnóstico IDI-MIPG (v6.1)  -> backend original, sin ningún cambio,
     ahora vive en diagnostico_idi_mipg.py (función render_diagnostico_idi_mipg()).
  2. Administración y Gestión del Riesgo (piloto) -> módulo nuevo construido
     sobre la Guía para la Gestión Integral del Riesgo en Entidades Públicas
     v7 (DAFP, agosto 2025) y validado por la docente/experta temática
     (Norma Elizabeth Álvarez Grajales) el 19-sep-2026; vive en
     modulo_riesgo_piloto.py (función render_modulo_riesgo()).

Ninguno de los dos módulos fue reescrito para esta integración: cada uno
conserva exactamente sus pantallas, cálculos y validaciones. Lo único nuevo
aquí es el selector que decide cuál de los dos se ejecuta en cada corrida.

Ambos módulos también se pueden seguir ejecutando por separado, sin pasar
por este router, con:
    streamlit run diagnostico_idi_mipg.py
    streamlit run modulo_riesgo_piloto.py
"""
import streamlit as st

from diagnostico_idi_mipg import render_diagnostico_idi_mipg
from modulo_riesgo_piloto import render_modulo_riesgo

st.set_page_config(
    page_title="SIIEAP — Sistema de Inteligencia Artificial para la Evaluación "
    "Integral del Desempeño Institucional",
    layout="wide",
)

MODULO_DIAGNOSTICO = "📊 Diagnóstico IDI-MIPG (v6.1)"
MODULO_RIESGO = "🛡️ Administración y Gestión del Riesgo (piloto)"

with st.sidebar:
    st.markdown("## SIIEAP")
    st.caption(
        "Escuela Superior de Administración Pública (ESAP) · Docente: "
        "Norma Elizabeth Álvarez Grajales"
    )
    modulo_elegido = st.radio(
        "Módulo",
        [MODULO_DIAGNOSTICO, MODULO_RIESGO],
        key="modulo_siieap_elegido",
    )
    st.divider()

if modulo_elegido == MODULO_DIAGNOSTICO:
    render_diagnostico_idi_mipg()
else:
    render_modulo_riesgo()
