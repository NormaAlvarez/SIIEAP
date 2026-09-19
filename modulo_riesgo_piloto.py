"""
Módulo de Administración y Gestión del Riesgo — PILOTO SIIEAP
================================================================

Este archivo es un prototipo INDEPENDIENTE y NO modifica app.py ni ningún
otro archivo del SIIEAP v6.1. Se ejecuta por separado:

    streamlit run modulo_riesgo_piloto.py

Implementa las Pantallas 1-4 (metodología genérica), la Pantalla 5 (riesgo
fiscal, ya con el catálogo oficial de 50 puntos del Anexo 3) y la Pantalla 9
(riesgos ambientales), como caso de prueba con los procesos CONFIRMADOS de
la Alcaldía de Carepa (los 4 tipos de proceso del Decreto 092 de 2021; el
listado nominal de cada proceso individual sigue pendiente de verificar
directamente con la Alcaldía).

Actualización (18-sep-2026): se incorporaron los Anexos 1, 3 y 5 oficiales
de la Guía v7 (DAFP). Esto corrigió la fórmula de severidad (ahora usa la
matriz de calor oficial del Anexo 1, no una aproximación) y la regla de
qué eje afecta cada control (preventivo/detectivo -> probabilidad,
correctivo -> impacto, según el Anexo 1). Las Pantallas 6-8 (seguridad de
la información, SIGRIP, KRI) siguen fuera de este piloto.

Antes de integrarlo a app.py, este piloto debe:
  1) Validarse con la docente/experta temática (campos, textos, catálogos).
  2) Recibir el listado nominal real de procesos de Carepa.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd

from backend.base_conocimiento.catalogo_riesgos import (
    opciones_factor_riesgo,
    opciones_tipologia_riesgo,
    opciones_proceso_carepa,
    instrumentos_insumo_ambientales,
    catalogo_puntos_riesgo_fiscal_disponible,
    opciones_punto_riesgo_fiscal,
    circunstancia_inmediata_de,
    tratamiento_riesgo_opciones,
)
from backend.modelos.riesgos import Riesgo, Control, TIPOLOGIA_AMBIENTAL
from backend.motores.motor_riesgos import (
    valorar_riesgo_inherente,
    calcular_riesgo_residual,
    consolidar_mapa,
)

st.set_page_config(page_title="SIIEAP — Módulo de Riesgo (piloto)", layout="wide")

if "riesgos_capturados" not in st.session_state:
    st.session_state.riesgos_capturados = []  # list[Riesgo]

st.title("🛡️ Módulo de Administración y Gestión del Riesgo — PILOTO")
st.caption(
    "Prototipo independiente de SIIEAP v6.1. Caso de prueba: Alcaldía de Carepa "
    "(Antioquia), vigencia 2026-2027. No sustituye ni modifica el diagnóstico IDI-MIPG."
)
st.warning(
    "⚠ Este es un PILOTO para validación funcional y de contenido con la docente/experta "
    "temática. Los procesos de Carepa que se listan abajo son solo los 4 TIPOS confirmados "
    "(Decreto 092 de 2021); el listado nominal de cada proceso específico está pendiente "
    "de verificar con la Alcaldía.",
    icon="⚠️",
)

tab1, tab2, tab3, tab4, tab_fiscal, tab_amb, tab_mapa = st.tabs([
    "1. Identificación",
    "2. Análisis inherente",
    "3. Controles",
    "4. Riesgo residual",
    "5. Fiscal",
    "9. Ambiental",
    "Mapa consolidado",
])

# ---------------------------------------------------------------------
# Estado temporal del riesgo en construcción
# ---------------------------------------------------------------------
if "riesgo_en_edicion" not in st.session_state:
    st.session_state.riesgo_en_edicion = {}

with tab1:
    st.subheader("Paso 1 — Identificación del riesgo")
    col1, col2 = st.columns(2)
    with col1:
        proceso = st.selectbox("Proceso institucional (Carepa)", opciones_proceso_carepa(), key="f_proceso")
        nombre = st.text_input("Nombre del riesgo", key="f_nombre")
        factor = st.selectbox("Factor de riesgo (Tabla 2, Guía v7)", opciones_factor_riesgo(), key="f_factor")
        tipologia = st.selectbox("Tipología del riesgo", opciones_tipologia_riesgo(), key="f_tipologia")
    with col2:
        causa_raiz = st.text_area("Causa raíz", key="f_causa_raiz")
        causa_inmediata = st.text_area("Causa inmediata", key="f_causa_inmediata")
        evento = st.text_area("Evento no deseado", key="f_evento")
        impacto_desc = st.text_area("Impacto (descripción cualitativa)", key="f_impacto_desc")

    if tipologia == "Ambiental":
        st.info(
            "Tipología Ambiental: complete también los campos específicos en la pestaña "
            "'9. Ambiental' antes de guardar."
        )

    if st.button("Guardar identificación y continuar al Paso 2 →", type="primary"):
        if not nombre or not causa_raiz or not evento:
            st.error("Nombre, causa raíz y evento no deseado son obligatorios.")
        else:
            st.session_state.riesgo_en_edicion = dict(
                entidad="Alcaldía de Carepa",
                proceso=proceso,
                nombre=nombre,
                causa_raiz=causa_raiz,
                causa_inmediata=causa_inmediata,
                evento_no_deseado=evento,
                impacto_descripcion=impacto_desc,
                factor_riesgo=factor,
                tipologia=tipologia,
            )
            st.success("Identificación guardada. Pase a la pestaña '2. Análisis inherente'.")

with tab2:
    st.subheader("Paso 2 — Análisis de riesgo inherente")
    if not st.session_state.riesgo_en_edicion:
        st.info("Primero complete el Paso 1 en la pestaña 'Identificación'.")
    else:
        st.write(f"Riesgo en edición: **{st.session_state.riesgo_en_edicion.get('nombre', '')}**")
        frecuencia = st.number_input(
            "Frecuencia anual (n.º de veces al año que se pasa por el punto de riesgo)",
            min_value=0, step=1, key="f_frecuencia",
        )
        smlmv = st.number_input(
            "Afectación económica estimada (en SMLMV)", min_value=0.0, step=1.0, key="f_smlmv"
        )
        nivel_reputacional = st.selectbox(
            "Afectación reputacional (Tabla 5, Guía v7)",
            ["Leve", "Menor", "Moderado", "Mayor", "Catastrófico"],
            index=2, key="f_nivel_reputacional",
        )
        if st.button("Calcular probabilidad e impacto inherente"):
            r_temp = Riesgo(
                **st.session_state.riesgo_en_edicion,
                frecuencia_anual=int(frecuencia),
                smlmv_afectacion=float(smlmv),
                nivel_reputacional=nivel_reputacional,
            )
            inh = valorar_riesgo_inherente(r_temp)
            st.session_state.riesgo_en_edicion.update(
                frecuencia_anual=int(frecuencia),
                smlmv_afectacion=float(smlmv),
                nivel_reputacional=nivel_reputacional,
            )
            st.session_state["inherente_calculado"] = inh
            c1, c2, c3 = st.columns(3)
            c1.metric("Probabilidad inherente", f"{inh.probabilidad*100:.0f}%", inh.nivel_probabilidad)
            c2.metric("Impacto inherente", f"{inh.impacto*100:.0f}%", inh.nivel_impacto)
            c3.metric("Severidad inherente", inh.severidad)

with tab3:
    st.subheader("Paso 3 — Diseño y análisis de controles")
    if "inherente_calculado" not in st.session_state:
        st.info("Primero calcule el riesgo inherente en la pestaña '2. Análisis inherente'.")
    else:
        if "controles_temp" not in st.session_state:
            st.session_state.controles_temp = []

        with st.form("form_control", clear_on_submit=True):
            col1, col2, col3, col4 = st.columns(4)
            responsable = col1.text_input("Responsable")
            accion = col2.selectbox("Acción", ["Verificar", "Validar", "Conciliar", "Comparar", "Revisar", "Cotejar", "Detectar"])
            tipo = col3.selectbox("Tipo", ["preventivo", "detectivo", "correctivo"])
            implementacion = col4.selectbox("Implementación", ["automatico", "manual"])
            agregar = st.form_submit_button("+ Agregar control")
            if agregar and responsable:
                st.session_state.controles_temp.append(
                    Control(responsable=responsable, accion=accion, tipo=tipo, implementacion=implementacion)
                )

        if st.session_state.controles_temp:
            st.write("Controles registrados para este riesgo:")
            st.dataframe(
                pd.DataFrame([c.__dict__ for c in st.session_state.controles_temp]),
                use_container_width=True,
            )

with tab4:
    st.subheader("Paso 4 — Valoración del riesgo residual")
    if "inherente_calculado" not in st.session_state:
        st.info("Complete primero los pasos 1 y 2.")
    else:
        if st.button("Calcular riesgo residual y guardar riesgo en el mapa", type="primary"):
            datos = dict(st.session_state.riesgo_en_edicion)
            riesgo = Riesgo(**datos)
            riesgo.controles = st.session_state.get("controles_temp", [])
            riesgo.inherente = st.session_state["inherente_calculado"]
            residual = calcular_riesgo_residual(riesgo)
            riesgo.residual = residual

            st.session_state.riesgos_capturados.append(riesgo)

            c1, c2, c3 = st.columns(3)
            c1.metric("Probabilidad residual", f"{residual.probabilidad*100:.1f}%")
            c2.metric("Impacto residual", f"{residual.impacto*100:.1f}%")
            c3.metric("Severidad residual", residual.severidad)
            st.success(
                f"Riesgo '{riesgo.nombre}' guardado en el mapa consolidado "
                f"({len(st.session_state.riesgos_capturados)} riesgo(s) capturado(s))."
            )
            # limpiar estado del riesgo en edición para iniciar uno nuevo
            st.session_state.riesgo_en_edicion = {}
            st.session_state.pop("inherente_calculado", None)
            st.session_state.controles_temp = []

with tab_fiscal:
    st.subheader("Pantalla 5 — Riesgo fiscal (Capítulo IV, catálogo oficial Anexo 3)")
    st.caption(
        "Use esta pestaña ANTES del Paso 1 si la tipología del riesgo es 'Fiscal': elija el punto de "
        "riesgo fiscal del catálogo oficial (50 puntos, Anexo 3 de la Guía v7) y la circunstancia inmediata "
        "se autocompleta. Luego complete la identificación en la pestaña 1 con Tipología = Fiscal."
    )
    punto = st.selectbox("Punto de riesgo fiscal (Anexo 3, catálogo oficial DAFP)", opciones_punto_riesgo_fiscal())
    id_punto = punto.split(" — ")[0]
    circunstancia = circunstancia_inmediata_de(id_punto)
    st.text_area("Circunstancia inmediata (autocompletada del catálogo)", value=circunstancia or "", disabled=True)
    efecto_economico = st.text_area("Efecto económico potencial (redactar según el caso concreto de la entidad)")
    tratamiento = st.selectbox("Tratamiento del riesgo propuesto", tratamiento_riesgo_opciones())
    st.info(
        "Después de elegir el punto de riesgo fiscal aquí, vaya a la pestaña '1. Identificación', use estos "
        "datos para redactar el riesgo (Tipología = Fiscal) y continúe con los pasos 2 a 4 normalmente: el "
        "motor de cálculo es el mismo para todas las tipologías."
    )

with tab_amb:
    st.subheader("Pantalla 9 — Riesgos ambientales (Política de Gestión Ambiental Institucional, MIPG v7, 4.10)")
    st.caption(
        "Use estos campos ANTES del Paso 1 si la tipología del riesgo es 'Ambiental': "
        "seleccione el aspecto y el instrumento de origen, y luego complete la identificación "
        "en la pestaña 1 con Tipología = Ambiental."
    )
    aspecto = st.selectbox(
        "Aspecto ambiental",
        ["Consumo de agua", "Consumo de energía", "Consumo de papel", "Generación de residuos", "Otro (especificar en el nombre del riesgo)"],
    )
    instrumento = st.selectbox("Instrumento de origen", instrumentos_insumo_ambientales())
    riesgo_climatico = st.checkbox("¿Incluye riesgo climático asociado (Ley 1523 de 2012)?")
    if riesgo_climatico:
        st.text_area("Descripción del riesgo climático")
    st.info(
        "La valoración de impactos ambientales con la Metodología Conesa simplificada "
        "(Subanexo 3 de la política) requiere campos adicionales (naturaleza, intensidad, "
        "extensión, momento, persistencia, reversibilidad, entre otros) que no se incluyeron "
        "en este piloto porque el detalle completo del Subanexo 3 aún no se transcribió del "
        "Manual Operativo MIPG v7. Pendiente antes de la versión final."
    )

with tab_mapa:
    st.subheader("Mapa de riesgos consolidado (piloto)")
    if not st.session_state.riesgos_capturados:
        st.info("Todavía no se ha guardado ningún riesgo. Complete los pasos 1 a 4.")
    else:
        filas = []
        for r in st.session_state.riesgos_capturados:
            filas.append({
                "Proceso": r.proceso,
                "Riesgo": r.nombre,
                "Tipología": r.tipologia,
                "Prob. inherente": f"{r.inherente.probabilidad*100:.0f}%" if r.inherente else "-",
                "Impacto inherente": f"{r.inherente.impacto*100:.0f}%" if r.inherente else "-",
                "Severidad inherente": r.inherente.severidad if r.inherente else "-",
                "Prob. residual": f"{r.residual.probabilidad*100:.1f}%" if r.residual else "-",
                "Severidad residual": r.residual.severidad if r.residual else "-",
                "N.º controles": len(r.controles),
            })
        df = pd.DataFrame(filas)
        st.dataframe(df, use_container_width=True)

        st.download_button(
            "⬇️ Descargar mapa de riesgos (CSV)",
            df.to_csv(index=False).encode("utf-8"),
            file_name="mapa_riesgos_carepa_piloto.csv",
            mime="text/csv",
        )

        mapa = consolidar_mapa(st.session_state.riesgos_capturados)
        st.write("Agrupado por proceso y tipología:")
        for proceso, tipologias in mapa.items():
            st.markdown(f"**{proceso}**")
            for tip, riesgos in tipologias.items():
                st.write(f"- {tip}: {len(riesgos)} riesgo(s)")

    if not catalogo_puntos_riesgo_fiscal_disponible():
        st.divider()
        st.caption(
            "Nota: el catálogo oficial de puntos de riesgo fiscal (Anexo 3 de la Guía v7) "
            "todavía no está cargado en el sistema; las Pantallas 5-8 (fiscal, seguridad de "
            "la información, SIGRIP, KRI) no se incluyen en este piloto."
        )
