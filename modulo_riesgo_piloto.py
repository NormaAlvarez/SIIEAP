# NOTA (integración a app.py): este archivo ahora expone
# render_modulo_riesgo(), invocada desde el router principal app.py.
# Sigue funcionando también de forma independiente con:
#     streamlit run modulo_riesgo_piloto.py
"""
Módulo de Administración y Gestión del Riesgo — PILOTO SIIEAP
================================================================

Este archivo es un prototipo INDEPENDIENTE y NO modifica app.py ni ningún
otro archivo del SIIEAP v6.1. Se ejecuta por separado:

    streamlit run modulo_riesgo_piloto.py

Implementa las 9 pantallas: Procesos (caracterización), 1-4 (metodología
genérica), 5 (fiscal, 50 puntos oficiales del Anexo 3), 6 (seguridad de la
información, campos del Anexo 5), 7 (SIGRIP), 8 (KRI) y 9 (autodiagnóstico
de madurez ERM, 77 puntos del Anexo 4).

Validado por la docente/experta temática (Norma Elizabeth Álvarez Grajales)
el 19-sep-2026: las 9 pantallas del diseño original quedaron APROBADAS.
Instrucciones adicionales ya incorporadas:
  - NO se unificó el peso de "Manual" entre fuentes: la tipología
    'Seguridad de la información' usa 10% (Anexo 5); las demás usan 15%
    (Guía v7 / Anexo 1). El motor elige la tabla según la tipología.
  - Se agregó la pestaña "Procesos" para capturar la caracterización
    completa de cada proceso (objetivo, alcance, entradas, salidas,
    responsable, indicadores), más allá de los 4 tipos generales.
  - Se transcribió completo el Glosario oficial (Anexo 2, 46 términos, 5
    secciones) y se construyó la Pantalla de autodiagnóstico de madurez
    ERM con los 77 puntos de reflexión oficiales del Anexo 4.
  - Se programaron completas las Pantallas 6, 7 y 8.

ACTUALIZACIÓN (19-sep-2026, tarde): se recibió el Mapa de Procesos oficial
de la Alcaldía de Carepa (19 procesos, código C-ADM-xx-01, en 4 tipos:
Estratégico, Misional, De Apoyo, De Evaluación). Los 19 quedaron cargados
con su caracterización completa (objetivo, alcance, entradas, salidas,
responsable) en la pestaña "Procesos" y disponibles para identificar
riesgos desde la Pantalla 1 — el listado nominal que antes estaba
pendiente de la Alcaldía ya no lo está.

ACTUALIZACIÓN (19-sep-2026, noche): por instrucción de la docente/experta
temática, se RETIRÓ del piloto la tipología y la pantalla de riesgos
Ambientales (que se habían agregado tomando como fuente el numeral 4.10
del Manual Operativo MIPG v7, un documento distinto de la Guía para la
Gestión Integral del Riesgo v7). El módulo queda ceñido estrictamente al
alcance de la Guía v7 de riesgos (DAFP, agosto de 2025) y sus 5 anexos:
Gestión, Fiscal, Seguridad de la Información, Integridad Pública (SIGRIP)
y KRI — sin la tipología Ambiental.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd

from backend.base_conocimiento.catalogo_riesgos import (
    opciones_factor_riesgo,
    opciones_tipologia_riesgo,
    opciones_proceso_carepa,
    procesos_confirmados_carepa,
    procesos_confirmados_disponibles,
    catalogo_puntos_riesgo_fiscal_disponible,
    opciones_punto_riesgo_fiscal,
    circunstancia_inmediata_de,
    tratamiento_riesgo_opciones,
    campos_matriz_seguridad_informacion,
    esquema_caracterizacion_proceso,
    modelo_madurez_componentes,
    escala_madurez_erm,
    sugerencias_para,
)
from backend.modelos.riesgos import Riesgo, Control, KRI, ProcesoCaracterizado
from backend.motores.motor_riesgos import (
    valorar_riesgo_inherente,
    calcular_riesgo_residual,
    consolidar_mapa,
    evaluar_kri,
)



def render_modulo_riesgo() -> None:
    """Ejecuta las 10 pantallas + mapa consolidado del piloto de riesgo."""
    try:
        st.set_page_config(page_title="SIIEAP — Módulo de Riesgo (piloto)", layout="wide")
    except st.errors.StreamlitAPIException:
        pass  # ya fue configurada por app.py (router) u otra corrida previa


    if "riesgos_capturados" not in st.session_state:
        st.session_state.riesgos_capturados = []  # list[Riesgo]
    if "procesos_caracterizados" not in st.session_state:
        st.session_state.procesos_caracterizados = []  # list[ProcesoCaracterizado]
    if "kris_capturados" not in st.session_state:
        st.session_state.kris_capturados = []  # list[KRI]

    # Precarga (una sola vez por sesión) de los 19 procesos reales de Carepa,
    # tomados del Mapa de Procesos oficial (C-ADM-xx-01) entregado por la
    # docente/experta temática el 19-sep-2026. El usuario puede seguir
    # agregando o editando procesos con el formulario de abajo.
    if not st.session_state.get("_procesos_confirmados_precargados"):
        for p in procesos_confirmados_carepa():
            st.session_state.procesos_caracterizados.append(
                ProcesoCaracterizado(
                    codigo=p["codigo"], nombre=p["nombre"], tipo=p["tipo"],
                    objetivo=p.get("objetivo", ""), alcance=p.get("alcance", ""),
                    entradas=p.get("entradas", ""), salidas=p.get("salidas", ""),
                    responsable=p.get("responsable", ""),
                    indicadores=p.get("indicadores", []),
                )
            )
        st.session_state._procesos_confirmados_precargados = True


    def opciones_proceso_combinadas() -> list[str]:
        """Los 4 tipos generales confirmados + cualquier proceso ya
        caracterizado con nombre propio (pedido de la docente/experta
        temática, 19-sep-2026)."""
        base = opciones_proceso_carepa()
        extra = [f'{p.codigo} — {p.nombre} ({p.tipo})' for p in st.session_state.procesos_caracterizados]
        return base + extra

    st.title("🛡️ Módulo de Administración y Gestión del Riesgo — PILOTO")
    st.caption(
        "Prototipo independiente de SIIEAP v6.1. Caso de prueba: Alcaldía de Carepa "
        "(Antioquia), vigencia 2026-2027. No sustituye ni modifica el diagnóstico IDI-MIPG."
    )
    if procesos_confirmados_disponibles():
        st.success(
            f"✅ Los {len(procesos_confirmados_carepa())} procesos institucionales de la Alcaldía de "
            "Carepa (Mapa de Procesos oficial, código C-ADM-xx-01, entregado el 19-sep-2026) ya están "
            "cargados en la pestaña 'Procesos', con su objetivo, alcance, entradas, salidas y "
            "responsable. Puede usarlos directamente para identificar riesgos, o agregar/editar procesos "
            "adicionales con el formulario.",
            icon="✅",
        )
    else:
        st.warning(
            "⚠ Este es un PILOTO para validación funcional y de contenido con la docente/experta "
            "temática. Los procesos de Carepa que se listan abajo son solo los 4 TIPOS confirmados "
            "(Decreto 092 de 2021); el listado nominal de cada proceso específico está pendiente "
            "de verificar con la Alcaldía.",
            icon="⚠️",
        )

    tab_procesos, tab1, tab2, tab3, tab4, tab_fiscal, tab_seg, tab_sigrip, tab_kri, tab_madurez, tab_mapa = st.tabs([
        "Procesos",
        "1. Identificación",
        "2. Análisis inherente",
        "3. Controles",
        "4. Riesgo residual",
        "5. Fiscal",
        "6. Seguridad Info",
        "7. SIGRIP",
        "8. KRI",
        "9. Madurez ERM",
        "Mapa consolidado",
    ])

    with tab_procesos:
        st.subheader("Caracterización de procesos institucionales")
        if procesos_confirmados_disponibles():
            st.caption(
                "Los 19 procesos oficiales de la Alcaldía de Carepa (Mapa de Procesos, código C-ADM-xx-01) "
                "ya están precargados abajo, con su objetivo, alcance, entradas, salidas y responsable "
                "tomados literalmente de la caracterización entregada por la Alcaldía. Puede agregar más "
                "procesos o ajustar alguno con el formulario. Campo agregado a petición de la docente/"
                "experta temática (validación 19-sep-2026)."
            )
        else:
            st.caption(
                "Además de los 4 tipos generales confirmados (Decreto 092 de 2021), aquí se puede capturar la "
                "caracterización completa de cada proceso individual de la entidad, en cuanto la Alcaldía la entregue "
                "o se levante con los líderes de proceso. Campo agregado a petición de la docente/experta temática "
                "(validación 19-sep-2026)."
            )
        esquema = esquema_caracterizacion_proceso()
        with st.expander("Ver los campos de la caracterización (Anexo del diseño técnico)"):
            for c in esquema["campos"]:
                st.write(f"**{c['campo']}** — {c['descripcion']}")

        # Selección hecha en la tabla de más abajo (clic en una fila): se lee
        # ANTES de dibujar el formulario para que sus campos se llenen solos
        # con los datos de ese proceso, sin tener que copiar y pegar nada.
        # (Streamlit no distingue clic sencillo de doble clic en una tabla;
        # esta es la forma más confiable de lograr "seleccionar y que se
        # llene todo": un clic en la fila ya autocompleta el formulario.)
        seleccion_tabla = st.session_state.get("tabla_procesos_sel")
        proceso_para_editar = None
        if seleccion_tabla:
            filas_sel = seleccion_tabla.get("selection", {}).get("rows", [])
            if filas_sel:
                idx_sel = filas_sel[0]
                lista_actual = st.session_state.procesos_caracterizados
                if 0 <= idx_sel < len(lista_actual):
                    proceso_para_editar = lista_actual[idx_sel]

        if proceso_para_editar:
            st.info(
                f"✏️ Editando **{proceso_para_editar.codigo} — {proceso_para_editar.nombre}** (se "
                "autocompletó al seleccionarlo en la tabla de abajo). Ajuste lo que necesite y presione "
                "'Guardar cambios'. Para capturar un proceso nuevo en blanco, presione 'Nuevo proceso en "
                "blanco'.",
                icon="✏️",
            )
            if st.button("🆕 Nuevo proceso en blanco (quitar selección)"):
                st.session_state["tabla_procesos_sel"] = None
                st.rerun()

        with st.form("form_proceso", clear_on_submit=False):
            col1, col2 = st.columns(2)
            codigo = col1.text_input("Código del proceso", value=proceso_para_editar.codigo if proceso_para_editar else "")
            nombre_p = col2.text_input("Nombre del proceso", value=proceso_para_editar.nombre if proceso_para_editar else "")
            opciones_tipo = ["Estratégico", "Misional", "De Apoyo", "De Evaluación"]
            tipo_p = st.selectbox(
                "Tipo", opciones_tipo,
                index=opciones_tipo.index(proceso_para_editar.tipo) if proceso_para_editar and proceso_para_editar.tipo in opciones_tipo else 0,
            )
            objetivo_p = st.text_area("Objetivo del proceso", value=proceso_para_editar.objetivo if proceso_para_editar else "")
            alcance_p = st.text_area("Alcance del proceso", value=proceso_para_editar.alcance if proceso_para_editar else "")
            col3, col4 = st.columns(2)
            entradas_p = col3.text_area("Entradas / insumos", value=proceso_para_editar.entradas if proceso_para_editar else "")
            salidas_p = col4.text_area("Salidas / productos", value=proceso_para_editar.salidas if proceso_para_editar else "")
            responsable_p = st.text_input("Responsable (líder de proceso)", value=proceso_para_editar.responsable if proceso_para_editar else "")
            indicadores_p = st.text_area(
                "Indicadores de gestión (uno por línea)",
                value="\n".join(proceso_para_editar.indicadores) if proceso_para_editar else "",
            )
            texto_boton = "💾 Guardar cambios" if proceso_para_editar else "+ Agregar proceso caracterizado"
            guardar_proceso = st.form_submit_button(texto_boton, type="primary")
            if guardar_proceso and codigo and nombre_p:
                nuevo_proceso = ProcesoCaracterizado(
                    codigo=codigo, nombre=nombre_p, tipo=tipo_p, objetivo=objetivo_p, alcance=alcance_p,
                    entradas=entradas_p, salidas=salidas_p, responsable=responsable_p,
                    indicadores=[i.strip() for i in indicadores_p.splitlines() if i.strip()],
                )
                lista = st.session_state.procesos_caracterizados
                codigos_existentes = [p.codigo for p in lista]
                if codigo in codigos_existentes:
                    lista[codigos_existentes.index(codigo)] = nuevo_proceso
                    st.success(f"Proceso {codigo} actualizado.")
                else:
                    lista.append(nuevo_proceso)
                    st.success(f"Proceso {codigo} agregado.")
                st.session_state["tabla_procesos_sel"] = None
                st.rerun()

        if st.session_state.procesos_caracterizados:
            st.write(
                f"Procesos caracterizados hasta ahora ({len(st.session_state.procesos_caracterizados)}). "
                "Haga clic en una fila para cargarla en el formulario de arriba y editarla:"
            )
            st.dataframe(
                pd.DataFrame([{
                    "Código": p.codigo, "Nombre": p.nombre, "Tipo": p.tipo,
                    "Responsable": p.responsable, "N.º indicadores": len(p.indicadores),
                } for p in st.session_state.procesos_caracterizados]),
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row",
                key="tabla_procesos_sel",
            )
        else:
            st.info("Todavía no se ha caracterizado ningún proceso con nombre propio. Mientras tanto, la "
                    "identificación de riesgos (pestaña 1) puede usar los 4 tipos generales.")

    # ---------------------------------------------------------------------
    # Estado temporal del riesgo en construcción
    # ---------------------------------------------------------------------
    if "riesgo_en_edicion" not in st.session_state:
        st.session_state.riesgo_en_edicion = {}

    with tab1:
        st.subheader("Paso 1 — Identificación del riesgo")

        # Banco de riesgos sugeridos (PROPUESTA DE APOYO, no oficial): si ya
        # hay un proceso y un factor elegidos de una vez anterior, se ofrece
        # la sugerencia correspondiente con un botón para prellenar el
        # formulario de abajo (nombre, causa raíz, evento no deseado).
        proceso_previo = st.session_state.get("f_proceso")
        factor_previo = st.session_state.get("f_factor")
        if proceso_previo and factor_previo:
            codigo_proceso_previo = proceso_previo.split(" — ")[0].strip()
            codigo_factor_previo = factor_previo.split(" — ")[0].strip()
            sugerencias = sugerencias_para(codigo_proceso_previo, codigo_factor_previo)
            if sugerencias:
                sug = sugerencias[0]
                with st.expander(
                    "💡 Sugerencia de riesgo para este proceso y factor (banco de apoyo — NO es catálogo del "
                    "DAFP, valide y ajuste antes de usarla)",
                    expanded=False,
                ):
                    st.caption(
                        "Propuesta de apoyo construida cruzando los 19 procesos reales de Carepa con los 6 "
                        "factores de riesgo de la Tabla 2 (Guía v7). No sustituye el análisis del equipo del "
                        "proceso."
                    )
                    st.markdown(f"**Riesgo sugerido:** {sug['riesgo_sugerido']}")
                    st.markdown(f"**Posible causa raíz:** {sug['causa_raiz_sugerida']}")
                    st.markdown(f"**Posible evento no deseado:** {sug['evento_no_deseado_sugerido']}")
                    if st.button("↳ Usar esta sugerencia para prellenar el formulario", key="btn_usar_sugerencia"):
                        st.session_state["f_nombre"] = sug["riesgo_sugerido"]
                        st.session_state["f_causa_raiz"] = sug["causa_raiz_sugerida"]
                        st.session_state["f_evento"] = sug["evento_no_deseado_sugerido"]
                        st.rerun()

        col1, col2 = st.columns(2)
        with col1:
            proceso = st.selectbox("Proceso institucional (Carepa)", opciones_proceso_combinadas(), key="f_proceso")
            nombre = st.text_input("Nombre del riesgo", key="f_nombre")
            factor = st.selectbox("Factor de riesgo (Tabla 2, Guía v7)", opciones_factor_riesgo(), key="f_factor")
            tipologia = st.selectbox("Tipología del riesgo", opciones_tipologia_riesgo(), key="f_tipologia")
        with col2:
            causa_raiz = st.text_area("Causa raíz", key="f_causa_raiz")
            causa_inmediata = st.text_area("Causa inmediata", key="f_causa_inmediata")
            evento = st.text_area("Evento no deseado", key="f_evento")
            impacto_desc = st.text_area("Impacto (descripción cualitativa)", key="f_impacto_desc")

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

    with tab_seg:
        st.subheader("Pantalla 6 — Riesgo de seguridad de la información (Capítulo V, campos oficiales Anexo 5)")
        st.caption(
            "Use esta pestaña ANTES del Paso 1 si la tipología del riesgo es 'Seguridad de la información'. "
            "Nota: esta tipología usa su PROPIA tabla de pesos de controles (Manual = 10%, según el Anexo 5), "
            "distinta de las demás tipologías (Manual = 15%, Guía v7/Anexo 1) — el motor la aplica automáticamente."
        )
        with st.expander("Ver los campos oficiales de la matriz (Anexo 5, DAFP)"):
            for c in campos_matriz_seguridad_informacion():
                st.write(f"- {c}")
        col1, col2 = st.columns(2)
        activo = col1.text_input("Activo de información")
        tipo_activo = col2.selectbox("Tipo de activo", ["Información", "Software/Aplicación", "Hardware", "Servicio", "Personal", "Instalación/Sede"])
        amenaza = st.text_area("Amenazas (Causa Inmediata)")
        vulnerabilidad = st.text_area("Vulnerabilidades (Causa raíz)")
        clasificacion = st.selectbox("Clasificación de la información", ["Pública", "Pública clasificada", "Pública reservada"])
        control_anexo_a = st.text_input("Control de referencia (Anexo A - ISO/IEC 27001), si aplica")
        st.info(
            "Después de diligenciar esta pestaña, vaya a '1. Identificación', use Tipología = 'Seguridad de la "
            "información' y redacte el riesgo retomando el activo, la amenaza y la vulnerabilidad de aquí. "
            "Continúe con los pasos 2 a 4 normalmente."
        )

    with tab_sigrip:
        st.subheader("Pantalla 7 — Riesgos para la integridad pública — SIGRIP (Capítulo VI)")
        st.caption(
            "Use esta pestaña ANTES del Paso 1 si la tipología del riesgo es 'Integridad pública (SIGRIP)'."
        )
        amenaza_integridad = st.selectbox(
            "Amenaza de integridad asociada",
            ["Soborno entrante", "Soborno saliente", "Fraude interno", "Fraude externo", "Conflicto de interés",
             "Corrupción", "Lavado de activos (LA)", "Financiación del terrorismo (FT)",
             "Financiación de la proliferación de armas de destrucción masiva (FP)", "Otra"],
        )
        contraparte = st.text_input("Contraparte relacionada (si aplica debida diligencia sobre terceros)")
        tipo_contraparte = st.selectbox("Tipo de contraparte (si aplica)", ["No aplica", "Parte Interesada", "Parte Vinculada", "Parte Relacionada"])
        vinculo_paac = st.text_input("Vínculo con el mapa de riesgos de corrupción (PAAC, componente 1), si existe")
        st.caption(
            "Definiciones de estos términos (soborno, fraude, conflicto de interés, LA/FT/FP, contraparte, etc.) "
            "están disponibles en el Glosario oficial (Anexo 2, sección IV) cargado en el catálogo del módulo."
        )
        st.info(
            "Después de diligenciar esta pestaña, vaya a '1. Identificación', use Tipología = 'Integridad pública "
            "(SIGRIP)' y redacte el riesgo. Continúe con los pasos 2 a 4 normalmente."
        )

    with tab_kri:
        st.subheader("Pantalla 8 — Indicadores clave de riesgo (KRI, Capítulo VIII)")
        st.caption("Un KRI se asocia a un riesgo YA GUARDADO en el mapa consolidado (complete primero los pasos 1 a 4).")
        if not st.session_state.riesgos_capturados:
            st.info("Todavía no hay riesgos guardados. Complete al menos un riesgo (pestañas 1 a 4) antes de crear un KRI.")
        else:
            opciones_riesgo = [r.nombre for r in st.session_state.riesgos_capturados]
            with st.form("form_kri", clear_on_submit=True):
                riesgo_asociado = st.selectbox("Riesgo asociado", opciones_riesgo)
                nombre_kri = st.text_input("Nombre del KRI")
                col1, col2, col3 = st.columns(3)
                umbral_alerta = col1.number_input("Umbral de alerta", min_value=0.0, step=1.0)
                umbral_apetito = col2.number_input("Umbral de apetito al riesgo", min_value=0.0, step=1.0)
                valor_actual = col3.number_input("Valor actual del KRI", min_value=0.0, step=1.0)
                responsable_kri = st.text_input("Responsable del KRI")
                guardar_kri = st.form_submit_button("+ Agregar KRI", type="primary")
                if guardar_kri and nombre_kri:
                    st.session_state.kris_capturados.append(
                        KRI(riesgo_asociado=riesgo_asociado, nombre=nombre_kri, umbral_alerta=umbral_alerta,
                            umbral_apetito=umbral_apetito, valor_actual=valor_actual, responsable=responsable_kri)
                    )

            if st.session_state.kris_capturados:
                st.write("KRI registrados:")
                filas_kri = []
                for k in st.session_state.kris_capturados:
                    semaforo = evaluar_kri(k)
                    icono = {"verde": "🟢", "amarillo": "🟡", "rojo": "🔴"}[semaforo]
                    filas_kri.append({
                        "Riesgo": k.riesgo_asociado, "KRI": k.nombre, "Umbral alerta": k.umbral_alerta,
                        "Umbral apetito": k.umbral_apetito, "Valor actual": k.valor_actual,
                        "Semáforo": f"{icono} {semaforo}", "Responsable": k.responsable,
                    })
                st.dataframe(pd.DataFrame(filas_kri), use_container_width=True)

    with tab_madurez:
        st.subheader("Pantalla 9 — Autodiagnóstico de madurez de la gestión del riesgo (Anexo 4, marco COSO-ERM)")
        st.caption(
            "77 puntos de reflexión oficiales, agrupados en 5 componentes COSO-ERM. Para cada punto, seleccione "
            "qué tanto se cumple en la entidad. Al final se calcula el promedio de madurez por componente."
        )
        escala = escala_madurez_erm()
        etiquetas_escala = [e["etiqueta"] for e in escala]
        valor_de_etiqueta = {e["etiqueta"]: e["valor"] for e in escala}

        if "respuestas_madurez" not in st.session_state:
            st.session_state.respuestas_madurez = {}

        componentes = modelo_madurez_componentes()
        componente_sel = st.selectbox(
            "Componente a diligenciar",
            [f'{c["numero"]}. {c["nombre"]} ({len(c["puntos_reflexion"])} puntos)' for c in componentes],
        )
        idx_componente = int(componente_sel.split(".")[0]) - 1
        comp = componentes[idx_componente]

        with st.form(f"form_madurez_{idx_componente}"):
            for i, p in enumerate(comp["puntos_reflexion"]):
                key = f'madurez_{comp["numero"]}_{i}'
                valor_previo = st.session_state.respuestas_madurez.get(key, "3.- A veces")
                st.session_state.respuestas_madurez[key] = st.select_slider(
                    p["punto_reflexion"], options=etiquetas_escala,
                    value=valor_previo if valor_previo in etiquetas_escala else "3.- A veces",
                    key=f"widget_{key}",
                )
            st.form_submit_button("Guardar respuestas de este componente")

        st.divider()
        st.write("**Resultado del autodiagnóstico (promedio por componente, sobre 5):**")
        filas_madurez = []
        for c in componentes:
            valores = []
            for i in range(len(c["puntos_reflexion"])):
                key = f'madurez_{c["numero"]}_{i}'
                if key in st.session_state.respuestas_madurez:
                    valores.append(valor_de_etiqueta[st.session_state.respuestas_madurez[key]])
            promedio = sum(valores) / len(valores) if valores else None
            filas_madurez.append({
                "Componente": c["nombre"],
                "Puntos respondidos": f"{len(valores)}/{len(c['puntos_reflexion'])}",
                "Promedio de madurez (1-5)": f"{promedio:.2f}" if promedio else "Sin responder",
            })
        st.dataframe(pd.DataFrame(filas_madurez), use_container_width=True)

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


if __name__ == "__main__":
    render_modulo_riesgo()
