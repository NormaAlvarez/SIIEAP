"""
SIIEAP — punto de entrada de la interfaz (Streamlit).

Ejecutar con:
    streamlit run app.py

V2: agrega la carga de entidades REALES desde los archivos oficiales de
Función Pública (Resultados_vigXXXX_nacion.xlsx / _territorio.xlsx),
además del modo de captura manual de V1.
"""
import streamlit as st
from pathlib import Path

from backend.base_conocimiento.catalogo import todos_los_indices, dimensiones
from backend.base_conocimiento.cargar_resultados_oficiales import (
    cargar_hoja,
    listar_entidades,
    entidad_por_nombre_exacto,
)
from backend.base_conocimiento.cargar_recomendaciones import (
    cargar_recomendaciones,
    cargar_consolidado,
    recomendaciones_de_entidad,
    cruzar_brechas_con_recomendaciones,
)
from backend.modelos.entidades import Entidad, ResultadoIndice
from backend.motores.motor_diagnostico import diagnosticar
from backend.motores.motor_analisis_360 import analizar_360, filtrar_grupo, subregiones_disponibles_para
from backend.motores.motor_isvpt import calcular_isvpt
from backend.motores.estudio_de_caso import generar_estudio_de_caso_docx, generar_estudio_de_caso_pdf
from backend.motores.informe_alcaldes_gobernadores import generar_informe_alcaldes_docx, generar_informe_alcaldes_pdf
from backend.motores.generador_informe import (
    REGIMEN_ESPECIAL_NINGUNO,
    REGIMEN_ESPECIAL_UNIVERSIDAD_AUTONOMA,
    REGIMEN_ESPECIAL_ORGANO_CONTROL,
    REGIMEN_ESPECIAL_BANCO_REPUBLICA,
    REGIMEN_ESPECIAL_CORPORACION_AUTONOMA,
    REGIMEN_ESPECIAL_CONCEJO_ASAMBLEA,
    REGIMEN_ESPECIAL_PERSONERIA,
)
from backend.base_conocimiento.subregiones_antioquia import todas_las_subregiones

CARPETA_DATA = Path(__file__).resolve().parent / "data"
RUTA_AUTO_NACION = CARPETA_DATA / "resultados_nacion.xlsx"
RUTA_AUTO_TERRITORIO = CARPETA_DATA / "resultados_territorio.xlsx"
RUTA_AUTO_RECOMENDACIONES = CARPETA_DATA / "recomendaciones_consolidado.json"
RUTA_AUTO_RECOMENDACIONES_GZ = CARPETA_DATA / "recomendaciones_consolidado.json.gz"

OPCIONES_REGIMEN_ESPECIAL = {
    "Rama Ejecutiva — MIPG íntegro aplica (alcaldías, gobernaciones, EICE, ESE...)": REGIMEN_ESPECIAL_NINGUNO,
    "Ente universitario autónomo (ej. Universidad de Antioquia)": REGIMEN_ESPECIAL_UNIVERSIDAD_AUTONOMA,
    "Contraloría territorial (departamental/municipal/distrital)": REGIMEN_ESPECIAL_ORGANO_CONTROL,
    "Personería municipal o distrital": REGIMEN_ESPECIAL_PERSONERIA,
    "Concejo municipal o Asamblea departamental": REGIMEN_ESPECIAL_CONCEJO_ASAMBLEA,
    "Banco de la República": REGIMEN_ESPECIAL_BANCO_REPUBLICA,
    "Corporación Autónoma Regional": REGIMEN_ESPECIAL_CORPORACION_AUTONOMA,
}

st.set_page_config(page_title="SIIEAP — Diagnóstico IDI-MIPG", layout="wide")


def verificar_acceso():
    """Puerta simple de usuario/clave. La clave se guarda en st.secrets['APP_PASSWORD']
    (nunca en el código). Si no hay clave configurada, deja pasar (modo local/pruebas)."""
    clave_esperada = st.secrets.get("APP_PASSWORD") if hasattr(st, "secrets") else None
    if not clave_esperada:
        return True  # sin clave configurada (uso local) -> no bloquear

    if st.session_state.get("acceso_concedido"):
        return True

    st.title("SIIEAP — Acceso")
    clave = st.text_input("Clave de acceso", type="password")
    if st.button("Ingresar"):
        if clave == clave_esperada:
            st.session_state.acceso_concedido = True
            st.rerun()
        else:
            st.error("Clave incorrecta.")
    return False


if not verificar_acceso():
    st.stop()

st.title("Modelo de Conocimiento Institucional del Sistema de Inteligencia Artificial para la Evaluación Integral del Desempeño Institucional en Entidades Públicas (SIIEAP)")
st.caption(
    "Docente: Norma Elizabeth Álvarez Grajales · Área del conocimiento "
    "Organizaciones Públicas y Gestión. Escuela Superior de Administración Pública (ESAP)"
)
st.caption(
    "Catálogo oficial: "
    f"{len(dimensiones())} dimensiones, 19 políticas, {len(todos_los_indices())} índices."
)


def mostrar_diagnostico(diag, idi_oficial=None, grupo_par=None, cruce_recomendaciones=None):
    """Renderiza el resultado de diagnosticar(), reutilizable en ambos modos."""
    col1, col2 = st.columns(2)
    col1.metric("IDI estimado por SIIEAP", diag.idi_estimado)
    if idi_oficial is not None:
        col2.metric(
            "IDI oficial publicado por Función Pública",
            idi_oficial,
            delta=round((diag.idi_estimado or 0) - idi_oficial, 2),
        )
    if grupo_par:
        st.caption(f"Grupo par: {grupo_par}")

    st.markdown("#### Resultado por dimensión")
    tabla = [
        {
            "Dimensión": r.nombre,
            "Promedio": r.promedio,
            "Nivel de riesgo": r.nivel_riesgo,
            "Índices/políticas evaluados": f"{r.n_indices_evaluados}/{r.n_indices_esperados}",
        }
        for r in diag.resultados_por_dimension
    ]
    st.dataframe(tabla, use_container_width=True)

    st.markdown("#### Brechas priorizadas (menor puntaje primero)")
    if diag.brechas:
        for b in diag.brechas:
            recos_b = (cruce_recomendaciones or {}).get(b.codigo_indice, [])
            with st.expander(f"🔴 {b.codigo_indice} ({b.puntaje}) — {b.nombre_indice} · {b.politica}"):
                if recos_b:
                    st.markdown(f"**{len(recos_b)} recomendación(es) oficial(es) para esta brecha:**")
                    for texto in recos_b:
                        st.markdown(f"- {texto}")
                else:
                    st.caption(
                        "Sin recomendación oficial cargada para esta política todavía "
                        "(suba el archivo de recomendaciones de la entidad para verlas)."
                    )
    else:
        st.success("No se detectaron brechas por debajo del umbral con los datos ingresados.")


with st.sidebar:
    st.markdown("### Base de recomendaciones")

    if "consolidado_cargado" not in st.session_state:
        ruta_auto = None
        if RUTA_AUTO_RECOMENDACIONES.exists():
            ruta_auto = RUTA_AUTO_RECOMENDACIONES
        elif RUTA_AUTO_RECOMENDACIONES_GZ.exists():
            ruta_auto = RUTA_AUTO_RECOMENDACIONES_GZ

        if ruta_auto is not None:
            with st.spinner(f"Cargando {ruta_auto.name} desde data/..."):
                st.session_state.consolidado_cargado = cargar_consolidado(ruta_auto)
            st.success(f"Cargado automáticamente: {len(st.session_state.consolidado_cargado)} entidades ({ruta_auto.name}).")

    if "consolidado_cargado" not in st.session_state:
        st.caption(
            "No se encontró data/recomendaciones_consolidado.json. Puede colocarlo ahí "
            "para que cargue solo, o subirlo aquí manualmente esta sesión."
        )
        archivo_consolidado = st.file_uploader(
            "Consolidado de recomendaciones (.json o .json.gz)", type=["json", "gz"], key="consolidado"
        )
        if archivo_consolidado is not None:
            with st.spinner("Cargando base consolidada..."):
                st.session_state.consolidado_cargado = cargar_consolidado(archivo_consolidado)
            st.success(f"{len(st.session_state.consolidado_cargado)} entidades cargadas en memoria.")
    else:
        st.success(f"{len(st.session_state.consolidado_cargado)} entidades ya cargadas en memoria.")

tab_real, tab_manual = st.tabs(["📊 Cargar entidad real (Excel oficial)", "✍️ Captura manual"])

# ----------------------------------------------------------------------
# MODO 1: cargar entidad real desde el Excel oficial de Función Pública
# ----------------------------------------------------------------------
with tab_real:
    st.markdown(
        "Suba el archivo oficial de resultados de Función Pública "
        "(`Resultados_vigXXXX_nacion.xlsx` o `_territorio.xlsx`, hoja 'Nacion' o 'Territorio')."
    )

    fuentes_disponibles = {}
    if RUTA_AUTO_NACION.exists():
        fuentes_disponibles["Nación (data/resultados_nacion.xlsx)"] = (RUTA_AUTO_NACION, "Nacion")
    if RUTA_AUTO_TERRITORIO.exists():
        fuentes_disponibles["Territorio (data/resultados_territorio.xlsx)"] = (RUTA_AUTO_TERRITORIO, "Territorio")

    df = None
    if fuentes_disponibles:
        opcion = st.selectbox(
            "Fuente detectada automáticamente en data/",
            list(fuentes_disponibles.keys()) + ["Subir otro archivo..."],
        )
        if opcion != "Subir otro archivo...":
            ruta_sel, hoja_sel = fuentes_disponibles[opcion]
            try:
                df = cargar_hoja(ruta_sel, hoja_sel)
                st.caption(f"{len(df)} entidades cargadas automáticamente desde {ruta_sel.name}.")
            except Exception as e:
                st.error(f"Error leyendo {ruta_sel.name}: {e}")

    if df is None:
        col_a, col_b = st.columns(2)
        archivo = col_a.file_uploader("Archivo de resultados (.xlsx)", type=["xlsx"], key="archivo_real")
        nombre_hoja = col_b.selectbox("Hoja del Excel", ["Nacion", "Territorio"], key="hoja_real")
        if archivo is not None:
            try:
                df = cargar_hoja(archivo, nombre_hoja)
            except Exception as e:
                st.error(f"No se pudo leer el archivo con la hoja '{nombre_hoja}': {e}")
                df = None

    if df is not None:
        nombres = listar_entidades(df)
        st.caption(f"{len(nombres)} entidades encontradas en el archivo.")
        entidad_elegida = st.selectbox(
            "Busque y seleccione la entidad", nombres, key="entidad_real", index=None,
            placeholder="Escriba para buscar...",
        )

        archivo_recos = None
        if "consolidado_cargado" not in st.session_state:
            archivo_recos = st.file_uploader(
                "(Opcional) Archivo de recomendaciones de ESA MISMA entidad, si no cargó "
                "el consolidado en la barra lateral",
                type=["xlsx"], key="archivo_recos",
            )

        etiqueta_regimen_elegida = st.selectbox(
            "Tipo de entidad (régimen especial MIPG/MECI)",
            list(OPCIONES_REGIMEN_ESPECIAL.keys()),
            key="tipo_regimen_especial_select",
            help=(
                "Si la entidad NO pertenece a la Rama Ejecutiva (universidades autónomas, "
                "órganos de control territoriales, Banco de la República, Corporaciones "
                "Autónomas Regionales), el MIPG no le aplica en su integralidad — solo la "
                "política de Control Interno (MECI). Selecciónelo aquí para que los 3 "
                "informes incluyan la nota jurídica correcta (art. 40 Ley 489/1998 y "
                "Decreto 1499/2017) en vez de leer las demás dimensiones como brechas."
            ),
        )
        tipo_regimen_especial_elegido = OPCIONES_REGIMEN_ESPECIAL[etiqueta_regimen_elegida]

        if entidad_elegida and st.button("Generar diagnóstico", type="primary", key="btn_real"):
            entidad, idi_oficial, grupo_par = entidad_por_nombre_exacto(df, entidad_elegida)
            entidad.regimen_especial = tipo_regimen_especial_elegido
            diag = diagnosticar(entidad)

            cruce = None
            total_recos_entidad = None
            if "consolidado_cargado" in st.session_state:
                recos = recomendaciones_de_entidad(st.session_state.consolidado_cargado, entidad_elegida)
                if recos:
                    cruce = cruzar_brechas_con_recomendaciones(diag.brechas, recos)
                    total_recos_entidad = len(recos)
            elif archivo_recos is not None:
                try:
                    recos = cargar_recomendaciones(archivo_recos)
                    cruce = cruzar_brechas_con_recomendaciones(diag.brechas, recos)
                    total_recos_entidad = len(recos)
                except Exception as e:
                    st.warning(f"No se pudieron leer las recomendaciones: {e}")

            # Se guarda en session_state para que TODO lo que sigue (el propio
            # diagnóstico, el Análisis 360 y sus resultados) sobreviva a los
            # reruns que dispara Streamlit al tocar cualquier otro control
            # (el selector de Departamento/Subregión, el botón "Calcular
            # análisis 360", etc.). Antes, todo este bloque dependía del
            # valor de un solo clic de st.button("Generar diagnóstico"), que
            # vuelve a False en el siguiente rerun — por eso la pantalla
            # "volvía" al formulario inicial al interactuar con el selector
            # de departamento/provincia (o con cualquier otro control de esta
            # sección, incluido "Calcular análisis 360" mismo).
            st.session_state["resultado_real_actual"] = {
                "entidad": entidad,
                "idi_oficial": idi_oficial,
                "grupo_par": grupo_par,
                "diag": diag,
                "cruce": cruce,
                "total_recos_entidad": total_recos_entidad,
                "etiqueta_regimen_elegida": etiqueta_regimen_elegida,
                "tipo_regimen_especial_elegido": tipo_regimen_especial_elegido,
            }

        if "resultado_real_actual" in st.session_state:
            _r = st.session_state["resultado_real_actual"]
            entidad = _r["entidad"]
            idi_oficial = _r["idi_oficial"]
            grupo_par = _r["grupo_par"]
            diag = _r["diag"]
            cruce = _r["cruce"]
            total_recos_entidad = _r["total_recos_entidad"]
            etiqueta_regimen_elegida = _r["etiqueta_regimen_elegida"]
            tipo_regimen_especial_elegido = _r["tipo_regimen_especial_elegido"]

            st.subheader(entidad.nombre)
            if not entidad.aplica_mipg_integral():
                st.info(
                    f"⚠️ Régimen especial: **{etiqueta_regimen_elegida}** — las brechas y el "
                    "semáforo de este diagnóstico solo consideran la política de Control "
                    "Interno (MECI) como exigible; las demás se muestran de forma informativa "
                    "pero no se marcan como brecha."
                )
            st.caption(
                f"Índices/políticas con información: {len(entidad.resultados) + len(entidad.resultados_politica_directa)} de 66"
            )
            if total_recos_entidad is not None:
                st.caption(f"{total_recos_entidad} recomendaciones encontradas para esta entidad.")

            mostrar_diagnostico(diag, idi_oficial, grupo_par, cruce)

            st.markdown("---")
            st.markdown("### 📈 Análisis 360 (comparación contra el grupo)")
            st.caption(
                "Compara esta entidad contra un grupo de referencia, usando el 'Grupo par' "
                "oficial de Función Pública (ya calculado en el Excel) y, opcionalmente, la "
                "subregión del departamento (disponible para Antioquia, Chocó y Santander). No inventa "
                "cifras: todo sale de las columnas reales del archivo cargado arriba."
            )
            col_360_a, col_360_b, col_360_c = st.columns(3)
            opciones_departamento = ["ANTIOQUIA", "CHOCÓ", "SANTANDER", "Otro (escribir abajo)"]
            departamento_elegido = col_360_a.selectbox(
                "Departamento", opciones_departamento, key="departamento_360_select"
            )
            if departamento_elegido == "Otro (escribir abajo)":
                departamento_360 = col_360_a.text_input(
                    "Escriba el departamento", value="", key="departamento_360_manual"
                )
                st.caption(
                    "ℹ️ Este departamento no tiene subregiones registradas todavía en el "
                    "sistema — el análisis seguirá funcionando por Departamento y Grupo par, "
                    "solo sin el filtro adicional de subregión. Para agregar la "
                    "subregionalización de un nuevo departamento, avísele a quien administra "
                    "el sistema."
                )
            else:
                departamento_360 = departamento_elegido

            st.session_state["departamento_360_actual"] = departamento_360

            opciones_subregion = ["(ninguna)"] + subregiones_disponibles_para(departamento_360)
            subregion_360 = col_360_b.selectbox(
                f"Subregión de {departamento_360 or '—'} (opcional)", opciones_subregion, key="subregion_360"
            )
            grupo_par_360 = col_360_c.text_input(
                "Grupo par contiene (opcional)",
                value=grupo_par or "",
                key="grupo_par_360",
                help="Ej.: 'ALCALDÍA GRUPO 4'. Se deja vacío para no filtrar por grupo par.",
            )
            if st.button("Calcular análisis 360", key="btn_360"):
                try:
                    resultado_360 = analizar_360(
                        df,
                        departamento=departamento_360 or None,
                        subregion=None if subregion_360 == "(ninguna)" else subregion_360,
                        grupo_par_contiene=grupo_par_360 or None,
                        entidad_referencia=entidad_elegida,
                    )
                    st.caption(f"Filtro aplicado: {resultado_360.filtro_descripcion}")
                    col_r1, col_r2, col_r3 = st.columns(3)
                    col_r1.metric("Entidades en el grupo", resultado_360.n_entidades)
                    col_r2.metric("IDI promedio del grupo", resultado_360.promedio_idi)
                    if resultado_360.percentil_entidad_referencia is not None:
                        col_r3.metric(
                            f"Percentil de {entidad_elegida.split()[-1] if entidad_elegida else ''}",
                            f"{resultado_360.percentil_entidad_referencia}%",
                            delta=round(
                                (resultado_360.idi_entidad_referencia or 0) - (resultado_360.promedio_idi or 0), 2
                            ),
                        )

                    if resultado_360.promedio_por_dimension:
                        st.markdown("#### Promedio del grupo por dimensión")
                        st.dataframe(
                            [{"Dimensión": k, "Promedio del grupo": v} for k, v in resultado_360.promedio_por_dimension.items()],
                            use_container_width=True,
                        )

                    col_top, col_bottom = st.columns(2)
                    with col_top:
                        st.markdown("#### Top 5 del grupo")
                        for nombre_e, idi_e in resultado_360.top5:
                            st.markdown(f"- {nombre_e}: **{idi_e}**")
                    with col_bottom:
                        st.markdown("#### Últimos 5 del grupo")
                        for nombre_e, idi_e in resultado_360.bottom5:
                            st.markdown(f"- {nombre_e}: **{idi_e}**")

                    st.session_state.ultimo_analisis_360 = resultado_360

                    # Índice Sintético de Valor Público Territorial (ISVPT) — novedad metodológica
                    df_grupo_isvpt, _desc_isvpt = filtrar_grupo(
                        df,
                        departamento=departamento_360 or None,
                        subregion=None if subregion_360 == "(ninguna)" else subregion_360,
                        grupo_par_contiene=grupo_par_360 or None,
                    )
                    resultado_isvpt = calcular_isvpt(df_grupo_isvpt, entidad_referencia=entidad_elegida)
                    st.markdown("---")
                    st.markdown("#### 🧮 Índice Sintético de Valor Público Territorial (ISVPT) — novedad")
                    st.caption(
                        "Normaliza (min-max) las 7 dimensiones del IDI-MIPG dentro de ESTE grupo de "
                        "comparación y las agrega en un solo valor relativo (0 a 1), siguiendo la "
                        "metodología académica del ISDEL (Vélez Tamayo et al., 2026) y las directrices "
                        "de la OCDE (2008) para indicadores compuestos."
                    )
                    if resultado_isvpt.isvpt_entidad_referencia is not None:
                        col_isvpt_a, col_isvpt_b = st.columns(2)
                        col_isvpt_a.metric("ISVPT de la entidad (0 a 1)", resultado_isvpt.isvpt_entidad_referencia)
                        col_isvpt_b.metric(
                            "Posición en el grupo",
                            f"{resultado_isvpt.posicion_entidad_referencia} de {resultado_isvpt.n_entidades}",
                        )
                        if resultado_isvpt.subindices_por_dimension_entidad:
                            st.dataframe(
                                [{"Dimensión": k, "Subíndice normalizado": v} for k, v in resultado_isvpt.subindices_por_dimension_entidad.items()],
                                use_container_width=True,
                            )
                        st.session_state.ultimo_diagnostico_real["isvpt"] = resultado_isvpt
                    else:
                        st.caption("No se pudo ubicar la entidad dentro del grupo filtrado para calcular su ISVPT.")
                except Exception as e:
                    st.error(f"No se pudo calcular el análisis 360: {e}")

            texto_recos = ""
            if cruce:
                for lista in cruce.values():
                    texto_recos += "\n".join(lista) + "\n"
            # Se preservan las claves calculadas "a demanda" (ISVPT, análisis IA
            # y los documentos ya generados) al reconstruir este diccionario en
            # cada rerun. Si no se preservaran, cada interacción (incluido
            # hacer clic en cualquier botón de descarga) las borraría, y habría
            # que volver a generar el análisis con IA (20-40 seg) y los 6
            # documentos desde cero en cada descarga.
            _anterior = st.session_state.get("ultimo_diagnostico_real", {})
            st.session_state.ultimo_diagnostico_real = {
                "nombre": entidad.nombre,
                "diag": diag,
                "texto_recos": texto_recos,
                "cruce": cruce,
                "total_recos_entidad": total_recos_entidad,
                "idi_oficial": idi_oficial,
                "grupo_par": grupo_par,
                "tipo_regimen_especial": tipo_regimen_especial_elegido,
                "isvpt": _anterior.get("isvpt"),
                "analisis_ia": _anterior.get("analisis_ia"),
                "docs_generados": _anterior.get("docs_generados"),
            }

        if "ultimo_diagnostico_real" in st.session_state:
            st.markdown("---")
            st.markdown("### 🧠 Análisis integral con IA (a demanda, solo para esta entidad)")
            st.caption(
                "Genera, con Claude, la lectura desde las tres teorías del curso (NGP, post-NGP, "
                "Nuevo Institucionalismo), recomendaciones técnicas/jurídicas/financieras, identificación "
                "inicial de riesgo (sin probabilidad/impacto — eso corresponde al líder de proceso) y una "
                "prospectiva orientada a valor público. Se genera SOLO para la entidad ya diagnosticada arriba, "
                "una vez por clic — no se ejecuta en bloque para otras entidades."
            )
            _n_brechas_ia = len(st.session_state.ultimo_diagnostico_real["diag"].brechas)
            if _n_brechas_ia <= 10:
                _mensaje_espera_ia = f"Generando análisis con IA ({_n_brechas_ia} brechas — normalmente 1-2 minutos)..."
            elif _n_brechas_ia <= 20:
                _mensaje_espera_ia = (
                    f"Generando análisis con IA ({_n_brechas_ia} brechas — puede tardar 3-6 minutos, "
                    "ya que hay que desarrollar el ciclo completo de riesgo para cada brecha). "
                    "No cierre ni recargue la página."
                )
            else:
                _mensaje_espera_ia = (
                    f"Generando análisis con IA ({_n_brechas_ia} brechas — puede tardar 6-10 minutos, "
                    "ya que con tantas brechas el texto a generar es largo y la generación tiene un "
                    "ritmo fijo, sin importar cuántas llamadas se necesiten). No cierre ni recargue la página."
                )
            if st.button("Generar análisis integral con IA", key="btn_ia"):
                try:
                    from backend.motores.motor_analisis_ia import generar_analisis_integral
                    with st.spinner(_mensaje_espera_ia):
                        datos = st.session_state.ultimo_diagnostico_real
                        resultado = generar_analisis_integral(
                            datos["nombre"], datos["diag"], datos["texto_recos"]
                        )
                    st.markdown(resultado)
                    st.session_state.ultimo_diagnostico_real["analisis_ia"] = resultado
                except Exception as e:
                    st.error(
                        f"No se pudo generar el análisis: {e}\n\n"
                        "Verifique que ANTHROPIC_API_KEY esté configurada en st.secrets "
                        "(Streamlit Community Cloud → Settings → Secrets)."
                    )

            if st.session_state.ultimo_diagnostico_real.get("analisis_ia"):
                datos_informe = st.session_state.ultimo_diagnostico_real

                # Los 6 documentos (docx+pdf de los 3 informes) son costosos de
                # generar (tablas, gráficas). Antes se regeneraban en CADA
                # rerun de Streamlit —incluido cada clic en un botón de
                # descarga— porque st.download_button necesita los bytes ya
                # calculados para poder dibujarse. Aquí se cachean y solo se
                # regeneran cuando cambia algo real (nuevo análisis IA, nuevo
                # cálculo de Análisis 360/ISVPT); de lo contrario se reusan
                # los bytes ya generados, sin recalcular nada.
                _firma_actual = (
                    id(datos_informe.get("analisis_ia")),
                    id(st.session_state.get("ultimo_analisis_360")),
                    id(datos_informe.get("isvpt")),
                )
                _cache = datos_informe.get("docs_generados")
                if _cache and _cache.get("firma") == _firma_actual:
                    _docs = _cache["bytes"]
                else:
                    _docs = {}

                def _generar_si_falta(clave, funcion):
                    """Genera (y cachea) el documento solo si no está ya en _docs."""
                    if clave not in _docs:
                        _docs[clave] = funcion().getvalue()
                    return _docs[clave]

                st.markdown("---")
                st.markdown("### 📄 Informe técnico descargable")
                st.caption(
                    "Empaqueta el diagnóstico real, la contextualización académica fija "
                    "(desde los griegos hasta la Administración Pública contemporánea) y "
                    "el análisis integral con IA ya generado arriba, en un solo archivo "
                    "descargable para entregar como evidencia de ejecución del sistema."
                )
                col_docx, col_pdf = st.columns(2)
                try:
                    from backend.motores.generador_informe import generar_reporte_docx, generar_reporte_pdf

                    docx_bytes = _generar_si_falta("docx_tecnico", lambda: generar_reporte_docx(
                        datos_informe["nombre"], datos_informe["diag"], datos_informe["analisis_ia"],
                        resultado_isvpt=datos_informe.get("isvpt"),
                        resultado_360=st.session_state.get("ultimo_analisis_360"),
                        idi_oficial=datos_informe.get("idi_oficial"),
                        cruce_recomendaciones=datos_informe.get("cruce"),
                        total_recomendaciones_entidad=datos_informe.get("total_recos_entidad"),
                        tipo_regimen_especial=datos_informe.get("tipo_regimen_especial"),
                    ))
                    col_docx.download_button(
                        "⬇️ Descargar informe en Word (.docx)",
                        data=docx_bytes,
                        file_name=f"informe_{datos_informe['nombre'].replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="descarga_docx",
                    )

                    pdf_bytes = _generar_si_falta("pdf_tecnico", lambda: generar_reporte_pdf(
                        datos_informe["nombre"], datos_informe["diag"], datos_informe["analisis_ia"],
                        resultado_isvpt=datos_informe.get("isvpt"),
                        resultado_360=st.session_state.get("ultimo_analisis_360"),
                        idi_oficial=datos_informe.get("idi_oficial"),
                        cruce_recomendaciones=datos_informe.get("cruce"),
                        total_recomendaciones_entidad=datos_informe.get("total_recos_entidad"),
                        tipo_regimen_especial=datos_informe.get("tipo_regimen_especial"),
                    ))
                    col_pdf.download_button(
                        "⬇️ Descargar informe en PDF",
                        data=pdf_bytes,
                        file_name=f"informe_{datos_informe['nombre'].replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        key="descarga_pdf",
                    )
                except Exception as e:
                    st.error(f"No se pudo generar el informe descargable: {e}")

                st.markdown("---")
                st.markdown("### 🎓 Estudio de Caso Académico (Unidad 2 — 35% de la nota)")
                st.caption(
                    "Genera el 'Informe técnico de análisis del caso' que exige el "
                    "microcurrículo de Enfoques y Teorías de la Administración Pública II: "
                    "descripción del problema público, contexto local/regional/global, "
                    "análisis desde la NGP/post-NGP/Nuevo Institucionalismo, fuentes de "
                    "datos abiertos, fundamento jurídico-fiscal-disciplinario-contable-"
                    "control interno, fortalezas, debilidades y recomendaciones."
                )
                col_ec_docx, col_ec_pdf = st.columns(2)
                try:
                    docx_bytes_ec = _generar_si_falta("docx_estudio_caso", lambda: generar_estudio_de_caso_docx(
                        datos_informe["nombre"], datos_informe["diag"], datos_informe["analisis_ia"],
                        cruce_recomendaciones=datos_informe.get("cruce"),
                        resultado_360=st.session_state.get("ultimo_analisis_360"),
                        resultado_isvpt=datos_informe.get("isvpt"),
                        idi_oficial=datos_informe.get("idi_oficial"),
                        departamento=st.session_state.get("departamento_360_actual"),
                        tipo_regimen_especial=datos_informe.get("tipo_regimen_especial"),
                    ))
                    col_ec_docx.download_button(
                        "⬇️ Descargar Estudio de Caso (.docx)",
                        data=docx_bytes_ec,
                        file_name=f"estudio_de_caso_{datos_informe['nombre'].replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="descarga_ec_docx",
                    )

                    pdf_bytes_ec = _generar_si_falta("pdf_estudio_caso", lambda: generar_estudio_de_caso_pdf(
                        datos_informe["nombre"], datos_informe["diag"], datos_informe["analisis_ia"],
                        cruce_recomendaciones=datos_informe.get("cruce"),
                        resultado_360=st.session_state.get("ultimo_analisis_360"),
                        resultado_isvpt=datos_informe.get("isvpt"),
                        idi_oficial=datos_informe.get("idi_oficial"),
                        departamento=st.session_state.get("departamento_360_actual"),
                        tipo_regimen_especial=datos_informe.get("tipo_regimen_especial"),
                    ))
                    col_ec_pdf.download_button(
                        "⬇️ Descargar Estudio de Caso (PDF)",
                        data=pdf_bytes_ec,
                        file_name=f"estudio_de_caso_{datos_informe['nombre'].replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        key="descarga_ec_pdf",
                    )
                    if not st.session_state.get("ultimo_analisis_360"):
                        st.caption(
                            "💡 Sugerencia: calcule primero el 'Análisis 360' más arriba para "
                            "que el Estudio de Caso incluya el contexto regional (percentil, "
                            "grupo de comparación) con datos reales."
                        )
                except Exception as e:
                    st.error(f"No se pudo generar el Estudio de Caso: {e}")

                st.markdown("---")
                st.markdown("### 🏛️ Informe Ejecutivo para Representantes Legales")
                st.caption(
                    "Versión en lenguaje llano, sin jerga académica, pensada para la Alta "
                    "Dirección: semáforo de 5 quintiles del MIPG, brechas más críticas, "
                    "comparación con entidades similares, recomendaciones oficiales completas "
                    "de la Función Pública y sus implicaciones legales, fiscales, "
                    "administrativas y disciplinarias (incluye el artículo 6 de la "
                    "Constitución y una matriz de riesgo)."
                )
                col_alc_docx, col_alc_pdf = st.columns(2)
                try:
                    docx_bytes_alc = _generar_si_falta("docx_ejecutivo", lambda: generar_informe_alcaldes_docx(
                        datos_informe["nombre"], datos_informe["diag"],
                        resultado_isvpt=datos_informe.get("isvpt"),
                        resultado_360=st.session_state.get("ultimo_analisis_360"),
                        idi_oficial=datos_informe.get("idi_oficial"),
                        cruce_recomendaciones=datos_informe.get("cruce"),
                        total_recomendaciones_entidad=datos_informe.get("total_recos_entidad"),
                        tipo_regimen_especial=datos_informe.get("tipo_regimen_especial"),
                    ))
                    col_alc_docx.download_button(
                        "⬇️ Descargar Informe Ejecutivo (.docx)",
                        data=docx_bytes_alc,
                        file_name=f"informe_ejecutivo_{datos_informe['nombre'].replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="descarga_alc_docx",
                    )

                    pdf_bytes_alc = _generar_si_falta("pdf_ejecutivo", lambda: generar_informe_alcaldes_pdf(
                        datos_informe["nombre"], datos_informe["diag"],
                        resultado_isvpt=datos_informe.get("isvpt"),
                        resultado_360=st.session_state.get("ultimo_analisis_360"),
                        idi_oficial=datos_informe.get("idi_oficial"),
                        cruce_recomendaciones=datos_informe.get("cruce"),
                        total_recomendaciones_entidad=datos_informe.get("total_recos_entidad"),
                        tipo_regimen_especial=datos_informe.get("tipo_regimen_especial"),
                    ))
                    col_alc_pdf.download_button(
                        "⬇️ Descargar Informe Ejecutivo (PDF)",
                        data=pdf_bytes_alc,
                        file_name=f"informe_ejecutivo_{datos_informe['nombre'].replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        key="descarga_alc_pdf",
                    )
                except Exception as e:
                    st.error(f"No se pudo generar el Informe Ejecutivo: {e}")

                # Se guarda el caché de los 6 documentos junto con la "firma"
                # que describe con qué datos se generaron, para reusarlos sin
                # recalcular mientras nada relevante cambie.
                datos_informe["docs_generados"] = {"firma": _firma_actual, "bytes": _docs}
    else:
        st.info("Suba un archivo para habilitar el selector de entidades.")

# ----------------------------------------------------------------------
# MODO 2: captura manual (V1) — útil para simular escenarios o entidades
# que aún no tienen resultado oficial (ej. formulario en curso).
# ----------------------------------------------------------------------
with tab_manual:
    col1, col2, col3 = st.columns(3)
    nombre_entidad = col1.text_input("Nombre de la entidad", "Alcaldía de Ejemplo")
    codigo_dane = col2.text_input("Código DANE (opcional)", "")
    vigencia = col3.number_input("Vigencia", min_value=2018, max_value=2030, value=2025, step=1)

    st.caption(
        "Ingrese el puntaje FURAG (0-100) de cada índice que la entidad ya tiene "
        "reportado. Los índices sin puntaje simplemente no entran al promedio."
    )

    if "puntajes" not in st.session_state:
        st.session_state.puntajes = {}

    with st.expander("Capturar puntajes por dimensión", expanded=True):
        for cod_dim, dim in dimensiones().items():
            if not any(pol["indices"] for pol in dim["politicas"].values()):
                continue  # dimensión sin índices capturables manualmente (ej. D4/POL14)
            st.markdown(f"**{dim['nombre']}** ({cod_dim})")
            cols = st.columns(4)
            i = 0
            for cod_pol, pol in dim["politicas"].items():
                for cod_idx, idx in pol["indices"].items():
                    with cols[i % 4]:
                        valor = st.number_input(
                            f"{cod_idx}: {idx['nombre'][:35]}",
                            min_value=0.0, max_value=100.0, value=0.0, step=1.0,
                            key=f"puntaje_{cod_idx}", help=idx["descripcion"],
                        )
                        if valor > 0:
                            st.session_state.puntajes[cod_idx] = valor
                    i += 1

    if st.button("Generar diagnóstico", type="primary", key="btn_manual"):
        resultados = [
            ResultadoIndice(codigo_indice=cod, puntaje=val, vigencia=int(vigencia))
            for cod, val in st.session_state.puntajes.items()
        ]
        if not resultados:
            st.warning("Ingrese al menos un puntaje antes de generar el diagnóstico.")
        else:
            entidad = Entidad(
                nombre=nombre_entidad, codigo_dane=codigo_dane or None,
                vigencia=int(vigencia), resultados=resultados,
            )
            diag = diagnosticar(entidad)
            mostrar_diagnostico(diag)
