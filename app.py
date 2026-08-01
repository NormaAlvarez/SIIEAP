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

CARPETA_DATA = Path(__file__).resolve().parent / "data"
RUTA_AUTO_NACION = CARPETA_DATA / "resultados_nacion.xlsx"
RUTA_AUTO_TERRITORIO = CARPETA_DATA / "resultados_territorio.xlsx"
RUTA_AUTO_RECOMENDACIONES = CARPETA_DATA / "recomendaciones_consolidado.json"
RUTA_AUTO_RECOMENDACIONES_GZ = CARPETA_DATA / "recomendaciones_consolidado.json.gz"

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

st.title("SIIEAP — Sistema de Diagnóstico del Desempeño Institucional (IDI-MIPG)")
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

        if entidad_elegida and st.button("Generar diagnóstico", type="primary", key="btn_real"):
            entidad, idi_oficial, grupo_par = entidad_por_nombre_exacto(df, entidad_elegida)
            st.subheader(entidad.nombre)
            st.caption(
                f"Índices/políticas con información: {len(entidad.resultados) + len(entidad.resultados_politica_directa)} de 66"
            )
            diag = diagnosticar(entidad)

            cruce = None
            if "consolidado_cargado" in st.session_state:
                recos = recomendaciones_de_entidad(st.session_state.consolidado_cargado, entidad_elegida)
                if recos:
                    cruce = cruzar_brechas_con_recomendaciones(diag.brechas, recos)
                    st.caption(f"{len(recos)} recomendaciones encontradas en el consolidado para esta entidad.")
                else:
                    st.caption("Esta entidad no aparece en el consolidado de recomendaciones cargado.")
            elif archivo_recos is not None:
                try:
                    recos = cargar_recomendaciones(archivo_recos)
                    cruce = cruzar_brechas_con_recomendaciones(diag.brechas, recos)
                    st.caption(f"{len(recos)} recomendaciones cargadas, {len(cruce)} brechas con recomendación vinculada.")
                except Exception as e:
                    st.warning(f"No se pudieron leer las recomendaciones: {e}")

            mostrar_diagnostico(diag, idi_oficial, grupo_par, cruce)

            texto_recos = ""
            if cruce:
                for lista in cruce.values():
                    texto_recos += "\n".join(lista) + "\n"
            st.session_state.ultimo_diagnostico_real = {
                "nombre": entidad.nombre,
                "diag": diag,
                "texto_recos": texto_recos,
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
            if st.button("Generar análisis integral con IA", key="btn_ia"):
                try:
                    from backend.motores.motor_analisis_ia import generar_analisis_integral
                    with st.spinner("Generando análisis con IA (puede tardar 20-40 segundos)..."):
                        datos = st.session_state.ultimo_diagnostico_real
                        resultado = generar_analisis_integral(
                            datos["nombre"], datos["diag"], datos["texto_recos"]
                        )
                    st.markdown(resultado)
                except Exception as e:
                    st.error(
                        f"No se pudo generar el análisis: {e}\n\n"
                        "Verifique que ANTHROPIC_API_KEY esté configurada en st.secrets "
                        "(Streamlit Community Cloud → Settings → Secrets)."
                    )
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
