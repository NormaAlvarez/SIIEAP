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
    REGIMEN_ESPECIAL_RAMA_LEGISLATIVA,
    REGIMEN_ESPECIAL_ORGANO_CONTROL_NACIONAL,
    REGIMEN_ESPECIAL_RAMA_JUDICIAL,
    REGIMEN_ESPECIAL_ORGANIZACION_ELECTORAL,
)
from backend.base_conocimiento.subregiones_antioquia import todas_las_subregiones

CARPETA_DATA = Path(__file__).resolve().parent / "data"
RUTA_AUTO_NACION = CARPETA_DATA / "resultados_nacion.xlsx"
RUTA_AUTO_TERRITORIO = CARPETA_DATA / "resultados_territorio.xlsx"
RUTA_AUTO_RECOMENDACIONES = CARPETA_DATA / "recomendaciones_consolidado.json"
RUTA_AUTO_RECOMENDACIONES_GZ = CARPETA_DATA / "recomendaciones_consolidado.json.gz"
CARPETA_PARTES_RECOMENDACIONES = CARPETA_DATA / "partes_recomendaciones"


def _reconstruir_gz_desde_partes() -> None:
    """El archivo data/recomendaciones_consolidado.json.gz pesa ~52 MB, por
    encima del límite de subida por arrastrar-y-soltar del navegador de
    GitHub (~25 MB). Por eso se sube dividido en partes pequeñas dentro de
    data/partes_recomendaciones/ (parte_00, parte_01, ...). Esta función las
    reensambla en el .gz completo la primera vez que la app arranca en un
    entorno nuevo. Si el .gz ya existe (por ejemplo, en desarrollo local),
    no hace nada."""
    if RUTA_AUTO_RECOMENDACIONES_GZ.exists():
        return
    if not CARPETA_PARTES_RECOMENDACIONES.exists():
        return
    partes = sorted(CARPETA_PARTES_RECOMENDACIONES.glob("parte_*"))
    if not partes:
        return
    CARPETA_DATA.mkdir(parents=True, exist_ok=True)
    with open(RUTA_AUTO_RECOMENDACIONES_GZ, "wb") as salida:
        for parte in partes:
            salida.write(parte.read_bytes())


_reconstruir_gz_desde_partes()

OPCIONES_REGIMEN_ESPECIAL = {
    "Rama Ejecutiva — MIPG íntegro aplica (alcaldías, gobernaciones, EICE, ESE, ministerios...)": REGIMEN_ESPECIAL_NINGUNO,
    "Ente universitario autónomo (ej. Universidad de Antioquia)": REGIMEN_ESPECIAL_UNIVERSIDAD_AUTONOMA,
    "Contraloría territorial (departamental/municipal/distrital)": REGIMEN_ESPECIAL_ORGANO_CONTROL,
    "Personería municipal o distrital": REGIMEN_ESPECIAL_PERSONERIA,
    "Concejo municipal o Asamblea departamental": REGIMEN_ESPECIAL_CONCEJO_ASAMBLEA,
    "Banco de la República": REGIMEN_ESPECIAL_BANCO_REPUBLICA,
    "Corporación Autónoma Regional": REGIMEN_ESPECIAL_CORPORACION_AUTONOMA,
    "Rama Legislativa (Senado, Cámara de Representantes)": REGIMEN_ESPECIAL_RAMA_LEGISLATIVA,
    "Órgano de control nacional (Procuraduría, Contraloría General, Defensoría, Auditoría General)": REGIMEN_ESPECIAL_ORGANO_CONTROL_NACIONAL,
    "Rama Judicial (Fiscalía, Consejo Superior de la Judicatura, Medicina Legal)": REGIMEN_ESPECIAL_RAMA_JUDICIAL,
    "Organización Electoral (Registraduría, Consejo Nacional Electoral)": REGIMEN_ESPECIAL_ORGANIZACION_ELECTORAL,
}


def _sugerir_regimen_especial(fila_entidad) -> str | None:
    """Sugiere (adivina) el régimen especial correcto a partir de columnas
    reales del Excel oficial de Función Pública — "Tipo Formulario",
    "Naturaleza Jurídica" y "Clasificación orgánica" —, para preseleccionar
    el valor correcto en el selector y que no dependa de que alguien se
    acuerde de elegirlo manualmente (ese olvido es justo lo que hacía que
    entidades MECI-only, como el Senado, se evaluaran por error con el MIPG
    íntegro). Nunca es definitivo ni oculta el selector: el usuario siempre
    puede cambiarlo. Devuelve None si "Tipo Formulario" no es "MECI" (no hace
    falta régimen especial) o si no hay suficiente certeza sobre cuál
    categoría específica aplica — en ese caso se deja el valor por defecto y
    es responsabilidad del usuario revisarlo y elegirlo a mano.
    """
    if fila_entidad is None or not len(fila_entidad):
        return None
    fila = fila_entidad.iloc[0]
    tipo_formulario = str(fila.get("Tipo Formulario", "")).strip().upper()
    if tipo_formulario != "MECI":
        return None  # Reporta MIPG completo: no hace falta régimen especial.

    naturaleza = str(fila.get("Naturaleza Jurídica", "")).strip().upper()
    clasificacion = str(fila.get("Clasificación orgánica", "")).strip().upper()
    nombre_entidad = str(fila.get("Entidad", "")).strip().upper()

    if naturaleza == "ENTE UNIVERSITARIO AUTÓNOMO":
        return REGIMEN_ESPECIAL_UNIVERSIDAD_AUTONOMA
    if "CONTRALORÍA" in naturaleza or "CONTRALORIA" in naturaleza:
        return REGIMEN_ESPECIAL_ORGANO_CONTROL
    if "PERSONERÍA" in naturaleza or "PERSONERIA" in naturaleza:
        return REGIMEN_ESPECIAL_PERSONERIA
    if naturaleza in ("CONCEJO MUNICIPAL", "ASAMBLEA DEPARTAMENTAL"):
        return REGIMEN_ESPECIAL_CONCEJO_ASAMBLEA
    if "BANCO DE LA REPÚBLICA" in naturaleza or "BANCO DE LA REPUBLICA" in nombre_entidad:
        return REGIMEN_ESPECIAL_BANCO_REPUBLICA
    if "CORPORACION AUTONOMA" in nombre_entidad or "CORPORACIÓN AUTÓNOMA" in nombre_entidad:
        return REGIMEN_ESPECIAL_CORPORACION_AUTONOMA
    if clasificacion == "RAMA LEGISLATIVA":
        return REGIMEN_ESPECIAL_RAMA_LEGISLATIVA
    if clasificacion == "RAMA JUDICIAL":
        return REGIMEN_ESPECIAL_RAMA_JUDICIAL
    if clasificacion == "ORGANIZACIÓN ELECTORAL":
        return REGIMEN_ESPECIAL_ORGANIZACION_ELECTORAL
    if clasificacion == "ORGANOS DE CONTROL":
        return REGIMEN_ESPECIAL_ORGANO_CONTROL_NACIONAL
    return None  # No hay suficiente certeza: se deja para elección manual.

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

tab_real, tab_manual, tab_lotes = st.tabs([
    "📊 Cargar entidad real (Excel oficial)", "✍️ Captura manual", "🗂️ Modo por lotes",
])

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

        # Se sugiere (preselecciona) el régimen especial correcto a partir de
        # columnas reales del Excel oficial ("Tipo Formulario", "Naturaleza
        # Jurídica", "Clasificación orgánica") — así el usuario no depende de
        # acordarse de elegirlo manualmente cada vez (ese olvido es lo que
        # causaba que entidades MECI-only, como el Senado, se evaluaran por
        # error con el MIPG íntegro). Sigue siendo 100% editable: es solo el
        # valor con el que arranca el selector para esta entidad puntual.
        _regimen_sugerido = None
        if entidad_elegida:
            _fila_regimen = df[df["Entidad"].astype(str).str.strip() == entidad_elegida.strip()]
            _regimen_sugerido = _sugerir_regimen_especial(_fila_regimen)

        _etiquetas_regimen = list(OPCIONES_REGIMEN_ESPECIAL.keys())
        _indice_por_defecto = 0
        if _regimen_sugerido:
            for _i, _etq in enumerate(_etiquetas_regimen):
                if OPCIONES_REGIMEN_ESPECIAL[_etq] == _regimen_sugerido:
                    _indice_por_defecto = _i
                    break

        etiqueta_regimen_elegida = st.selectbox(
            "Tipo de entidad (régimen especial MIPG/MECI)",
            _etiquetas_regimen,
            index=_indice_por_defecto,
            # La clave incluye la entidad para que, al cambiar de entidad, el
            # selector vuelva a partir de la sugerencia recién calculada para
            # ESA entidad, en vez de arrastrar la última selección manual de
            # una entidad distinta.
            key=f"tipo_regimen_especial_select_{entidad_elegida}",
            help=(
                "Si la entidad NO pertenece a la Rama Ejecutiva (universidades autónomas, "
                "órganos de control, Rama Legislativa/Judicial, Organización Electoral, "
                "Banco de la República, Corporaciones Autónomas Regionales), el MIPG no le "
                "aplica en su integralidad — solo la política de Control Interno (MECI). "
                "El sistema intenta preseleccionar la categoría correcta a partir del "
                "archivo oficial, pero SIEMPRE verifique que sea la adecuada antes de "
                "generar el diagnóstico."
            ),
        )
        if _regimen_sugerido:
            st.caption(
                f"ℹ️ Se preseleccionó automáticamente '{_etiquetas_regimen[_indice_por_defecto]}' "
                "a partir del 'Tipo Formulario' y la 'Naturaleza Jurídica' de esta entidad en "
                "el archivo oficial. Verifique que sea correcto antes de continuar."
            )
        tipo_regimen_especial_elegido = OPCIONES_REGIMEN_ESPECIAL[etiqueta_regimen_elegida]

        if entidad_elegida and st.button("Generar diagnóstico", type="primary", key="btn_real"):
            entidad, idi_oficial, grupo_par, dimensiones_oficiales = entidad_por_nombre_exacto(df, entidad_elegida)
            entidad.regimen_especial = tipo_regimen_especial_elegido
            diag = diagnosticar(entidad, dimensiones_oficiales=dimensiones_oficiales)

            cruce = None
            total_recos_entidad = None
            recos = None
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
            # Cálculo automático por defecto del Análisis 360 y el ISVPT,
            # apenas se genera el diagnóstico — para CUALQUIER entidad, de
            # cualquier municipio o departamento del país, sin depender de
            # que alguien haga clic aparte en "Calcular análisis 360" (si no
            # se hacía ese clic, el informe salía sin sección de ISVPT).
            # Se usa como comparación por defecto el mismo Departamento y el
            # mismo "Grupo par" oficial de Función Pública de la propia
            # entidad — un campo que Función Pública ya asigna a TODAS las
            # entidades del país (alcaldías, gobernaciones, etc.), por lo que
            # este filtro por defecto funciona sin importar el departamento,
            # tenga o no subregiones registradas en el sistema. Si el
            # usuario quiere un grupo de comparación distinto (otra
            # subregión, sin filtro de grupo par, etc.), puede seguir
            # ajustándolo y recalculando manualmente más abajo — eso
            # simplemente sobrescribe este valor por defecto.
            resultado_360_auto = None
            resultado_isvpt_auto = None
            try:
                _fila_geo = df[df["Entidad"].astype(str).str.strip() == entidad_elegida.strip()]
                _departamento_entidad = (
                    str(_fila_geo.iloc[0].get("Departamento", "")).strip() if len(_fila_geo) else ""
                ) or None
                # Para entidades del nivel NACIONAL (Senado, DAPRE, Ministerios,
                # Procuraduría, etc.), Función Pública ya asigna un "Grupo par"
                # que empieza con "NACIÓN" (p. ej. "NACIÓN - MIPG" o "NACIÓN -
                # MECI") — esa clasificación YA ES el grupo de comparación
                # completo y correcto a nivel país. Si además se filtrara por
                # Departamento, se excluirían por error pares legítimos cuya
                # sede simplemente no está en Bogotá (p. ej. universidades
                # públicas regionales o Corporaciones Autónomas Regionales,
                # que también reportan solo MECI) — reduciendo el grupo de
                # comparación de forma arbitraria y ajena al criterio real de
                # Función Pública. Por eso, para estas entidades, el filtro
                # automático usa SOLO el Grupo par, sin Departamento.
                _es_entidad_nacional = bool(grupo_par) and grupo_par.strip().upper().startswith("NACIÓN")
                if _es_entidad_nacional:
                    resultado_360_auto = analizar_360(
                        df, grupo_par_contiene=grupo_par, entidad_referencia=entidad_elegida,
                    )
                    df_grupo_auto, _ = filtrar_grupo(df, grupo_par_contiene=grupo_par)
                    resultado_isvpt_auto = calcular_isvpt(df_grupo_auto, entidad_referencia=entidad_elegida)
                elif _departamento_entidad and grupo_par:
                    resultado_360_auto = analizar_360(
                        df, departamento=_departamento_entidad, grupo_par_contiene=grupo_par,
                        entidad_referencia=entidad_elegida,
                    )
                    df_grupo_auto, _ = filtrar_grupo(
                        df, departamento=_departamento_entidad, grupo_par_contiene=grupo_par,
                    )
                    resultado_isvpt_auto = calcular_isvpt(df_grupo_auto, entidad_referencia=entidad_elegida)
            except Exception:
                # El cálculo automático es un "mejor esfuerzo": si falla por
                # cualquier razón (columna faltante, etc.), simplemente no se
                # rellena por defecto y el usuario puede calcularlo manualmente
                # más abajo — nunca debe romper la generación del diagnóstico.
                pass

            st.session_state["resultado_real_actual"] = {
                "entidad": entidad,
                "idi_oficial": idi_oficial,
                "dimensiones_oficiales": dimensiones_oficiales,
                "grupo_par": grupo_par,
                "diag": diag,
                "cruce": cruce,
                "recomendaciones_completas": recos,
                "total_recos_entidad": total_recos_entidad,
                "etiqueta_regimen_elegida": etiqueta_regimen_elegida,
                "tipo_regimen_especial_elegido": tipo_regimen_especial_elegido,
                "resultado_360_auto": resultado_360_auto,
                "resultado_isvpt_auto": resultado_isvpt_auto,
            }
            if resultado_360_auto is not None:
                st.session_state.ultimo_analisis_360 = resultado_360_auto

        if "resultado_real_actual" in st.session_state:
            _r = st.session_state["resultado_real_actual"]
            entidad = _r["entidad"]
            idi_oficial = _r["idi_oficial"]
            dimensiones_oficiales = _r.get("dimensiones_oficiales", {})
            grupo_par = _r["grupo_par"]
            diag = _r["diag"]
            cruce = _r["cruce"]
            total_recos_entidad = _r["total_recos_entidad"]
            recomendaciones_completas = _r.get("recomendaciones_completas")
            etiqueta_regimen_elegida = _r["etiqueta_regimen_elegida"]
            tipo_regimen_especial_elegido = _r["tipo_regimen_especial_elegido"]
            resultado_360_auto = _r.get("resultado_360_auto")
            resultado_isvpt_auto = _r.get("resultado_isvpt_auto")

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
            if resultado_isvpt_auto is not None and resultado_isvpt_auto.isvpt_entidad_referencia is not None:
                st.caption(
                    f"✅ Ya se calculó automáticamente un grupo de comparación por defecto "
                    f"(Departamento + Grupo par oficial '{grupo_par}', {resultado_isvpt_auto.n_entidades} "
                    "entidades) y su ISVPT — visible más abajo, junto al diagnóstico, y ya incluido en "
                    "los informes descargables. Si prefiere otro grupo de comparación (p. ej. una "
                    "subregión específica), ajústelo aquí abajo y haga clic en 'Calcular análisis 360' "
                    "para reemplazarlo."
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
            # cada rerun — pero SOLO si siguen correspondiendo a la MISMA
            # entidad. Si no se preservaran nunca, cada interacción (incluido
            # hacer clic en cualquier botón de descarga) las borraría, y habría
            # que volver a generar el análisis con IA y los 6 documentos desde
            # cero en cada descarga. Pero si se preservaran SIEMPRE sin
            # verificar la entidad, al diagnosticar una entidad nueva se
            # arrastrarían por error el ISVPT, el análisis IA y los documentos
            # de la entidad ANTERIOR (bug real detectado: dos entidades
            # distintas mostrando el mismo ISVPT y el mismo grupo de
            # comparación, copiados de una a otra). Por eso: solo se preserva
            # si "nombre" coincide con el de la entidad que se acaba de
            # diagnosticar; si es una entidad distinta, se parte de cero.
            _anterior = st.session_state.get("ultimo_diagnostico_real", {})
            _es_misma_entidad = _anterior.get("nombre") == entidad.nombre
            # Si es una entidad nueva, se parte del ISVPT calculado
            # automáticamente por defecto (arriba, con Departamento + Grupo
            # par de la propia entidad) — así el informe SIEMPRE trae una
            # sección de ISVPT, sin depender de un clic manual aparte. Si el
            # usuario ya había calculado uno manualmente para esta MISMA
            # entidad (con un filtro más específico, como subregión), ese se
            # respeta y no se pisa con el automático.
            _isvpt_por_defecto = _anterior.get("isvpt") if _es_misma_entidad else resultado_isvpt_auto
            st.session_state.ultimo_diagnostico_real = {
                "nombre": entidad.nombre,
                "diag": diag,
                "texto_recos": texto_recos,
                "cruce": cruce,
                "recomendaciones_completas": recos,
                "total_recos_entidad": total_recos_entidad,
                "idi_oficial": idi_oficial,
                "dimensiones_oficiales": dimensiones_oficiales,
                "grupo_par": grupo_par,
                "tipo_regimen_especial": tipo_regimen_especial_elegido,
                "isvpt": _isvpt_por_defecto,
                "analisis_ia": _anterior.get("analisis_ia") if _es_misma_entidad else None,
                "docs_generados": _anterior.get("docs_generados") if _es_misma_entidad else None,
            }
            if not _es_misma_entidad and resultado_360_auto is None:
                # Es una entidad distinta a la última diagnosticada en esta
                # sesión, y no se pudo calcular un Análisis 360 automático
                # por defecto (p. ej. faltó Departamento o Grupo par): el
                # Análisis 360 (y su ISVPT) calculado antes pertenecía a la
                # entidad anterior y ya no aplica aquí.
                st.session_state.pop("ultimo_analisis_360", None)

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
                    st.caption(_mensaje_espera_ia)
                    # Se muestra el texto EN VIVO, a medida que va llegando por
                    # streaming, en vez de un spinner ciego: además de dejar
                    # ver que sigue avanzando, esto mantiene la conexión con
                    # el navegador activa durante los varios minutos que
                    # puede tardar una entidad con muchas brechas — sin esto,
                    # el navegador no recibe NADA hasta el final y la conexión
                    # puede cortarse por inactividad antes de terminar.
                    marcador_texto_en_vivo = st.empty()
                    _ultimo_largo_mostrado = [0]

                    def _mostrar_avance(texto_parcial):
                        if len(texto_parcial) - _ultimo_largo_mostrado[0] >= 200:
                            marcador_texto_en_vivo.markdown(texto_parcial + " ▌")
                            _ultimo_largo_mostrado[0] = len(texto_parcial)

                    datos = st.session_state.ultimo_diagnostico_real
                    resultado = generar_analisis_integral(
                        datos["nombre"], datos["diag"], datos["texto_recos"],
                        on_texto_parcial=_mostrar_avance,
                        idi_oficial=datos.get("idi_oficial"),
                    )
                    marcador_texto_en_vivo.markdown(resultado)
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
                        recomendaciones_completas=datos_informe.get("recomendaciones_completas"),
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
                        recomendaciones_completas=datos_informe.get("recomendaciones_completas"),
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
                        recomendaciones_completas=datos_informe.get("recomendaciones_completas"),
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


# ----------------------------------------------------------------------
# MODO 3: generación por lotes — varias entidades, sin intervención manual
# entidad por entidad. Usa la MISMA API key ya configurada en los Secrets
# de esta app (nunca sale de aquí). Diseñado para poder detenerse y
# reanudar: cada entidad ya generada se guarda en session_state y se
# omite si se vuelve a correr, y se procesa un número acotado de
# entidades por clic para no dejar la conexión bloqueada por horas.
# ----------------------------------------------------------------------
with tab_lotes:
    import io
    import re
    import time
    import zipfile

    st.markdown(
        "Genera automáticamente los 3 informes (6 archivos: docx + PDF de cada uno) "
        "para una lista de entidades, sin repetir el proceso manual una por una."
    )
    st.warning(
        "⏱️ **Tiempos reales, sin exagerar:** cada entidad necesita su propio análisis "
        "con IA (1 a 10+ minutos, según su número de brechas) — no hay forma de "
        "acelerar eso, es el mismo límite de velocidad de generación de texto que ya "
        "vimos antes. Con 102 entidades, el total puede sumar varias HORAS. Por eso "
        "este modo procesa **una tanda acotada por clic** (usted decide cuántas) y "
        "recuerda lo que ya generó, para que pueda ir y volver en varias sesiones sin "
        "perder el avance ni repetir entidades ya hechas."
    )

    fuentes_lotes = []
    if RUTA_AUTO_NACION.exists():
        fuentes_lotes.append(("Nacion", RUTA_AUTO_NACION))
    if RUTA_AUTO_TERRITORIO.exists():
        fuentes_lotes.append(("Territorio", RUTA_AUTO_TERRITORIO))

    if "df_combinado_lotes" not in st.session_state and fuentes_lotes:
        with st.spinner("Cargando Nación + Territorio para el modo por lotes..."):
            import pandas as pd
            partes = []
            for hoja, ruta in fuentes_lotes:
                try:
                    partes.append(cargar_hoja(ruta, hoja))
                except Exception as e:
                    st.warning(f"No se pudo cargar {ruta.name} ({hoja}): {e}")
            if partes:
                st.session_state.df_combinado_lotes = pd.concat(partes, ignore_index=True)

    df_lotes = st.session_state.get("df_combinado_lotes")
    if df_lotes is None:
        st.error(
            "No se encontraron data/resultados_nacion.xlsx ni data/resultados_territorio.xlsx. "
            "El modo por lotes busca cada entidad en ambos archivos combinados, para no "
            "tener que saber de antemano si es una entidad territorial o nacional."
        )
    else:
        st.caption(f"Buscando entidades entre {len(df_lotes)} registros combinados (Nación + Territorio).")

        st.markdown("#### 1. Lista de entidades")
        archivo_lista = st.file_uploader(
            "Suba un .docx con una tabla de entidades (como el que ya me compartió), o pegue "
            "la lista abajo — una entidad por línea.",
            type=["docx"], key="archivo_lista_lotes",
        )
        texto_lista_default = ""
        if archivo_lista is not None:
            try:
                from docx import Document as DocxDocument
                doc_lista = DocxDocument(archivo_lista)
                nombres_extraidos = []
                for t in doc_lista.tables:
                    for fila in t.rows[1:]:  # se salta el encabezado
                        celdas = [c.text.strip() for c in fila.cells]
                        if len(celdas) >= 2 and celdas[1]:
                            nombres_extraidos.append(celdas[1])
                        elif len(celdas) == 1 and celdas[0]:
                            nombres_extraidos.append(celdas[0])
                if nombres_extraidos:
                    texto_lista_default = "\n".join(nombres_extraidos)
                    st.success(f"Se extrajeron {len(nombres_extraidos)} entidades del .docx.")
                else:
                    st.warning("No se encontró una tabla reconocible en el .docx — pegue la lista a mano abajo.")
            except Exception as e:
                st.warning(f"No se pudo leer el .docx ({e}) — pegue la lista a mano abajo.")

        texto_lista = st.text_area(
            "Lista de entidades (una por línea)", value=texto_lista_default, height=200,
            key="texto_lista_lotes",
        )
        nombres_lote = [n.strip() for n in texto_lista.split("\n") if n.strip()]

        if nombres_lote:
            st.caption(f"{len(nombres_lote)} entidades en la lista.")

            st.markdown("#### 2. Confirmar cada entidad contra el archivo oficial")
            st.caption(
                "Los nombres de la lista rara vez coinciden letra por letra con el archivo "
                "oficial (ej. 'Municipio de San Jerónimo' vs. 'ALCALDIA DE SAN JERONIMO'). "
                "Para las coincidencias EXACTAS se confirma solo; para todo lo demás, elija "
                "usted la correcta de una lista de candidatas — nunca se adivina en silencio, "
                "porque un emparejamiento equivocado generaría el informe de OTRA entidad."
            )
            nombres_oficiales_lotes = sorted(set(df_lotes["Entidad"].astype(str).str.strip()))

            PALABRAS_VACIAS_EMPAREJE = {
                "DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "SAS", "SA", "ESP", "ESE", "EICE", "SD",
                "EMPRESA", "EMPRESAS", "SERVICIOS", "PUBLICOS", "PUBLICA", "PUBLICAS", "SOCIAL", "ESTADO",
                "MUNICIPIO", "MUNICIPAL", "ALCALDIA", "HOSPITAL", "INSTITUTO", "INSTITUCION", "UNIVERSITARIA",
                "CONTRALORIA", "PERSONERIA", "CONCEJO", "ASAMBLEA", "DEPARTAMENTAL", "DEPARTAMENTO", "REGIONAL",
                "LOCAL", "CAMU", "CENTRO", "SALUD", "SOCIEDAD", "ADMINISTRATIVO", "CORPORACION", "UNIDAD",
            }

            def _normalizar_empareje(t):
                t = t.strip().upper()
                return "".join(c for c in __import__("unicodedata").normalize("NFKD", t) if not __import__("unicodedata").combining(c))

            def _tokenizar_empareje(t):
                t = _normalizar_empareje(t)
                t = __import__("re").sub(r"[^A-Z0-9\s]", " ", t)
                return set(p for p in t.split() if p and p not in PALABRAS_VACIAS_EMPAREJE)

            def _es_coincidencia_exacta(pedido, oficial):
                return _normalizar_empareje(pedido) == _normalizar_empareje(oficial)

            def _top_candidatas(nombre_pedido, top_n=8):
                tokens_pedido = _tokenizar_empareje(nombre_pedido)
                puntajes = []
                for oficial in nombres_oficiales_lotes:
                    tokens_of = _tokenizar_empareje(oficial)
                    if not tokens_pedido or not tokens_of:
                        continue
                    interseccion = tokens_pedido & tokens_of
                    union = tokens_pedido | tokens_of
                    score = len(interseccion) / len(union) if union else 0
                    if score > 0:
                        puntajes.append((score, oficial))
                puntajes.sort(key=lambda x: (-x[0], len(x[1])))
                return puntajes[:top_n]

            if st.session_state.get("_nombres_lote_previo") != nombres_lote:
                st.session_state["_nombres_lote_previo"] = nombres_lote
                st.session_state.confirmaciones_lotes = {}
                for n in nombres_lote:
                    exacta = next((o for o in nombres_oficiales_lotes if _es_coincidencia_exacta(n, o)), None)
                    st.session_state.confirmaciones_lotes[n] = exacta  # None si no hay coincidencia exacta

            exactas = {n: m for n, m in st.session_state.confirmaciones_lotes.items() if m}
            pendientes_confirmar = [n for n, m in st.session_state.confirmaciones_lotes.items() if not m]

            st.success(f"✅ {len(exactas)} coincidencia(s) exacta(s), confirmadas automáticamente.")

            if pendientes_confirmar:
                st.warning(f"⚠️ {len(pendientes_confirmar)} entidad(es) necesitan que usted elija la correcta:")
                for n in pendientes_confirmar:
                    candidatas = _top_candidatas(n)
                    opciones = ["— Sin coincidencia clara, omitir esta entidad —"] + [c[1] for c in candidatas]
                    etiquetas_score = {c[1]: c[0] for c in candidatas}
                    seleccion = st.selectbox(
                        f"'{n}' →",
                        opciones,
                        key=f"select_empareje_{n}",
                        format_func=lambda o: o if o not in etiquetas_score else f"{o}  (similitud {etiquetas_score[o]:.0%})",
                    )
                    if seleccion != opciones[0]:
                        st.session_state.confirmaciones_lotes[n] = seleccion

            with st.expander(f"Ver los {len(nombres_lote)} emparejamientos actuales"):
                for n, m in st.session_state.confirmaciones_lotes.items():
                    st.caption(f"✅ '{n}' → **{m}**" if m else f"⏳ '{n}' → aún sin confirmar")

            entidades_validas = [m for m in st.session_state.confirmaciones_lotes.values() if m]

            st.markdown("#### 3. Generar")
            if "lote_completadas" not in st.session_state:
                st.session_state.lote_completadas = {}  # nombre -> {clave_archivo: bytes}
            if "lote_fallidas" not in st.session_state:
                st.session_state.lote_fallidas = {}  # nombre -> motivo
            if "lote_objetivo" not in st.session_state:
                st.session_state.lote_objetivo = 0
            if "lote_procesadas_en_corrida" not in st.session_state:
                st.session_state.lote_procesadas_en_corrida = 0

            ya_hechas = [e for e in entidades_validas if e in st.session_state.lote_completadas]
            pendientes = [
                e for e in entidades_validas
                if e not in st.session_state.lote_completadas and e not in st.session_state.lote_fallidas
            ]
            st.info(
                f"✅ {len(ya_hechas)} ya generadas en esta sesión · "
                f"⏳ {len(pendientes)} pendientes · "
                f"❌ {len(st.session_state.lote_fallidas)} fallidas"
            )

            # ¿Hay una tanda en curso? Se considera "en curso" mientras queden
            # entidades pendientes por procesar EN ESTA corrida y falten por
            # completar el objetivo pedido.
            corrida_activa = (
                st.session_state.lote_objetivo > 0
                and st.session_state.lote_procesadas_en_corrida < st.session_state.lote_objetivo
                and pendientes
            )

            if not corrida_activa:
                cuantas_procesar = st.number_input(
                    "¿Cuántas entidades procesar en esta corrida?",
                    min_value=1, max_value=max(len(pendientes), 1),
                    value=min(3, max(len(pendientes), 1)), step=1,
                )
                if st.button("▶️ Procesar esta tanda", type="primary", disabled=not pendientes):
                    st.session_state.lote_objetivo = int(cuantas_procesar)
                    st.session_state.lote_procesadas_en_corrida = 0
                    st.rerun()
            else:
                # IMPORTANTE: aquí se procesa UNA SOLA entidad por cada
                # ejecución del script, y al terminar se llama st.rerun()
                # para encadenar con la siguiente — en vez de recorrer TODA
                # la tanda dentro de un único bucle bloqueante. Procesar
                # varias entidades seguidas sin ningún "respiro" entre ellas
                # tiene el mismo riesgo de desconexión por inactividad que
                # ya se resolvió para una sola entidad, pero multiplicado:
                # si la conexión se cae procesando la entidad 2 de 5, TODO
                # el resto de la tanda se pierde en silencio (sin error
                # visible), aunque la entidad 1 sí haya quedado guardada.
                # Con un rerun por entidad, cada ejecución del script está
                # acotada al tiempo de UNA sola entidad, y el navegador
                # tiene la oportunidad de reconectarse limpiamente entre
                # una y otra.
                nombre_of = pendientes[0]
                col_prog, col_stop = st.columns([4, 1])
                col_prog.info(
                    f"Procesando {st.session_state.lote_procesadas_en_corrida + 1} de "
                    f"{st.session_state.lote_objetivo} de esta corrida: **{nombre_of}**"
                )
                if col_stop.button("⏸️ Detener"):
                    st.session_state.lote_objetivo = 0
                    st.session_state.lote_procesadas_en_corrida = 0
                    st.rerun()
                else:
                    from backend.motores.motor_analisis_ia import generar_analisis_integral
                    from backend.motores.generador_informe import generar_reporte_docx, generar_reporte_pdf

                    marcador_avance_ia = st.empty()
                    try:
                        t0 = time.time()
                        fila = df_lotes[df_lotes["Entidad"].astype(str).str.strip() == nombre_of]
                        entidad, idi_oficial, grupo_par, dimensiones_oficiales = entidad_por_nombre_exacto(df_lotes, nombre_of)
                        regimen_sug = _sugerir_regimen_especial(fila) or REGIMEN_ESPECIAL_NINGUNO
                        entidad.regimen_especial = regimen_sug
                        diag = diagnosticar(entidad, dimensiones_oficiales=dimensiones_oficiales)

                        cruce, total_recos = None, None
                        recos = None
                        if "consolidado_cargado" in st.session_state:
                            recos = recomendaciones_de_entidad(st.session_state.consolidado_cargado, nombre_of)
                            if recos:
                                cruce = cruzar_brechas_con_recomendaciones(diag.brechas, recos)
                                total_recos = len(recos)

                        # ISVPT/Análisis 360 automático (Departamento+Grupo par,
                        # o solo Grupo par si es una entidad de nivel nacional).
                        resultado_360, resultado_isvpt = None, None
                        depto_ent = None
                        try:
                            depto_ent = str(fila.iloc[0].get("Departamento", "")).strip() or None
                            es_nacional = bool(grupo_par) and grupo_par.strip().upper().startswith("NACIÓN")
                            if es_nacional:
                                resultado_360 = analizar_360(df_lotes, grupo_par_contiene=grupo_par, entidad_referencia=nombre_of)
                                df_grupo, _ = filtrar_grupo(df_lotes, grupo_par_contiene=grupo_par)
                            elif depto_ent and grupo_par:
                                resultado_360 = analizar_360(df_lotes, departamento=depto_ent, grupo_par_contiene=grupo_par, entidad_referencia=nombre_of)
                                df_grupo, _ = filtrar_grupo(df_lotes, departamento=depto_ent, grupo_par_contiene=grupo_par)
                            else:
                                df_grupo = None
                            if df_grupo is not None:
                                resultado_isvpt = calcular_isvpt(df_grupo, entidad_referencia=nombre_of)
                        except Exception:
                            pass  # mejor esfuerzo, como en el modo individual

                        texto_recos = ""
                        if cruce:
                            for lista in cruce.values():
                                texto_recos += "\n".join(lista) + "\n"

                        def _avance(texto_parcial, _nombre=nombre_of):
                            marcador_avance_ia.caption(
                                f"IA escribiendo para {_nombre}... ({len(texto_parcial)} caracteres, "
                                f"{int(time.time() - t0)}s transcurridos)"
                            )

                        analisis_ia = generar_analisis_integral(
                            nombre_of, diag, texto_recos, on_texto_parcial=_avance,
                            idi_oficial=idi_oficial,
                        )

                        archivos_entidad = {}
                        docx_t = generar_reporte_docx(
                            nombre_of, diag, analisis_ia, resultado_isvpt=resultado_isvpt,
                            resultado_360=resultado_360, idi_oficial=idi_oficial,
                            cruce_recomendaciones=cruce, total_recomendaciones_entidad=total_recos,
                            tipo_regimen_especial=regimen_sug,
                            recomendaciones_completas=recos,
                        )
                        archivos_entidad["informe_tecnico.docx"] = docx_t.getvalue()
                        pdf_t = generar_reporte_pdf(
                            nombre_of, diag, analisis_ia, resultado_isvpt=resultado_isvpt,
                            resultado_360=resultado_360, idi_oficial=idi_oficial,
                            cruce_recomendaciones=cruce, total_recomendaciones_entidad=total_recos,
                            tipo_regimen_especial=regimen_sug,
                        )
                        archivos_entidad["informe_tecnico.pdf"] = pdf_t.getvalue()

                        docx_ec = generar_estudio_de_caso_docx(
                            nombre_of, diag, analisis_ia, cruce_recomendaciones=cruce,
                            resultado_360=resultado_360, resultado_isvpt=resultado_isvpt,
                            idi_oficial=idi_oficial, departamento=depto_ent,
                            tipo_regimen_especial=regimen_sug,
                            recomendaciones_completas=recos,
                        )
                        archivos_entidad["estudio_de_caso.docx"] = docx_ec.getvalue()
                        pdf_ec = generar_estudio_de_caso_pdf(
                            nombre_of, diag, analisis_ia, cruce_recomendaciones=cruce,
                            resultado_360=resultado_360, resultado_isvpt=resultado_isvpt,
                            idi_oficial=idi_oficial, departamento=depto_ent,
                            tipo_regimen_especial=regimen_sug,
                        )
                        archivos_entidad["estudio_de_caso.pdf"] = pdf_ec.getvalue()

                        docx_ej = generar_informe_alcaldes_docx(
                            nombre_of, diag, resultado_isvpt=resultado_isvpt,
                            resultado_360=resultado_360, idi_oficial=idi_oficial,
                            cruce_recomendaciones=cruce, total_recomendaciones_entidad=total_recos,
                            tipo_regimen_especial=regimen_sug,
                            recomendaciones_completas=recos,
                        )
                        archivos_entidad["informe_ejecutivo.docx"] = docx_ej.getvalue()
                        pdf_ej = generar_informe_alcaldes_pdf(
                            nombre_of, diag, resultado_isvpt=resultado_isvpt,
                            resultado_360=resultado_360, idi_oficial=idi_oficial,
                            cruce_recomendaciones=cruce, total_recomendaciones_entidad=total_recos,
                            tipo_regimen_especial=regimen_sug,
                        )
                        archivos_entidad["informe_ejecutivo.pdf"] = pdf_ej.getvalue()

                        st.session_state.lote_completadas[nombre_of] = archivos_entidad
                        st.session_state.lote_fallidas.pop(nombre_of, None)
                    except Exception as e:
                        st.session_state.lote_fallidas[nombre_of] = str(e)

                    st.session_state.lote_procesadas_en_corrida += 1
                    st.rerun()

            if st.session_state.lote_fallidas:
                with st.expander(f"❌ {len(st.session_state.lote_fallidas)} entidad(es) fallida(s) — ver motivo"):
                    for n, motivo in st.session_state.lote_fallidas.items():
                        st.caption(f"**{n}**: {motivo}")
                        if st.button(f"🔁 Reintentar '{n}'", key=f"reintentar_{n}"):
                            st.session_state.lote_fallidas.pop(n, None)
                            st.rerun()

            if st.session_state.lote_completadas:
                st.markdown("#### 4. Descargar lo generado hasta ahora")
                st.caption(
                    f"Puede descargar el .zip en cualquier momento con lo que ya esté listo — "
                    f"no hace falta esperar a que las {len(entidades_validas)} entidades terminen."
                )
                buffer_zip = io.BytesIO()
                with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for nombre_ent, archivos in st.session_state.lote_completadas.items():
                        carpeta = re.sub(r"[^\w\s-]", "", nombre_ent).strip().replace(" ", "_")
                        for nombre_archivo, contenido in archivos.items():
                            zf.writestr(f"{carpeta}/{nombre_archivo}", contenido)
                st.download_button(
                    f"⬇️ Descargar .zip ({len(st.session_state.lote_completadas)} entidades, "
                    f"{len(st.session_state.lote_completadas) * 6} archivos)",
                    data=buffer_zip.getvalue(),
                    file_name="informes_SIIEAP_lote.zip",
                    mime="application/zip",
                )
        else:
            st.caption("Pegue o suba la lista de entidades para comenzar.")

