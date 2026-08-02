"""Generador de Informe Técnico — Word (.docx) y PDF descargables.

Empaqueta, para UNA entidad, un informe técnico completo que el docente
puede entregar a sus estudiantes como evidencia de ejecución del sistema
y como fuente de trabajo para la asignatura. El informe integra:

  1. Portada institucional (SIIEAP, entidad, fecha)
  2. Contextualización de la Administración Pública Contemporánea:
     un bloque FIJO (igual en todos los informes) que recorre el arco
     teórico completo del microcurrículo — desde los antecedentes griegos
     de lo público/privado, pasando por el Nuevo Institucionalismo, la
     Nueva Gestión Pública (NGP) y la post-NGP, hasta el esquema
     integrador de la Administración Pública contemporánea (Estado
     Digital, Transformación Digital, Gobernanza Inteligente, IA,
     Gobierno Abierto, Gobernanza de Datos, Valor Público, Administración
     Pública Basada en Evidencia, Capacidades Estatales, Resiliencia
     Institucional, Agenda 2030/ODS, OCDE, CEPAL y Naciones Unidas).
  3. Resultado real del diagnóstico IDI-MIPG de la entidad (dimensiones,
     brechas).
  4. El análisis integral generado por IA (motor_analisis_ia.py): lectura
     desde las tres teorías del curso, recomendaciones técnicas/jurídicas/
     financieras, valoración de riesgo y prospectiva.
  5. Nota de trazabilidad y disclaimer académico.

Este módulo NO llama a la API de Claude — solo da formato a un texto de
análisis que ya fue generado antes por motor_analisis_ia.py.

Dependencias nuevas a agregar en requirements.txt:
    python-docx
    reportlab
"""
from __future__ import annotations

import io
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
)


# ---------------------------------------------------------------------------
# Bloque fijo: Contextualización de la Administración Pública Contemporánea
# ---------------------------------------------------------------------------

CONTEXTUALIZACION_TITULO = "Contextualización de la Administración Pública Contemporánea"

CONTEXTUALIZACION_INTRO = (
    "Este informe se enmarca en el recorrido teórico de la asignatura Enfoques y "
    "Teorías de la Administración Pública, que inicia en los antecedentes de la "
    "cultura griega para la distinción entre lo público y lo privado, y llega hasta "
    "los debates más recientes sobre la transformación digital del Estado. A "
    "continuación se sintetiza ese arco completo, como marco de lectura para el "
    "diagnóstico institucional que se presenta más adelante."
)

CONTEXTUALIZACION_ARCO_HISTORICO = [
    (
        "Antecedentes griegos de lo público y lo privado",
        "La distinción entre la esfera pública (la polis, los asuntos comunes) y la "
        "esfera privada (el oikos, el hogar) es el punto de partida clásico para "
        "pensar qué es \"lo público\" y cómo se organiza su gobierno.",
    ),
    (
        "Administración Pública clásica y campos múltiples de lo público",
        "Lo público se extiende más allá de las instituciones estatales: aparece en "
        "muchas esferas de la vida social, aunque el Estado y sus organizaciones "
        "sigan siendo su expresión institucional central.",
    ),
    (
        "Nuevo Institucionalismo",
        "Distingue las instituciones (reglas, normas, rutinas) de las organizaciones "
        "públicas, y explica cómo esas reglas moldean el comportamiento y los "
        "resultados de la gestión pública.",
    ),
    (
        "Nueva Gestión Pública (NGP)",
        "Busca homogenizar el fenómeno organizacional público retomando técnicas de "
        "gestión privada: eficiencia, orientación a resultados, indicadores de "
        "desempeño.",
    ),
    (
        "post-Nueva Gestión Pública (post-NGP)",
        "Da un giro para retomar conceptos del pasado (lo colectivo, lo público como "
        "valor) y enfrentar nuevas realidades: coordinación interinstitucional, "
        "gobernanza en red, valor público.",
    ),
    (
        "Gobernanza Pública y Valor Público",
        "La gobernanza reconoce múltiples actores (Estado, mercado, sociedad civil) "
        "coordinando la acción pública; el valor público mide el éxito de la gestión "
        "no solo en eficiencia, sino en el bienestar colectivo generado.",
    ),
    (
        "MIPG — Modelo Integrado de Planeación y Gestión",
        "Es la traducción institucional colombiana de estas corrientes: articula "
        "planeación, gestión del riesgo, talento humano, control interno y "
        "evaluación de resultados en un solo modelo de gestión pública.",
    ),
]

CONTEXTUALIZACION_ESQUEMA_INTEGRADOR = [
    "Estado Digital",
    "Transformación Digital",
    "Gobernanza Inteligente",
    "Inteligencia Artificial",
    "Gobierno Abierto",
    "Gobernanza de Datos",
    "Valor Público",
    "Administración Pública Basada en Evidencia",
    "Capacidades Estatales",
    "Resiliencia Institucional",
    "Agenda 2030 y Objetivos de Desarrollo Sostenible (ODS)",
]

CONTEXTUALIZACION_CIERRE = (
    "Este esquema integrador evidencia que la Administración Pública contemporánea "
    "no es una ruptura frente a las teorías clásicas, sino su evolución: el Estado "
    "Digital y la Transformación Digital habilitan una Gobernanza Inteligente "
    "apoyada en Inteligencia Artificial; el Gobierno Abierto y la Gobernanza de "
    "Datos fortalecen el Valor Público y una Administración Pública Basada en "
    "Evidencia; y todo ello construye Capacidades Estatales y Resiliencia "
    "Institucional, en línea con la Agenda 2030 y los ODS y los marcos de "
    "referencia de organismos internacionales como la OCDE, la CEPAL y las "
    "Naciones Unidas, que orientan a los países en la modernización de su gestión "
    "pública hacia la innovación, la transparencia y la generación de valor "
    "público."
)

DISCLAIMER_INFORME = (
    "Este informe fue generado por el Sistema Integral de Diagnóstico del "
    "Desempeño Institucional (SIIEAP), a partir de datos reales del Índice de "
    "Desempeño Institucional (IDI-MIPG) de Función Pública, complementado con un "
    "análisis generado por inteligencia artificial (Claude, Anthropic) como punto "
    "de partida académico y metodológico. No constituye un dictamen oficial de "
    "Función Pública ni sustituye la validación técnica, jurídica y del líder de "
    "proceso de la entidad analizada. Se entrega con fines académicos, como "
    "evidencia de ejecución del sistema y como insumo de trabajo para la "
    "asignatura."
)


def _fecha_hoy_es():
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
        "agosto", "septiembre", "octubre", "noviembre", "diciembre",
    ]
    hoy = datetime.now()
    return f"{hoy.day} de {meses[hoy.month - 1]} de {hoy.year}"


# ---------------------------------------------------------------------------
# Generación del .docx
# ---------------------------------------------------------------------------

def generar_reporte_docx(nombre_entidad, diag, analisis_ia_texto):
    """Devuelve un BytesIO con el informe técnico en formato Word (.docx)."""
    doc = Document()

    # Portada
    titulo = doc.add_heading("SIIEAP — Informe Técnico de Diagnóstico Institucional", level=0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(nombre_entidad)
    run.bold = True
    run.font.size = Pt(16)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"Generado el {_fecha_hoy_es()} · Índice de Desempeño Institucional (IDI-MIPG)")

    doc.add_page_break()

    # Contextualización académica (bloque fijo)
    doc.add_heading(CONTEXTUALIZACION_TITULO, level=1)
    doc.add_paragraph(CONTEXTUALIZACION_INTRO)

    doc.add_heading("Arco teórico: de los griegos a la Administración Pública contemporánea", level=2)
    for titulo_hito, texto_hito in CONTEXTUALIZACION_ARCO_HISTORICO:
        p = doc.add_paragraph()
        run_hito = p.add_run(f"{titulo_hito}. ")
        run_hito.bold = True
        p.add_run(texto_hito)

    doc.add_heading("Esquema integrador de la Administración Pública contemporánea", level=2)
    p_esquema = doc.add_paragraph()
    run_esquema = p_esquema.add_run(" → ".join(CONTEXTUALIZACION_ESQUEMA_INTEGRADOR))
    run_esquema.italic = True
    doc.add_paragraph(CONTEXTUALIZACION_CIERRE)

    doc.add_page_break()

    # Resultado real del diagnóstico
    doc.add_heading("Resultado del diagnóstico institucional (datos reales)", level=1)
    doc.add_paragraph(f"IDI estimado: {diag.idi_estimado}")

    tabla = doc.add_table(rows=1, cols=4)
    tabla.style = "Light Grid Accent 1"
    encabezados = tabla.rows[0].cells
    for celda, texto in zip(encabezados, ["Dimensión", "Promedio", "Riesgo", "Índices"]):
        celda.text = texto
        for parrafo in celda.paragraphs:
            for run_enc in parrafo.runs:
                run_enc.bold = True

    for r in diag.resultados_por_dimension:
        fila = tabla.add_row().cells
        fila[0].text = f"{r.codigo} {r.nombre}"
        fila[1].text = str(r.promedio)
        fila[2].text = str(r.nivel_riesgo)
        fila[3].text = f"{r.n_indices_evaluados}/{r.n_indices_esperados}"

    doc.add_paragraph()
    doc.add_heading("Brechas priorizadas", level=2)
    for b in diag.brechas[:15]:
        doc.add_paragraph(
            f"{b.codigo_indice} ({b.puntaje}) — {b.nombre_indice} · {b.politica}",
            style="List Bullet",
        )

    doc.add_page_break()

    # Análisis integral IA
    doc.add_heading("Análisis integral (generado por IA a partir de los datos reales)", level=1)
    for parrafo in analisis_ia_texto.split("\n"):
        if parrafo.strip():
            doc.add_paragraph(parrafo)

    doc.add_page_break()

    # Disclaimer
    doc.add_heading("Nota metodológica", level=2)
    p_disc = doc.add_paragraph()
    run_disc = p_disc.add_run(DISCLAIMER_INFORME)
    run_disc.italic = True
    run_disc.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# Generación del PDF (reportlab)
# ---------------------------------------------------------------------------

def generar_reporte_pdf(nombre_entidad, diag, analisis_ia_texto):
    """Devuelve un BytesIO con el informe técnico en formato PDF."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=LETTER,
        topMargin=2 * cm, bottomMargin=2 * cm,
        leftMargin=2 * cm, rightMargin=2 * cm,
    )
    estilos = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle("TituloSIIEAP", parent=estilos["Title"], fontSize=18)
    estilo_h2 = ParagraphStyle("H2SIIEAP", parent=estilos["Heading2"], spaceBefore=14)
    estilo_normal = estilos["Normal"]
    estilo_cursiva = ParagraphStyle("Cursiva", parent=estilos["Normal"], fontName="Helvetica-Oblique", textColor=colors.grey)

    elementos = []

    # Portada
    elementos.append(Paragraph("SIIEAP — Informe Técnico de Diagnóstico Institucional", estilo_titulo))
    elementos.append(Spacer(1, 12))
    elementos.append(Paragraph(f"<b>{nombre_entidad}</b>", estilos["Heading2"]))
    elementos.append(Paragraph(f"Generado el {_fecha_hoy_es()} · Índice de Desempeño Institucional (IDI-MIPG)", estilo_normal))
    elementos.append(PageBreak())

    # Contextualización
    elementos.append(Paragraph(CONTEXTUALIZACION_TITULO, estilos["Heading1"]))
    elementos.append(Paragraph(CONTEXTUALIZACION_INTRO, estilo_normal))
    elementos.append(Paragraph("Arco teórico: de los griegos a la Administración Pública contemporánea", estilo_h2))
    for titulo_hito, texto_hito in CONTEXTUALIZACION_ARCO_HISTORICO:
        elementos.append(Paragraph(f"<b>{titulo_hito}.</b> {texto_hito}", estilo_normal))
        elementos.append(Spacer(1, 6))

    elementos.append(Paragraph("Esquema integrador de la Administración Pública contemporánea", estilo_h2))
    elementos.append(Paragraph("<i>" + " → ".join(CONTEXTUALIZACION_ESQUEMA_INTEGRADOR) + "</i>", estilo_normal))
    elementos.append(Spacer(1, 8))
    elementos.append(Paragraph(CONTEXTUALIZACION_CIERRE, estilo_normal))
    elementos.append(PageBreak())

    # Resultado real del diagnóstico
    elementos.append(Paragraph("Resultado del diagnóstico institucional (datos reales)", estilos["Heading1"]))
    elementos.append(Paragraph(f"IDI estimado: {diag.idi_estimado}", estilo_normal))
    elementos.append(Spacer(1, 8))

    datos_tabla = [["Dimensión", "Promedio", "Riesgo", "Índices"]]
    for r in diag.resultados_por_dimension:
        datos_tabla.append([
            f"{r.codigo} {r.nombre}", str(r.promedio), str(r.nivel_riesgo),
            f"{r.n_indices_evaluados}/{r.n_indices_esperados}",
        ])
    tabla = Table(datos_tabla, hAlign="LEFT", colWidths=[7 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elementos.append(tabla)
    elementos.append(Spacer(1, 12))

    elementos.append(Paragraph("Brechas priorizadas", estilo_h2))
    for b in diag.brechas[:15]:
        elementos.append(Paragraph(
            f"• {b.codigo_indice} ({b.puntaje}) — {b.nombre_indice} · {b.politica}", estilo_normal,
        ))
    elementos.append(PageBreak())

    # Análisis integral IA
    elementos.append(Paragraph("Análisis integral (generado por IA a partir de los datos reales)", estilos["Heading1"]))
    for parrafo in analisis_ia_texto.split("\n"):
        if parrafo.strip():
            texto_escapado = parrafo.replace("&amp;", "&amp;amp;").replace("&lt;", "&amp;lt;").replace("&gt;", "&amp;gt;")
            elementos.append(Paragraph(texto_escapado, estilo_normal))
            elementos.append(Spacer(1, 4))
    elementos.append(PageBreak())

    # Disclaimer
    elementos.append(Paragraph("Nota metodológica", estilos["Heading2"]))
    elementos.append(Paragraph(DISCLAIMER_INFORME, estilo_cursiva))

    doc.build(elementos)
    buffer.seek(0)
    return buffer
