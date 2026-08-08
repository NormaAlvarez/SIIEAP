"""Generador de gráficas para los informes SIIEAP.

Usa matplotlib para producir imágenes PNG a partir de los datos REALES del
diagnóstico (nunca datos inventados), pensadas para insertarse directamente
en los informes Word y PDF y así dar una presentación visual profesional
(gráfica de barras por dimensión, coloreada según el nivel de riesgo real).
"""
from __future__ import annotations

import io

import matplotlib
matplotlib.use("Agg")  # backend sin pantalla, seguro para servidores
import matplotlib.pyplot as plt

COLOR_POR_RIESGO = {
    "alta": "#C0392B",
    "media": "#E67E22",
    "baja": "#27AE60",
}
COLOR_DEFECTO = "#5B7C99"


def _color_de_riesgo(nivel_riesgo: str) -> str:
    if not nivel_riesgo:
        return COLOR_DEFECTO
    return COLOR_POR_RIESGO.get(str(nivel_riesgo).strip().lower(), COLOR_DEFECTO)


def generar_grafica_dimensiones(diag, titulo: str = "Resultado por dimensión (IDI-MIPG)") -> io.BytesIO:
    """
    Genera una gráfica de barras horizontales con el promedio real de cada
    dimensión del diagnóstico, coloreada según su nivel de riesgo real
    (Alta = rojo, Media = naranja, Baja = verde). Devuelve un BytesIO PNG.
    """
    dimensiones = list(diag.resultados_por_dimension)
    # Ordenar de menor a mayor promedio para que la barra más crítica quede arriba
    dimensiones = sorted(dimensiones, key=lambda r: (r.promedio if r.promedio is not None else 0))

    etiquetas = [f"{r.codigo} {r.nombre}" for r in dimensiones]
    valores = [r.promedio or 0 for r in dimensiones]
    colores = [_color_de_riesgo(r.nivel_riesgo) for r in dimensiones]

    alto = max(2.5, 0.55 * len(dimensiones) + 1)
    fig, ax = plt.subplots(figsize=(8.5, alto), dpi=160)
    barras = ax.barh(etiquetas, valores, color=colores, edgecolor="white", height=0.62)

    for barra, valor in zip(barras, valores):
        ax.text(
            barra.get_width() + 1.2, barra.get_y() + barra.get_height() / 2,
            f"{valor:.1f}", va="center", ha="left", fontsize=9, color="#333333",
        )

    ax.set_xlim(0, 105)
    ax.set_xlabel("Promedio (0 a 100)", fontsize=9)
    ax.set_title(titulo, fontsize=12, fontweight="bold", color="#1F3864", pad=12)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="y", labelsize=9)
    ax.tick_params(axis="x", labelsize=8)
    ax.axvline(60, color="#999999", linestyle="--", linewidth=0.8)
    ax.text(60.5, len(etiquetas) - 0.3 if etiquetas else 0, "umbral 60", fontsize=7, color="#777777")

    # Leyenda manual de colores de riesgo
    from matplotlib.patches import Patch
    leyenda = [
        Patch(facecolor=COLOR_POR_RIESGO["alta"], label="Riesgo Alto"),
        Patch(facecolor=COLOR_POR_RIESGO["media"], label="Riesgo Medio"),
        Patch(facecolor=COLOR_POR_RIESGO["baja"], label="Riesgo Bajo"),
    ]
    ax.legend(handles=leyenda, loc="lower right", fontsize=8, frameon=False)

    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer


def generar_matriz_riesgo_probabilidad_impacto(titulo: str = "Matriz de Riesgo — Probabilidad × Impacto") -> io.BytesIO:
    """
    Genera la matriz de riesgo 5x5 (Probabilidad × Impacto) usada oficialmente en
    auditoría al proceso financiero por auditores no financieros (metodología de
    Contraloría), con las mismas zonas de clasificación: Baja (verde), Moderada
    (amarillo), Alta (naranja) y Extrema (rojo). No usa datos de una entidad
    específica: es la matriz metodológica de referencia.
    """
    valores = [
        [5, 10, 15, 20, 25],
        [4, 8, 12, 16, 20],
        [3, 6, 9, 12, 15],
        [2, 4, 6, 8, 10],
        [1, 2, 3, 4, 5],
    ]

    def _color_de(v):
        if v < 3:
            return "#27AE60"  # Baja
        elif v < 6:
            return "#F1C40F"  # Moderada
        elif v < 15:
            return "#E67E22"  # Alta
        else:
            return "#C0392B"  # Extrema

    fig, ax = plt.subplots(figsize=(6.5, 6), dpi=160)
    for i in range(5):
        for j in range(5):
            valor = valores[i][j]
            color = _color_de(valor)
            ax.add_patch(plt.Rectangle((j, 4 - i), 1, 1, facecolor=color, edgecolor="white"))
            ax.text(j + 0.5, 4 - i + 0.5, str(valor), ha="center", va="center", fontsize=13, fontweight="bold", color="white")

    ax.set_xlim(0, 5)
    ax.set_ylim(0, 5)
    ax.set_xticks([0.5, 1.5, 2.5, 3.5, 4.5])
    ax.set_xticklabels(["1", "2", "3", "4", "5"])
    ax.set_yticks([0.5, 1.5, 2.5, 3.5, 4.5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"][::-1])
    ax.set_xlabel("Probabilidad", fontsize=10)
    ax.set_ylabel("Impacto", fontsize=10)
    ax.set_title(titulo, fontsize=12, fontweight="bold", color="#1F3864", pad=12)

    from matplotlib.patches import Patch
    leyenda = [
        Patch(facecolor="#27AE60", label="Baja (< 3)"),
        Patch(facecolor="#F1C40F", label="Moderada (3 a <6)"),
        Patch(facecolor="#E67E22", label="Alta (6 a <15)"),
        Patch(facecolor="#C0392B", label="Extrema (≥ 15)"),
    ]
    ax.legend(handles=leyenda, loc="upper left", bbox_to_anchor=(1.02, 1), fontsize=8, frameon=False)

    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer


def generar_grafica_brechas(diag, titulo: str = "Brechas priorizadas por puntaje", maximo: int = 15) -> io.BytesIO | None:
    """
    Genera una gráfica de barras horizontales con las brechas (índices con
    puntaje bajo), para dar contexto visual rápido. Devuelve None si no hay
    brechas que graficar.
    """
    if not diag.brechas:
        return None

    brechas = sorted(diag.brechas, key=lambda b: b.puntaje)[:maximo]
    etiquetas = [f"{b.codigo_indice} — {b.nombre_indice[:40]}" for b in brechas]
    valores = [b.puntaje for b in brechas]

    alto = max(2.5, 0.4 * len(brechas) + 1)
    fig, ax = plt.subplots(figsize=(8.5, alto), dpi=160)
    barras = ax.barh(etiquetas, valores, color="#C0392B", edgecolor="white", height=0.6)

    for barra, valor in zip(barras, valores):
        ax.text(
            barra.get_width() + 1.0, barra.get_y() + barra.get_height() / 2,
            f"{valor:.1f}", va="center", ha="left", fontsize=8, color="#333333",
        )

    ax.set_xlim(0, max(65, max(valores) + 10 if valores else 65))
    ax.set_xlabel("Puntaje (0 a 100)", fontsize=9)
    ax.set_title(titulo, fontsize=12, fontweight="bold", color="#1F3864", pad=12)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="y", labelsize=8)
    ax.tick_params(axis="x", labelsize=8)
    ax.invert_yaxis()

    fig.tight_layout()
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer
