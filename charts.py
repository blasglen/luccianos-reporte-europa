"""
Genera los dos graficos del reporte como PNG (para incrustar en el mail via CID).
  charts/comparativo.png -> barras 2026 vs 2025 por sucursal (acumulado del mes)
  charts/progreso.png    -> barra horizontal total mes en curso vs anio anterior

Imagenes estaticas de alta resolucion (nitidas en pantallas retina).
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# Paleta corporativa: azul marino (en curso) + celeste (anio anterior)
AZUL = "#1f3a6e"        # azul marino - 2026
CELESTE = "#7fb3e0"     # celeste - 2025
GRIS_TXT = "#8a8a8a"
VERDE = "#1a7d2e"
ROJO = "#c62828"

DPI = 200  # alta resolucion para que no se vea pixelado


def _money_k(x, _):
    if x >= 1000:
        return f"{x/1000:.0f}k €"
    return f"{x:.0f} €"


def chart_comparativo(rows, a26_lbl, a25_lbl, out_path):
    branches = [r["branch"] for r in rows]
    v26 = [r["a26"] for r in rows]
    v25 = [(r["a25"] or 0) for r in rows]  # locales nuevos (a25=None) -> barra 2025 en cero

    n = len(branches)
    x = range(n)
    w = 0.38

    fig, ax = plt.subplots(figsize=(8.4, 3.8), dpi=DPI)
    ax.bar([i - w/2 for i in x], v26, width=w, label=a26_lbl, color=AZUL, zorder=3)
    ax.bar([i + w/2 for i in x], v25, width=w, label=a25_lbl, color=CELESTE, zorder=3)

    ax.set_xticks(list(x))
    ax.set_xticklabels(branches, fontsize=9, color="#333333")
    ax.yaxis.set_major_formatter(FuncFormatter(_money_k))
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TXT)
    ax.tick_params(axis="x", length=0)

    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color("#dddddd")
    ax.grid(axis="y", color="#eeeeee", zorder=0)

    ax.legend(loc="upper right", frameon=False, fontsize=9)
    fig.tight_layout(pad=0.6)
    fig.savefig(out_path, transparent=True, bbox_inches="tight")
    plt.close(fig)


def chart_progreso(a26, a25, pct, out_path):
    fig, ax = plt.subplots(figsize=(8.4, 1.25), dpi=DPI)

    base = max(a26, a25)
    ax.barh([0], [a25], height=0.5, color="#e9e9e9", zorder=2)
    color = VERDE if a26 >= a25 else ROJO
    ax.barh([0], [a26], height=0.5, color=color, zorder=3)
    ax.axvline(a25, color=GRIS_TXT, linestyle=":", linewidth=1.2, zorder=4)

    ax.set_xlim(0, base * 1.12)
    ax.set_ylim(-0.6, 0.6)
    ax.axis("off")

    ax.text(a26, 0, f"  {a26:,.0f} €", va="center", ha="left", fontsize=12,
            fontweight="bold", color="#111111", zorder=5)
    ax.text(a25, 0.42, f"Año ant. {a25:,.0f} €", va="bottom", ha="center",
            fontsize=8, color=GRIS_TXT)

    fig.tight_layout(pad=0.3)
    fig.savefig(out_path, transparent=True, bbox_inches="tight")
    plt.close(fig)


def chart_dias(dias, out_path):
    """Venta por dia de la semana, SOLO el anio en curso, con linea de promedio.

    Por que no compara contra 2025: como el espejo es por fecha calendario, el
    "Lun 13" de 2026 caeria al lado del 13/07/2025 que fue DOMINGO. En una
    heladeria un domingo puede duplicar a un lunes, asi que esa barra no compara
    performance: compara dias de semana distintos. Preferimos no mostrar una
    comparacion antes que mostrar una que enganie.

    El total de la semana SI se compara contra 2025 (en las KPI y en el
    comparativo por sucursal) porque cualquier ventana de 7 dias tiene
    exactamente un lunes, un martes... y un domingo. Ahi la composicion es
    identica y la comparacion es limpia.
    """
    etiquetas = [d["etiqueta"] for d in dias]
    v = [d["actual"] for d in dias]
    prom = sum(v) / len(v) if v else 0

    fig, ax = plt.subplots(figsize=(8.4, 3.2), dpi=DPI)
    # El mejor y el peor dia van resaltados; el resto en gris azulado.
    mx, mn = max(v), min(v)
    colores = [AZUL if x == mx else ("#c9d3e3" if x == mn else CELESTE) for x in v]
    ax.bar(range(len(v)), v, width=0.62, color=colores, zorder=3)

    ax.axhline(prom, color="#9a9a9a", linestyle="--", linewidth=1.1, zorder=4)
    ax.annotate(f"promedio {_money_k(prom, None)}", xy=(len(v) - 0.45, prom),
                xytext=(0, 5), textcoords="offset points", ha="right",
                fontsize=8.5, color="#777777")

    for i, x in enumerate(v):
        ax.annotate(_money_k(x, None), xy=(i, x), xytext=(0, 4),
                    textcoords="offset points", ha="center", fontsize=8.5,
                    color="#333333", fontweight="bold")

    ax.set_xticks(range(len(v)))
    ax.set_xticklabels(etiquetas, fontsize=9.5, color="#333333")
    ax.yaxis.set_major_formatter(FuncFormatter(_money_k))
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TXT)
    ax.tick_params(axis="x", length=0)
    ax.set_ylim(0, max(v) * 1.16 if v else 1)

    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color("#dddddd")
    ax.grid(axis="y", color="#eeeeee", zorder=0)

    fig.tight_layout(pad=0.6)
    fig.savefig(out_path, transparent=True, bbox_inches="tight")
    plt.close(fig)


def build_charts_semanal(rows, totals, dias, lbl_actual, lbl_ant, out_dir="charts"):
    """Los 3 graficos del semanal. Nombres de archivo distintos a los del diario
    para que los dos workflows no se pisen si corren el mismo viernes."""
    d = Path(out_dir)
    d.mkdir(exist_ok=True)
    p1 = d / "sem_comparativo.png"
    p2 = d / "sem_progreso.png"
    p3 = d / "sem_dias.png"
    chart_comparativo(rows, lbl_actual, lbl_ant, p1)
    chart_progreso(totals["a26"], totals["a25"], totals["pct"], p2)
    chart_dias(dias, p3)
    return str(p1), str(p2), str(p3)


def build_charts(rows, totals, mes, a26_lbl, a25_lbl, out_dir="charts"):
    d = Path(out_dir)
    d.mkdir(exist_ok=True)
    p1 = d / "comparativo.png"
    p2 = d / "progreso.png"
    chart_comparativo(rows, a26_lbl, a25_lbl, p1)
    chart_progreso(totals["a26"], totals["a25"], totals["pct"], p2)
    return str(p1), str(p2)


def chart_ytd(labels, v_ant, v_act, ytd_ant, ytd_act, lbl_ant, lbl_act, out_path):
    """Acumulado del anio (YTD): barras agrupadas por mes, anio anterior (celeste)
    vs actual (azul), enero..mes cerrado. Arriba, el total YTD de cada anio."""
    from matplotlib.patches import Patch
    x = list(range(len(labels)))
    w = 0.38
    fig, ax = plt.subplots(figsize=(8.4, 3.5), dpi=DPI)
    ax.bar([i - w / 2 for i in x], v_ant, width=w, color=CELESTE, zorder=3)
    ax.bar([i + w / 2 for i in x], v_act, width=w, color=AZUL, zorder=3)
    top = max(max(v_ant), max(v_act))
    ax.set_ylim(0, top * 1.30)
    ax.yaxis.set_major_formatter(FuncFormatter(_money_k))
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TXT)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=9.5, color="#333333")
    ax.tick_params(axis="x", length=0)
    for sp in ["top", "right", "left"]:
        ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color("#dddddd")
    ax.grid(axis="y", color="#f0f0f0", zorder=0)
    ax.annotate(f"Acum. {lbl_ant}: {ytd_ant/1e6:.2f}M €", xy=(0, top * 1.26), ha="left", va="top",
                fontsize=9.5, color="#3f5c86", fontweight="bold")
    ax.annotate(f"Acum. {lbl_act}: {ytd_act/1e6:.2f}M €", xy=(len(labels) - 1, top * 1.26), ha="right", va="top",
                fontsize=9.5, color=AZUL, fontweight="bold")
    ax.legend(handles=[Patch(color=CELESTE, label=lbl_ant), Patch(color=AZUL, label=lbl_act)],
              loc="lower center", bbox_to_anchor=(0.5, -0.26), ncol=2, frameon=False, fontsize=9)
    fig.tight_layout(pad=0.6)
    fig.savefig(out_path, transparent=True, bbox_inches="tight")
    plt.close(fig)
