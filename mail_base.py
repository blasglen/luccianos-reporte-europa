"""
mail_base.py — Diseño común de los mails de Lucciano's Europa (diario, semanal,
cierre mensual). Paleta, formatos, componentes y el "inliner".

Un cambio de estilo se hace ACA y lo toman los tres reportes.

Como se dibuja un mail (para que se vea bien en Gmail y Outlook):
  - Se escribe con clases + la hoja CSS_BASE. Al final `documento()` pasa todo
    por `inline()`, que copia cada estilo DENTRO de cada elemento (style="...").
    Gmail y Outlook ignoran muchas hojas de estilo; lo "en linea" lo respetan.
  - En <style> solo quedan las reglas de celular (@media max-width) y de modo
    oscuro (Apple Mail / iPhone), con !important para ganarle a lo en linea.
  - Barras y graficos son TABLAS con fondo de color, no imagenes: se ven aunque
    el cliente bloquee imagenes. La unica imagen es el logo (cid:logo).

Paleta: negro #1a1a1a · papel #faf9f7 · indigo #4f46e5 · gris calido #857a6b
        verde #15803d / rojo #b91c1c solo para variaciones · ambar #b45309 atencion
"""
import json
import re
from datetime import date, timedelta
from html import escape
from html.parser import HTMLParser
from pathlib import Path

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
DIAS_C = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]
MESES = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MES_C = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
NB = "\u00a0"  # espacio que no corta: "28.228 €" nunca se parte en dos renglones
BASE = Path(__file__).parent


# ================================================================ FORMATOS ==
def eur(v, dec=0, signo=False):
    s = f"{abs(v):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    pre = "−" if v < 0 and round(abs(v), dec) else ("+" if signo and v > 0 else "")
    return f"{pre}{s}{NB}€"


def eur_corto(v, euro=False):
    """1.820.000 -> '1,82 M€' · 934.000 -> '934k' (o '934 k€') · 4.700 -> '4,7k'.
    Para graficos, donde no entra el numero completo."""
    a = abs(v)
    k = f"{NB}k€" if euro else "k"
    if a >= 1e6:
        return f"{v / 1e6:.2f}".replace(".", ",") + f"{NB}M€"
    if a >= 1e4:
        return f"{v / 1e3:.0f}{k}"
    if a >= 1e3:
        return f"{v / 1e3:.1f}".replace(".", ",") + k
    return f"{v:.0f}" + (f"{NB}€" if euro else "")


def pct(v):
    return ("+" if v >= 0 else "−") + f"{abs(v):.1f}".replace(".", ",") + "%"


def num(v):
    return f"{v:,.0f}".replace(",", ".")


def chip(v, extra=""):
    cls, flecha = ("pos", "▲") if v >= 0 else ("neg", "▼")
    return f'<span class="chip {cls} {extra}">{flecha}{NB}{pct(v)}</span>'


def flecha_pct(v):
    return f"{'▲' if v >= 0 else '▼'} {pct(v)}"


def lista(xs):
    xs = list(xs)
    return " y ".join(xs) if len(xs) <= 2 else ", ".join(xs[:-1]) + " y " + xs[-1]


# =================================================================== DATOS ==
_hist_cache, _master_cache = {}, {}


def historico(anio):
    if anio not in _hist_cache:
        p = BASE / "data" / f"historico_{anio}.json"
        _hist_cache[anio] = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
    return _hist_cache[anio]


def _master():
    """{fecha: {local: venta_neta}} del master del anio anterior (una sola lectura)."""
    if not _master_cache:
        from generar_acum_ant import _leer_master, IVA
        for f, locs in _leer_master().items():
            _master_cache[f] = {l: round(b / (1 + IVA), 2) for l, (b, _tk) in locs.items()}
    return _master_cache


def venta_dia(d):
    """{local: venta neta} de un dia, de donde haya dato: historial diario (2026 en
    adelante) o master del anio anterior. None si no hay nada."""
    h = historico(d.year).get(d.isoformat())
    if h:
        return {l: (v["venta"] if isinstance(v, dict) else float(v)) for l, v in h.items()}
    return _master().get(d)


def tickets_dia(d):
    h = historico(d.year).get(d.isoformat()) or {}
    return {l: (v.get("tickets", 0) if isinstance(v, dict) else 0) for l, v in h.items()}


def mismo_dia_semana(d):
    """El mismo dia de la semana, un anio antes (364 dias): martes contra martes.
    Para comparar UN dia. Los acumulados se siguen comparando por fecha calendario."""
    return d - timedelta(days=364)


def espejo(d):
    """Misma fecha calendario, anio anterior (29/02 -> 28/02)."""
    try:
        return date(d.year - 1, d.month, d.day)
    except ValueError:
        return date(d.year - 1, d.month, 28)


# ============================================================= COMPONENTES ==
def header(titulo, fecha_txt):
    return f'''
 <table class="head" role="presentation"><tr>
   <td class="hl"><img src="cid:logo" alt="Lucciano's" width="120" class="logo"></td>
   <td class="hr"><div class="hk">{titulo}</div><div class="hd">{fecha_txt}</div></td>
 </tr></table>'''


def celda(lbl, val, sub, cls_val="", extra=""):
    """Una celda de la franja de 4 datos del bloque oscuro."""
    return (f'<td class="sc {extra}"><div class="sl">{lbl}</div>'
            f'<div class="sv {cls_val}">{val}</div><div class="ss">{sub}</div></td>')


def celda_var(lbl, v, sub, extra=""):
    if v is None:
        return celda(lbl, "—", "sin dato", extra=extra)
    return celda(lbl, flecha_pct(v), sub, "pos-d" if v >= 0 else "neg-d", extra)


def hero(lbl, total, sub, celdas, propias_v, franquicias_v):
    """Bloque oscuro: total grande + franja de 4 datos + barra propias/franquicias.
    celdas: 4 strings de celda(); las 2 primeras arriba en el celular."""
    c = list(celdas)
    c[0] = c[0].replace('class="sc ', 'class="sc top ', 1)
    c[1] = c[1].replace('class="sc ', 'class="sc top ', 1)
    c[2] = c[2].replace('class="sc ', 'class="sc bot ', 1)
    c[3] = c[3].replace('class="sc ', 'class="sc bot ', 1)
    tot = propias_v + franquicias_v
    wp = propias_v / tot * 100 if tot else 50
    return f'''
 <div class="sec">
  <div class="hero">
   <div class="hlbl">{lbl}</div>
   <div class="big">{eur(total)}</div>
   <div class="hsub">{sub}</div>
   <table class="strip" role="presentation"><tr>{"".join(c)}</tr></table>
   <table class="splitlbl" role="presentation"><tr>
     <td class="spl"><span class="dot sp"></span>Propias <b class="wh">{eur(propias_v)}</b></td>
     <td class="spr">Franquicias <b class="wh">{eur(franquicias_v)}</b><span class="dot sf dotr"></span></td>
   </tr></table>
   <table class="split" role="presentation"><tr><td class="sp spa" style="width:{wp:.1f}%"></td><td class="sf spb"></td></tr></table>
  </div>
 </div>'''


def barra_partida(c26, nuevos, c25, lbl26, lbl25, nombres_nuevos):
    """Dos barras: este anio partido en 'con historia' + 'nuevos', y el anterior."""
    tot = c26 + nuevos
    esc = max(tot, c25, 1)
    if nombres_nuevos and nuevos > 0:
        fila = (f'<td class="s-comp" style="width:{c26 / esc * 100:.2f}%"></td>'
                f'<td class="s-new" style="width:{nuevos / esc * 100:.2f}%"></td><td class="s-bg"></td>')
        ley = (f'<div class="legend"><span class="it"><span class="dot s-comp"></span>Con historia {eur(c26)}</span>'
               f'<span class="it"><span class="dot s-new"></span>{lista(nombres_nuevos)} {eur(nuevos)}</span></div>')
    else:
        fila = f'<td class="s-comp" style="width:{c26 / esc * 100:.2f}%"></td><td class="s-bg"></td>'
        ley = ""
    return f'''
  <table class="av" role="presentation">
   <tr><td class="yr">{lbl26}</td><td class="trk"><table class="track" role="presentation"><tr>{fila}</tr></table></td><td class="amt">{eur(tot)}</td></tr>
   <tr><td class="yr mut">{lbl25}</td><td class="trk"><table class="track" role="presentation"><tr><td class="s-25" style="width:{c25 / esc * 100:.2f}%"></td><td class="s-bg"></td></tr></table></td><td class="amt mut">{eur(c25)}</td></tr>
  </table>
  {ley}'''


def _barra(cls, alto):
    # Cada barra va en su propia tablita apoyada abajo: si fuera una celda suelta,
    # se estiraria hasta el alto de la fila y todas medirian igual.
    alto = int(round(alto))
    if alto < 1:
        return '<td class="dslot"><div class="dgap"></div></td>'
    return (f'<td class="dslot" valign="bottom"><table class="dbt" role="presentation"><tr>'
            f'<td class="dbar {cls}" style="height:{alto}px"></td></tr></table></td>')


def ritmo(items, alto=46, futuros=()):
    """Mini barras pareadas (2025 gris / 2026 indigo) para muchos dias.
    items: [(etiqueta, v26, v25, clases_extra)]. futuros: etiquetas de dias por venir."""
    vmax = max([max(a or 0, b or 0) for _, a, b, _ in items] + [1])
    cols = ""
    for et, v26, v25, extra in items:
        cols += (f'<td class="dcol {extra}"><table class="dbars" role="presentation" align="center"><tr>'
                 f'{_barra("d25", (v25 or 0) / vmax * alto)}<td class="dsep"></td>'
                 f'{_barra("d26", (v26 or 0) / vmax * alto)}</tr></table><div class="dl">{et}</div></td>')
    for et, extra in futuros:
        cols += (f'<td class="dcol fut {extra}"><table class="dbars" role="presentation" align="center"><tr>'
                 f'{_barra("dfut", 3)}</tr></table><div class="dl">{et}</div></td>')
    return f'<table class="ritmo" role="presentation"><tr>{cols}</tr></table>'


def columnas(items, alto=96, ocultar_valores_celular=False):
    """Barras verticales grandes, pareadas, con el valor arriba de la barra de este anio.
    items: [(etiqueta, v26, v25_o_None, resaltar)]. Para 7 dias o 12 meses."""
    vmax = max([max(a or 0, b or 0) for _, a, b, _ in items] + [1])
    ocultar = " vlm" if ocultar_valores_celular else ""
    cols = ""
    for et, v26, v25, res in items:
        b25 = _barra("c25", (v25 or 0) / vmax * alto) if v25 is not None else ""
        sep = '<td class="csep"></td>' if v25 is not None else ""
        cls26 = "c26"
        cols += (f'<td class="ccol"><div class="cval{ocultar}">{eur_corto(v26) if v26 else "—"}</div>'
                 f'<table class="cbars" role="presentation" align="center"><tr>{b25}{sep}{_barra(cls26, (v26 or 0) / vmax * alto)}'
                 f'</tr></table><div class="clab{" b" if res else ""}">{et}</div></td>')
    return f'<table class="cols" role="presentation"><tr>{cols}</tr></table>'


def divergente(var_b, nuevos, anio_nuevos):
    """Grafico de variacion por sucursal, de mayor a menor. El cero se corre segun
    donde esten los datos (si casi todo es negativo, la mitad roja tiene mas lugar)."""
    mneg = max([-v for v in var_b.values() if v < 0] + [0])
    mpos = max([v for v in var_b.values() if v > 0] + [0])
    rango = (mneg + mpos) or 1
    WL = min(max(mneg / rango * 100, 12), 88)
    WR = 100 - WL
    ESP = 0.62   # las dos mitades usan el 62% del ancho; el resto: nombre y %
    filas = ""
    for b in sorted(var_b, key=var_b.get, reverse=True):
        v = var_b[b]
        ref = mneg if v < 0 else mpos
        w = max(abs(v) / ref * 100, 2) if ref else 0
        neg = (f'<table class="dv" role="presentation" align="right" style="width:{w:.0f}%"><tr><td class="dvn"></td></tr></table>'
               if v < 0 else "")
        pos = (f'<table class="dv" role="presentation" style="width:{w:.0f}%"><tr><td class="dvp"></td></tr></table>'
               if v >= 0 else "")
        filas += (f'<tr><td class="dn">{b}</td><td class="dh dhl" style="width:{WL * ESP:.1f}%">{neg}</td>'
                  f'<td class="dh dhr" style="width:{WR * ESP:.1f}%">{pos}</td>'
                  f'<td class="dp {"pos-t" if v >= 0 else "neg-t"}">{pct(v)}</td></tr>')
    for b in nuevos:
        filas += (f'<tr><td class="dn mut">{b}</td><td colspan="2" class="dh dhc">'
                  f'<span class="mut sm0">abrió en {anio_nuevos}, sin comparación</span></td>'
                  f'<td class="dp mut">—</td></tr>')
    return f'<table class="divg" role="presentation">{filas}</table>'


def destacados(dif_b, c26, c25, nuevos_nombres, nuevos_monto, periodo, anio_ant):
    """Dos frases con lo que el grafico no dice. periodo: 'el mes', 'la semana'..."""
    dest = []
    if dif_b:
        brecha = min(dif_b, key=dif_b.get)
        if dif_b[brecha] < 0 and c26 < c25:
            dest.append(("warn", f"<b>{brecha}</b> explica el {abs(dif_b[brecha]) / (c25 - c26) * 100:.0f}% "
                                 f"de lo que falta para igualar {anio_ant} ({eur(dif_b[brecha])})."))
        mejor = max(dif_b, key=dif_b.get)
        if c26 >= c25 and dif_b[mejor] > 0:
            dest.append(("warn", f"<b>{mejor}</b> es el que más suma contra {anio_ant}: "
                                 f"{eur(dif_b[mejor], signo=True)}."))
    if nuevos_nombres and nuevos_monto > 0:
        q = f"<b>{lista(nuevos_nombres)}</b>"
        var = (c26 / c25 - 1) * 100 if c25 else 0
        tot = c26 + nuevos_monto
        if c26 < c25 <= tot:
            dest.append(("info", f"Sin {q} {periodo} viene {pct(var)}; con ellos, el total supera a "
                                 f"{anio_ant} por {eur(tot - c25)}."))
        elif tot < c25:
            dest.append(("info", f"Aun sumando {q} ({eur(nuevos_monto)}), el total queda "
                                 f"{eur(c25 - tot)} abajo de {anio_ant}."))
        else:
            dest.append(("info", f"{q} suman {eur(nuevos_monto)} extra sobre una base que ya viene {pct(var)}."))
    if not dest:
        return ""
    return ('<table class="dest" role="presentation">' +
            "".join(f'<tr><td class="mk"><div class="mkd {c}"></div></td><td class="mkt">{t}</td></tr>'
                    for c, t in dest) + "</table>")


def barras_horizontales(valores, total=None):
    """Una barra por sucursal, de mayor a menor, con monto y participacion."""
    total = total or sum(valores.values()) or 1
    vmax = max(list(valores.values()) + [1])
    filas = ""
    for b in sorted(valores, key=valores.get, reverse=True):
        v = valores[b]
        filas += (f'<tr><td class="hn">{b}</td><td class="hb"><table class="hbt" role="presentation" '
                  f'style="width:{max(v / vmax * 100, 1):.1f}%"><tr><td class="hbar"></td></tr></table></td>'
                  f'<td class="hv">{eur_corto(v, euro=True)}<span class="mut hs"> · {v / total * 100:.0f}%</span></td></tr>')
    return f'<table class="hbars" role="presentation">{filas}</table>'


def callout(texto):
    return f'<table class="callout" role="presentation"><tr><td class="cot">{texto}</td></tr></table>'


def tarjeta(*secciones, pie=""):
    return f'<div class="card">{"".join(secciones)}{pie}</div>'


def pie(texto):
    return f'<div class="foot">{texto}</div>'


# ===================================================================== CSS ==
FUENTE = '"Inter Tight",-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif'
CSS_BASE = r'''
.mail{background:#faf9f7;width:100%}
.mailtd{padding:24px 8px 28px;font-family:FUENTE;color:#1a1a1a;font-variant-numeric:tabular-nums;-webkit-text-size-adjust:100%}
.pre{display:none;max-height:0;overflow:hidden;font-size:1px;line-height:1px;color:#faf9f7}
.wrap{width:100%;max-width:600px}
.wraptd{padding:0}
.card{background:#ffffff;border:1px solid #eae7e1;border-radius:16px;overflow:hidden;font-family:FUENTE;color:#1a1a1a;font-variant-numeric:tabular-nums}
.head{width:100%;background:#1a1a1a}
.hl{padding:20px 24px;vertical-align:middle}
.hr{padding:20px 24px;vertical-align:middle;text-align:right}
.logo{display:block;width:120px;height:auto;border:0}
.hk{color:#a8a39a;font-size:12px;font-weight:600}
.hd{color:#ffffff;font-size:15px;font-weight:700;margin-top:3px}
.sec{padding:22px 24px;border-bottom:1px solid #eae7e1}
.hero{background:#1a1a1a;border-radius:14px;padding:20px 20px 18px;color:#ffffff}
.hlbl{font-size:12px;font-weight:600;color:#a8a39a}
.big{font-size:46px;font-weight:800;letter-spacing:-1.5px;line-height:1.05;margin-top:6px;color:#ffffff}
.hsub{font-size:13px;color:#a8a39a;margin-top:2px}
.strip{width:100%;margin-top:16px;background:#2a2826;border-radius:10px}
.sc{width:25%;text-align:center;padding:11px 6px;border-left:1px solid #3a3733;vertical-align:top;color:#ffffff}
.sc:first-child{border-left:0}
.sl{font-size:11px;color:#a8a39a;font-weight:600}
.sv{font-size:18px;font-weight:800;margin-top:3px;white-space:nowrap;color:#ffffff}
.ss{font-size:11px;color:#a8a39a;margin-top:2px}
.pos-d{color:#86efac}
.neg-d{color:#fca5a5}
.splitlbl{width:100%;margin-top:16px}
.spl{font-size:12px;color:#d6d0c6;padding:0 0 6px}
.spr{font-size:12px;color:#d6d0c6;padding:0 0 6px;text-align:right}
.wh{color:#ffffff}
.split{width:100%}
.spa{height:7px;font-size:0;line-height:0;padding:0;border-radius:4px 0 0 4px}
.spb{height:7px;font-size:0;line-height:0;padding:0;border-radius:0 4px 4px 0}
.sp{background:#818cf8}
.sf{background:#d6d0c6}
.dot{display:inline-block;width:9px;height:9px;border-radius:2px;margin:0 5px 0 0;vertical-align:-1px}
.dotr{margin:0 0 0 5px}
.lbl{font-size:13px;font-weight:600;color:#6b665e;margin-bottom:4px}
.lbl0{font-size:13px;font-weight:600;color:#6b665e;padding:0}
.h{font-size:15px;font-weight:700;margin:0 0 4px;color:#1a1a1a}
.hsubt{font-size:13px;color:#6b665e;margin:0 0 14px;line-height:1.45}
.mh{width:100%}
.mhr{text-align:right;vertical-align:top}
.big2{font-size:32px;font-weight:800;letter-spacing:-.8px;line-height:1.1;color:#1a1a1a}
.msub{font-size:13px;color:#6b665e;margin:6px 0 14px;line-height:1.45}
.chip{display:inline-block;font-size:12px;font-weight:700;border-radius:999px;padding:3px 9px;white-space:nowrap}
.lg{font-size:14px;padding:5px 12px}
.pos{color:#15803d;background:#e8f3ec}
.neg{color:#b91c1c;background:#fbecec}
.flat{color:#6b665e;background:#f5f3ef;font-weight:600}
.inv.neg{background:#3a1e1e;color:#fca5a5}
.inv.pos{background:#16301f;color:#86efac}
.pos-t{color:#15803d}
.neg-t{color:#b91c1c}
.mut{color:#6b665e}
.sm{font-size:12px;margin-top:2px}
.sm0{font-size:12px}
.b{font-weight:700}
.nw{white-space:nowrap}
.it2{font-style:italic;font-size:13px}
.av{width:100%}
.yr{width:52px;padding:4px 0;font-size:13px;font-weight:700;vertical-align:middle}
.trk{padding:4px 0;vertical-align:middle}
.amt{width:92px;padding:4px 0;text-align:right;font-weight:700;font-size:13px;white-space:nowrap;vertical-align:middle}
.track{width:100%}
.track td{height:16px;padding:0;font-size:0;line-height:0}
.track td:first-child{border-radius:5px 0 0 5px}
.track td:last-child{border-radius:0 5px 5px 0}
.s-comp{background:#4f46e5}
.s-new{background:#c7c4f5}
.s-25{background:#857a6b}
.s-bg{background:#f5f3ef}
.legend{font-size:12px;color:#6b665e;margin-top:8px}
.legend .it{margin-right:14px;white-space:nowrap}
.dhead{width:100%;margin-top:20px;border-top:1px dashed #eae7e1}
.dhead td{padding-top:14px;padding-bottom:8px}
.dhr{text-align:right;font-size:12px;color:#6b665e}
.ritmo{width:100%;table-layout:fixed}
.dcol{text-align:center;vertical-align:bottom;padding:0}
.dbars{margin:0 auto}
.dslot{padding:0;vertical-align:bottom}
.dbt{border-collapse:collapse}
.dbar{width:5px;padding:0;font-size:0;line-height:0;border-radius:2px 2px 0 0}
.dgap{width:5px;height:0;font-size:0;line-height:0}
.dsep{width:1px;padding:0;font-size:0;line-height:0}
.d26{background:#4f46e5}
.d25{background:#d6d0c6}
.dfut{background:#eae7e1;width:9px;border-radius:2px}
.dl{font-size:10px;color:#6b665e;margin-top:4px}
.we .dl{color:#1a1a1a;font-weight:700}
.fut .dl{color:#c4beb3}
.cols{width:100%;table-layout:fixed}
.ccol{text-align:center;vertical-align:bottom;padding:0 2px}
.cval{font-size:11px;font-weight:700;color:#1a1a1a;margin-bottom:4px;white-space:nowrap}
.cbars{margin:0 auto}
.csep{width:2px;padding:0;font-size:0;line-height:0}
.c26{width:14px;padding:0;font-size:0;line-height:0;background:#4f46e5;border-radius:3px 3px 0 0}
.c26h{width:14px;padding:0;font-size:0;line-height:0;background:#1a1a1a;border-radius:3px 3px 0 0}
.c25{width:14px;padding:0;font-size:0;line-height:0;background:#d6d0c6;border-radius:3px 3px 0 0}
.clab{font-size:11px;color:#6b665e;margin-top:5px;white-space:nowrap}
.clab.b{color:#1a1a1a}
.divg{width:100%}
.dn{width:130px;padding:6px 8px 6px 0;font-size:14px;vertical-align:middle}
.dh{padding:6px 0;vertical-align:middle}
.dhl{border-right:1px solid #dad6ce}
.dhc{text-align:center}
.dv{border-collapse:collapse}
.dvn{height:14px;padding:0;font-size:0;line-height:0;background:#b91c1c;border-radius:4px 0 0 4px}
.dvp{height:14px;padding:0;font-size:0;line-height:0;background:#15803d;border-radius:0 4px 4px 0}
.dp{width:64px;padding:6px 0;text-align:right;font-weight:700;font-size:14px;white-space:nowrap;vertical-align:middle}
.dest{width:100%;margin-top:14px;background:#f5f3ef;border-radius:10px}
.mk{width:14px;padding:15px 0 0 12px;vertical-align:top}
.mkd{width:8px;height:8px;border-radius:999px;font-size:0;line-height:0}
.mkd.warn{background:#b45309}
.mkd.info{background:#4f46e5}
.mkt{padding:9px 12px 9px 8px;font-size:13px;line-height:1.45;vertical-align:top;color:#1a1a1a}
.callout{width:100%;margin-top:14px;background:#fcf1e1;border-radius:10px}
.cot{padding:11px 14px;font-size:13px;line-height:1.5;color:#92400e}
.hbars{width:100%}
.hn{width:120px;padding:5px 8px 5px 0;font-size:13px;vertical-align:middle}
.hb{padding:5px 0;vertical-align:middle}
.hbt{border-collapse:collapse}
.hbar{height:12px;padding:0;font-size:0;line-height:0;background:#857a6b;border-radius:0 4px 4px 0}
.hv{width:104px;padding:5px 0 5px 8px;text-align:right;font-size:13px;font-weight:700;white-space:nowrap;vertical-align:middle}
.hs{font-weight:500;font-size:12px}
.tbl{width:100%}
.th{font-size:12px;font-weight:600;color:#6b665e;text-align:right;padding:0 0 8px 10px;border-bottom:1px solid #dad6ce;white-space:nowrap}
.thl{text-align:left;padding-left:0}
.tbl td{padding-top:11px;padding-bottom:11px;border-bottom:1px solid #eae7e1;vertical-align:top;font-size:14px}
.c-name{text-align:left;padding-right:8px}
.c-num{text-align:right;padding-left:10px;white-space:nowrap}
.c-acum{text-align:right;padding-left:10px;white-space:nowrap;width:112px}
.c-var{text-align:right;padding-left:10px;white-space:nowrap;width:96px}
.nm{font-weight:700}
.tbl td.grp{font-size:12px;font-weight:700;color:#4f46e5;padding-top:16px;padding-bottom:2px;border-bottom:0}
.tbl td.sub{background:#f5f3ef;font-size:13px}
.subl{padding-left:10px;border-radius:8px 0 0 8px}
.subr{padding-right:10px;border-radius:0 8px 8px 0}
.tot{width:100%;margin-top:12px;background:#1a1a1a;border-radius:12px}
.tt{color:#ffffff;padding:14px 12px;font-weight:700;font-size:15px;vertical-align:top}
.r{text-align:right;white-space:nowrap}
.tv{width:96px}
.ts{font-size:12px;font-weight:500;color:#a8a39a;margin-top:2px}
.neg-tot{color:#fca5a5;font-weight:700}
.pos-tot{color:#86efac;font-weight:700}
.nota{font-size:12px;color:#6b665e;line-height:1.55;margin-top:10px}
.foot{text-align:center;font-size:12px;color:#a8a39a;background:#1a1a1a;padding:14px 16px}
'''.replace("FUENTE", FUENTE)

# Celular (va en <style>, con !important para ganarle a lo en linea)
CSS_MOBILE = r'''
.mailtd{padding:0 0 20px}
.card{border-radius:0;border-left:0;border-right:0}
.sec{padding:18px 14px}
.hl{padding:16px 14px}
.hr{padding:16px 14px}
.logo{width:96px}
.hd{font-size:13px}
.big{font-size:40px}
.sc{display:inline-block;width:50%;box-sizing:border-box;border-left:0}
.bot{border-top:1px solid #3a3733}
.c-num{display:none}
.m-hide{display:none}
.dn{width:104px;font-size:13px}
.hn{width:96px}
.c-acum{width:100px}
.dbar{width:3px}
.dgap{width:3px}
.c26{width:9px}
.c26h{width:9px}
.c25{width:9px}
.vlm{visibility:hidden}
.dl{visibility:hidden}
.k .dl{visibility:visible}
'''

# Modo oscuro (Apple Mail / iPhone). Gmail lo ignora y aplica su propia inversion.
CSS_DARK = r'''
.mail{background:#0f141b}
.card{background:#161c25;border-color:#28303c;color:#e6e9ee}
.head{background:#0b0e13}
.hero{background:#0b0e13}
.tot{background:#0b0e13}
.foot{background:#0b0e13}
.strip{background:#1d2530}
.sc{border-color:#28303c}
.sec{border-color:#28303c}
.tbl td{border-color:#28303c}
.th{border-color:#38424f;color:#8b93a1}
.h{color:#e6e9ee}
.big2{color:#e6e9ee}
.mkt{color:#e6e9ee}
.cval{color:#e6e9ee}
.clab.b{color:#e6e9ee}
.mut{color:#8b93a1}
.lbl{color:#8b93a1}
.lbl0{color:#8b93a1}
.msub{color:#8b93a1}
.hsubt{color:#8b93a1}
.legend{color:#8b93a1}
.dl{color:#8b93a1}
.clab{color:#8b93a1}
.dhr{color:#8b93a1}
.nota{color:#8b93a1}
.sub{background:#1d2530}
.flat{background:#1d2530}
.dest{background:#1d2530}
.s-bg{background:#1d2530}
.callout{background:#3a2e14}
.cot{color:#fbbf24}
.pos{background:#16301f;color:#4ade80}
.neg{background:#3a1e1e;color:#f87171}
.pos-t{color:#4ade80}
.neg-t{color:#f87171}
.s-comp{background:#818cf8}
.d26{background:#818cf8}
.c26{background:#818cf8}
.c26h{background:#e6e9ee}
.s-new{background:#3b3f73}
.s-25{background:#a79d8f}
.hbar{background:#a79d8f}
.d25{background:#3a4250}
.c25{background:#3a4250}
.dvn{background:#f87171}
.dvp{background:#4ade80}
.dhl{border-color:#38424f}
.grp{color:#818cf8}
.we .dl{color:#e6e9ee}
.fut .dl{color:#3a4250}
.dfut{background:#28303c}
.mkd.warn{background:#fbbf24}
.mkd.info{background:#818cf8}
'''


def _importante(css):
    """Agrega !important a cada declaracion (para las reglas de <style>)."""
    def fix(m):
        decls = [d.strip() for d in m.group(2).split(";") if d.strip()]
        return m.group(1) + "{" + ";".join(d if "!important" in d else d + " !important" for d in decls) + "}"
    return re.sub(r"([^{}]+)\{([^{}]*)\}", fix, css)


def documento(card_html, titulo, preheader):
    """Envuelve la tarjeta, pasa todo por el inliner y arma el HTML final."""
    cuerpo = f'''<table class="mail" role="presentation" width="100%"><tr><td class="mailtd">
<div class="pre">{escape(preheader)}</div>
<table class="wrap" role="presentation" align="center"><tr><td class="wraptd">
{card_html}
</td></tr></table>
</td></tr></table>'''
    return f'''<!DOCTYPE html>
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="x-apple-disable-message-reformatting">
<meta name="color-scheme" content="light dark">
<meta name="supported-color-schemes" content="light dark">
<title>{escape(titulo)}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter+Tight:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{{color-scheme:light dark;supported-color-schemes:light dark}}
body{{margin:0;padding:0;background:#faf9f7;-webkit-text-size-adjust:100%;font-variant-numeric:tabular-nums}}
table{{border-collapse:collapse}}
@media only screen and (max-width:600px){{
{_importante(CSS_MOBILE)}
}}
@media (prefers-color-scheme:dark){{
body{{background:#0f141b !important}}
{_importante(CSS_DARK)}
}}
</style>
</head>
<body style="margin:0;padding:0;background:#faf9f7;">
{inline(cuerpo, CSS_BASE)}
</body>
</html>'''


# ================================================================= INLINER ==
# Copia los estilos de CSS_BASE adentro de cada elemento (style="...").
# Soporta lo que usan estos mails: .clase, tag, tag.clase, .a.b, descendientes
# (".a .b td"), :first-child y :last-child. Respeta especificidad y orden, como
# un navegador. Lo que ya estaba en linea (anchos y altos de barras) gana.
VACIOS = {"img", "br", "meta", "link", "hr", "input", "col", "source"}


class _Nodo:
    __slots__ = ("tag", "attrs", "hijos", "padre")

    def __init__(self, tag, attrs, padre):
        self.tag, self.attrs, self.hijos, self.padre = tag, dict(attrs), [], padre

    def clases(self):
        return set((self.attrs.get("class") or "").split())

    def hermanos(self):
        return [h for h in self.padre.hijos if isinstance(h, _Nodo)] if self.padre else [self]


class _Parser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=False)
        self.raiz = _Nodo("#raiz", {}, None)
        self.actual = self.raiz

    def handle_starttag(self, tag, attrs):
        n = _Nodo(tag, attrs, self.actual)
        self.actual.hijos.append(n)
        if tag not in VACIOS:
            self.actual = n

    def handle_startendtag(self, tag, attrs):
        self.actual.hijos.append(_Nodo(tag, attrs, self.actual))

    def handle_endtag(self, tag):
        n = self.actual
        while n is not self.raiz and n.tag != tag:
            n = n.padre
        if n is not self.raiz:
            self.actual = n.padre

    def handle_data(self, d):
        self.actual.hijos.append(d)

    def handle_entityref(self, name):
        self.actual.hijos.append(f"&{name};")

    def handle_charref(self, name):
        self.actual.hijos.append(f"&#{name};")


_COMP = re.compile(r"^([a-z][a-z0-9]*)?((?:\.[\w-]+)*)(:first-child|:last-child)?$")


def _compilar(selector):
    partes = []
    for p in selector.split():
        m = _COMP.match(p)
        if not m:
            raise ValueError(f"Selector no soportado por el inliner: {selector!r}")
        partes.append((m.group(1), {c for c in m.group(2).split(".") if c}, m.group(3)))
    espec = (sum(len(c) + (1 if ps else 0) for _, c, ps in partes), sum(1 for t, _, _ in partes if t))
    return partes, espec


def _calza(n, comp):
    tag, clases, pseudo = comp
    if tag and n.tag != tag:
        return False
    if clases and not clases <= n.clases():
        return False
    if pseudo:
        hs = n.hermanos()
        if pseudo == ":first-child" and hs[0] is not n:
            return False
        if pseudo == ":last-child" and hs[-1] is not n:
            return False
    return True


def _calza_sel(n, partes):
    if not _calza(n, partes[-1]):
        return False
    i, a = len(partes) - 2, n.padre
    while i >= 0 and a is not None:
        if a.tag != "#raiz" and _calza(a, partes[i]):
            i -= 1
        a = a.padre
    return i < 0


def _reglas(css):
    reglas, orden = [], 0
    for sel, cuerpo in re.findall(r"([^{}]+)\{([^{}]*)\}", css):
        decls = []
        for d in cuerpo.split(";"):
            if ":" in d:
                k, v = d.split(":", 1)
                decls.append((k.strip(), v.replace("!important", "").strip(), "!important" in v))
        for s in sel.split(","):
            partes, espec = _compilar(s.strip())
            reglas.append((partes, espec, orden, decls))
            orden += 1
    return sorted(reglas, key=lambda r: (r[1], r[2]))


def _poner(d, k, v):
    # Cascada: lo que se escribe despues gana. Se borra y se vuelve a poner para
    # que quede DESPUES en el style ("padding:0" posterior pisa a "padding-top").
    d.pop(k, None)
    d[k] = v


def _aplicar(n, reglas):
    if not isinstance(n, _Nodo):
        return
    if n.tag != "#raiz":
        normal, importante = {}, {}
        for partes, _espec, _orden, decls in reglas:
            if _calza_sel(n, partes):
                for k, v, imp in decls:
                    _poner(importante if imp else normal, k, v)
        final = dict(normal)
        for d in (n.attrs.get("style") or "").split(";"):
            if ":" in d:
                k, v = d.split(":", 1)
                _poner(final, k.strip(), v.strip())
        for k, v in importante.items():
            _poner(final, k, v)
        if final:
            n.attrs["style"] = ";".join(f"{k}:{v}" for k, v in final.items())
        if n.tag == "table":
            for a, v in (("cellpadding", "0"), ("cellspacing", "0"), ("border", "0")):
                n.attrs.setdefault(a, v)
    for h in n.hijos:
        _aplicar(h, reglas)


def _serializar(n):
    if isinstance(n, str):
        return n
    if n.tag == "#raiz":
        return "".join(_serializar(h) for h in n.hijos)
    attrs = "".join(f' {k}="{escape(v, quote=True)}"' if v is not None else f" {k}" for k, v in n.attrs.items())
    if n.tag in VACIOS:
        return f"<{n.tag}{attrs}>"
    return f"<{n.tag}{attrs}>" + "".join(_serializar(h) for h in n.hijos) + f"</{n.tag}>"


def inline(html, css):
    p = _Parser()
    p.feed(html)
    p.close()
    _aplicar(p.raiz, _reglas(css))
    return _serializar(p.raiz)
