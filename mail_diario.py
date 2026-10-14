"""
mail_diario.py — Arma el HTML del reporte diario de Europa (diseño "combinado").

Lo llama report.py (build_report) con los numeros ya calculados. Este modulo
solo DIBUJA: no toca el acumulado ni el historial.

Estructura del mail:
  1. Header negro (logo + fecha)
  2. Bloque "Venta de ayer": total, tickets, ticket promedio, vs mismo dia de
     la semana pasada, vs misma fecha del anio anterior, barra propias/franquicias
  3. Acumulado del mes: barra partida (locales con historia vs locales nuevos)
     + "dia por dia" contra el anio anterior
  4. Variacion del mes por sucursal (grafico divergente) + 2 destacados
  5. Tabla por sucursal + total
  6. Pie

Como esta hecho para que se vea bien en Gmail y Outlook:
  - Se escribe con clases + una hoja de estilos (CSS_BASE), pero al final
    `inline()` copia cada estilo DENTRO de cada elemento (style="..."). Gmail y
    Outlook ignoran muchas hojas de estilo; lo que esta "en linea" lo respetan.
  - Solo quedan en <style> las reglas de celular (@media max-width) y de modo
    oscuro (Apple Mail / iPhone). Llevan !important para ganarle a lo en linea.
  - Barras y graficos son TABLAS con fondo de color, no imagenes: se ven aunque
    el cliente bloquee imagenes, y nitidas en el celular.
"""
import calendar
import json
import re
from datetime import date, timedelta
from html import escape
from html.parser import HTMLParser
from pathlib import Path

from report import (BRANCH_ORDER, PROPIAS, FRANQUICIAS, SIN_HISTORICO_2025,
                    motivo_cierre)

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
DIAS_C = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]
MESES = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
MES_C = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
NB = "\u00a0"  # espacio que no corta (para que "28.228 €" no se parta en dos renglones)


# ----------------------------------------------------------------- formatos --
def eur(v, dec=0, signo=False):
    s = f"{abs(v):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    pre = "−" if v < 0 and round(abs(v), dec) else ("+" if signo and v > 0 else "")
    return f"{pre}{s}{NB}€"


def pct(v):
    return ("+" if v >= 0 else "−") + f"{abs(v):.1f}".replace(".", ",") + "%"


def num(v):
    return f"{v:,.0f}".replace(",", ".")


def chip(v, extra=""):
    cls, flecha = ("pos", "▲") if v >= 0 else ("neg", "▼")
    return f'<span class="chip {cls} {extra}">{flecha}{NB}{pct(v)}</span>'


def _lista(xs):
    return " y ".join(xs) if len(xs) <= 2 else ", ".join(xs[:-1]) + " y " + xs[-1]


# ------------------------------------------------------------------- datos --
def _historico(base_dir, anio, cache={}):
    if anio not in cache:
        p = Path(base_dir) / "data" / f"historico_{anio}.json"
        cache[anio] = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
    return cache[anio]


def _dia_hist(base_dir, d):
    """{local: {venta, tickets}} de un dia del historial 2026+, o None."""
    return _historico(base_dir, d.year).get(d.isoformat())


def _master_2025():
    """{fecha: {local: venta_neta}} del master del anio anterior (una sola lectura)."""
    from generar_acum_ant import _leer_master, IVA
    out = {}
    for f, locs in _leer_master().items():
        out[f] = {l: round(b / (1 + IVA), 2) for l, (b, _tk) in locs.items()}
    return out


def _espejo(d):
    try:
        return date(d.year - 1, d.month, d.day)
    except ValueError:            # 29/02 -> 28/02
        return date(d.year - 1, d.month, 28)


# -------------------------------------------------------------------- HTML --
def render(fecha, rows, totals, propias, franquicias, tickets, base_dir):
    """
    rows/totals/propias/franquicias: como los arma report.build_report.
    tickets: {local: tickets del dia} (de Ventas_ayer.xlsx).
    """
    base_dir = Path(base_dir)
    R = {r["branch"]: r for r in rows}
    NUE = [b for b in BRANCH_ORDER if b in SIN_HISTORICO_2025]
    COMP = [b for b in BRANCH_ORDER if b not in SIN_HISTORICO_2025]
    m25 = _master_2025()

    dtot = totals["dia"]
    tkt = sum(tickets.get(b, 0) for b in BRANCH_ORDER)
    tot = totals["a26"]
    comp26, comp25 = totals["ca26"], totals["ca25"]
    nuevos = tot - comp26
    var_m = totals["pct"]
    ndias = calendar.monthrange(fecha.year, fecha.month)[1]

    # Tickets del mes (historial: historico.py ya registro el dia de hoy)
    tk_mes = 0
    for i in range(1, fecha.day + 1):
        h = _dia_hist(base_dir, fecha.replace(day=i)) or {}
        tk_mes += sum((h.get(b) or {}).get("tickets", 0) for b in BRANCH_ORDER)
    tk_mes = max(tk_mes, tkt)

    # vs mismo dia de la semana pasada (total, todos los locales)
    sem = fecha - timedelta(days=7)
    hs = _dia_hist(base_dir, sem)
    dsem = sum(v["venta"] for v in hs.values()) if hs else None
    var_s = (dtot / dsem - 1) * 100 if dsem else None

    # vs misma fecha del anio anterior, a igual base (solo locales con historia)
    d25 = m25.get(_espejo(fecha))
    if d25:
        base25 = sum(d25.get(b, 0) for b in COMP)
        var_a = (sum(R[b]["dia"] for b in COMP) / base25 - 1) * 100 if base25 else None
    else:
        var_a = None

    # Cerrados por descanso (fijo o puntual)
    cerrados = [r["branch"] for r in rows if r.get("cerrado")]
    abiertos = len(rows) - len(cerrados)

    # Variacion por local
    var_b = {b: R[b]["pct"] for b in COMP}
    dif_b = {b: R[b]["diff"] for b in COMP}

    # ---------------- 1. header ----------------
    fecha_larga = f"{DIAS[fecha.weekday()]} {fecha.day} de {MESES[fecha.month]} de {fecha.year}"
    head = f'''
 <table class="head" role="presentation"><tr>
   <td class="hl"><img src="cid:logo" alt="Lucciano's" width="120" class="logo"></td>
   <td class="hr"><div class="hk">Reporte diario · Europa</div><div class="hd">{fecha_larga}</div></td>
 </tr></table>'''

    # ---------------- 2. venta de ayer ----------------
    if cerrados:
        sub = f"{abiertos} sucursales abiertas · {_lista(cerrados)} cerrad{'a' if len(cerrados) == 1 else 'as'} por descanso"
    else:
        sub = f"{len(rows)} sucursales abiertas"

    def celda(lbl, val, sub_, cls_val="", extra_cls=""):
        return (f'<td class="sc {extra_cls}"><div class="sl">{lbl}</div>'
                f'<div class="sv {cls_val}">{val}</div><div class="ss">{sub_}</div></td>')

    if var_s is None:
        c3 = celda(f"vs {DIAS_C[sem.weekday()]} {sem.day}/{sem.month}", "—", "sin dato", extra_cls="bot")
    else:
        c3 = celda(f"vs {DIAS_C[sem.weekday()]} {sem.day}/{sem.month}",
                   f"{'▲' if var_s >= 0 else '▼'} {pct(var_s)}", eur(dsem),
                   "pos-d" if var_s >= 0 else "neg-d", "bot")
    e25 = _espejo(fecha)
    if var_a is None:
        c4 = celda(f"vs {e25.day}/{e25.month}/{e25.year}", "—", "sin dato", extra_cls="bot")
    else:
        c4 = celda(f"vs {e25.day}/{e25.month}/{e25.year}", f"{'▲' if var_a >= 0 else '▼'} {pct(var_a)}",
                   "a igual base", "pos-d" if var_a >= 0 else "neg-d", "bot")

    tp_dia = eur(dtot / tkt, 2) if tkt else "—"
    tp_mes = eur(tot / tk_mes, 2) if tk_mes else "—"
    wp = propias["dia"] / dtot * 100 if dtot else 50
    hero = f'''
 <div class="sec">
  <div class="hero">
   <div class="hlbl">Venta de ayer · {DIAS_C[fecha.weekday()]} {fecha.day:02d}/{fecha.month:02d}</div>
   <div class="big">{eur(dtot)}</div>
   <div class="hsub">{sub}</div>
   <table class="strip" role="presentation"><tr>
     {celda("Tickets", num(tkt), f"{num(tk_mes)} en el mes", extra_cls="top")}
     {celda("Ticket prom.", tp_dia, f"{tp_mes} en el mes", extra_cls="top")}
     {c3}
     {c4}
   </tr></table>
   <table class="splitlbl" role="presentation"><tr>
     <td class="spl"><span class="dot sp"></span>Propias <b class="wh">{eur(propias["dia"])}</b></td>
     <td class="spr">Franquicias <b class="wh">{eur(franquicias["dia"])}</b><span class="dot sf dotr"></span></td>
   </tr></table>
   <table class="split" role="presentation"><tr><td class="sp spa" style="width:{wp:.1f}%"></td><td class="sf spb"></td></tr></table>
  </div>
 </div>'''

    # ---------------- 3. acumulado del mes ----------------
    mes_n = MESES[fecha.month]
    yy, yy1 = str(fecha.year)[2:], str(fecha.year - 1)[2:]
    escala = max(tot, comp25, 1)
    wcomp, wnue, w25 = comp26 / escala * 100, nuevos / escala * 100, comp25 / escala * 100

    if NUE:
        n_locales = len(COMP)
        sub_mes = (f'<b class="{"pos-t" if comp26 >= comp25 else "neg-t"}">{eur(comp26 - comp25, signo=True)}</b> '
                   f'vs {mes_n} {fecha.year - 1} en los {n_locales} locales con historia, al mismo día')
        fila26 = (f'<td class="s-comp" style="width:{wcomp:.2f}%"></td>'
                  f'<td class="s-new" style="width:{wnue:.2f}%"></td><td class="s-bg"></td>')
        leyenda = (f'<div class="legend"><span class="it"><span class="dot s-comp"></span>Con historia {eur(comp26)}</span>'
                   f'<span class="it"><span class="dot s-new"></span>{_lista(NUE)} {eur(nuevos)}</span></div>')
    else:
        sub_mes = (f'<b class="{"pos-t" if comp26 >= comp25 else "neg-t"}">{eur(comp26 - comp25, signo=True)}</b> '
                   f'vs {mes_n} {fecha.year - 1} al mismo día')
        fila26 = f'<td class="s-comp" style="width:{wcomp:.2f}%"></td><td class="s-bg"></td>'
        leyenda = ""

    # Dia por dia (locales con historia, contra la misma fecha del anio anterior)
    serie = []
    for i in range(1, fecha.day + 1):
        dd = fecha.replace(day=i)
        h = _dia_hist(base_dir, dd) or {}
        v26 = sum((h.get(b) or {}).get("venta", 0) for b in COMP)
        v25 = sum((m25.get(_espejo(dd)) or {}).get(b, 0) for b in COMP)
        serie.append((dd, v26, v25))
    smax = max([max(a, b) for _, a, b in serie] + [1])
    H = 46

    def barra(cls, alto):
        # Cada barra va en su propia tablita apoyada abajo: si fuera una celda
        # suelta, se estiraria hasta el alto de la fila (y todas medirian igual).
        alto = int(round(alto))
        if alto < 1:
            return '<td class="dslot"><div class="dgap"></div></td>'
        return (f'<td class="dslot" valign="bottom"><table class="dbt" role="presentation"><tr>'
                f'<td class="dbar {cls}" style="height:{alto}px"></td></tr></table></td>')

    cols = ""
    for dd, v26, v25 in serie:
        k = " k" if dd.day in (1, 5, 10, 15, 20, 25, 30) or dd == fecha else ""
        we = " we" if dd.weekday() >= 5 else ""
        cols += (f'<td class="dcol{we}{k}"><table class="dbars" role="presentation" align="center"><tr>'
                 f'{barra("d25", v25 / smax * H)}<td class="dsep"></td>{barra("d26", v26 / smax * H)}'
                 f'</tr></table><div class="dl">{dd.day}</div></td>')
    for i in range(fecha.day + 1, ndias + 1):
        k = " k" if i in (15, 20, 25, 30) else ""
        cols += (f'<td class="dcol fut{k}"><table class="dbars" role="presentation" align="center"><tr>'
                 f'<td class="dslot" valign="bottom"><table class="dbt" role="presentation"><tr>'
                 f'<td class="dbar dfut" style="height:3px"></td></tr></table></td></tr></table><div class="dl">{i}</div></td>')

    mes = f'''
 <div class="sec">
  <table class="mh" role="presentation"><tr>
    <td><div class="lbl">Acumulado de {mes_n}</div><div class="big2">{eur(tot)}</div></td>
    <td class="mhr">{chip(var_m, "lg")}</td></tr></table>
  <div class="msub">{sub_mes}</div>
  <table class="av" role="presentation">
   <tr><td class="yr">{MES_C[fecha.month]}/{yy}</td><td class="trk"><table class="track" role="presentation"><tr>{fila26}</tr></table></td><td class="amt">{eur(tot)}</td></tr>
   <tr><td class="yr mut">{MES_C[fecha.month]}/{yy1}</td><td class="trk"><table class="track" role="presentation"><tr><td class="s-25" style="width:{w25:.2f}%"></td><td class="s-bg"></td></tr></table></td><td class="amt mut">{eur(comp25)}</td></tr>
  </table>
  {leyenda}
  <table class="dhead" role="presentation"><tr><td class="lbl0">Día por día contra {fecha.year - 1}</td><td class="dhr">Día {fecha.day} de {ndias}</td></tr></table>
  <table class="ritmo" role="presentation"><tr>{cols}</tr></table>
 </div>'''

    # ---------------- 4. variacion por sucursal ----------------
    # El cero se corre segun donde esten los datos: si casi todo es negativo,
    # la mitad roja ocupa mas lugar. Misma escala para los dos lados.
    mneg = max([-v for v in var_b.values() if v < 0] + [0])
    mpos = max([v for v in var_b.values() if v > 0] + [0])
    total_rango = (mneg + mpos) or 1
    WL = min(max(mneg / total_rango * 100, 12), 88)
    WR = 100 - WL
    ESP = 0.62   # las dos mitades usan el 62% del ancho; el resto: nombre y %
    filas_div = ""
    for b in sorted(var_b, key=var_b.get, reverse=True):
        v = var_b[b]
        ref = mneg if v < 0 else mpos
        w = max(abs(v) / ref * 100, 2) if ref else 0
        neg = (f'<table class="dv" role="presentation" align="right" style="width:{w:.0f}%"><tr><td class="dvn"></td></tr></table>'
               if v < 0 else "")
        pos = (f'<table class="dv" role="presentation" style="width:{w:.0f}%"><tr><td class="dvp"></td></tr></table>'
               if v >= 0 else "")
        filas_div += (f'<tr><td class="dn">{b}</td><td class="dh dhl" style="width:{WL * ESP:.1f}%">{neg}</td>'
                      f'<td class="dh dhr" style="width:{WR * ESP:.1f}%">{pos}</td>'
                      f'<td class="dp {"pos-t" if v >= 0 else "neg-t"}">{pct(v)}</td></tr>')
    for b in NUE:
        filas_div += (f'<tr><td class="dn mut">{b}</td><td colspan="2" class="dh dhc">'
                      f'<span class="mut sm0">abrió en {fecha.year}, sin comparación</span></td>'
                      f'<td class="dp mut">—</td></tr>')

    # Destacados: lo que el grafico no dice
    dest = []
    if dif_b:
        brecha = min(dif_b, key=dif_b.get)
        if dif_b[brecha] < 0 and comp26 < comp25:
            dest.append(("warn", f"<b>{brecha}</b> explica el {abs(dif_b[brecha]) / (comp25 - comp26) * 100:.0f}% "
                                 f"de lo que falta para igualar {fecha.year - 1} ({eur(dif_b[brecha])})."))
        mejor = max(dif_b, key=dif_b.get)
        if comp26 >= comp25 and dif_b[mejor] > 0:
            dest.append(("warn", f"<b>{mejor}</b> es el que más suma contra {fecha.year - 1}: "
                                 f"{eur(dif_b[mejor], signo=True)}."))
    if NUE and nuevos > 0:
        quienes = f"<b>{_lista(NUE)}</b>"
        if comp26 < comp25 and tot >= comp25:
            dest.append(("info", f"Sin {quienes} el mes viene {pct(var_m)}; con ellos, el total supera a "
                                 f"{mes_n} {fecha.year - 1} por {eur(tot - comp25)}."))
        elif tot < comp25:
            dest.append(("info", f"Aun sumando {quienes} ({eur(nuevos)}), el total queda "
                                 f"{eur(comp25 - tot)} abajo de {mes_n} {fecha.year - 1}."))
        else:
            dest.append(("info", f"{quienes} suman {eur(nuevos)} extra sobre una base que ya viene {pct(var_m)}."))
    dest_html = ""
    if dest:
        dest_html = ('<table class="dest" role="presentation">' +
                     "".join(f'<tr><td class="mk"><div class="mkd {c}"></div></td><td class="mkt">{t}</td></tr>'
                             for c, t in dest) + "</table>")

    variacion = f'''
 <div class="sec">
  <div class="h">Variación del mes por sucursal</div>
  <table class="divg" role="presentation">{filas_div}</table>
  {dest_html}
 </div>'''

    # ---------------- 5. tabla ----------------
    def fila(b):
        r = R[b]
        t = tickets.get(b, 0)
        if r.get("cerrado"):
            linea = "cerrado por descanso"
            ayer = '<span class="mut it2">cerrado</span>'
        else:
            linea = (f'{num(t)} tickets · <span class="nw">prom. {eur(r["dia"] / t, 2)}</span>'
                     if t else "sin tickets")
            ayer = eur(r["dia"])
        if r["comparable"]:
            sub25 = f'<div class="mut sm">vs {eur(r["a25"])}</div>'
            var = (f'{chip(r["pct"])}<div class="sm b {"pos-t" if r["diff"] >= 0 else "neg-t"}">'
                   f'{eur(r["diff"], signo=True)}</div>')
        else:
            sub25 = f'<div class="mut sm">sin {fecha.year - 1}</div>'
            var = '<span class="chip flat">nuevo</span>'
        return (f'<tr><td class="c-name"><div class="nm">{b}</div><div class="mut sm">{linea}</div></td>'
                f'<td class="c-num">{ayer}</td><td class="c-acum"><div class="b">{eur(r["a26"])}</div>{sub25}</td>'
                f'<td class="c-var">{var}</td></tr>')

    def subtotal(nombre, s, locs):
        st = sum(tickets.get(b, 0) for b in locs)
        prom = f"prom. {eur(s['dia'] / st, 2)}" if st else ""
        ast = "*" if s.get("hay_nuevo") else ""
        return (f'<tr><td class="c-name sub subl"><div class="nm">{nombre}</div><div class="mut sm">{prom}</div></td>'
                f'<td class="c-num sub">{eur(s["dia"])}</td>'
                f'<td class="c-acum sub"><div class="b">{eur(s["a26"])}</div><div class="mut sm">vs {eur(s["ca25"])}{ast}</div></td>'
                f'<td class="c-var sub subr">{chip(s["pct"])}<div class="sm b {"pos-t" if s["diff"] >= 0 else "neg-t"}">'
                f'{eur(s["diff"], signo=True)}</div></td></tr>')

    def grupo(t):
        return f'<tr><td colspan="4" class="grp">{t}</td></tr>'

    filas = (grupo("Propias") + "".join(fila(b) for b in PROPIAS) + subtotal("Subtotal propias", propias, PROPIAS) +
             grupo("Franquicias") + "".join(fila(b) for b in FRANQUICIAS) +
             subtotal("Subtotal franquicias", franquicias, FRANQUICIAS))
    ast = "*" if totals.get("hay_nuevo") else ""
    tabla = f'''
 <div class="sec">
  <div class="h">Detalle por sucursal</div>
  <table class="tbl" role="presentation">
   <tr><th class="th thl">Sucursal</th><th class="th c-num">Ayer</th><th class="th c-acum">Acum. {MES_C[fecha.month].lower()}</th><th class="th c-var">vs {fecha.year - 1}</th></tr>
   {filas}
  </table>
  <table class="tot" role="presentation"><tr>
   <td class="tt"><div>Total general</div><div class="ts">prom. {tp_dia}</div></td>
   <td class="tt r c-num">{eur(dtot)}</td>
   <td class="tt r"><div>{eur(tot)}</div><div class="ts">vs {eur(comp25)}{ast}</div></td>
   <td class="tt r tv">{chip(var_m, "inv")}<div class="ts {"pos-tot" if comp26 >= comp25 else "neg-tot"}">{eur(comp26 - comp25, signo=True)}</div></td>
  </tr></table>
  {_notas(fecha, rows, NUE)}
 </div>'''

    pie = f'<div class="foot">Lucciano\'s Europa · Ventas netas sin IVA · Datos al {fecha.strftime("%d/%m/%Y")}</div>'

    pre = f"Ayer {eur(dtot)} · {mes_n.capitalize()} {pct(var_m)} vs {fecha.year - 1}"
    if NUE:
        pre += " en locales con historia"

    cuerpo = f'''<table class="mail" role="presentation" width="100%"><tr><td class="mailtd">
<div class="pre">{escape(pre)}</div>
<table class="wrap" role="presentation" align="center"><tr><td class="wraptd">
<div class="card">{head}{hero}{mes}{variacion}{tabla}{pie}</div>
</td></tr></table>
</td></tr></table>'''

    return _documento(inline(cuerpo, CSS_BASE), f"Reporte diario {fecha.strftime('%d/%m/%Y')}")


def _notas(fecha, rows, NUE):
    notas = []
    if NUE:
        notas.append(f"* La variación y el año anterior se calculan solo sobre locales con historia "
                     f"{fecha.year - 1}; {_lista(NUE)} suman al total pero no a la variación.")
    fijos = [r["branch"] for r in rows if r.get("cerrado") and motivo_cierre(r["branch"], fecha) == "fijo"]
    punt = [r["branch"] for r in rows if r.get("cerrado") and motivo_cierre(r["branch"], fecha) == "puntual"]
    if fijos:
        verbo = "cierra" if len(fijos) == 1 else "cierran"
        notas.append(f"{_lista(fijos)} {verbo} los {DIAS[fecha.weekday()].lower()} (descanso fijo): "
                     f"hoy no tiene venta y no es un faltante.")
    if punt:
        verbo = "no abrió" if len(punt) == 1 else "no abrieron"
        notas.append(f"{_lista(punt)} {verbo} hoy (descanso corrido de día por un feriado): no es un faltante.")
    return "".join(f'<div class="nota">{n}</div>' for n in notas)


# ===================================================================== CSS ==
# Paleta: negro #1a1a1a · papel #faf9f7 · indigo #4f46e5 · gris calido #857a6b
#         verde #15803d / rojo #b91c1c solo para variaciones · ambar #b45309 atencion
FUENTE = '"Inter Tight",-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif'
CSS_BASE = r'''
.mail{background:#faf9f7;width:100%}
.mailtd{padding:24px 8px 28px;font-family:FUENTE;color:#1a1a1a}
.pre{display:none;max-height:0;overflow:hidden;font-size:1px;line-height:1px;color:#faf9f7}
.wrap{width:100%;max-width:600px}
.wraptd{padding:0}
.card{background:#ffffff;border:1px solid #eae7e1;border-radius:16px;overflow:hidden;font-family:FUENTE;color:#1a1a1a}
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
.h{font-size:15px;font-weight:700;margin:0 0 12px;color:#1a1a1a}
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
.dbars{height:46px;margin:0 auto}
.dslot{padding:0;vertical-align:bottom;height:46px}
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
.dn{width:104px;font-size:13px}
.c-acum{width:100px}
.dbar{width:3px}
.dgap{width:3px}
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
.mut{color:#8b93a1}
.lbl{color:#8b93a1}
.lbl0{color:#8b93a1}
.msub{color:#8b93a1}
.legend{color:#8b93a1}
.dl{color:#8b93a1}
.dhr{color:#8b93a1}
.nota{color:#8b93a1}
.sub{background:#1d2530}
.flat{background:#1d2530}
.dest{background:#1d2530}
.s-bg{background:#1d2530}
.pos{background:#16301f;color:#4ade80}
.neg{background:#3a1e1e;color:#f87171}
.pos-t{color:#4ade80}
.neg-t{color:#f87171}
.s-comp{background:#818cf8}
.d26{background:#818cf8}
.s-new{background:#3b3f73}
.s-25{background:#a79d8f}
.d25{background:#3a4250}
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


def _documento(cuerpo, titulo):
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
{cuerpo}
</body>
</html>'''


# ================================================================= INLINER ==
# Copia los estilos de CSS_BASE adentro de cada elemento (style="...").
# Soporta lo que usa este archivo: .clase, tag, tag.clase, .a.b, descendientes
# (".a .b td"), :first-child y :last-child. Respeta especificidad y orden, como
# un navegador. Los estilos que ya estaban en linea (anchos de barras) ganan.
VACIOS = {"img", "br", "meta", "link", "hr", "input", "col", "source"}


class _Nodo:
    __slots__ = ("tag", "attrs", "hijos", "padre")

    def __init__(self, tag, attrs, padre):
        self.tag, self.attrs, self.hijos, self.padre = tag, dict(attrs), [], padre

    def clases(self):
        return set((self.attrs.get("class") or "").split())

    def elementos_hermanos(self):
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
        tag, clases, pseudo = m.group(1), [c for c in m.group(2).split(".") if c], m.group(3)
        partes.append((tag, set(clases), pseudo))
    espec = (sum(len(c) + (1 if ps else 0) for _, c, ps in partes), sum(1 for t, _, _ in partes if t))
    return partes, espec


def _calza(n, comp):
    tag, clases, pseudo = comp
    if tag and n.tag != tag:
        return False
    if clases and not clases <= n.clases():
        return False
    if pseudo:
        hs = n.elementos_hermanos()
        if pseudo == ":first-child" and hs[0] is not n:
            return False
        if pseudo == ":last-child" and hs[-1] is not n:
            return False
    return True


def _calza_sel(n, partes):
    if not _calza(n, partes[-1]):
        return False
    i = len(partes) - 2
    a = n.padre
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
                imp = "!important" in v
                decls.append((k.strip(), v.replace("!important", "").strip(), imp))
        for s in sel.split(","):
            partes, espec = _compilar(s.strip())
            reglas.append((partes, espec, orden, decls))
            orden += 1
    return reglas


def _aplicar(n, reglas):
    if not isinstance(n, _Nodo):
        return
    if n.tag != "#raiz":
        # Cascada: lo que se escribe despues gana. Se borra y se vuelve a poner
        # cada propiedad para que quede DESPUES en el style (si no, "padding:0"
        # escrito mas tarde no pisaria a un "padding-top" anterior).
        def poner(d, k, v):
            d.pop(k, None)
            d[k] = v
        normal, importante = {}, {}
        for partes, espec, orden, decls in sorted(reglas, key=lambda r: (r[1], r[2])):
            if _calza_sel(n, partes):
                for k, v, imp in decls:
                    poner(importante if imp else normal, k, v)
        final = dict(normal)
        for d in (n.attrs.get("style") or "").split(";"):
            if ":" in d:
                k, v = d.split(":", 1)
                poner(final, k.strip(), v.strip())
        for k, v in importante.items():
            poner(final, k, v)
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
