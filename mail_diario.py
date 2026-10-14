"""
mail_diario.py — HTML del reporte DIARIO de Europa (diseño "combinado").

Lo llama report.py (build_report) con los numeros ya calculados. Este modulo
solo DIBUJA: no toca el acumulado ni el historial. El estilo, los componentes y
el inliner estan en mail_base.py (compartidos con el semanal y el cierre).

Estructura:
  1. Header negro (logo + fecha)
  2. "Venta de ayer": total, tickets, ticket promedio, vs mismo dia de la semana
     pasada, vs mismo dia de la semana del anio anterior, propias/franquicias
  3. Acumulado del mes: barra partida (con historia / locales nuevos) + dia por dia
  4. Variacion del mes por sucursal (grafico divergente) + 2 destacados
  5. Tabla por sucursal + total
  6. Pie

Comparaciones de UN dia: contra el MISMO DIA DE LA SEMANA del anio anterior
(364 dias antes: martes contra martes). Si se comparara por fecha, el 13/10/2026
(martes) iria contra el 13/10/2025 (lunes), y en una heladeria un sabado no se
compara con un viernes. Los ACUMULADOS siguen por fecha calendario.
"""
import calendar
from datetime import timedelta

from report import BRANCH_ORDER, PROPIAS, FRANQUICIAS, SIN_HISTORICO_2025, motivo_cierre
from mail_base import (DIAS, DIAS_C, MESES, MES_C, eur, pct, num, chip, lista,
                       venta_dia, tickets_dia, mismo_dia_semana, header, celda, celda_var,
                       hero, barra_partida, ritmo, divergente, destacados, tarjeta, pie,
                       documento)


def render(fecha, rows, totals, propias, franquicias, tickets, base_dir=None):
    """rows/totals/propias/franquicias: como los arma report.build_report.
    tickets: {local: tickets del dia} (de Ventas_ayer.xlsx)."""
    R = {r["branch"]: r for r in rows}
    NUE = [b for b in BRANCH_ORDER if b in SIN_HISTORICO_2025]
    COMP = [b for b in BRANCH_ORDER if b not in SIN_HISTORICO_2025]
    ant = fecha.year - 1

    dtot = totals["dia"]
    tkt = sum(tickets.get(b, 0) for b in BRANCH_ORDER)
    tot, c26, c25 = totals["a26"], totals["ca26"], totals["ca25"]
    nuevos = tot - c26
    var_m = totals["pct"]
    ndias = calendar.monthrange(fecha.year, fecha.month)[1]
    mes_n = MESES[fecha.month]

    # Tickets del mes (historico.py ya registro el dia de hoy antes de report.py)
    tk_mes = sum(sum(tickets_dia(fecha.replace(day=i)).values()) for i in range(1, fecha.day + 1))
    tk_mes = max(tk_mes, tkt)

    # vs mismo dia de la semana pasada (todos los locales)
    sem = fecha - timedelta(days=7)
    vs = venta_dia(sem)
    dsem = sum(vs.values()) if vs else None
    var_s = (dtot / dsem - 1) * 100 if dsem else None

    # vs mismo dia de la semana del anio anterior, a igual base (locales con historia)
    f25 = mismo_dia_semana(fecha)
    v25 = venta_dia(f25)
    base25 = sum((v25 or {}).get(b, 0) for b in COMP)
    var_a = (sum(R[b]["dia"] for b in COMP) / base25 - 1) * 100 if base25 else None

    cerrados = [r["branch"] for r in rows if r.get("cerrado")]
    if cerrados:
        sub = (f"{len(rows) - len(cerrados)} sucursales abiertas · {lista(cerrados)} "
               f"cerrad{'a' if len(cerrados) == 1 else 'as'} por descanso")
    else:
        sub = f"{len(rows)} sucursales abiertas"

    # ---- 2. venta de ayer
    bloque_dia = hero(
        f"Venta de ayer · {DIAS_C[fecha.weekday()]} {fecha.day:02d}/{fecha.month:02d}", dtot, sub,
        [celda("Tickets", num(tkt), f"{num(tk_mes)} en el mes"),
         celda("Ticket prom.", eur(dtot / tkt, 2) if tkt else "—",
               f"{eur(tot / tk_mes, 2)} en el mes" if tk_mes else ""),
         celda_var(f"vs {DIAS_C[sem.weekday()]} {sem.day}/{sem.month}", var_s, eur(dsem) if dsem else ""),
         celda_var(f"vs {DIAS_C[f25.weekday()]} {f25.day}/{f25.month}/{f25.year}", var_a, "a igual base")],
        propias["dia"], franquicias["dia"])

    # ---- 3. acumulado del mes + dia por dia
    if NUE:
        sub_mes = (f'<b class="{"pos-t" if c26 >= c25 else "neg-t"}">{eur(c26 - c25, signo=True)}</b> '
                   f'vs {mes_n} {ant} en los {len(COMP)} locales con historia, al mismo día')
    else:
        sub_mes = (f'<b class="{"pos-t" if c26 >= c25 else "neg-t"}">{eur(c26 - c25, signo=True)}</b> '
                   f'vs {mes_n} {ant} al mismo día')
    items = []
    for i in range(1, fecha.day + 1):
        dd = fecha.replace(day=i)
        a, b = venta_dia(dd) or {}, venta_dia(mismo_dia_semana(dd)) or {}
        k = "k" if i in (1, 5, 10, 15, 20, 25, 30) or dd == fecha else ""
        we = "we" if dd.weekday() >= 5 else ""
        items.append((str(i), sum(a.get(x, 0) for x in COMP), sum(b.get(x, 0) for x in COMP), f"{we} {k}"))
    futuros = [(str(i), "k" if i in (15, 20, 25, 30) else "") for i in range(fecha.day + 1, ndias + 1)]

    bloque_mes = f'''
 <div class="sec">
  <table class="mh" role="presentation"><tr>
    <td><div class="lbl">Acumulado de {mes_n}</div><div class="big2">{eur(tot)}</div></td>
    <td class="mhr">{chip(var_m, "lg")}</td></tr></table>
  <div class="msub">{sub_mes}</div>
  {barra_partida(c26, nuevos, c25, f"{MES_C[fecha.month]}/{str(fecha.year)[2:]}", f"{MES_C[fecha.month]}/{str(ant)[2:]}", NUE)}
  <table class="dhead" role="presentation"><tr><td class="lbl0">Día por día contra {ant}</td><td class="dhr">Día {fecha.day} de {ndias}</td></tr></table>
  {ritmo(items, 46, futuros)}
  <div class="legend">Cada día contra el mismo día de la semana de {ant}. En negrita, los fines de semana.</div>
 </div>'''

    # ---- 4. variacion por sucursal
    var_b = {b: R[b]["pct"] for b in COMP}
    dif_b = {b: R[b]["diff"] for b in COMP}
    bloque_var = f'''
 <div class="sec">
  <div class="h" style="margin-bottom:12px">Variación del mes por sucursal</div>
  {divergente(var_b, NUE, fecha.year)}
  {destacados(dif_b, c26, c25, NUE, nuevos, "el mes", ant)}
 </div>'''

    # ---- 5. tabla
    def fila(b):
        r, t = R[b], tickets.get(b, 0)
        if r.get("cerrado"):
            linea, ayer = "cerrado por descanso", '<span class="mut it2">cerrado</span>'
        else:
            linea = (f'{num(t)} tickets · <span class="nw">prom. {eur(r["dia"] / t, 2)}</span>'
                     if t else "sin tickets")
            ayer = eur(r["dia"])
        if r["comparable"]:
            sub25 = f'<div class="mut sm">vs {eur(r["a25"])}</div>'
            var = (f'{chip(r["pct"])}<div class="sm b {"pos-t" if r["diff"] >= 0 else "neg-t"}">'
                   f'{eur(r["diff"], signo=True)}</div>')
        else:
            sub25, var = f'<div class="mut sm">sin {ant}</div>', '<span class="chip flat">nuevo</span>'
        return (f'<tr><td class="c-name"><div class="nm">{b}</div><div class="mut sm">{linea}</div></td>'
                f'<td class="c-num">{ayer}</td><td class="c-acum"><div class="b">{eur(r["a26"])}</div>{sub25}</td>'
                f'<td class="c-var">{var}</td></tr>')

    def subtotal(nombre, s, locs):
        st = sum(tickets.get(b, 0) for b in locs)
        ast = "*" if s.get("hay_nuevo") else ""
        return (f'<tr><td class="c-name sub subl"><div class="nm">{nombre}</div>'
                f'<div class="mut sm">{f"prom. {eur(s["dia"] / st, 2)}" if st else ""}</div></td>'
                f'<td class="c-num sub">{eur(s["dia"])}</td>'
                f'<td class="c-acum sub"><div class="b">{eur(s["a26"])}</div><div class="mut sm">vs {eur(s["ca25"])}{ast}</div></td>'
                f'<td class="c-var sub subr">{chip(s["pct"])}<div class="sm b {"pos-t" if s["diff"] >= 0 else "neg-t"}">'
                f'{eur(s["diff"], signo=True)}</div></td></tr>')

    grupo = lambda t: f'<tr><td colspan="4" class="grp">{t}</td></tr>'
    filas = (grupo("Propias") + "".join(fila(b) for b in PROPIAS) + subtotal("Subtotal propias", propias, PROPIAS) +
             grupo("Franquicias") + "".join(fila(b) for b in FRANQUICIAS) +
             subtotal("Subtotal franquicias", franquicias, FRANQUICIAS))
    ast = "*" if totals.get("hay_nuevo") else ""
    bloque_tabla = f'''
 <div class="sec">
  <div class="h" style="margin-bottom:12px">Detalle por sucursal</div>
  <table class="tbl" role="presentation">
   <tr><th class="th thl">Sucursal</th><th class="th c-num">Ayer</th><th class="th c-acum">Acum. {MES_C[fecha.month].lower()}</th><th class="th c-var">vs {ant}</th></tr>
   {filas}
  </table>
  <table class="tot" role="presentation"><tr>
   <td class="tt"><div>Total general</div><div class="ts">{f"prom. {eur(dtot / tkt, 2)}" if tkt else ""}</div></td>
   <td class="tt r c-num">{eur(dtot)}</td>
   <td class="tt r"><div>{eur(tot)}</div><div class="ts">vs {eur(c25)}{ast}</div></td>
   <td class="tt r tv">{chip(var_m, "inv")}<div class="ts {"pos-tot" if c26 >= c25 else "neg-tot"}">{eur(c26 - c25, signo=True)}</div></td>
  </tr></table>
  {_notas(fecha, rows, NUE)}
 </div>'''

    card = tarjeta(header("Reporte diario · Europa",
                          f"{DIAS[fecha.weekday()]} {fecha.day} de {mes_n} de {fecha.year}"),
                   bloque_dia, bloque_mes, bloque_var, bloque_tabla,
                   pie=pie(f"Lucciano's Europa · Ventas netas sin IVA · Datos al {fecha.strftime('%d/%m/%Y')}"))
    pre = f"Ayer {eur(dtot)} · {mes_n.capitalize()} {pct(var_m)} vs {ant}" + (" en locales con historia" if NUE else "")
    return documento(card, f"Reporte diario {fecha.strftime('%d/%m/%Y')}", pre)


def _notas(fecha, rows, NUE):
    notas = []
    if NUE:
        notas.append(f"* La variación y el año anterior se calculan solo sobre locales con historia "
                     f"{fecha.year - 1}; {lista(NUE)} suman al total pero no a la variación.")
    fijos = [r["branch"] for r in rows if r.get("cerrado") and motivo_cierre(r["branch"], fecha) == "fijo"]
    punt = [r["branch"] for r in rows if r.get("cerrado") and motivo_cierre(r["branch"], fecha) == "puntual"]
    if fijos:
        notas.append(f"{lista(fijos)} {'cierra' if len(fijos) == 1 else 'cierran'} los "
                     f"{DIAS[fecha.weekday()].lower()} (descanso fijo): hoy no tiene venta y no es un faltante.")
    if punt:
        notas.append(f"{lista(punt)} {'no abrió' if len(punt) == 1 else 'no abrieron'} hoy "
                     f"(descanso corrido de día por un feriado): no es un faltante.")
    return "".join(f'<div class="nota">{n}</div>' for n in notas)
