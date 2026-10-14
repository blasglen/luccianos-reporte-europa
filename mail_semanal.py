"""
mail_semanal.py — HTML del reporte SEMANAL de Europa (diseño "combinado").

Lo llama report_semanal.py con los numeros ya calculados (construir()). Solo
dibuja. Estilo, componentes e inliner: mail_base.py.

Estructura:
  1. Header
  2. "Venta de la semana": total, tickets y ticket promedio (vs anio anterior a
     igual base), vs semana anterior, vs misma semana del anio anterior
  3. Dia por dia: 7 columnas, cada dia contra el mismo dia de la semana de 2025
  4. Variacion de la semana por sucursal (divergente) + destacados
  5. Acumulado del mes (barra partida con historia / nuevos)
  6. Tabla por sucursal: semana, variacion de la semana, acumulado del mes
  7. Ultimos 12 meses cerrados, por sucursal
"""
from datetime import timedelta

from report import BRANCH_ORDER, PROPIAS, FRANQUICIAS, SIN_HISTORICO_2025, cerrado_por_descanso
from mail_base import (DIAS_C, MESES, MES_C, eur, pct, num, chip, flecha_pct, lista,
                       venta_dia, mismo_dia_semana, header, celda, celda_var, hero,
                       barra_partida, columnas, divergente, destacados, barras_horizontales,
                       tarjeta, pie, documento, historico)


def _rango_txt(desde, hasta):
    if desde.month == hasta.month:
        return f"{desde.day} al {hasta.day} de {MESES[hasta.month]} de {hasta.year}"
    return f"{desde.day} de {MESES[desde.month]} al {hasta.day} de {MESES[hasta.month]} de {hasta.year}"


def render(desde, hasta, rows, totals, propias, franquicias, serie, mejor, peor, dia1, sin_tks, v12,
           v12_hasta=None):
    """Mismos argumentos que el render_html viejo de report_semanal.py.
    v12_hasta: (anio, mes) del ultimo mes cerrado que suma v12 (para el titulo)."""
    R = {r["branch"]: r for r in rows}
    NUE = [b for b in BRANCH_ORDER if b in SIN_HISTORICO_2025]
    COMP = [b for b in BRANCH_ORDER if b not in SIN_HISTORICO_2025]
    ant = hasta.year - 1

    s26 = totals["sem26"]
    cs26, cs25 = totals["csem26"], totals["csem25"]
    var_sem = totals["sem_pct"]

    # Tickets y ticket promedio a igual base (solo locales con historia)
    ct26 = sum(R[b]["tks26"] for b in COMP)
    ct25 = sum(R[b]["tks25"] for b in COMP)
    var_tk = (ct26 / ct25 - 1) * 100 if (ct25 and not sin_tks) else None
    tp_c26 = cs26 / ct26 if ct26 else 0
    tp_c25 = cs25 / ct25 if ct25 else 0
    var_tp = (tp_c26 / tp_c25 - 1) * 100 if (tp_c25 and not sin_tks) else None

    # Semana anterior (todos los locales)
    prev = [venta_dia(desde - timedelta(days=7 - i)) for i in range(7)]
    sprev = sum(sum(p.values()) for p in prev) if all(prev) else None
    var_prev = (s26 / sprev - 1) * 100 if sprev else None
    pd, ph = desde - timedelta(days=7), hasta - timedelta(days=7)

    tks = totals["tks26"]
    tp = s26 / tks if tks else 0
    sub_tk = f"{flecha_pct(var_tk)} vs {ant}" if var_tk is not None else ""
    sub_tp = f"{flecha_pct(var_tp)} vs {ant}" if var_tp is not None else ""

    bloque_sem = hero(
        f"Venta de la semana · {DIAS_C[desde.weekday()]} {desde.day}/{desde.month} al "
        f"{DIAS_C[hasta.weekday()]} {hasta.day}/{hasta.month}",
        s26, f"Promedio {eur(s26 / 7)} por día · {len(rows)} sucursales",
        [celda("Tickets", num(tks) if not sin_tks else "—", sub_tk),
         celda("Ticket prom.", eur(tp, 2) if (tks and not sin_tks) else "—", sub_tp),
         celda_var(f"vs {pd.day}/{pd.month} al {ph.day}/{ph.month}", var_prev, eur(sprev) if sprev else ""),
         celda_var(f"vs {ant}", var_sem, "misma semana, igual base")],
        propias["sem26"], franquicias["sem26"])

    # ---- 3. dia por dia
    items = []
    for s in serie:
        f = s["fecha"]
        v25 = venta_dia(mismo_dia_semana(f))
        items.append((f"{DIAS_C[f.weekday()]} {f.day}", s["actual"], sum(v25.values()) if v25 else None,
                      s["fecha"] == mejor["fecha"]))
    prom = s26 / 7
    cerr = {}
    for s in serie:
        for b in BRANCH_ORDER:
            h = historico(s["fecha"].year).get(s["fecha"].isoformat(), {})
            v = h.get(b)
            vta = v.get("venta", 0) if isinstance(v, dict) else (v or 0)
            if cerrado_por_descanso(b, s["fecha"]) and not vta:
                cerr.setdefault(b, []).append(f"{DIAS_C[s['fecha'].weekday()]} {s['fecha'].day}")
    nota_cerr = "".join(f" {b} no abrió el {' y el '.join(ds)} (descanso)." for b, ds in cerr.items())
    nota_dias = []
    if NUE:
        nota_dias.append(f"El {ant} no incluye a {lista(NUE)}, que todavía no existían.")
    if nota_cerr:
        nota_dias.append(nota_cerr.strip())
    bloque_dias = f'''
 <div class="sec">
  <div class="h">Día por día</div>
  <div class="hsubt">Mejor día: <b>{DIAS_C[mejor["fecha"].weekday()]} {mejor["fecha"].day}</b> con {eur(mejor["actual"])} · Más flojo: <b>{DIAS_C[peor["fecha"].weekday()]} {peor["fecha"].day}</b> con {eur(peor["actual"])} · Promedio {eur(prom)}</div>
  {columnas(items, 96)}
  <div class="legend"><span class="it"><span class="dot c26"></span>{hasta.year}</span><span class="it"><span class="dot c25"></span>Mismo día de la semana de {ant}</span></div>
  {"".join(f'<div class="nota">{n}</div>' for n in nota_dias)}
 </div>'''

    # ---- 4. variacion de la semana
    var_b = {b: R[b]["sem_pct"] for b in COMP}
    dif_b = {b: R[b]["sem_diff"] for b in COMP}
    nuevos_sem = s26 - cs26
    bloque_var = f'''
 <div class="sec">
  <div class="h" style="margin-bottom:12px">Variación de la semana por sucursal</div>
  {divergente(var_b, NUE, hasta.year)}
  {destacados(dif_b, cs26, cs25, NUE, nuevos_sem, "la semana", ant)}
 </div>'''

    # ---- 5. acumulado del mes
    a26, ca26, ca25 = totals["a26"], totals["ca26"], totals["ca25"]
    mes_n = MESES[hasta.month]
    sub_mes = (f'<b class="{"pos-t" if ca26 >= ca25 else "neg-t"}">{eur(ca26 - ca25, signo=True)}</b> '
               f'vs {mes_n} {ant}' + (f' en los {len(COMP)} locales con historia' if NUE else '') +
               f', del {dia1.day} al {hasta.day}')
    bloque_mes = f'''
 <div class="sec">
  <table class="mh" role="presentation"><tr>
    <td><div class="lbl">Acumulado de {mes_n}</div><div class="big2">{eur(a26)}</div></td>
    <td class="mhr">{chip(totals["pct"], "lg")}</td></tr></table>
  <div class="msub">{sub_mes}</div>
  {barra_partida(ca26, a26 - ca26, ca25, f"{MES_C[hasta.month]}/{str(hasta.year)[2:]}", f"{MES_C[hasta.month]}/{str(ant)[2:]}", NUE)}
 </div>'''

    # ---- 6. tabla
    def fila(b):
        r = R[b]
        t = r["tks26"]
        linea = (f'{num(t)} tickets · <span class="nw">prom. {eur(r["sem26"] / t, 2)}</span>'
                 if (t and not sin_tks) else "")
        if r["comparable"]:
            sub25 = f'<div class="mut sm">vs {eur(r["sem25"])}</div>'
            var = (f'{chip(r["sem_pct"])}<div class="sm b {"pos-t" if r["sem_diff"] >= 0 else "neg-t"}">'
                   f'{eur(r["sem_diff"], signo=True)}</div>')
            mes_v = f'<div class="sm b {"pos-t" if r["diff"] >= 0 else "neg-t"}">{pct(r["pct"])}</div>'
        else:
            sub25, var = f'<div class="mut sm">sin {ant}</div>', '<span class="chip flat">nuevo</span>'
            mes_v = '<div class="mut sm">nuevo</div>'
        return (f'<tr><td class="c-name"><div class="nm">{b}</div><div class="mut sm">{linea}</div></td>'
                f'<td class="c-acum"><div class="b">{eur(r["sem26"])}</div>{sub25}</td>'
                f'<td class="c-var">{var}</td>'
                f'<td class="c-num"><div>{eur(r["a26"])}</div>{mes_v}</td></tr>')

    def subtotal(nombre, s):
        ast = "*" if s.get("hay_nuevo") else ""
        prom_g = f"prom. {eur(s['tp26'], 2)}" if (s.get("tks26") and not sin_tks) else ""
        return (f'<tr><td class="c-name sub subl"><div class="nm">{nombre}</div><div class="mut sm">{prom_g}</div></td>'
                f'<td class="c-acum sub"><div class="b">{eur(s["sem26"])}</div><div class="mut sm">vs {eur(s["csem25"])}{ast}</div></td>'
                f'<td class="c-var sub">{chip(s["sem_pct"])}<div class="sm b {"pos-t" if s["sem_diff"] >= 0 else "neg-t"}">'
                f'{eur(s["sem_diff"], signo=True)}</div></td>'
                f'<td class="c-num sub subr"><div>{eur(s["a26"])}</div>'
                f'<div class="sm b {"pos-t" if s["diff"] >= 0 else "neg-t"}">{pct(s["pct"])}{ast}</div></td></tr>')

    grupo = lambda t: f'<tr><td colspan="4" class="grp">{t}</td></tr>'
    filas = (grupo("Propias") + "".join(fila(b) for b in PROPIAS) + subtotal("Subtotal propias", propias) +
             grupo("Franquicias") + "".join(fila(b) for b in FRANQUICIAS) + subtotal("Subtotal franquicias", franquicias))
    ast = "*" if totals.get("hay_nuevo") else ""
    bloque_tabla = f'''
 <div class="sec">
  <div class="h" style="margin-bottom:12px">Detalle por sucursal</div>
  <table class="tbl" role="presentation">
   <tr><th class="th thl">Sucursal</th><th class="th c-acum">Semana</th><th class="th c-var">vs {ant}</th><th class="th c-num">Acum. {MES_C[hasta.month].lower()}</th></tr>
   {filas}
  </table>
  <table class="tot" role="presentation"><tr>
   <td class="tt"><div>Total general</div><div class="ts">{f"prom. {eur(tp, 2)}" if (tks and not sin_tks) else ""}</div></td>
   <td class="tt r"><div>{eur(s26)}</div><div class="ts">vs {eur(cs25)}{ast}</div></td>
   <td class="tt r tv">{chip(var_sem, "inv")}<div class="ts {"pos-tot" if cs26 >= cs25 else "neg-tot"}">{eur(cs26 - cs25, signo=True)}</div></td>
   <td class="tt r c-num"><div>{eur(a26)}</div><div class="ts {"pos-tot" if ca26 >= ca25 else "neg-tot"}">{pct(totals["pct"])}{ast}</div></td>
  </tr></table>
 </div>'''

    # ---- 7. ultimos 12 meses
    bloque_12 = ""
    if v12 and sum(v12.values()) > 0:
        tit = ""
        if v12_hasta:
            ay, am = v12_hasta
            d_ini_m = am % 12 + 1
            d_ini_y = ay if am == 12 else ay - 1
            tit = f"{MES_C[d_ini_m].lower()} {d_ini_y} a {MES_C[am].lower()} {ay}"
        bloque_12 = f'''
 <div class="sec">
  <table class="mh" role="presentation"><tr>
    <td><div class="lbl">Últimos 12 meses cerrados{f" · {tit}" if tit else ""}</div><div class="big2">{eur(sum(v12.values()))}</div></td></tr></table>
  <div class="msub">Venta de cada sucursal en el último año móvil y su peso sobre el total. No se compara contra nada: es la escala de cada local.</div>
  {barras_horizontales(v12)}
 </div>'''

    # ---- notas
    notas = [f"Semana del {desde.strftime('%d/%m')} al {hasta.strftime('%d/%m/%Y')}, comparada contra las mismas "
             f"fechas de {ant}: cualquier semana de 7 días tiene un día de cada tipo, así que la comparación es pareja."]
    if NUE:
        notas.append(f"* La variación y el año anterior se calculan solo sobre locales con historia {ant}; "
                     f"{lista(NUE)} suman al total pero no a la variación.")
    if cerr:
        notas.append(f"{lista(cerr)}: los días de descanso no tienen venta y explican parte de la "
                     f"variación del local contra {ant}.")
    if sin_tks:
        notas.append("Hay días sin tickets en el historial: el reporte sale con la venta pero sin ticket promedio.")
    # las notas van al final del bloque de la tabla (antes de cerrar su <div>)
    cierre_div = "\n </div>"
    bloque_tabla = (bloque_tabla[:-len(cierre_div)] +
                    "".join(f'<div class="nota">{n}</div>' for n in notas) + cierre_div)

    card = tarjeta(header("Reporte semanal · Europa", _rango_txt(desde, hasta)),
                   bloque_sem, bloque_dias, bloque_var, bloque_mes, bloque_tabla, bloque_12,
                   pie=pie(f"Lucciano's Europa · Ventas netas sin IVA · Semana del {desde.strftime('%d/%m')} "
                           f"al {hasta.strftime('%d/%m/%Y')}"))
    pre = (f"Semana {eur(s26)} · {pct(var_sem)} vs {ant}" + (" a igual base" if NUE else "") +
           f" · {MESES[hasta.month].capitalize()} {pct(totals['pct'])}")
    return documento(card, f"Reporte semanal {desde.strftime('%d/%m')} al {hasta.strftime('%d/%m/%Y')}", pre)
