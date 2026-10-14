"""
mail_cierre.py — HTML del CIERRE MENSUAL de Europa (diseño "combinado").

Lo llama report_cierre.py con los numeros ya calculados (construir() + YTD).
Solo dibuja. Estilo, componentes e inliner: mail_base.py.

Estructura:
  1. Header
  2. "Venta del mes": total, tickets y ticket promedio (vs anio anterior a igual
     base), vs mes anterior, vs mismo mes del anio anterior
  3. Aviso de calendario (si el mes tuvo distinta composicion de dias)
  4. Variacion del mes por sucursal (divergente) + destacados
  5. Tabla por sucursal (con participacion y ticket promedio)
  6. Acumulado del anio: total, barra partida, venta mes a mes, por sucursal
"""
from report import BRANCH_ORDER, PROPIAS, FRANQUICIAS, SIN_HISTORICO_2025
from mail_base import (MESES, MES_C, eur, pct, num, chip, flecha_pct, lista, header, celda,
                       celda_var, hero, barra_partida, columnas, divergente, destacados, callout,
                       tarjeta, pie, documento)


def render(dia1, ultimo, rows, totals, propias, franquicias, ia, fa, con_tks, ytd,
           texto_cal=None, serie_meses=None, mes_prev=None):
    """
    Mismos argumentos que el render_html viejo de report_cierre.py, mas:
      texto_cal:   nota de composicion del calendario (report_cierre.composicion) o None
      serie_meses: [(etiqueta, venta_anio, venta_anio_anterior)] de enero al mes cerrado
      mes_prev:    (etiqueta, venta total del mes anterior) o None
    """
    R = {r["branch"]: r for r in rows}
    NUE = [b for b in BRANCH_ORDER if b in SIN_HISTORICO_2025]
    COMP = [b for b in BRANCH_ORDER if b not in SIN_HISTORICO_2025]
    anio, mes = ultimo.year, ultimo.month
    ant = anio - 1
    mes_n = MESES[mes]
    ndias = (ultimo - dia1).days + 1

    tot, c26, c25 = totals["a26"], totals["ca26"], totals["ca25"]
    nuevos = tot - c26
    var_m = totals["pct"]

    # Tickets y ticket promedio a igual base
    var_tk = var_tp = None
    tks = tp = 0
    if con_tks:
        tks = totals.get("tks26", 0)
        tp = tot / tks if tks else 0
        ct26 = sum(R[b].get("tks26", 0) for b in COMP)
        ct25 = sum(R[b].get("tks25", 0) for b in COMP)
        if ct25:
            var_tk = (ct26 / ct25 - 1) * 100
            var_tp = ((c26 / ct26) / (c25 / ct25) - 1) * 100 if ct26 else None

    if mes_prev and mes_prev[1]:
        var_prev = (tot / mes_prev[1] - 1) * 100
        c3 = celda_var(f"vs {mes_prev[0]}", var_prev, eur(mes_prev[1]))
    else:
        c3 = celda_var("vs mes anterior", None, "")

    bloque_mes = hero(
        f"Venta de {mes_n} {anio}", tot,
        f"{ndias} días · promedio {eur(tot / ndias)} por día",
        [celda("Tickets", num(tks) if con_tks else "—", f"{flecha_pct(var_tk)} vs {ant}" if var_tk is not None else ""),
         celda("Ticket prom.", eur(tp, 2) if (con_tks and tks) else "—",
               f"{flecha_pct(var_tp)} vs {ant}" if var_tp is not None else ""),
         c3,
         celda_var(f"vs {MES_C[mes].lower()} {ant}", var_m, "a igual base")],
        propias["a26"], franquicias["a26"])

    # ---- 3+4. calendario + variacion por sucursal
    var_b = {b: R[b]["pct"] for b in COMP}
    dif_b = {b: R[b]["diff"] for b in COMP}
    bloque_var = f'''
 <div class="sec">
  <div class="h" style="margin-bottom:12px">Variación del mes por sucursal</div>
  {divergente(var_b, NUE, anio)}
  {destacados(dif_b, c26, c25, NUE, nuevos, "el mes", ant)}
  {callout("<b>Ojo con el calendario.</b> " + texto_cal) if texto_cal else ""}
 </div>'''

    # ---- 5. tabla
    def fila(b):
        r = R[b]
        partes = [f'{r.get("share", 0):.1f}'.replace(".", ",") + "% del total"]
        if con_tks and r.get("tp26"):
            partes.append(f'<span class="nw">prom. {eur(r["tp26"], 2)}</span>')
        if r["comparable"]:
            sub25 = f'<div class="mut sm">vs {eur(r["a25"])}</div>'
            var = (f'{chip(r["pct"])}<div class="sm b {"pos-t" if r["diff"] >= 0 else "neg-t"}">'
                   f'{eur(r["diff"], signo=True)}</div>')
        else:
            sub25, var = f'<div class="mut sm">sin {ant}</div>', '<span class="chip flat">nuevo</span>'
        return (f'<tr><td class="c-name"><div class="nm">{b}</div><div class="mut sm">{" · ".join(partes)}</div></td>'
                f'<td class="c-acum"><div class="b">{eur(r["a26"])}</div>{sub25}</td>'
                f'<td class="c-var">{var}</td></tr>')

    def subtotal(nombre, s):
        ast = "*" if s.get("hay_nuevo") else ""
        part = f'{s["a26"] / tot * 100:.1f}'.replace(".", ",") + "% del total" if tot else ""
        return (f'<tr><td class="c-name sub subl"><div class="nm">{nombre}</div><div class="mut sm">{part}</div></td>'
                f'<td class="c-acum sub"><div class="b">{eur(s["a26"])}</div><div class="mut sm">vs {eur(s["ca25"])}{ast}</div></td>'
                f'<td class="c-var sub subr">{chip(s["pct"])}<div class="sm b {"pos-t" if s["diff"] >= 0 else "neg-t"}">'
                f'{eur(s["diff"], signo=True)}</div></td></tr>')

    grupo = lambda t: f'<tr><td colspan="3" class="grp">{t}</td></tr>'
    filas = (grupo("Propias") + "".join(fila(b) for b in PROPIAS) + subtotal("Subtotal propias", propias) +
             grupo("Franquicias") + "".join(fila(b) for b in FRANQUICIAS) + subtotal("Subtotal franquicias", franquicias))
    ast = "*" if totals.get("hay_nuevo") else ""
    notas = [f"Mes cerrado del {dia1.strftime('%d/%m')} al {ultimo.strftime('%d/%m/%Y')}, comparado contra "
             f"{ia.strftime('%d/%m')} al {fa.strftime('%d/%m/%Y')} (mismas fechas del año anterior)."]
    if NUE:
        notas.append(f"* La variación y el año anterior se calculan solo sobre locales con historia {ant}; "
                     f"{lista(NUE)} suman al total pero no a la variación.")
    bloque_tabla = f'''
 <div class="sec">
  <div class="h" style="margin-bottom:12px">Detalle por sucursal</div>
  <table class="tbl" role="presentation">
   <tr><th class="th thl">Sucursal</th><th class="th c-acum">{MES_C[mes]}/{str(anio)[2:]}</th><th class="th c-var">vs {ant}</th></tr>
   {filas}
  </table>
  <table class="tot" role="presentation"><tr>
   <td class="tt"><div>Total {MES_C[mes].lower()}/{str(anio)[2:]}</div><div class="ts">{f"prom. {eur(tp, 2)}" if (con_tks and tks) else ""}</div></td>
   <td class="tt r"><div>{eur(tot)}</div><div class="ts">vs {eur(c25)}{ast}</div></td>
   <td class="tt r tv">{chip(var_m, "inv")}<div class="ts {"pos-tot" if c26 >= c25 else "neg-tot"}">{eur(c26 - c25, signo=True)}</div></td>
  </tr></table>
  {"".join(f'<div class="nota">{n}</div>' for n in notas)}
 </div>'''

    # ---- 6. acumulado del anio
    bloque_anio = ""
    if ytd:
        p26, p25 = ytd["por26"], ytd["por25"]
        y26, y25 = ytd["t26"], ytd["t25"]
        yc26 = sum(p26.get(b, 0) for b in COMP)
        yc25 = sum(p25.get(b, 0) for b in COMP)
        ynue = y26 - yc26
        rango = f"enero a {mes_n}" if mes > 1 else "enero"
        sub_y = (f'<b class="{"pos-t" if yc26 >= yc25 else "neg-t"}">{eur(yc26 - yc25, signo=True)}</b> '
                 f'vs {rango} {ant}' + (f' en los {len(COMP)} locales con historia' if NUE else '') +
                 (f' · {num(ytd["tks"])} tickets, prom. {eur(ytd["tp"], 2)}' if ytd.get("tks") else ""))
        cols = ""
        if serie_meses:
            cols = f'''
  <table class="dhead" role="presentation"><tr><td class="lbl0">Venta mes a mes</td><td class="dhr">{anio} contra {ant}</td></tr></table>
  {columnas([(et, v26, v25, et == MES_C[mes]) for et, v26, v25 in serie_meses], 96, ocultar_valores_celular=len(serie_meses) > 6)}
  <div class="legend"><span class="it"><span class="dot c26"></span>{anio}</span><span class="it"><span class="dot c25"></span>{ant}</span></div>'''

        filas_y = ""
        for b in sorted(BRANCH_ORDER, key=lambda x: p26.get(x, 0), reverse=True):
            v = p26.get(b, 0)
            share = f"{v / y26 * 100:.1f}".replace(".", ",") + "% del total" if y26 else ""
            if b in SIN_HISTORICO_2025 or not p25.get(b):
                var = '<span class="chip flat">nuevo</span>'
                sub25 = f'<div class="mut sm">sin {ant}</div>'
            else:
                d = v - p25[b]
                var = (f'{chip((v / p25[b] - 1) * 100)}<div class="sm b {"pos-t" if d >= 0 else "neg-t"}">'
                       f'{eur(d, signo=True)}</div>')
                sub25 = f'<div class="mut sm">vs {eur(p25[b])}</div>'
            filas_y += (f'<tr><td class="c-name"><div class="nm">{b}</div><div class="mut sm">{share}</div></td>'
                        f'<td class="c-acum"><div class="b">{eur(v)}</div>{sub25}</td><td class="c-var">{var}</td></tr>')

        bloque_anio = f'''
 <div class="sec">
  <table class="mh" role="presentation"><tr>
    <td><div class="lbl">Acumulado del año · {rango} {anio}</div><div class="big2">{eur(y26)}</div></td>
    <td class="mhr">{chip(ytd["pct"], "lg")}</td></tr></table>
  <div class="msub">{sub_y}</div>
  {barra_partida(yc26, ynue, yc25, str(anio), str(ant), NUE)}
  {cols}
  <table class="dhead" role="presentation"><tr><td class="lbl0">Por sucursal en el año</td><td class="dhr"></td></tr></table>
  <table class="tbl" role="presentation">
   <tr><th class="th thl">Sucursal</th><th class="th c-acum">Acum. {anio}</th><th class="th c-var">vs {ant}</th></tr>
   {filas_y}
  </table>
  <div class="nota">Acumulado del año calendario a la misma fecha en los dos años: lo que va de {anio} contra lo que iba de {ant} a esta altura.</div>
 </div>'''

    card = tarjeta(header("Cierre mensual · Europa", f"{mes_n.capitalize()} {anio}"),
                   bloque_mes, bloque_var, bloque_tabla, bloque_anio,
                   pie=pie(f"Lucciano's Europa · Ventas netas sin IVA · {mes_n.capitalize()} {anio}"))
    pre = (f"{mes_n.capitalize()} {eur(tot)} · {pct(var_m)} vs {ant}" + (" a igual base" if NUE else "") +
           (f" · Año {pct(ytd['pct'])}" if ytd else ""))
    return documento(card, f"Cierre mensual {mes_n} {anio}", pre)
