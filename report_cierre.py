"""
Lucciano's - Reporte de CIERRE MENSUAL
--------------------------------------
Se manda el DIA 1 de cada mes y cubre el mes que acaba de cerrar, completo.

POR QUE EL DIA 1 Y NO OTRO:
  report.py mira la fecha del EXCEL, no la de hoy. Entonces:
    - Dia 1, 7:00 -> procesa el ultimo dia del mes anterior -> mismo mes -> SUMA.
      acumulado.json queda con el mes CERRADO y COMPLETO.
    - Dia 2, 7:00 -> procesa el dia 1 del mes nuevo -> cambio de mes -> REINICIA.
      El mes anterior se pierde de acumulado.json.
  O sea que el mes cerrado existe en una ventana de 24 horas. Este reporte corre
  adentro de esa ventana (dia 1, 8:00) y ademas lo GUARDA en data/cierres.json
  antes de que el dia 2 lo pise. Ese archivo es el que, a fin de anio, te da la
  serie mes a mes de 2026 sin haber hecho nada extra.

El snapshot ES el candado: si el mes ya esta en cierres.json, no se reenvia.
"""
import argparse
import json
import os
import sys
from calendar import monthrange
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from report import BRANCH_ORDER, PROPIAS, FRANQUICIAS, MESES_ES, MES_CORTO, money, money2, parse_excel_full, SIN_HISTORICO_2025
from generar_acum_ant import acumular_rango, escribir_excel, espejo
import mensual

BASE = Path(__file__).parent
SALIDA_ANT = BASE / "Acumulado_cierre_ant.xlsx"
ESTADO_DIARIO = BASE / "data" / "acumulado.json"
CIERRES = BASE / "data" / "cierres.json"
PREVIEW = BASE / "preview_cierre.html"


def mes_a_cerrar(hoy=None):
    """El mes que cerro. Corriendo el dia 1, es el mes anterior.

    Si corre otro dia (prueba manual, o se atraso), igual devuelve el mes anterior
    al de hoy. La validacion contra acumulado.json despues avisa si no da.
    """
    hoy = hoy or date.today()
    primero = date(hoy.year, hoy.month, 1)
    ultimo_ant = primero - timedelta(days=1)
    return ultimo_ant.year, ultimo_ant.month


def rango_mes(anio, mes):
    return date(anio, mes, 1), date(anio, mes, monthrange(anio, mes)[1])


def leer_acum_cerrado(dia1, ultimo):
    """Trae el mes cerrado de acumulado.json, con validacion dura.

    La ventana es de 24 horas. Si este reporte corre tarde (dia 2 o despues), el
    diario ya reinicio y el mes se perdio: acumulado.json va a decir que es del
    mes nuevo y esto corta. Es el escenario que hay que mirar si falla.
    """
    if not ESTADO_DIARIO.exists():
        raise SystemExit(f"[ERROR] No existe {ESTADO_DIARIO.name}.")
    st = json.loads(ESTADO_DIARIO.read_text(encoding="utf-8"))

    mes_esperado = dia1.strftime("%Y-%m")
    if st.get("month") != mes_esperado:
        raise SystemExit(
            f"[ERROR] El acumulado del diario dice ser del mes {st.get('month')} y "
            f"se esperaba {mes_esperado}.\n"
            f"        Si dice el mes SIGUIENTE, el diario ya reinicio y el cierre de "
            f"{mes_esperado} se perdio de acumulado.json. Se puede recuperar del "
            f"historial de commits del repo (buscar el commit del dia 1).\n"
            f"        Este reporte tiene que correr el DIA 1, despues del diario."
        )
    if st.get("last_date") != ultimo.isoformat():
        raise SystemExit(
            f"[ERROR] El acumulado del diario esta cerrado al {st.get('last_date')} y "
            f"el mes termina el {ultimo.isoformat()}.\n"
            f"        Causa tipica: el diario del dia 1 (que procesa el ultimo dia del "
            f"mes) no corrio o fallo. Arreglar eso y volver a correr este.\n"
            f"        NO se manda: faltaria el ultimo dia del mes."
        )

    acum = st.get("acumulado", {})
    faltan = [b for b in BRANCH_ORDER if b not in acum]
    if faltan:
        raise SystemExit(f"[ERROR] El acumulado del diario no tiene: {', '.join(faltan)}")
    return {b: round(float(acum[b]), 2) for b in BRANCH_ORDER}


DIAS_ES_C = {0: "lunes", 1: "martes", 2: "miércoles", 3: "jueves", 4: "viernes",
             5: "sábado", 6: "domingo"}


def leer_tickets_mes(dia1, ultimo):
    """Tickets del mes por sucursal, desde el historial.

    acumulado.json solo arrastra plata, no tickets, asi que esto sale del
    historial. Si al historial le falta algun dia del mes, devuelve None y el
    reporte sale sin ticket promedio (pero con la venta, que es lo importante).
    No invento un promedio con datos parciales: seria peor que no mostrarlo.
    """
    path = BASE / "data" / f"historico_{ultimo.year}.json"
    if not path.exists():
        return None
    hist = json.loads(path.read_text(encoding="utf-8"))
    tks = {b: 0 for b in BRANCH_ORDER}
    d = dia1
    while d <= ultimo:
        v = hist.get(d.isoformat())
        if not v:
            return None
        for b in BRANCH_ORDER:
            x = v.get(b)
            if not isinstance(x, dict) or "tickets" not in x:
                return None   # historial viejo, sin tickets
            tks[b] += int(x["tickets"])
        d += timedelta(days=1)
    return tks


def composicion(dia1, ultimo):
    """Cuenta los dias de semana del mes y del mismo mes del anio anterior.

    Por que importa: julio 2026 tiene 5 viernes y 4 martes; julio 2025 tenia 4
    viernes y 5 martes. En una heladeria un viernes vale bastante mas que un
    martes, asi que parte de la variacion interanual es CALENDARIO, no
    performance. Un cierre que no lo aclara induce a error.

    (En el semanal esto no hace falta: cualquier ventana de 7 dias tiene
    exactamente un dia de cada tipo, asi que la composicion siempre coincide.)

    Devuelve (texto_o_None, dias_mes, dias_mes_ant).
    """
    def contar(y, m):
        c = Counter()
        d = date(y, m, 1)
        while d.month == m:
            c[d.weekday()] += 1
            d += timedelta(days=1)
        return c

    act = contar(ultimo.year, ultimo.month)
    ant = contar(ultimo.year - 1, ultimo.month)
    de_mas = [DIAS_ES_C[k] for k in range(7) if act[k] > ant[k]]
    de_menos = [DIAS_ES_C[k] for k in range(7) if act[k] < ant[k]]
    n_act = sum(act.values())
    n_ant = sum(ant.values())

    partes = []
    if de_mas or de_menos:
        if de_mas:
            partes.append("un " + " y un ".join(de_mas) + " más")
        if de_menos:
            partes.append("un " + " y un ".join(de_menos) + " menos")
    if n_act != n_ant:
        partes.append(f"{n_act} días contra {n_ant}")
    if not partes:
        return None, n_act, n_ant
    return (f"{MESES_ES[ultimo.month].capitalize()} {ultimo.year} tuvo "
            + " y ".join(partes) + f" que {MESES_ES[ultimo.month].lower()} "
            f"{ultimo.year - 1}. Parte de la variación responde a la composición "
            f"del calendario y no a la performance de los locales."), n_act, n_ant


# --- Snapshot / candado ---------------------------------------------------------
def cargar_cierres():
    if not CIERRES.exists():
        return {}
    return json.loads(CIERRES.read_text(encoding="utf-8"))


def guardar_cierre(clave, dia1, ultimo, mes_act, mes_ant, totals):
    cierres = cargar_cierres()
    cierres[clave] = {
        "desde": dia1.isoformat(),
        "hasta": ultimo.isoformat(),
        "total": round(totals["a26"], 2),
        "total_anio_anterior": round(totals["a25"], 2),
        "variacion_pct": round(totals["pct"], 2),
        "por_sucursal": {b: mes_act[b] for b in BRANCH_ORDER},
        "por_sucursal_anio_anterior": {b: round(mes_ant[b], 2) for b in BRANCH_ORDER},
        "cerrado": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    CIERRES.parent.mkdir(parents=True, exist_ok=True)
    CIERRES.write_text(
        json.dumps({k: cierres[k] for k in sorted(cierres)}, indent=2, ensure_ascii=False),
        encoding="utf-8")


# --- Calculo --------------------------------------------------------------------
def construir(dia1, ultimo):
    mes_act = leer_acum_cerrado(dia1, ultimo)
    tks_act = leer_tickets_mes(dia1, ultimo)

    ia, fa = espejo(dia1), espejo(ultimo)
    acum_ant, tks_ant, faltan = acumular_rango(ia, fa)
    if faltan:
        print(f"[ERROR] Al master del anio anterior le faltan dias del rango {ia}..{fa}:")
        for f in faltan:
            print(f"  - {f.isoformat()}")
        raise SystemExit(1)
    escribir_excel(ia, fa, acum_ant, SALIDA_ANT, tks_ant)
    _, _, mes_ant, mes_ant_tks = parse_excel_full(SALIDA_ANT)

    rows = []
    for b in BRANCH_ORDER:
        a26 = mes_act[b]
        if b in SIN_HISTORICO_2025:
            # Local nuevo 2026: no existia en el mes espejo de 2025. Sin comparable.
            r = {"branch": b, "a26": a26, "a25": None, "diff": None, "pct": None,
                 "comparable": False}
            if tks_act:
                t26 = tks_act[b]
                r["tks26"], r["tks25"] = t26, 0
                r["tp26"] = (a26 / t26) if t26 else 0.0
                r["tp25"], r["tp_pct"] = None, None
        else:
            a25 = round(mes_ant[b], 2)
            diff = a26 - a25
            r = {"branch": b, "a26": a26, "a25": a25, "diff": diff,
                 "pct": (diff / a25 * 100) if a25 else 0.0, "comparable": True}
            if tks_act:
                t26, t25 = tks_act[b], mes_ant_tks[b]
                r["tks26"], r["tks25"] = t26, t25
                r["tp26"] = (a26 / t26) if t26 else 0.0
                r["tp25"] = (a25 / t25) if t25 else 0.0
                r["tp_pct"] = ((r["tp26"] - r["tp25"]) / r["tp25"] * 100) if r["tp25"] else 0.0
        rows.append(r)

    def agregar(items):
        comp = [r for r in items if r["comparable"]]
        t = {"a26": round(sum(r["a26"] for r in items), 2),
             "a25": round(sum((r["a25"] or 0.0) for r in items), 2),
             "hay_nuevo": any(not r["comparable"] for r in items)}
        c26 = round(sum(r["a26"] for r in comp), 2)
        c25 = round(sum(r["a25"] for r in comp), 2)
        t["ca26"], t["ca25"] = c26, c25
        t["diff"] = c26 - c25
        t["pct"] = (t["diff"] / c25 * 100) if c25 else 0.0
        return t

    totals = agregar(rows)
    if tks_act:
        totals["tks26"] = sum(r["tks26"] for r in rows)
        totals["tks25"] = sum(r["tks25"] for r in rows)
        # Ticket promedio = venta total / tickets totales. NUNCA el promedio de
        # los promedios de cada sucursal: eso le da el mismo peso a Aventura que
        # a Sawgrass y el numero sale mal.
        totals["tp26"] = totals["a26"] / totals["tks26"] if totals["tks26"] else 0.0
        totals["tp25"] = totals["a25"] / totals["tks25"] if totals["tks25"] else 0.0
        totals["tks_pct"] = ((totals["tks26"] - totals["tks25"]) / totals["tks25"] * 100) if totals["tks25"] else 0.0
        totals["tp_pct"] = ((totals["tp26"] - totals["tp25"]) / totals["tp25"] * 100) if totals["tp25"] else 0.0
    # Participacion de cada sucursal sobre el total del mes: en el cierre si tiene
    # sentido (en el diario no, porque un dia flojo de una sucursal te distorsiona).
    for r in rows:
        r["share"] = (r["a26"] / totals["a26"] * 100) if totals["a26"] else 0.0

    return (rows, totals, agregar([r for r in rows if r["branch"] in PROPIAS]),
            agregar([r for r in rows if r["branch"] in FRANQUICIAS]),
            mes_act, mes_ant, ia, fa, bool(tks_act))


# --- HTML -----------------------------------------------------------------------
def chip(pct, diff, grande=False, nota=False):
    if pct is None:
        return ('<span style="display:inline-block;background:#eeeeee;color:#8a8a8a;'
                'font-weight:700;font-size:12px;padding:4px 11px;border-radius:20px;'
                'white-space:nowrap;">s/ comp.</span>')
    up = pct >= 0
    col = "#1a7d2e" if up else "#c62828"
    bg = "#eaf5ec" if up else "#fbecec"
    s = "+" if up else ""
    fs = "17px" if grande else "12px"
    dd = ""
    if diff:
        txt = f"(+{money(diff)})" if diff >= 0 else f"({money(diff)})"
        dd = f'<div style="color:{col};font-size:11px;margin-top:4px;">{txt}</div>'
    return (f'<span style="display:inline-block;background:{bg};color:{col};'
            f'font-weight:700;font-size:{fs};padding:4px 11px;border-radius:20px;'
            f'white-space:nowrap;">{s}{pct:.1f}%{"*" if nota else ""}</span>{dd}')


def render_html(dia1, ultimo, rows, totals, propias, franquicias, ia, fa, con_tks, ytd):
    anio = ultimo.year
    mes_txt = MESES_ES[ultimo.month]
    a26_lbl = f"{MES_CORTO[ultimo.month].upper()}/{str(anio)[2:]}"
    a25_lbl = f"{MES_CORTO[ultimo.month].upper()}/{str(anio - 1)[2:]}"
    dias_mes = ultimo.day
    prom = totals["a26"] / dias_mes

    texto_cal, n_act, n_ant = composicion(dia1, ultimo)
    nota_cal = ""
    if texto_cal:
        nota_cal = (f'<div style="background:#fff8e6;border-left:3px solid #e0a800;'
                    f'border-radius:6px;padding:12px 14px;margin-top:16px;color:#6b5200;'
                    f'font-size:12px;line-height:1.6;">'
                    f'<b>Nota sobre el calendario.</b> {texto_cal}</div>')

    if not con_tks:
        banda_tks = ('<div style="background:#fbecec;border-radius:12px;padding:14px 18px;'
                     'color:#c62828;font-size:12px;">Sin datos de tickets para este mes: '
                     'el historial diario no cubre el período completo. La venta no está afectada.</div>')
    else:
        banda_tks = f"""
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td width="50%" style="padding-right:7px;vertical-align:top;">
          <div style="background:#f5f5f5;border-radius:12px;padding:18px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">TICKETS DEL MES</div>
            <table role="presentation" width="100%"><tr>
              <td style="color:#111111;font-size:20px;font-weight:800;padding-top:6px;">{totals['tks26']:,}</td>
              <td style="text-align:right;">{chip(totals['tks_pct'], 0)}</td>
            </tr></table>
            <div style="color:#9a9a9a;font-size:11px;margin-top:2px;">{totals['tks25']:,} en {anio - 1}</div>
          </div>
        </td>
        <td width="50%" style="padding-left:7px;vertical-align:top;">
          <div style="background:#f5f5f5;border-radius:12px;padding:18px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">TICKET PROMEDIO</div>
            <table role="presentation" width="100%"><tr>
              <td style="color:#111111;font-size:20px;font-weight:800;padding-top:6px;">{money2(totals['tp26'])}</td>
              <td style="text-align:right;">{chip(totals['tp_pct'], 0)}</td>
            </tr></table>
            <div style="color:#9a9a9a;font-size:11px;margin-top:2px;">{money2(totals['tp25'])} en {anio - 1}</div>
          </div>
        </td>
      </tr>
    </table>"""

    def fila(r, zebra):
        return f"""
        <tr style="background:{zebra};">
          <td style="padding:14px 18px;font-weight:700;color:#111111;font-size:14px;">{r['branch']}
            <div style="color:#9a9a9a;font-size:11px;font-weight:400;margin-top:2px;">{r['share']:.1f}% del total{'' if not con_tks else f" · ticket prom. {money2(r['tp26'])}"}</div>
          </td>
          <td style="padding:14px 12px;text-align:right;color:#111111;font-weight:700;font-size:14px;">{money(r['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#9a9a9a;font-size:14px;">{money(r['a25']) if r['a25'] is not None else "—"}</td>
          <td style="padding:14px 18px;text-align:right;">{chip(r['pct'], r['diff'])}</td>
        </tr>"""

    def encabezado_grupo(nombre):
        return f"""
        <tr style="background:#eef1f6;">
          <td colspan="4" style="padding:9px 18px;color:#1f3a6e;font-size:10px;font-weight:800;letter-spacing:2px;">{nombre}</td>
        </tr>"""

    def fila_subtotal(nombre, s):
        return f"""
        <tr style="background:#f5f5f5;border-top:1px solid #e2e2e2;">
          <td style="padding:14px 18px;font-weight:800;color:#1f3a6e;font-size:13px;">{nombre}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#1f3a6e;font-size:13px;">{money(s['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#7a8aa5;font-size:13px;">{money(s['a25'])}</td>
          <td style="padding:14px 18px;text-align:right;">{chip(s['pct'], s['diff'], nota=s.get('hay_nuevo'))}</td>
        </tr>"""

    by_name = {r["branch"]: r for r in rows}
    cuerpo = encabezado_grupo("PROPIAS")
    for i, b in enumerate(PROPIAS):
        cuerpo += fila(by_name[b], "#ffffff" if i % 2 == 0 else "#fafafa")
    cuerpo += fila_subtotal("Subtotal Propias", propias)
    cuerpo += encabezado_grupo("FRANQUICIAS")
    for i, b in enumerate(FRANQUICIAS):
        cuerpo += fila(by_name[b], "#ffffff" if i % 2 == 0 else "#fafafa")
    cuerpo += fila_subtotal("Subtotal Franquicias", franquicias)

    # --- Bloque ACUMULADO DEL AÑO (YTD calendario) ---
    filas_ytd = ""
    for i, b in enumerate(BRANCH_ORDER):
        v, va = round(ytd["por26"][b], 2), round(ytd["por25"][b], 2)
        sc = (b in SIN_HISTORICO_2025) or (va == 0)
        dp = None if sc else ((v - va) / va * 100)
        share = (v / ytd["t26"] * 100) if ytd["t26"] else 0.0
        z = "#ffffff" if i % 2 == 0 else "#fafafa"
        filas_ytd += f"""
        <tr style="background:{z};">
          <td style="padding:11px 16px;color:#111111;font-size:13px;font-weight:600;">{b}<span style="color:#9a9a9a;font-size:11px;font-weight:400;"> &middot; {share:.1f}%</span></td>
          <td style="padding:11px 12px;text-align:right;color:#111111;font-size:13px;">{money(v)}</td>
          <td style="padding:11px 16px;text-align:right;">{chip(dp, 0)}</td>
        </tr>"""
    bloque_ytd = f"""
  <!-- YTD -->
  <tr><td style="padding:26px 32px 6px 32px;">
    <div style="border-top:2px solid #0f1c33;padding-top:22px;">
      <div style="color:#0f1c33;font-size:12px;letter-spacing:3px;font-weight:800;">ACUMULADO DEL A&Ntilde;O &middot; INTERANUAL</div>
      <div style="color:#9a9a9a;font-size:11px;margin-top:4px;line-height:1.5;">Acumulado del a&ntilde;o calendario hasta el mes cerrado ({ytd["rango"]} {anio}) contra el mismo per&iacute;odo de {anio - 1}.</div>
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:16px;"><tr>
        <td width="50%" style="padding-right:7px;vertical-align:top;">
          <div style="background:#0f1c33;border-radius:12px;padding:20px;height:118px;">
            <div style="color:#8ea6cf;font-size:10px;letter-spacing:1px;">ACUM. {ytd["lbl"]} {anio}</div>
            <div style="color:#ffffff;font-size:24px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(ytd["t26"])}</div>
            <div style="color:#8ea6cf;font-size:11px;margin-top:8px;">ticket prom. {money(ytd["tp"])} &middot; {ytd["tks"]:,} tickets</div>
          </div>
        </td>
        <td width="50%" style="padding-left:7px;vertical-align:top;">
          <div style="background:#f5f5f5;border-radius:12px;padding:20px;height:118px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">VARIACI&Oacute;N INTERANUAL</div>
            <div style="margin-top:12px;">{chip(ytd["pct"], ytd["diff"], grande=True, nota=ytd.get("hay_nuevo"))}</div>
            <div style="color:#9a9a9a;font-size:11px;margin-top:10px;">{ytd["rango"]} {anio - 1} ({money(ytd["t25"])})</div>
          </div>
        </td>
      </tr></table>
      <div style="background:#fafafa;border-radius:12px;padding:18px 20px;margin-top:16px;">
        <div style="color:#9a9a9a;font-size:11px;letter-spacing:2px;margin-bottom:4px;">VENTA MENSUAL &middot; {ytd["lbl"]} {anio} vs {anio - 1}</div>
        <div style="color:#9a9a9a;font-size:11px;margin-bottom:8px;line-height:1.5;">Cada par de barras es un mes (a&ntilde;o anterior vs actual); arriba, el acumulado del a&ntilde;o de cada uno.</div>
        <img src="cid:ytd" alt="Acumulado del a&ntilde;o" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
      </div>
      <div style="margin-top:16px;">
        <div style="color:#9a9a9a;font-size:11px;letter-spacing:2px;margin-bottom:10px;">POR SUCURSAL &middot; ACUMULADO DEL A&Ntilde;O vs {anio - 1}</div>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-radius:10px;overflow:hidden;box-shadow:0 1px 5px rgba(0,0,0,0.05);">
          <thead><tr style="background:#0f1c33;">
            <th style="padding:11px 16px;text-align:left;color:#ffffff;font-size:10px;letter-spacing:1px;">SUCURSAL &middot; PART.</th>
            <th style="padding:11px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;">ACUM. A&Ntilde;O</th>
            <th style="padding:11px 16px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;">vs. A&Ntilde;O ANT.</th>
          </tr></thead>
          <tbody>{filas_ytd}
            <tr style="background:#0f1c33;">
              <td style="padding:13px 16px;font-weight:800;color:#ffffff;font-size:13px;">TOTAL</td>
              <td style="padding:13px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:13px;">{money(ytd["t26"])}</td>
              <td style="padding:13px 16px;text-align:right;">{chip(ytd["pct"], 0)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div style="background:#fff8e6;border-left:3px solid #e0a800;border-radius:6px;padding:11px 14px;margin-top:14px;color:#6b5200;font-size:11px;line-height:1.6;"><b>C&oacute;mo leerlo.</b> Es el acumulado del a&ntilde;o calendario a la misma fecha en los dos a&ntilde;os, as&iacute; que compara lo que va del a&ntilde;o contra el a&ntilde;o pasado a esta altura.</div>
    </div>
  </td></tr>
"""
    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#eeeeee;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eeeeee;">
<tr><td align="center" style="padding:24px 12px;">

<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:600px;max-width:600px;background:#ffffff;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,0.12);">

  <!-- HEADER -->
  <tr><td style="background:#000000;padding:38px 32px 32px 32px;text-align:center;">
    <img src="cid:logo" alt="Lucciano's" width="190" style="display:block;margin:0 auto;max-width:190px;height:auto;">
    <div style="color:#bdbdbd;font-size:12px;letter-spacing:4px;margin-top:18px;">CIERRE MENSUAL DE VENTAS</div>
    <div style="color:#ffffff;font-size:17px;font-weight:800;letter-spacing:3px;margin-top:14px;">{mes_txt} {anio}</div>
    <div style="color:#777777;font-size:11px;margin-top:8px;">mes completo · {dias_mes} días · 9 sucursales</div>
  </td></tr>

  <!-- KPIs -->
  <tr><td style="padding:30px 32px 6px 32px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td width="33%" style="padding-right:7px;vertical-align:top;">
          <div style="background:#111111;border-radius:12px;padding:20px;height:150px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">VENTA DE {a26_lbl}</div>
            <div style="color:#ffffff;font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['a26'])}</div>
            <div style="color:#777777;font-size:11px;margin-top:6px;">promedio {money(prom)}/día</div>
          </div>
        </td>
        <td width="34%" style="padding:0 7px;vertical-align:top;">
          <div style="background:#f5f5f5;border-radius:12px;padding:20px;height:150px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">VENTA DE {a25_lbl}</div>
            <div style="color:#111111;font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['a25'])}</div>
            <div style="color:#9a9a9a;font-size:10px;margin-top:6px;">{ia.strftime('%d/%m')} al {fa.strftime('%d/%m/%Y')}</div>
          </div>
        </td>
        <td width="33%" style="padding-left:7px;vertical-align:top;">
          <div style="background:#f5f5f5;border-radius:12px;padding:20px;height:150px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">VARIACIÓN INTERANUAL</div>
            <div style="margin-top:14px;">{chip(totals['pct'], totals['diff'], grande=True, nota=totals.get('hay_nuevo'))}</div>
            {f'<div style="color:#9a9a9a;font-size:11px;margin-top:10px;line-height:1.5;">{money(totals["ca26"])} vs {money(totals["ca25"])}<br>{sum(1 for b in BRANCH_ORDER if b not in SIN_HISTORICO_2025)} locales comparables</div>' if totals.get('hay_nuevo') else ''}
          </div>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- TICKETS -->
  <tr><td style="padding:18px 32px 6px 32px;">
    {banda_tks}
  </td></tr>

  <!-- COMPARATIVO -->
  <tr><td style="padding:22px 32px 6px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:12px;">COMPARATIVO POR SUCURSAL · {str(anio)[2:]} vs {str(anio - 1)[2:]}</div>
    <img src="cid:comparativo" alt="Comparativo por sucursal" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
  </td></tr>

  <!-- DETALLE -->
  <tr><td style="padding:24px 32px 36px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:14px;">DETALLE POR SUCURSAL</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-radius:12px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,0.06);">
      <thead>
        <tr style="background:#111111;">
          <th style="padding:13px 18px;text-align:left;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">SUCURSAL</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">{a26_lbl}</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">{a25_lbl}</th>
          <th style="padding:13px 18px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">VARIACIÓN</th>
        </tr>
      </thead>
      <tbody>{cuerpo}
        <tr style="background:#111111;">
          <td style="padding:17px 18px;font-weight:800;color:#ffffff;font-size:14px;">TOTAL {a26_lbl}</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['a26'])}</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['a25'])}</td>
          <td style="padding:17px 18px;text-align:right;">{chip(totals['pct'], totals['diff'])}</td>
        </tr>
      </tbody>
    </table>
    {nota_cal}
    <div style="color:#9a9a9a;font-size:11px;margin-top:14px;line-height:1.6;">
      Mes cerrado del {dia1.strftime('%d/%m')} al {ultimo.strftime('%d/%m/%Y')}, comparado contra
      {ia.strftime('%d/%m')} al {fa.strftime('%d/%m/%Y')} (mismas fechas calendario del año anterior).
      Ventas netas (Net Sales), sin impuestos. Las dos unidades de Vineland se informan consolidadas.
    </div>
  </td></tr>
{bloque_ytd}
  <!-- FOOTER -->
  <tr><td style="background:#000000;padding:20px 32px;text-align:center;">
    <div style="color:#777777;font-size:11px;letter-spacing:1px;">LUCCIANO'S EUROPA · Reporte automático generado el {date.today().strftime('%d/%m/%Y')}</div>
  </td></tr>

</table>

</td></tr>
</table>
</body>
</html>"""


def _gh_out(**kv):
    p = os.environ.get("GITHUB_OUTPUT")
    if not p:
        return
    with open(p, "a", encoding="utf-8") as f:
        for k, v in kv.items():
            f.write(f"{k}={v}\n")


def main():
    ap = argparse.ArgumentParser(description="Reporte de cierre mensual.")
    ap.add_argument("--mes", help="Mes a cerrar (YYYY-MM). Default: el mes anterior a hoy.")
    ap.add_argument("--forzar", action="store_true", help="Regenera aunque ya este cerrado.")
    a = ap.parse_args()

    if a.mes:
        anio, mes = int(a.mes[:4]), int(a.mes[5:7])
    else:
        anio, mes = mes_a_cerrar()
    dia1, ultimo = rango_mes(anio, mes)
    clave = dia1.strftime("%Y-%m")

    print(f"[INFO] Mes a cerrar: {clave} ({dia1} .. {ultimo})")

    if clave in cargar_cierres() and not a.forzar:
        print(f"[SKIP] El mes {clave} ya fue cerrado y enviado. No se reenvia.")
        _gh_out(send="false")
        return 0

    rows, totals, propias, franquicias, mes_act, mes_ant, ia, fa, con_tks = construir(dia1, ultimo)

    # --- Master mensual: registrar el mes cerrado y armar el ACUMULADO DEL ANIO (YTD) ---
    # Registramos el mes que acaba de cerrar (idempotente) ANTES de calcular el YTD, asi
    # el YTD incluye este mes. El anio anterior ya esta en el master (backfill).
    by = {r["branch"]: r for r in rows}
    por_suc_mes = {b: {"venta": mes_act[b], "tickets": by[b].get("tks26", 0)} for b in BRANCH_ORDER}
    mensual.registrar_mes(anio, mes, por_suc_mes)
    hist_path = BASE / "data" / f"historico_{anio}.json"
    if hist_path.exists():
        msg = mensual.conciliar_con_historial(anio, mes, json.loads(hist_path.read_text(encoding="utf-8")))
        if msg:
            print(msg)
    y26, ytks26 = mensual.ytd_por_sucursal(anio, mes)
    y25, _ = mensual.ytd_por_sucursal(anio - 1, mes)
    t26, t25 = round(sum(y26.values()), 2), round(sum(y25.values()), 2)
    _comp = [b for b in BRANCH_ORDER if b not in SIN_HISTORICO_2025]
    _c26 = round(sum(y26[b] for b in _comp), 2)
    _c25 = round(sum(y25[b] for b in _comp), 2)
    ytks = sum(ytks26.values())
    ytd = {"por26": y26, "por25": y25, "t26": t26, "t25": t25,
           "hay_nuevo": any(b in SIN_HISTORICO_2025 for b in BRANCH_ORDER),
           "diff": _c26 - _c25, "pct": ((_c26 - _c25) / _c25 * 100) if _c25 else 0.0,
           "tks": ytks, "tp": (t26 / ytks) if ytks else 0.0,
           "lbl": f"{MES_CORTO[1].upper()}-{MES_CORTO[mes].upper()}",
           "rango": f"{MESES_ES[1].lower()} a {MESES_ES[mes].lower()}"}

    # Saque el grafico de "progreso": en un mes cerrado son las mismas dos barras
    # que ya estan en las KPI y en el comparativo. Un cierre tiene que ser concreto.
    from charts import chart_comparativo, chart_ytd
    d = BASE / "charts"
    d.mkdir(exist_ok=True)
    chart_comparativo(rows, f"{MES_CORTO[mes]}/{str(anio)[2:]}",
                      f"{MES_CORTO[mes]}/{str(anio - 1)[2:]}", d / "cierre_comparativo.png")
    # Grafico YTD: barras por mes (ene..mes), anio anterior vs actual
    master = mensual.cargar()
    labels = [MES_CORTO[m] for m in range(1, mes + 1)]
    v_ant = [sum(master[f"{anio - 1}-{m:02d}"][b]["venta"] for b in BRANCH_ORDER) for m in range(1, mes + 1)]
    v_act = [sum(master[f"{anio}-{m:02d}"][b]["venta"] for b in BRANCH_ORDER) for m in range(1, mes + 1)]
    chart_ytd(labels, v_ant, v_act, t25, t26, str(anio - 1), str(anio), d / "cierre_ytd.png")

    PREVIEW.write_text(render_html(dia1, ultimo, rows, totals, propias, franquicias, ia, fa, con_tks, ytd),
                       encoding="utf-8")
    guardar_cierre(clave, dia1, ultimo, mes_act, mes_ant, totals)

    subject = f"Cierre {MESES_ES[mes].capitalize()} {anio} - Lucciano's Europa"
    _gh_out(send="true", subject=subject, mes=clave)

    print(f"OK - cierre {clave}: {money(totals['a26'])} vs {money(totals['a25'])} "
          f"({totals['pct']:+.1f}%) | snapshot guardado en cierres.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
