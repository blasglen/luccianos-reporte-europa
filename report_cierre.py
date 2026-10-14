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
# El HTML lo arma mail_cierre.py (diseño "combinado", oct/2026). La version
# anterior (render_html + graficos PNG de charts.py) quedo en el historial de git.


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

    # Mail (diseño "combinado"): lo arma mail_cierre.py. Ya no lleva graficos PNG:
    # barras y graficos son tablas HTML. El unico adjunto es el logo.
    import mail_cierre
    master = mensual.cargar()
    labels = [MES_CORTO[m] for m in range(1, mes + 1)]
    v_ant = [sum(master[f"{anio - 1}-{m:02d}"][b]["venta"] for b in BRANCH_ORDER) for m in range(1, mes + 1)]
    v_act = [sum(master[f"{anio}-{m:02d}"][b]["venta"] for b in BRANCH_ORDER) for m in range(1, mes + 1)]
    # Mes anterior (para "vs agosto"): si es enero, diciembre del anio anterior
    pa, pm = (anio, mes - 1) if mes > 1 else (anio - 1, 12)
    clave_prev = f"{pa}-{pm:02d}"
    mes_prev = None
    if clave_prev in master:
        mes_prev = (f"{MESES_ES[pm].lower()} {pa}",
                    round(sum(master[clave_prev][b]["venta"] for b in BRANCH_ORDER), 2))
    texto_cal, _, _ = composicion(dia1, ultimo)

    PREVIEW.write_text(
        mail_cierre.render(dia1, ultimo, rows, totals, propias, franquicias, ia, fa, con_tks, ytd,
                           texto_cal, list(zip(labels, v_act, v_ant)), mes_prev),
        encoding="utf-8")
    guardar_cierre(clave, dia1, ultimo, mes_act, mes_ant, totals)

    subject = f"Cierre {MESES_ES[mes].capitalize()} {anio} - Lucciano's Europa"
    _gh_out(send="true", subject=subject, mes=clave)

    print(f"OK - cierre {clave}: {money(totals['a26'])} vs {money(totals['a25'])} "
          f"({totals['pct']:+.1f}%) | snapshot guardado en cierres.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
