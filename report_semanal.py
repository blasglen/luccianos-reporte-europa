"""
Lucciano's - Reporte de Ventas SEMANAL (lunes a domingo)
--------------------------------------------------------
Se manda los LUNES y cubre la semana calendario anterior: lunes a domingo (ayer).

Flujo:
  1. Calcula el rango lun..dom (por defecto, la semana que termino ayer).
  2. Lee data/historico_<anio>.json -> los 7 dias del rango, por sucursal.
     Si falta alguno, CORTA EN ROJO. No se manda una semana incompleta.
  3. Arma el comparativo del anio anterior con generar_acum_ant (fechas calendario)
     y lo deja en Acumulado_semanal_ant.xlsx (queda como evidencia auditable).
  4. Calcula variaciones, subtotales Propias/Franquicias, y la serie por dia.
  5. Genera preview_semanal.html + los 3 graficos.

DIFERENCIAS DE DISENIO CONTRA EL DIARIO (a proposito):
  - No hay acumulador incremental: la semana se recalcula entera desde el
    historial cada vez. Si un dia se corrige, el semanal lo toma solo.
  - Si hay estado, pero es solo para no MANDAR dos veces la misma semana
    (data/semanal.json). Es un candado de envio, no un acumulador de plata.
"""
import argparse
import json
import os
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from report import (
    BRANCH_ORDER, PROPIAS, FRANQUICIAS, MESES_ES, MES_CORTO, money, money2, parse_excel_full, SIN_HISTORICO_2025,
    DIAS_ES, cerrado_por_descanso,
)
from generar_acum_ant import acumular_rango, escribir_excel, espejo
import mensual

BASE = Path(__file__).parent
SALIDA_ACUM_ANT = BASE / "Acumulado_semanal_ant.xlsx"
SALIDA_ACUM_MES_ANT = BASE / "Acumulado_mensual_ant.xlsx"
ESTADO_DIARIO = BASE / "data" / "acumulado.json"
ESTADO = BASE / "data" / "semanal.json"
PREVIEW = BASE / "preview_semanal.html"

DIAS_CORTOS = {0: "Lun", 1: "Mar", 2: "Mié", 3: "Jue", 4: "Vie", 5: "Sáb", 6: "Dom"}
LUNES = 0  # date.weekday(): lunes=0 ... domingo=6


# --- 1. Rango de la semana -------------------------------------------------------
def rango_semana(hoy=None):
    """Devuelve (lunes, domingo) de la semana calendario a reportar.

    Regla: el reporte sale el lunes y cubre hasta AYER (domingo). O sea que
    'hasta' = hoy - 1 y 'desde' = hasta - 6 = el lunes anterior.

    Si por lo que sea corre otro dia (prueba manual, o se atraso), retrocede al
    ultimo domingo cerrado en vez de reventar. Asi el reporte siempre cubre una
    semana lun..dom completa, corra el dia que corra.
    """
    hoy = hoy or date.today()
    # Dias a retroceder desde 'hoy' hasta el lunes de arranque.
    # Lunes -> 7 (el lunes anterior). Martes -> 8. Domingo -> 13.
    delta = (hoy.weekday() - LUNES) % 7 + 7
    desde = hoy - timedelta(days=delta)
    return desde, desde + timedelta(days=6)


def _venta(dia, b):
    """El historial viejo guardaba {suc: float}; el nuevo {suc: {venta, tickets}}.
    Leo los dos para que un rescate a medias no rompa el reporte."""
    v = dia.get(b, 0.0)
    return float(v["venta"]) if isinstance(v, dict) else float(v)


def _tickets(dia, b):
    v = dia.get(b, 0.0)
    return int(v.get("tickets", 0)) if isinstance(v, dict) else 0


def _sin_tickets(dias):
    """True si algun dia del rango no tiene tickets (historial viejo)."""
    return any(not isinstance(d.get(b), dict) for _, d in dias for b in BRANCH_ORDER)


# --- 2. Historial ---------------------------------------------------------------
def leer_semana(desde, hasta):
    """Levanta los dias del rango del historial. Falla RUIDOSA si falta alguno."""
    path = BASE / "data" / f"historico_{hasta.year}.json"
    if not path.exists():
        raise SystemExit(
            f"[ERROR] No existe {path.name}. El historial lo escribe historico.py "
            f"en el workflow diario; todavia no corrio ninguna vez."
        )
    hist = json.loads(path.read_text(encoding="utf-8"))

    dias, faltantes = [], []
    d = desde
    while d <= hasta:
        k = d.isoformat()
        if k not in hist:
            faltantes.append(d)
        else:
            dias.append((d, hist[k]))
        d += timedelta(days=1)

    if faltantes:
        print(f"[ERROR] Al historial le faltan {len(faltantes)} dia(s) de la semana "
              f"{desde} .. {hasta}:")
        for f in faltantes:
            print(f"  - {f.isoformat()} ({DIAS_CORTOS[f.weekday()]})")
        print("Probablemente fallo el reporte diario de esos dias. Revisar Actions.")
        print("NO se manda una semana incompleta: el total quedaria subestimado y la "
              "variacion contra el anio anterior seria falsa.")
        raise SystemExit(1)

    # Si el rango cruza el 1ro de enero, el historial esta partido en dos archivos.
    # Con semanas vie..jue esto pasa una vez al anio; lo aviso en vez de mentir.
    if desde.year != hasta.year:
        raise SystemExit(
            f"[ERROR] La semana {desde}..{hasta} cruza el cambio de anio y el "
            f"historial esta separado por anio. Hay que unificar a mano esta vez."
        )
    return dias


# --- 2 bis. Acumulado del mes (viene del diario) --------------------------------
def leer_acum_mes(hasta):
    """Devuelve el acumulado del mes por sucursal, al cierre del jueves.

    NO lo recalculo: lo toma tal cual de data/acumulado.json, que es el estado que
    viene manteniendo report.py dia a dia desde el 1ro del mes. Es LA fuente del
    numero que vos ya venis mirando todos los dias, asi que el semanal y el diario
    no pueden decir cosas distintas.

    La dependencia con el diario es real, asi que la hago EXPLICITA y la valido:
    si el diario no cerro el jueves, esto corta. Prefiero eso a mandarle a los
    socios un acumulado que va un dia atrasado sin que se note.
    """
    if not ESTADO_DIARIO.exists():
        raise SystemExit(f"[ERROR] No existe {ESTADO_DIARIO.name}.")
    st = json.loads(ESTADO_DIARIO.read_text(encoding="utf-8"))

    mes_esperado = hasta.strftime("%Y-%m")
    if st.get("month") != mes_esperado:
        raise SystemExit(
            f"[ERROR] El acumulado del diario es del mes {st.get('month')} y el "
            f"cierre de esta semana es {mes_esperado}. No se manda."
        )
    if st.get("last_date") != hasta.isoformat():
        raise SystemExit(
            f"[ERROR] El acumulado del diario esta cerrado al {st.get('last_date')} "
            f"y el semanal cierra el {hasta.isoformat()}.\n"
            f"        Causa tipica: el reporte diario del domingo no corrio o fallo "
            f"(mail de TouchBistro demorado). Revisar Actions y volver a correr el "
            f"semanal despues.\n"
            f"        NO se manda: el acumulado del mes quedaria incompleto."
        )

    acum = st.get("acumulado", {})
    faltan = [b for b in BRANCH_ORDER if b not in acum]
    if faltan:
        raise SystemExit(f"[ERROR] El acumulado del diario no tiene: {', '.join(faltan)}")
    return {b: round(float(acum[b]), 2) for b in BRANCH_ORDER}


def conciliar(hasta, acum_mes):
    """CONTROL: si el historial ya cubre el mes entero hasta el jueves, la suma
    del historial TIENE que dar igual al acumulado del diario. Son dos caminos
    independientes al mismo numero.

    Hoy el historial arranca a mitad de julio, asi que este control se saltea solo
    y se va a prender automaticamente en agosto, cuando haya mes completo. No hay
    nada que activar despues.

    Devuelve un texto para el log, o None si no habia con que conciliar.
    """
    path = BASE / "data" / f"historico_{hasta.year}.json"
    if not path.exists():
        return None
    hist = json.loads(path.read_text(encoding="utf-8"))

    d = date(hasta.year, hasta.month, 1)
    suma = {b: 0.0 for b in BRANCH_ORDER}
    while d <= hasta:
        k = d.isoformat()
        if k not in hist:
            return (f"[CONCILIACION] Omitida: el historial no cubre todo "
                    f"{hasta.strftime('%m/%Y')} (falta al menos {k}). "
                    f"Se prende sola cuando haya un mes completo.")
        for b in BRANCH_ORDER:
            suma[b] += _venta(hist[k], b)
        d += timedelta(days=1)

    difs = {b: round(suma[b] - acum_mes[b], 2) for b in BRANCH_ORDER
            if abs(suma[b] - acum_mes[b]) > 0.01}
    if difs:
        det = " | ".join(f"{b}: {v:+,.2f}" for b, v in difs.items())
        raise SystemExit(
            f"[ERROR] CONCILIACION FALLIDA. El historial y el acumulado del diario "
            f"no cierran:\n        {det}\n"
            f"        Alguno de los dos tiene un dia de mas, de menos, o mal sumado. "
            f"NO se manda hasta entender por que."
        )
    return (f"[CONCILIACION] OK: historial y acumulado del diario coinciden "
            f"al centavo (${sum(suma.values()):,.2f}).")


# --- 3 y 4. Calculo -------------------------------------------------------------
def construir(desde, hasta):
    dias = leer_semana(desde, hasta)

    # (a) Semana actual, por sucursal: venta y tickets
    sem_act = {b: 0.0 for b in BRANCH_ORDER}
    sem_tks = {b: 0 for b in BRANCH_ORDER}
    for _, d in dias:
        for b in BRANCH_ORDER:
            sem_act[b] += _venta(d, b)
            sem_tks[b] += _tickets(d, b)

    # (b) Acumulado del mes 2026: lo trae el diario, ya validado al jueves
    mes_act = leer_acum_mes(hasta)
    msg = conciliar(hasta, mes_act)
    if msg:
        print(msg)

    # (c) Comparativo anio anterior: escribo el Excel y lo leo con el MISMO parser
    # que usa el diario. Redundante? No: garantiza que las dos comparaciones
    # apliquen identico criterio de consolidacion, y me deja el archivo como
    # respaldo por si alguien pregunta de donde salio el numero.
    def comparativo(ini, fin, salida):
        ia, fa = espejo(ini), espejo(fin)
        acum, tks, faltan = acumular_rango(ia, fa)
        if faltan:
            print(f"[ERROR] Al master del anio anterior le faltan dias del rango "
                  f"{ia}..{fa}:")
            for f in faltan:
                print(f"  - {f.isoformat()}")
            raise SystemExit(1)
        escribir_excel(ia, fa, acum, salida, tks)
        _, _, ventas, tickets = parse_excel_full(salida)
        return ventas, tickets

    sem_ant, sem_ant_tks = comparativo(desde, hasta, SALIDA_ACUM_ANT)
    # El acumulado del mes anterior arranca SIEMPRE el 1ro del mes del jueves.
    # Si la semana cruza fin de mes (ej. vie 31/07 a jue 06/08), el acumulado es
    # el de agosto (1 al 6), igual que hace el diario al reiniciar el 1ro.
    dia1 = date(hasta.year, hasta.month, 1)
    mes_ant, _ = comparativo(dia1, hasta, SALIDA_ACUM_MES_ANT)

    rows = []
    for b in BRANCH_ORDER:
        s26, s25 = round(sem_act[b], 2), round(sem_ant[b], 2)
        a26, a25 = mes_act[b], round(mes_ant[b], 2)
        t26, t25 = sem_tks[b], sem_ant_tks[b]
        tp26 = (s26 / t26) if t26 else 0.0
        tp25 = (s25 / t25) if t25 else 0.0
        comparable = b not in SIN_HISTORICO_2025
        rows.append({
            "branch": b, "comparable": comparable,
            "sem26": s26, "sem25": s25 if comparable else None,
            "sem_diff": (s26 - s25) if comparable else None,
            "sem_pct": (((s26 - s25) / s25 * 100) if s25 else 0.0) if comparable else None,
            "a26": a26, "a25": a25 if comparable else None,
            "diff": (a26 - a25) if comparable else None,
            "pct": (((a26 - a25) / a25 * 100) if a25 else 0.0) if comparable else None,
            "tks26": t26, "tks25": t25,
            "tks_pct": ((t26 - t25) / t25 * 100) if t25 else 0.0,
            "tp26": tp26, "tp25": tp25,
            "tp_pct": ((tp26 - tp25) / tp25 * 100) if tp25 else 0.0,
        })

    def agregar(items):
        comp = [r for r in items if r["comparable"]]
        t = {"hay_nuevo": any(not r["comparable"] for r in items)}
        t["sem26"] = round(sum(r["sem26"] for r in items), 2)
        t["a26"]   = round(sum(r["a26"] for r in items), 2)
        t["sem25"] = round(sum((r["sem25"] or 0.0) for r in items), 2)
        t["a25"]   = round(sum((r["a25"] or 0.0) for r in items), 2)
        t["_csem26"] = round(sum(r["sem26"] for r in comp), 2)
        t["_csem25"] = round(sum(r["sem25"] for r in comp), 2)
        t["_ca26"]   = round(sum(r["a26"] for r in comp), 2)
        t["_ca25"]   = round(sum(r["a25"] for r in comp), 2)
        t["ca26"], t["ca25"] = t["_ca26"], t["_ca25"]
        t["csem26"] = round(sum(r["sem26"] for r in comp), 2)
        t["csem25"] = round(sum(r["sem25"] for r in comp), 2)
        for k in ("tks26", "tks25"):
            t[k] = sum(r[k] for r in items)
        # OJO: el ticket promedio del grupo NO es el promedio de los promedios.
        # Hay que dividir la venta total por los tickets totales, sino una
        # sucursal chica pesa igual que una grande y el numero sale mal.
        t["tp26"] = (t["sem26"] / t["tks26"]) if t["tks26"] else 0.0
        t["tp25"] = (t["sem25"] / t["tks25"]) if t["tks25"] else 0.0
        t["tks_pct"] = ((t["tks26"] - t["tks25"]) / t["tks25"] * 100) if t["tks25"] else 0.0
        t["tp_pct"] = ((t["tp26"] - t["tp25"]) / t["tp25"] * 100) if t["tp25"] else 0.0
        t["sem_diff"] = t["_csem26"] - t["_csem25"]
        t["sem_pct"] = (t["sem_diff"] / t["_csem25"] * 100) if t["_csem25"] else 0.0
        t["diff"] = t["_ca26"] - t["_ca25"]
        t["pct"] = (t["diff"] / t["_ca25"] * 100) if t["_ca25"] else 0.0
        return t

    totals = agregar(rows)
    propias = agregar([r for r in rows if r["branch"] in PROPIAS])
    franquicias = agregar([r for r in rows if r["branch"] in FRANQUICIAS])

    # Serie por dia: el actual sale del historial; el anio anterior, del master
    # dia por dia (por eso acumulo rangos de un solo dia).
    # Serie por dia: SOLO 2026. No comparo contra 2025 dia a dia porque el espejo
    # calendario alinea "Lun 13" contra un domingo. Ver el docstring de chart_dias.
    serie = []
    for f, d in dias:
        serie.append({
            "fecha": f,
            "etiqueta": f"{DIAS_CORTOS[f.weekday()]} {f.day}",
            "actual": round(sum(_venta(d, b) for b in BRANCH_ORDER), 2),
        })

    # El mejor y el peor dia, para el texto del mail
    mejor = max(serie, key=lambda s: s["actual"])
    peor = min(serie, key=lambda s: s["actual"])

    return rows, totals, propias, franquicias, serie, mejor, peor, dia1, _sin_tickets(dias)


# --- Candado de envio -----------------------------------------------------------
def ya_enviada(desde, hasta):
    if not ESTADO.exists():
        return False
    st = json.loads(ESTADO.read_text(encoding="utf-8"))
    return st.get("last_week") == f"{desde.isoformat()}..{hasta.isoformat()}"


def marcar_enviada(desde, hasta, totals):
    ESTADO.parent.mkdir(parents=True, exist_ok=True)
    ESTADO.write_text(json.dumps({
        "last_week": f"{desde.isoformat()}..{hasta.isoformat()}",
        "enviado": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "total": round(totals["a26"], 2),
    }, indent=2, ensure_ascii=False), encoding="utf-8")


# --- 5. HTML --------------------------------------------------------------------
# El HTML lo arma mail_semanal.py (diseño "combinado", oct/2026). La version
# anterior (render_html + graficos PNG de charts.py) quedo en el historial de git.


def _gh_out(**kv):
    p = os.environ.get("GITHUB_OUTPUT")
    if not p:
        return
    with open(p, "a", encoding="utf-8") as f:
        for k, v in kv.items():
            f.write(f"{k}={v}\n")


def main():
    ap = argparse.ArgumentParser(description="Reporte semanal viernes a jueves.")
    ap.add_argument("--hasta", help="Domingo de cierre (YYYY-MM-DD). Default: ayer.")
    ap.add_argument("--forzar", action="store_true",
                    help="Ignora el candado y regenera aunque ya se haya enviado.")
    a = ap.parse_args()

    if a.hasta:
        hasta = datetime.strptime(a.hasta, "%Y-%m-%d").date()
        desde = hasta - timedelta(days=6)
    else:
        desde, hasta = rango_semana()

    print(f"[INFO] Semana a reportar: {desde} ({DIAS_CORTOS[desde.weekday()]}) .. "
          f"{hasta} ({DIAS_CORTOS[hasta.weekday()]})")

    if ya_enviada(desde, hasta) and not a.forzar:
        print(f"[SKIP] La semana {desde}..{hasta} ya fue enviada. No se reenvia.")
        _gh_out(send="false")
        return 0

    rows, totals, propias, franquicias, serie, mejor, peor, dia1, sin_tks = construir(desde, hasta)
    v12, _ = mensual.acum_12m_por_sucursal()  # acum 12 meses moviles por sucursal (ultimo cerrado)
    if sin_tks:
        print("[AVISO] Hay dias sin tickets en el historial (formato viejo). "
              "El reporte sale con la venta pero sin ticket promedio. Para "
              "arreglarlo: correr el rescate de historial con pisar=true.")

    # Mail (diseño "combinado"): lo arma mail_semanal.py. Ya no lleva graficos PNG:
    # barras y graficos son tablas HTML. El unico adjunto es el logo.
    import mail_semanal
    try:
        v12_hasta = mensual.ultimo_mes(mensual.cargar())
    except Exception:
        v12_hasta = None
    PREVIEW.write_text(
        mail_semanal.render(desde, hasta, rows, totals, propias, franquicias, serie, mejor, peor,
                            dia1, sin_tks, v12, v12_hasta),
        encoding="utf-8")
    marcar_enviada(desde, hasta, totals)

    subject = (f"Reporte Semanal {desde.strftime('%d/%m')} al {hasta.strftime('%d/%m/%Y')} "
               f"- Lucciano's Europa")
    _gh_out(send="true", subject=subject, week=f"{desde.isoformat()}..{hasta.isoformat()}")

    print(f"OK - semana {money(totals['sem26'])} ({totals['sem_pct']:+.1f}%) | "
          f"acum mes {money(totals['a26'])} vs {money(totals['a25'])} "
          f"({totals['pct']:+.1f}%) | mejor dia {mejor['etiqueta']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
