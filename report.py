"""
Lucciano's - Reporte de Ventas Diario
Genera el mail HTML comparando el mes en curso (2026) contra el mismo periodo del anio anterior (2025).

Flujo:
  1. Lee Ventas_ayer.xlsx       -> venta del dia (2026)
  2. Lee acum_jun26.json        -> acumulado del mes previo (persistido en el repo)
  3. Acum.26 = acum previo + venta del dia  (se vuelve a guardar)
  4. Lee Acumulado_interanual.xlsx -> acumulado 2025 (comparativo)
  5. Variacion = (acum26 - acum25) / acum25
  6. Genera el HTML del mail
"""
import json
import os
import re
import sys
from pathlib import Path
from datetime import date, datetime, timedelta

import openpyxl

# --- Configuracion de sucursales -------------------------------------------------
# fetch_cierres.py ya consolido los ~17 PDF en los 9 locales canonicos y dejo el
# Ventas_ayer.xlsx con esos nombres. Aca el mapeo es identidad: cada local es el
# suyo. Si aparece uno que no esta, revienta (igual que el "venue desconocido" de USA).
VENUE_MAP = {b: b for b in [
    "Barcelona 1", "Barcelona 2", "Madrid", "Roma",
    "Málaga 1", "Málaga 3", "Valencia", "Alicante", "Granada",
]}

# Orden de aparicion en la tabla de detalle (propias primero, franquicias despues)
BRANCH_ORDER = ["Barcelona 1", "Barcelona 2", "Madrid", "Roma",
                "Málaga 1", "Málaga 3", "Valencia", "Alicante", "Granada"]

# Sucursales "PROPIAS" y "FRANQUICIAS" (clasificacion 2026)
PROPIAS = ["Barcelona 1", "Barcelona 2", "Madrid", "Roma"]
FRANQUICIAS = ["Málaga 1", "Málaga 3", "Valencia", "Alicante", "Granada"]

# Locales SIN historico 2025 (abrieron en 2026): no tienen contra que compararse.
# Se muestran con "—" y "s/ comp." y quedan FUERA del calculo de la variacion %
# (pero su venta 2026 SI suma al total, es plata real que entro).
SIN_HISTORICO_2025 = {"Madrid", "Alicante"}

# Dias de DESCANSO FIJO por local: {local: {dia_semana: vigente_desde}}.
# dia_semana = date.weekday() -> lunes=0 ... domingo=6.
# Ese dia, si el local no manda cierre, NO es un error: se registra en 0 y el
# reporte sale igual. Si lo manda (ej. abre un feriado), el PDF se suma normal.
# El "vigente_desde" evita que un rescate de un lunes VIEJO (cuando Granada si
# abria) tome por bueno un 0 que en realidad es un PDF que se perdio.
# Un solo lugar: lo leen fetch_cierres.py, report.py y report_semanal.py.
DIAS_CERRADO = {
    "Granada": {0: date(2026, 10, 5)},   # el encargado confirmo: los lunes ya no abre
}


# Cierres PUNTUALES (un dia concreto): ej. un feriado cae lunes, Granada abre
# el lunes y descansa el martes. Se cargan en data/descansos_puntuales.json,
# que se edita desde la web de GitHub sin tocar codigo:
#     "2026-10-13": ["Granada"]
DESCANSOS_PUNTUALES = Path(__file__).parent / "data" / "descansos_puntuales.json"
_puntuales_cache = None


def descansos_puntuales():
    """{date: set(locales)}. Valida fechas y nombres para que un typo corte claro."""
    global _puntuales_cache
    if _puntuales_cache is None:
        _puntuales_cache = {}
        if DESCANSOS_PUNTUALES.exists():
            data = json.loads(DESCANSOS_PUNTUALES.read_text(encoding="utf-8"))
            for k, locales in data.items():
                if k.startswith("_"):          # claves de ayuda/comentario
                    continue
                try:
                    f = date.fromisoformat(k)
                except ValueError:
                    raise SystemExit(f"[ERROR] descansos_puntuales.json: fecha invalida {k!r} "
                                     f"(formato AAAA-MM-DD).")
                malos = [l for l in locales if l not in BRANCH_ORDER]
                if malos:
                    raise SystemExit(f"[ERROR] descansos_puntuales.json ({k}): local invalido "
                                     f"{malos}. Validos: {', '.join(BRANCH_ORDER)}.")
                _puntuales_cache[f] = set(locales)
    return _puntuales_cache


def motivo_cierre(local, fecha):
    """'puntual' / 'fijo' si ese dia el local esta de descanso; None si debe abrir."""
    if local in descansos_puntuales().get(fecha, set()):
        return "puntual"
    desde = DIAS_CERRADO.get(local, {}).get(fecha.weekday())
    if desde is not None and fecha >= desde:
        return "fijo"
    return None


def cerrado_por_descanso(local, fecha):
    """True si 'fecha' es dia de descanso (fijo o puntual) del local."""
    return motivo_cierre(local, fecha) is not None

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}
DIAS_ES = {0: "LUNES", 1: "MARTES", 2: "MIÉRCOLES", 3: "JUEVES", 4: "VIERNES", 5: "SÁBADO", 6: "DOMINGO"}
MES_CORTO = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
             7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}


def _to_float(v):
    if v is None:
        return 0.0
    return float(str(v).replace(",", "").replace("$", "").strip() or 0)


def parse_excel_full(path):
    """Devuelve (fecha_ini, fecha_fin, ventas_por_sucursal, tickets_por_sucursal).

    Net Sales = columna C. Bill Count = columna F (los tickets del dia).
    parse_excel() quedo como wrapper de esta para no cambiarle la firma a nadie:
    el diario la sigue llamando igual y ni se entera de que ahora tambien hay
    tickets. Un solo parser, un solo VENUE_MAP, un solo lugar donde tocar.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    title = str(rows[0][0] or "")
    m = re.search(r"(\d{4}-\d{2}-\d{2})\s*/\s*(\d{4}-\d{2}-\d{2})", title)
    if not m:
        raise ValueError(f"No pude leer la fecha del titulo del Excel: {title!r}")
    fecha_ini = datetime.strptime(m.group(1), "%Y-%m-%d").date()
    fecha_fin = datetime.strptime(m.group(2), "%Y-%m-%d").date()

    ventas = {b: 0.0 for b in BRANCH_ORDER}
    tickets = {b: 0 for b in BRANCH_ORDER}
    for r in rows[2:]:
        name = str(r[0] or "").strip()
        if not name or name.upper().startswith("REPORT"):
            continue
        if name not in VENUE_MAP:
            raise ValueError(f"Venue desconocido en {path}: {name!r}. Agregalo a VENUE_MAP.")
        ventas[VENUE_MAP[name]] += _to_float(r[2])          # col C = Net Sales
        tickets[VENUE_MAP[name]] += int(_to_float(r[5]))    # col F = Bill Count
    return fecha_ini, fecha_fin, ventas, tickets


def parse_excel(path):
    """Devuelve (fecha_inicio, fecha_fin, dict_consolidado_por_sucursal) usando Net Sales (columna C)."""
    ini, fin, ventas, _ = parse_excel_full(path)
    return ini, fin, ventas


def load_accumulator(path):
    """Estado: {month: 'YYYY-MM', last_date: 'YYYY-MM-DD', acumulado: {sucursal: monto}}.
    Si no existe, arranca vacio (el reinicio mensual lo deja en cero al primer uso)."""
    p = Path(path)
    if not p.exists():
        return {"month": None, "last_date": None, "acumulado": {}}
    return json.loads(p.read_text(encoding="utf-8"))


def save_accumulator(path, data):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def money(v):
    # Sin decimales (redondeado) y con espacio DURO antes del €, para que el
    # simbolo nunca caiga a otro renglon. Formato europeo: 28.976&nbsp;€
    s = f"{round(v):,}".replace(",", ".")
    return f"{s}\u00a0€"

def money2(v):
    # Como money() pero CON 2 decimales. Se usa para el TICKET PROMEDIO, donde
    # los centavos importan (8,80 EUR no es lo mismo que 9 EUR).
    s = f"{v:,.2f}".replace(",", "·").replace(".", ",").replace("·", ".")
    return f"{s}\u00a0€"


def build_report(ventas_path, acum25_path, acum_state_path):
    fecha_ini_v, fecha, venta_dia = parse_excel(ventas_path)
    a25_ini, a25_fin, acum25 = parse_excel(acum25_path)

    # Validacion: ambos Excel deben referirse al mismo mes y terminar el mismo dia (comparacion espejo).
    # Ventas_ayer = un dia (fecha). Acumulado_interanual = rango del anio anterior, mismo mes, hasta el mismo dia.
    errores = []
    if a25_fin.year != fecha.year - 1:
        errores.append(
            f"El Acumulado_interanual parece NO ser del año anterior: termina en {a25_fin.year}, "
            f"se esperaba {fecha.year - 1}."
        )
    if a25_fin.month != fecha.month:
        errores.append(
            f"Los meses no coinciden: Ventas es de {MES_CORTO[fecha.month]} y "
            f"Acumulado_interanual es de {MES_CORTO[a25_fin.month]}. "
            f"Subí el acumulado {fecha.year - 1} del mismo mes."
        )
    if a25_fin.day != fecha.day:
        errores.append(
            f"Los días de corte no coinciden: Ventas es del día {fecha.day} y "
            f"el Acumulado_interanual llega hasta el día {a25_fin.day}. "
            f"Deben terminar el mismo día para comparar período espejo."
        )

    if errores:
        gh_out = os.environ.get("GITHUB_OUTPUT")
        if gh_out:
            with open(gh_out, "a", encoding="utf-8") as f:
                f.write("send=false\n")
        print("[ERROR DE VALIDACION] No se envía el mail:")
        for e in errores:
            print("  - " + e)
        sys.exit(1)

    state = load_accumulator(acum_state_path)
    mes_actual = fecha.strftime("%Y-%m")  # ej. "2026-06"

    # Proteccion anti doble-conteo (misma fecha ya procesada)
    if state.get("last_date") == fecha.isoformat():
        gh_out = os.environ.get("GITHUB_OUTPUT")
        if gh_out:
            with open(gh_out, "a", encoding="utf-8") as f:
                f.write("send=false\n")
        print(f"[SKIP] La fecha {fecha} ya fue procesada. No se suma de nuevo ni se envia mail.")
        sys.exit(0)

    # Reinicio mensual: si cambio el mes (o es la primera corrida), el acumulado arranca en cero.
    if state.get("month") != mes_actual:
        print(f"[MES NUEVO] {state.get('month')} -> {mes_actual}. Acumulado reiniciado a cero.")
        acum_prev = {b: 0.0 for b in BRANCH_ORDER}
    else:
        acum_prev = state.get("acumulado", {})

    # Acum.26 = previo (del mes en curso) + venta del dia
    acum26 = {b: round(acum_prev.get(b, 0.0) + venta_dia[b], 2) for b in BRANCH_ORDER}

    rows = []
    for b in BRANCH_ORDER:
        d = venta_dia[b]
        a26 = acum26[b]
        comparable = b not in SIN_HISTORICO_2025
        if comparable:
            a25 = acum25[b]
            diff = a26 - a25
            pct = (diff / a25 * 100) if a25 else 0.0
        else:
            # Local nuevo: no existia en 2025. No hay variacion posible.
            a25 = None
            diff = None
            pct = None
        rows.append({"branch": b, "dia": d, "a26": a26, "a25": a25,
                     "diff": diff, "pct": pct, "comparable": comparable,
                     # Descanso fijo y sin venta: se muestra "cerrado" en vez de "0 €"
                     "cerrado": d == 0 and cerrado_por_descanso(b, fecha)})

    def _agregar(grupo):
        """Suma un grupo de locales. La VENTA (dia/a26/a25) es real: incluye a los
        nuevos. La VARIACION se calcula SOLO sobre comparables (apples-to-apples),
        asi Madrid/Alicante no inflan el %. 'hay_nuevo' prende la nota al pie."""
        elegidos = [r for r in rows if r["branch"] in grupo]
        comp = [r for r in elegidos if r["comparable"]]
        s = {
            "dia": sum(r["dia"] for r in elegidos),
            "a26": sum(r["a26"] for r in elegidos),
            "a25": sum((r["a25"] or 0.0) for r in elegidos),
            "hay_nuevo": any(not r["comparable"] for r in elegidos),
        }
        comp_a26 = sum(r["a26"] for r in comp)
        comp_a25 = sum(r["a25"] for r in comp)
        s["ca26"], s["ca25"] = round(comp_a26, 2), round(comp_a25, 2)
        s["diff"] = comp_a26 - comp_a25
        s["pct"] = (s["diff"] / comp_a25 * 100) if comp_a25 else 0.0
        return s

    totals = _agregar(BRANCH_ORDER)

    def subtotal(grupo):
        return _agregar(grupo)

    propias = subtotal(PROPIAS)
    franquicias = subtotal(FRANQUICIAS)

    # Persistir el nuevo acumulado, el mes activo y la fecha procesada
    new_state = {"month": mes_actual, "last_date": fecha.isoformat(), "acumulado": acum26}

    # Mail (diseño "combinado"): lo arma mail_diario.py. Ya no lleva graficos
    # PNG: barras y graficos son tablas HTML (se ven aunque el cliente de mail
    # bloquee imagenes). El unico adjunto es el logo.
    # Tickets del dia: de Ventas_ayer.xlsx (columna Bill Count).
    import mail_diario
    _, _, _, tickets_dia = parse_excel_full(ventas_path)
    base_dir = Path(acum_state_path).parent.parent
    html = mail_diario.render(fecha, rows, totals, propias, franquicias, tickets_dia, base_dir)
    chart_paths = []
    return html, new_state, fecha, totals, chart_paths


# El HTML del mail lo arma mail_diario.py (diseño "combinado", oct/2026).
# La version anterior (render_html + graficos PNG de charts.py) quedo en el
# historial de git; charts.py lo siguen usando el semanal y el cierre mensual.


if __name__ == "__main__":
    base = Path(__file__).parent
    ventas = base / "Ventas_ayer.xlsx"
    acum25 = base / "Acumulado_interanual.xlsx"
    state = base / "data" / "acumulado.json"

    html, new_state, fecha, totals, chart_paths = build_report(ventas, acum25, state)

    (base / "preview.html").write_text(html, encoding="utf-8")
    save_accumulator(state, new_state)

    # Asunto: "Reporte Ventas DD/MM/YYYY - Lucciano's USA"
    subject = f"Reporte Ventas {fecha.strftime('%d/%m/%Y')} - Lucciano's Europa"

    # Exponer salidas al workflow de GitHub Actions
    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as f:
            f.write("send=true\n")
            f.write(f"subject={subject}\n")
            f.write(f"date={fecha.isoformat()}\n")

    print(f"OK - fecha {fecha} | venta dia {money(totals['dia'])} | acum26 {money(totals['a26'])} | var {totals['pct']:.1f}%")
