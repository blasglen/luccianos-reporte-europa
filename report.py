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
from datetime import datetime, timedelta

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
    # Formato europeo: 28.975,69 €  (punto de miles, coma decimal, simbolo al final)
    s = f"{v:,.2f}".replace(",", "·").replace(".", ",").replace("·", ".")
    return f"{s} €"


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
                     "diff": diff, "pct": pct, "comparable": comparable})

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

    # Graficos (PNG para incrustar via CID)
    from charts import build_charts
    mes = MES_CORTO[fecha.month]
    a26_lbl = f"Acum. {mes}/{str(fecha.year)[2:]}"
    a25_lbl = f"Acum. {mes}/{str(fecha.year - 1)[2:]}"
    base_dir = Path(acum_state_path).parent.parent
    chart_paths = build_charts(rows, totals, mes, a26_lbl, a25_lbl, out_dir=str(base_dir / "charts"))

    html = render_html(fecha, rows, totals, propias, franquicias)
    return html, new_state, fecha, totals, chart_paths


def _pct_html(pct, diff):
    color = "#2e7d32" if pct >= 0 else "#c62828"
    sign = "+" if pct >= 0 else ""
    diff_str = f"({sign}{money(diff)})" if diff < 0 else f"(+{money(diff)})"
    # diff puede ser negativo; money ya pone el signo? No: money no pone signo de resta para negativos formateados con :,.2f -> si lo pone
    diff_str = f"({money(diff)})" if diff < 0 else f"(+{money(diff)})"
    return f'<span style="color:{color};font-weight:700;">{sign}{pct:.1f}%</span><br><span style="color:{color};font-size:12px;">{diff_str}</span>'


def render_html(fecha, rows, totals, propias, franquicias):
    mes = MES_CORTO[fecha.month]
    anio_corto = str(fecha.year)[2:]
    anio_ant = str(fecha.year - 1)[2:]
    fecha_larga = f"{DIAS_ES[fecha.weekday()]} {fecha.day} DE {MESES_ES[fecha.month]} DE {fecha.year}"
    a26_lbl = f"ACUM. {mes.upper()}/{anio_corto}"
    a25_lbl = f"ACUM. {mes.upper()}/{anio_ant}"

    def chip(pct, diff, nota=False):
        up = pct >= 0
        col = "#1a7d2e" if up else "#c62828"
        bg = "#eaf5ec" if up else "#fbecec"
        s = "+" if up else ""
        ast = "*" if nota else ""
        dd = f"(+{money(diff)})" if diff >= 0 else f"({money(diff)})"
        return (f'<span style="display:inline-block;background:{bg};color:{col};'
                f'font-weight:700;font-size:12px;padding:3px 9px;border-radius:20px;white-space:nowrap;">'
                f'{s}{pct:.1f}%{ast}</span>'
                f'<div style="color:{col};font-size:11px;margin-top:3px;">{dd}</div>')

    def chip_sc():
        # "sin comparable": local nuevo, no existia en 2025. Cartelito gris neutro.
        return ('<span style="display:inline-block;background:#eeeeee;color:#8a8a8a;'
                'font-weight:700;font-size:12px;padding:3px 9px;border-radius:20px;white-space:nowrap;">'
                's/ comp.</span>')

    def fila(r, zebra):
        if r["comparable"]:
            a25_cell = money(r["a25"])
            var_cell = chip(r["pct"], r["diff"])
        else:
            a25_cell = "—"
            var_cell = chip_sc()
        return f"""
        <tr style="background:{zebra};">
          <td style="padding:14px 18px;font-weight:700;color:#111111;font-size:14px;">{r['branch']}</td>
          <td style="padding:14px 12px;text-align:right;color:#111111;font-size:14px;">{money(r['dia'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#111111;font-weight:700;font-size:14px;">{money(r['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;color:#9a9a9a;font-size:14px;">{a25_cell}</td>
          <td style="padding:14px 18px;text-align:right;">{var_cell}</td>
        </tr>"""

    def encabezado_grupo(nombre):
        return f"""
        <tr style="background:#eef1f6;">
          <td colspan="5" style="padding:9px 18px;color:#1f3a6e;font-size:10px;font-weight:800;letter-spacing:2px;">{nombre}</td>
        </tr>"""

    def fila_subtotal(nombre, s):
        return f"""
        <tr style="background:#f5f5f5;border-top:1px solid #e2e2e2;">
          <td style="padding:14px 18px;font-weight:800;color:#1f3a6e;font-size:13px;">{nombre}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#1f3a6e;font-size:13px;">{money(s['dia'])}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#1f3a6e;font-size:13px;">{money(s['a26'])}</td>
          <td style="padding:14px 12px;text-align:right;font-weight:800;color:#7a8aa5;font-size:13px;">{money(s['a25'])}</td>
          <td style="padding:14px 18px;text-align:right;">{chip(s['pct'], s['diff'], nota=s.get('hay_nuevo'))}</td>
        </tr>"""

    # Nota al pie automatica: aparece sola si algun local no tiene comparable 2025.
    nuevos = [r["branch"] for r in rows if not r["comparable"]]
    if nuevos:
        lista = " y ".join(nuevos) if len(nuevos) <= 2 else ", ".join(nuevos[:-1]) + " y " + nuevos[-1]
        nota_pie = (f'\n    <div style="color:#9a9a9a;font-size:11px;margin-top:12px;line-height:1.5;">'
                    f'* La variación se calcula solo sobre locales con histórico 2025. '
                    f'{lista} abrieron en 2026, no tienen comparable (su venta sí suma al total).</div>')
    else:
        nota_pie = ""

    by_name = {r["branch"]: r for r in rows}
    cuerpo = ""
    cuerpo += encabezado_grupo("PROPIAS")
    for i, b in enumerate(PROPIAS):
        cuerpo += fila(by_name[b], "#ffffff" if i % 2 == 0 else "#fafafa")
    cuerpo += fila_subtotal("Subtotal Propias", propias)
    cuerpo += encabezado_grupo("FRANQUICIAS")
    for i, b in enumerate(FRANQUICIAS):
        cuerpo += fila(by_name[b], "#ffffff" if i % 2 == 0 else "#fafafa")
    cuerpo += fila_subtotal("Subtotal Franquicias", franquicias)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#eeeeee;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eeeeee;">
<tr><td align="center" style="padding:24px 12px;">

<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:600px;max-width:600px;background:#ffffff;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,0.12);">

  <!-- HEADER -->
  <tr><td style="background:#000000;padding:38px 32px 32px 32px;text-align:center;">
    <img src="cid:logo" alt="Lucciano's" width="190" style="display:block;margin:0 auto;max-width:190px;height:auto;">
    <div style="color:#bdbdbd;font-size:12px;letter-spacing:4px;margin-top:18px;">REPORTE DE VENTAS DIARIO</div>
    <div style="color:#ffffff;font-size:13px;font-weight:700;letter-spacing:2px;margin-top:14px;">{fecha_larga}</div>
  </td></tr>

  <!-- KPIs (alturas igualadas con height fijo en el contenido) -->
  <tr><td style="padding:30px 32px 6px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;">CONSOLIDADO · 9 SUCURSALES</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:16px;">
      <tr>
        <td width="33%" style="padding-right:7px;vertical-align:top;">
          <div style="background:#111111;border-radius:12px;padding:20px;height:118px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">VENTA DEL DÍA</div>
            <div style="color:#ffffff;font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['dia'])}</div>
            <div style="color:#777777;font-size:11px;margin-top:6px;">9 sucursales</div>
          </div>
        </td>
        <td width="34%" style="padding:0 7px;vertical-align:top;">
          <div style="background:#111111;border-radius:12px;padding:20px;height:118px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">{a26_lbl}</div>
            <div style="color:#ffffff;font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['a26'])}</div>
            <div style="color:#777777;font-size:11px;margin-top:6px;">mes en curso</div>
          </div>
        </td>
        <td width="33%" style="padding-left:7px;vertical-align:top;">
          <div style="background:#f5f5f5;border-radius:12px;padding:20px;height:118px;">
            <div style="color:#9a9a9a;font-size:10px;letter-spacing:1px;">{a25_lbl}</div>
            <div style="color:#111111;font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px;">{money(totals['a25'])}</div>
            <div style="margin-top:8px;">{chip(totals['pct'], totals['diff'])}</div>
          </div>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- PROGRESO -->
  <tr><td style="padding:22px 32px 6px 32px;">
    <div style="background:#fafafa;border-radius:12px;padding:20px 22px;">
      <div style="color:#9a9a9a;font-size:11px;letter-spacing:2px;margin-bottom:6px;">AVANCE DEL MES vs AÑO ANTERIOR</div>
      <img src="cid:progreso" alt="Avance del mes" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
    </div>
  </td></tr>

  <!-- GRAFICO COMPARATIVO -->
  <tr><td style="padding:22px 32px 6px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:12px;">COMPARATIVO POR SUCURSAL · {anio_corto} vs {anio_ant}</div>
    <img src="cid:comparativo" alt="Comparativo por sucursal" width="536" style="display:block;width:100%;max-width:536px;height:auto;">
  </td></tr>

  <!-- DETALLE -->
  <tr><td style="padding:24px 32px 36px 32px;">
    <div style="color:#9a9a9a;font-size:11px;letter-spacing:3px;margin-bottom:14px;">DETALLE POR SUCURSAL</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-radius:12px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,0.06);">
      <thead>
        <tr style="background:#111111;">
          <th style="padding:13px 18px;text-align:left;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">SUCURSAL</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">DÍA</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">{a26_lbl}</th>
          <th style="padding:13px 12px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">{a25_lbl}</th>
          <th style="padding:13px 18px;text-align:right;color:#ffffff;font-size:10px;letter-spacing:1px;font-weight:700;">VARIACIÓN</th>
        </tr>
      </thead>
      <tbody>{cuerpo}
        <tr style="background:#111111;">
          <td style="padding:17px 18px;font-weight:800;color:#ffffff;font-size:14px;">TOTAL GENERAL</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['dia'])}</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['a26'])}</td>
          <td style="padding:17px 12px;text-align:right;font-weight:800;color:#ffffff;font-size:14px;">{money(totals['a25'])}</td>
          <td style="padding:17px 18px;text-align:right;">{chip(totals['pct'], totals['diff'], nota=totals.get('hay_nuevo'))}</td>
        </tr>
      </tbody>
    </table>{nota_pie}
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="background:#000000;padding:20px 32px;text-align:center;">
    <div style="color:#777777;font-size:11px;letter-spacing:1px;">LUCCIANO'S EUROPA · Reporte automático generado el {fecha.strftime('%d/%m/%Y')}</div>
  </td></tr>

</table>

</td></tr>
</table>
</body>
</html>"""
    return html


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
