"""
generar_acum_ant.py  (Europa)
-----------------------------
Comparativo del AÑO ANTERIOR para CUALQUIER rango de fechas, leyendo el master
diario data/Ventas_Master_2025.xlsx (hoja "Ventas 2025", columnas
Local | Fecha | Venta Bruta (EUR) | Tickets).

Version generalizada de generar_acum2025.py:
  - El semanal pide "del lunes al domingo"  -> rango.
  - El cierre pide "del 1ro al ultimo del mes" -> rango.
Mismo calculo, distintos limites.

CRITERIO DE ESPEJO: fechas calendario (mismo dia/mes, año - 1). NO alinea por dia
de semana (decision de negocio, documentada).

Venta en NETO (bruto del master ÷ 1.10). Locales sin datos en el rango (Madrid,
Alicante en 2025) simplemente no aparecen; report.py los toma como "sin comparable".

Falla RUIDOSA si al master le falta algun dia del rango.
"""
import argparse
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from collections import defaultdict

import openpyxl

BASE = Path(__file__).parent
MASTER = BASE / "data" / "Ventas_Master_2025.xlsx"
IVA = 0.10


def _as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()


def espejo(d):
    """Mismo dia y mes, año anterior."""
    return date(d.year - 1, d.month, d.day)


def _leer_master():
    wb = openpyxl.load_workbook(MASTER, data_only=True)
    ws = wb["Ventas 2025"]
    h = [str(c) for c in next(ws.iter_rows(values_only=True))]
    cl, cf, cv = h.index("Local"), h.index("Fecha"), h.index("Venta Bruta (EUR)")
    por_dia = defaultdict(dict)  # fecha -> {local: bruto}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[cl] is None:
            continue
        f = _as_date(r[cf])
        por_dia[f][str(r[cl]).strip()] = float(r[cv] or 0)
    return por_dia


def acumular_rango(desde, hasta):
    """Suma NETA por local entre [desde, hasta] (inclusive). Falla si falta un dia."""
    por_dia = _leer_master()
    acum = defaultdict(float)
    faltan = []
    d = desde
    while d <= hasta:
        if d not in por_dia:
            faltan.append(d.isoformat())
        else:
            for loc, bruto in por_dia[d].items():
                acum[loc] += round(bruto / (1 + IVA), 2)
        d += timedelta(days=1)
    if faltan:
        raise SystemExit(
            f"[ERROR] Al master 2025 le faltan dias del rango {desde}..{hasta}: "
            f"{', '.join(faltan[:8])}{'...' if len(faltan) > 8 else ''}. "
            f"No genero el comparativo con huecos."
        )
    return dict(acum)


def escribir_excel(desde, hasta, acum, salida):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = f"Ventas {desde.isoformat()} / {hasta.isoformat()}"
    ws.append(["Sucursal", "", "Net Sales", "", "", "Bill Count"])
    for loc in sorted(acum):
        ws.append([loc, "", round(acum[loc], 2), "", "", 0])
    wb.save(salida)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--desde", required=True)
    ap.add_argument("--hasta", required=True)
    ap.add_argument("--salida", required=True)
    a = ap.parse_args()
    desde = datetime.strptime(a.desde, "%Y-%m-%d").date()
    hasta = datetime.strptime(a.hasta, "%Y-%m-%d").date()
    acum = acumular_rango(desde, hasta)
    escribir_excel(desde, hasta, acum, a.salida)
    print(f"OK - comparativo {desde}..{hasta} (neto) | {len(acum)} locales | "
          f"total {sum(acum.values()):,.2f} EUR")
