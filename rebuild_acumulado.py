"""
rebuild_acumulado.py
--------------------
Reconstruye data/acumulado.json desde cero a partir del historial del mes.

Para que sirve:
  El acumulado del diario es INCREMENTAL (acum_prev + venta del dia) y su unica
  proteccion es no procesar dos veces la misma fecha. NO valida que los dias
  vengan consecutivos, asi que si un dia del medio no se proceso (ej. Roma no
  mando el PDF y el fetch aborto), el acumulado avanza igual salteando el hueco
  y queda subestimado sin que se note.

  Cuando se backfillea ese hueco en el historial, este script recalcula el
  acumulado del mes sumando TODOS los dias presentes del historial. Deja
  last_date = ultimo dia presente del mes. Asi el semanal vuelve a cerrar y la
  conciliacion (historial vs acumulado) da al centavo.

Uso: python rebuild_acumulado.py 2026-09
"""
import json
import sys
from pathlib import Path

from report import BRANCH_ORDER

BASE = Path(__file__).parent


def _v(x):
    """El historial viejo guardaba {suc: float}; el nuevo {suc: {venta, tickets}}."""
    return x["venta"] if isinstance(x, dict) else x


def main(mes):  # mes = "YYYY-MM"
    anio = int(mes.split("-")[0])
    path_hist = BASE / "data" / f"historico_{anio}.json"
    if not path_hist.exists():
        raise SystemExit(f"[ERROR] No existe {path_hist.name}.")
    hist = json.loads(path_hist.read_text(encoding="utf-8"))

    dias = sorted(k for k in hist if k.startswith(mes))
    if not dias:
        raise SystemExit(f"[ERROR] No hay dias de {mes} en el historial.")

    acum = {b: 0.0 for b in BRANCH_ORDER}
    for k in dias:
        for b in BRANCH_ORDER:
            acum[b] = round(acum[b] + _v(hist[k].get(b, 0.0)), 2)

    estado = {"month": mes, "last_date": dias[-1], "acumulado": acum}
    (BASE / "data" / "acumulado.json").write_text(
        json.dumps(estado, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"[OK] Acumulado reconstruido: {mes} | {len(dias)} dias "
          f"({dias[0]} .. {dias[-1]}) | total {sum(acum.values()):,.2f} EUR")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Uso: python rebuild_acumulado.py YYYY-MM")
    sys.exit(main(sys.argv[1]))
