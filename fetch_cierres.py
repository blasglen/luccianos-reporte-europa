"""
fetch_cierres.py  (Europa)
--------------------------
Reemplaza a fetch_touchbistro.py. En USA llegaba UN Excel con todos los locales;
en Europa llegan MUCHOS PDF (uno por punto de venta) en mails separados de
informatica@luccianos.com.ar. Este script:

  1. Entra por IMAP a la casilla que RECIBE (reportesluccianos@gmail.com).
  2. Junta TODOS los mails del remitente de los ultimos 2 dias.
  3. Extrae cada PDF adjunto y lo parsea (bilingue ES/IT: Roma viene en italiano).
  4. Se queda con los cierres cuyo DIA DE NEGOCIO == ayer.
     -> El dia NO se saca del asunto ni del nombre del PDF (mienten: los locales
        que cierran pasada la medianoche quedan estampados con la fecha del dia
        siguiente). El dia real es el de la linea "INICIO DE CAJA" / "APERTURA DI
        CASSA", que en los 17 cierres del 29/07 dice 29/07 en todos.
  5. Consolida los ~17 PDF en los 9 locales sumando por el campo "Sucursal:".
  6. Convierte BRUTO -> NETO dividiendo por 1.10 (IVA 10%).
  7. Deja Ventas_ayer.xlsx con UNA fila por local (col A = local, col C = neto,
     col F = tickets), que es el mismo formato que espera report.py. Asi el resto
     del sistema no se entera de que la fuente cambio.

Falla RUIDOSA (exit 1) si: aparece un local que no esta en la lista blanca, o si
a un local esperado le falta el cierre del dia. Preferimos cortar en rojo antes
que mandar un total incompleto.

Secrets (GitHub Actions), igual esquema que USA:
  IMAP_USER      -> reportesluccianos@gmail.com
  IMAP_APP_PASS  -> App Password de 16 caracteres de ESA casilla
"""
import email
import imaplib
import io
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

import openpyxl
import pdfplumber

SENDER = os.environ.get("EU_SENDER", "informatica@luccianos.com.ar")
SALIDA = Path(__file__).parent / "Ventas_ayer.xlsx"
IVA = 0.10  # neto = bruto / (1 + IVA). Un solo lugar donde vive la alicuota.

# --- Locales ---------------------------------------------------------------
# Lo que dice el campo "Sucursal:" del PDF  ->  nombre canonico (el del Master
# 2025 y el que se muestra en el mail). Los PDF dicen "Barcelona"/"Malaga";
# el negocio los llama "Barcelona 1"/"Malaga 1".
DISPLAY = {
    "Barcelona": "Barcelona 1",
    "Barcelona 2": "Barcelona 2",
    "Madrid": "Madrid",
    "Roma": "Roma",
    "Malaga": "Málaga 1",
    "Malaga 3": "Málaga 3",
    "Valencia": "Valencia",
    "Alicante": "Alicante",
    "Granada": "Granada",
}
# Los 9 que TIENEN que estar todos los dias. Si falta uno, corta en rojo.
LOCALES_ESPERADOS = set(DISPLAY.values())

# Regex bilingues (ES / IT).
RE_SUC = re.compile(r"Sucursal:\s*(.+?)\s+Punto (?:de venta|vendita)\s+(\d+)")
RE_DIA = re.compile(r"(?:INICIO DE CAJA|APERTURA DI CASSA)\s+(\d{2}/\d{2}/\d{4})")
RE_VENTA = re.compile(r"(?:TOTAL VENTA BRUTA|TOTALE VENDITA LORDA)\s+([\-\d.]+)")
RE_TICK = re.compile(r"(?:FACTURAS B|FATTURE B)\s+[\-\d.]+\s+(\d+)")


def parse_pdf(data):
    """bytes de un PDF -> dict del cierre, o None si no se pudo leer la cabecera."""
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        txt = "\n".join((p.extract_text() or "") for p in pdf.pages)

    suc = pdv = dia = None
    bruto = 0.0
    tickets = 0
    for ln in txt.splitlines():
        m = RE_SUC.search(ln)
        if m:
            suc, pdv = m.group(1).strip(), m.group(2)
        m = RE_DIA.search(ln)
        if m and dia is None:
            dia = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        m = RE_VENTA.search(ln)
        if m:
            bruto = float(m.group(1))          # TOTAL VENTA BRUTA ya = facturas - notas de credito
        m = RE_TICK.search(ln)
        if m:
            tickets = int(m.group(1))          # cantidad de FACTURAS B = tickets
    if suc is None or dia is None:
        return None
    if suc not in DISPLAY:
        raise ValueError(f"Local desconocido en un PDF: {suc!r} (PDV {pdv}). "
                         f"Agregalo a DISPLAY o revisa el cierre.")
    return {"local": DISPLAY[suc], "pdv": pdv, "dia": dia,
            "bruto": bruto, "tickets": tickets}


def consolidar(cierres, dia_objetivo):
    """Suma los PDF del dia objetivo por local, deduplicando por (pdv, dia)."""
    vistos = set()
    neto = defaultdict(float)
    tickets = defaultdict(int)
    usados = 0
    for c in cierres:
        if c["dia"] != dia_objetivo:
            continue
        clave = (c["pdv"], c["dia"])          # candado: un terminal, un dia, una vez
        if clave in vistos:
            continue
        vistos.add(clave)
        neto[c["local"]] += c["bruto"] / (1 + IVA)
        tickets[c["local"]] += c["tickets"]
        usados += 1

    faltan = LOCALES_ESPERADOS - set(neto.keys())
    if faltan:
        raise SystemExit(
            f"[ERROR] Faltan cierres del {dia_objetivo} para: {', '.join(sorted(faltan))}. "
            f"No genero Ventas_ayer.xlsx (prefiero cortar antes que mandar un total incompleto)."
        )
    return neto, tickets, usados


def escribir_xlsx(dia, neto, tickets):
    """Deja Ventas_ayer.xlsx en el formato que ya lee report.py:
    fila 0 = titulo con el rango de fechas (mismo dia dos veces, como el diario de USA),
    fila 1 = encabezados, luego una fila por local (A=local, C=neto, F=tickets)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    d = dia.strftime("%Y-%m-%d")
    ws["A1"] = f"Ventas {d} / {d}"
    ws.append(["Sucursal", "", "Net Sales", "", "", "Bill Count"])
    for local in sorted(neto):
        ws.append([local, "", round(neto[local], 2), "", "", tickets[local]])
    wb.save(SALIDA)


def main():
    user = os.environ["IMAP_USER"]
    pwd = os.environ["IMAP_APP_PASS"]
    # Dia de negocio = ayer. Corremos 7:00 AR; a esa hora los cierres de anoche
    # (Europa va varias horas adelantada) ya entraron todos.
    dia_objetivo = (datetime.now(timezone.utc) - timedelta(days=1)).date()

    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(user, pwd)
    M.select("INBOX")
    since = (datetime.utcnow() - timedelta(days=2)).strftime("%d-%b-%Y")
    typ, data = M.search(None, f'(FROM "{SENDER}" SINCE {since})')
    ids = data[0].split()
    if not ids:
        print(f"[ERROR] No hay mails de {SENDER} en los ultimos 2 dias.")
        return 1

    cierres = []
    for mid in ids:
        typ, msgdata = M.fetch(mid, "(RFC822)")
        msg = email.message_from_bytes(msgdata[0][1])
        for part in msg.walk():
            fn = part.get_filename()
            if fn and fn.lower().endswith(".pdf"):
                c = parse_pdf(part.get_payload(decode=True))
                if c:
                    cierres.append(c)
    M.logout()

    neto, tickets, usados = consolidar(cierres, dia_objetivo)
    escribir_xlsx(dia_objetivo, neto, tickets)
    total = sum(neto.values())
    print(f"[OK] {usados} PDF -> {len(neto)} locales | dia {dia_objetivo} | "
          f"neto total {total:,.2f} EUR -> {SALIDA.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
