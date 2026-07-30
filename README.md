# Lucciano's Europa — Reportes de Ventas

Reporte diario automático por mail sobre los 9 locales de Lucciano's Europa,
comparando el mes en curso (2026) contra el mismo período del año anterior (2025).

| Reporte | Cuándo | Período | Destinatario |
|---|---|---|---|
| **Diario** | Todos los días 7:00 AR | El día de ayer + acumulado del mes | `MAIL_TO` |

> El semanal y el cierre están pendientes de adaptar (reusan `fetch_cierres.py`
> pero necesitan su propio backfill de histórico mensual). No dispares esos crons todavía.

---

## Cómo se dispara

cron-job.org le pega por API a
`POST https://api.github.com/repos/blasglen/luccianos-reporte-europa/dispatches`
con `event_type: reporte-diario-europa`. Si se cae, se corre a mano desde Actions
(`workflow_dispatch` habilitado).

Corre 7:00 AR: Europa va varias horas adelantada, así que a esa hora ya entraron
los ~17 cierres de anoche.

---

## Flujo de datos (la gran diferencia con USA)

En USA llegaba UN Excel de TouchBistro. Acá llegan **~17 PDF** ("Cierre de Caja"),
uno por punto de venta, en mails separados de `informatica@luccianos.com.ar`.

```
        ~17 mails de informatica@luccianos.com.ar (PDF por punto de venta)
                            |
                   fetch_cierres.py  (IMAP + pdfplumber)
                   - junta todos los PDF del día
                   - agrupa por campo "Sucursal:" -> 9 locales
                   - dedup por (punto de venta, día de INICIO DE CAJA)
                   - BRUTO -> NETO (÷ 1.10, IVA 10%)
                            |
                     Ventas_ayer.xlsx   (mismo formato que esperaba report.py)
                            |
            +---------------+---------------+
      historico.py                      report.py
            |                               |
 data/historico_2026.json          data/acumulado.json
                                           |
                                     report.py (diario)
```

**El día de negocio NO se saca del asunto ni del nombre del PDF** (mienten: los
locales que cierran pasada la medianoche quedan estampados con la fecha del día
siguiente). Se saca de la línea `INICIO DE CAJA` / `APERTURA DI CASSA`. Roma viene
en italiano y el parser es bilingüe.

El comparativo 2025 sale de `data/Ventas_Master_2025.xlsx` (7 locales; también en
bruto, se pasa a neto ÷1.10) vía `generar_acum2025.py`.

---

## Locales

**Propias (4):** Barcelona 1, Barcelona 2, Madrid, Roma
**Franquicias (5):** Málaga 1, Málaga 3, Valencia, Alicante, Granada

Cada local puede tener varios puntos de venta (ej. Barcelona 2 = PDV 801/802/803/804);
`fetch_cierres.py` los suma por el campo `Sucursal:`. Si aparece un local que no está
en la lista blanca, **revienta** en vez de ignorarlo.

**Madrid y Alicante abrieron en 2026**: no tienen histórico 2025. Se muestran con
"—" y chip gris `s/ comp.`, y quedan FUERA del cálculo de la variación % (pero su
venta SÍ suma al total). Una nota al pie automática lo aclara.

---

## Archivos

| Archivo | Qué hace |
|---|---|
| `fetch_cierres.py` | Baja los PDF por IMAP, los parsea (bilingüe), consolida por local, pasa a neto y deja `Ventas_ayer.xlsx` |
| `report.py` | Reporte diario. Formato €, lógica de "sin comparable", 9 locales |
| `historico.py` | Registra el día en `data/historico_<año>.json` |
| `generar_acum2025.py` | Comparativo del año anterior desde el Master europeo (neto) |
| `charts.py` | Gráficos (€) |
| `send_mail.py` | Envía por SMTP. Remitente "Lucciano's Europa" |

### Secrets
`IMAP_USER` / `IMAP_APP_PASS` (casilla que **recibe**: reportesluccianos@gmail.com) ·
`GMAIL_USER` / `GMAIL_APP_PASS` (casilla que **envía**: contabilidad@luccianos.com.ar) ·
`MAIL_TO` (destinatario).

---

## Criterios de negocio

- **Venta = neto** (bruto de los cierres ÷ 1.10, IVA 10%).
- **Comparación interanual: mismas fechas calendario.**
- **Ticket = cantidad de FACTURAS B** por local.
- Locales sin histórico 2025 → "—" y fuera del %.
