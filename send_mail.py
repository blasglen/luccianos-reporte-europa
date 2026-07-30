"""
Envia el reporte HTML por Gmail via SMTP usando un App Password.
Incrusta imagenes inline (Content-ID) referenciadas en el HTML como cid:<nombre>:
  - cid:logo        -> Logo.png
  - cid:comparativo -> charts/comparativo.png
  - cid:progreso    -> charts/progreso.png

Variables de entorno (inyectadas por GitHub Actions desde Secrets):
  GMAIL_USER      -> casilla emisora (la cuenta gmail que envia)
  GMAIL_APP_PASS  -> App Password de 16 caracteres (NO la clave normal)
  MAIL_TO         -> destinatario(s). Si hay varios, separados por coma.
                     El PRIMERO va en 'Para' (To); el resto en copia (CC).
"""
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from pathlib import Path

# (cid, ruta)
INLINE_IMAGES = [
    ("logo", "Logo.png"),
    ("comparativo", "charts/comparativo.png"),
    ("progreso", "charts/progreso.png"),
]


def _attach_inline(root, cid, path):
    p = Path(path)
    if not p.exists():
        print(f"[AVISO] No encontre {path}; el mail sale sin esa imagen (cid:{cid}).")
        return
    img = MIMEImage(p.read_bytes())
    img.add_header("Content-ID", f"<{cid}>")
    img.add_header("Content-Disposition", "inline", filename=p.name)
    root.attach(img)


def send(subject, html, to=None, imagenes=None, mail_to_env="MAIL_TO"):
    """imagenes: lista de (cid, path). Si es None usa INLINE_IMAGES (el diario).
    mail_to_env: que Secret leer para los destinatarios. El semanal usa
    MAIL_TO_SOCIOS, asi que el diario y el semanal nunca comparten lista."""
    user = os.environ["GMAIL_USER"]
    pwd = os.environ["GMAIL_APP_PASS"]
    raw = to or os.environ[mail_to_env]

    destinatarios = [d.strip() for d in raw.split(",") if d.strip()]
    if not destinatarios:
        raise ValueError("MAIL_TO no tiene ninguna direccion valida.")

    to_addr = destinatarios[0]
    cc_addrs = destinatarios[1:]
    all_rcpts = destinatarios

    root = MIMEMultipart("related")
    root["Subject"] = subject
    root["From"] = f"Lucciano's Europa <{user}>"
    root["To"] = to_addr
    if cc_addrs:
        root["Cc"] = ", ".join(cc_addrs)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText("Tu cliente no soporta HTML. Abri el reporte en un cliente compatible.", "plain", "utf-8"))
    alt.attach(MIMEText(html, "html", "utf-8"))
    root.attach(alt)

    for cid, path in (imagenes if imagenes is not None else INLINE_IMAGES):
        _attach_inline(root, cid, path)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(user, pwd)
        s.sendmail(user, all_rcpts, root.as_string())
    print(f"Mail enviado. Para: {to_addr} | CC: {', '.join(cc_addrs) if cc_addrs else '(ninguno)'}")


if __name__ == "__main__":
    # Uso:
    #   python send_mail.py "Asunto" preview.html
    #   python send_mail.py "Asunto" preview_semanal.html --to-env MAIL_TO_SOCIOS \
    #          logo=Logo.png comparativo=charts/sem_comparativo.png
    # Sin extras se comporta EXACTAMENTE como antes (el diario no se entera).
    subject = sys.argv[1] if len(sys.argv) > 1 else "Reporte de Ventas"
    path = sys.argv[2] if len(sys.argv) > 2 else "preview.html"

    extras = sys.argv[3:]
    to_env = "MAIL_TO"
    imagenes = []
    i = 0
    while i < len(extras):
        if extras[i] == "--to-env":
            to_env = extras[i + 1]
            i += 2
            continue
        cid, _, ruta = extras[i].partition("=")
        if not ruta:
            raise ValueError(f"Argumento invalido: {extras[i]!r}. Se espera cid=ruta.")
        imagenes.append((cid, ruta))
        i += 1

    with open(path, encoding="utf-8") as f:
        send(subject, f.read(), imagenes=imagenes or None, mail_to_env=to_env)
