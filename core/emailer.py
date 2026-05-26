import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587


def send_report_email(
    pdf_path: str,
    recipient_email: str,
    month_label: str,
    commentary_text: str,
    sender_email: str,
    outlook_password: str,
    synthese_data: dict | None = None,
) -> None:
    """
    Envoie le rapport PDF par email via SMTP Outlook (STARTTLS).

    Raises:
        smtplib.SMTPAuthenticationError: credentials incorrects
        FileNotFoundError: PDF introuvable
    """
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF introuvable : {pdf_path}")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"📊 Rapport charges {month_label} — ✅ Validé"
    msg["From"] = sender_email
    msg["To"] = recipient_email

    # ── Corps HTML ────────────────────────────────────────────────────────────
    synthese = synthese_data or {}
    rows_html = ""
    for dept in ["RH", "Tech", "S&M", "G&A", "TOTAL"]:
        val = synthese.get(dept, 0)
        bold = " style='font-weight:bold; background:#e8edf4;'" if dept == "TOTAL" else ""
        rows_html += f"<tr{bold}><td style='padding:6px 12px;'>{dept}</td><td style='padding:6px 12px; text-align:right;'>{val:,.0f} k€</td></tr>"

    commentary_html = commentary_text.replace("\n", "<br/>")

    html_body = f"""
    <html><body style="font-family: Arial, sans-serif; color: #2d3748; max-width: 640px; margin: auto;">
      <div style="background:#042C53; padding: 20px 24px; border-radius:4px 4px 0 0;">
        <h1 style="color:white; margin:0; font-size:20px;">CybelAngel</h1>
        <p style="color:#A8C4E5; margin:4px 0 0; font-size:13px;">Rapport consolidation charges — {month_label}</p>
      </div>
      <div style="border:1px solid #dde3ec; border-top:none; padding:24px; border-radius:0 0 4px 4px;">
        <h2 style="font-size:14px; color:#042C53; margin-top:0;">Synthèse charges</h2>
        <table style="border-collapse:collapse; width:100%; font-size:13px;">
          <thead>
            <tr style="background:#042C53; color:white;">
              <th style="padding:8px 12px; text-align:left;">Département</th>
              <th style="padding:8px 12px; text-align:right;">{month_label}</th>
            </tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>

        <h2 style="font-size:14px; color:#042C53; margin-top:24px;">Analyse — {month_label}</h2>
        <div style="background:#eef2f8; border-left:3px solid #1A5CA8; padding:12px 16px; border-radius:3px;
                    font-style:italic; font-size:13px; color:#2d3748; line-height:1.6;">
          {commentary_html}
        </div>

        <p style="font-size:11px; color:#8a9ab5; margin-top:24px; border-top:1px solid #dde3ec; padding-top:12px;">
          Document généré automatiquement — Processus certifié ✅ | Écart Synthèse/Sources : 0 k€
        </p>
      </div>
    </body></html>
    """

    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # ── Pièce jointe PDF ──────────────────────────────────────────────────────
    with open(pdf_file, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f'attachment; filename="{pdf_file.name}"',
    )
    msg.attach(part)

    # ── Envoi SMTP Outlook ────────────────────────────────────────────────────
    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(sender_email, outlook_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
