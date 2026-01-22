import gspread
import json
import os
import smtplib
from email.message import EmailMessage
from oauth2client.service_account import ServiceAccountCredentials
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter
from copy import deepcopy

# ======================================================
# EMAIL CONFIG
# ======================================================
SENDER_EMAIL = os.environ.get("GMAIL_USER")
APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")

if not SENDER_EMAIL or not APP_PASSWORD:
    raise EnvironmentError("❌ Gmail credentials missing")

# ======================================================
# FILE CONFIG
# ======================================================
CERT_TEMPLATE = "Certificate.pdf"
FONT_PATH = "PlayfairDisplay-Regular.ttf"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ======================================================
# GOOGLE SHEETS AUTH
# ======================================================
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

google_creds = json.loads(os.environ["GOOGLE_CREDENTIALS"])
creds = ServiceAccountCredentials.from_json_keyfile_dict(google_creds, scope)
client = gspread.authorize(creds)

SPREADSHEET_ID = "1IIMYLmvvFIXcMChASZEBxmBovdCv-CuKK8bgvZsNJxM"
sheet = client.open_by_key(SPREADSHEET_ID).sheet1
records = sheet.get_all_records()
headers = sheet.row_values(1)

if "Status" not in headers:
    sheet.update_cell(1, len(headers) + 1, "Status")
    headers.append("Status")

STATUS_COL = headers.index("Status") + 1

# ======================================================
# LOAD PDF SIZE
# ======================================================
base_pdf = PdfReader(CERT_TEMPLATE)
PAGE_WIDTH = float(base_pdf.pages[0].mediabox.width)
PAGE_HEIGHT = float(base_pdf.pages[0].mediabox.height)

# ======================================================
# FONT
# ======================================================
pdfmetrics.registerFont(TTFont("Playfair", FONT_PATH))

# ======================================================
# DASH MEASUREMENTS (FROM LIBREOFFICE)
# ======================================================
# NAME DASH
NAME_DASH_X = 260
NAME_DASH_WIDTH = 235
NAME_Y = PAGE_HEIGHT - (4.23 * 72) + 6

# EVENT DASH
EVENT_DASH_X = 56
EVENT_DASH_WIDTH = 165
EVENT_Y = PAGE_HEIGHT - (4.85 * 72) + 14

# ======================================================
# CENTER TEXT ON DASH (CORE FIX)
# ======================================================
def draw_centered_on_dash(c, text, font, size, dash_x, dash_width, y):
    c.setFont(font, size)
    text_width = pdfmetrics.stringWidth(text, font, size)
    center_x = dash_x + (dash_width / 2)
    text_x = center_x - (text_width / 2)
    c.drawString(text_x, y, text)

# ======================================================
# CERTIFICATE CREATOR
# ======================================================
def create_certificate(name, event):
    overlay_path = os.path.join(OUTPUT_DIR, "overlay.pdf")
    final_path = os.path.join(OUTPUT_DIR, f"{name.replace(' ', '_')}.pdf")

    c = canvas.Canvas(overlay_path, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    # NAME
    draw_centered_on_dash(
        c, name, "Playfair", 26,
        NAME_DASH_X, NAME_DASH_WIDTH, NAME_Y
    )

    # EVENT
    draw_centered_on_dash(
        c, event, "Playfair", 20,
        EVENT_DASH_X, EVENT_DASH_WIDTH, EVENT_Y
    )

    c.save()

    overlay = PdfReader(overlay_path)
    writer = PdfWriter()

    # VERY IMPORTANT: fresh page every time
    fresh_page = deepcopy(base_pdf.pages[0])
    fresh_page.merge_page(overlay.pages[0])
    writer.add_page(fresh_page)

    with open(final_path, "wb") as f:
        writer.write(f)

    return final_path

# ======================================================
# EMAIL SENDER
# ======================================================
def send_email(to_email, name, pdf_path):
    msg = EmailMessage()
    msg["Subject"] = "Certificate of Participation – ITRONIX"
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email

    msg.set_content(f"""
Dear {name},

Greetings from Guru Nanak College.

Please find attached your Certificate of Participation
for the event conducted during ITRONIX IT Fest.

Regards,
Department of Information Technology
Guru Nanak College
""")

    with open(pdf_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(pdf_path)
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)

# ======================================================
# MAIN LOOP
# ======================================================
for row_index, row in enumerate(records, start=2):
    name = row.get("Full Name")
    event = row.get("EVENT")
    email = row.get("Email Address")
    status = row.get("Status", "")

    if status.startswith("✅"):
        continue

    if not name or not event or not email:
        sheet.update_cell(row_index, STATUS_COL, "❌ FAILED (Missing data)")
        continue

    try:
        sheet.update_cell(row_index, STATUS_COL, "⏳ PENDING")
        pdf_path = create_certificate(name, event)
        send_email(email, name, pdf_path)
        sheet.update_cell(row_index, STATUS_COL, "✅ SENT")
        print(f"✔ Sent to {email}")
    except Exception as e:
        sheet.update_cell(row_index, STATUS_COL, "❌ FAILED")
        print(f"❌ Error for {email}: {e}")

print("🎉 ALL CERTIFICATES GENERATED & SENT")
