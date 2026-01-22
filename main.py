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

# ======================================================
# EMAIL CONFIG
# ======================================================
SENDER_EMAIL = os.environ["GMAIL_USER"]
APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

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

# Status column
if "Status" not in headers:
    sheet.update_cell(1, len(headers) + 1, "Status")
    headers.append("Status")

STATUS_COL = headers.index("Status") + 1

# ======================================================
# LOAD PDF SIZE
# ======================================================
base_pdf = PdfReader(CERT_TEMPLATE)
PAGE = base_pdf.pages[0]
PAGE_WIDTH = float(PAGE.mediabox.width)
PAGE_HEIGHT = float(PAGE.mediabox.height)

CENTER_X = PAGE_WIDTH / 2

# ======================================================
# FONT
# ======================================================
pdfmetrics.registerFont(TTFont("Playfair", FONT_PATH))

# ======================================================
# 🎯 PERFECT POSITIONS (LOCKED)
# ======================================================
# NAME (centered exactly on dash)
NAME_Y = PAGE_HEIGHT - (4.23 * 72) + 10

# EVENT (lifted so BGMI sits perfectly)
EVENT_Y = PAGE_HEIGHT - (4.85 * 72) + 14

# ======================================================
# FONT AUTO SCALE (LONG NAMES SAFE)
# ======================================================
def fit_font(text, max_width, base_size):
    size = base_size
    while pdfmetrics.stringWidth(text, "Playfair", size) > max_width:
        size -= 1
    return size

# ======================================================
# CERTIFICATE CREATOR
# ======================================================
def create_certificate(name, event):
    overlay_path = f"{OUTPUT_DIR}/overlay.pdf"
    final_path = f"{OUTPUT_DIR}/{name.replace(' ', '_')}.pdf"

    c = canvas.Canvas(overlay_path, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    # NAME
    name_font = fit_font(name, 360, 28)
    c.setFont("Playfair", name_font)
    c.drawCentredString(CENTER_X, NAME_Y, name)

    # EVENT
    event_font = fit_font(event, 200, 20)
    c.setFont("Playfair", event_font)
    c.drawCentredString(CENTER_X, EVENT_Y, event)

    c.save()

    overlay = PdfReader(overlay_path)
    writer = PdfWriter()

    page = PdfReader(CERT_TEMPLATE).pages[0]
    page.merge_page(overlay.pages[0])
    writer.add_page(page)

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
    except Exception as e:
        sheet.update_cell(row_index, STATUS_COL, "❌ FAILED")
        print(e)

print("🎉 ALL CERTIFICATES GENERATED & SENT")
