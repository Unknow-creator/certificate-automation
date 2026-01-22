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

# ================= CONFIG =================
CERT_TEMPLATE = "Certificate.pdf"
FONT_PATH = "PlayfairDisplay-Regular.ttf"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

SENDER_EMAIL = os.environ["GMAIL_USER"]
APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

# ================= GOOGLE SHEET =================
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

google_creds = json.loads(os.environ["GOOGLE_CREDENTIALS"])
creds = ServiceAccountCredentials.from_json_keyfile_dict(google_creds, scope)
client = gspread.authorize(creds)

sheet = client.open_by_key("1IIMYLmvvFIXcMChASZEBxmBovdCv-CuKK8bgvZsNJxM").sheet1
records = sheet.get_all_records()
headers = sheet.row_values(1)

if "Status" not in headers:
    sheet.update_cell(1, len(headers) + 1, "Status")
headers = sheet.row_values(1)
STATUS_COL = headers.index("Status") + 1

# ================= PDF SETUP =================
pdfmetrics.registerFont(TTFont("Playfair", FONT_PATH))
base_pdf = PdfReader(CERT_TEMPLATE)
base_page = base_pdf.pages[0]

PAGE_W = float(base_page.mediabox.width)
PAGE_H = float(base_page.mediabox.height)

# ================= EXACT BOX POSITIONS (IN POINTS) =================
def inch(v): return v * 72

NAME_BOX = {
    "x": inch(3.87),
    "y": inch(3.87),
    "w": inch(2.63),
    "h": inch(0.38),
    "size": 32
}

EVENT_BOX = {
    "x": inch(0.50),
    "y": inch(4.50),
    "w": inch(2.38),
    "h": inch(0.37),
    "size": 22
}

# ================= DRAW TEXT PERFECTLY CENTERED =================
def draw_centered_text(c, text, box, font="Playfair"):
    c.setFont(font, box["size"])

    text_width = pdfmetrics.stringWidth(text, font, box["size"])
    x = box["x"] + (box["w"] - text_width) / 2

    # baseline correction
    y = box["y"] + (box["h"] / 2) - (box["size"] * 0.35)

    c.drawString(x, y, text)

# ================= CERTIFICATE GENERATOR =================
def create_certificate(name, event):
    overlay_path = f"{OUTPUT_DIR}/overlay.pdf"
    final_path = f"{OUTPUT_DIR}/{name.replace(' ', '_')}.pdf"

    c = canvas.Canvas(overlay_path, pagesize=(PAGE_W, PAGE_H))

    draw_centered_text(c, name, NAME_BOX)
    draw_centered_text(c, event, EVENT_BOX)

    c.save()

    overlay = PdfReader(overlay_path)
    writer = PdfWriter()

    page = base_pdf.pages[0]
    page.merge_page(overlay.pages[0])
    writer.add_page(page)

    with open(final_path, "wb") as f:
        writer.write(f)

    return final_path

# ================= EMAIL =================
def send_email(to_email, name, pdf_path):
    msg = EmailMessage()
    msg["Subject"] = "Certificate of Participation – ITRONIX"
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email

    msg.set_content(f"""
Dear {name},

Please find attached your Certificate of Participation
for ITRONIX IT Fest.

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

# ================= MAIN =================
for i, row in enumerate(records, start=2):
    if row.get("Status", "").startswith("✅"):
        continue

    name = row.get("Full Name")
    event = row.get("EVENT")
    email = row.get("Email Address")

    if not name or not event or not email:
        sheet.update_cell(i, STATUS_COL, "❌ Missing data")
        continue

    try:
        sheet.update_cell(i, STATUS_COL, "⏳ Processing")
        pdf = create_certificate(name, event)
        send_email(email, name, pdf)
        sheet.update_cell(i, STATUS_COL, "✅ Sent")
    except Exception as e:
        sheet.update_cell(i, STATUS_COL, "❌ Failed")
        print(e)

print("🎉 All certificates generated & emailed")
