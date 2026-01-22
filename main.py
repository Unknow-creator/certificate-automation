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

# ================= EMAIL CONFIG =================
SENDER_EMAIL = os.environ.get("GMAIL_USER")
APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")

if not SENDER_EMAIL or not APP_PASSWORD:
    raise EnvironmentError("Missing Gmail credentials")

# ================= FILE CONFIG =================
CERT_TEMPLATE = "Certificate.pdf"
FONT_PATH = "PlayfairDisplay-Regular.ttf"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ================= GOOGLE SHEET AUTH =================
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

# ================= LOAD TEMPLATE SIZE =================
base_pdf = PdfReader(CERT_TEMPLATE)
PAGE_WIDTH = float(base_pdf.pages[0].mediabox.width)
PAGE_HEIGHT = float(base_pdf.pages[0].mediabox.height)
CENTER_X = PAGE_WIDTH / 2

# ================= FONT =================
pdfmetrics.registerFont(TTFont("Playfair", FONT_PATH))

# ================= EXACT DASH POSITIONS =================
# These values MATCH your Inkscape dashed lines
NAME_BOX_X = PAGE_WIDTH * 0.32
NAME_BOX_Y = PAGE_HEIGHT * 0.46
NAME_BOX_W = PAGE_WIDTH * 0.36

EVENT_BOX_X = PAGE_WIDTH * 0.17
EVENT_BOX_Y = PAGE_HEIGHT * 0.395
EVENT_BOX_W = PAGE_WIDTH * 0.26

# ================= TEXT DRAWING =================
def draw_centered_in_box(c, text, font, max_size, x, y, box_width):
    size = max_size
    while size > 12:
        text_width = pdfmetrics.stringWidth(text, font, size)
        if text_width <= box_width:
            break
        size -= 1

    c.setFont(font, size)
    c.drawCentredString(x + box_width / 2, y, text)

# ================= CERTIFICATE CREATOR =================
def create_certificate(name, event):
    overlay_path = f"{OUTPUT_DIR}/overlay.pdf"
    final_path = f"{OUTPUT_DIR}/{name.replace(' ', '_')}.pdf"

    c = canvas.Canvas(overlay_path, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    draw_centered_in_box(
        c, name, "Playfair", 26,
        NAME_BOX_X, NAME_BOX_Y, NAME_BOX_W
    )

    draw_centered_in_box(
        c, event, "Playfair", 20,
        EVENT_BOX_X, EVENT_BOX_Y, EVENT_BOX_W
    )

    c.save()

    overlay = PdfReader(overlay_path)

    # 🔥 IMPORTANT FIX — CLONE PAGE
    fresh_page = deepcopy(base_pdf.pages[0])
    fresh_page.merge_page(overlay.pages[0])

    writer = PdfWriter()
    writer.add_page(fresh_page)

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

# ================= MAIN LOOP =================
for i, row in enumerate(records, start=2):
    name = row.get("Full Name")
    event = row.get("EVENT")
    email = row.get("Email Address")
    status = row.get("Status", "")

    if status.startswith("✅"):
        continue

    if not name or not event or not email:
        sheet.update_cell(i, STATUS_COL, "❌ FAILED (Missing data)")
        continue

    try:
        sheet.update_cell(i, STATUS_COL, "⏳ PENDING")
        pdf = create_certificate(name, event)
        send_email(email, name, pdf)
        sheet.update_cell(i, STATUS_COL, "✅ SENT")
    except Exception as e:
        sheet.update_cell(i, STATUS_COL, "❌ FAILED")
        print(e)

print("🎉 All certificates generated successfully")
