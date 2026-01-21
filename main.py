import os
import json
import gspread
import smtplib
from email.message import EmailMessage
from oauth2client.service_account import ServiceAccountCredentials
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter

# ================= CONFIG =================
SPREADSHEET_ID = "1IIMYLmvvFIXcMChASZEBxmBovdCv-CuKK8bgvZsNJxM"

SENDER_EMAIL = os.environ["GMAIL_USER"]
APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

CERT_TEMPLATE = "certificate.pdf"
FONT_PATH = "PlayfairDisplay-Regular.ttf"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

PAGE_WIDTH, PAGE_HEIGHT = 842, 595
CENTER_X = PAGE_WIDTH / 2
NAME_Y = 310
EVENT_Y = 265

# ================= GOOGLE AUTH =================
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

google_creds = json.loads(os.environ["GOOGLE_CREDENTIALS"])
creds = ServiceAccountCredentials.from_json_keyfile_dict(google_creds, scope)
client = gspread.authorize(creds)

sheet = client.open_by_key(SPREADSHEET_ID).sheet1
records = sheet.get_all_records()
headers = sheet.row_values(1)

if "Status" not in headers:
    sheet.update_cell(1, len(headers) + 1, "Status")
    headers.append("Status")

status_col = headers.index("Status") + 1

pdfmetrics.registerFont(TTFont("Playfair", FONT_PATH))

# ================= FUNCTIONS =================
def create_certificate(name, event):
    overlay = f"{OUTPUT_DIR}/overlay.pdf"
    output = f"{OUTPUT_DIR}/{name.replace(' ', '_')}.pdf"

    c = canvas.Canvas(overlay, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
    c.setFont("Playfair", 30)
    c.drawCentredString(CENTER_X, NAME_Y, name)
    c.setFont("Playfair", 22)
    c.drawCentredString(CENTER_X, EVENT_Y, event)
    c.save()

    base = PdfReader(CERT_TEMPLATE)
    layer = PdfReader(overlay)
    writer = PdfWriter()

    page = base.pages[0]
    page.merge_page(layer.pages[0])
    writer.add_page(page)

    with open(output, "wb") as f:
        writer.write(f)

    return output

def send_email(to, name, pdf):
    msg = EmailMessage()
    msg["Subject"] = "Certificate of Participation – ITRONIX 2026"
    msg["From"] = f"ITRONIX <{SENDER_EMAIL}>"
    msg["To"] = to

    msg.set_content(f"""
Dear {name},

Congratulations 🎉

Please find attached your Certificate of Participation
for the ITRONIX 2026 event.

Regards,
Department of Information Technology
Guru Nanak College
""")

    with open(pdf, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=os.path.basename(pdf))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(SENDER_EMAIL, APP_PASSWORD)
        s.send_message(msg)

# ================= MAIN =================
for i, row in enumerate(records, start=2):
    name = row.get("Full Name")
    email = row.get("Email Address")
    event = row.get("EVENT")
    status = row.get("Status", "")

    if status == "SENT":
        continue

    if not name or not email or not event:
        sheet.update_cell(i, status_col, "FAILED")
        continue

    try:
        pdf = create_certificate(name, event)
        send_email(email, name, pdf)
        sheet.update_cell(i, status_col, "SENT")
        print(f"Sent to {email}")

    except Exception as e:
        sheet.update_cell(i, status_col, "FAILED")
        print(e)
