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

# ================= EMAIL CONFIG =================
SENDER_EMAIL = os.environ["GMAIL_USER"]
APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

# ================= FILES =================
CERT_TEMPLATE = "certificate.pdf"
FONT_PATH = "PlayfairDisplay-Regular.ttf"

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ================= PAGE & TEXT POSITION =================
# A4 Landscape = 842 x 595 points
PAGE_WIDTH = 842
PAGE_HEIGHT = 595
CENTER_X = PAGE_WIDTH // 2

# Calibrated for your certificate image
NAME_Y = 315      # Upper dotted line
EVENT_Y = 270     # Lower dotted line

# ================= GOOGLE SHEETS AUTH =================
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

google_creds = json.loads(os.environ["GOOGLE_CREDS"])

creds = ServiceAccountCredentials.from_json_keyfile_dict(
    google_creds, scope
)

client = gspread.authorize(creds)

# OPEN BY SPREADSHEET NAME + SHEET TAB NAME
spreadsheet = client.open("CertificateData")
sheet = spreadsheet.worksheet("Form_Responses")

records = sheet.get_all_records()

# ================= CERTIFICATE CREATION =================
def create_certificate(name, event):
    overlay_pdf = f"{OUTPUT_DIR}/overlay.pdf"
    final_pdf = f"{OUTPUT_DIR}/{name}.pdf"

    pdfmetrics.registerFont(TTFont("Playfair", FONT_PATH))

    # Create overlay PDF
    c = canvas.Canvas(overlay_pdf, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    # NAME
    c.setFont("Playfair", 30)
    c.drawCentredString(CENTER_X, NAME_Y, name)

    # EVENT
    c.setFont("Playfair", 22)
    c.drawCentredString(CENTER_X, EVENT_Y, event)

    c.save()

    # Merge overlay with certificate template
    base_pdf = PdfReader(CERT_TEMPLATE)
    overlay = PdfReader(overlay_pdf)
    writer = PdfWriter()

    page = base_pdf.pages[0]
    page.merge_page(overlay.pages[0])
    writer.add_page(page)

    with open(final_pdf, "wb") as f:
        writer.write(f)

    return final_pdf

# ================= EMAIL SENDER =================
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

# ================= MAIN EXECUTION =================
for row_index, row in enumerate(records, start=2):  # start=2 = actual sheet row
    name = row.get("Full Name")
    event = row.get("EVENT")
    email = row.get("Email Address")
    status = row.get("Status", "")

    if not name or not event or not email:
        continue

    if status == "SENT":
        continue

    pdf_path = create_certificate(name, event)
    send_email(email, name, pdf_path)

    # Update Status column (H = 8)
    sheet.update_cell(row_index, 8, "SENT")

    print(f"✔ Certificate sent to {email}")
