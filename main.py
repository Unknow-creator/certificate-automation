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
SENDER_EMAIL = os.environ.get("GMAIL_USER")
APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")

if not SENDER_EMAIL or not APP_PASSWORD:
    raise EnvironmentError(
        "Missing Gmail credentials! Set GMAIL_USER and GMAIL_APP_PASSWORD environment variables."
    )

CERT_TEMPLATE = "certificate.pdf"
FONT_PATH = "PlayfairDisplay-Regular.ttf"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ================= PAGE & TEXT POSITION =================
PAGE_WIDTH = 842   # A4 Landscape
PAGE_HEIGHT = 595
CENTER_X = PAGE_WIDTH // 2

NAME_Y = 315
EVENT_Y = 270

# ================= GOOGLE SHEETS AUTH =================
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

google_creds_env = os.environ.get("GOOGLE_CREDENTIALS")

if google_creds_env:
    google_creds = json.loads(google_creds_env)
else:
    # fallback for local testing
    with open("credentials.json") as f:
        google_creds = json.load(f)

creds = ServiceAccountCredentials.from_json_keyfile_dict(google_creds, scope)
client = gspread.authorize(creds)

spreadsheet = client.open("CertificateData")
sheet = spreadsheet.get_worksheet(0)

records = sheet.get_all_records()

if not records:
    print("✅ Sheet has headers only. Nothing to process.")
    exit(0)

headers = sheet.row_values(1)

# Ensure Status column exists
if "Status" not in headers:
    sheet.update_cell(1, len(headers) + 1, "Status")
    headers.append("Status")

status_col = headers.index("Status") + 1

# ================= CERTIFICATE CREATION =================
pdfmetrics.registerFont(TTFont("Playfair", FONT_PATH))

def create_certificate(name, event):
    overlay_pdf = f"{OUTPUT_DIR}/overlay.pdf"
    final_pdf = f"{OUTPUT_DIR}/{name}.pdf"

    c = canvas.Canvas(overlay_pdf, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    # NAME
    c.setFont("Playfair", 30)
    c.drawCentredString(CENTER_X, NAME_Y, name)

    # EVENT
    c.setFont("Playfair", 22)
    c.drawCentredString(CENTER_X, EVENT_Y, event)

    c.save()

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
    msg["From"] = f"ITRONIX Certificates <{SENDER_EMAIL}>"
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
for row_index, row in enumerate(records, start=2):
    name = row.get("Full Name")
    event = row.get("EVENT")
    email = row.get("Email Address")
    status = row.get("Status", "").strip()

    if status.startswith("✅"):
        continue

    if not name or not event or not email:
        sheet.update_cell(row_index, status_col, "❌ FAILED (Missing data)")
        continue

    try:
        sheet.update_cell(row_index, status_col, "⏳ PENDING")
        pdf_path = create_certificate(name, event)
        send_email(email, name, pdf_path)
        sheet.update_cell(row_index, status_col, "✅ SENT")
        print(f"✔ Certificate sent to {email}")

    except Exception as e:
        sheet.update_cell(row_index, status_col, "❌ FAILED")
        print(f"❌ Failed for {email}: {e}")

print("🎉 Certificate automation completed successfully")
