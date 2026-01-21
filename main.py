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
page = base_pdf.pages[0]

PAGE_WIDTH = float(page.mediabox.width)
PAGE_HEIGHT = float(page.mediabox.height)

# ================= DASHED LINE LOCKED POSITIONS =================
NAME_X = PAGE_WIDTH * 0.18
NAME_WIDTH = PAGE_WIDTH * 0.64
NAME_Y = PAGE_HEIGHT * 0.47

EVENT_X = PAGE_WIDTH * 0.18
EVENT_WIDTH = PAGE_WIDTH * 0.64
EVENT_Y = PAGE_HEIGHT * 0.40

# ================= FONT =================
pdfmetrics.registerFont(TTFont("Playfair", FONT_PATH))

# ================= TEXT BOX DRAWER (NO FLYING) =================
def draw_text_in_box(c, text, x, width, y, font, max_size, min_size):
    size = max_size
    while size >= min_size:
        text_width = pdfmetrics.stringWidth(text, font, size)
        if text_width <= width:
            break
        size -= 1

    c.setFont(font, size)
    x_pos = x + (width - text_width) / 2
    c.drawString(x_pos, y, text)

# ================= CERTIFICATE CREATOR =================
def create_certificate(name, event):
    overlay_path = f"{OUTPUT_DIR}/overlay.pdf"
    final_path = f"{OUTPUT_DIR}/{name.replace(' ', '_')}.pdf"

    c = canvas.Canvas(overlay_path, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))

    # NAME (LOCKED EXACTLY ON DASH)
    draw_text_in_box(
        c,
        name,
        NAME_X,
        NAME_WIDTH,
        NAME_Y,
        "Playfair",
        max_size=32,
        min_size=22
    )

    # EVENT (LOCKED EXACTLY ON DASH)
    draw_text_in_box(
        c,
        event,
        EVENT_X,
        EVENT_WIDTH,
        EVENT_Y,
        "Playfair",
        max_size=22,
        min_size=16
    )

    c.save()

    base = PdfReader(CERT_TEMPLATE)
    overlay = PdfReader(overlay_path)

    writer = PdfWriter()
    page = base.pages[0]
    page.merge_page(overlay.pages[0])
    writer.add_page(page)

    with open(final_path, "wb") as f:
        writer.write(f)

    os.remove(overlay_path)
    return final_path

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
for row_index, row in enumerate(records, start=2):
    name = row.get("Full Name")
    event = row.get("EVENT")
    email = row.get("Email Address")
    status = row.get("Status", "").strip()

    if status.startswith("✅"):
        continue

    if not name or not event or not email:
        sheet.update_cell(row_index, STATUS_COL, "❌ FAILED (Missing data)")
        continue

    try:
        sheet.update_cell(row_index, STATUS_COL, "⏳ PENDING")
        pdf = create_certificate(name, event)
        send_email(email, name, pdf)
        sheet.update_cell(row_index, STATUS_COL, "✅ SENT")
        print(f"✔ Sent to {email}")
    except Exception as e:
        sheet.update_cell(row_index, STATUS_COL, "❌ FAILED")
        print(f"❌ Error for {email}: {e}")

print("🎉 Certificate automation completed successfully")
