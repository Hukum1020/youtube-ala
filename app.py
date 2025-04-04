import os
import time
import smtplib
import ssl
import gspread
import json
import traceback
import random
from email.message import EmailMessage
from oauth2client.service_account import ServiceAccountCredentials
from flask import Flask
import threading

app = Flask(__name__)

# ------------------------------
# Google Sheets API Setup
# ------------------------------
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
if not SPREADSHEET_ID:
    raise ValueError("❌ SPREADSHEET_ID not found!")

CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
if not CREDENTIALS_JSON:
    raise ValueError("❌ GOOGLE_CREDENTIALS_JSON not found!")

try:
    creds_dict = json.loads(CREDENTIALS_JSON)
    # Fix private key formatting
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n").strip()
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
except Exception as e:
    raise ValueError(f"❌ Error connecting to Google Sheets: {e}")

# ------------------------------
# SMTP Setup
# ------------------------------
SMTP_SERVER = "smtp-relay.brevo.com"  # or your preferred SMTP server
SMTP_PORT = 587
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")

if not SMTP_USER or not SMTP_PASSWORD:
    raise ValueError("❌ SMTP_USER or SMTP_PASSWORD not found!")

def send_email(email, language):
    """
    Sends an email using an HTML template.
    - 'email' is the recipient's address (from column B).
    - 'language' (from column D) determines the subject and template.
    """
    try:
        if language == "ru":
            subject = "Подключайтесь к эфиру и выиграйте Iphone16 🎁 Уже завтра — BI Ecosystem! "
        else:
            subject = "Эфирге қосылып, Iphone16 ұтып алыңыз🎁 Ертең BI Ecosystem болады!"

        msg = EmailMessage()
        msg["From"] = "noreply@biecosystem.kz"
        msg["To"] = email
        msg["Subject"] = subject
        msg.set_type("multipart/related")

        # Load HTML template (AlaRu.html or AlaKz.html)
        template_filename = f"Ala{language}.html"
        if os.path.exists(template_filename):
            with open(template_filename, "r", encoding="utf-8") as template_file:
                html_content = template_file.read()
        else:
            print(f"❌ Template file {template_filename} not found.")
            return False

        # Insert a unique identifier (optional)
        unique_id = random.randint(100000, 999999)
        html_content = html_content.replace("<!--UNIQUE_PLACEHOLDER-->", str(unique_id))

        # Embed logo if available
        logo_path = "logo2.png"
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as logo_file:
                msg.add_related(
                    logo_file.read(),
                    maintype="image",
                    subtype="png",
                    filename="logo2.png",
                    cid="logo"
                )
            html_content = html_content.replace('src="logo2.png"', 'src="cid:logo"')
        else:
            print("⚠️ Logo not found. Sending email without logo.")

        # Add HTML content
        msg.add_alternative(html_content, subtype="html")

        # Send the email
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)

        print(f"✅ Email sent to {email}")
        return True

    except Exception as e:
        print(f"❌ Error sending email to {email}: {e}")
        traceback.print_exc()
        return False

def process_new_guests():
    """
    Iterates through rows in the Google Sheet.
    Expects:
      - Column B (index 1): Email address.
      - Column D (index 3): Language ("ru" or "kz").
      - Column K (index 10): Status.
    If status is not "Done", sends an email and then marks status as "Done" in column K.
    Sends one email every 0.5 seconds.
    """
    try:
        all_values = sheet.get_all_values()
        # Skip header row; start at row index 1
        for i in range(1, len(all_values)):
            row = all_values[i]
            if len(row) < 11:
                continue

            email = row[1].strip()               # Column B
            language = row[3].strip().lower()     # Column D
            status = row[10].strip().lower()       # Column K

            if status == "done":
                continue

            if send_email(email, language):
                # Update status to "Done" in column K (11th column)
                sheet.update_cell(i + 1, 11, "Done")

            # Wait 0.5 seconds before processing the next email
            time.sleep(1)

    except Exception as e:
        print(f"[Error] processing guests: {e}")
        traceback.print_exc()

def background_task():
    """
    Runs in a separate thread, checking the sheet every 30 seconds.
    """
    while True:
        try:
            process_new_guests()
        except Exception as e:
            print(f"[Error] in background task: {e}")
            traceback.print_exc()
        time.sleep(10)

# Start the background task
threading.Thread(target=background_task, daemon=True).start()

@app.route("/")
def home():
    return "Email-only sender is running!", 200

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
