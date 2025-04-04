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
    # Fix the private key formatting
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n").strip()
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
except Exception as e:
    raise ValueError(f"❌ Error connecting to Google Sheets: {e}")

# ------------------------------
# SMTP Setup
# ------------------------------
SMTP_SERVER = "smtp-relay.brevo.com"  # Or your preferred SMTP server
SMTP_PORT = 587
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")

if not SMTP_USER or not SMTP_PASSWORD:
    raise ValueError("❌ SMTP_USER or SMTP_PASSWORD not found!")

def send_email(email):
    """
    Sends an email using an HTML template (Ala.html).
    Assumes 'email' is the recipient's email address.
    """
    try:
        subject = "Завтра встречаемся на BI Ecosystem — ждём Вас!"
        msg = EmailMessage()
        msg["From"] = "noreply@biecosystem.kz"
        msg["To"] = email
        msg["Subject"] = subject
        msg.set_type("multipart/related")

        # Load the HTML template
        template_filename = "Ala.html"
        if os.path.exists(template_filename):
            with open(template_filename, "r", encoding="utf-8") as template_file:
                html_content = template_file.read()
        else:
            print(f"❌ Template file {template_filename} not found.")
            return False

        # Insert a unique placeholder (optional)
        unique_id = random.randint(100000, 999999)
        html_content = html_content.replace("<!--UNIQUE_PLACEHOLDER-->", str(unique_id))

        # Embed logo if it exists
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
    Checks each row in the Google Sheet.
    Columns in your doc:
        A -> Name
        B -> Email  (index 1)
        C -> Phone
        D -> language
        E -> Checkbox
        F -> referr
        G -> formId
        H -> sent
        I -> requestID
        J -> dbbase
        K -> admin
        L -> status (index 11)

    If 'status' (column L) is not "Done", send an email to 'Email' (column B),
    then mark 'status' as "Done" in column L.
    """
    try:
        all_values = sheet.get_all_values()
        # Skip the header row (starting from index 1)
        for i in range(1, len(all_values)):
            row = all_values[i]
            # Ensure the row has at least 12 columns
            if len(row) < 12:
                continue

            email = row[1].strip()    # Column B
            status = row[11].strip().lower()  # Column L

            # If status is already "Done", skip
            if status == "done":
                continue

            # Attempt to send the email
            if send_email(email):
                # Update the status to "Done" in column L (12th column)
                sheet.update_cell(i + 1, 12, "Done")

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
        time.sleep(30)

# Start the background task
threading.Thread(target=background_task, daemon=True).start()

@app.route("/")
def home():
    return "Email-only sender is running!", 200

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
