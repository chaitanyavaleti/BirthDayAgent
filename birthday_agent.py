"""
birthday_agent_outlook_once_per_day.py

Requirements:
    pip install pywin32 pandas openpyxl

Project layout:
    birthday_agent/
    ├── TeamBirthDay.xlsx            # columns: Name, Email, Birthday (YYYY-MM-DD or Excel date)
    ├── template.html        # HTML template using {name} and {year}
    ├── anniversary_template.html
    ├── images/
    │     ├── birthday_bg.jpg
    │     └── anniversary_bg.jpg
    └── birthday_agent.py  # this file

Behavior:
    - Checks once per day at CHECK_TIME (local system time, format "HH:MM")
    - Records last run date to last_run_date.txt (created in the same dir)
"""

import os
import time
import datetime
import base64
import pandas as pd
import win32com.client as win32

import smtplib
from email.message import EmailMessage

SMTP_SERVER = "ban-mail01-l001.otxlab.net"
SMTP_PORT = 25

# ---------- Configuration ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(BASE_DIR, "TeamBirthDay.xlsx")
TEMPLATE_FILE = os.path.join(BASE_DIR, "birthday_template.html")
ANNIVERSARY_TEMPLATE_FILE = os.path.join(BASE_DIR, "anniversary_template.html")
IMAGES_FOLDER = os.path.join(BASE_DIR, "images")
STATE_FILE = os.path.join(BASE_DIR, "last_run_date.txt")  # stores YYYY-MM-DD of last successful run
CC_LIST_GLOBAL = []                # e.g. ["manager@company.com", "hr@company.com"] or leave empty
# -----------------------------------

def read_last_run_date():
    """Return a date object for last run, or None if not present."""
    if not os.path.exists(STATE_FILE):
        return None
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            txt = f.read().strip()
            if not txt:
                return None
            return datetime.datetime.strptime(txt, "%Y-%m-%d").date()
    except Exception:
        return None

def write_last_run_date(d: datetime.date):
    """Write last run date as YYYY-MM-DD."""
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        f.write(d.isoformat())

def load_template(template_file, name, years=None):
    """Load template.html and replace placeholders"""
    with open(template_file, "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace("{name}", name)
    html = html.replace("{year}", str(datetime.datetime.now().year))

    if years is not None:
        html = html.replace("{years}", str(years))

    return html

def embed_images_in_html(html):
    """
    Replace cid: references with data URIs (base64). This approach ensures
    most email clients display the image inline when using Outlook automation.
    Note: keep image filenames and cids consistent with your template.
    """
    images = {
        "confetti": "birthday_bg.jpg",
    }
    for cid, filename in images.items():
        path = os.path.join(IMAGES_FOLDER, filename)
        if os.path.exists(path):
            try:
                with open(path, "rb") as img:
                    b64 = base64.b64encode(img.read()).decode("utf-8")
                    # assume png; if jpg change to image/jpeg
                    data_uri = f"data:image/png;base64,{b64}"
                    html = html.replace(f"cid:{cid}", data_uri)
            except Exception as e:
                print(f"[WARN] Could not embed image {path}: {e}")
    return html

# def send_birthday_email(name, receiver_email, cc_list=None):
    # """Send email via Outlook Desktop. cc_list may be None, string, or list."""
    # try:
        # outlook = win32.Dispatch("Outlook.Application")
        # mail = outlook.CreateItem(0)

        # html_body = load_template(name)
        # html_body = embed_images_in_html(html_body)
        # bg_image_path = os.path.join(IMAGES_FOLDER, "birthday_bg.jpg")  # <-- rename accordingly

        # bg_image_path = os.path.abspath(os.path.join(IMAGES_FOLDER, "birthday_bg.jpg"))

        # if not os.path.exists(bg_image_path):
            # print("[ERROR] Background image NOT FOUND at:", bg_image_path)
        # else:
            # print("[OK] Background image found:", bg_image_path)
            # attachment = mail.Attachments.Add(bg_image_path)
            # attachment.PropertyAccessor.SetProperty(
                # "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                # "birthday_bg"
            # )

        # mail.To = receiver_email

        # if cc_list:
            # if isinstance(cc_list, list):
                # mail.CC = "; ".join(cc_list)
            # else:
                # mail.CC = cc_list

        # mail.Subject = f"🎉 Happy Birthday, {name}!"
        # mail.HTMLBody = html_body

        # # Send the mail
        # mail.Send()
        # print(f"[SENT] Email sent to {name} <{receiver_email}> ")
        # return True
    # except Exception as e:
        # print(f"[ERROR] Could not send email to {receiver_email}: {e}")
        # return False

def send_wish_email( name, receiver_email, subject, template_file, image_filename, image_cid, cc_list=None, years=None):
    SENDER_EMAIL = "cvaleti@opentext.com"

    try:
        html_body = load_template( template_file, name, years=years)

        msg = EmailMessage()

        msg["From"] = SENDER_EMAIL
        msg["To"] = receiver_email
        msg["Subject"] = subject

        # CC from Excel if present
        if cc_list is not None and not pd.isna(cc_list):
            cc_value = str(cc_list).strip()

            if cc_value:
                msg["Cc"] = cc_value

        # Plain-text fallback
        if years is not None:
            msg.set_content( f"Happy {years} Year Work Anniversary, {name}!" )
        else:
            msg.set_content( f"Happy Birthday, {name}!" )

        # HTML version
        msg.add_alternative( html_body, subtype="html" )

        image_path = os.path.join( IMAGES_FOLDER, image_filename )

        if not os.path.exists(image_path):
            print( f"[ERROR] Image NOT FOUND: {image_path}" )
            return False

        print(f"[OK] Image found: {image_path}")

        with open(image_path, "rb") as image_file:
            image_data = image_file.read()

        html_part = msg.get_payload()[-1]

        html_part.add_related( image_data, maintype="image", subtype="jpeg", cid=f"<{image_cid}>", filename=image_filename )

        with smtplib.SMTP( SMTP_SERVER, SMTP_PORT, timeout=30 ) as smtp:
            smtp.set_debuglevel(1)
            smtp.ehlo()
            smtp.send_message(msg)

        print( f"[SENT] {subject} -> " f"{name} <{receiver_email}>" )

        return True

    except Exception as e:
        print( f"[ERROR] Could not send email " f"to {receiver_email}: {e}" )
        return False

def send_birthday_email(name, receiver_email, cc_list=None):
    return send_wish_email( name=name, receiver_email=receiver_email,
        subject=f"🎉 Happy Birthday, {name}!", template_file= TEMPLATE_FILE,
        image_filename="birthday_bg.jpg", image_cid="birthday_bg",
        cc_list=cc_list
    )

def send_anniversary_email( name, receiver_email, years, cc_list=None):
    return send_wish_email( name=name, receiver_email=receiver_email,
        subject=f"🎉 Happy {years} Year Work Anniversary, {name}!", template_file=ANNIVERSARY_TEMPLATE_FILE,
        image_filename="anniversary_bg.jpg", image_cid="anniversary_bg",
        cc_list=cc_list, years=years
    )

def run_wishes_today():
    """Read Excel and send emails for birthdays matching today's month/day."""
    now = datetime.datetime.now()
    current_year = now.year

    try:
        df = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        print(f"[ERROR] Could not read Excel file '{EXCEL_FILE}': {e}")
        return None

    required_columns = [
        "Full Name",
        "Email Address",
        "Birthdate",
        "Joining Date"
    ]

    missing_columns = [ column
        for column in required_columns
            if column not in df.columns
    ]

    if missing_columns:
        print( "[ERROR] Missing Excel columns: " + ", ".join(missing_columns) )
        return None

    try:
        df["Birthdate"] = pd.to_datetime(df["Birthdate"])
        df["Joining Date"] = pd.to_datetime(df["Joining Date"])
    except Exception as e:
        print(f"[WARN] Could not parse Birthday column to datetimes: {e}")

    sent_any = False
    for _, row in df.iterrows():
        try:
            name = row.get("Full Name")
            print({name})
            if pd.isna(name):
                continue
            email = row.get("Email Address")
            if not email or not str(email).strip():
                print(f"[WARN] Missing email for {name}, skipping.")
                continue

            # Optionally read CC column if exists
            cc_in_row = row.get("CC") if "CC" in df.columns else None
            # -------------------------
            # Birthday check
            # -------------------------
            bday = row["Birthdate"]
            # If Birthday is NaT or missing, skip
            if not pd.isna(bday):
                if bday.month == now.month and bday.day == now.day:
                    success = send_birthday_email(str(name), str(email).strip(), cc_list=cc_in_row)
                    sent_any = sent_any or success
            # -------------------------
            # Work Anniversary check
            # -------------------------
            joining_date = row.get("Joining Date")
            if not pd.isna(joining_date):
                if joining_date.month == now.month and joining_date.day == now.day:
                    years = current_year - joining_date.year
                    if years >= 1:
                        success = send_anniversary_email(str(name), str(email).strip(), years, cc_list=cc_in_row)
                        sent_any = sent_any or success
        except Exception as e:
            print(f"[ERROR] Processing row failed: {e}")
    return sent_any

def main_loop():
    print("Birthday Agent - once per day runner started.")
    # print(f"CHECK_TIME = {CHECK_TIME}. State file = {STATE_FILE}")

    # parse CHECK_TIME into a time object
    # ht, mt = CHECK_TIME.split(":")
    # scheduled_time = datetime.time(int(ht), int(mt))

    # while True:
    now = datetime.datetime.now()
    today = now.date()
    last_run = read_last_run_date()

    print(f"[INFO] Scheduled time reached ({now.strftime('%Y-%m-%d %H:%M')}). Running birthday check...")
    try:
        result = run_wishes_today()
        # Record the run date regardless of whether any emails were sent, to avoid repeated runs
        if result is None:
            print( "[ERROR] Birthday/Anniversary check failed. " "Last-run date NOT updated." )
        elif result:
            write_last_run_date(today)
            print( "[INFO] Birthday/Anniversary check completed and email(s) sent." )
        else:
            write_last_run_date(today)
            print( "[INFO] Birthday/Anniversary check completed. No wishes to send today." )
    except Exception as e:
        print(f"[ERROR] Unexpected error during run: {e}")

if __name__ == "__main__":
    main_loop()
