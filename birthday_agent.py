"""
birthday_agent_outlook_once_per_day.py

Requirements:
    pip install pywin32 pandas openpyxl

Project layout:
    birthday_agent/
    ├── data.xlsx            # columns: Name, Email, Birthday (YYYY-MM-DD or Excel date)
    ├── template.html        # HTML template using {name} and {year}
    ├── images/
    │     ├── header_banner.png
    │     ├── confetti.png
    │     └── footer_logo.png
    └── birthday_agent_outlook_once_per_day.py  # this file

Behavior:
    - Checks once per day at CHECK_TIME (local system time, format "HH:MM")
    - Records last run date to last_run_date.txt (created in the same dir)
    - Uses Outlook Desktop to send HTML emails with embedded base64 images
"""

import os
import time
import datetime
import base64
import pandas as pd
import win32com.client as win32

# ---------- Configuration ----------
EXCEL_FILE = "TeamBirthDay.xlsx"
TEMPLATE_FILE = "template.html"
IMAGES_FOLDER = os.path.join(os.path.dirname(__file__), "images")
STATE_FILE = "last_run_date.txt"   # stores YYYY-MM-DD of last successful run
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

def load_template(name):
    """Load template.html and replace placeholders"""
    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace("{name}", name)
    html = html.replace("{year}", str(datetime.datetime.now().year))
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

def send_birthday_email(name, receiver_email, cc_list=None):
    """Send email via Outlook Desktop. cc_list may be None, string, or list."""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        html_body = load_template(name)
        html_body = embed_images_in_html(html_body)
        bg_image_path = os.path.join(IMAGES_FOLDER, "birthday_bg.jpg")  # <-- rename accordingly

        bg_image_path = os.path.abspath(os.path.join(IMAGES_FOLDER, "birthday_bg.jpg"))

        if not os.path.exists(bg_image_path):
            print("[ERROR] Background image NOT FOUND at:", bg_image_path)
        else:
            print("[OK] Background image found:", bg_image_path)
            attachment = mail.Attachments.Add(bg_image_path)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                "birthday_bg"
            )

        mail.To = receiver_email

        if cc_list:
            if isinstance(cc_list, list):
                mail.CC = "; ".join(cc_list)
            else:
                mail.CC = cc_list

        mail.Subject = f"🎉 Happy Birthday, {name}!"
        mail.HTMLBody = html_body

        # Send the mail
        mail.Send()
        print(f"[SENT] Email sent to {name} <{receiver_email}> ")
        return True
    except Exception as e:
        print(f"[ERROR] Could not send email to {receiver_email}: {e}")
        return False

def run_birthdays_today():
    """Read Excel and send emails for birthdays matching today's month/day."""
    now = datetime.datetime.now()
    today = now.date()

    try:
        df = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        print(f"[ERROR] Could not read Excel file '{EXCEL_FILE}': {e}")
        return False

    # Ensure Birthday column parsed
    if "Birthdate" not in df.columns:
        print(f"[ERROR] Excel must contain a 'Birthdate' column (found columns: {list(df.columns)})")
        return False

    try:
        df["Birthdate"] = pd.to_datetime(df["Birthdate"])
    except Exception as e:
        print(f"[WARN] Could not parse Birthday column to datetimes: {e}")

    sent_any = False
    for _, row in df.iterrows():
        try:
            bday = row["Birthdate"]
            # If Birthday is NaT or missing, skip
            if pd.isna(bday):
                continue
            if bday.month == now.month and bday.day == now.day:
                name = row.get("Full Name")
                email = row.get("Email Address")
                if not email or not str(email).strip():
                    print(f"[WARN] Missing email for {name}, skipping.")
                    continue
                # Optionally read CC column if exists
                cc_in_row = row.get("CC") if "CC" in df.columns else None
                success = send_birthday_email(str(name), str(email).strip(), cc_list=cc_in_row)
                sent_any = sent_any or success
        except Exception as e:
            print(f"[ERROR] Processing row failed: {e}")
    return sent_any

def main_loop():
    print("Birthday Agent (Outlook Desktop) - once per day runner started.")
    # print(f"CHECK_TIME = {CHECK_TIME}. State file = {STATE_FILE}")

    # parse CHECK_TIME into a time object
    # ht, mt = CHECK_TIME.split(":")
    # scheduled_time = datetime.time(int(ht), int(mt))

    # while True:
    now = datetime.datetime.now()
    today = now.date()
    last_run = read_last_run_date()

        # # If we've already run today, sleep until tomorrow morning to reduce work
        # if last_run == today:
            # # Sleep long (e.g., 30 minutes) to reduce CPU usage - we'll wake up tomorrow
            # time.sleep(1800)
            # continue

        # # If current local time is at or after scheduled_time and we haven't run today -> run
        # if now.time() >= scheduled_time and last_run != today:
    print(f"[INFO] Scheduled time reached ({now.strftime('%Y-%m-%d %H:%M')}). Running birthday check...")
    try:
        success = run_birthdays_today()
        # Record the run date regardless of whether any emails were sent, to avoid repeated runs
        write_last_run_date(today)
        if success:
            print(f"[INFO] Birthday check completed and emails sent ")
        else:
            print(f"[INFO] Birthday check completed. No emails sent or errors occurred.")
    except Exception as e:
        print(f"[ERROR] Unexpected error during run: {e}")
            # Still write last_run to avoid repeat spam; you can decide to not write on failure
        write_last_run_date(today)
        # else:
            # # Not yet time: sleep a short while and check again
            # time.sleep(20)

if __name__ == "__main__":
    main_loop()
