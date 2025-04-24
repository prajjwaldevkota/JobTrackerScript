# === HELPERS ===
import imaplib
import email
import re
import os
import openpyxl
from email.header import decode_header
from datetime import datetime
from email.utils import parsedate_to_datetime
from dotenv import load_dotenv

# Configuration constants
load_dotenv()
EMAIL = os.getenv("EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
IMAP_SERVER = "imap.gmail.com"
EXCEL_FILE = "job_applications.xlsx"
DATE_LIMIT = "24-Apr-2025"


def connect_email():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, APP_PASSWORD)
    return mail


def parse_subject(subj):
    try:
        decoded, enc = decode_header(subj)[0]
        return decoded.decode(enc or 'utf-8') if isinstance(decoded, bytes) else decoded
    except:
        return subj or ""


def extract_job_info(subject, sender):
    sl = subject.lower()

    # Status detection
    if any(k in sl for k in ["rejected", "unsuccessful", "not selected", "not moving forward", "not proceeding", "thank you for the interest"]):
        status = "Rejected"
    elif any(k in sl for k in ["offer", "congratulations", "welcome aboard", "pleased to offer"]):
        status = "Offer"
    elif any(k in sl for k in ["interview", "shortlisted", "next round", "next steps", "assessment"]):
        status = "Interview"
    else:
        status = "Applied"

    # Improved regex for job title extraction
    title_patterns = [
        r'for\s+["\']?(.*?)["\']?\s+at',  # "for [title] at"
        # "application for [title]"
        r'application\s+for\s+["\']?(.*?)["\']?\s+',
        r're:\s+["\']?(.*?)["\']?\s+application',  # "re: [title] application"
        r'position:?\s+["\']?(.*?)["\']?(?:\s|$)'  # "position: [title]"
    ]

    title = subject  # Default to full subject
    for pattern in title_patterns:
        tm = re.search(pattern, subject, re.IGNORECASE)
        if tm:
            title = tm.group(1)
            break

    # Improved company extraction
    company_patterns = [
        r'at\s+["\']?([\w\s&\-\.]+?)["\']?(?:$|:|\.|\s\()',  # "at [company]"
        # "from [company]"
        r'from\s+["\']?([\w\s&\-\.]+?)["\']?(?:$|:|\.|\s\()',
    ]

    comp = None
    for pattern in company_patterns:
        cm = re.search(pattern, subject, re.IGNORECASE)
        if cm:
            comp = cm.group(1)
            break

    # Fallback to email domain if no company found
    if not comp:
        try:
            email_parts = sender.split('<')[-1].split('>')[0].split('@')
            if len(email_parts) > 1:
                domain = email_parts[1]
                comp = domain.split('.')[0].capitalize()
        except:
            comp = "Unknown"

    return title.strip(), comp.strip(), status

# === MAIN ===


def fetch_jobs():
    # Prepare workbook and existing sets
    existing_uids = set()
    existing_keys = set()
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb.active
        headers = [c.value for c in sheet[1]]
        # ensure UID column
        if "UID" not in headers:
            sheet.cell(row=1, column=len(headers)+1, value="UID")
            wb.save(EXCEL_FILE)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Handle variable number of columns
            row_data = list(row) + [None] * (6 - len(row))
            date_val, title_val, comp_val, stat_val, sender_val, uid_val = row_data
            t = (title_val or "").strip().lower()
            c = (comp_val or "").strip().lower()
            s = (stat_val or "").strip()
            if uid_val:
                existing_uids.add(str(uid_val))
            else:
                existing_keys.add((t, c, s))
    else:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["Date", "Job Title", "Company",
                     "Status", "Sender", "UID"])
        wb.save(EXCEL_FILE)

    try:
        # connect and search
        mail = connect_email()
        mail.select("INBOX")
        typ, data = mail.uid('search', None,
                             f'X-GM-RAW "category:primary after:{DATE_LIMIT}"'
                             )
        if typ != 'OK':
            print("❌ IMAP search failed.")
            return
        uids = data[0].split()
        print(f"🔍 Found {len(uids)} messages in Primary since {DATE_LIMIT}")

        # Simple inclusion keywords to filter relevant emails
        include_kw = [
            "thank you for your application",
            "application received",
            "your application for",
            "we've received your application",
            "interview",
            "shortlisted",
            "internship",
            "service desk",
            "technical support",
            "support analyst",
            "application was sent",
            "Your online application has been successfully submitted",
            "Help Desk Analyst"
        ]

        # Blacklist - more specific patterns to exclude
        blacklist = [
            "visa", "ircc", "immigration",
            "30+ new tech internships posted this week",
            "job alert", "apply now",
            "sql interview challenge",
            "is hiring", "new internships",
            "internship", "hiring",
            "job opportunity", "one of the first",
            "be a great fit", "webinar",
            "newsletter", "unsubscribe",
            "career fair", "notification", "event", "first", "apply to", "actively recruiting", "you would be a great fit", "and more"
        ]

        new_rows = []
        for uid in uids:
            uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
            if uid_str in existing_uids:
                continue

            try:
                typ, msg_data = mail.uid('fetch', uid, '(RFC822)')
                if typ != 'OK' or not msg_data or not msg_data[0]:
                    continue

                msg = email.message_from_bytes(msg_data[0][1])

                gmail_labels = msg.get('X-Gmail-Labels', '')

                if any(category in gmail_labels for category in ["DRAFT", "SPAM", "PROMOTIONS", "UPDATES", "SOCIAL"]):
                    print(f"⏭️ Skipping email in category: {gmail_labels}")
                    continue

                subj = parse_subject(msg.get('Subject', ''))
                sl = subj.lower()
                sender = msg.get('From', '')

                # Check inclusion criteria - must have one of the keywords
                if not any(kw.lower() in sl for kw in include_kw):
                    continue

                # Improved blacklist filtering - check each word individually
                # Skip if ANY blacklist term matches
                blacklisted = False
                for term in blacklist:
                    if term.lower() in sl:
                        print(
                            f"⏭️ Skipping blacklisted: {subj} [matched: {term}]")
                        blacklisted = True
                        break

                if blacklisted:
                    continue

                # Parse actual sent date
                try:
                    dt = parsedate_to_datetime(msg.get('Date', ''))
                    date_str = dt.strftime('%Y-%m-%d')
                except:
                    date_str = datetime.now().strftime('%Y-%m-%d')

                title, comp, status = extract_job_info(subj, sender)
                key = (title.lower(), comp.lower(), status)
                if key in existing_keys:
                    continue

                new_rows.append(
                    (date_str, title, comp, status, sender, uid_str))
                existing_uids.add(uid_str)
                existing_keys.add(key)
                print(f"📌 New job: {title} at {comp} ({status}) - {subj}")

            except Exception as e:
                print(f"⚠️ Error processing email UID {uid_str}: {str(e)}")

        mail.logout()

        if new_rows:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            sheet = wb.active
            for row in new_rows:
                sheet.append(row)
            wb.save(EXCEL_FILE)
            print(f"✅ Appended {len(new_rows)} new jobs.")
        else:
            print("ℹ️ No new jobs to add.")

    except Exception as e:
        print(f"❌ Error: {str(e)}")


if __name__ == '__main__':
    fetch_jobs()
