Job Applications Email Tracker

A Python script to automatically track your job application emails from Gmail and export them to an Excel spreadsheet. The script:

Connects to your Gmail account via IMAP

Searches only the Primary category for messages sent after a specified date

Filters for application‐related messages (applied, interview invites, offers, rejections)

Deduplicates by IMAP UID and existing records

Extracts the email’s actual sent date, job title, company name, and status

Appends new entries to job_applications.xlsx

🚀 Features

Primary‐only fetch using Gmail’s X-GM-RAW "category:primary" search

Keyword filtering: include only application‐centric subjects, exclude noise (recruiter pitches, promotions)

Robust parsing: regex patterns to pull job title and company from email subjects

Deduplication: tracks existing UIDs and legacy entries to prevent duplicates

Excel output: stores data in a simple .xlsx file for easy review and reporting

🔧 Prerequisites

Python 3.7 or higher

Gmail account with IMAP enabled and an App Password (if using 2FA)

Required Python packages:

pip install openpyxl
pip install pandas

(Note: pandas is optional if you prefer built‐in Excel handling only.)

⚙️ Configuration

Edit the top of job_tracker.py and set:

EMAIL       = "your_email@gmail.com"
APP_PASSWORD= "your_app_password"
IMAP_SERVER = "imap.gmail.com"
EXCEL_FILE  = "job_applications.xlsx"
DATE_LIMIT  = "YYYY-MM-DD"   # e.g. "2024-10-01"

EMAIL: your Gmail address

APP_PASSWORD: an App Password generated under Google Account → Security → App passwords

DATE_LIMIT: ISO date after which messages are considered

📥 Installation

Clone or download the repository.

Install dependencies:

pip install openpyxl

Ensure your IMAP settings are enabled in Gmail Settings → Forwarding and POP/IMAP.

Save your App Password into the script configuration.

🎬 Usage

Run the script from your project folder:

python job_tracker.py

On first run, it will create job_applications.xlsx with headers:

Date

Job Title

Company

Status

Sender

UID

Subsequent runs will append only new, filtered messages.

⚙️ Customizing Filters

Inclusion keywords (line 50+): only subjects containing these terms will be considered.

Blacklist terms: exclude subjects matching any of these patterns.

Regex patterns: refine extract_job_info() with additional title/company patterns as needed.

🐞 Troubleshooting

No messages found?

Check DATE_LIMIT format and ensure there are matching emails in Primary.

Verify IMAP access and correct App Password.

Duplicates still appear?

Inspect the UID column in job_applications.xlsx; UIDs must be persistent.

Clear out legacy entries manually if needed.

Parsing errors?

Log the raw subject printed by the script and adjust regex patterns under extract_job_info().

📄 License

This project is released under the MIT License. Feel free to adapt and extend!
