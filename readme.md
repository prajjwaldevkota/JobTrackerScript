# Job Applications Email Tracker 📬

A Python script to automatically track your job application emails from Gmail and export them to an Excel spreadsheet.

## 📦 What It Does

- **Connects** to your Gmail account via IMAP  
- **Searches only** the **Primary** category for messages sent after a specified date  
- **Filters** for application-related messages (applied, interview invites, offers, rejections)  
- **Deduplicates** by IMAP UID and existing records  
- **Extracts** the email’s actual sent date, job title, company name, and status  
- **Appends** new entries to `job_applications.xlsx`  

---

## 🚀 Features

- **Primary-only fetch** using Gmail’s `X-GM-RAW "category:primary"` search  
- **Keyword filtering**: include only application-centric subjects, exclude noise (e.g. recruiter pitches, newsletters)  
- **Robust parsing**: regex patterns to extract job title and company from email subject lines  
- **Deduplication**: avoids reprocessing messages using persistent IMAP UID and historical record comparison  
- **Excel output**: stores structured data in a `.xlsx` file for easy review and tracking  

---

## 🔧 Prerequisites

- Python **3.7 or higher**
- Gmail account with **IMAP enabled**
- A Gmail **App Password** (required if using two-factor authentication)

### Required Python packages

```bash
pip install openpyxl
pip install pandas


