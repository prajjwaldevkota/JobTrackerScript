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
```
Pandas is optional

## ⚙️ Configuration

Edit the top of `job_tracker.py` and set the following variables:

```python
EMAIL       = "your_email@gmail.com"
APP_PASSWORD= "your_app_password"
IMAP_SERVER = "imap.gmail.com"
EXCEL_FILE  = "job_applications.xlsx"
DATE_LIMIT  = "YYYY-MM-DD"   # e.g. "2024-10-01"
```

## 📥 Installation

1.  **Clone or download the repository:** Obtain the `job_tracker.py` script and any associated files. You can do this by cloning the Git repository if available, or by downloading the files directly.

2.  **Install dependencies:** Open your terminal or command prompt, navigate to the directory where you saved the `job_tracker.py` file, and run the following command to install the necessary library:

    ```bash
    pip install openpyxl
    ```

    This command will install the `openpyxl` library, which is required for the script to work with Excel files.

3.  **Enable IMAP in Gmail:**
    * Go to your Gmail settings. You can usually find this by clicking the gear icon in the top right corner and then selecting "See all settings".
    * Navigate to the **Forwarding and POP/IMAP** tab.
    * Ensure that **IMAP access** is enabled. If it's disabled, select "Enable IMAP" and click "Save Changes".

4.  **Generate and save App Password:**
    * Go to your [Google Account](https://myaccount.google.com/).
    * Navigate to the **Security** section.
    * Under "How you sign in," find and click on **App passwords**. (You might need to have 2-Step Verification enabled for this option to appear.)
    * Select "Mail" from the "Select the app" dropdown.
    * Select "Other (Custom name)" from the "Select the device" dropdown and enter a name like "Job Tracker".
    * Click **Generate**. This will provide you with a 16-digit App Password.
    * **Copy this App Password.**
    * Open the `job_tracker.py` file in a text editor.
    * Locate the line `APP_PASSWORD= "your_app_password"` and replace `"your_app_password"` with the App Password you just generated, ensuring you keep the quotation marks:

        ```python
        APP_PASSWORD= "your_generated_app_password"
        ```
        or add an env file and you can change it there.
    * Save the `job_tracker.py` file.
  
    ## 🎬 Usage

Run the script from your project folder using the following command in your terminal or command prompt:

```bash
python job_tracker.py
```

## 💡 Feel Free to Contribute
If you encounter any issues, have suggestions for improvements, or would like to add new features to this job tracking script, please feel free to contribute! You can fork the repository (if this is shared on a platform like GitHub) and submit pull requests with your changes. Your contributions are welcome and can help make this tool even better for everyone.

  
    
