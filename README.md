# Email Search Script (`Email_Search_v1.0.174.py`)

This script automatically searches Microsoft Outlook emails based on provided keywords, saves matched emails and attachments as PDFs, performs OCR if needed, and optionally downloads matching transcripts from Google Drive. All configurations are handled through a `config.ini` file.

---

## ✅ Requirements

- Windows machine with Outlook installed and configured
- Python 3.10+
- Required Python libraries (see installation below)
- Optional: `ocrmypdf` for scanned PDF processing

---

## 📦 Installation

### 1. Open Command Prompt

```bash
cd path\to\your\script\folder
```

### 2. Install Python packages

```bash
pip install -r requirements.txt
```

### 3. Optional: Install `ocrmypdf` (needed for scanned attachments)

Using Chocolatey on Windows:
```bash
choco install ocrmypdf
```

Or on Linux/WSL:
```bash
sudo apt install ocrmypdf
```

---

## 🛠️ Configuration (`config.ini`)

Edit the `config.ini` file to control behavior.

### [GENERAL]
```ini
keywords = what, ever, you, want, to, search, for
```

### [EMAIL]
```ini
outlook_email = your.email@domain.com  ; Leave blank for default Outlook profile
limit_to_days_back = 0                ; Only emails newer than X days (0 = no limit)
process_only_with_keywords = yes      ; yes = only match keywords
```

### [ATTACHMENTS]
```ini
allowed_extensions = .pdf, .docx, .xlsx
convert_office_docs = yes
max_attachment_size_mb = 40
```

### [PDF]
```ini
split_emails = yes
split_attachments = yes
max_split_size_mb = 90
ocr_required = yes
ocr_timeout = 60
```

### [GOOGLE_DRIVE]
```ini
enable_transcript_download = yes
client_secret_file = client_secret.json
token_file = token.json
transcript_folder_id = your_folder_id
```

### [PATHS]
```ini
base_output_dir = ~/Downloads
```

---

## ▶️ Running the Script

Run the script with the config file like this:

```bash
python Email_Search_v1.0.174.py --config config.ini
```

---

## 📂 Output

Output will be saved to:

```
<base_output_dir>/Email_Search_<keywords>/
```

Includes:
- `Emails_*.pdf` — merged emails
- `Attachments_*.pdf` — merged and OCR-processed attachments
- `Transcripts_*.pdf` — Google Drive transcripts (optional)
- `project_index_*.csv` — master index of all documents
- `*_Log_*.txt` — log file with all operations

---

## 🧪 Testing Tips

- Use simple keywords (like your name) for initial runs.
- Set `limit_to_days_back = 3` to test only recent emails.
- Disable transcripts (`enable_transcript_download = no`) during testing.

---

