## 📂 What You'll Need

- **Outlook Desktop App** (for email extraction via COM)
- **Python 3.8+**
- **Google Drive API credentials** (`client_secret.json`)
- Dependencies:
  ```bash
  pip install -r requirements.txt

If you're using Outlook on Windows, you'll also need:
pip install pywin32==306

📦 Installation
1. Clone this repository
git clone https://github.com/Picklesasdf/outlook-transcript-exporter.git
cd outlook-transcript-exporter

2. Install dependencies
pip install -r requirements.txt
pip install pywin32==306  # Required for Outlook automation on Windows

3. Run the script
python email_transcript_exporter.py

3. When prompted:
- Enter your search keywords
- Choose where to save results
- Point to your Google Drive transcript folder
- (First time only) Authenticate your Google account

5.Done! Your PDFs will be saved and merged into output folders.

This version clearly separates `pywin32==306` as a Windows-specific requirement so users know exactly when to install it.

Let me know if you want me to help scaffold:
- A `config_template.json`
- The refactored `email_transcript_exporter.py` with user prompts
- Sample output folder structure or dummy PDFs for demo purposes
