diff --git a/README.md b/README.md
index c662cf30d9f09e3cc9c649c144ef25acc04835a4..45593382c15365e4a9f6d002816f788cbb8262d6 100644
--- a/README.md
+++ b/README.md
@@ -2,43 +2,43 @@
 
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
-- Choose where to save results
-- Point to your Google Drive transcript folder
+- Specify paths for your `client_secret.json` and token file (or accept the defaults)
+- Provide the Google Drive folder ID that contains your transcripts
 - (First time only) Authenticate your Google account
 
 5.Done! Your PDFs will be saved and merged into output folders.
 
 📌 Security Notes
 ✅ Everything runs locally — your emails, files, and transcripts never leave your machine
 🔐 You must authorize Google access yourself via the official OAuth window
 📁 You choose the folders where content is saved
 📄 License
 This project is licensed under the MIT License.
 You're free to use, modify, and share — just don’t hold me liable if something breaks.
 
 🙋 Need Help?
 If you're unsure about the code, paste it into ChatGPT and ask:
 “Can you explain what this script does and if it looks safe?”
 That's what I would do too.
