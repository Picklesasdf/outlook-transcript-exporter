# Email & Transcript Exporter 📨📄

A Python tool to **search Outlook emails**, **download meeting transcripts from Google Drive**, and **export everything into organized, searchable PDF reports** — including **OCR (text recognition)** for easy searchability.

---

## 🔧 Features

- 📥 Searches **Outlook emails** using your provided keywords
- 📎 Exports **emails and attachments** as PDFs
- 📄 Downloads **Google Docs transcripts** from a Drive folder
- 📚 Merges all content into categorized PDF bundles
- 🔍 Runs **OCR** (optional) so you can search inside the PDFs
- 💬 Prompts you for inputs — no hardcoded paths or credentials

---

## 📂 What You'll Need

- **Outlook Desktop App** (for email extraction via COM)
- **Python 3.8+**
- **Google Drive API credentials** (`client_secret.json`)
- Dependencies:
  ```bash
  pip install -r requirements.txt
