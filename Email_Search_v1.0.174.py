# Filename: Email_Search_v1.0.174.py

# Import statements
import os
# Optional Windows COM imports
try:
    import win32com.client  # Outlook integration
    import pywintypes
except ImportError:
    win32com = None
    pywintypes = None
import re
import sys
import subprocess
import random
import shutil
import time
from datetime import datetime
from io import BytesIO
from tqdm import tqdm
from concurrent.futures import ProcessPoolExecutor, as_completed, TimeoutError
import fitz  # PyMuPDF (if needed for advanced PDF parsing)
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen.canvas import Canvas as rlc
from pypdf import PdfReader, PdfWriter
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.auth.exceptions import RefreshError
import argparse
import configparser


# --- Script Configuration ---
__version__ = 'v1.0.174'
SCRIPT_NAME = "Email_Search"

# --- Constants ---
MAX_ATTACHMENT_SIZE_MB = 40
MAX_ATTACHMENT_SIZE_BYTES = MAX_ATTACHMENT_SIZE_MB * 1024 * 1024
OCR_TIMEOUT_SECONDS = 60  # 1 minute per file
OCR_JOBS = os.cpu_count() or 1  # concurrency for OCR
MAX_SPLIT_SIZE_MB = 90
OCR_CHECK_PERCENTAGE = 0.05
OCR_CHECK_MAX_PAGES = 25
OCR_TEXT_THRESHOLD = 10

# Google Drive settings
GDRIVE_CLIENT_SECRET_FILE = ''
GDRIVE_TOKEN_FILE = ''
GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID = ''

SIGNATURE_IMAGE_EXTENSIONS = ('.png', '.gif', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff')
WORD_EXTENSIONS = ('.doc', '.docx')
EXCEL_EXTENSIONS = ('.xls', '.xlsx', '.xlsm')
# Attachment types to include for saving and conversion (Word, Excel, PDF)
ALLOWED_ATTACHMENT_EXTENSIONS = WORD_EXTENSIONS + EXCEL_EXTENSIONS + ('.pdf',)
EXCLUDED_FOLDERS = ['sent items', 'deleted items', 'junk e-mail', 'drafts', 'outbox']
 
# --- Configuration from INI ---
parser = argparse.ArgumentParser()
parser.add_argument('--config', default='config.ini', help='Path to INI configuration file')
args, _ = parser.parse_known_args()
CONFIG = configparser.ConfigParser()
CONFIG.read(args.config)

# Override defaults with config values
# Load and normalize base output directory (expand user and get absolute path)
_raw_base_dir = CONFIG.get('PATHS', 'base_output_dir',
                           fallback=os.path.join(os.path.expanduser('~'), 'Downloads'))
BASE_OUTPUT_DIR = os.path.abspath(os.path.expanduser(_raw_base_dir))
LOG_LEVEL = CONFIG.get('LOGGING', 'log_level', fallback='INFO')
# Email settings
OUTLOOK_EMAIL = CONFIG.get('EMAIL', 'outlook_email', fallback=None)
EXCLUDED_FOLDERS = [e.strip().lower() for e in CONFIG.get('EMAIL', 'excluded_folders', fallback=', '.join(EXCLUDED_FOLDERS)).split(',') if e.strip()]
PROCESS_ONLY_WITH_KEYWORDS = CONFIG.getboolean('EMAIL', 'process_only_with_keywords', fallback=True)
LIMIT_TO_DAYS_BACK = CONFIG.getint('EMAIL', 'limit_to_days_back', fallback=0)
# Attachment settings
ALLOWED_ATTACHMENT_EXTENSIONS = tuple(e.strip().lower() for e in CONFIG.get('ATTACHMENTS', 'allowed_extensions', fallback=', '.join(ALLOWED_ATTACHMENT_EXTENSIONS)).split(',') if e.strip())
CONVERT_OFFICE_DOCS = CONFIG.getboolean('ATTACHMENTS', 'convert_office_docs', fallback=True)
MAX_ATTACHMENT_SIZE_MB = CONFIG.getint('ATTACHMENTS', 'max_attachment_size_mb', fallback=MAX_ATTACHMENT_SIZE_MB)
MAX_ATTACHMENT_SIZE_BYTES = MAX_ATTACHMENT_SIZE_MB * 1024 * 1024
# PDF settings
SPLIT_EMAILS = CONFIG.getboolean('PDF', 'split_emails', fallback=True)
SPLIT_ATTACHMENTS = CONFIG.getboolean('PDF', 'split_attachments', fallback=True)
MAX_SPLIT_SIZE_MB = CONFIG.getint('PDF', 'max_split_size_mb', fallback=MAX_SPLIT_SIZE_MB)
OCR_REQUIRED = CONFIG.getboolean('PDF', 'ocr_required', fallback=True)
OCR_TIMEOUT_SECONDS = CONFIG.getint('PDF', 'ocr_timeout', fallback=OCR_TIMEOUT_SECONDS)
# Google Drive settings
GOOGLE_DRIVE_ENABLE = CONFIG.getboolean('GOOGLE_DRIVE', 'enable_transcript_download', fallback=True)
GDRIVE_CLIENT_SECRET_FILE = CONFIG.get('GOOGLE_DRIVE', 'client_secret_file', fallback=GDRIVE_CLIENT_SECRET_FILE)
GDRIVE_TOKEN_FILE = CONFIG.get('GOOGLE_DRIVE', 'token_file', fallback=GDRIVE_TOKEN_FILE)
GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID = CONFIG.get('GOOGLE_DRIVE', 'transcript_folder_id', fallback=GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID)

# Globals
keywords = []
BASE_FOLDER = EMAIL_SAVE_PATH = ATTACHMENT_SAVE_PATH = TRANSCRIPT_SAVE_PATH = LOG_FILE = None
CONSOLIDATED_EMAIL_PDF_PATH = CONSOLIDATED_ATTACHMENT_PDF_PATH = CONSOLIDATED_TRANSCRIPT_PDF_PATH = None
log_messages = []
# Index metadata storage
EMAIL_INDEX_LIST = []
ATTACHMENT_INDEX_LIST = []
TRANSCRIPT_INDEX_LIST = []

# Logging helper
def log(msg, worker=False):
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    prefix = '[WORKER] ' if worker else ''
    entry = f"[{ts}] {prefix}{msg}"
    log_messages.append(entry)
    # Always print to keep progress bar visible
    print(entry)

# Path initialization
def initialize_paths_and_logging():
    global keywords, BASE_FOLDER, EMAIL_SAVE_PATH, ATTACHMENT_SAVE_PATH, TRANSCRIPT_SAVE_PATH, LOG_FILE
    global CONSOLIDATED_EMAIL_PDF_PATH, CONSOLIDATED_ATTACHMENT_PDF_PATH, CONSOLIDATED_TRANSCRIPT_PDF_PATH
    global PROJECT_SAFE, DATE_STR
    # Load keywords from configuration
    try:
        keywords[:] = [k.strip().lower() for k in CONFIG.get('GENERAL', 'keywords').split(',') if k.strip()]
    except Exception:
        keywords[:] = []
    if not keywords:
        print("No keywords provided in config, exiting.")
        sys.exit(1)
    safe = '_'.join(re.sub(r'[^A-Za-z0-9_]', '', k) for k in keywords) or 'search'
    PROJECT_SAFE = safe
    # Use configured base output directory
    dl = BASE_OUTPUT_DIR
    BASE_FOLDER = os.path.join(dl, f"{SCRIPT_NAME}_{safe}")
    EMAIL_SAVE_PATH = os.path.join(BASE_FOLDER, 'Emails')
    ATTACHMENT_SAVE_PATH = os.path.join(BASE_FOLDER, 'Attachments')
    TRANSCRIPT_SAVE_PATH = os.path.join(BASE_FOLDER, 'Meeting_Transcripts')
    # Use MMDDYYYY format for consistency
    DATE_STR = datetime.now().strftime('%m%d%Y')
    date = DATE_STR
    LOG_FILE = os.path.join(BASE_FOLDER, f"{SCRIPT_NAME}_{safe}_Log_{__version__}.txt")
    CONSOLIDATED_EMAIL_PDF_PATH = os.path.join(BASE_FOLDER, f"Emails_{safe}_{date}.pdf")
    CONSOLIDATED_ATTACHMENT_PDF_PATH = os.path.join(BASE_FOLDER, f"Attachments_{safe}_{date}.pdf")
    CONSOLIDATED_TRANSCRIPT_PDF_PATH = os.path.join(BASE_FOLDER, f"Transcripts_{safe}_{date}.pdf")
    for path in [EMAIL_SAVE_PATH, ATTACHMENT_SAVE_PATH, TRANSCRIPT_SAVE_PATH]:
        os.makedirs(path, exist_ok=True)

# PDF validation
def is_valid_pdf(path):
    try:
        reader = PdfReader(path)
        return len(reader.pages) > 0
    except Exception:
        log(f"Invalid PDF: {os.path.basename(path)}")
        return False

# Merge PDFs
def merge_pdfs(paths, out_path):
    writer = PdfWriter()
    valid = [p for p in paths if os.path.exists(p) and is_valid_pdf(p)]
    if not valid:
        log(f"No PDFs to merge for {os.path.basename(out_path)}")
        return
    # Merge each PDF file into output, with progress bar
    for p in tqdm(valid,
                    desc=f"Merging {os.path.basename(out_path)}",
                    unit='file', position=1, leave=True):
        for page in PdfReader(p).pages:
            writer.add_page(page)
    try:
        with open(out_path, 'wb') as f:
            writer.write(f)
        log(f"Merged PDF: {out_path}")
    except Exception as e:
        log(f"Merge save failed: {e}")

# Split large PDF into parts
def split_pdf_by_size(path, max_mb=MAX_SPLIT_SIZE_MB):
    parts = []
    if os.path.getsize(path) <= max_mb * 1024 * 1024:
        return [path]
    reader = PdfReader(path)
    writer = PdfWriter()
    # Iterate pages with progress
    pages = reader.pages
    page_iter = tqdm(pages,
                     desc=f"Splitting {os.path.basename(path)}",
                     unit='page', position=1, leave=True)
    part_num = 1
    for page in page_iter:
        writer.add_page(page)
        temp = BytesIO()
        writer.write(temp)
        if temp.tell() >= max_mb * 1024 * 1024:
            out_path = f"{os.path.splitext(path)[0]}_part{part_num}.pdf"
            with open(out_path, 'wb') as f:
                writer.write(f)
            parts.append(out_path)
            part_num += 1
            writer = PdfWriter()
    if writer.pages:
        out_path = f"{os.path.splitext(path)[0]}_part{part_num}.pdf"
        with open(out_path, 'wb') as f:
            writer.write(f)
        parts.append(out_path)
    return parts
   
def update_attachment_index_after_split(part_paths, index_list):
    """
    Update attachment index entries with correct merged_file and start_page values
    after the merged attachments PDF has been split into parts.
    """
    # Compute page counts for each part
    part_page_counts = []
    for p in part_paths:
        try:
            count = len(PdfReader(p).pages)
        except Exception:
            count = 0
        part_page_counts.append(count)
    # Compute cumulative boundaries
    boundaries = []
    cum = 0
    for count in part_page_counts:
        cum += count
        boundaries.append(cum)
    # Update each entry with merged_file and start_page
    global_page = 0
    for entry in index_list:
        pc = entry.get('page_count', 0)
        start_global = global_page + 1
        # Determine part index for this start page
        part_idx = 0
        for idx, boundary in enumerate(boundaries):
            if start_global <= boundary:
                part_idx = idx
                break
        # Assign merged file name
        merged_filename = os.path.basename(part_paths[part_idx])
        entry['merged_file'] = merged_filename
        # Compute start page relative to the part
        prev_boundary = boundaries[part_idx - 1] if part_idx > 0 else 0
        entry['start_page'] = start_global - prev_boundary
        global_page += pc

# Save email as PDF, convert attachments
def save_email_as_pdf(item, out_path):
    mail = item
    try:
        rlc_canvas = rlc(out_path, pagesize=letter)
        # Determine recipients for 'to' field, with fallback for items lacking 'To' attribute
        if hasattr(mail, 'To'):
            to_field = mail.To
        elif hasattr(mail, 'Recipients'):
            recips = []
            try:
                for i in range(1, mail.Recipients.Count + 1):
                    r = mail.Recipients.Item(i)
                    name = getattr(r, 'Name', None)
                    if name:
                        recips.append(name)
                    else:
                        addr = getattr(r, 'Address', None)
                        if addr:
                            recips.append(addr)
                to_field = ', '.join(recips)
            except Exception:
                to_field = ''
        else:
            to_field = ''
        # Safely extract email metadata fields with fallbacks
        from_field = getattr(mail, 'SenderName', '') or ''
        subject_field = getattr(mail, 'Subject', '') or ''
        # SentOn may not exist (e.g., ReportItem); default to empty
        sent_on = ''
        sent_attr = getattr(mail, 'SentOn', None)
        if sent_attr:
            try:
                sent_on = sent_attr.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                sent_on = ''
        body_field = getattr(mail, 'Body', '') or ''
        details = {
            'from': from_field,
            'to': to_field,
            'subject': subject_field,
            'sent': sent_on,
            'body': body_field
        }
        y = 750
        def writeline(txt):
            nonlocal y
            for l in re.findall('.{1,100}(?:\\s+|$)', txt):
                if y < 50:
                    rlc_canvas.showPage()
                    rlc_canvas.setFont('Helvetica', 10)
                    y = 750
                rlc_canvas.drawString(50, y, l.strip())
                y -= 12
        for h in [
            f"From: {details['from']}",
            f"To: {details['to']}",
            f"Subject: {details['subject']}",
            f"Sent: {details['sent']}", '', 'Body:']:
            writeline(h)
        writeline(details['body'])
        rlc_canvas.save()
        return out_path
    except Exception as e:
        import traceback
        log(f"[save_email_as_pdf] FAILED for subject='{mail.Subject}': {e}\n{traceback.format_exc()}")
        return None

# Convert Office docs to PDF
 

def convert_office_to_pdf(path):
    ext = os.path.splitext(path)[1].lower()
    output = path.replace(ext, '.pdf')
    try:
        if ext in WORD_EXTENSIONS:
            word = win32com.client.Dispatch('Word.Application')
            try:
                doc = word.Documents.Open(path)
                doc.SaveAs(output, FileFormat=17)
                doc.Close()
            finally:
                word.Quit()
                del doc, word
        elif ext in EXCEL_EXTENSIONS:
            # Convert Excel workbooks to PDF using Workbook.ExportAsFixedFormat,
            # restricting to used data range to avoid empty colored columns
            excel = win32com.client.Dispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                # Open workbook in read-only mode
                wb = excel.Workbooks.Open(path, ReadOnly=1)
                # Set print area on each sheet to used range only
                for sheet in wb.Sheets:
                    try:
                        used = sheet.UsedRange
                        addr = used.Address
                        sheet.PageSetup.PrintArea = addr
                    except Exception:
                        continue
                # Export workbook to PDF respecting print areas
                wb.ExportAsFixedFormat(0, output)
                wb.Close(False)
            finally:
                excel.Quit()
                del wb, excel
        else:
            return None
        return output
    except Exception as e:
        log(f"Office convert failed: {e}")
        return None

# Fetch Outlook mail items
# Fetch Outlook mail items (search inbox and subfolders for keywords in subject or body)
# Fetch Outlook mail items
# Can filter by keywords or fetch all based on config
def get_all_mail_items(_keywords=None):
    # Initialize Outlook COM application
    try:
        outlook_app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    except Exception:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")
    # Select mailbox based on configured email address, if provided
    if OUTLOOK_EMAIL:
        # Retrieve Accounts as iterable (handles COM collection)
        try:
            accounts = list(namespace.Accounts)
        except (TypeError, AttributeError):
            accounts = []
            try:
                count = namespace.Accounts.Count
                for i in range(1, count + 1):
                    accounts.append(namespace.Accounts.Item(i))
            except Exception:
                accounts = []
        # Find matching account by SMTP address
        account = None
        for acc in accounts:
            try:
                smtp = getattr(acc, 'SmtpAddress', '') or ''
            except Exception:
                smtp = ''
            if smtp.lower() == OUTLOOK_EMAIL.lower():
                account = acc
                break
        if account:
            inbox = account.DeliveryStore.GetDefaultFolder(6)
        else:
            inbox = namespace.GetDefaultFolder(6)
    else:
        inbox = namespace.GetDefaultFolder(6)
    def search_folder(folder):
        items = []
        # Safely get items in folder
        try:
            all_items = list(folder.Items)
        except Exception as e:
            log(f"Error accessing items in folder {folder.Name}: {e}")
            all_items = []
        for item in all_items:
            # Skip items older than configured days back
            if LIMIT_TO_DAYS_BACK > 0:
                sent_attr = getattr(item, 'SentOn', None)
                if not sent_attr:
                    continue
                try:
                    if (datetime.now() - sent_attr).days > LIMIT_TO_DAYS_BACK:
                        continue
                except Exception:
                    continue
            try:
                subj = (item.Subject or "").lower()
                # prefer plaintext body, fallback to HTMLBody for HTML-only messages
                raw_body = (getattr(item, "Body", "") or "").strip()
                if not raw_body:
                    raw_body = (getattr(item, "HTMLBody", "") or "")
                body = raw_body.lower()
                # Include all if not limiting to keywords, else filter
                if (not PROCESS_ONLY_WITH_KEYWORDS) or any(kw in subj or kw in body for kw in keywords):
                    items.append(item)
            except Exception as e:
                log(f"Error processing item in folder {folder.Name}: {e}")
        # Recursively search subfolders
        for sub in folder.Folders:
            if sub.Name.lower() not in EXCLUDED_FOLDERS:
                items.extend(search_folder(sub))
        return items
    return search_folder(inbox)

# Process emails and attachments
def process_emails():
    # Collect metadata for index
    global EMAIL_INDEX_LIST, ATTACHMENT_INDEX_LIST
    EMAIL_INDEX_LIST.clear()
    ATTACHMENT_INDEX_LIST.clear()
    items = get_all_mail_items(keywords)
    email_pdfs = []
    attachments = []
    # Export matching emails to PDF with progress bar
    for itm in tqdm(items,
                      desc="Exporting Emails",
                      unit='email', position=1, leave=True):
        # Export email to PDF and record metadata
        out = save_email_as_pdf(
            itm,
            os.path.join(EMAIL_SAVE_PATH, f"email_{random.randint(1000,9999)}.pdf"),
        )
        if out and is_valid_pdf(out):
            email_pdfs.append(out)
            # Extract email metadata
            subject = getattr(itm, 'Subject', '') or ''
            sender = getattr(itm, 'SenderName', '') or getattr(itm, 'SenderEmailAddress', '') or ''
            sent_attr = getattr(itm, 'SentOn', None)
            sent_on = ''
            if sent_attr:
                try:
                    sent_on = sent_attr.strftime('%Y-%m-%d %H:%M:%S')
                except:
                    sent_on = ''
            EMAIL_INDEX_LIST.append({
                'source_filename': os.path.basename(out),
                'email_subject': subject,
                'sender': sender,
                'sent_on': sent_on
            })
        for att in itm.Attachments:
            fn = att.FileName
            ext = os.path.splitext(fn)[1].lower()
            if ext in SIGNATURE_IMAGE_EXTENSIONS:
                continue
            if ext not in ALLOWED_ATTACHMENT_EXTENSIONS:
                log(f"Skipping unsupported attachment type: {fn}")
                continue
            # ensure unique filename
            base, extension = os.path.splitext(fn)
            safe_base = re.sub(r'[\\/:"*?<>|]+', '_', base)
            dest = os.path.join(ATTACHMENT_SAVE_PATH, safe_base + extension)
            count = 1
            while os.path.exists(dest):
                dest = os.path.join(
                    ATTACHMENT_SAVE_PATH, f"{safe_base}_{count}{extension}"
                )
                count += 1
            try:
                att.SaveAsFile(dest)
                # Convert non-PDF files (Word/Excel) to PDF with timeout and progress
                if ext in WORD_EXTENSIONS + EXCEL_EXTENSIONS:
                    log(f"Converting attachment to PDF: {fn}")
                    start_conv = datetime.now()
                    try:
                        with ProcessPoolExecutor(max_workers=1) as exe:
                            future = exe.submit(convert_office_to_pdf, dest)
                            pdf_path = future.result(timeout=OCR_TIMEOUT_SECONDS)
                    except TimeoutError:
                        log(f"Office conversion timed out after {OCR_TIMEOUT_SECONDS}s: {fn}")
                        continue
                    except Exception as e:
                        log(f"Office conversion failed for {fn}: {e}")
                        continue
                    # Validate PDF output
                        if pdf_path and is_valid_pdf(pdf_path):
                            elapsed = datetime.now() - start_conv
                            log(f"Converted {fn} to PDF in {elapsed}")
                            attachments.append(pdf_path)
                            # Record attachment metadata with detailed info
                            subject = getattr(itm, 'Subject', '') or ''
                            sender = getattr(itm, 'SenderName', '') or getattr(itm, 'SenderEmailAddress', '') or ''
                            sent_attr = getattr(itm, 'SentOn', None)
                            sent_on = ''
                            if sent_attr:
                                try:
                                    sent_on = sent_attr.strftime('%Y-%m-%d %H:%M:%S')
                                except:
                                    sent_on = ''
                            try:
                                page_count = len(PdfReader(pdf_path).pages)
                            except Exception:
                                page_count = 0
                            ATTACHMENT_INDEX_LIST.append({
                                'source_filename': os.path.basename(pdf_path),
                                'attachment_name': fn,
                                'email_subject': subject,
                                'sender': sender,
                                'sent_on': sent_on,
                                'page_count': page_count,
                                'start_page': 0,
                                'merged_file': os.path.basename(CONSOLIDATED_ATTACHMENT_PDF_PATH)
                            })
                    else:
                        log(f"Office convert produced invalid PDF for attachment: {fn}")
                else:
                    # PDF attachment
                    attachments.append(dest)
                    # Record attachment metadata with detailed info
                    subject = getattr(itm, 'Subject', '') or ''
                    sender = getattr(itm, 'SenderName', '') or getattr(itm, 'SenderEmailAddress', '') or ''
                    sent_attr = getattr(itm, 'SentOn', None)
                    sent_on = ''
                    if sent_attr:
                        try:
                            sent_on = sent_attr.strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            sent_on = ''
                    try:
                        page_count = len(PdfReader(dest).pages)
                    except Exception:
                        page_count = 0
                    ATTACHMENT_INDEX_LIST.append({
                        'source_filename': os.path.basename(dest),
                        'attachment_name': fn,
                        'email_subject': subject,
                        'sender': sender,
                        'sent_on': sent_on,
                        'page_count': page_count,
                        'start_page': 0,
                        'merged_file': os.path.basename(CONSOLIDATED_ATTACHMENT_PDF_PATH)
                    })
            except Exception as e:
                log(f"Attachment save failed: {e}")
    return email_pdfs, attachments

# Google Drive download
def download_google_docs_from_drive(keywords, out_dir):
    creds = None
    if os.path.exists(GDRIVE_TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(GDRIVE_TOKEN_FILE)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(GDRIVE_CLIENT_SECRET_FILE, 
                      ['https://www.googleapis.com/auth/drive.readonly'])
            creds = flow.run_local_server(port=0)
        with open(GDRIVE_TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    service = build('drive', 'v3', credentials=creds)
    query = "mimeType='application/pdf'"
    files = service.files().list(q=query, fields="files(id,name)").execute().get('files', [])
    downloaded = []
    # Download matching transcripts with progress bar
    for f in tqdm(files,
                    desc="Downloading Transcripts",
                    unit='file', position=1, leave=True):
        if any(kw in f['name'].lower() for kw in keywords):
            request = service.files().get_media(fileId=f['id'])
            fh = BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
            path = os.path.join(out_dir, f['name'])
            with open(path, 'wb') as out_f:
                out_f.write(fh.getvalue())
            downloaded.append(path)
    return downloaded

# OCR status checker
def check_ocr_status(path):
    try:
        reader = PdfReader(path)
        text = ''.join(page.extract_text() or '' for page in reader.pages[:OCR_CHECK_MAX_PAGES])
        return len(text) > OCR_TEXT_THRESHOLD
    except Exception:
        return False

# OCR PDF task wrapper
def ocr_pdf_task(path):
    out_path = path.replace('.pdf', '_ocr.pdf')
    try:
        cmd = ['ocrmypdf', '--jobs', str(OCR_JOBS), '--timeout', str(OCR_TIMEOUT_SECONDS), path, out_path]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        ok = check_ocr_status(out_path)
        if ok:
            return True, out_path
        else:
            return False, path
    except Exception:
        return False, path
    
# Project index builder
def build_project_index(emails_dir='emails', attachments_dir='attachments', transcripts_dir='transcripts', output_csv='project_index.csv'):
    import csv, os
    global EMAIL_INDEX_LIST, ATTACHMENT_INDEX_LIST, TRANSCRIPT_INDEX_LIST
    global CONSOLIDATED_EMAIL_PDF_PATH, CONSOLIDATED_ATTACHMENT_PDF_PATH, CONSOLIDATED_TRANSCRIPT_PDF_PATH
    # Basenames for merged files
    email_merged_basename = os.path.basename(CONSOLIDATED_EMAIL_PDF_PATH)
    attachment_merged_basename = os.path.basename(CONSOLIDATED_ATTACHMENT_PDF_PATH)
    transcript_merged_basename = os.path.basename(CONSOLIDATED_TRANSCRIPT_PDF_PATH)
    # CSV headers
    headers = ['type', 'source_filename', 'email_subject', 'sender', 'sent_on', 'attachment_name', 'transcript_subject', 'meeting_date', 'page_count', 'start_page', 'merged_file']
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(headers)
        # Emails
        start_page = 1
        for entry in EMAIL_INDEX_LIST:
            fname = entry.get('source_filename', '')
            path = os.path.join(emails_dir, fname)
            try:
                reader = PdfReader(path)
                page_count = len(reader.pages)
            except Exception:
                page_count = 0
            writer.writerow([
                'email', fname,
                entry.get('email_subject', ''),
                entry.get('sender', ''),
                entry.get('sent_on', ''),
                '', '', '',
                page_count, start_page,
                email_merged_basename
            ])
            start_page += page_count
        # Attachments
        for entry in ATTACHMENT_INDEX_LIST:
            writer.writerow([
                'attachment',
                entry.get('source_filename', ''),
                '', '', '',
                entry.get('attachment_name', ''),
                '', '',
                entry.get('page_count', 0),
                entry.get('start_page', 0),
                entry.get('merged_file', attachment_merged_basename)
            ])
        # Transcripts
        start_page = 1
        for entry in TRANSCRIPT_INDEX_LIST:
            fname = entry.get('source_filename', '')
            path = os.path.join(transcripts_dir, fname)
            try:
                reader = PdfReader(path)
                page_count = len(reader.pages)
            except Exception:
                page_count = 0
            writer.writerow([
                'transcript', fname,
                '', '', '', '',
                entry.get('transcript_subject', ''),
                entry.get('meeting_date', ''),
                page_count, start_page,
                transcript_merged_basename
            ])
            start_page += page_count

# Updated transcript processing
def process_transcripts():
    # Collect transcript metadata for index
    global TRANSCRIPT_INDEX_LIST
    TRANSCRIPT_INDEX_LIST.clear()
    paths = download_google_docs_from_drive(keywords, TRANSCRIPT_SAVE_PATH)
    valid = [p for p in paths if is_valid_pdf(p)]
    # Build transcript metadata entries
    for p in valid:
        fname = os.path.basename(p)
        base = os.path.splitext(fname)[0]
        # Extract date pattern YYYY-MM-DD
        m = re.search(r"(\d{4}-\d{2}-\d{2})", base)
        if m:
            meeting_date = m.group(1)
            subj = base.replace(meeting_date, '').rstrip('_- ')
        else:
            meeting_date = ''
            subj = base
        # Human-friendly subject
        transcript_subject = subj.replace('_', ' ').strip()
        TRANSCRIPT_INDEX_LIST.append({
            'source_filename': fname,
            'transcript_subject': transcript_subject,
            'meeting_date': meeting_date
        })
    if valid:
        merge_pdfs(valid, CONSOLIDATED_TRANSCRIPT_PDF_PATH)
        # Split transcripts PDF if over size limit
        parts = split_pdf_by_size(CONSOLIDATED_TRANSCRIPT_PDF_PATH)
        if len(parts) > 1:
            log(f"Split merged transcripts into {len(parts)} parts under {MAX_SPLIT_SIZE_MB}MB")
            for p in parts:
                log(f"Transcript part: {p}")
    return valid

# Main execution
if __name__ == '__main__':
    initialize_paths_and_logging()
    log(f"--- {SCRIPT_NAME} {__version__} STARTED ---")
    overall_start = datetime.now()

    # Setup overall progress bar
    total_stages = 3
    stage_pct = 100 / total_stages
    overall_bar = tqdm(total=100, desc="Overall Progress", position=0, leave=True)

    # Stage 1: Email processing
    stage1_start = datetime.now()
    emails, atts = process_emails()
    stage1_end = datetime.now()
    stage1_elapsed = stage1_end - stage1_start
    overall_bar.update(stage_pct)

    # Stage 2: Transcript processing
    stage2_start = datetime.now()
    if GOOGLE_DRIVE_ENABLE:
        trans_paths = process_transcripts()
    else:
        log("Transcript download disabled by config.")
        trans_paths = []
    stage2_end = datetime.now()
    stage2_elapsed = stage2_end - stage2_start
    overall_bar.update(stage_pct)

    # Stage 3: Attachment OCR / processing
    stage3_start = datetime.now()
    attachments_to_merge, failures = [], []
    total_atts = len(atts)
    if total_atts == 0:
        log("ℹ️ No attachments to process.")
    else:
        # Process attachments with per-file progress bar
        with tqdm(atts,
                  desc="Attachment OCR/Processing",
                  unit='file', position=1, leave=True) as attach_bar:
            for pdf in attach_bar:
                # If OCR not required, include all attachments as-is
                if not OCR_REQUIRED:
                    attachments_to_merge.append(pdf)
                    overall_bar.update(stage_pct / total_atts)
                    continue
                name = os.path.basename(pdf)
                attach_bar.set_postfix_str(name)
                log(f"Processing attachment: {name}")
                # If PDF already contains text, skip OCR
                if is_valid_pdf(pdf) and check_ocr_status(pdf):
                    log(f"{name}: existing text detected, skipping OCR")
                    attachments_to_merge.append(pdf)
                else:
                    # Perform OCR
                    if shutil.which('ocrmypdf') is None:
                        log(f"ocrmypdf not found, cannot OCR: {name}")
                        failures.append(pdf)
                    else:
                        log(f"{name}: performing OCR (timeout {OCR_TIMEOUT_SECONDS}s)")
                        try:
                            ok, outp = ocr_pdf_task(pdf)
                            if ok:
                                log(f"{name}: OCR succeeded -> {os.path.basename(outp)}")
                                attachments_to_merge.append(outp)
                            else:
                                log(f"{name}: OCR failed, skipping merge")
                                failures.append(pdf)
                        except Exception as e:
                            log(f"OCR exception for {name}: {e}")
                            failures.append(pdf)
                # Update overall progress
                overall_bar.update(stage_pct / total_atts)
    # End of attachments processing timing
    stage3_end = datetime.now()
    stage3_elapsed = stage3_end - stage3_start
    # Close overall progress bar at end of processing
    overall_bar.close()

    # Merge and finalize
    # Merge emails and split if necessary
    if emails:
        merge_pdfs(emails, CONSOLIDATED_EMAIL_PDF_PATH)
        # Split emails PDF if over size limit
        email_parts = split_pdf_by_size(CONSOLIDATED_EMAIL_PDF_PATH)
        if len(email_parts) > 1:
            log(f"Split merged emails into {len(email_parts)} parts under {MAX_SPLIT_SIZE_MB}MB")
            for p in email_parts:
                log(f"Email part: {p}")
    else:
        log("ℹ️ No email PDFs to merge.")

    if attachments_to_merge:
        merge_pdfs(attachments_to_merge, CONSOLIDATED_ATTACHMENT_PDF_PATH)
        # Split attachments or retain single file based on config
        if SPLIT_ATTACHMENTS:
            parts = split_pdf_by_size(CONSOLIDATED_ATTACHMENT_PDF_PATH)
            if len(parts) > 1:
                log(f"Split merged attachments into {len(parts)} parts under {MAX_SPLIT_SIZE_MB}MB")
                for p in parts:
                    log(f"Attachment part: {p}")
        else:
            parts = [CONSOLIDATED_ATTACHMENT_PDF_PATH]
        # Update attachment index entries after merge/split
        update_attachment_index_after_split(parts, ATTACHMENT_INDEX_LIST)
    else:
        log("ℹ️ No attachments merged.")

    overall_end = datetime.now()
    overall_elapsed = overall_end - overall_start
    # Compute OCR success count for attachments
    success_count = len(attachments_to_merge)

    # Summary output
    print("\n=== Processing Summary ===")
    print(f"Emails downloaded: {len(emails)}")
    print(f"Emails merged: {len(emails)}")
    print(f"Attachments downloaded: {len(atts)}")
    print(f"Attachments OCR succeeded: {success_count}")
    print(f"Attachments OCR failed: {len(failures)}")
    print(f"Attachments merged: {len(attachments_to_merge)}")
    print(f"Transcripts downloaded: {len(trans_paths)}")
    print(f"Transcripts merged: {len(trans_paths)}")
    if failures:
        print("\nFailed OCR attachments:")
        for f in failures:
            print(f" - {os.path.basename(f)}")
    # Print timing information
    print("\nStage durations:")
    print(f" - Email processing: {stage1_elapsed}")
    print(f" - Transcript processing: {stage2_elapsed}")
    print(f" - Attachment processing: {stage3_elapsed}")
    print(f"Total runtime: {overall_elapsed}")
    
    # Log summary to log_messages
    log("=== Processing Summary ===")
    log(f"Emails downloaded: {len(emails)}")
    log(f"Emails merged: {len(emails)}")
    log(f"Attachments downloaded: {len(atts)}")
    log(f"Attachments OCR succeeded: {success_count}")
    log(f"Attachments OCR failed: {len(failures)}")
    log(f"Attachments merged: {len(attachments_to_merge)}")
    log(f"Transcripts downloaded: {len(trans_paths)}")
    log(f"Transcripts merged: {len(trans_paths)}")
    if failures:
        log("Failed OCR attachments:")
        for f in failures:
            log(f" - {os.path.basename(f)}")
    # Log timing information
    log("Stage durations:")
    log(f" - Email processing: {stage1_elapsed}")
    log(f" - Transcript processing: {stage2_elapsed}")
    log(f" - Attachment processing: {stage3_elapsed}")
    log(f"Total runtime: {overall_elapsed}")

    # Write log to file
    with open(LOG_FILE, 'w', encoding='utf-8') as lf:
        lf.write("\n".join(log_messages))
    # Build project index CSV for merged PDFs
    try:
        # Generate index for individual PDFs in email, attachment, and transcript folders
        # Name index CSV with project and date
        output_csv = os.path.join(BASE_FOLDER, f"project_index_{PROJECT_SAFE}_{DATE_STR}.csv")
        build_project_index(
            emails_dir=EMAIL_SAVE_PATH,
            attachments_dir=ATTACHMENT_SAVE_PATH,
            transcripts_dir=TRANSCRIPT_SAVE_PATH,
            output_csv=output_csv
        )
        msg = f"Project index generated: {output_csv}"
        print(msg)
        log(msg)
    except Exception as e:
        # Report failure with dynamic index filename
        err = f"Failed to generate {os.path.basename(output_csv)}: {e}"
        print(err)
        log(err)
    # Cleanup temporary files and folders
    try:
        shutil.rmtree(EMAIL_SAVE_PATH)
        shutil.rmtree(ATTACHMENT_SAVE_PATH)
        shutil.rmtree(TRANSCRIPT_SAVE_PATH)
        log("Cleaned up temporary files")
    except Exception as e:
        log(f"Cleanup failed: {e}")
