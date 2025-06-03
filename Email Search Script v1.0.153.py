import os
import win32com.client
import re
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter
from io import BytesIO
from tqdm import tqdm
import subprocess
import shutil
import sys
import fitz  # PyMuPDF
import pywintypes # Ensure this is imported if not already for pywintypes.com_error
import concurrent.futures
import traceback
import random

# --- Script Configuration ---
__version__ = 'v1.0.153' # Version updated
SCRIPT_NAME = "Email_Search"

# --- Setup ---
MAX_SPLIT_SIZE_MB = 90
OCR_TIMEOUT_SECONDS = 1800 # 30 minutes per file for OCR
MAX_ATTACHMENT_SIZE_MB = 40
MAX_ATTACHMENT_SIZE_BYTES = MAX_ATTACHMENT_SIZE_MB * 1024 * 1024
SPLIT_CHECK_INTERVAL = 25 # Check PDF part size every 25 pages during splitting
OCR_CHECK_PERCENTAGE = 0.05 # Check 5% of pages for existing text
OCR_CHECK_MAX_PAGES = 25 # Max number of pages to check for existing text
OCR_MIN_TEXT_LENGTH_THRESHOLD = 10 # Min text length on a page to consider it OCR'd

GDRIVE_CLIENT_SECRET_FILE = "client_secret.json" # Path to your Google client_secret.json
GDRIVE_TOKEN_FILE = 'token.json' # Path where the token will be stored
GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID = "1y84cNQAnSsr7UYvK84TXH4MmW-GZmuC8" # Specific Google Drive Folder ID

SIGNATURE_IMAGE_EXTENSIONS = ('.png', '.gif', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff')
WORD_EXTENSIONS = ('.doc', '.docx')
EXCEL_EXTENSIONS = ('.xls', '.xlsx', '.xlsm')

# --- Global Variables for Logging and Paths ---
USE_TQDM_FOR_LOGGING = False # Set to True to use tqdm.write for all logs, False for print

keywords_input = ""
keywords = []
BASE_FOLDER = ""
EMAIL_SAVE_PATH = ""
ATTACHMENT_SAVE_PATH = ""
TRANSCRIPT_SAVE_PATH = ""
LOG_FILE = ""
CONSOLIDATED_EMAIL_PDF_PATH = ""
CONSOLIDATED_ATTACHMENT_PDF_PATH = ""
CONSOLIDATED_TRANSCRIPT_PDF_PATH = ""

# Outlook Folder Types (constants from OlDefaultFolders enumeration)
olFolderSentMail = 5
olFolderDeletedItems = 3
olFolderOutbox = 4
olFolderDrafts = 16
olFolderJunk = 23 # Often referred to as Junk E-mail

EXCLUDED_FOLDER_TYPES = [
    olFolderSentMail, olFolderDeletedItems, olFolderOutbox,
    olFolderDrafts, olFolderJunk
]
EXCLUDED_FOLDER_NAMES_LOWER = [ # Case-insensitive list of folder names to exclude
    "sent items", "deleted items", "junk e-mail", "junk email", "drafts", "outbox",
    "archive", "archives", "conversation history", "rss feeds", "sync issues", "clutter"
]

log_messages = []
def log(message, is_worker_log=False):
    """Logs messages to a global list and prints them."""
    global USE_TQDM_FOR_LOGGING
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prefix = "[WORKER] " if is_worker_log else ""
    formatted_message = f"[{current_time}] {prefix}{message}"
    log_messages.append(formatted_message)

    if is_worker_log: # Worker logs should always print directly for immediate visibility
        print(formatted_message, flush=True)
    elif USE_TQDM_FOR_LOGGING:
        tqdm.write(formatted_message, file=sys.stdout)
    else:
        print(formatted_message)

def initialize_paths_and_logging():
    """Initializes global paths and sets up the base folder and log file name."""
    global keywords_input, keywords, BASE_FOLDER, EMAIL_SAVE_PATH, ATTACHMENT_SAVE_PATH
    global TRANSCRIPT_SAVE_PATH, LOG_FILE, CONSOLIDATED_EMAIL_PDF_PATH
    global CONSOLIDATED_ATTACHMENT_PDF_PATH, CONSOLIDATED_TRANSCRIPT_PDF_PATH

    keywords_input = input("Enter keywords to search for (separated by commas): ")
    keywords = [kw.strip().lower() for kw in keywords_input.split(',') if kw.strip()]

    if not keywords:
        print("No valid keywords entered. Exiting.")
        sys.exit(1)

    # Create a file-system-safe string from keywords for folder/file naming
    safe_keyword_string = '_'.join(filter(None, (re.sub(r'[^a-zA-Z0-9_-]', '', kw) for kw in keywords)))
    if not safe_keyword_string: # Fallback if all keywords become empty after sanitizing
        safe_keyword_string = "search"
        print("Keywords resulted in an empty string after sanitization. Using 'search' for folder name.")


    HOME_DOWNLOADS_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
    BASE_FOLDER = os.path.join(HOME_DOWNLOADS_PATH, f"{SCRIPT_NAME}_{safe_keyword_string}")

    EMAIL_SAVE_PATH = os.path.join(BASE_FOLDER, "Emails")
    ATTACHMENT_SAVE_PATH = os.path.join(BASE_FOLDER, "Attachments")
    TRANSCRIPT_SAVE_PATH = os.path.join(BASE_FOLDER, "Meeting_Transcripts")

    LOG_FILE = os.path.join(BASE_FOLDER, f"{SCRIPT_NAME}_{safe_keyword_string}_Log_{__version__}.txt")

    # Consolidated PDF paths
    current_date_str = datetime.now().strftime('%d%m%Y')
    CONSOLIDATED_EMAIL_PDF_PATH = os.path.join(BASE_FOLDER, f"Emails_Complete_{safe_keyword_string}_{current_date_str}.pdf")
    CONSOLIDATED_ATTACHMENT_PDF_PATH = os.path.join(BASE_FOLDER, f"{safe_keyword_string}_Attachments_{current_date_str}.pdf")
    CONSOLIDATED_TRANSCRIPT_PDF_PATH = os.path.join(BASE_FOLDER, f"{safe_keyword_string}_Meeting_Transcripts_{current_date_str}.pdf")

    try:
        os.makedirs(EMAIL_SAVE_PATH, exist_ok=True)
        os.makedirs(ATTACHMENT_SAVE_PATH, exist_ok=True)
        os.makedirs(TRANSCRIPT_SAVE_PATH, exist_ok=True)
    except OSError as e:
        print(f"Error creating base directories: {e}. Exiting.")
        sys.exit(1)

def is_valid_pdf(path):
    """Checks if a PDF file is valid and has pages."""
    try:
        reader = PdfReader(path)
        if not reader.pages:
            log(f"⚠️ PDF '{os.path.basename(path)}' has no pages.")
            return False
        return True
    except Exception as e:
        log(f"❌ Invalid PDF '{os.path.basename(path)}': {e}")
        return False

def merge_pdfs(pdf_files, output_path):
    """Merges a list of PDF files into a single output PDF."""
    global USE_TQDM_FOR_LOGGING
    writer = PdfWriter()
    valid_pdfs_to_merge = []

    if not pdf_files:
        log(f"ℹ️ No PDF files provided to merge for output: {os.path.basename(output_path)}")
        return

    for pdf_path in pdf_files:
        if not os.path.exists(pdf_path):
            log(f"⚠️ Skipping missing PDF for merge: {pdf_path}")
            continue
        if not is_valid_pdf(pdf_path): # is_valid_pdf already logs
            log(f"⚠️ Skipping invalid PDF for merge: {pdf_path}") # Additional context
            continue
        valid_pdfs_to_merge.append(pdf_path)

    if not valid_pdfs_to_merge:
        log(f"ℹ️ No valid PDF files found to merge for output: {os.path.basename(output_path)}")
        return

    log(f"Merging {len(valid_pdfs_to_merge)} PDF(s) into {os.path.basename(output_path)}...")
    original_tqdm_state = USE_TQDM_FOR_LOGGING
    USE_TQDM_FOR_LOGGING = True # Force tqdm for this specific operation's progress
    try:
        for pdf in tqdm(valid_pdfs_to_merge, desc=f"Merging for {os.path.basename(output_path)}", unit="file"):
            try:
                reader = PdfReader(pdf)
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                log(f"❌ Error reading '{os.path.basename(pdf)}' during merge: {e}")
    finally:
        USE_TQDM_FOR_LOGGING = original_tqdm_state # Restore original tqdm state

    if not writer.pages:
        log(f"⚠️ No pages were added to the writer for {os.path.basename(output_path)}. Output PDF will not be created.")
        return

    try:
        with open(output_path, "wb") as out:
            writer.write(out)
        log(f"✅ Merged PDF saved to: {output_path} ({len(writer.pages)} pages)")
    except Exception as e:
        log(f"❌ Failed to save merged PDF {os.path.basename(output_path)}: {e}")

def save_email_as_pdf(email_data, pdf_path):
    """Saves email content (headers and body) as a PDF file."""
    try:
        c = canvas.Canvas(pdf_path, pagesize=letter)
        c.setFont("Helvetica", 10)
        y_position, line_height = 750, 12
        margin = 50

        def draw_wrapped(text_content, current_y):
            """Draws text, wrapping it and creating new pages if necessary."""
            lines = []
            # Split by newlines first, then wrap each resulting line
            for paragraph in text_content.split('\n'):
                 lines.extend(re.findall('.{1,100}(?:\s+|$)', paragraph)) # Wrap lines at 100 chars

            for line in lines:
                if current_y < margin + line_height : # Check if new page is needed
                    c.showPage()
                    c.setFont("Helvetica", 10)
                    current_y = 750
                c.drawString(margin, current_y, line.strip())
                current_y -= line_height
            return current_y

        headers = [
            f"From: {email_data.get('from', '[N/A]')}",
            f"To: {email_data.get('to', '[N/A]')}",
            f"CC: {email_data.get('cc', '[N/A]')}",
            f"Subject: {email_data.get('subject', '[N/A]')}",
            f"Sent: {email_data.get('sent', '[N/A]')}",
            "", # Blank line before body
            "Body:"
        ]
        for header_text in headers:
            y_position = draw_wrapped(header_text, y_position)
            if y_position < margin + line_height : # Redundant check, draw_wrapped handles it
                 c.showPage(); c.setFont("Helvetica", 10); y_position = 750

        y_position = draw_wrapped(email_data.get('body', '[No Body Content]'), y_position)
        c.save()
    except Exception as e:
        log(f"❌ Failed to save email as PDF '{os.path.basename(pdf_path)}': {e}")

def convert_office_to_pdf(office_file_path, output_pdf_path):
    """Converts Word or Excel files to PDF using COM."""
    office_file_path = os.path.abspath(office_file_path) # Ensure absolute paths
    output_pdf_path = os.path.abspath(output_pdf_path)
    file_ext = os.path.splitext(office_file_path)[1].lower()
    app = None
    doc = None
    # Excel constants
    xlLandscape = 2 # For PageSetup.Orientation
    # wdFormatPDF = 17 # For Word.SaveAs
    # xlTypePDF = 0 # For Workbook.ExportAsFixedFormat Type

    try:
        if file_ext in WORD_EXTENSIONS:
            log(f"  🔄 Converting Word file to PDF: {os.path.basename(office_file_path)}")
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = False # Run in background
            app.DisplayAlerts = 0 # Suppress alerts
            doc = app.Documents.Open(office_file_path, ReadOnly=True)
            doc.SaveAs(output_pdf_path, FileFormat=17) # 17 is wdFormatPDF
            log(f"  📄 Word file converted to PDF: {os.path.basename(output_pdf_path)}")
            return True
        elif file_ext in EXCEL_EXTENSIONS:
            log(f"  🔄 Converting Excel file to PDF: {os.path.basename(office_file_path)}")
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Workbooks.Open(office_file_path, ReadOnly=True)
            try:
                log(f"    Configuring PageSetup for sheets in '{os.path.basename(office_file_path)}'.")
                for sheet_index in range(1, doc.Sheets.Count + 1):
                    sheet = doc.Sheets(sheet_index)
                    sheet_name = getattr(sheet, 'Name', f'Index {sheet_index}')
                    try:
                        ps = sheet.PageSetup
                        ps.Orientation = xlLandscape
                        ps.Zoom = False # Important to allow FitToPagesWide/Tall
                        ps.FitToPagesWide = 1 # Fit to 1 page wide
                        try:
                            # To allow automatic height, FitToPagesTall should be False
                            ps.FitToPagesTall = False
                        except pywintypes.com_error as e_fittotall:
                            log(f"      ⚠️ COM Error setting FitToPagesTall for sheet '{sheet_name}': {e_fittotall}. Continuing without it.")
                        log(f"      📄 Configured basic PageSetup for sheet '{sheet_name}': Landscape, Fit to 1 page wide.")
                    except pywintypes.com_error as e_sheet_ps:
                        err_code = e_sheet_ps.hresult if hasattr(e_sheet_ps, 'hresult') else 'N/A'
                        err_msg = str(e_sheet_ps)
                        log(f"      ⚠️ COM Error configuring PageSetup for sheet '{sheet_name}' (Code: {err_code}): {err_msg}.")
                    except Exception as e_sheet_ps_general: # Catch any other exception during PageSetup
                        log(f"      ⚠️ General error configuring PageSetup for sheet '{sheet_name}': {e_sheet_ps_general}.")

                log(f"    Finished basic PageSetup configuration for all sheets.")
            except pywintypes.com_error as e_sheets_iter: # Error iterating sheets
                log(f"    ⚠️ COM Error iterating through sheets to configure PageSetup: {e_sheets_iter}.")
            except Exception as e_sheets_iter_general: # Other errors during sheet iteration
                 log(f"    ⚠️ General error iterating through sheets to configure PageSetup: {e_sheets_iter_general}.")

            doc.ExportAsFixedFormat(0, output_pdf_path, IgnorePrintAreas=True) # 0 is xlTypePDF
            log(f"  📄 Excel file converted to PDF: {os.path.basename(output_pdf_path)}")
            return True
        else:
            log(f"  ℹ️ File type {file_ext} is not supported for Office to PDF conversion: {os.path.basename(office_file_path)}")
            return False
    except pywintypes.com_error as e_com: # Catch COM errors specifically
        log(f"  ❌ COM Error during Office to PDF conversion for '{os.path.basename(office_file_path)}': {e_com}")
        if hasattr(e_com, 'hresult') and e_com.hresult == -2147221005: # RPC_E_SERVERFAULT
            log(f"    ⚠️  The required Office application (Word/Excel) might not be installed, registered correctly, or might have crashed/become unresponsive.")
        return False
    except Exception as e: # Catch any other exceptions
        log(f"  ❌ General error during Office to PDF conversion for '{os.path.basename(office_file_path)}': {e}")
        return False
    finally: # Ensure Office applications are closed
        if doc:
            try: doc.Close(SaveChanges=False)
            except: pass # Ignore errors on close
        if app:
            try: app.Quit()
            except: pass # Ignore errors on quit
        doc = None # Dereference
        app = None

def get_all_mail_items(folder, path=""):
    """Recursively retrieves all mail items from a given Outlook folder and its subfolders."""
    items = []
    current_path = os.path.join(path, folder.Name) # Build current folder path for logging
    folder_name_lower = folder.Name.lower()

    # Check if folder should be excluded by type or name
    try:
        if hasattr(folder, 'DefaultFolderType') and folder.DefaultFolderType in EXCLUDED_FOLDER_TYPES:
            log(f"  🚫 Skipping excluded folder (by type {folder.DefaultFolderType}): {current_path}")
            return items
    except pywintypes.com_error: # Handle cases where DefaultFolderType might not be accessible
        log(f"  ⚠️ COM error accessing DefaultFolderType for '{current_path}'. Checking by name.")

    if folder_name_lower in EXCLUDED_FOLDER_NAMES_LOWER:
        log(f"  🚫 Skipping excluded folder (by name): {current_path}")
        return items

    log(f"  📂 Scanning folder: {current_path}")

    # Process items in the current folder
    try:
        folder_items_collection = folder.Items
        if folder_items_collection is not None:
            # Attempt to restrict items to only emails for efficiency
            restriction = "[MessageClass] = 'IPM.Note'"
            filtered_items = None
            try:
                filtered_items = folder_items_collection.Restrict(restriction)
            except pywintypes.com_error as e_restrict_call:
                log(f"    ⚠️ COM error calling .Restrict on folder '{current_path}': {e_restrict_call}. Will attempt to iterate all items.")
            
            if filtered_items is not None:
                for item in filtered_items:
                    try:
                        # Ensure it's actually an email item (IPM.Note or derived)
                        if hasattr(item, 'MessageClass') and item.MessageClass.startswith('IPM.Note'):
                            items.append(item)
                    except pywintypes.com_error:
                        log(f"    ⚠️ COM error accessing MessageClass for an item in '{current_path}' (restricted list). Skipping item.")
            else: # Fallback if restriction failed or returned None
                log(f"  ℹ️ Restriction returned None or failed for folder '{current_path}'. Iterating all items in this folder.")
                for item in folder_items_collection: # Iterate all items if restriction fails
                    try:
                        if hasattr(item, 'MessageClass') and item.MessageClass.startswith('IPM.Note'):
                            items.append(item)
                    except pywintypes.com_error:
                        log(f"    ⚠️ COM error accessing MessageClass for an item in '{current_path}' (full list). Skipping item.")

    except pywintypes.com_error as e_items:
        log(f"  ❌ COM error accessing items collection in folder '{current_path}': {e_items}")
    except Exception as e_general_items: # Catch other potential errors
        log(f"  ❌ General error accessing items in folder '{current_path}': {e_general_items}")

    # Recursively process subfolders
    try:
        sub_folders_collection = folder.Folders
        if sub_folders_collection is not None:
            for subfolder in sub_folders_collection:
                try:
                    items.extend(get_all_mail_items(subfolder, current_path))
                except pywintypes.com_error as e_sub_item_access:
                    log(f"  ⚠️ COM Error accessing/processing subfolder of '{current_path}' (Name: {getattr(subfolder, 'Name', 'Unknown')}): {e_sub_item_access}. Skipping this subfolder.")
                except Exception as e_sub_general:
                    log(f"  ⚠️ General error processing subfolder of '{current_path}' (Name: {getattr(subfolder, 'Name', 'Unknown')}): {e_sub_general}. Skipping this subfolder.")
    except pywintypes.com_error as e_subfolders:
        log(f"  ❌ COM error accessing subfolders collection of '{current_path}': {e_subfolders}")
    except Exception as e_general_subfolders:
        log(f"  ❌ General error accessing subfolders of '{current_path}': {e_general_subfolders}")
    return items

def process_emails():
    """Main function to connect to Outlook, retrieve, filter, and process emails."""
    global USE_TQDM_FOR_LOGGING
    log("Starting email processing...")
    outlook = None
    all_mail_items_from_stores = []

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if not namespace.Folders or namespace.Folders.Count == 0:
            log("❌ No mail stores (accounts/PSTs) found in Outlook.")
            return [], []

        log(f"Found {namespace.Folders.Count} mail store(s)/top-level folder(s) in Outlook profile.")
        for store_root_folder in namespace.Folders:
            try:
                store_name = "[Unknown Store Name]"
                try:
                    store_name = getattr(store_root_folder, 'Name', f"Store (ID: {getattr(store_root_folder, 'StoreID', 'N/A')})")
                except pywintypes.com_error: # Handle if store name itself is problematic
                    log(f"  ⚠️ COM error getting name for a store. Using generic ID.")

                log(f"Scanning Store/Top-Level Folder: '{store_name}'")
                all_mail_items_from_stores.extend(get_all_mail_items(store_root_folder, path=store_name))
            except pywintypes.com_error as e_store:
                log(f"  ❌ COM error processing a store/top-level folder ('{store_name}'): {e_store}. Skipping this store.")
                continue # Skip to the next store
            except Exception as e_store_general:
                log(f"  ❌ General error processing a store/top-level folder ('{store_name}'): {e_store_general}. Skipping this store.")

    except pywintypes.com_error as e_com:
        log(f"❌ COM Error connecting to Outlook or accessing namespace: {e_com}")
        log("   Ensure Outlook is running, accessible, and not showing any dialogs.")
        return [], []
    except Exception as e:
        log(f"❌ Failed to connect to Outlook or access namespace: {e}")
        return [], []

    log(f"Retrieved {len(all_mail_items_from_stores)} total mail-like items from all scanned Outlook folders.")

    email_pdf_files, attachment_pdf_files = [], []
    email_count, saved_attachment_count = 0, 0
    skipped_image_attachment_count = 0
    converted_office_files_count = 0
    skipped_large_attachment_count = 0

    valid_messages_for_sorting = []
    log("Filtering and preparing messages for sorting...")
    original_tqdm_state_filter = USE_TQDM_FOR_LOGGING # Save current tqdm state
    # USE_TQDM_FOR_LOGGING = False # Disable tqdm for this potentially noisy part if needed
    try:
        for item_idx, item in enumerate(all_mail_items_from_stores):
            try:
                if hasattr(item, 'ReceivedTime'): # Check if item has ReceivedTime (emails do)
                    received_time = getattr(item, 'ReceivedTime') # Get the time
                    if received_time is not None: # Ensure it's not None
                        valid_messages_for_sorting.append(item)
            except pywintypes.com_error as e_com_attr:
                subject_for_log = "[Unretrievable Subject]"
                try: subject_for_log = getattr(item, 'Subject', subject_for_log)
                except: pass
                log(f"  ⚠️ COM Error accessing attributes for item {item_idx} (Subject: '{subject_for_log}'): {e_com_attr}. Skipping.")
            except Exception as e_gen_attr: # Catch other errors
                subject_for_log = "[Unretrievable Subject]"
                try: subject_for_log = getattr(item, 'Subject', subject_for_log)
                except: pass
                log(f"  ⚠️ General error accessing attributes for item {item_idx} (Subject: '{subject_for_log}'): {e_gen_attr}. Skipping.")
    finally:
        USE_TQDM_FOR_LOGGING = original_tqdm_state_filter # Restore tqdm state

    log(f"Found {len(valid_messages_for_sorting)} messages with valid ReceivedTime for sorting.")

    # Sort messages by ReceivedTime, newest first
    try:
        messages = sorted(valid_messages_for_sorting, key=lambda x: x.ReceivedTime, reverse=True)
    except Exception as e: # Fallback if sorting fails (e.g., inconsistent time types)
        log(f"⚠️ Unable to sort messages by ReceivedTime: {e}. Using unsorted valid messages.")
        messages = valid_messages_for_sorting

    log(f"Processing {len(messages)} email messages after sorting attempt...")
    original_tqdm_state_process = USE_TQDM_FOR_LOGGING
    USE_TQDM_FOR_LOGGING = True # Enable tqdm for the main email processing loop
    try:
        for message_item in tqdm(messages, desc="Processing Emails", unit="email"):
            try:
                # Initialize attributes with defaults
                subject_attr = ""
                body_attr = ""
                sender_name_attr = "[Unknown Sender]"
                sender_email_attr = "unknown@example.com"
                received_time_attr = "[No Date]"
                to_attr = "[Unknown Recipient]"
                cc_attr = "[No CC]"

                # Safely get attributes
                try: subject_attr = getattr(message_item, "Subject", "")
                except: pass
                try: body_attr = getattr(message_item, "Body", "")
                except: pass

                subject_lower = (subject_attr or "").lower()
                body_lower = (body_attr or "").lower()

                # Keyword matching
                if not any(k in subject_lower or k in body_lower for k in keywords):
                    continue # Skip if no keywords match

                email_count += 1

                # Get other email details
                try: sender_name_attr = getattr(message_item, 'SenderName', sender_name_attr)
                except: pass
                try: sender_email_attr = getattr(message_item, 'SenderEmailAddress', sender_email_attr)
                except: pass
                try:
                    rt = getattr(message_item, "ReceivedTime", None)
                    if rt: # Ensure rt is not None
                        if isinstance(rt, (datetime, pywintypes.TimeType)): # Check type
                            if isinstance(rt, pywintypes.TimeType): # Convert pywintypes.TimeType to datetime
                                rt = datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second, tzinfo=rt.tzinfo)
                            received_time_attr = rt.strftime('%Y-%m-%d %H:%M:%S %Z') if rt.tzinfo else rt.strftime('%Y-%m-%d %H:%M:%S')
                        else:
                            received_time_attr = str(rt) # Fallback if not a recognized time type
                except: pass
                try: to_attr = getattr(message_item, "To", to_attr)
                except: pass
                try: cc_attr = getattr(message_item, "CC", cc_attr)
                except: pass

                email_details = {
                    "from": f"{sender_name_attr} <{sender_email_attr}>", "sent": received_time_attr,
                    "to": to_attr, "cc": cc_attr, "subject": subject_attr or "[No Subject]",
                    "body": body_attr or "[No Body]",
                }

                # Create a safe filename for the email PDF
                safe_subject_filename = re.sub(r'[<>:"/\\|?*]', '_', subject_attr or "No_Subject")
                safe_subject_filename = re.sub(r'\s+', '_', safe_subject_filename).strip('_')
                if not safe_subject_filename: safe_subject_filename = "No_Subject"
                safe_subject_filename = (safe_subject_filename[:75] + '...') if len(safe_subject_filename) > 78 else safe_subject_filename # Truncate long names

                # Use email's received time for timestamp if available, else current time
                timestamp_obj = datetime.now() # Default
                try:
                    raw_received_time = getattr(message_item, "ReceivedTime", datetime.now())
                    if isinstance(raw_received_time, (datetime, pywintypes.TimeType)):
                         timestamp_obj = raw_received_time
                         if isinstance(timestamp_obj, pywintypes.TimeType): # Convert if necessary
                            timestamp_obj = datetime(timestamp_obj.year, timestamp_obj.month, timestamp_obj.day,
                                                     timestamp_obj.hour, timestamp_obj.minute, timestamp_obj.second,
                                                     tzinfo=timestamp_obj.tzinfo)
                except: pass # Use default if error

                timestamp_str = timestamp_obj.strftime("%Y%m%d_%H%M%S") if isinstance(timestamp_obj, datetime) else datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]

                email_pdf_filename = f"{timestamp_str}_{safe_subject_filename}.pdf"
                email_pdf_path = os.path.join(EMAIL_SAVE_PATH, email_pdf_filename)
                save_email_as_pdf(email_details, email_pdf_path)
                if os.path.exists(email_pdf_path): # Check if PDF was actually created
                     email_pdf_files.append(email_pdf_path)

                # Process attachments
                if hasattr(message_item, 'Attachments') and message_item.Attachments.Count > 0:
                    for att_idx in range(1, message_item.Attachments.Count + 1): # Attachments are 1-indexed
                        att = message_item.Attachments.Item(att_idx)
                        original_att_filename_loop = "[Unknown Attachment]"
                        saved_att_path_loop = None # Path where attachment is saved

                        try:
                            original_att_filename_loop = getattr(att, 'FileName', f'unknown_attachment_{saved_attachment_count}')

                            # Check attachment size
                            try:
                                attachment_size = att.Size
                                if attachment_size > MAX_ATTACHMENT_SIZE_BYTES:
                                    log(f"  📎 Skipping attachment '{original_att_filename_loop}' ({(attachment_size / (1024*1024)):.2f} MB) due to size exceeding {MAX_ATTACHMENT_SIZE_MB} MB limit.")
                                    skipped_large_attachment_count += 1
                                    continue # Skip this attachment
                            except: # If size cannot be determined, attempt to save anyway
                                log(f"  ⚠️ Error getting size for attachment '{original_att_filename_loop}'. Will attempt to save.")

                            _, att_ext = os.path.splitext(original_att_filename_loop)
                            att_ext_lower = att_ext.lower()

                            # Skip common signature images
                            if att_ext_lower in SIGNATURE_IMAGE_EXTENSIONS:
                                log(f"  📎 Skipping signature image attachment: {original_att_filename_loop} from email (Subject: '{subject_attr}')")
                                skipped_image_attachment_count +=1
                                continue

                            # Create safe filename for attachment
                            safe_att_filename_base = re.sub(r'[<>:"/\\|?*]', '_', os.path.splitext(original_att_filename_loop)[0])
                            safe_att_filename_base = re.sub(r'\s+', '_', safe_att_filename_base).strip('_')
                            if not safe_att_filename_base: safe_att_filename_base = f"unknown_attachment_{saved_attachment_count + 1}"
                            safe_att_filename_base = (safe_att_filename_base[:75] + '...') if len(safe_att_filename_base) > 78 else safe_att_filename_base

                            saved_att_filename_with_ext = f"{timestamp_str}_{safe_att_filename_base}{att_ext}"
                            temp_saved_att_path = os.path.join(ATTACHMENT_SAVE_PATH, saved_att_filename_with_ext)

                            # Handle potential filename conflicts
                            counter = 1
                            saved_att_path_loop = temp_saved_att_path
                            while os.path.exists(saved_att_path_loop):
                                saved_att_path_loop = os.path.join(ATTACHMENT_SAVE_PATH, f"{timestamp_str}_{safe_att_filename_base}_{counter}{att_ext}")
                                counter += 1

                            att.SaveAsFile(saved_att_path_loop)
                            log(f"  📎 Attachment saved: {os.path.basename(saved_att_path_loop)}")
                            saved_attachment_count += 1

                            # Convert Office attachments to PDF
                            if att_ext_lower in WORD_EXTENSIONS or att_ext_lower in EXCEL_EXTENSIONS:
                                converted_pdf_filename = f"{timestamp_str}_{safe_att_filename_base}.pdf"
                                temp_converted_pdf_path = os.path.join(ATTACHMENT_SAVE_PATH, converted_pdf_filename)
                                pdf_counter = 1
                                final_converted_pdf_path = temp_converted_pdf_path
                                while os.path.exists(final_converted_pdf_path): # Handle conflict for converted PDF name
                                     final_converted_pdf_path = os.path.join(ATTACHMENT_SAVE_PATH, f"{timestamp_str}_{safe_att_filename_base}_{pdf_counter}.pdf")
                                     pdf_counter += 1

                                if convert_office_to_pdf(saved_att_path_loop, final_converted_pdf_path):
                                    if os.path.exists(final_converted_pdf_path) and is_valid_pdf(final_converted_pdf_path):
                                        attachment_pdf_files.append(final_converted_pdf_path)
                                        converted_office_files_count += 1
                                    else:
                                        log(f"  ⚠️ Office to PDF conversion reported success, but output PDF is invalid or missing: {os.path.basename(final_converted_pdf_path)}")
                                else:
                                    log(f"  ⚠️ Failed to convert Office attachment to PDF: {os.path.basename(saved_att_path_loop)}")
                            elif att_ext_lower == '.pdf': # If attachment is already PDF
                                if is_valid_pdf(saved_att_path_loop):
                                    attachment_pdf_files.append(saved_att_path_loop)
                                else:
                                    log(f"  ⚠️ Saved PDF attachment is invalid: {os.path.basename(saved_att_path_loop)}")
                            else:
                                log(f"  ℹ️ Attachment '{os.path.basename(saved_att_path_loop)}' is not a PDF or supported Office file. Saved, but not added for PDF consolidation.")
                        except Exception as e_att:
                            log(f"  ❌ Error processing attachment '{original_att_filename_loop}' from email (Subject: '{subject_attr}') - {e_att}")
                        finally:
                            att = None # Release COM object
            except Exception as e: # Catch errors processing a single email message
                subject_for_log = "[Unreadable Subject]"
                try: subject_for_log = message_item.Subject
                except: pass
                log(f"❌ General error processing email (Subject: '{subject_for_log}'): {e}\n{traceback.format_exc(limit=2)}")
            finally:
                message_item = None # Release COM object

    finally: # Restore tqdm state and release Outlook objects
        USE_TQDM_FOR_LOGGING = original_tqdm_state_process
        if 'namespace' in locals() and namespace is not None: namespace = None
        if 'outlook' in locals() and outlook is not None: outlook = None

    log(f"✅ Matched emails found and processed: {email_count}")
    log(f"✅ Total attachments saved (excluding common images & oversized): {saved_attachment_count}")
    log(f"ℹ️ Skipped image attachments (likely signatures): {skipped_image_attachment_count}")
    log(f"ℹ️ Skipped large attachments (>{MAX_ATTACHMENT_SIZE_MB}MB): {skipped_large_attachment_count}")
    log(f"✅ Office files converted to PDF: {converted_office_files_count}")
    log(f"✅ PDF attachments (original or converted) collected for merging: {len(attachment_pdf_files)}")
    return email_pdf_files, attachment_pdf_files

def download_google_docs_from_drive(search_keywords, local_transcript_folder,
                                    creds_filename=GDRIVE_CLIENT_SECRET_FILE,
                                    token_filename=GDRIVE_TOKEN_FILE,
                                    drive_folder_id=GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID):
    """Downloads Google Docs matching keywords from a specified Drive folder as PDFs."""
    global USE_TQDM_FOR_LOGGING
    log(f"Attempting to download Google Docs from Drive folder ID: {drive_folder_id}")
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google.auth.exceptions import RefreshError
    except ImportError:
        log("❌ Google API client libraries not found. Please install: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
        return []

    if not os.path.exists(local_transcript_folder):
        try:
            os.makedirs(local_transcript_folder)
            log(f"Created transcript download folder: {local_transcript_folder}")
        except Exception as e_mkdir:
            log(f"❌ Failed to create transcript download folder {local_transcript_folder}: {e_mkdir}")
            return []

    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
    creds = None
    downloaded_paths = []

    # Load credentials from token file if it exists
    if os.path.exists(token_filename):
        try:
            creds = Credentials.from_authorized_user_file(token_filename, SCOPES)
            log(f"🔐 Using cached token: {token_filename}")
        except Exception as e:
            log(f"⚠️ Failed to load token from {token_filename}: {e}. Will attempt re-authentication.")
            creds = None

    # If no valid credentials, authenticate or refresh
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            log("🔑 Token expired, attempting refresh...")
            try:
                creds.refresh(Request())
            except RefreshError as e:
                log(f"⚠️ Token refresh failed: {e}. Removing stale token and attempting full re-auth...")
                if os.path.exists(token_filename):
                    try: os.remove(token_filename)
                    except Exception as e_rm_token: log(f"  Could not remove token file {token_filename}: {e_rm_token}")
                creds = None # Force re-auth
            except Exception as e: # Other refresh errors
                log(f"⚠️ An unexpected error occurred during token refresh: {e}. Attempting full re-auth...")
                if os.path.exists(token_filename):
                    try: os.remove(token_filename)
                    except Exception as e_rm_token: log(f"  Could not remove token file {token_filename}: {e_rm_token}")
                creds = None
        
        if not creds: # If still no creds (initial run or refresh failed)
            log("🚀 No valid credentials, starting authentication flow...")
            if not os.path.exists(creds_filename):
                log(f"❌ Credentials file '{creds_filename}' not found. Please download it from Google Cloud Console and place it in the script's directory.")
                return []
            try:
                flow = InstalledAppFlow.from_client_secrets_file(creds_filename, SCOPES)
                creds = flow.run_local_server(port=0) # Opens browser for auth
                log("✅ Authentication successful.")
            except Exception as e:
                log(f"❌ Authentication flow failed: {e}")
                return []
        
        # Save the credentials for the next run
        try:
            with open(token_filename, 'w') as token_file_handle:
                token_file_handle.write(creds.to_json())
            log(f"🔑 Credentials saved to {token_filename}")
        except Exception as e:
            log(f"⚠️ Could not save token to {token_filename}: {e}")

    if not creds: # Final check if credentials could be obtained
        log("❌ Failed to obtain Google Drive credentials. Cannot download transcripts.")
        return []

    try:
        service = build('drive', 'v3', credentials=creds)
        query = f"'{drive_folder_id}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        items = results.get('files', [])
        log(f"Found {len(items)} Google Docs in Drive folder '{drive_folder_id}'.")
        download_count = 0
        original_tqdm_state = USE_TQDM_FOR_LOGGING
        USE_TQDM_FOR_LOGGING = True # Enable tqdm for downloads
        try:
            for item in tqdm(items, desc="Downloading Transcripts", unit="doc"):
                file_name = item['name']
                file_id = item['id']

                if not any(k.lower() in file_name.lower() for k in search_keywords):
                    continue # Skip if keywords not in GDoc name

                # Sanitize filename for local saving
                safe_name = "".join(c if c.isalnum() or c in " ._-" else "_" for c in file_name)
                safe_name = (safe_name[:75] + '...') if len(safe_name) > 78 else safe_name # Truncate
                output_pdf_path = os.path.join(local_transcript_folder, f"{safe_name}.pdf")

                # Check if already downloaded and valid
                if os.path.exists(output_pdf_path) and os.path.getsize(output_pdf_path) > 0 : # Basic check
                    log(f"  ➡️ Transcript already downloaded and seems valid: {os.path.basename(output_pdf_path)}")
                    downloaded_paths.append(output_pdf_path)
                else:
                    if os.path.exists(output_pdf_path): # Exists but is empty/invalid
                        log(f"  ⚠️ Existing transcript '{os.path.basename(output_pdf_path)}' is empty or invalid. Re-downloading...")
                    log(f"  ⬇️ Downloading GDoc '{file_name}' as PDF...")
                    try:
                        request = service.files().export_media(fileId=file_id, mimeType='application/pdf')
                        fh = BytesIO()
                        downloader = MediaIoBaseDownload(fh, request)
                        done = False
                        while not done:
                            status, done = downloader.next_chunk() # Download in chunks
                        with open(output_pdf_path, 'wb') as f_out:
                            f_out.write(fh.getvalue())
                        log(f"  ✅ GDoc '{file_name}' downloaded to: {os.path.basename(output_pdf_path)}")
                        downloaded_paths.append(output_pdf_path)
                        download_count += 1
                    except Exception as e_download:
                        log(f"  ❌ Failed to download GDoc '{file_name}': {e_download}")
        finally:
            USE_TQDM_FOR_LOGGING = original_tqdm_state # Restore tqdm state

        log(f"Total matching Google Docs downloaded in this run: {download_count}")
    except pywintypes.com_error as e_com_drive: # Should not happen with Google API, but good to have
        log(f"❌ A COM error occurred during Google Drive operations (unexpected): {e_com_drive}")
    except Exception as e:
        log(f"❌ An error occurred during Google Drive operations: {e}")
    return downloaded_paths

def process_transcripts():
    """Downloads and verifies transcripts, then merges them."""
    log("Starting transcript processing...")
    downloaded_transcript_paths = download_google_docs_from_drive(keywords, TRANSCRIPT_SAVE_PATH)
    valid_transcript_pdfs_for_merging = []

    if os.path.exists(TRANSCRIPT_SAVE_PATH):
        log(f"Verifying downloaded and existing transcripts in: {TRANSCRIPT_SAVE_PATH}")
        all_potential_paths = set(downloaded_transcript_paths) # Start with newly downloaded

        # Add existing files in the directory that match keywords
        for file_name in os.listdir(TRANSCRIPT_SAVE_PATH):
            if file_name.lower().endswith('.pdf') and any(k.lower() in file_name.lower() for k in keywords):
                all_potential_paths.add(os.path.join(TRANSCRIPT_SAVE_PATH, file_name))
        
        for full_path in all_potential_paths:
            if os.path.exists(full_path): # Ensure file still exists
                if is_valid_pdf(full_path): # Check if it's a valid PDF
                    if full_path not in valid_transcript_pdfs_for_merging: # Avoid duplicates
                        valid_transcript_pdfs_for_merging.append(full_path)
                    log(f"  ✔️ Transcript confirmed valid for merging: {os.path.basename(full_path)}")
                else:
                    log(f"  ⚠️ Transcript matched but invalid, skipping: {os.path.basename(full_path)}")
            else:
                log(f"  ℹ️ Path {os.path.basename(full_path)} from download list no longer exists. Skipping.")
    else:
        log(f"⚠️ Transcript save path {TRANSCRIPT_SAVE_PATH} not found. No transcripts to process.")

    if valid_transcript_pdfs_for_merging:
        log(f"Found {len(valid_transcript_pdfs_for_merging)} unique and valid transcript PDFs to merge.")
        merge_pdfs(valid_transcript_pdfs_for_merging, CONSOLIDATED_TRANSCRIPT_PDF_PATH)
        return CONSOLIDATED_TRANSCRIPT_PDF_PATH # Return path of merged file
    else:
        log("ℹ️ No matching transcript PDFs found or downloaded to merge.")
        # Check if an old, empty consolidated file exists and remove it
        if os.path.exists(CONSOLIDATED_TRANSCRIPT_PDF_PATH):
             try:
                 if os.path.getsize(CONSOLIDATED_TRANSCRIPT_PDF_PATH) == 0:
                     log(f"Removing empty or potentially corrupt consolidated transcript PDF: {CONSOLIDATED_TRANSCRIPT_PDF_PATH}")
                     os.remove(CONSOLIDATED_TRANSCRIPT_PDF_PATH)
             except Exception as e_remove_empty:
                 log(f"Could not remove empty transcript PDF {CONSOLIDATED_TRANSCRIPT_PDF_PATH}: {e_remove_empty}")
        return None

def check_ocr_status_random_pages_worker(file_path, text_threshold=OCR_MIN_TEXT_LENGTH_THRESHOLD):
    """Worker function to check if a PDF likely already has OCR text by sampling random pages."""
    worker_log_prefix = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [WORKER_OCR_CHECK] "
    try:
        # Use print directly for worker logs for immediate feedback in multiprocessing
        print(f"{worker_log_prefix}Opening with fitz: {os.path.basename(file_path)} for random page check.", flush=True)
        doc = fitz.open(file_path)
        page_count = doc.page_count

        if page_count == 0:
            print(f"{worker_log_prefix}  ℹ️ PDF '{os.path.basename(file_path)}' has no pages. OCR cannot be checked/applied.", flush=True)
            doc.close()
            return False, "empty_pdf_no_pages"

        # Determine number of pages to sample
        num_pages_to_check_float = page_count * OCR_CHECK_PERCENTAGE
        pages_to_sample_count = max(1, int(num_pages_to_check_float)) # At least 1 page
        pages_to_sample_count = min(pages_to_sample_count, OCR_CHECK_MAX_PAGES) # Cap at max
        pages_to_sample_count = min(pages_to_sample_count, page_count) # Cannot sample more than available

        if pages_to_sample_count == 0 : # Should not happen with max(1,...) but defensive
             print(f"{worker_log_prefix}  ℹ️ PDF '{os.path.basename(file_path)}' - No pages selected for random check (count: {pages_to_sample_count}). Assuming OCR needed.", flush=True)
             doc.close()
             return False, "no_pages_selected_for_random_check"

        all_page_indices = list(range(page_count))
        k_sample = min(pages_to_sample_count, len(all_page_indices)) # Ensure k is not > population
        sampled_indices = random.sample(all_page_indices, k_sample)

        print(f"{worker_log_prefix}  Checking {k_sample} random page(s) (target: {OCR_CHECK_PERCENTAGE*100:.1f}%, max: {OCR_CHECK_MAX_PAGES}) from {page_count} total for '{os.path.basename(file_path)}'. Indices: {sampled_indices}", flush=True)

        for i, page_idx in enumerate(sampled_indices):
            page = doc.load_page(page_idx)
            text = page.get_text("text").strip()
            if len(text) < text_threshold:
                print(f"{worker_log_prefix}  ⚠️ No sufficient text (length {len(text)} < {text_threshold}) found on random page {page_idx + 1} of '{os.path.basename(file_path)}'. Assuming OCR needed.", flush=True)
                doc.close()
                return False, f"no_sufficient_text_on_random_page_{page_idx + 1}"
        
        print(f"{worker_log_prefix}  ✅ Sufficient text found on all {k_sample} randomly checked page(s) of '{os.path.basename(file_path)}'. Assuming OCR already applied.", flush=True)
        doc.close()
        return True, "sufficient_text_found_on_random_pages"

    except Exception as e:
        print(f"{worker_log_prefix}  ⚠️ Error checking OCR status (random pages) for '{os.path.basename(file_path)}' with fitz: {e}", flush=True)
        return False, f"fitz_error_random_check: {e}"

def ocr_pdf_task(file_path):
    """Task function to apply OCR to a single PDF file using ocrmypdf."""
    worker_log_prefix = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [WORKER_OCR_TASK] "
    print(f"{worker_log_prefix}Starting OCR task for: {os.path.basename(file_path)}", flush=True)

    ocrmypdf_exe_path_for_task = None # Initialize
    try:
        ocrmypdf_exe_path_for_task = shutil.which('ocrmypdf') # Find ocrmypdf in PATH
        print(f"{worker_log_prefix}  Path to ocrmypdf found by script worker: {ocrmypdf_exe_path_for_task}", flush=True)
        if ocrmypdf_exe_path_for_task:
            version_cmd = [ocrmypdf_exe_path_for_task, '--version']
            result = subprocess.run(version_cmd, capture_output=True, text=True, check=True, timeout=15)
            print(f"{worker_log_prefix}  OCRMypdf version successfully called by script worker: {result.stdout.strip()}", flush=True)
        else:
            print(f"{worker_log_prefix}  OCRMypdf command not found in PATH by script worker. Please check installation and PATH.", flush=True)
            return False, file_path, "ocrmypdf_not_found_in_worker_path"
    except Exception as e_diag:
        print(f"{worker_log_prefix}  ❌ DIAGNOSTIC ERROR trying to get ocrmypdf path/version: {e_diag}", flush=True)
        ocrmypdf_exe_path_for_task = 'ocrmypdf' # Fallback to command name if shutil.which fails but it might be in PATH

    # Pre-OCR checks
    if not os.path.exists(file_path):
        print(f"{worker_log_prefix}  ⚠️ File not found, cannot OCR: {file_path}", flush=True)
        return False, file_path, "file_not_found_at_ocr_task_start"
    try:
        reader_check = PdfReader(file_path) # Check if PyPDF can open it
        if not reader_check.pages:
            print(f"{worker_log_prefix}  ⚠️ Skipping OCR for empty PDF (no pages based on PyPDF2): {os.path.basename(file_path)}", flush=True)
            return False, file_path, "empty_pdf_pypdf_check"
    except Exception as e_read:
        print(f"{worker_log_prefix}  ⚠️ File '{os.path.basename(file_path)}' is not a valid PDF or is corrupt (PyPDF2 Read check failed: {e_read}). Skipping OCR.", flush=True)
        return False, file_path, f"invalid_pdf_pypdf_check: {e_read}"

    # Check if OCR is actually needed using random page sampling
    text_already_present, reason_ocr_check = check_ocr_status_random_pages_worker(file_path)
    print(f"{worker_log_prefix}  check_ocr_status_random_pages_worker for {os.path.basename(file_path)} returned: {text_already_present}, Reason: {reason_ocr_check}", flush=True)
    if text_already_present:
        print(f"{worker_log_prefix}  ✅ OCR not needed for {os.path.basename(file_path)} based on random page check. Reason: {reason_ocr_check}", flush=True)
        return True, file_path, f"already_ocred_text_found_random_check ({reason_ocr_check})"
    
    print(f"{worker_log_prefix}Applying OCR to: {os.path.basename(file_path)} (Reason: {reason_ocr_check})", flush=True)
    base, ext = os.path.splitext(file_path)
    temp_file = f"{base}_ocr_temp_{random.randint(1000,9999)}{ext}" # Unique temp file name

    # Ensure old temp file is removed if it exists
    if os.path.exists(temp_file):
        try: os.remove(temp_file)
        except Exception as e_rm_old_temp: print(f"{worker_log_prefix}Could not remove pre-existing temp file {temp_file}: {e_rm_old_temp}", flush=True)
    
    executable_to_run = ocrmypdf_exe_path_for_task if ocrmypdf_exe_path_for_task else 'ocrmypdf'

    try:
        # OCRmyPDF command arguments
        command = [
            executable_to_run, '--jobs', '1', # Use 1 job for ocrmypdf internal parallelism
            '--deskew', '--rotate-pages',
            '--force-ocr', # Force OCR even if some text is found (random check might miss some)
            # '--skip-encrypted', # Removed as per log, can be re-added if needed
            '--optimize', '1', # Basic optimization
            '--output-type', 'pdf', # Ensure output is PDF
            file_path, temp_file
        ]
        print(f"{worker_log_prefix}  Running command: {' '.join(command)}", flush=True)
        process_result = subprocess.run(command, check=True, capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=OCR_TIMEOUT_SECONDS)
        
        # Log stdout/stderr from ocrmypdf
        stdout_log = process_result.stdout.strip()
        stderr_log = process_result.stderr.strip()
        if stdout_log: print(f"{worker_log_prefix}  OCRMypdf STDOUT for {os.path.basename(file_path)}:\n{stdout_log}", flush=True)
        if stderr_log: print(f"{worker_log_prefix}  OCRMypdf STDERR for {os.path.basename(file_path)}:\n{stderr_log}", flush=True)

        if os.path.exists(temp_file) and os.path.getsize(temp_file) > 0:
            try:
                # Validate the OCR'd temp file before replacing
                try:
                    ocred_reader = PdfReader(temp_file)
                    if not ocred_reader.pages: raise ValueError("OCR'd PDF has no pages.")
                except Exception as e_val_ocr:
                    print(f"{worker_log_prefix}  ❌ OCR'd temp file '{os.path.basename(temp_file)}' appears invalid: {e_val_ocr}. Not replacing original.", flush=True)
                    if os.path.exists(temp_file): os.remove(temp_file)
                    return False, file_path, "ocr_output_invalid"
                
                os.remove(file_path) # Remove original
                shutil.move(temp_file, file_path) # Replace with OCR'd version
                print(f"{worker_log_prefix}  ✅ OCR complete, original replaced: {os.path.basename(file_path)}", flush=True)
                return True, file_path, "ocr_applied_successfully"
            except Exception as e_replace:
                print(f"{worker_log_prefix}  ❌ Error replacing original with OCR'd file '{os.path.basename(file_path)}': {e_replace}", flush=True)
                # Try to save the OCR'd file with a different name if replacement fails
                final_ocr_name = f"{base}_ocr_completed{ext}"
                try:
                    shutil.move(temp_file, final_ocr_name)
                    print(f"{worker_log_prefix}  ⚠️ OCR'd file saved as {os.path.basename(final_ocr_name)} due to replacement error. Original file ({os.path.basename(file_path)}) remains unchanged.", flush=True)
                except Exception as e_move_alt:
                    print(f"{worker_log_prefix}  ❌ Failed to even rename temp OCR file {os.path.basename(temp_file)} to {final_ocr_name}: {e_move_alt}", flush=True)
                return False, file_path, "ocr_replace_error_original_kept"
        else:
            print(f"{worker_log_prefix}  ⚠️ OCR process completed for {os.path.basename(file_path)} but temp file '{os.path.basename(temp_file)}' is missing or empty.", flush=True)
            return False, file_path, "ocr_temp_file_issue_after_processing"

    except subprocess.TimeoutExpired:
        print(f"{worker_log_prefix}  ❌ OCR timed out for {os.path.basename(file_path)} after {OCR_TIMEOUT_SECONDS} seconds. Skipping file.", flush=True)
        if os.path.exists(temp_file): # Clean up temp file on timeout
            try: os.remove(temp_file)
            except Exception as e_remove_timeout: print(f"{worker_log_prefix}    ⚠️ Could not remove temp file {temp_file} after timeout: {e_remove_timeout}", flush=True)
        return False, file_path, "ocr_timeout"
    except subprocess.CalledProcessError as e_proc:
        print(f"{worker_log_prefix}  ❌ OCR failed for {os.path.basename(file_path)} with CalledProcessError. Return code: {e_proc.returncode}", flush=True)
        # Decode stdout/stderr safely
        stdout_decoded = e_proc.stdout if isinstance(e_proc.stdout, str) else (e_proc.stdout.decode('utf-8', errors='replace') if isinstance(e_proc.stdout, bytes) else "")
        stderr_decoded = e_proc.stderr if isinstance(e_proc.stderr, str) else (e_proc.stderr.decode('utf-8', errors='replace') if isinstance(e_proc.stderr, bytes) else "")
        if stdout_decoded and stdout_decoded.strip(): print(f"{worker_log_prefix}  OCRMypdf STDOUT:\n{stdout_decoded.strip()}", flush=True)
        if stderr_decoded and stderr_decoded.strip(): print(f"{worker_log_prefix}  OCRMypdf STDERR:\n{stderr_decoded.strip()}", flush=True)
        if os.path.exists(temp_file): # Clean up temp file on error
            try: os.remove(temp_file)
            except Exception as e_remove_cpe: print(f"{worker_log_prefix}    ⚠️ Could not remove temp file {temp_file} after CalledProcessError: {e_remove_cpe}", flush=True)
        return False, file_path, f"ocr_calledprocesserror_{e_proc.returncode}"
    except FileNotFoundError: # ocrmypdf command not found
        print(f"{worker_log_prefix}  ❌ OCR failed for {os.path.basename(file_path)}: ocrmypdf command not found.", flush=True)
        print(f"{worker_log_prefix}     Please ensure ocrmypdf is installed and in your system's PATH.", flush=True)
        return False, file_path, "ocrmypdf_not_found"
    except Exception as e_ocr_general: # Catch-all for other OCR errors
        print(f"{worker_log_prefix}  ❌ An unexpected error occurred during OCR for {os.path.basename(file_path)}: {e_ocr_general}", flush=True)
        print(traceback.format_exc(), flush=True) # Print full traceback for unexpected errors
        if os.path.exists(temp_file): # Clean up temp file
            try: os.remove(temp_file)
            except Exception as e_remove_general: print(f"{worker_log_prefix}    ⚠️ Could not remove temp file {temp_file} after other Exception: {e_remove_general}", flush=True)
        return False, file_path, f"ocr_unexpected_error: {e_ocr_general}"
    
    # Fallback return, should ideally not be reached if logic is exhaustive
    return False, file_path, "ocr_unknown_exit_path"

def split_pdf_by_size(input_pdf_path, max_size_mb=MAX_SPLIT_SIZE_MB):
    """Splits a PDF into multiple parts if it exceeds max_size_mb."""
    global USE_TQDM_FOR_LOGGING
    log(f"Checking if PDF needs splitting: {os.path.basename(input_pdf_path)}")
    if not os.path.exists(input_pdf_path):
        log(f"  ⚠️ Input PDF for splitting not found: {input_pdf_path}")
        return [] # Return empty list if source PDF doesn't exist
    
    try:
        input_file_size = os.path.getsize(input_pdf_path)
    except OSError as e:
        log(f"  ⚠️ Could not get size of input PDF {input_pdf_path}: {e}")
        return [] # Return empty if size cannot be determined

    max_bytes = max_size_mb * 1024 * 1024

    if input_file_size <= max_bytes:
        log(f"  ℹ️ PDF '{os.path.basename(input_pdf_path)}' ({input_file_size / (1024*1024):.2f} MB) is within size limit ({max_size_mb} MB). No splitting needed.")
        return [input_pdf_path] # Return original path in a list

    log(f"  Splitting PDF '{os.path.basename(input_pdf_path)}' ({input_file_size / (1024*1024):.2f} MB) as it exceeds {max_size_mb} MB...")
    
    try:
        reader = PdfReader(input_pdf_path)
        if not reader.pages:
            log(f"  ⚠️ PDF for splitting has no pages: {os.path.basename(input_pdf_path)}")
            return []
    except Exception as e:
        log(f"  ❌ Error reading PDF for splitting '{os.path.basename(input_pdf_path)}': {e}")
        return []

    part_num = 1
    current_writer = PdfWriter()
    split_paths = []
    current_part_page_count = 0

    def write_current_part_func(writer_to_write, current_part_num_func):
        """Helper to write the current PDF part to a file."""
        if not writer_to_write.pages:
            log(f"  Skipping writing empty part {current_part_num_func} for {os.path.basename(input_pdf_path)}")
            return None
        base_name, ext_name = os.path.splitext(input_pdf_path)
        part_path = f"{base_name}_part{current_part_num_func}{ext_name}"
        try:
            with open(part_path, "wb") as f_out:
                writer_to_write.write(f_out)
            actual_size_mb = os.path.getsize(part_path) / (1024 * 1024)
            log(f"  ✅ Created part: {os.path.basename(part_path)} ({actual_size_mb:.2f} MB, {len(writer_to_write.pages)} pages)")
            # Warning if a part is still too large (e.g., single large page)
            if actual_size_mb > max_size_mb + (max_size_mb * 0.05): # Allow 5% overshoot for warning
                log(f"  ❗ WARNING: Part {os.path.basename(part_path)} ({actual_size_mb:.2f} MB) still exceeds {max_size_mb}MB limit significantly. This may happen if a single page is very large.")
            return part_path
        except Exception as e_write_func:
            log(f"  ❌ Error writing PDF part {os.path.basename(part_path)}: {e_write_func}")
            return None

    original_tqdm_state = USE_TQDM_FOR_LOGGING
    USE_TQDM_FOR_LOGGING = True # Enable tqdm for splitting progress
    try:
        total_pages = len(reader.pages)
        for i, page in enumerate(tqdm(reader.pages, desc=f"Splitting {os.path.basename(input_pdf_path)}", unit="page")):
            current_writer.add_page(page)
            current_part_page_count += 1
            
            # Check size periodically or if it's the last page
            perform_check = (current_part_page_count % SPLIT_CHECK_INTERVAL == 0 or (i == total_pages - 1)) and len(current_writer.pages) > 0

            if perform_check:
                temp_stream_for_check = BytesIO()
                try:
                    current_writer.write(temp_stream_for_check)
                    current_writer_size = temp_stream_for_check.tell()

                    if current_writer_size >= max_bytes:
                        if current_part_page_count == 1: # Single page is already too large
                            path = write_current_part_func(current_writer, part_num)
                            if path: split_paths.append(path)
                            current_writer = PdfWriter() # Reset for next part
                            current_part_page_count = 0
                            part_num += 1
                        else: # More than one page in current writer, split before this last page
                            writer_to_save_before_last = PdfWriter()
                            # Add all pages except the last one that caused overflow
                            for page_idx_to_save in range(len(current_writer.pages) - 1):
                                writer_to_save_before_last.add_page(current_writer.pages[page_idx_to_save])
                            
                            if writer_to_save_before_last.pages: # If there's anything to save
                                path = write_current_part_func(writer_to_save_before_last, part_num)
                                if path: split_paths.append(path)
                                part_num += 1
                            
                            # Start new writer with the page that caused overflow
                            last_page_causing_overflow = current_writer.pages[-1]
                            current_writer = PdfWriter()
                            current_writer.add_page(last_page_causing_overflow)
                            current_part_page_count = 1
                except Exception as e_stream: # Error during in-memory size check
                    log(f"  ⚠️ Error writing current part to memory for size check (page {i+1}): {e_stream}. Continuing to accumulate pages.")
    finally:
        USE_TQDM_FOR_LOGGING = original_tqdm_state # Restore tqdm state

    # Write any remaining pages in the current_writer
    if current_writer.pages:
        path = write_current_part_func(current_writer, part_num)
        if path: split_paths.append(path)

    if not split_paths and os.path.exists(input_pdf_path) and reader.pages:
        log(f"  ⚠️ Splitting {os.path.basename(input_pdf_path)} resulted in no output files, but it had pages. Original file might be used if within limits or if splitting failed.")
        if os.path.getsize(input_pdf_path) <= max_bytes: # If original is now somehow small enough
             return [input_pdf_path]
        return [] # Return empty if splitting truly failed to produce parts

    return split_paths

if __name__ == "__main__":
    initialize_paths_and_logging() # Get keywords, set up folders
    try:
        log(f"--- {SCRIPT_NAME} Script {__version__} Started ---")
        log(f"Searching for keywords: {', '.join(keywords)}")
        log(f"Base output folder: {BASE_FOLDER}")

        email_pdfs, attachment_pdfs_for_consolidation = process_emails()
        consolidated_transcript_file_path = process_transcripts() # This returns the path of the merged transcript PDF or None

        # Merge emails
        if email_pdfs:
            merge_pdfs(email_pdfs, CONSOLIDATED_EMAIL_PDF_PATH)
        else:
            log("ℹ️ No email PDFs to merge.")

        # Merge attachments
        if attachment_pdfs_for_consolidation:
            merge_pdfs(attachment_pdfs_for_consolidation, CONSOLIDATED_ATTACHMENT_PDF_PATH)
        else:
            log("ℹ️ No PDF attachments (original or converted from Office) to merge.")

        # --- OCR Processing ---
        files_to_ocr_final = []

        # Option to skip OCR for emails (usually text-based already)
        if os.path.exists(CONSOLIDATED_EMAIL_PDF_PATH) and os.path.getsize(CONSOLIDATED_EMAIL_PDF_PATH) > 0:
            log(f"ℹ️ Skipping OCR for consolidated email PDF as per user request: {CONSOLIDATED_EMAIL_PDF_PATH}")
        else:
            log(f"ℹ️ Consolidated email PDF not found or empty, skipping for OCR: {CONSOLIDATED_EMAIL_PDF_PATH}")

        # Prepare attachments for OCR (split if necessary)
        if os.path.exists(CONSOLIDATED_ATTACHMENT_PDF_PATH) and os.path.getsize(CONSOLIDATED_ATTACHMENT_PDF_PATH) > 0:
            attachment_paths_for_ocr = split_pdf_by_size(CONSOLIDATED_ATTACHMENT_PDF_PATH, max_size_mb=MAX_SPLIT_SIZE_MB)
            if attachment_paths_for_ocr:
                log(f"  ➡️ Consolidated Attachments PDF(s) for OCR: {', '.join(map(os.path.basename, attachment_paths_for_ocr))}")
                files_to_ocr_final.extend(attachment_paths_for_ocr)
        else:
            log(f"ℹ️ Consolidated PDF of attachments not found or empty, skipping for OCR: {CONSOLIDATED_ATTACHMENT_PDF_PATH}")

        # Prepare transcripts for OCR (split if necessary)
        actual_transcript_path_to_check = None
        if consolidated_transcript_file_path and os.path.exists(consolidated_transcript_file_path) and os.path.getsize(consolidated_transcript_file_path) > 0:
            actual_transcript_path_to_check = consolidated_transcript_file_path
        elif os.path.exists(CONSOLIDATED_TRANSCRIPT_PDF_PATH) and os.path.getsize(CONSOLIDATED_TRANSCRIPT_PDF_PATH) > 0: # Fallback to default path if return was None but file exists
            log(f"ℹ️ Using existing consolidated transcript PDF for OCR: {CONSOLIDATED_TRANSCRIPT_PDF_PATH} (returned path was not valid or not provided).")
            actual_transcript_path_to_check = CONSOLIDATED_TRANSCRIPT_PDF_PATH

        if actual_transcript_path_to_check:
            transcript_paths_for_ocr = split_pdf_by_size(actual_transcript_path_to_check, max_size_mb=MAX_SPLIT_SIZE_MB)
            if transcript_paths_for_ocr:
                log(f"  ➡️ Transcripts for OCR: {', '.join(map(os.path.basename, transcript_paths_for_ocr))}")
                files_to_ocr_final.extend(transcript_paths_for_ocr)
        else:
            log(f"ℹ️ Consolidated transcript PDF not found, not created, or empty. Skipping for OCR.")
        
        # Ensure unique list of files for OCR
        unique_files_to_ocr = []
        seen_paths = set()
        for f_path in files_to_ocr_final:
            abs_path = os.path.abspath(f_path) # Normalize path for uniqueness check
            if abs_path not in seen_paths:
                unique_files_to_ocr.append(f_path)
                seen_paths.add(abs_path)

        if unique_files_to_ocr:
            log(f"--- Starting OCR for {len(unique_files_to_ocr)} unique PDF file(s) using parallel processing ---")
            cpu_cores = os.cpu_count()
            if cpu_cores:
                num_workers = max(1, int(cpu_cores * 0.80)) # Use 80% of cores, at least 1
            else:
                num_workers = 2 # Fallback if cpu_count fails
            log(f"Using {num_workers} worker processes for parallel OCR (targeting 80% of CPU cores: {cpu_cores}).")
            
            ocr_results_summary = []
            original_tqdm_state_ocr = USE_TQDM_FOR_LOGGING
            USE_TQDM_FOR_LOGGING = False # Disable main thread tqdm for OCR part, workers print directly

            with concurrent.futures.ProcessPoolExecutor(max_workers=num_workers) as executor:
                future_to_file = {executor.submit(ocr_pdf_task, file_path): file_path for file_path in unique_files_to_ocr}
                # Use tqdm for iterating over futures completion
                for future in tqdm(concurrent.futures.as_completed(future_to_file), total=len(unique_files_to_ocr), desc="Applying OCR (Parallel)", unit="file"):
                    file_path_processed = future_to_file[future]
                    try:
                        is_ocred_valid, _, reason_msg = future.result() # Unpack result
                        ocr_results_summary.append({'file': os.path.basename(file_path_processed), 'processed_ok': is_ocred_valid, 'message': reason_msg})
                        if is_ocred_valid:
                            if "already_ocred" in reason_msg.lower():
                                log(f"✅ OCR check passed (already OCR'd): {os.path.basename(file_path_processed)}. Reason: {reason_msg}")
                            else:
                                log(f"✅ OCR successfully completed for: {os.path.basename(file_path_processed)}. Reason: {reason_msg}")
                        else:
                            log(f"ℹ️ OCR processing issue for {os.path.basename(file_path_processed)}. Reason: {reason_msg}")
                    except Exception as exc:
                        log(f"❌ OCR task for {os.path.basename(file_path_processed)} generated an exception in main thread: {exc}")
                        ocr_results_summary.append({'file': os.path.basename(file_path_processed), 'processed_ok': False, 'message': f"exception_in_future_result: {exc}"})
            
            USE_TQDM_FOR_LOGGING = original_tqdm_state_ocr # Restore tqdm state
            successful_ocr_processing_count = sum(1 for r in ocr_results_summary if r['processed_ok'])
            log(f"--- Parallel OCR processing completed. {successful_ocr_processing_count}/{len(unique_files_to_ocr)} files are now considered OCR'd and valid. ---")
            if len(unique_files_to_ocr) > 0 and successful_ocr_processing_count < len(unique_files_to_ocr):
                log("Summary of OCR attempts with issues or skips:")
                for res in ocr_results_summary:
                    if not res['processed_ok']:
                        log(f"  - File: {res['file']}, Processed OK: {res['processed_ok']}, Message: {res['message']}")
        else:
            log("ℹ️ No unique files to OCR.")

        log(f"--- {SCRIPT_NAME} Script {__version__} Finished Successfully ---")

    except Exception as e: # Catch-all for any unhandled exceptions in main block
        log(f"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        log(f"!!!!!!!!!!!!!!!!! CRITICAL SCRIPT ERROR !!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        log(f"An unhandled error occurred: {e}")
        log(f"Traceback:\n{traceback.format_exc()}")
        log(f"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print(f"\n❌ CRITICAL ERROR: Script terminated due to an unhandled exception. Check log file for details.")
    finally:
        if LOG_FILE: # Ensure LOG_FILE is defined
            try:
                with open(LOG_FILE, "w", encoding="utf-8") as f:
                    f.write(f"--- {SCRIPT_NAME} Log --- Version: {__version__} ---\n")
                    if keywords: f.write(f"Keywords: {', '.join(keywords)}\n")
                    if BASE_FOLDER: f.write(f"Base Folder: {BASE_FOLDER}\n")
                    f.write(f"Execution Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("--- Messages ---\n")
                    f.write('\n'.join(log_messages))
                print(f"\n[FINAL LOG] Log file saved to: {LOG_FILE}")
            except Exception as e_log:
                print(f"\n[FINAL LOG ERROR] Failed to write final log file to {LOG_FILE}: {e_log}")
                print("\n--- Collected Log Messages (Fallback due to log write error) ---")
                for msg in log_messages: print(msg)
        else:
            print("\n[FINAL LOG ERROR] LOG_FILE path was not initialized. Cannot save log file.")
            print("\n--- Collected Log Messages (Fallback as LOG_FILE not set) ---")
            for msg in log_messages: print(msg) # Print to console if log file fails

        print(f"\n✅ Script {SCRIPT_NAME} {__version__} execution attempt concluded. See messages above and log file for details.")

