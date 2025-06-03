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
 
-GDRIVE_CLIENT_SECRET_FILE = "client_secret.json" # Path to your Google client_secret.json
-GDRIVE_TOKEN_FILE = 'token.json' # Path where the token will be stored
-GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID = "1y84cNQAnSsr7UYvK84TXH4MmW-GZmuC8" # Specific Google Drive Folder ID
+# Default Google Drive configuration. These are overwritten by user input at
+# runtime so that the script works for anyone without editing the code.
+GDRIVE_CLIENT_SECRET_FILE = "client_secret.json"  # Default path to client_secret.json
+GDRIVE_TOKEN_FILE = "token.json"  # Default path where the OAuth token will be stored
+GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID = ""  # Google Drive folder ID containing transcripts
 
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
diff --git a/Email Search Script v1.0.153.py	 b/Email Search Script v1.0.153.py	
index 6ba95b4203163e68c9323cabaa744f570251796b..4e9667d8172a4a70153bf80a7eef75df779cb039 100644
--- a/Email Search Script v1.0.153.py	
+++ b/Email Search Script v1.0.153.py	
@@ -67,64 +69,78 @@ EXCLUDED_FOLDER_NAMES_LOWER = [ # Case-insensitive list of folder names to exclu
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
+    global GDRIVE_CLIENT_SECRET_FILE, GDRIVE_TOKEN_FILE, GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID
 
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
 
+    # Prompt for Google Drive configuration so the script can work for any user
+    user_input = input(f"Path to Google client_secret.json [{GDRIVE_CLIENT_SECRET_FILE}]: ").strip()
+    if user_input:
+        GDRIVE_CLIENT_SECRET_FILE = user_input
+
+    user_input = input(f"Path for Google Drive token file [{GDRIVE_TOKEN_FILE}]: ").strip()
+    if user_input:
+        GDRIVE_TOKEN_FILE = user_input
+
+    user_input = input("Google Drive folder ID for meeting transcripts (leave blank to skip download): ").strip()
+    if user_input:
+        GDRIVE_MEETING_TRANSCRIPTS_FOLDER_ID = user_input
+
 
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
diff --git a/Email Search Script v1.0.153.py	 b/Email Search Script v1.0.153.py	
index 6ba95b4203163e68c9323cabaa744f570251796b..4e9667d8172a4a70153bf80a7eef75df779cb039 100644
--- a/Email Search Script v1.0.153.py	
+++ b/Email Search Script v1.0.153.py	
@@ -628,50 +644,54 @@ def process_emails():
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
+    if not drive_folder_id:
+        log("ℹ️ No Google Drive folder ID provided. Skipping transcript download.")
+        return []
+
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
