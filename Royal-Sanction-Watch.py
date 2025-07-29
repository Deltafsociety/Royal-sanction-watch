import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox, font
import threading
import queue
import os
import re
import time
from urllib.parse import urljoin
from io import StringIO, BytesIO
import zipfile
import xml.etree.ElementTree as ET
import sys
import csv # For My Vessels CSV persistence
import datetime # Import datetime for timestamping files

# --- Configuration for report freshness ---
MAX_REPORT_AGE_SECONDS = 3600 * 24 # 24 hours (adjust as needed)

# --- Global Queues for thread-safe GUI updates ---
gui_update_queue = queue.Queue() # For SanctionsAppFrame
excel_comparator_log_queue = queue.Queue() # For ExcelComparatorFrame
my_vessels_log_queue = queue.Queue() # For MyVesselsAppFrame


# --- Configuration ---
OFAC_SDN_CSV_URL = "https://www.treasury.gov/ofac/downloads/sdn.csv"
UK_SANCTIONS_PUBLICATION_PAGE_URL = "https://www.gov.uk/government/publications/the-uk-sanctions-list"
DMA_XLSX_URL = "https://www.dma.dk/Media/638834044135010725/2025118019-7%20Importversion%20-%20List%20of%20EU%20designated%20vessels%20(20-05-2025)%203010691_2_0.XLSX"
UANI_SHEET_ID = "19SBq7N1Ety5fCfaTOZUf61QY-hJIouptx9Gv-uosR_k"
UANI_CSV_URL = f"https://docs.google.com/spreadsheets/d/{UANI_SHEET_ID}/export?format=csv&gid=0"


# --- Global User-Agent Initializer & Retry setup ---
try:
    import requests # Ensure requests is imported
    import pandas as pd # Import pandas as pd
    from bs4 import BeautifulSoup
    import lxml # Explicitly try to import lxml parser
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
    from fake_useragent import UserAgent # Import for random user agents
    from requests.adapters import HTTPAdapter # Import for retries
    from urllib3.util.retry import Retry # Import for retries

except ImportError as e:
    missing_lib_name = str(e).split()[-1]
    if "requests" in str(e): missing_lib_name = "requests"
    elif "pandas" in str(e): missing_lib_name = "pandas" # Still check if it's missing for user instruction
    elif "BeautifulSoup" in str(e): missing_lib_name = "beautifulsoup4"
    elif "lxml" in str(e): missing_lib_name = "lxml"
    elif "openpyxl" in str(e): missing_lib_name = "openpyxl"
    elif "fake_useragent" in str(e): missing_lib_name = "fake-useragent"
    elif "Retry" in str(e): missing_lib_name = "urllib3"
    
    print(f"Initial import error: {e}")
    messagebox.showerror("Dependency Error",
                          f"A critical library is missing or cannot be imported: '{missing_lib_name}'.\n"
                          f"Please install it using:\n\npip install {missing_lib_name}\n\n"
                          f"For 'lxml' specifically, ensure you run:\npip install lxml")
    sys.exit(1)


# --- Global User-Agent Initializer & Requests Session Setup ---
try:
    ua = UserAgent()

    # Configure retry strategy for network robustness
    retry_strategy = Retry(
        total=5, # Increased retries to 5
        backoff_factor=2, # Exponential backoff (1s, 2s, 4s, 8s, 16s)
        status_forcelist=[429, 500, 502, 503, 504], # Status codes to retry on
        allowed_methods=["HEAD", "GET", "OPTIONS"] # Methods to retry
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    http_session = requests.Session()
    http_session.mount("https://", adapter)
    http_session.mount("http://", adapter)

except Exception as e:
    # This block handles any errors during the setup, but the ImportError is already handled at top.
    # We set ua and http_session to None so the rest of the code can fall back gracefully.
    print(f"Warning: An unexpected error occurred during networking setup: {e}. Traceback: {sys.exc_info()}")
    ua = None
    http_session = None


# --- Helper: Log to GUI and Console ---
class TextRedirector(object):
    """A class to redirect stdout to a Tkinter Text widget via a queue."""
    def __init__(self, queue):
        self.queue = queue
        self.stdout = sys.stdout # Keep a reference to original stdout

    def write(self, str_val):
        self.queue.put({'type': 'log', 'message': str_val})
        self.stdout.write(str_val) # Also print to original stdout/console

    def flush(self):
        self.stdout.flush()

def log_to_gui_and_console(message):
    """Sends a message to the console and to the GUI queue for thread-safe UI updates for SanctionsApp."""
    gui_update_queue.put({'type': 'log', 'message': str(message) + "\n"}) # Add newline for log consistency

# --- Helper: Get Soup from URL ---
def get_soup(url, is_xml=False, delay_seconds=1):
    log_to_gui_and_console(f"Attempting to fetch: {url} (delay: {delay_seconds}s)")
    time.sleep(delay_seconds) # Respect delays
    try:
        current_session = http_session if http_session else requests # Use global session or fallback to plain requests
        
        user_agent_string = ua.random if ua else 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        headers = {'User-Agent': user_agent_string}
        
        response = current_session.get(url, timeout=60, headers=headers) # Increased timeout to 60s
        response.raise_for_status()
        log_to_gui_and_console(f"Successfully fetched: {url} with status {response.status_code}")
        
        parser_type = 'lxml-xml' if is_xml else 'lxml'
        soup = BeautifulSoup(response.content, parser_type)
        return soup
    except requests.exceptions.RequestException as e:
        log_to_gui_and_console(f"Network error fetching URL {url}: {e}. Traceback: {sys.exc_info()}")
        return None
    except Exception as e:
        log_to_gui_and_console(f"A general error occurred in get_soup for {url}: {e}. Traceback: {sys.exc_info()}")
        return None

# Helper function for IMO checksum validation (Optional but highly recommended)
def is_valid_imo_checksum(imo_number):
    """
    Validates a 7-digit IMO number using its checksum.
    Returns True if valid, False otherwise.
    """
    if not re.fullmatch(r'^\d{7}$', imo_number):
        return False # Not a 7-digit number

    imo_str = str(imo_number)
    checksum = int(imo_str[-1])
    
    calculated_sum = 0
    for i in range(6):
        calculated_sum += int(imo_str[i]) * (7 - i)
        
    return (calculated_sum % 10) == checksum


# --- Data Fetching Functions ---
def fetch_ofac_vessels(pd_module): # Added pd_module parameter
    log_to_gui_and_console("\n--- Fetching OFAC SDN list ---")
    vessels_data = [] # This will store dicts like {'Vessel Name': ..., 'IMO Number': ...}
    try:
        log_to_gui_and_console("OFAC: Initiating request.")
        current_session = http_session if http_session else requests # Use global session or fallback
        response = current_session.get(OFAC_SDN_CSV_URL, timeout=60, headers={'User-Agent': ua.random if ua else 'Mozilla/5.0'}) # Increased timeout
        response.raise_for_status()
        csv_content = response.content.decode('utf-8')
        log_to_gui_and_console("OFAC: CSV content downloaded. Parsing...")
        
        reader = csv.reader(StringIO(csv_content))
        
        for row_idx, row in enumerate(reader):
            if row_idx < 5: # Log first few rows to confirm data structure
                log_to_gui_and_console(f"OFAC Row {row_idx}: {row}")

            if len(row) > 2 and row[2].strip().upper() == 'VESSEL':
                name = row[1].strip() # Assuming vessel name is at index 1
                imo_number = "N/A"
                if len(row) > 11 and row[11]:
                    remarks = row[11].strip().upper()
                    imo_match = re.search(r'IMO\s*(\d{7})', remarks)
                    if imo_match:
                        imo_number = imo_match.group(1) # Extract the IMO number
                        # Clean IMO number immediately at extraction point for consistency
                        imo_number = re.sub(r'\D', '', imo_number).strip() 
                
                # Only add to list if an IMO number was successfully found for a VESSEL entry
                if imo_number != "N/A" and re.fullmatch(r'^\d{7}$', imo_number): # Final check for 7 digits
                    vessels_data.append({"Vessel Name": name, "IMO Number": imo_number, "Source": "OFAC"})
        
        log_to_gui_and_console(f"OFAC: Found {len(vessels_data)} vessels based on 'VESSEL' type and extracted IMO.")

    except requests.exceptions.RequestException as e:
        log_to_gui_and_console(f"Network error fetching OFAC data: {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module
    except Exception as e:
        log_to_gui_and_console(f"Error processing OFAC data: {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module
    
    # Convert to DataFrame (this DataFrame construction is still needed for consistency with other fetchers)
    df_vessels = pd_module.DataFrame(vessels_data) # Use pd_module
    
    if not df_vessels.empty:
        # Drop duplicates based on IMO number. Keeping the first occurrence is typical.
        df_vessels.drop_duplicates(subset=['IMO Number'], keep='first', inplace=True)
    
    log_to_gui_and_console(f"OFAC: Final {len(df_vessels)} vessels after filtering duplicates and validating IMO.")
    return df_vessels


def fetch_uk_sanctions_vessels(progress_callback, pd_module): # Added pd_module parameter
    log_to_gui_and_console("\n--- Fetching UK Sanctions list (ODT format) ---")
    try:
        log_to_gui_and_console("UK: Fetching main page...")
        main_page_soup = get_soup(UK_SANCTIONS_PUBLICATION_PAGE_URL)
        progress_callback() # Step 1: Fetched main page
        if not main_page_soup:
            log_to_gui_and_console("UK: Failed to get main page soup.")
            for _ in range(2): progress_callback() # Account for remaining steps if failed
            return pd_module.DataFrame() # Use pd_module
        
        odt_link_tag = main_page_soup.find('a', href=re.compile(r'\.odt$', re.IGNORECASE))
        if not odt_link_tag:
            log_to_gui_and_console("UK: Error: Could not find ODT download link on UK Sanctions page.")
            for _ in range(2): progress_callback() # Account for remaining steps if failed
            return pd_module.DataFrame() # Use pd_module
        
        odt_url = urljoin(UK_SANCTIONS_PUBLICATION_PAGE_URL, odt_link_tag['href'])
        
        log_to_gui_and_console(f"UK: Found ODT link: {odt_url}. Downloading...")
        odt_response = http_session.get(odt_url, timeout=60, headers={'User-Agent': ua.random if ua else 'Mozilla/5.0'})
        odt_response.raise_for_status()
        log_to_gui_and_console("UK: ODT file downloaded successfully. Parsing...")
        progress_callback() # Step 2: Downloaded ODT
        
        vessels_data = []
        with zipfile.ZipFile(BytesIO(odt_response.content)) as odt_zip:
            content_xml_name = 'content.xml'
            if content_xml_name not in odt_zip.namelist():
                log_to_gui_and_console("UK: Warning: 'content.xml' not found directly in ODT. Attempting to list contents.")
                xml_files = [name for name in odt_zip.namelist() if name.endswith('.xml') and 'content' in name.lower()]
                if xml_files:
                    content_xml_name = xml_files[0]
                    log_to_gui_and_console(f"UK: Using '{content_xml_name}' as content file.")
                else:
                    log_to_gui_and_console("UK: Error: No suitable content XML file found in ODT.")
                    for _ in range(1): progress_callback() # Account for remaining steps if failed
                    return pd_module.DataFrame() # Use pd_module

            content_xml = odt_zip.read(content_xml_name)
            root = ET.fromstring(content_xml)
            full_text = "".join(root.itertext()).strip()
            log_to_gui_and_console("UK: Extracted text from ODT. Searching for IMOs...")

            imo_matches = list(re.finditer(r'IMO\s*[:;]?\s*(\d{7})', full_text, re.IGNORECASE))

            for match in imo_matches:
                imo = match.group(1)
                context_start = max(0, match.start() - 200)
                context = full_text[context_start : match.start()]
                
                name_pattern = r'(?:Name\s*[:;]\s*|Vessel Name\s*[:;]\s*)([^\n,;]+)'
                name_match = re.search(name_pattern, context, re.IGNORECASE)
                
                name = "Unknown UK Vessel"
                if name_match:
                    name = name_match.group(1).strip()
                    name = re.sub(r'^(?:Name|Vessel Name)\s*[:;]?\s*', '', name, flags=re.IGNORECASE).strip()
                    if len(name) > 100: name = name[:100] + "..."
                elif "vessel" in context.lower():
                    generic_name_match = re.search(r'([^\n,;]{5,100}?)\s+vessel', context, re.IGNORECASE)
                    if generic_name_match:
                        name = generic_name_match.group(1).strip()
                        if len(name) > 100: name = name[:100] + "..."

                vessels_data.append({"Vessel Name": name, "IMO Number": imo, "Source": "UK (ODT)"})
        
        progress_callback() # Step 3: Parsed ODT content
        log_to_gui_and_console(f"UK: Found {len(vessels_data)} potential UK vessels.")
        return pd_module.DataFrame(vessels_data) # Use pd_module
    except requests.exceptions.RequestException as e:
        log_to_gui_and_console(f"Network error during UK Sanctions ODT download: {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module
    except zipfile.BadZipFile:
        log_to_gui_and_console(f"UK: Error: Downloaded UK ODT file is not a valid zip file. It might be corrupted or not an ODT. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module
    except ET.ParseError as e:
        log_to_gui_and_console(f"UK: XML parsing error for UK ODT content: {e}. File might be malformed. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module
    except Exception as e:
        log_to_gui_and_console(f"UK: An unexpected error occurred during UK Sanctions processing: {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module

def fetch_dma_vessels(pd_module): # Added pd_module parameter
    log_to_gui_and_console("\n--- Fetching EU (DMA) Sanctions list ---")
    try:
        log_to_gui_and_console("DMA: Initiating request...")
        current_session = http_session if http_session else requests # Use global session or fallback
        response = current_session.get(DMA_XLSX_URL, timeout=60, headers={'User-Agent': ua.random if ua else 'Mozilla/5.0'})
        response.raise_for_status()
        log_to_gui_and_console("DMA: XLSX file downloaded. Parsing...")
        df = pd_module.read_excel(BytesIO(response.content)) # Use pd_module
        log_to_gui_and_console("DMA: XLSX content parsed.")
        
        imo_col = None
        name_col = None
        for col in df.columns:
            cleaned_col = str(col).lower().strip()
            if 'imo' in cleaned_col and not imo_col:
                imo_col = col
            if ('vessel' in cleaned_col or 'name' in cleaned_col) and not name_col:
                name_col = col
            if imo_col and name_col: break

        if not imo_col or not name_col:
            log_to_gui_and_console(f"DMA: Error: Could not find 'IMO' ({imo_col}) or 'Vessel Name' ({name_col}) columns in EU (DMA) Excel. Found columns: {df.columns.tolist()}. Traceback: {sys.exc_info()}")
            return pd_module.DataFrame() # Use pd_module
        
        df = df[[name_col, imo_col]].copy()
        df.rename(columns={name_col: "Vessel Name", imo_col: "IMO Number"}, inplace=True)
        df["Source"] = "EU (DMA)"
        df.dropna(subset=["IMO Number"], inplace=True)
        df["IMO Number"] = df["IMO Number"].astype(str).apply(lambda x: re.sub(r'\D', '', str(x)))
        
        df = df[df['IMO Number'].str.match(r'^\d{7}$')]
        
        log_to_gui_and_console(f"DMA: Found {len(df)} EU (DMA) vessels.")
        return df
    except requests.exceptions.RequestException as e:
        log_to_gui_and_console(f"Network error fetching EU (DMA) data: {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module
    except Exception as e:
        log_to_gui_and_console(f"DMA: Error processing DMA data: {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module

def fetch_uani_vessels(pd_module): # Added pd_module parameter
    log_to_gui_and_console(f"\n--- Fetching UANI data from Google Sheet ---")
    try:
        log_to_gui_and_console("UANI: Initiating request...")
        current_session = http_session if http_session else requests # Use global session or fallback
        response = current_session.get(UANI_CSV_URL, timeout=60, headers={'User-Agent': ua.random if ua else 'Mozilla/5.0'})
        response.raise_for_status()
        csv_text = response.content.decode('utf-8')
        log_to_gui_and_console("UANI: CSV content downloaded. Parsing...")
        
        lines = csv_text.splitlines()
        header_row_index = -1
        for i, line in enumerate(lines):
            if 'imo' in line.lower() and ('vessel' in line.lower() or 'name' in line.lower()):
                header_row_index = i
                break
        
        if header_row_index == -1:
            log_to_gui_and_console(f"UANI: Error: Could not find header row with 'IMO' and 'Vessel/Name' in UANI Google Sheet. Traceback: {sys.exc_info()}")
            return pd_module.DataFrame() # Use pd_module
        
        clean_csv_text = "\n".join(lines[header_row_index:])
        df = pd_module.read_csv(StringIO(clean_csv_text)) # Use pd_module
        log_to_gui_and_console("UANI: Google Sheet content parsed.")
        
        original_columns = list(df.columns)
        cleaned_columns_map = {str(col).strip().lower(): col for col in original_columns}
        
        imo_col_key = None
        name_col_key = None

        for pattern in ['imo', 'imo number', 'imo no']:
            if pattern in cleaned_columns_map:
                imo_col_key = cleaned_columns_map[pattern]
                break
        
        for pattern in ['vessel name', 'vessel', 'name']:
            if pattern in cleaned_columns_map:
                name_col_key = cleaned_columns_map[pattern]
                break

        if not imo_col_key or not name_col_key:
            log_to_gui_and_console(f"UANI: Error: Missing 'imo' ({imo_col_key}) or 'vessel name' ({name_col_key}) column in UANI Google Sheet after header parsing. Found cols: {df.columns.tolist()}. Traceback: {sys.exc_info()}")
            return pd_module.DataFrame() # Use pd_module

        df = df[[name_col_key, imo_col_key]].copy()
        df.rename(columns={name_col_key: "Vessel Name", imo_col_key: "IMO Number"}, inplace=True)
        df["Source"] = "UANI (Google Sheet)"
        df.dropna(subset=["IMO Number"], inplace=True)
        df["IMO Number"] = df["IMO Number"].astype(str).apply(lambda x: re.sub(r'\D', '', str(x)))
        
        df = df[df['IMO Number'].str.match(r'^\d{7}$')]
        
        log_to_gui_and_console(f"UANI: Found {len(df)} UANI vessels.")
        return df
    except requests.exceptions.RequestException as e:
        log_to_gui_and_console(f"Network error processing UANI Google Sheet data: {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module
    except Exception as e:
        log_to_gui_and_console(f"UANI: Error processing UANI Google Sheet data: {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame() # Use pd_module

# --- Worker function for Sanctions Report ---
def process_all_data(output_filepath, on_complete_callback=None):
    # Total steps for the main progress bar:
    # 1. OFAC Fetch
    # 2. UK Fetch (get page)
    # 3. UK Fetch (ODT download)
    # 4. UK Fetch (ODT parse)
    # 5. DMA Fetch
    # 6. UANI Fetch
    # 7. Excel Write
    total_steps = 7
    current_step = 0
    gui_update_queue.put({'type': 'progress_config', 'max': total_steps})
    
    # Helper to increment progress and update GUI
    def increment_progress():
        nonlocal current_step
        current_step += 1
        gui_update_queue.put({'type': 'progress_update', 'value': current_step, 'max': total_steps})
        log_to_gui_and_console(f"Overall Progress: {current_step}/{total_steps}") # Diagnostic log

    log_to_gui_and_console("Starting sanctions report generation...")
    
    try:
        # Import pandas inside this function to ensure it's available in the thread
        import pandas as pd_local 

        # Fetch OFAC data (pass pd_local explicitly)
        df_ofac = fetch_ofac_vessels(pd_local)
        increment_progress() 
        
        # Fetch UK sanctions data (pass pd_local explicitly)
        df_uk = fetch_uk_sanctions_vessels(progress_callback=increment_progress, pd_module=pd_local)
        
        # Fetch DMA data (pass pd_local explicitly)
        df_dma = fetch_dma_vessels(pd_local)
        increment_progress()
        
        # Fetch UANI data (pass pd_local explicitly)
        df_uani = fetch_uani_vessels(pd_local)
        increment_progress()
        
        log_to_gui_and_console("\nWriting data to Excel file...")
        timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
        base_name = os.path.basename(output_filepath).replace(".xlsx", "").replace(".xls", "")
        final_output_filepath = os.path.join(os.path.dirname(output_filepath), f"{base_name}_{timestamp}.xlsx")

        with pd_local.ExcelWriter(final_output_filepath, engine='openpyxl') as writer: # Use pd_local
            def write_sheet(df, sheet_name):
                if not df.empty:
                    df_copy = df.copy()
                    if 'IMO Number' in df_copy.columns:
                        df_copy['IMO Number'] = df_copy['IMO Number'].astype(str)
                    
                    if 'IMO Number' in df_copy.columns:
                        df_copy.drop_duplicates(subset=['IMO Number'], keep='first').to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        df_copy.drop_duplicates(keep='first').to_excel(writer, sheet_name=sheet_name, index=False)

                    log_to_gui_and_console(f"{sheet_name} data written.")
                else:
                    log_to_gui_and_console(f"No data for {sheet_name} to write.")
            
            write_sheet(df_ofac, "OFAC_Vessels")
            write_sheet(df_uk, "UK_Sanctions_Vessels")
            write_sheet(df_dma, "EU_DMA_Vessels")
            write_sheet(df_uani, "UANI_Vessels_Tracked")
        
        gui_update_queue.put({'type': 'show_message', 'kind': 'info', 'title': 'Success',
                               'message': f"Report generated successfully:\n{final_output_filepath}"})
        if on_complete_callback:
            on_complete_callback(final_output_filepath) # Pass the final path back

    except Exception as e:
        log_to_gui_and_console(f"Error during main report processing: {e}. Traceback: {sys.exc_info()}")
        gui_update_queue.put({'type': 'show_message', 'kind': 'error', 'title': 'Excel Write Error',
                               'message': f"Could not write the Excel file.\nError: {e}"})
        if on_complete_callback:
            on_complete_callback(None) # Signal failure
            
    increment_progress() # Final increment for Excel write completion
    gui_update_queue.put({'type': 'processing_done'})
    log_to_gui_and_console("\nScript finished.")

# --- Excel Comparison Logic ---

# Global DataFrames to hold loaded data for search functionality
# g_dfs1 and g_dfs2 will be populated from outside.
# Access to pd here is fine as it's outside threaded function.
g_dfs1 = {}
g_dfs2 = {}

# Define common column patterns for numerical identifiers like IMO numbers
IMO_LIKE_COLUMN_PATTERNS = ['imo', 'id', 'number', 'code', 'vesselimo', 'imo no']

def is_imo_like_column(column_name):
    """Checks if a column name indicates an IMO-like number column for targeted highlighting."""
    cleaned_col_name = str(column_name).strip().lower()
    for pattern in IMO_LIKE_COLUMN_PATTERNS:
        if pattern in cleaned_col_name:
            return True
    return False

def compare_and_highlight_excel_threaded(file1_path, file2_path, output_choice, output_folder, status_callback=None):
    """
    Worker function for the Excel comparison, run in a separate thread.
    Highlights common numerical identifiers in identified 'IMO-like' columns in red.
    """
    def update_status(message):
        if status_callback:
            status_callback(message)
        else:
            excel_comparator_log_queue.put(f"[Excel Comparator] {message}")

    update_status("Starting comparison...")

    try:
        # Import pandas inside this function to ensure it's available in the thread
        import pandas as pd_local 

        if not os.path.exists(file1_path):
            raise FileNotFoundError(f"File not found at {file1_path}")
        if not os.path.exists(file2_path):
            raise FileNotFoundError(f"File not found at {file2_path}")

        output_dir = os.path.join(output_folder, "highlighted_excel_files")
        os.makedirs(output_dir, exist_ok=True)

        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        update_status(f"Loading {file1_path}...")
        global g_dfs1
        g_dfs1 = pd_local.read_excel(file1_path, sheet_name=None) # Use pd_local
        update_status(f"Loading {file2_path}...")
        global g_dfs2
        g_dfs2 = pd_local.read_excel(file2_path, sheet_name=None) # Use pd_local

        timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M%S")

        file_to_highlight_path = ""
        if output_choice == "file1":
            file_to_highlight_path = file1_path
            output_base_name = os.path.basename(file1_path).replace(".xlsx", "").replace(".xls", "")
            output_file_name = f"{output_base_name}_{timestamp}_highlighted.xlsx"
            df_to_save_dict = g_dfs1
            wb_to_save = load_workbook(file1_path)
            update_status(f"Selected to highlight: {os.path.basename(file1_path)}")
        elif output_choice == "file2":
            file_to_highlight_path = file2_path
            output_base_name = os.path.basename(file2_path).replace(".xlsx", "").replace(".xls", "")
            output_file_name = f"{output_base_name}_{timestamp}_highlighted.xlsx"
            df_to_save_dict = g_dfs2
            wb_to_save = load_workbook(file2_path)
            update_status(f"Selected to highlight: {os.path.basename(file2_path)}")
        else:
            update_status("No output file chosen for highlighting.")
            return

        output_file_path = os.path.join(output_dir, output_file_name)

        def extract_imo_like_numbers(dfs_dict):
            numbers_set = set()
            for sheet_name, df in dfs_dict.items():
                if df.empty: continue
                normalized_columns = {str(col).strip().lower(): col for col in df.columns}
                
                imo_candidate_cols = []
                for pattern in IMO_LIKE_COLUMN_PATTERNS:
                    for norm_col, orig_col in normalized_columns.items():
                        if pattern in norm_col:
                            original_col_name = next(col_orig for col_orig in df.columns if col_orig.strip().lower() == norm_col)
                            imo_candidate_cols.append(original_col_name)
                
                for col_name in imo_candidate_cols:
                    for val in df[col_name].dropna():
                        cleaned_val = re.sub(r'\D', '', str(val)).strip() # Apply aggressive cleaning
                        if re.fullmatch(r'^\d{7}$', cleaned_val):
                            numbers_set.add(cleaned_val)
            return numbers_set

        imo_numbers_file1_set = extract_imo_like_numbers(g_dfs1)
        imo_numbers_file2_set = extract_imo_like_numbers(g_dfs2)
        
        common_red_numbers = imo_numbers_file1_set.intersection(imo_numbers_file2_set)
        excel_comparator_log_queue.put(f"[Excel Comparator] Found {len(common_red_numbers)} common 7-digit IMO numbers for highlighting.")

        def apply_highlights(workbook, dataframes_dict, target_red_nums, red_fill, status_update):
            for sheet_name, df in dataframes_dict.items():
                if sheet_name not in workbook.sheetnames:
                    status_update(f"Warning: Sheet '{sheet_name}' not found in workbook for highlighting. Skipping.")
                    continue
                ws = workbook[sheet_name]
                
                header_values = [cell.value for cell in ws[1]] # Assuming header is always in row 1
                
                imo_like_col_indices = []
                for idx, cell_header_value in enumerate(header_values):
                    if cell_header_value is not None and is_imo_like_column(str(cell_header_value)):
                        imo_like_col_indices.append(idx)

                for r_idx, row in enumerate(ws.iter_rows(min_row=2)): # Start from second row
                    for c_idx, cell in enumerate(row):
                        if c_idx in imo_like_col_indices:
                            cell_value = cell.value
                            if pd_local.notna(cell_value): # Use pd_local.notna
                                # Apply identical aggressive cleaning here
                                cleaned_cell_str_value = re.sub(r'\D', '', str(cell_value)).strip()
                                
                                # Diagnostic print for unhighlighted items that should be
                                if cleaned_cell_str_value in target_red_nums and (cell.fill is None or cell.fill.bgColor.rgb != 'FFFF0000'):
                                    excel_comparator_log_queue.put(f"DEBUG: MISSING HIGHLIGHT: Sheet='{sheet_name}', Row={r_idx+2}, Col={chr(65+c_idx)}, Raw='{cell_value}', Cleaned='{cleaned_cell_str_value}'")
                                
                                if re.fullmatch(r'^\d{7}$', cleaned_cell_str_value) and cleaned_cell_str_value in target_red_nums:
                                    cell.fill = red_fill

        update_status(f"Applying highlights to {os.path.basename(file_to_highlight_path)}...")
        apply_highlights(wb_to_save, df_to_save_dict, common_red_numbers, red_fill, update_status)
        
        wb_to_save.save(output_file_path)
        final_message = (f"Comparison complete. Highlighted file saved to:\n- {output_file_path}")

        update_status(final_message)
        excel_comparator_log_queue.put({'type': 'messagebox', 'kind': 'info', 'title': 'Success', 'message': final_message})
    except Exception as e:
        update_status(f"Error during comparison: {e}")
        excel_comparator_log_queue.put({'type': 'messagebox', 'kind': 'error', 'title': 'Error', 'message': f"An error occurred during comparison: {e}"})
    finally:
        excel_comparator_log_queue.put({'type': 'processing_done'})


# --- SanctionsAppFrame (Program 1 GUI encapsulated) ---
class SanctionsAppFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.is_processing = False
        self.excel_comparator_file1_callback = None # Callback to update ExcelComparator's file1 entry
        self.automatic_generation_complete_callback = None # New: Callback for automatic runs

        self.BG_COLOR = "#002060"
        self.FRAME_COLOR = "#103070"
        self.TEXT_COLOR = "#E0E0E0"
        self.ACCENT_COLOR = "#FFD700"
        self.BUTTON_HOVER_COLOR = "#184090"
        self.PULSE_COLOR = "#FFEEAA"

        self.title_font = font.Font(family="Georgia", size=16, weight="bold")
        self.label_font = font.Font(family="Verdana", size=10)
        self.button_font = font.Font(family="Verdana", size=11, weight="bold")
        self.log_font = font.Font(family="Consolas", size=9)

        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('TFrame', background=self.BG_COLOR)
        self.style.configure('TLabelframe', background=self.BG_COLOR, bordercolor=self.ACCENT_COLOR, relief="solid", borderwidth=2)
        self.style.configure('TLabelframe.Label', background=self.BG_COLOR, foreground=self.ACCENT_COLOR, font=self.title_font, padding=(10, 5))
        self.style.configure('TLabel', background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=self.label_font)
        self.style.configure('Royal.TButton', background=self.FRAME_COLOR, foreground=self.ACCENT_COLOR, 
                             font=self.button_font, bordercolor=self.ACCENT_COLOR, 
                             borderwidth=2, padding=(20, 10), relief="raised")
        self.style.map('Royal.TButton', 
                        background=[('active', self.BUTTON_HOVER_COLOR), ('pressed', self.BUTTON_HOVER_COLOR)], 
                        foreground=[('active', self.TEXT_COLOR), ('pressed', self.TEXT_COLOR)],
                        relief=[('pressed', 'sunken')])
        self.style.configure('Royal.Horizontal.TProgressbar', troughcolor=self.FRAME_COLOR, background=self.ACCENT_COLOR, 
                             bordercolor=self.ACCENT_COLOR, lightcolor=self.ACCENT_COLOR, darkcolor=self.ACCENT_COLOR)

        main_content_frame = ttk.Frame(self, padding="20")
        main_content_frame.pack(fill=tk.BOTH, expand=True)

        controls_frame = ttk.LabelFrame(main_content_frame, text="Controls", padding=(20,10))
        controls_frame.pack(fill=tk.X, pady=(0, 20), ipady=10)
        
        self.run_button = ttk.Button(controls_frame, text="Generate Sanctions Report", command=self.start_processing_thread, style='Royal.TButton')
        self.run_button.pack(pady=10)

        self.progress_canvas = tk.Canvas(main_content_frame, height=20, bg=self.FRAME_COLOR, highlightthickness=0)
        self.progress_canvas.pack(pady=10, fill=tk.X)
        self.pulse_rect = None

        log_frame = ttk.LabelFrame(main_content_frame, text="Log Output", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.log_area = tk.Text(log_frame, wrap=tk.WORD, state='disabled', font=self.log_font, bg=self.FRAME_COLOR, fg=self.TEXT_COLOR, relief=tk.FLAT, selectbackground=self.BUTTON_HOVER_COLOR, insertbackground=self.ACCENT_COLOR)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_area.yview)
        self.log_area['yscrollcommand'] = scrollbar.set
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Redirect stdout to log_area
        self.log_stream = TextRedirector(gui_update_queue)
        sys.stdout = self.log_stream
        sys.stderr = self.log_stream # Also redirect stderr for errors
        
        self.after(100, self.process_gui_queue)

    def set_excel_comparator_file1_callback(self, callback):
        """Sets the callback function to update ExcelComparator's file1 field."""
        self.excel_comparator_file1_callback = callback

    def start_processing_thread(self, on_complete_override=None):
        """
        Starts the report generation thread.
        Can be triggered manually (on_complete_override=None) or automatically.
        """
        output_filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save Sanctions Report As", initialfile="Sanctions_Report.xlsx")
        if not output_filepath:
            log_to_gui_and_console("File save cancelled.")
            if on_complete_override:
                on_complete_override(None) # Signal cancellation
            return

        self.is_processing = True
        self.run_button.config(state=tk.DISABLED)
        
        # Initialize progress_bar BEFORE animate_pulse and ensure it has a size
        self.progress_canvas.delete("progress_background") # Clear old background
        self.progress_canvas.delete("progress_bar") # Clear any old bar
        # Create background rectangle
        self.progress_canvas.create_rectangle(0, 0, self.progress_canvas.winfo_width(), 25, fill=self.FRAME_COLOR, outline="", tags="progress_background")
        # Create actual progress bar (initially zero width)
        self.progress_canvas.create_rectangle(0, 0, 0, 25, fill=self.ACCENT_COLOR, outline="", tags="progress_bar")
        self.master.update_idletasks() # Force update to ensure 'progress_bar' exists before animation starts
        time.sleep(0.01) # Small delay to help ensure canvas update

        self.update_progress(0, 7) # Initialize visual progress
        self.log_area.config(state='normal'); self.log_area.delete('1.0', tk.END); self.log_area.config(state='disabled')
        
        # Use provided callback, or default to updating Excel comparator
        callback_to_use = on_complete_override if on_complete_override else self.excel_comparator_file1_callback
        threading.Thread(target=process_all_data, args=(output_filepath, callback_to_use,), daemon=True).start()
        self.animate_pulse()

    def start_automatic_generation(self, callback_after_generation):
        """Public method to trigger a report generation automatically."""
        self.automatic_generation_complete_callback = callback_after_generation
        # Use a dummy filepath (won't be directly used, but needed for file dialog context)
        dummy_filepath = os.path.join(os.getcwd(), "Sanctions_Report.xlsx")
        self.start_processing_thread(on_complete_override=self._automatic_generation_complete)

    def _automatic_generation_complete(self, filepath):
        """Internal callback for automatic generation completion."""
        if self.automatic_generation_complete_callback:
            self.automatic_generation_complete_callback(filepath)
        self.automatic_generation_complete_callback = None # Clear callback

    def update_progress(self, value, max_val):
        canvas_width = self.progress_canvas.winfo_width()
        if canvas_width <= 1: # Canvas not fully rendered yet
            self.after(50, lambda: self.update_progress(value, max_val))
            return
            
        progress_width = (value / max_val) * canvas_width
        self.progress_canvas.coords("progress_bar", 0, 0, progress_width, 25) # Update coordinates of existing bar
        self.master.update_idletasks() # Force update after drawing
            
    def animate_pulse(self, x_pos=0, direction=1):
        if not self.is_processing:
            self.progress_canvas.delete("pulse")
            return

        canvas_width = self.progress_canvas.winfo_width()
        pulse_width = 40
        
        if self.pulse_rect:
            self.progress_canvas.delete(self.pulse_rect)

        self.pulse_rect = self.progress_canvas.create_rectangle(x_pos, 0, x_pos + pulse_width, 25, fill=self.PULSE_COLOR, tags="pulse", outline="")
        self.progress_canvas.tag_raise("pulse", "progress_bar") # Ensure pulse is on top of the bar
        
        new_x = x_pos + 5 * direction
        if new_x > canvas_width or new_x < -pulse_width:
             new_x = -pulse_width if direction == 1 else canvas_width
        
        self.after(15, lambda: self.animate_pulse(new_x, direction))

    def update_log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message) # Message already has newline from TextRedirector.write
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def process_gui_queue(self):
        try:
            while not gui_update_queue.empty():
                item = gui_update_queue.get_nowait()
                msg_type = item.get('type')
                if msg_type == 'log': self.update_log(item['message'])
                elif msg_type == 'progress_config': self.max_progress = item['max']
                elif msg_type == 'progress_update': self.update_progress(item['value'], item['max'])
                elif msg_type == 'processing_done':
                    self.is_processing = False
                    self.run_button.config(state=tk.NORMAL)
                    self.update_progress(self.max_progress, self.max_progress)
                elif msg_type == 'show_message':
                    if item['kind'] == 'info': messagebox.showinfo(item['title'], item['message'])
                    elif item['kind'] == 'error': messagebox.showerror(item['title'], item['message'])
        finally:
            self.after(100, self.process_gui_queue)


# --- ExcelComparatorFrame ---
class ExcelComparatorFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.is_processing = False
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_file_choice = tk.StringVar(value="file1")

        self.BG_COLOR = "#F0F0F0"
        self.TEXT_COLOR = "#333333"
        self.BUTTON_COLOR = "#0056B3"
        self.BUTTON_TEXT_COLOR = "white"
        self.FRAME_BG_COLOR = "#FFFFFF"

        self.style = ttk.Style(self)
        self.style.configure('TFrame', background=self.BG_COLOR)
        self.style.configure('TLabelframe', background=self.FRAME_BG_COLOR, bordercolor="#AAAAAA", relief="groove", borderwidth=1)
        self.style.configure('TLabelframe.Label', background=self.FRAME_BG_COLOR, foreground="#555555", font=("Arial", 11, "bold"))
        
        self.style.configure('TLabel', background=self.FRAME_BG_COLOR, foreground=self.TEXT_COLOR, font=("Arial", 10))
        self.style.configure('TRadiobutton', background=self.FRAME_BG_COLOR, foreground=self.TEXT_COLOR, font=("Arial", 10))
        
        self.style.configure('Excel.TButton', background=self.BUTTON_COLOR, foreground=self.BUTTON_TEXT_COLOR,
                             font=("Arial", 10, "bold"), relief="raised", borderwidth=2)
        self.style.map('Excel.TButton', background=[('active', '#004085')])

        main_frame = ttk.Frame(self, padding="20", style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="Select Excel Files", padding=(15,10))
        file_frame.pack(pady=10, fill=tk.X)

        ttk.Label(file_frame, text="Baseline Excel File:").grid(row=0, column=0, sticky="w", pady=5)
        self.entry1 = ttk.Entry(file_frame, textvariable=self.file1_path, width=60)
        self.entry1.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=lambda: self.browse_file(self.file1_path), style='Excel.TButton').grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(file_frame, text="Comparison Excel File:").grid(row=1, column=0, sticky="w", pady=5)
        self.entry2 = ttk.Entry(file_frame, textvariable=self.file2_path, width=60)
        self.entry2.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=lambda: self.browse_file(self.file2_path), style='Excel.TButton').grid(row=1, column=2, padx=5, pady=5)

        output_frame = ttk.LabelFrame(main_frame, text="Choose Output File for Highlighting", padding=(15,10))
        output_frame.pack(pady=10, fill=tk.X)
        ttk.Radiobutton(output_frame, text="Highlight Baseline File", variable=self.output_file_choice, value="file1", style='TRadiobutton').pack(anchor="w")
        ttk.Radiobutton(output_frame, text="Highlight Comparison File", variable=self.output_file_choice, value="file2", style='TRadiobutton').pack(anchor="w")

        self.compare_button = ttk.Button(main_frame, text="Compare and Highlight", command=self.run_comparison_threaded, style='Excel.TButton')
        self.compare_button.pack(pady=15)

        search_frame = ttk.LabelFrame(main_frame, text="Search Numeric Values", padding=(15,10))
        search_frame.pack(pady=10, fill=tk.X)

        ttk.Label(search_frame, text="Search Number:").grid(row=0, column=0, sticky="w", pady=5)
        self.search_entry = ttk.Entry(search_frame, width=30)
        self.search_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(search_frame, text="Search", command=self.perform_search, style='Excel.TButton').grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(main_frame, text="Status:", background=self.BG_COLOR, foreground=self.TEXT_COLOR).pack(anchor="w", padx=10)
        self.status_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, width=70, height=10, font=("Consolas", 9),
                                                     bg=self.FRAME_BG_COLOR, fg=self.TEXT_COLOR, relief=tk.FLAT)
        self.status_text.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        self.status_text.insert(tk.END, "Ready to compare and search Excel files.\n")
        self.status_text.config(state=tk.DISABLED)

        footnote_label = ttk.Label(main_frame, text="For any problems or queries, please contact: it@rcsclass.org",
                                     background=self.BG_COLOR, foreground="#666666", font=("Arial", 9, "italic"))
        footnote_label.pack(side=tk.BOTTOM, pady=(10, 0))

        self.after(100, self.process_excel_comparator_queue)

    def set_file1_path(self, filepath):
        self.file1_path.set(filepath)
        self.update_status_text(f"Sanctions report output loaded as Baseline Excel File: {filepath}\n")

    def browse_file(self, path_var):
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filepath:
            path_var.set(filepath)
            self.update_status_text(f"Selected: {filepath}\n")

    def update_status_text(self, message):
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message)
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

    def run_comparison_threaded(self):
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        output_choice = self.output_file_choice.get()

        if not file1 or not file2:
            messagebox.showwarning("Missing Files", "Please select both Excel files.")
            self.update_status_text("Error: Please select both Excel files.\n")
            return
        
        output_folder = filedialog.askdirectory(title="Select Output Folder for Highlighted Files")
        if not output_folder:
            self.update_status_text("Output folder selection cancelled.\n")
            return

        self.update_status_text("Comparison initiated...\n")
        self.is_processing = True
        self.compare_button.config(state=tk.DISABLED)
        self.search_entry.config(state=tk.DISABLED)
        threading.Thread(target=compare_and_highlight_excel_threaded, args=(file1, file2, output_choice, output_folder, self.update_status_text), daemon=True).start()

    def perform_search(self):
        search_term_str = self.search_entry.get().strip()
        if not search_term_str:
            messagebox.showwarning("Search Input", "Please enter a number to search.")
            self.update_status_text("Please enter a number to search.\n")
            return

        if not re.fullmatch(r'\d+', search_term_str):
            messagebox.showwarning("Invalid Input", "Please enter a valid whole number for search.")
            self.update_status_text("Invalid input for search. Please enter a whole number.\n")
            return
        
        search_num_str = search_term_str

        self.update_status_text(f"\nSearching for '{search_num_str}' (numeric only)...\n")
        
        if not g_dfs1 and not g_dfs2:
            self.status_text.config(state='normal')
            self.status_text.insert(tk.END, "No Excel files loaded for search. Please run comparison first.\n")
            self.status_text.config(state='disabled')
            return

        found_in_file1 = []
        found_in_file2 = []

        def search_df(dfs_dict, search_val_str, filename):
            found_locations = []
            for sheet_name, df in dfs_dict.items():
                if df.empty: continue
                for r_idx, row_series in df.iterrows():
                    for c_idx, cell_value in enumerate(row_series):
                        if pd.notna(cell_value):
                            cleaned_cell_val = re.sub(r'\D', '', str(cell_value)).strip()
                            if cleaned_cell_val == search_val_str:
                                found_locations.append(f"Sheet: '{sheet_name}', Row: {r_idx + 2}, Column Index: {c_idx} (Excel Column: {chr(65 + c_idx)})")
                
                for c_idx, col_name in enumerate(df.columns):
                    if pd.notna(col_name):
                        cleaned_col_name = re.sub(r'\D', '', str(col_name)).strip()
                        if cleaned_col_name == search_val_str:
                            found_locations.append(f"Sheet: '{sheet_name}', Header Row: 1, Column Index: {c_idx} (Excel Column: {chr(65 + c_idx)}) (Header)")
            return found_locations

        found_in_file1 = search_df(g_dfs1, search_num_str, os.path.basename(self.file1_path.get()))
        found_in_file2 = search_df(g_dfs2, search_num_str, os.path.basename(self.file2_path.get()))

        if found_in_file1:
            self.update_status_text(f"Found in Baseline Excel File ({os.path.basename(self.file1_path.get())}):\n" + "\n".join(found_in_file1) + "\n")
        else:
            self.update_status_text(f"'{search_num_str}' not found in Baseline Excel File (numeric only).\n")

        if found_in_file2:
            self.update_status_text(f"Found in Comparison Excel File ({os.path.basename(self.file2_path.get())}):\n" + "\n".join(found_in_file2) + "\n")
        else:
            self.update_status_text(f"'{search_num_str}' not found in Comparison Excel File (numeric only).\n")

    def process_excel_comparator_queue(self):
        try:
            while not excel_comparator_log_queue.empty():
                item = excel_comparator_log_queue.get_nowait()
                if isinstance(item, dict) and item.get('type') == 'messagebox':
                    if item['kind'] == 'info': messagebox.showinfo(item['title'], item['message'])
                    elif item['kind'] == 'error': messagebox.showerror(item['title'], item['message'])
                elif isinstance(item, dict) and item.get('type') == 'processing_done':
                    self.is_processing = False
                    self.compare_button.config(state=tk.NORMAL)
                    self.search_entry.config(state=tk.NORMAL)
                else:
                    self.update_status_text(item + "\n")
        finally:
            self.after(100, self.process_excel_comparator_queue)


# --- MyVesselsAppFrame (New Tab) ---
MY_VESSELS_CSV = "my_vessels.csv"

class MyVesselsAppFrame(ttk.Frame):
    def __init__(self, parent, get_last_sanctions_report_path_callback):
        super().__init__(parent)
        self.get_last_sanctions_report_path = get_last_sanctions_report_path_callback
        self.vessels_data = [] # List of {'name': 'Vessel Name', 'imo': 'IMO Number'} dicts
        self.load_vessels_from_csv()

        # Stores sanctioned IMO numbers from the last report run
        self.sanctioned_imos_from_report = set()
        # Stores {IMO: [Source1, Source2]} for sanctioned vessels
        self.sanctioned_vessel_sources = {} 

        # Internal flag for My Vessels sanctions check progress
        self.is_checking_sanctions = False 

        # Define colors specific to this frame's styling needs
        self.BG_COLOR = "#E8E8E8"
        self.TEXT_COLOR = "#333333"
        self.BUTTON_COLOR = "#007BFF"
        self.BUTTON_TEXT_COLOR = "white"
        self.FRAME_BG_COLOR = "#FFFFFF"
        self.ACCENT_COLOR = "#FFD700" # ACCENT_COLOR needed for pulse animation

        self.style = ttk.Style(self)
        self.style.configure('TFrame', background=self.BG_COLOR)
        self.style.configure('TLabelframe', background=self.FRAME_BG_COLOR, bordercolor="#AAAAAA", relief="groove", borderwidth=1)
        self.style.configure('TLabelframe.Label', background=self.FRAME_BG_COLOR, foreground="#555555", font=("Arial", 11, "bold"))
        
        self.style.configure('TLabel', background=self.FRAME_BG_COLOR, foreground=self.TEXT_COLOR, font=("Arial", 10))
        self.style.configure('MyVessels.TButton', background=self.BUTTON_COLOR, foreground=self.BUTTON_TEXT_COLOR,
                             font=("Arial", 10, "bold"), relief="raised", borderwidth=2)
        self.style.map('MyVessels.TButton', background=[('active', '#0056B3')])

        self.style.configure("Treeview",
                             background="#F8F8F8",
                             foreground=self.TEXT_COLOR,
                             rowheight=25,
                             fieldbackground="#F8F8F8",
                             font=("Arial", 10))
        # tag_configure must be called AFTER the Treeview widget is created.
        self.style.map('Treeview', background=[('selected', '#007BFF')]) # Keep selected row style
        self.style.configure("Treeview.Heading",
                             font=("Arial", 10, "bold"),
                             background="#DDDDDD",
                             foreground="#333333",
                             relief="raised")

        main_frame = ttk.Frame(self, padding="20", style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input Frame
        input_frame = ttk.LabelFrame(main_frame, text="Add New Vessel", padding=(15,10))
        input_frame.pack(pady=10, fill=tk.X)

        ttk.Label(input_frame, text="Vessel Name:").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        self.vessel_name_entry = ttk.Entry(input_frame, width=40)
        self.vessel_name_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="IMO Number:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        self.imo_number_entry = ttk.Entry(input_frame, width=40)
        self.imo_number_entry.grid(row=1, column=1, padx=5, pady=5)

        add_button = ttk.Button(input_frame, text="Add Vessel", command=self.add_vessel, style='MyVessels.TButton')
        add_button.grid(row=2, column=0, columnspan=2, pady=10)

        # Search & Sanction Check Frame
        search_check_frame = ttk.LabelFrame(main_frame, text="Search & Sanction Check", padding=(15,10))
        search_check_frame.pack(pady=10, fill=tk.X)

        ttk.Label(search_check_frame, text="Search Term (Name/IMO):").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        self.search_vessel_entry = ttk.Entry(search_check_frame, width=30)
        self.search_vessel_entry.grid(row=0, column=1, padx=5, pady=5)
        self.search_vessel_entry.bind("<KeyRelease>", self.apply_vessel_search_filter) # Live search filter
        
        self.check_sanctions_button = ttk.Button(search_check_frame, text="Check Sanctions", command=self.compare_my_vessels_threaded, style='MyVessels.TButton')
        self.check_sanctions_button.grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Button(search_check_frame, text="Clear Search", command=lambda: (self.search_vessel_entry.delete(0, tk.END), self.apply_vessel_search_filter(event=None)), style='MyVessels.TButton').grid(row=0, column=3, padx=5, pady=5)

        # Progress bar for sanctions check
        self.check_progress_canvas = tk.Canvas(search_check_frame, height=15, bg=self.FRAME_BG_COLOR, highlightthickness=0)
        self.check_progress_canvas.grid(row=1, column=0, columnspan=4, pady=5, sticky="ew")
        self.check_pulse_rect = None


        # Vessel List Frame
        list_frame = ttk.LabelFrame(main_frame, text="My Vessels List", padding=(15,10))
        list_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # Define columns including the new "No." column
        self.vessel_tree = ttk.Treeview(list_frame, columns=("No.", "Vessel Name", "IMO Number", "Sanctioned?", "Sources"), show="headings")
        self.vessel_tree.heading("No.", text="No.")
        self.vessel_tree.heading("Vessel Name", text="Vessel Name")
        self.vessel_tree.heading("IMO Number", text="IMO Number")
        self.vessel_tree.heading("Sanctioned?", text="Sanctioned?")
        self.vessel_tree.heading("Sources", text="Sources")
        
        # Configure column widths
        self.vessel_tree.column("No.", width=40, anchor="center")
        self.vessel_tree.column("Vessel Name", width=200, anchor="w")
        self.vessel_tree.column("IMO Number", width=100, anchor="center")
        self.vessel_tree.column("Sanctioned?", width=80, anchor="center")
        self.vessel_tree.column("Sources", width=250, anchor="w")

        # Corrected placement: tag_configure belongs to the Treeview instance
        self.vessel_tree.tag_configure('sanctioned', background='#FFCCCC', foreground='#CC0000', font=('Arial', 10, 'bold')) # Light red background, dark red text


        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.vessel_tree.yview)
        self.vessel_tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y")
        self.vessel_tree.pack(fill="both", expand=True)

        remove_button = ttk.Button(list_frame, text="Remove Selected Vessel(s)", command=self.remove_selected_vessels, style='MyVessels.TButton')
        remove_button.pack(pady=5)

        # Export My Vessels to Excel button (now standalone)
        export_button = ttk.Button(main_frame, text="Export My Vessels to Excel", command=self.export_my_vessels, style='MyVessels.TButton')
        export_button.pack(pady=10, fill=tk.X)


        ttk.Label(main_frame, text="Status:", background=self.BG_COLOR, foreground=self.TEXT_COLOR).pack(anchor="w", padx=10)
        self.status_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=5, font=("Consolas", 9),
                                                     bg=self.FRAME_BG_COLOR, fg=self.TEXT_COLOR, relief=tk.FLAT)
        self.status_text.pack(padx=10, pady=5, fill=tk.X, expand=False)
        self.status_text.insert(tk.END, "Manage your personal vessel list here.\n")
        self.status_text.config(state=tk.DISABLED)

        self.populate_vessel_tree() # Initial population
        self.after(100, self.process_my_vessels_queue) # Start queue processing

    def load_vessels_from_csv(self):
        self.vessels_data = []
        if os.path.exists(MY_VESSELS_CSV):
            try:
                with open(MY_VESSELS_CSV, mode='r', newline='', encoding='utf-8') as file:
                    reader = csv.DictReader(file)
                    for row in reader:
                        if 'name' in row and 'imo' in row:
                            self.vessels_data.append({'name': row['name'], 'imo': row['imo']})
            except Exception as e:
                my_vessels_log_queue.put(f"Error loading vessels from CSV: {e}")
        my_vessels_log_queue.put(f"Loaded {len(self.vessels_data)} vessels from {MY_VESSELS_CSV}")

    def save_vessels_to_csv(self):
        try:
            with open(MY_VESSELS_CSV, mode='w', newline='', encoding='utf-8') as file:
                fieldnames = ['name', 'imo']
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(self.vessels_data)
            my_vessels_log_queue.put(f"Saved {len(self.vessels_data)} vessels to {MY_VESSELS_CSV}")
        except Exception as e:
            my_vessels_log_queue.put(f"Error saving vessels to CSV: {e}")

    def populate_vessel_tree(self, search_term=""):
        """Populates the Treeview, optionally filtering by search_term and applying highlights."""
        for item in self.vessel_tree.get_children():
            self.vessel_tree.delete(item)
        
        search_term_lower = search_term.lower().strip()

        displayed_count = 0
        row_number = 1 
        for vessel in self.vessels_data:
            vessel_name_lower = vessel['name'].lower()
            vessel_imo_cleaned = re.sub(r'\D', '', str(vessel['imo'])).strip()
            
            # Apply search filter
            if search_term_lower:
                if search_term_lower not in vessel_name_lower and \
                   search_term_lower not in vessel_imo_cleaned:
                    continue # Skip if no match
            
            # Determine sanction status and sources
            tags = ()
            sanctioned_text = "No"
            sources_text = ""
            if vessel_imo_cleaned in self.sanctioned_imos_from_report:
                tags = ('sanctioned',) # Apply the red highlight tag
                sanctioned_text = "Yes"
                sources_list = self.sanctioned_vessel_sources.get(vessel_imo_cleaned, [])
                sources_text = ", ".join(sources_list)

            # Insert with the new sequential number as the first value
            self.vessel_tree.insert("", "end", values=(row_number, vessel['name'], vessel['imo'], sanctioned_text, sources_text), tags=tags)
            displayed_count += 1
            row_number += 1 # Increment row number for next visible item
        
        self.update_status_text(f"Displaying {displayed_count} vessels. Sanction check results visible.")


    def add_vessel(self):
        name = self.vessel_name_entry.get().strip()
        imo = self.imo_number_entry.get().strip()

        if not name or not imo:
            messagebox.showwarning("Input Error", "Both Vessel Name and IMO Number are required.")
            return
        
        if not re.fullmatch(r'^\d{7}$', imo):
            messagebox.showwarning("Input Error", "IMO Number must be exactly 7 digits.")
            return

        for vessel in self.vessels_data:
            if vessel['imo'] == imo:
                messagebox.showwarning("Duplicate Entry", f"Vessel with IMO Number {imo} already exists.")
                return

        self.vessels_data.append({'name': name, 'imo': imo})
        self.save_vessels_to_csv()
        self.vessel_name_entry.delete(0, tk.END)
        self.imo_number_entry.delete(0, tk.END)
        self.update_status_text(f"Added vessel: {name} ({imo}).")
        self.populate_vessel_tree(search_term=self.search_vessel_entry.get()) # Refresh display after adding

    def remove_selected_vessels(self):
        selected_items = self.vessel_tree.selection()
        if not selected_items:
            messagebox.showwarning("Selection Error", "Please select at least one vessel to remove.")
            return

        if messagebox.askyesno("Confirm Removal", f"Are you sure you want to remove {len(selected_items)} selected vessel(s)?"):
            removed_count = 0
            # Get IMO numbers from Treeview value (index 2 for IMO after 'No.' and 'Vessel Name')
            imos_to_remove = [self.vessel_tree.item(item, 'values')[2] for item in selected_items] 
            
            self.vessels_data = [v for v in self.vessels_data if v['imo'] not in imos_to_remove]
            removed_count = len(imos_to_remove)
            
            self.save_vessels_to_csv()
            self.update_status_text(f"Removed {removed_count} vessel(s).")
            self.populate_vessel_tree(search_term=self.search_vessel_entry.get()) # Refresh display after removal

    def export_my_vessels(self):
        if not self.vessels_data:
            messagebox.showinfo("No Data", "No vessels to export.")
            return

        timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
        initial_filename = f"My_Vessels_List_{timestamp}.xlsx"

        output_filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Export My Vessels As", initialfile=initial_filename)
        if not output_filepath:
            self.update_status_text("Export cancelled.\n")
            return

        try:
            # Import pandas inside this function to ensure it's available in the thread
            import pandas as pd_local 

            # Create a DataFrame that matches the Treeview's displayed columns
            export_data = []
            for item_id in self.vessel_tree.get_children():
                # Extract values as they appear in the Treeview, including No., Sanctioned?, Sources
                values = self.vessel_tree.item(item_id, 'values')
                export_data.append({
                    "No.": values[0],
                    "Vessel Name": values[1],
                    "IMO Number": values[2],
                    "Sanctioned?": values[3],
                    "Sources": values[4]
                })
            df = pd_local.DataFrame(export_data) # Use pd_local
            
            df.to_excel(output_filepath, index=False)
            messagebox.showinfo("Export Successful", f"My Vessels list exported to:\n{output_filepath}")
            self.update_status_text(f"My Vessels list exported to: {output_filepath}\n")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export vessels: {e}")
            self.update_status_text(f"Error exporting vessels: {e}\n")

    def compare_my_vessels_threaded(self):
        """
        Initiates the sanctions check for My Vessels.
        Does NOT automatically trigger report generation. User must generate report separately.
        """
        if not self.vessels_data:
            messagebox.showinfo("No Vessels", "Please add vessels to 'My Vessels' list before checking sanctions.")
            return

        sanctions_report_path = self.get_last_sanctions_report_path()
        if not sanctions_report_path or not os.path.exists(sanctions_report_path):
            messagebox.showwarning("Sanctions Report Missing",
                                   "No sanctions report found. Please generate a report in the 'Sanctions Report Generator' tab first, or load one into the 'Excel Comparator' tab as 'Baseline File'.")
            return

        self.update_status_text("Starting sanctions check for My Vessels...")
        self.is_checking_sanctions = True
        self.check_sanctions_button.config(state=tk.DISABLED) # Disable button during check
        
        # Clear previous sanction status before running new check
        self.sanctioned_imos_from_report.clear()
        self.sanctioned_vessel_sources.clear()
        self.populate_vessel_tree(search_term=self.search_vessel_entry.get()) # Clear old highlights
        
        # Reset and start progress animation
        # Initialize canvas background and bar here explicitly to avoid TclError
        self.check_progress_canvas.delete("check_progress_background")
        self.check_progress_canvas.delete("check_progress_bar")
        self.check_progress_canvas.create_rectangle(0, 0, self.check_progress_canvas.winfo_width(), 15, fill=self.FRAME_BG_COLOR, outline="", tags="check_progress_background")
        self.check_progress_canvas.create_rectangle(0, 0, 0, 15, fill=self.BUTTON_COLOR, outline="", tags="check_progress_bar")
        self.master.update_idletasks() # Force update
        time.sleep(0.01) # Small delay to help ensure canvas update
        
        self.update_check_progress(0, 100) # Progress in percentage (0-100)
        self.animate_check_pulse()

        threading.Thread(target=self.run_vessel_sanctions_check,
                         args=(list(self.vessels_data), sanctions_report_path),
                         daemon=True).start()

    def run_vessel_sanctions_check(self, my_vessels, sanctions_report_path):
        """Worker function to perform the actual sanctions comparison."""
        try:
            # Import pandas inside this function to ensure it's available in the thread
            import pandas as pd_local 

            # Load all sanctions data from the generated report
            sanctions_dfs = pd_local.read_excel(sanctions_report_path, sheet_name=None)
            
            all_sanctioned_imos_temp = set()
            sanctioned_imo_details_temp = {}

            # Phase 1: Populate sanctioned_imos_temp from all loaded sanctions sheets
            # This loop contributes to the first half (50%) of the progress bar.
            total_sanctioned_sheets_rows = sum(len(df) for df in sanctions_dfs.values())
            processed_sanction_rows = 0

            my_vessels_log_queue.put({'type': 'log', 'message': f"MV Check: Loading {len(sanctions_dfs)} sanctions lists from report..."})

            for source_name, df in sanctions_dfs.items():
                if df.empty:
                    my_vessels_log_queue.put({'type': 'log', 'message': f"MV Check: Skipping empty sheet: {source_name}"})
                    continue
                
                imo_col_for_sheet = None
                for col in df.columns:
                    cleaned_col_name = str(col).strip().lower()
                    for pattern in IMO_LIKE_COLUMN_PATTERNS:
                        if pattern in cleaned_col_name:
                            imo_col_for_sheet = col # Found the original column name
                            break
                    if imo_col_for_sheet: break

                if imo_col_for_sheet and imo_col_for_sheet in df.columns:
                    for imo_val in df[imo_col_for_sheet].dropna().astype(str).tolist():
                        imo_val = re.sub(r'\D', '', str(imo_val)).strip() # Clean consistently
                        if re.fullmatch(r'^\d{7}$', imo_val):
                            all_sanctioned_imos_temp.add(imo_val)
                            sanctioned_imo_details_temp.setdefault(imo_val, []).append(source_name)
                
                processed_sanction_rows += len(df) # Update progress based on rows processed in each sheet
                # Progress is for loading sanctions data, maxing out at 50%
                progress_percentage = int((processed_sanction_rows / total_sanctioned_sheets_rows) * 50) if total_sanctioned_sheets_rows > 0 else 0
                my_vessels_log_queue.put({'type': 'progress_update', 'value': progress_percentage, 'max': 100})
                my_vessels_log_queue.put({'type': 'log', 'message': f"MV Check: Processed {len(df)} entries from {source_name}."})
            
            # Update the instance variables
            self.sanctioned_imos_from_report = all_sanctioned_imos_temp
            self.sanctioned_vessel_sources = sanctioned_imo_details_temp
            
            my_vessels_log_queue.put({'type': 'log', 'message': f"MV Check: Loaded {len(all_sanctioned_imos_temp)} unique sanctioned IMOs."})
            my_vessels_log_queue.put({'type': 'log', 'message': f"Comparing {len(my_vessels)} of your vessels against loaded sanctions data..."})

            # Phase 2: Iterate through my_vessels and update progress (remaining 50%)
            num_my_vessels = len(my_vessels)
            checked_vessel_count = 0
            for vessel in my_vessels:
                # The actual comparison is fast since sanctioned_imos_from_report is a set.
                vessel_imo_cleaned = re.sub(r'\D', '', str(vessel['imo'])).strip() # Ensure consistency
                # No need to explicitly check here, populate_vessel_tree will use the updated sets.
                
                checked_vessel_count += 1
                # Progress for checking My Vessels starts from 50% and goes to 100%
                progress_percentage = 50 + int((checked_vessel_count / num_my_vessels) * 50) if num_my_vessels > 0 else 50
                my_vessels_log_queue.put({'type': 'progress_update', 'value': progress_percentage, 'max': 100})
                # Add a small sleep to ensure the progress bar is visible
                time.sleep(0.005) 

            my_vessels_log_queue.put({'type': 'log', 'message': f"Sanctions check complete for My Vessels."})
            my_vessels_log_queue.put({'type': 'refresh_tree'}) # Signal GUI to refresh treeview

        except Exception as e:
            my_vessels_log_queue.put({'type': 'messagebox', 'kind': 'error', 'title': 'Sanctions Check Error',
                                       'message': f"An error occurred during sanctions check: {e}. Traceback: {sys.exc_info()}"})
            my_vessels_log_queue.put({'type': 'log', 'message': f"Error during sanctions check: {e}\n"})
        finally:
            my_vessels_log_queue.put({'type': 'processing_done'}) # Signal to re-enable button and stop animation

    def update_check_progress(self, value, max_val):
        """Updates the progress bar drawn on the canvas for My Vessels check."""
        # Ensure canvas dimensions are valid before drawing
        canvas_width = self.check_progress_canvas.winfo_width()
        canvas_height = self.check_progress_canvas.winfo_height()
        if canvas_width <= 1 or canvas_height <= 1:
            self.after(50, lambda: self.update_check_progress(value, max_val))
            return
            
        progress_width = (value / max_val) * canvas_width if max_val > 0 else 0
        
        # Update the 'check_progress_bar' item's coordinates
        self.check_progress_canvas.coords("check_progress_bar", 0, 0, progress_width, canvas_height)
        self.master.update_idletasks() # Force redraw

    def animate_check_pulse(self, x_pos=0, direction=1):
        """Animates the pulse on the My Vessels progress bar."""
        if not self.is_checking_sanctions: # Stop animation if check is done
            self.check_progress_canvas.delete("check_pulse")
            return

        canvas_width = self.check_progress_canvas.winfo_width()
        canvas_height = self.check_progress_canvas.winfo_height()
        pulse_width = 30
        
        if self.check_pulse_rect:
            self.check_progress_canvas.delete(self.check_pulse_rect)

        # Draw a transparent pulse effect over the progress bar, matching FRAME_BG_COLOR for contrast
        self.check_pulse_rect = self.check_progress_canvas.create_rectangle(x_pos, 0, x_pos + pulse_width, canvas_height, fill=self.ACCENT_COLOR, stipple="gray50", tags="check_pulse", outline="")
        self.check_progress_canvas.tag_raise("check_pulse", "check_progress_bar") # Ensure pulse is on top of bar
        
        # Move the pulse
        new_x = x_pos + 4 * direction # Speed of pulse
        # Reverse direction if hitting ends
        if new_x > canvas_width or new_x < -pulse_width:
             direction *= -1
             new_x = max(-pulse_width, min(new_x, canvas_width)) # Ensure it's within bounds after flip
        
        self.after(15, lambda: self.animate_check_pulse(new_x, direction)) # Recursive call for animation


    def apply_vessel_search_filter(self, event=None):
        """Applies filter to the Treeview based on search entry content."""
        self.populate_vessel_tree(search_term=self.search_vessel_entry.get())


    def update_status_text(self, message):
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message)
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

    def process_my_vessels_queue(self):
        try:
            while not my_vessels_log_queue.empty():
                item = my_vessels_log_queue.get_nowait()
                if isinstance(item, dict) and item.get('type') == 'messagebox':
                    if item['kind'] == 'info': messagebox.showinfo(item['title'], item['message'])
                    elif item['kind'] == 'error': messagebox.showerror(item['title'], item['message'])
                elif isinstance(item, dict) and item.get('type') == 'processing_done':
                    self.is_checking_sanctions = False # Stop animation
                    self.check_sanctions_button.config(state=tk.NORMAL) # Re-enable button
                    self.update_check_progress(100, 100) # Ensure bar is full at end
                    my_vessels_log_queue.put({'type': 'log', 'message': "My Vessels sanctions check process finished."})
                elif isinstance(item, dict) and item.get('type') == 'refresh_tree':
                    # Call refresh_vessel_tree_display with current search term
                    self.populate_vessel_tree(search_term=self.search_vessel_entry.get())
                elif isinstance(item, dict) and item.get('type') == 'progress_update':
                    self.update_check_progress(item['value'], item['max'])
                else:
                    self.update_status_text(item + "\n")
        finally:
            self.after(100, self.process_my_vessels_queue)


# --- AboutAppFrame (New Tab) ---
class AboutAppFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.BG_COLOR = "#F0F0F0"
        self.TEXT_COLOR = "#333333"
        self.TITLE_COLOR = "#002060"
        self.FOOTNOTE_COLOR = "#666666"

        self.style = ttk.Style(self)
        self.style.configure('About.TFrame', background=self.BG_COLOR)
        self.style.configure('About.TLabel', background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=("Arial", 10))
        self.style.configure('About.Title.TLabel', background=self.BG_COLOR, foreground=self.TITLE_COLOR, font=("Georgia", 16, "bold"))
        self.style.configure('About.SubTitle.TLabel', background=self.BG_COLOR, foreground=self.TITLE_COLOR, font=("Georgia", 12, "bold"))
        self.style.configure('About.Footnote.TLabel', background=self.BG_COLOR, foreground=self.FOOTNOTE_COLOR, font=("Arial", 9, "italic"))

        main_frame = ttk.Frame(self, padding="20", style='About.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="About Royal Classification Society Vessel Sanctions & Excel Tool", style='About.Title.TLabel').pack(pady=10)

        ttk.Label(main_frame, text="Purpose:", style='About.SubTitle.TLabel').pack(anchor="w", pady=(10, 0))
        ttk.Label(main_frame, text=(
            "This application serves as a comprehensive tool designed for the maritime industry, "
            "specifically for compliance and risk assessment. It automates the process of exploring "
            "the internet to fetch and consolidate critical vessel sanctions data from various "
            "authoritative sources, and provides robust tools for local data comparison and management."
        ), style='About.TLabel', wraplength=700, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))

        ttk.Label(main_frame, text="How to Use This Program:", style='About.SubTitle.TLabel').pack(anchor="w", pady=(10, 0))

        ttk.Label(main_frame, text="1. Sanctions Report Generator Tab:", style='About.SubTitle.TLabel').pack(anchor="w", pady=(5, 0))
        ttk.Label(main_frame, text=(
            "   • Click 'Generate Sanctions Report'.\n"
            "   • A 'Save As' dialog will appear. Choose a location and filename (e.g., 'Sanctions_Report.xlsx') "
            "to save the generated Excel file.\n"
            "   • The application will then connect to various online sources (OFAC, UK, EU DMA, UANI Google Sheet) "
            "to fetch the latest vessel sanctions data.\n"
            "   • Progress and status messages will be displayed in the 'Log Output' area.\n"
            "   • Once complete, an Excel file will be saved containing separate sheets for each sanctions source. "
            "This generated file will automatically populate the 'Baseline Excel File' field in the 'Excel Comparator' tab."
        ), style='About.TLabel', wraplength=700, justify=tk.LEFT).pack(anchor="w", pady=(0, 5))

        ttk.Label(main_frame, text="2. Excel Comparator Tab:", style='About.SubTitle.TLabel').pack(anchor="w", pady=(5, 0))
        ttk.Label(main_frame, text=(
            "   • Select your 'Baseline Excel File' and 'Comparison Excel File' using the 'Browse' buttons. "
            "The baseline file can be automatically loaded from the Sanctions Generator tab.\n"
            "   • Choose which of the two files you want to be highlighted as the output.\n"
            "   • Click 'Compare and Highlight'. A folder selection dialog will prompt you to choose where "
            "the 'highlighted_excel_files' subfolder should be created.\n"
            "   • Common numerical identifiers will be highlighted in red in the chosen output file, "
            "which will be saved in the selected output folder.\n"
            "   • Use the 'Search Numeric Values' section to find specific numbers within the currently loaded "
            "Excel files (after a comparison has been run)."
        ), style='About.TLabel', wraplength=700, justify=tk.LEFT).pack(anchor="w", pady=(0, 5))

        ttk.Label(main_frame, text="3. My Vessels Tab:", style='About.SubTitle.TLabel').pack(anchor="w", pady=(5, 0))
        ttk.Label(main_frame, text=(
            "   • Add your own vessel names and IMO numbers using the 'Add New Vessel' section. "
            "IMO numbers must be exactly 7 digits.\n"
            "   • Your list of vessels will be displayed in the table and automatically saved for future sessions.\n"
            "   • Select one or more vessels from the list and 'Remove Selected Vessel(s)' to delete them.\n"
            "   • 'Export My Vessels to Excel': Saves your current list of vessels to a new Excel file.\n"
            "   • 'Compare with Sanctions Report': Generates a new Excel report that checks your vessels "
            "against the most recently generated sanctions data. This report will indicate if your vessels "
            "are found in any sanctions list and specify the source(s)."
        ), style='About.TLabel', wraplength=700, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))
        
        ttk.Label(main_frame, text="Contact Information:", style='About.SubTitle.TLabel').pack(anchor="w", pady=(10, 0))
        ttk.Label(main_frame, text="For technical assistance or inquiries, please contact: it@rcsclass.org", 
                     style='About.TLabel').pack(anchor="w", pady=(0, 10))

        main_frame.grid_rowconfigure(0, weight=0)
        main_frame.grid_rowconfigure(1, weight=0)
        main_frame.grid_rowconfigure(2, weight=0)
        main_frame.grid_rowconfigure(3, weight=0)
        main_frame.grid_rowconfigure(4, weight=0)
        main_frame.grid_rowconfigure(5, weight=0)
        main_frame.grid_rowconfigure(6, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

# --- Main RoyalClassificationApp ---
class RoyalClassificationApp(tk.Tk):
    def __init__(self):
        # Dependencies Check: Ensures all required libraries are installed before starting GUI
        super().__init__()
        self.title("ROYAL CLASSIFICATION SOCIETY")
        self.geometry("1000x700")
        self.minsize(900, 600)

        # Store path to the last generated sanctions report and a flag if My Vessels is waiting
        self.last_sanctions_report_path = None
        self.my_vessels_tab_needs_sanction_check = False # This flag is now less crucial without auto-trigger

        self.configure(bg="#002060")

        app_style = ttk.Style(self)
        app_style.theme_use('clam')
        
        app_style.configure('TNotebook', background="#002060", borderwidth=0)
        app_style.map('TNotebook.Tab',
                      background=[('selected', '#103070'), ('!selected', '#002060')],
                      foreground=[('selected', '#FFD700'), ('!selected', '#E0E0E0')],
                      font=[('selected', ('Verdana', 10, 'bold')), ('!selected', ('Verdana', 10))]
                     )
        app_style.configure('TNotebook.Tab', padding=[10, 5], focuscolor=app_style.lookup('TNotebook.Tab', 'background'))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Instantiate tab frames - now correctly ordered AFTER their class definitions
        self.sanctions_tab = SanctionsAppFrame(self.notebook)
        self.my_vessels_tab = MyVesselsAppFrame(self.notebook, self.get_last_sanctions_report_path)
        self.excel_comparator_tab = ExcelComparatorFrame(self.notebook)
        self.about_tab = AboutAppFrame(self.notebook)

        # Set callbacks between tabs
        self.sanctions_tab.set_excel_comparator_file1_callback(self.set_excel_comparator_file1_and_sanctions_path)
        
        # This callback is now simpler as MyVessels no longer auto-triggers SanctionsGen.
        # It ensures that when SanctionsAppFrame completes, this master method is notified.
        self.sanctions_tab.automatic_generation_complete_callback = self.on_sanctions_report_complete

        # Add tabs to notebook in desired order
        self.notebook.add(self.sanctions_tab, text="Sanctions Report Generator")
        self.notebook.add(self.my_vessels_tab, text="My Vessels") # My Vessels is now tab 2
        self.notebook.add(self.excel_comparator_tab, text="Excel Comparator") # Excel Comparator is now tab 3
        self.notebook.add(self.about_tab, text="About This Application")

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def set_excel_comparator_file1_and_sanctions_path(self, filepath):
        """
        Callback from SanctionsAppFrame to update ExcelComparator's file1 and
        store the path as the last generated sanctions report.
        """
        self.excel_comparator_tab.set_file1_path(filepath)
        self.last_sanctions_report_path = filepath
        print(f"Last sanctions report path updated to: {self.last_sanctions_report_path}")

    def get_last_sanctions_report_path(self):
        """Provides the last generated sanctions report path to MyVesselsAppFrame."""
        return self.last_sanctions_report_path

    def on_sanctions_report_complete(self, filepath):
        """
        Callback triggered by SanctionsAppFrame when a report generation (manual or automatic) is complete.
        Used to update the last_sanctions_report_path.
        """
        self.set_excel_comparator_file1_and_sanctions_path(filepath) # Update Excel Comparator and last path
        
        # Removed the logic that would trigger MyVessels check automatically
        # Now, MyVessels "Check Sanctions" button will simply use this newly generated file.

    def on_closing(self):
        """Handles application closing, ensuring My Vessels data is saved."""
        self.my_vessels_tab.save_vessels_to_csv()
        self.destroy()

# --- Main Application Execution ---
if __name__ == "__main__":
    # Dependency Check: Ensures all required libraries are installed before starting GUI
    missing_libs_to_install = []
    if sys.version_info < (3, 7):
        messagebox.showerror("Python Version Error", "This application requires Python 3.7 or higher.")
        sys.exit(1)

    # Check common libraries
    try: import requests
    except ImportError: missing_libs_to_install.append('requests')
    try: import pandas # Removed 'as pd' here
    except ImportError: missing_libs_to_install.append('pandas')
    try: import openpyxl
    except ImportError: missing_libs_to_install.append('openpyxl')
    try: from bs4 import BeautifulSoup
    except ImportError: missing_libs_to_install.append('beautifulsoup4')
    try: import lxml
    except ImportError: missing_libs_to_install.append('lxml')
    # UserAgent import handled within try/except block at the top

    # Check for requests.adapters and urllib3 for retry logic, if requests is present
    if 'requests' not in missing_libs_to_install:
        try:
            from requests.adapters import HTTPAdapter
            from urllib3.util.retry import Retry
        except ImportError:
            missing_libs_to_install.append('urllib3') # urllib3 contains Retry functionality
            # requests is already in list if it failed, no need to add HTTPAdapter here.

    if missing_libs_to_install:
        try:
            root_check = tk.Tk()
            root_check.withdraw()
            messagebox.showerror("Dependency Error",
                f"The following libraries are required:\n\n{', '.join(sorted(list(set(missing_libs_to_install))))}\n\n"
                f"Please install them by running:\npip install {' '.join(sorted(list(set(missing_libs_to_install))))}")
            root_check.destroy()
        except tk.TclError:
            print(f"CRITICAL ERROR: Missing libraries: {', '.join(sorted(list(set(missing_libs_to_install))))}")
            print(f"Please run: pip install {' '.join(sorted(list(set(missing_libs_to_install))))}")
        sys.exit(1)

    app = RoyalClassificationApp()
    app.mainloop()