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
import importlib.resources # For accessing bundled files

# --- Configuration for report freshness ---
MAX_REPORT_AGE_SECONDS = 3600 * 24 # 24 hours (adjust as needed)

# --- Global Queues for thread-safe GUI updates ---
gui_update_queue = queue.Queue() # For SanctionsAppFrame
my_vessels_log_queue = queue.Queue() # For MyVesselsAppFrame
sanctions_display_queue = queue.Queue() # For SanctionsDisplayFrame

# --- Global Data Store for fetched Sanctions Data (in-memory) ---
# This will hold the fetched DataFrames for easy access across tabs
global_sanctions_data_store = {
    "OFAC_Vessels": None,
    "UK_Sanctions_Vessels": None,
    "EU_DMA_Vessels": None,
    "UANI_Vessels_Tracked": None
}

# --- Configuration ---
OFAC_SDN_CSV_URL = "https://www.treasury.gov/ofac/downloads/sdn.csv"
UK_SANCTIONS_PUBLICATION_PAGE_URL = "https://www.gov.uk/government/publications/the-uk-sanctions-list"
DMA_XLSX_URL = "https://www.dma.dk/Media/638834044135010725/2025118019-7%20Importversion%20-%20List%20of%20EU%20designated%20vessels%20(20-05-2025)%203010691_2_0.XLSX"
UANI_WEBSCRAPE_URL = "https://www.unitedagainstnucleariran.com/blog/switch-list-tankers-shift-from-carrying-iranian-oil-to-russian-oil"
UANI_BUNDLED_CSV_NAME = "UANI_Switch_List_Bundled.csv" # Name of the CSV file to bundle

# Define common column patterns for numerical identifiers like IMO numbers
IMO_LIKE_COLUMN_PATTERNS = ['imo', 'id', 'number', 'code', 'vesselimo', 'imo no']


# --- Global User-Agent Initializer & Retry setup ---
try:
    import requests # Ensure requests is imported
    import pandas as pd # Import pandas as pd
    from bs4 import BeautifulSoup
    import lxml # Explicitly try to import lxml parser
    import openpyxl # Keep for type hinting and potential future use, though not directly used in fetching anymore
    from openpyxl.styles import PatternFill # Keep for potential future use or if some logic still relies on it
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
ua = None # Initialize ua outside try-block to ensure it always exists
http_session = None # Initialize http_session outside try-block

try:
    # Initialize UserAgent first, with a more specific fallback if it fails
    try:
        ua = UserAgent()
        print("DEBUG: fake_useragent initialized successfully.")
    except Exception as e:
        print(f"WARNING: Failed to initialize fake_useragent: {e}. Using static User-Agent. Traceback: {sys.exc_info()}")
        # If fake_useragent fails, we proceed with ua = None, so the get_soup fallback is used.
        ua = None

    # Always attempt to set up http_session
    retry_strategy = Retry(
        total=5,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    http_session = requests.Session()
    http_session.mount("https://", adapter)
    http_session.mount("http://", adapter)
    print("DEBUG: requests.Session initialized successfully.")

except Exception as e:
    # This block now only catches issues with requests.Session setup itself
    print(f"CRITICAL ERROR: Failed to set up requests.Session: {e}. Network operations may fail. Traceback: {sys.exc_info()}")
    http_session = None # Set to None only if session setup fails entirely


# --- Helper: Log to GUI and Console ---
class TextRedirector(object):
    """A class to redirect stdout to a Tkinter Text widget via a queue."""
    def __init__(self, queue):
        self.queue = queue
        self.stdout = sys.stdout # Keep a reference to original stdout

    def write(self, str_val):
        self.queue.put({'type': 'log', 'message': str_val})
        # REMOVED: self.stdout.write(str_val)
        # This line was causing AttributeError: 'NoneType' object has no attribute 'write'
        # in compiled --windowed applications, as sys.stdout is detached.

    def flush(self):
        # The flush for the original stdout is no longer necessary if we don't write to it.
        # But keeping it might prevent other unexpected issues if some part of Python
        # still tries to flush a detached stdout.
        if self.stdout is not None:
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
        
        # Use a more generic but common User-Agent if ua is None (fake_useragent failed)
        # This specific User-Agent string is known to work for many sites.
        user_agent_string = ua.random if ua else 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
        headers = {'User-Agent': user_agent_string}
        
        # Ensure current_session is not None before making the request
        if current_session is None:
            log_to_gui_and_console(f"Network error: http_session is None. Cannot fetch {url}. Session setup failed.")
            return None

        response = current_session.get(url, timeout=60, headers=headers) # Increased timeout to 60s
        response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)
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
                
                name_pattern = r'(?:Name\s*[:;]\s*|Vessel Name\s*[:;]\s*)([^\n,仿佛;]{5,100}?)\s*vessel' # Updated regex
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

# MODIFIED fetch_uani_vessels to prioritize bundled CSV
def fetch_uani_vessels(pd_module): # Added pd_module parameter
    log_to_gui_and_console(f"\n--- Fetching UANI data ---")
    
    # Try to load from bundled CSV first
    try:
        # Determine the base path correctly for frozen (compiled) vs. unfrozen (script)
        base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
        bundled_csv_full_path = os.path.join(base_path, UANI_BUNDLED_CSV_NAME)
        
        if os.path.exists(bundled_csv_full_path):
            log_to_gui_and_console(f"UANI: Attempting to load from bundled CSV: {bundled_csv_full_path}")
            df = pd_module.read_csv(bundled_csv_full_path)
            
            # Normalize column names in the loaded DataFrame
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
            
            if imo_col_key and name_col_key:
                df = df[[name_col_key, imo_col_key]].copy()
                df.rename(columns={name_col_key: "Vessel Name", imo_col_key: "IMO Number"}, inplace=True)
                df["Source"] = "UANI (Bundled CSV)"
                df.dropna(subset=["IMO Number"], inplace=True)
                df["IMO Number"] = df["IMO Number"].astype(str).apply(lambda x: re.sub(r'\D', '', str(x)))
                df = df[df['IMO Number'].str.match(r'^\d{7}$')]
                log_to_gui_and_console(f"UANI: Successfully loaded {len(df)} vessels from bundled CSV.")
                return df
            else:
                log_to_gui_and_console(f"UANI: Bundled CSV headers missing 'IMO' or 'Vessel Name'. Headers: {original_columns}. Attempting web scrape fallback.")

    except Exception as e:
        log_to_gui_and_console(f"UANI: Error loading bundled CSV ({UANI_BUNDLED_CSV_NAME}): {e}. Falling back to web scrape. Traceback: {sys.exc_info()}")

    # Fallback to web scraping if bundled CSV fails or doesn't exist/is malformed
    log_to_gui_and_console(f"UANI: Attempting web scrape from {UANI_WEBSCRAPE_URL}...")
    try:
        soup = get_soup(UANI_WEBSCRAPE_URL)
        if not soup:
            log_to_gui_and_console("UANI: Failed to get soup from UANI URL (web scrape fallback).")
            return pd_module.DataFrame()

        log_to_gui_and_console("UANI: HTML content downloaded. Searching for tables...")
        
        tables = soup.find_all('table')
        vessels_data = []
        
        for table_idx, table in enumerate(tables):
            headers = [th.get_text(strip=True) for th in table.find_all('th')]
            
            imo_col_idx = -1
            name_col_idx = -1
            
            normalized_headers = [h.lower().replace('*', '').strip() for h in headers]
            
            for idx, header in enumerate(normalized_headers):
                if 'imo' in header and imo_col_idx == -1:
                    imo_col_idx = idx
                if ('vessel' in header or 'name' in header) and name_col_idx == -1:
                    name_col_idx = idx
            
            if imo_col_idx != -1 and name_col_idx != -1:
                log_to_gui_and_console(f"UANI: Found relevant table (Table {table_idx+1}) with IMO (col {imo_col_idx}) and Name (col {name_col_idx}) columns.")
                rows = table.find_all('tr')
                
                for r_idx, row in enumerate(rows):
                    if r_idx == 0:
                        continue
                        
                    cols = row.find_all(['td', 'th'])
                    cols = [ele.text.strip() for ele in cols]
                    
                    if len(cols) > max(imo_col_idx, name_col_idx):
                        vessel_name = cols[name_col_idx]
                        imo_number = cols[imo_col_idx]
                        
                        imo_number = re.sub(r'\D', '', str(imo_number)).strip()
                        
                        if re.fullmatch(r'^\d{7}$', imo_number):
                            vessels_data.append({"Vessel Name": vessel_name, "IMO Number": imo_number, "Source": "UANI (Web Scrape)"})
                
                log_to_gui_and_console(f"UANI: Processed {len(vessels_data)} entries from Table {table_idx+1}.")
                break
            else:
                log_to_gui_and_console(f"UANI: Table {table_idx+1} does not contain both 'IMO' and 'Vessel Name' columns. Headers: {headers}")

        df = pd_module.DataFrame(vessels_data)
        
        if not df.empty:
            df.drop_duplicates(subset=['IMO Number'], keep='first', inplace=True)
            
        log_to_gui_and_console(f"UANI: Final {len(df)} vessels from web scrape.")
        return df
        
    except requests.exceptions.RequestException as e:
        log_to_gui_and_console(f"Network error fetching UANI HTML data (web scrape fallback): {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame()
    except Exception as e:
        log_to_gui_and_console(f"UANI: Error processing UANI HTML data (web scrape fallback): {e}. Traceback: {sys.exc_info()}")
        return pd_module.DataFrame()

# --- Worker function for Sanctions Report ---
def process_all_data(on_complete_callback=None, display_callback=None): # Removed output_filepath
    # Total steps for the main progress bar:
    # 1. OFAC Fetch
    # 2. UK Fetch (get page)
    # 3. UK Fetch (ODT download)
    # 4. UK Fetch (ODT parse)
    # 5. DMA Fetch
    # 6. UANI Fetch
    total_steps = 6 # Reduced total steps as Excel write is removed
    current_step = 0
    gui_update_queue.put({'type': 'progress_config', 'max': total_steps})
    
    # Helper to increment progress and update GUI
    def increment_progress():
        nonlocal current_step
        current_step += 1
        gui_update_queue.put({'type': 'progress_update', 'value': current_step, 'max': total_steps})
        log_to_gui_and_console(f"Overall Progress: {current_step}/{total_steps}") # Diagnostic log

    log_to_gui_and_console("Starting sanctions report generation...")
    
    fetched_data = {} # Dictionary to store fetched DataFrames

    try:
        # Import pandas inside this function to ensure it's available in the thread
        import pandas as pd_local 

        # Fetch OFAC data (pass pd_local explicitly)
        df_ofac = fetch_ofac_vessels(pd_local)
        fetched_data["OFAC_Vessels"] = df_ofac
        increment_progress() 
        
        # Fetch UK sanctions data (pass pd_local explicitly)
        df_uk = fetch_uk_sanctions_vessels(progress_callback=increment_progress, pd_module=pd_local)
        fetched_data["UK_Sanctions_Vessels"] = df_uk
        
        # Fetch DMA data (pass pd_local explicitly)
        df_dma = fetch_dma_vessels(pd_local)
        fetched_data["EU_DMA_Vessels"] = df_dma
        increment_progress()
        
        # Fetch UANI data (pass pd_local explicitly)
        df_uani = fetch_uani_vessels(pd_local)
        fetched_data["UANI_Vessels_Tracked"] = df_uani
        increment_progress()
        
        log_to_gui_and_console("\nDisplaying data in Sanctions Report Viewer tab...")

        # Update the global data store
        global global_sanctions_data_store
        global_sanctions_data_store.update(fetched_data)
        
        # Call the display callback to show data in the new tab
        if display_callback:
            display_callback(global_sanctions_data_store)
        
        gui_update_queue.put({'type': 'show_message', 'kind': 'info', 'title': 'Success',
                              'message': "Sanctions report data fetched and displayed successfully in 'Sanctions Report Viewer' tab."})
        
        if on_complete_callback:
            # Pass the fact that data is loaded in-memory, no file path needed
            on_complete_callback(True) 

    except Exception as e:
        log_to_gui_and_console(f"Error during main report processing: {e}. Traceback: {sys.exc_info()}")
        gui_update_queue.put({'type': 'show_message', 'kind': 'error', 'title': 'Report Generation Error',
                              'message': f"An error occurred during report generation: {e}"})
        if on_complete_callback:
            on_complete_callback(False) # Signal failure
            
    # Final increment for data display completion
    increment_progress() 
    gui_update_queue.put({'type': 'processing_done'})
    log_to_gui_and_console("\nScript finished.")


# --- SanctionsAppFrame (Program 1 GUI encapsulated) ---
class SanctionsAppFrame(ttk.Frame):
    def __init__(self, parent, sanctions_display_callback): # Added display callback
        super().__init__(parent)
        self.is_processing = False
        self.sanctions_display_callback = sanctions_display_callback # Callback to update SanctionsDisplayFrame
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

        # Corrected: Use main_content_frame here for log_frame (Fixes NameError)
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
        """No longer used directly as ExcelComparator is removed. Kept for compatibility if external references exist."""
        pass # This callback is now effectively a no-op or can be removed if not called elsewhere.

    def start_processing_thread(self, on_complete_override=None):
        """
        Starts the report generation thread.
        Now calls process_all_data without output_filepath and passes display_callback.
        """
        # No file dialog needed as output is in-app
        
        self.is_processing = True
        self.run_button.config(state=tk.DISABLED)
        
        # Initialize progress_bar BEFORE animate_pulse and ensure it has a size
        self.progress_canvas.delete("progress_background") # Clear old background
        self.progress_canvas.delete("progress_bar") # Clear any old bar
        self.progress_canvas.create_rectangle(0, 0, self.progress_canvas.winfo_width(), 25, fill=self.FRAME_COLOR, outline="", tags="progress_background")
        self.progress_canvas.create_rectangle(0, 0, 0, 25, fill=self.ACCENT_COLOR, outline="", tags="progress_bar")
        self.master.update_idletasks() # Force update to ensure 'progress_bar' exists before animation starts
        time.sleep(0.01) # Small delay to help ensure canvas update

        self.update_progress(0, 7) # Initialize visual progress
        self.log_area.config(state='normal'); self.log_area.delete('1.0', tk.END); self.log_area.config(state='disabled')
        
        # Pass the sanctions_display_callback to process_all_data
        threading.Thread(target=process_all_data, 
                         args=(on_complete_override, self.sanctions_display_callback,), 
                         daemon=True).start()
        self.animate_pulse()

    def start_automatic_generation(self, callback_after_generation):
        """Public method to trigger a report generation automatically."""
        self.automatic_generation_complete_callback = callback_after_generation
        # No dummy filepath needed anymore.
        self.start_processing_thread(on_complete_override=self._automatic_generation_complete)

    def _automatic_generation_complete(self, success_status): # Changed parameter to success_status
        """Internal callback for automatic generation completion."""
        if self.automatic_generation_complete_callback:
            # Pass success status, not a file path
            self.automatic_generation_complete_callback(success_status) 
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
        self.progress_canvas.tag_raise("pulse", "progress_bar") # Ensure pulse is on top of bar
        
        new_x = x_pos + 5 * direction
        if new_x > canvas_width or new_x < -pulse_width:
              new_x = -pulse_width if direction == 1 else canvas_width
        
        self.after(15, lambda: self.animate_pulse(new_x, direction))

    def update_log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message)
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


# --- SanctionsDisplayFrame (NEW TAB for displaying fetched data) ---
class SanctionsDisplayFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.current_data = {} # Stores the DataFrames to display
        self.notebook_tabs = [] # To keep track of dynamically created tabs
        self.current_selected_data_source = tk.StringVar(value="All Data (Combined)") # Default selection

        self.BG_COLOR = "#F0F0F0"
        self.TEXT_COLOR = "#333333"
        self.FRAME_BG_COLOR = "#FFFFFF"
        self.BUTTON_COLOR = "#0056B3"
        self.BUTTON_TEXT_COLOR = "white"

        self.style = ttk.Style(self)
        self.style.configure('TFrame', background=self.BG_COLOR)
        self.style.configure('TLabelframe', background=self.FRAME_BG_COLOR, bordercolor="#AAAAAA", relief="groove", borderwidth=1)
        self.style.configure('TLabelframe.Label', background=self.FRAME_BG_COLOR, foreground="#555555", font=("Arial", 11, "bold"))
        self.style.configure("Treeview",
                             background="#F8F8F8",
                             foreground=self.TEXT_COLOR,
                             rowheight=25,
                             fieldbackground="#F8F8F8",
                             font=("Arial", 10))
        self.style.configure("Treeview.Heading",
                             font=("Arial", 10, "bold"),
                             background="#DDDDDD",
                             foreground="#333333",
                             relief="raised")
        self.style.configure('Display.TButton', background=self.BUTTON_COLOR, foreground=self.BUTTON_TEXT_COLOR,
                             font=("Arial", 10, "bold"), relief="raised", borderwidth=2)
        self.style.map('Display.TButton', background=[('active', '#004085')])

        main_frame = ttk.Frame(self, padding="20", style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)

        controls_frame = ttk.LabelFrame(main_frame, text="Report Viewer Controls", padding=(15,10))
        controls_frame.pack(pady=10, fill=tk.X)

        ttk.Label(controls_frame, text="Select Data Source:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.source_selector = ttk.Combobox(controls_frame, textvariable=self.current_selected_data_source, state="readonly", width=30)
        self.source_selector.grid(row=0, column=1, padx=5, pady=5)
        self.source_selector.bind("<<ComboboxSelected>>", self.display_selected_source)
        self.source_selector['values'] = ["No Data Available"] # Default value

        # Define methods that will be assigned to buttons BEFORE the buttons are created
        self.export_current_view = self._export_current_view_method
        self.clear_display = self._clear_display_method
        self.load_uani_data_from_file = self._load_uani_data_from_file_method # New: UANI manual load method

        self.export_button = ttk.Button(controls_frame, text="Export Current View to Excel", command=self.export_current_view, style='Display.TButton')
        self.export_button.grid(row=0, column=2, padx=10, pady=5)
        self.export_button.config(state=tk.DISABLED) # Disable until data is loaded

        self.clear_button = ttk.Button(controls_frame, text="Clear All Displayed Data", command=self.clear_display, style='Display.TButton')
        self.clear_button.grid(row=0, column=3, padx=10, pady=5)
        
        # New: UANI Manual Load Button
        self.load_uani_button = ttk.Button(controls_frame, text="Load UANI Data from File", command=self.load_uani_data_from_file, style='Display.TButton')
        self.load_uani_button.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")


        self.tree_frame = ttk.Frame(main_frame, style='TFrame')
        self.tree_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Treeview setup for displaying data
        self.treeview = ttk.Treeview(self.tree_frame)
        self.treeview.pack(side="left", fill="both", expand=True)
        
        tree_scroll_y = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.treeview.yview)
        tree_scroll_y.pack(side="right", fill="y")
        self.treeview.configure(yscrollcommand=tree_scroll_y.set)

        tree_scroll_x = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.treeview.xview)
        tree_scroll_x.pack(side="bottom", fill="x")
        self.treeview.configure(xscrollcommand=tree_scroll_x.set)
        
        ttk.Label(main_frame, text="Status:", background=self.BG_COLOR, foreground=self.TEXT_COLOR).pack(anchor="w", padx=10)
        self.status_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=3, font=("Consolas", 9),
                                                     bg=self.FRAME_BG_COLOR, fg=self.TEXT_COLOR, relief=tk.FLAT)
        self.status_text.pack(padx=10, pady=5, fill=tk.X, expand=False)
        self.status_text.insert(tk.END, "Awaiting sanctions report data...\n")
        self.status_text.config(state=tk.DISABLED)

        self.after(100, self.process_sanctions_display_queue)


    def update_status_text(self, message):
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message)
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')

    def display_data(self, fetched_data_dict):
        self.current_data = fetched_data_dict
        available_sources = ["All Data (Combined)"] + [key for key, df in fetched_data_dict.items() if df is not None and not df.empty]
        self.source_selector['values'] = available_sources
        self.current_selected_data_source.set("All Data (Combined)")
        
        self.export_button.config(state=tk.NORMAL)
        self.display_selected_source() # Display the combined data by default
        self.update_status_text(f"New sanctions report data loaded. {len(available_sources)-1} sources available.\n")

    def display_selected_source(self, event=None):
        selected_source = self.current_selected_data_source.get()
        self.clear_treeview()

        df_to_display = pd.DataFrame()
        
        if selected_source == "All Data (Combined)":
            all_dfs = []
            for source_name, df in self.current_data.items():
                if df is not None and not df.empty:
                    df_copy = df.copy()
                    if 'Source' not in df_copy.columns:
                        df_copy['Source'] = source_name # Add source column if not present
                    all_dfs.append(df_copy)
            if all_dfs:
                df_to_display = pd.concat(all_dfs, ignore_index=True)
                # Ensure 'IMO Number' is string for robust duplicate dropping
                if 'IMO Number' in df_to_display.columns:
                    df_to_display['IMO Number'] = df_to_display['IMO Number'].astype(str)
                    df_to_display.drop_duplicates(subset=['IMO Number'], keep='first', inplace=True)
            self.update_status_text(f"Displaying combined data: {len(df_to_display)} unique vessels.\n")
        elif selected_source in self.current_data and self.current_data[selected_source] is not None:
            df_to_display = self.current_data[selected_source].copy()
            self.update_status_text(f"Displaying {selected_source}: {len(df_to_display)} records.\n")
        else:
            self.update_status_text(f"No data available for {selected_source}.\n")
            return

        if not df_to_display.empty:
            self.treeview["columns"] = list(df_to_display.columns)
            self.treeview["show"] = "headings"

            for col in df_to_display.columns:
                self.treeview.heading(col, text=col)
                # Adjust column width based on content
                max_len = max(df_to_display[col].astype(str).map(len).max(), len(col))
                self.treeview.column(col, width=max_len * 9 + 10) # Approx width based on characters
            
            for index, row in df_to_display.iterrows():
                self.treeview.insert("", "end", values=list(row.values))
        else:
            self.update_status_text("Selected data source is empty.\n")
            self.clear_treeview()


    def clear_treeview(self):
        self.treeview.delete(*self.treeview.get_children())
        self.treeview["columns"] = () # Clear columns
        self.treeview["show"] = "tree headings" # Reset to default if no columns
        self.update_status_text("Display cleared.\n")

    def _clear_display_method(self): # Actual method for clear_display
        self.current_data = {}
        self.clear_treeview()
        self.source_selector['values'] = ["No Data Available"]
        self.current_selected_data_source.set("No Data Available")
        self.export_button.config(state=tk.DISABLED)
        self.update_status_text("All displayed sanctions data cleared.\n")
        global global_sanctions_data_store
        for key in global_sanctions_data_store:
            global_sanctions_data_store[key] = None

    def _export_current_view_method(self): # Actual method for export_current_view
        selected_source = self.current_selected_data_source.get()
        df_to_export = pd.DataFrame()

        if selected_source == "All Data (Combined)":
            all_dfs = []
            for source_name, df in self.current_data.items():
                if df is not None and not df.empty:
                    df_copy = df.copy()
                    if 'Source' not in df_copy.columns:
                        df_copy['Source'] = source_name
                    all_dfs.append(df_copy)
            if all_dfs:
                df_to_export = pd.concat(all_dfs, ignore_index=True)
                if 'IMO Number' in df_to_export.columns:
                    df_to_export['IMO Number'] = df_to_export['IMO Number'].astype(str)
                    df_to_export.drop_duplicates(subset=['IMO Number'], keep='first', inplace=True)
        elif selected_source in self.current_data and self.current_data[selected_source] is not None:
            df_to_export = self.current_data[selected_source].copy()
        
        if df_to_export.empty:
            messagebox.showwarning("Export Warning", "No data to export in the current view.")
            return

        timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
        initial_filename = f"Sanctions_Report_View_{selected_source.replace(' ', '_').replace('(', '').replace(')', '')}_{timestamp}.xlsx"

        output_filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Export Current View As", initialfile=initial_filename)
        
        if output_filepath:
            try:
                # Import pandas inside this function to ensure it's available in the thread
                import pandas as pd_local 
                df_to_export.to_excel(output_filepath, index=False)
                messagebox.showinfo("Export Successful", f"Current view exported to:\n{output_filepath}")
                self.update_status_text(f"Current view exported to: {output_filepath}\n")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export data: {e}")
                self.update_status_text(f"Error exporting data: {e}\n")
        else:
            self.update_status_text("Export cancelled.\n")

    def _load_uani_data_from_file_method(self):
        """Allows user to manually load a CSV into the UANI sanctions data, persisting it."""
        filepath = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")]
        )
        if not filepath:
            self.update_status_text("UANI CSV load cancelled by user.\n")
            return

        self.update_status_text(f"Attempting to load UANI data from user-selected file: {filepath}\n")
        
        try:
            import pandas as pd_local # Ensure pandas is available in this scope
            df_loaded = pd_local.read_csv(filepath)

            # Normalize column names in the loaded DataFrame
            original_columns = list(df_loaded.columns)
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
            
            if imo_col_key and name_col_key:
                df_uani = df_loaded[[name_col_key, imo_col_key]].copy()
                df_uani.rename(columns={name_col_key: "Vessel Name", imo_col_key: "IMO Number"}, inplace=True)
                df_uani["Source"] = "UANI (Manual File)"
                df_uani.dropna(subset=["IMO Number"], inplace=True)
                df_uani["IMO Number"] = df_uani["IMO Number"].astype(str).apply(lambda x: re.sub(r'\D', '', str(x)))
                df_uani = df_uani[df_uani['IMO Number'].str.match(r'^\d{7}$')]
                
                # Update global store
                global global_sanctions_data_store
                global_sanctions_data_store["UANI_Vessels_Tracked"] = df_uani
                
                # Save to the bundled CSV location for persistence
                try:
                    # Get the directory where the main executable/script is located
                    app_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
                    persist_path = os.path.join(app_dir, UANI_BUNDLED_CSV_NAME)
                    df_uani.to_csv(persist_path, index=False, encoding='utf-8')
                    self.update_status_text(f"UANI data successfully loaded from file and saved to {UANI_BUNDLED_CSV_NAME} for persistence.\n")
                except Exception as save_e:
                    self.update_status_text(f"WARNING: Could not save UANI data to {UANI_BUNDLED_CSV_NAME} for persistence: {save_e}. Data loaded but not saved for next session.\n")

                self.display_data(global_sanctions_data_store) # Refresh display
                messagebox.showinfo("UANI Load Success", f"UANI data loaded and displayed from: {filepath}")
            else:
                self.update_status_text(f"ERROR: Selected UANI CSV '{filepath}' missing 'IMO' or 'Vessel Name' columns after cleaning. Found headers: {df_loaded.columns.tolist()}\n")
                messagebox.showerror("UANI Load Error", "Selected UANI CSV file format is incorrect. Missing 'IMO' or 'Vessel Name' columns.")

        except Exception as e:
            self.update_status_text(f"ERROR: Failed to load UANI data from file {filepath}: {e}. Traceback: {sys.exc_info()}\n")
            messagebox.showerror("UANI Load Error", f"An error occurred while loading UANI data from file: {e}")

    def process_sanctions_display_queue(self):
        try:
            while not sanctions_display_queue.empty():
                item = sanctions_display_queue.get_nowait()
                msg_type = item.get('type')
                if msg_type == 'display_data':
                    self.display_data(item['data'])
                elif msg_type == 'update_status':
                    self.update_status_text(item['message'])
        finally:
            self.after(100, self.process_sanctions_display_queue)


# --- MyVesselsAppFrame (New Tab) ---
MY_VESSELS_CSV = "my_vessels.csv"

class MyVesselsAppFrame(ttk.Frame):
    def __init__(self, parent, get_last_sanctions_data_callback): # Renamed callback
        super().__init__(parent)
        self.get_last_sanctions_data = get_last_sanctions_data_callback # Callback to get in-memory data
        self.vessels_data = [] # List of {'name': 'Vessel Name', 'imo': 'IMO Number'} dicts
        
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
        tree_scroll.pack(side="right", fill="y")
        self.vessel_tree.configure(yscrollcommand=tree_scroll.set)
        self.vessel_tree.pack(fill="both", expand=True)

        remove_button = ttk.Button(list_frame, text="Remove Selected Vessel(s)", command=self.remove_selected_vessels, style='MyVessels.TButton')
        remove_button.pack(pady=5)

        # Export My Vessels to Excel button (now standalone)
        export_button = ttk.Button(main_frame, text="Export My Vessels to Excel", command=self.export_my_vessels, style='MyVessels.TButton')
        export_button.pack(pady=10, fill=tk.X)

        # --- MOVED THIS BLOCK UP (within MyVesselsAppFrame init) ---
        ttk.Label(main_frame, text="Status:", background=self.BG_COLOR, foreground=self.TEXT_COLOR).pack(anchor="w", padx=10)
        self.status_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=5, font=("Consolas", 9),
                                                     bg=self.FRAME_BG_COLOR, fg=self.TEXT_COLOR, relief=tk.FLAT)
        self.status_text.pack(padx=10, pady=5, fill=tk.X, expand=False)
        self.status_text.insert(tk.END, "Manage your personal vessel list here.\n")
        self.status_text.config(state=tk.DISABLED)
        # --- END MOVED BLOCK ---

        # New "Load CSV" button
        load_csv_button = ttk.Button(main_frame, text="Load Vessels from CSV", command=self.prompt_load_custom_vessels_csv, style='MyVessels.TButton')
        load_csv_button.pack(pady=10, fill=tk.X)

        self.load_vessels_from_csv() # Initial population - NOW status_text exists
        self.after(100, self.process_my_vessels_queue) # Start queue processing

    def _load_vessels_from_file(self, file_path):
        """Helper to load vessels from a specified CSV file, used internally."""
        temp_vessels_data = [] # Use a temporary list for loading
        self.update_status_text(f"Attempting to load vessels from: {file_path}\n")

        if not os.path.exists(file_path):
            self.update_status_text(f"File not found: {file_path}. Please ensure it exists and is accessible.\n")
            return False # Indicate failure

        try:
            with open(file_path, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                
                if reader.fieldnames:
                    # Log raw and stripped fieldnames for debugging
                    self.update_status_text(f"CSV Raw Fieldnames: {reader.fieldnames}\n")
                    reader.fieldnames = [field.strip() for field in reader.fieldnames]
                    self.update_status_text(f"CSV Stripped Fieldnames: {reader.fieldnames}\n")

                if reader.fieldnames is None:
                    self.update_status_text("CSV file has no headers. Expected 'name' and 'imo'.\n")
                    return False
                
                if 'name' not in reader.fieldnames or 'imo' not in reader.fieldnames:
                    self.update_status_text(f"CSV headers missing 'name' or 'imo'. Found: {reader.fieldnames}\n")
                    return False

                read_count = 0
                for row_idx, row in enumerate(reader):
                    # Create a cleaned_row dictionary with stripped keys for reliable access
                    cleaned_row = {k.strip(): v for k, v in row.items()}
                    
                    if 'name' in cleaned_row and 'imo' in cleaned_row:
                        temp_vessels_data.append({'name': cleaned_row['name'], 'imo': cleaned_row['imo']})
                        read_count += 1
                        if read_count <= 5: # Log first five rows for debugging
                            self.update_status_text(f"Read row {row_idx+1}: Name='{cleaned_row['name']}', IMO='{cleaned_row['imo']}'\n")
                    else:
                        self.update_status_text(f"Skipping row {row_idx+1} due to missing 'name' or 'imo' fields: {row}\n")

            # If loading successful, update the main vessels_data list
            self.vessels_data = temp_vessels_data
            self.update_status_text(f"Successfully loaded {len(self.vessels_data)} vessels from {file_path}.\n")
            self.populate_vessel_tree(search_term=self.search_vessel_entry.get()) # Refresh display
            self.save_vessels_to_csv() # Save the newly loaded list
            return True

        except Exception as e:
            self.update_status_text(f"Error loading vessels from CSV: {e}. Please check the file's integrity and permissions.\n")
            print(f"DEBUG: Exception in _load_vessels_from_file: {e}")
            import traceback
            traceback.print_exc()
            return False

    def load_vessels_from_csv(self):
        """Loads vessels from the default MY_VESSELS_CSV file on startup."""
        self._load_vessels_from_file(MY_VESSELS_CSV)

    def prompt_load_custom_vessels_csv(self):
        """Opens a file dialog for the user to select a CSV to load."""
        filepath = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")]
        )
        if filepath:
            self.update_status_text(f"User selected: {filepath}\n")
            if self._load_vessels_from_file(filepath):
                messagebox.showinfo("Load Successful", f"Vessels loaded successfully from:\n{filepath}")
            else:
                messagebox.showerror("Load Failed", f"Failed to load vessels from:\n{filepath}\nCheck status messages for details.")
        else:
            self.update_status_text("CSV file load cancelled by user.\n")

    def save_vessels_to_csv(self):
        try:
            with open(MY_VESSELS_CSV, mode='w', newline='', encoding='utf-8') as file:
                fieldnames = ['name', 'imo']
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(self.vessels_data)
            self.update_status_text(f"Saved {len(self.vessels_data)} vessels to {MY_VESSELS_CSV}.\n")
            print(f"DEBUG: Successfully saved {len(self.vessels_data)} vessels to {MY_VESSELS_CSV}") # Console log
        except Exception as e:
            error_message = f"Error saving vessels to CSV: {e}"
            self.update_status_text(f"ERROR: {error_message}. Please check file permissions or if another process is using 'my_vessels.csv'.\n")
            print(f"DEBUG: Error during save_vessels_to_csv: {e}") # Console log
            import traceback
            traceback.print_exc() # Print full traceback to console

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
        
        if output_filepath:
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
        else:
            self.update_status_text("Export cancelled.\n")

    def compare_my_vessels_threaded(self):
        """
        Initiates the sanctions check for My Vessels against the in-memory report.
        """
        if not self.vessels_data:
            messagebox.showinfo("No Vessels", "Please add vessels to 'My Vessels' list before checking sanctions.")
            return

        # Get the latest fetched sanctions data from the main application
        sanctions_data_dict = self.get_last_sanctions_data()
        
        # Check if any sanctions data was actually fetched
        has_sanctions_data = any(df is not None and not df.empty for df in sanctions_data_dict.values())

        if not has_sanctions_data:
            messagebox.showwarning("Sanctions Report Missing",
                                   "No sanctions report data found in memory. Please generate a report in the 'Sanctions Report Generator' tab first.")
            return

        self.update_status_text("Starting sanctions check for My Vessels...")
        self.is_checking_sanctions = True
        self.check_sanctions_button.config(state=tk.DISABLED) # Disable button during check
        
        # Clear previous sanction status before running new check
        self.sanctioned_imos_from_report.clear()
        self.sanctioned_vessel_sources.clear()
        self.populate_vessel_tree(search_term=self.search_vessel_entry.get()) # Clear old highlights
        
        # Reset and start progress animation
        self.check_progress_canvas.delete("check_progress_background")
        self.check_progress_canvas.delete("check_progress_bar")
        self.check_progress_canvas.create_rectangle(0, 0, self.check_progress_canvas.winfo_width(), 15, fill=self.FRAME_BG_COLOR, outline="", tags="check_progress_background")
        self.check_progress_canvas.create_rectangle(0, 0, 0, 15, fill=self.BUTTON_COLOR, outline="", tags="check_progress_bar")
        self.master.update_idletasks() # Force update
        time.sleep(0.01) # Small delay to help ensure canvas update
        
        self.update_check_progress(0, 100) # Progress in percentage (0-100)
        self.animate_check_pulse()

        threading.Thread(target=self.run_vessel_sanctions_check,
                         args=(list(self.vessels_data), sanctions_data_dict), # Pass the in-memory data
                         daemon=True).start()

    def run_vessel_sanctions_check(self, my_vessels, sanctions_data_dict):
        """Worker function to perform the actual sanctions comparison."""
        try:
            # Diagnostic print for my_vessels content BEFORE processing
            print(f"DEBUG: run_vessel_sanctions_check received my_vessels (len={len(my_vessels)}):")
            for i, item in enumerate(my_vessels[:5]): # Print first 5 elements for inspection
                print(f"DEBUG: my_vessels[{i}] type: {type(item)}, value: {item}")
            if len(my_vessels) > 5:
                print("DEBUG: ... (more vessels omitted)")

            # Import pandas inside this function to ensure it's available in the thread
            import pandas as pd_local 

            all_sanctioned_imos_temp = set()
            sanctioned_imo_details_temp = {}

            # Phase 1: Populate all_sanctioned_imos_temp from all loaded sanctions sheets
            total_sanctioned_sheets_rows = sum(len(df) for df in sanctions_data_dict.values() if df is not None and not df.empty)
            processed_sanction_rows = 0

            my_vessels_log_queue.put({'type': 'log', 'message': f"MV Check: Loading {len(sanctions_data_dict)} sanctions lists from memory..."})

            for source_name, df in sanctions_data_dict.items():
                if df is None or df.empty:
                    my_vessels_log_queue.put({'type': 'log', 'message': f"MV Check: Skipping empty or None sheet: {source_name}"})
                    continue
                
                imo_col_for_sheet = None
                for col in df.columns:
                    cleaned_col_name = str(col).strip().lower()
                    # Use the globally defined IMO_LIKE_COLUMN_PATTERNS
                    if any(pattern in cleaned_col_name for pattern in IMO_LIKE_COLUMN_PATTERNS):
                        imo_col_for_sheet = col # Found the original column name
                        break

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
                # Added check to ensure 'vessel' is a dictionary
                if not isinstance(vessel, dict):
                    my_vessels_log_queue.put({'type': 'log', 'message': f"ERROR: Expected dictionary for vessel, but got type {type(vessel)} with value '{vessel}'. Skipping."})
                    continue # Skip this malformed entry
                
                # Check for 'imo' key specifically within the dictionary
                if 'imo' not in vessel:
                    my_vessels_log_queue.put({'type': 'log', 'message': f"WARNING: Vessel entry missing 'imo' key: {vessel}. Skipping."})
                    continue

                vessel_imo_cleaned = re.sub(r'\D', '', str(vessel['imo'])).strip() 
                
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
            "   • The application will then connect to various online sources (OFAC, UK, EU DMA, UANI Website) "
            "to fetch the latest vessel sanctions data.\n"
            "   • Progress and status messages will be displayed in the 'Log Output' area.\n"
            "   • Once complete, the fetched data will be displayed in the 'Sanctions Report Viewer' tab."
        ), style='About.TLabel', wraplength=700, justify=tk.LEFT).pack(anchor="w", pady=(0, 5))

        ttk.Label(main_frame, text="2. Sanctions Report Viewer Tab:", style='About.SubTitle.TLabel').pack(anchor="w", pady=(5, 0))
        ttk.Label(main_frame, text=(
            "   • This tab automatically displays the data fetched by the 'Sanctions Report Generator'.\n"
            "   • Use the 'Select Data Source' dropdown to view individual lists (OFAC, UK, EU DMA, UANI) or a combined list.\n"
            "   • Click 'Export Current View to Excel' to save the currently displayed table to an Excel file.\n"
            "   • Click 'Clear All Displayed Data' to remove all fetched data from memory and the display."
        ), style='About.TLabel', wraplength=700, justify=tk.LEFT).pack(anchor="w", pady=(0, 5))

        ttk.Label(main_frame, text="3. My Vessels Tab:", style='About.SubTitle.TLabel').pack(anchor="w", pady=(5, 0))
        ttk.Label(main_frame, text=(
            "   • Add your own vessel names and IMO numbers using the 'Add New Vessel' section. "
            "IMO numbers must be exactly 7 digits.\n"
            "   • Your list of vessels will be displayed in the table and automatically saved for future sessions.\n"
            "   • Select one or more vessels from the list and 'Remove Selected Vessel(s)' to delete them.\n"
            "   • 'Export My Vessels to Excel': Saves your current list of vessels to a new Excel file.\n"
            "   • 'Check Sanctions': Compares your vessels against the most recently fetched sanctions data "
            "from the 'Sanctions Report Generator' tab. This will indicate if your vessels are found in any "
            "sanctions list and specify the source(s)."
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
        self.last_sanctions_data = global_sanctions_data_store # Direct reference to the global store
        self.my_vessels_tab_needs_sanction_check = False 

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

        # Instantiate tab frames
        self.sanctions_display_tab = SanctionsDisplayFrame(self.notebook) # New display tab first
        self.sanctions_tab = SanctionsAppFrame(self.notebook, self.sanctions_display_tab.display_data) # Pass display callback
        self.my_vessels_tab = MyVesselsAppFrame(self.notebook, self.get_current_sanctions_data) # Pass getter for in-memory data
        self.about_tab = AboutAppFrame(self.notebook)

        # Add tabs to notebook in desired order
        self.notebook.add(self.sanctions_tab, text="Sanctions Report Generator")
        self.notebook.add(self.sanctions_display_tab, text="Sanctions Report Viewer") # New tab added
        self.notebook.add(self.my_vessels_tab, text="My Vessels")
        self.notebook.add(self.about_tab, text="About This Application")

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def get_current_sanctions_data(self):
        """Provides the current in-memory sanctions data to other components."""
        return global_sanctions_data_store

    def on_sanctions_report_complete(self, success_status):
        """
        Callback triggered by SanctionsAppFrame when a report generation is complete.
        This now just logs success and updates UI, as display is handled directly.
        """
        if success_status:
            print("Sanctions report generation completed successfully (data in memory).")
        else:
            print("Sanctions report generation failed.")
        # No more file path to pass here.

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
    try: import pandas 
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
