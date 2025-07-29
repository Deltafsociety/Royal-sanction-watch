# ROYAL SANCTION WATCH

# Maritime Sanctions & Data Tool

This application is a comprehensive desktop tool designed to assist maritime professionals with compliance, risk assessment, and data management related to vessel sanctions. It automates the collection of sanctions data from various international sources and provides robust utilities for comparing Excel files and managing a personal vessel list.

## Features

* **Sanctions Report Generation:**
    * Automatically fetches the latest vessel sanctions data from:
        * OFAC (Office of Foreign Assets Control) SDN list.
        * UK Sanctions List (via ODT document).
        * EU (DMA) Designated Vessels List.
        * UANI (United Against Nuclear Iran) Google Sheet for tracked vessels.
    * Consolidates all data into a single Excel file with separate sheets for each source.
    * Extracts and normalizes IMO numbers for consistent data.
* **Excel File Comparator:**
    * Compares two Excel files and highlights common 7-digit numerical identifiers (e.g., IMO numbers) in red across all relevant sheets.
    * Allows choosing which file to highlight or saving highlighted copies of both.
    * Includes a search functionality to find specific numerical values within the loaded Excel data.
* **My Vessels Management:**
    * Allows users to maintain a personal list of vessels (Name and IMO Number).
    * Persists the list locally for future sessions.
    * Enables exporting the personal vessel list to an Excel file.
    * Provides a quick comparison feature to check vessels on your list against the most recently generated sanctions report, indicating if any of your vessels are sanctioned and by which source.
* **User-Friendly Interface:**
    * Intuitive tabbed interface for easy navigation between functionalities.
    * Real-time logging and progress updates.
    * Modern visual theme.

## Requirements

* Python 3.7 or higher
* Required Python libraries:
    * `tkinter` (usually comes with Python)
    * `pandas`
    * `requests`
    * `beautifulsoup4` (bs4)
    * `lxml` (for parsing HTML/XML content)
    * `openpyxl` (for reading/writing Excel .xlsx files)
    * `fake_useragent`

## Installation

1.  **Ensure Python is installed:**
    Download and install Python from [python.org](https://www.python.org/downloads/). Make sure to check "Add Python to PATH" during installation.

2.  **Install required libraries:**
    Open your terminal or command prompt and run the following command:
    ```bash
    pip install pandas requests beautifulsoup4 lxml openpyxl fake-useragent
    ```

## How to Use

1.  **Run the Application:**
    Save the provided Python code as a `.py` file (e.g., `sanctions_tool.py`) and run it from your terminal:
    ```bash
    python sanctions_tool.py
    ```

2.  **Sanctions Report Generator Tab:**
    * Navigate to the "Sanctions Report" tab.
    * Click the "Generate Sanctions Report" button.
    * A "Save As" dialog will appear. Choose a location and filename (e.g., `Sanctions_Report.xlsx`) for the generated Excel file.
    * The application will fetch data from various online sources. Progress and status messages will be displayed in the "Log Output" area.
    * Once complete, an Excel file with separate sheets for OFAC, UK, EU DMA, and UANI vessels will be saved. The path to this file will automatically populate "File 1" and "File 2" in the "Excel Comparator" tab for immediate comparison.

3.  **Excel Comparator Tab:**
    * Navigate to the "Excel Comparator" tab.
    * **Select Files:** Use the "Browse" buttons to select your "File 1" (Baseline Excel File) and "File 2" (Comparison Excel File). (The Sanctions Report output will automatically fill File 1 & 2 if generated).
    * **Output Options:** Choose whether to save the highlighted output files in the same directory as the input files or select a different output folder.
    * **Compare and Highlight:** Click "Compare and Highlight". The application will identify common 7-digit numbers (like IMO numbers) across both files and highlight them in red in newly generated Excel files (copies of your originals).
    * **Search Numeric Values:** After a comparison has been run (which loads the data into memory), you can enter a 7-digit number in the "Search Number" field and click "Search" to find its occurrences across the loaded Excel files. Results will appear in the "Search Results" tree.

4.  **My Vessels Tab:**
    * Navigate to the "My Vessels" tab.
    * **Add New Vessel:** Enter a "Vessel Name" and a 7-digit "IMO Number" in the respective fields and click "Add Vessel". Your vessel will be added to the list and saved locally for future use.
    * **Remove Vessels:** Select one or more vessels from the list and click "Remove Selected Vessel(s)" to delete them.
    * **Export My Vessels to Excel:** Click this button to save your current personal vessel list to a new Excel file.
    * **Compare with Sanctions Report:** Click this button to generate a new Excel report that checks your personal vessels against the most recently generated sanctions data. This report will indicate if your vessels are found on any sanctions list and specify the source(s).

5.  **About This Application Tab:**
    * Provides information about the application's purpose, usage instructions, and contact details.

## Configuration (Data Sources)

The application uses the following public data sources:

* **OFAC SDN List:** `https://www.treasury.gov/ofac/downloads/sdn.csv`
* **UK Sanctions List:** `https://www.gov.uk/government/publications/the-uk-sanctions-list`
* **EU (DMA) Designated Vessels List:** `https://www.dma.dk/Media/638834044135010725/2025118019-7%20Importversion%20-%20List%20of%20EU%20designated%20vessels%20(20-05-2025)%203010691_2_0.XLSX`
* **UANI Tracked Vessels (Google Sheet):** `https://docs.google.com/spreadsheets/d/19SBq7N1Ety5fCfaTOZUf61QY-hJIouptx9Gv-uosR_k/export?format=csv&gid=0`

## Troubleshooting

* **"Dependency Error" on startup:** Ensure all required libraries are installed (`pip install ...`). For `lxml` specifically, confirm it's installed.
* **"Network error fetching URL":** Check your internet connection. The application relies on external websites for sanctions data.
* **Highlighting not working as expected:**
    * Ensure the column containing IMO numbers in your Excel files is named clearly (e.g., "IMO Number", "IMO", "ID", "Vessel Name" if IMOs are embedded there). The tool searches for common patterns in column headers.
    * Confirm the numbers are consistently 7 digits.
* **"My Vessels" data not saving:** Check if the `my_vessels.csv` file is being created in the same directory as your script. Ensure you have write permissions to that directory.
## Compiling
```bash
   python -m venv venv

   pip install pandas requests beautifulsoup4 lxml openpyxl fake-useragent
    
   pyinstaller --onefile --windowed --icon=app_icon.ico your_script_name.py
```
## Contact

For any problems or queries, please contact: `Delta.fsociety@tutamail.com`

## License

This project is Free ( as in libre ) and under GPL3 License.
