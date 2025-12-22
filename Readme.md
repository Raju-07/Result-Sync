# Result Sync

## Overview
Result Sync is a Python-based GUI application designed to automate the extraction of BCA (Bachelor of Computer Applications) examination results from the Maharshi Dayananda University (MDU) result portal. The application uses web scraping techniques to fetch individual student results and compiles them into a formatted Excel report. This tool is particularly useful for educators, administrators, or students who need to process multiple result records efficiently.

**Author:** Raju Yadav  
**Copyright:** ©2025 Raju Yadav  
**Powered by:** Python, made with Love ❤

## Features
- **User-Friendly GUI:** Built with CustomTkinter for a modern, responsive interface.
- **Automated Web Scraping:** Utilizes Selenium to interact with the MDU result portal and extract data.
- **Excel Integration:** Reads student records from an input Excel file and writes results to an output Excel file with professional formatting.
- **Progress Tracking:** Real-time progress bar and status updates during processing.
- **Error Handling:** Robust error handling with resume capability and user notifications.
- **Threading:** Prevents GUI freezing by running the extraction process in a separate thread.
- **Feedback System:** Includes a link for user feedback.

## Requirements
### System Requirements
- **Operating System:** Windows, macOS, or Linux (tested on Linux)
- **Python Version:** Python 3.7 or higher
- **Browser:** Microsoft Edge (required for Selenium WebDriver)
- **Internet Connection:** Required for accessing the MDU result portal

### Python Libraries
The following libraries must be installed:
- `customtkinter` - For the GUI interface
- `openpyxl` - For Excel file handling
- `selenium` - For web automation
- `pillow` (PIL) - For image processing (used for icons)
- `webbrowser` - Built-in Python module
- `threading` - Built-in Python module
- `os` - Built-in Python module
- `tkinter` - Built-in Python module (usually comes with Python)

## Installation
### Step 1: Install Python
Ensure Python 3.7+ is installed on your system. Download from [python.org](https://www.python.org/downloads/).

### Step 2: Install Required Libraries
Open a terminal/command prompt and run the following commands:

```bash
pip install customtkinter
pip install openpyxl
pip install selenium
pip install pillow
```

### Step 3: Install Microsoft Edge
If not already installed, download and install Microsoft Edge from [microsoft.com](https://www.microsoft.com/en-us/edge).

### Step 4: Download the Project
Clone or download the project files to your local machine.

## Usage
### Preparing Input Data
1. Create an Excel file (.xlsx) with student records.
2. The first row must be headers.
3. Column A: Registration Number
4. Column B: Roll Number
5. Save the file in .xlsx format.

### Running the Application
1. Navigate to the project directory in your terminal.
2. Run the application:
   ```bash
   python Result-Sync/ResultSync.py
   ```
3. The GUI window will open.

### Step-by-Step Procedure
1. **Launch the App:** Run the script as described above. The main window titled "Result Sync" will appear.

2. **Select Input File:**
   - In the "Required Info" frame, click the folder icon next to "Select Record File (.xlsx)".
   - Choose your prepared Excel file containing student registration and roll numbers.
   - The file path will be displayed in the text box.

3. **Select Output Location:**
   - Click the folder icon next to "Location to Save".
   - Choose where to save the output Excel file and provide a filename (e.g., "BCA_Results_2025.xlsx").
   - The save path will be displayed in the text box.

4. **Proceed:**
   - Click the "Proceed" button.
   - The app will validate the inputs and display a progress bar, percentage, and estimated time.

5. **Start Processing:**
   - Click the "Start" button.
   - The application will begin extracting results from the MDU portal.
   - Monitor the progress bar and current student status.
   - The process may take several minutes depending on the number of records.

6. **Completion:**
   - Once finished, a success message will appear.
   - The output Excel file will contain formatted results with student details, marks, and pass/fail status.

### Output Format
The generated Excel file includes:
- **Title:** "BCA Result 2025 Sem - V" (merged cells A1:N1)
- **Section Headers:** Student Details, Marks in Each Subjects, Result
- **Columns:** Sr.No, Name, Father Name, Registration No, Roll No, (empty), BCA206, BCA207, BCA208, BCA209, BCA210, (empty), Total Marks, Result
- Professional formatting with colors and fonts.

## Important Notes
- **Data Format:** Ensure the input Excel file has registration numbers in column A and roll numbers in column B, starting from row 2 (row 1 is headers).
- **Error Handling:** If an error occurs during processing, the app will save progress and display an error message. You may need to reselect the output file and restart.
- **Resume Capability:** The app tracks processed entries and can resume from where it left off if interrupted.
- **Internet Dependency:** A stable internet connection is crucial for web scraping.
- **Browser Compatibility:** Currently configured for Microsoft Edge. Modifications may be needed for other browsers.
- **Legal Notice:** Ensure compliance with MDU's terms of service when using automated tools.

## Troubleshooting
- **File Not Found:** Verify the input file path and ensure the file exists.
- **WebDriver Issues:** Ensure Microsoft Edge is installed and up-to-date. Selenium should handle WebDriver automatically, but you may need to download it manually if issues persist.
- **Permission Errors:** Ensure write permissions for the output directory.
- **Slow Processing:** Processing time depends on internet speed and server response. Be patient for large datasets.
- **GUI Freezing:** The app uses threading to prevent freezing, but very large datasets may still take time.

## Feedback
Your feedback is valuable! Click the "Click Here" link in the app or visit: [Feedback Form](https://forms.gle/PoXZRNdXS1P14MsaA)

## License
This project is proprietary software. ©2025 Raju Yadav. All rights reserved.

## Contributing
For contributions or modifications, please contact the author.

## Changelog
- **v1.0:** Initial release with core functionality for BCA result extraction.