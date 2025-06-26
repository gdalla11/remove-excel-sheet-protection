# Excel Sheet Protection Remover

A Python utility that safely removes password protection from Excel worksheet sheets without requiring the original password. Uses string-based pattern matching to preserve file integrity and prevent data corruption.

## Features

- ✅ **Safe & Reliable** - Uses string replacement instead of XML parsing to prevent corruption
- 🔍 **Inspection Mode** - Preview protected sheets before making changes
- 💾 **Automatic Backup** - Creates backup files before any modifications
- 🛡️ **Data Preservation** - Maintains all formulas, formatting, and content
- 📊 **Multiple Sheet Support** - Handles workbooks with multiple protected sheets
- ✔️ **File Integrity Check** - Verifies output file validity

## Supported Formats

- `.xlsx` (Excel Workbook)
- `.xlsm` (Excel Macro-Enabled Workbook)

## What It Does

Removes sheet-level protection that prevents users from:
- Editing cells
- Inserting/deleting rows and columns
- Formatting cells
- Modifying sheet structure

**Note:** This tool removes *sheet protection* only, not workbook protection or file encryption.

## Quick Start

1. Install Python 3.7+ with PATH enabled
2. Download `remove_sheet_protection.py`
3. Run: `python remove_sheet_protection.py`
4. Follow the prompts

See the full setup guide in `README.txt` for detailed installation instructions.

## Legal Notice

This tool is intended for legitimate use only. Only use on Excel files you own or have explicit permission to modify. Users are responsible for compliance with their organization's policies and applicable laws.

## Requirements

- Python 3.7 or higher
- No additional dependencies required (uses built-in libraries only)








===============================================================================
                    EXCEL SHEET PROTECTION REMOVER
                     Setup and Usage Guide
===============================================================================

A Python script that safely removes password protection from Excel sheets 
without knowing the password. This tool preserves all data, formulas, and 
formatting while removing sheet-level protection.

===============================================================================
IMPORTANT NOTES
===============================================================================

- ALWAYS BACKUP YOUR FILES - While this tool creates automatic backups, 
  keep your own copies
- SHEET PROTECTION ONLY - This removes sheet protection, not workbook 
  protection or file encryption

===============================================================================
WHAT YOU'LL NEED
===============================================================================

1. Python (we'll install this)
2. The script file (remove_sheet_protection.py)  
3. Your protected Excel file (.xlsx or .xlsm)

===============================================================================
STEP 1: INSTALL PYTHON
===============================================================================

Download Python
----------------
1. Go to https://www.python.org/downloads/
2. Click the big yellow "Download Python" button
3. This will download the latest version (3.11+ recommended)

Install Python (CRITICAL STEP)
-------------------------------
1. Run the installer you just downloaded
2. BEFORE clicking "Install Now" - Look at the bottom of the installer window
3. CHECK THE BOX that says "Add Python to PATH" or 
   "Add Python to environment variables"
   
   *** THIS STEP IS CRUCIAL *** 
   Without it, you won't be able to run Python from Command Prompt

4. Click "Install Now"
5. Wait for installation to complete
6. Click "Close" when done

Verify Installation
-------------------
1. Press Windows Key + R
2. Type cmd and press Enter
3. In the black Command Prompt window, type:
   python --version
4. You should see something like: Python 3.11.4

If you get an error like "'python' is not recognized", you need to reinstall 
Python and make sure to check the "Add Python to PATH" box.

===============================================================================
STEP 2: GET THE SCRIPT
===============================================================================

1. Save the script - Copy the remove_sheet_protection.py file to a folder 
   on your computer
2. Recommended location: C:\Users\YourName\Documents\ or your Desktop
3. Note the full path - you'll need to navigate here in Command Prompt

===============================================================================
STEP 3: RUN THE SCRIPT
===============================================================================

Open Command Prompt
--------------------
1. Press Windows Key + R
2. Type cmd and press Enter
3. A black window (Command Prompt) will open

Navigate to Script Location
---------------------------
If you saved the script to your Documents folder:
cd C:\Users\YourName\Documents

Replace YourName with your actual Windows username. For example:
cd C:\Users\john.smith\Documents

Run the Script
--------------
Type this command and press Enter:
python remove_sheet_protection.py

===============================================================================
STEP 4: USING THE TOOL
===============================================================================

The Process
-----------
1. File Path Prompt:
   Enter full path to the .xlsm/.xlsx file:
   
   TIP: You can drag and drop your Excel file into the Command Prompt window 
   to automatically insert the full path!

2. Inspection Option:
   Do you want to inspect the file first to see what protection exists? (y/n):
   
   - Type y and press Enter to see what's protected before making changes
   - Type n to skip inspection and proceed directly

3. Review Results: The tool will show you:
   - Which sheets were protected
   - What type of protection was removed
   - Where the new unprotected file was saved

Example Run
-----------
Excel Sheet Protection Remover (Safe String Method)
=======================================================
Enter full path to the .xlsm/.xlsx file: C:\Users\john\Documents\MyFile.xlsx
Do you want to inspect the file first? (y/n): y
Backup created: C:\Users\john\Documents\MyFile.xlsx.backup
Extracting Excel archive...
Mapping sheet names...

INSPECTION MODE - Analyzing protection elements:
Found 3 worksheet files

Inspecting sheet1.xml:
   No protection elements found

Inspecting sheet2.xml:
   sheetProtection (self-closing): 1 found

Inspecting sheet3.xml:
   sheetProtection (self-closing): 1 found

Do you want to proceed with removing the protection? (y/n): y

Removing protection from sheets...
Successfully unprotected 2 sheet(s):
   • Data Sheet (sheet2.xml)
   • Analysis (sheet3.xml)

Creating unprotected file...
Done! New file created: C:\Users\john\Documents\MyFile_unprotected.xlsx
File integrity verified

===============================================================================
OUTPUT FILES
===============================================================================

The tool creates several files:

- OriginalName_unprotected.xlsx - Your unprotected Excel file
- OriginalName.xlsx.backup - Automatic backup of original file

===============================================================================
TROUBLESHOOTING
===============================================================================

"python is not recognized"
---------------------------
Problem: Python wasn't added to PATH during installation
Solution: Reinstall Python and check the "Add Python to PATH" box

"File not found"
----------------
Problem: Wrong file path or file doesn't exist
Solution: 
- Double-check the file path
- Try dragging and dropping the file into Command Prompt
- Make sure the file exists and isn't open in Excel

"No protection found" but sheets are protected
----------------------------------------------
Problem: Different type of protection (workbook protection, file encryption)
Solution: This tool only removes sheet-level protection

File corruption or Excel repair messages
-----------------------------------------
Problem: Rare issue with complex Excel files
Solution: 
- Use the backup file created automatically
- Try the inspection mode first to see what's being modified

Permission errors
-----------------
Problem: File is open in Excel or read-only
Solution: Close Excel completely and ensure the file isn't read-only

===============================================================================
SAFETY FEATURES
===============================================================================

- Automatic Backup: Creates .backup file before any changes
- Inspection Mode: Preview what will be changed before proceeding  
- File Integrity Check: Verifies the output file is valid
- Safe String Method: Uses text replacement instead of XML parsing to 
  prevent corruption
- Error Recovery: Attempts to restore original content if errors occur

===============================================================================
TIPS FOR SUCCESS
===============================================================================

1. Close Excel completely before running the tool
2. Use inspection mode first to see what protection exists
3. Test the output file before deleting backups
4. Keep the backup files until you're sure everything works
5. Run from a simple folder path (avoid spaces and special characters 
   in folder names)

===============================================================================
SHARING WITH COLLEAGUES
===============================================================================

When sharing this tool:

1. Share both files: remove_sheet_protection.py and this README
2. Emphasize the Python installation step - the PATH checkbox is critical
3. Recommend testing on a copy first

===============================================================================
QUICK REFERENCE COMMANDS
===============================================================================

Task                          | Command
------------------------------|----------------------------------------
Check Python is installed     | python --version
Navigate to Documents         | cd C:\Users\YourName\Documents
Run the script               | python remove_sheet_protection.py
Go back one folder           | cd ..
See current folder contents  | dir

===============================================================================
NEED HELP?
===============================================================================

If you run into issues:

1. Double-check Python installation - especially the PATH setting
2. Try the inspection mode to see what the tool finds
3. Check file permissions - make sure the file isn't read-only or open
4. Use the backup file if something goes wrong

Remember: This tool has been tested and works reliably, but always keep 
backups of important files!

===============================================================================
