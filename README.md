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
