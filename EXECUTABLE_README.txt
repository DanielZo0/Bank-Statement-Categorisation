================================================================================
  BANK STATEMENT PROCESSOR - EXECUTABLE VERSION
================================================================================

VERSION: 1.0
PLATFORM: Windows 64-bit

================================================================================
  QUICK START
================================================================================

1. Double-click "Bank Statement Processor.exe" to launch the application

2. A file selection dialog will open automatically
   - Navigate to your statement files
   - Select one or multiple files (PDF and/or CSV)
   - Hold Ctrl to select multiple files

3. Click "Open" to start processing

4. The application will:
   - Process each selected file
   - Create categorized Excel output files
   - Save outputs in the same folder as your input files
   - Display a summary of results

5. Press Enter to open the output folder

================================================================================
  OUTPUT FILES
================================================================================

For each input file, an Excel file will be created with the format:
"categorized_[original_filename].xlsx"

Each Excel file contains 3 sheets:
- SOURCE: Original transaction data
- INCOMING: Categorized incoming transactions
- OUTGOING: Categorized outgoing transactions

================================================================================
  SUPPORTED FORMATS
================================================================================

PDF Statements:
- Automatically extracts transaction data from PDF bank statements
- Supports Wamo and similar formats

CSV Statements:
- Processes CSV exports from bank systems
- Supports Bank of Valletta and similar formats

================================================================================
  FEATURES
================================================================================

- Smart categorization of transactions by type
- Invoice number extraction
- Counterparty identification
- Month-based color coding
- Accounting columns (Type, Account Reference, Nominal A/C Ref, etc.)
- Static fields: Tax Code (T9), Tax Amount (0.00)

================================================================================
  TROUBLESHOOTING
================================================================================

If the application doesn't start:
- Ensure you're running on Windows 64-bit
- Try running as Administrator (right-click > Run as administrator)
- Check that your antivirus isn't blocking the file

If processing fails:
- Ensure your statement files are not open in Excel
- Check that files are valid PDF or CSV format
- Close any previously generated Excel files

For more help or to report issues:
https://github.com/DanielZo0/Bank-Statement-Categorisation

================================================================================
  CUSTOMIZATION
================================================================================

To customize transaction categorization rules:
1. Download the full source code from GitHub
2. Edit the "common_categorization.py" file
3. Modify the get_transaction_type() function
4. Rebuild the executable or run from Python

================================================================================
  LICENSE
================================================================================

MIT License - Free to use, modify, and distribute

Copyright (c) 2025

================================================================================

