# Bank Statement Categorization - Usage Guide

## Quick Start

### Method 1: Double-Click (Easiest - Windows)

1. **Double-click** `Categorize Statement.bat`
2. The program will start and show you a menu
3. Choose option **1** to browse for your file
4. A file selection dialog will open - select your statement (PDF or CSV)
5. Press **Enter** to accept the default output filename, or type a custom name
6. Review the summary and press **Enter** to proceed
7. The program will process your statement and create the Excel file
8. Choose **Yes** to automatically open the result

### Method 2: Run from Terminal

Open a terminal/command prompt in the project folder and run:

```bash
python categorize_statement.py
```

## Step-by-Step Example

```
+============================================================+
|                                                            |
|           BANK STATEMENT CATEGORIZATION TOOL               |
|                                                            |
|       Supports: Wamo PDF & Bank of Valletta CSV           |
|                                                            |
+============================================================+

Step 1: Select your statement file
------------------------------------------------------------

Choose how to select your file:
  1. Browse for file (GUI dialog)
  2. Enter file path manually
  3. Exit

Your choice (1-3): 1

Opening file selection dialog...
Please select your bank statement file (PDF or CSV)

[OK] Selected file: statement_BOV.csv
File type: CSV (Bank of Valletta Statement)

------------------------------------------------------------
Step 2: Choose output filename
------------------------------------------------------------

Default output filename: categorized_statement_BOV.xlsx
Press Enter to use default, or type a custom filename:
Output filename: [press Enter]

[OK] Output will be saved as: categorized_statement_BOV.xlsx

------------------------------------------------------------
Summary
------------------------------------------------------------
  Input:  statement_BOV.csv
  Type:   CSV
  Output: categorized_statement_BOV.xlsx
------------------------------------------------------------

Proceed with processing? (Y/n): Y

============================================================
  PROCESSING STATEMENT
============================================================

Processing: statement_BOV.csv
Extracting transactions from CSV...
Found 338 transactions
Categorizing transactions...
  Incoming: 169 transactions
  Outgoing: 169 transactions
Exporting to Excel...
Excel file created: categorized_statement_BOV.xlsx

Complete! Output saved to: categorized_statement_BOV.xlsx

============================================================
  SUCCESS!
============================================================

Categorized statement saved to: categorized_statement_BOV.xlsx

The Excel file contains:
  - SOURCE sheet - Raw transaction data
  - INCOMING sheet - Income transactions with categorization
  - OUTGOING sheet - Expense transactions with categorization

All sheets are formatted as Excel tables with:
  - Month-based color coding
  - Proper date and currency formatting
  - Sortable/filterable columns

============================================================

Would you like to open the output file? (Y/n): Y
Opening file...

Press Enter to exit...
```

## Features

### File Selection Options

**Option 1: GUI File Browser**
- Opens a graphical file selection dialog
- Easy point-and-click interface
- Filters for PDF and CSV files only
- Shows file type immediately

**Option 2: Manual Path Entry**
- Type or paste the full path to your file
- Supports drag-and-drop (drag file into terminal)
- Automatically removes quotes
- Validates file existence

### Automatic File Type Detection

The tool automatically detects:
- **PDF files** → Uses Wamo categorization engine
- **CSV files** → Uses Bank of Valletta categorization engine

### Output Customization

- Default filename based on input: `categorized_[input_name].xlsx`
- Or specify your own custom filename
- Automatically adds `.xlsx` extension if missing

### Processing Summary

Before processing, you'll see:
- Input filename
- Detected file type
- Output filename
- Confirmation prompt (safety feature)

### Auto-Open Result

After successful processing:
- Option to automatically open the Excel file
- Works on Windows, macOS, and Linux
- Or manually navigate to the file

## Troubleshooting

### "Python is not installed"
- Install Python 3.8 or higher from [python.org](https://python.org)
- Make sure to check "Add Python to PATH" during installation

### "Module not found" errors
Run:
```bash
pip install -r requirements.txt
```

### File dialog doesn't appear
- The dialog might be behind other windows
- Check your taskbar for a new window
- Try using Option 2 (manual path entry) instead

### "No transactions found"
- Verify the file format matches the expected type
- PDF files should be Wamo statements
- CSV files should be Bank of Valletta exports

## Tips

1. **Keep your statements organized** - Store all statements in one folder for easy selection

2. **Use descriptive output names** - E.g., `expenses_q3_2025.xlsx` instead of default names

3. **Process multiple statements** - Run the tool multiple times to process different statements

4. **Review the output** - Always verify the categorization matches your expectations

5. **Drag and drop** - If using manual path entry, you can drag the file into the terminal window on most systems

## Excel Output Details

The generated Excel file contains three professionally formatted sheets:

### SOURCE Sheet
- Raw transaction data as extracted from the statement
- Formatted as an Excel table with filters
- Columns: Date, Detail, Amount

### INCOMING Sheet
- All positive transactions (money received)
- Additional columns: Type, Invoice, Counterparty
- Month-based color coding (12 distinct colors)
- Currency formatting for amounts

### OUTGOING Sheet
- All negative transactions (money spent)
- Same structure as INCOMING sheet
- Different color scheme for easy distinction

### Data Formatting
- **Dates**: yyyy-mm-dd format
- **Amounts**: Currency format with thousands separators (#,##0.00)
- **Tables**: Excel native tables with sort/filter capabilities
- **Colors**: Automatic month-based row coloring for visual grouping

## Support

For issues or questions:
1. Check the main [README.md](README.md) for detailed feature information
2. Review error messages carefully - they often indicate the exact issue
3. Ensure your Python environment has all required packages installed
4. Verify your statement file is in the correct format (Wamo PDF or BoV CSV)

