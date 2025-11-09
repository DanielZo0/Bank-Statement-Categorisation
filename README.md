# Bank Statement Categorization Tools

Python scripts to automatically categorize transactions from bank statements. This replicates Excel-based categorization logic with automated processing.

## Supported Banks

- **Wamo** (Wise) - PDF statements
- **Bank of Valletta (BoV)** - CSV statements

## Features

- Extracts transactions from Wamo PDF bank statements
- Automatically categorizes transactions by type (transfers, fees, salaries, etc.)
- Extracts invoice numbers and counterparty information
- Splits transactions into INCOMING and OUTGOING sheets
- Color-codes transactions by month for easy visualization
- Exports to Excel format with proper formatting

## Installation

1. Clone this repository:
```bash
git clone https://github.com/DanielZo0/Wamo-Categorisation.git
cd Wamo-Categorisation
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Quick Start (Easiest Method)

**Option 1: Double-click the batch file (Windows)**

Simply double-click `Categorize Statement.bat` and follow the prompts. The tool will:
1. Let you browse for your statement file (PDF or CSV)
2. Ask for an output filename
3. Automatically detect and process your statement
4. Open the result in Excel

**Option 2: Run the unified script**

```bash
python categorize_statement.py
```

The interactive tool provides:
- File browser dialog (GUI) to select your statement
- Or manual file path entry
- Automatic detection of file type (PDF/CSV)
- Custom output filename selection
- Summary before processing
- Option to automatically open the result

### Advanced Usage (Direct Scripts)

You can also run the individual scripts directly:

**For Wamo (Wise) PDF Statements:**

```bash
python wamo_categorization.py statement.pdf [output.xlsx]
```

**For Bank of Valletta CSV Statements:**

```bash
python bov_categorization.py statement.csv [output.xlsx]
```

## Output

Both scripts generate an Excel file with three sheets:

1. **SOURCE** - Raw transaction data (formatted as Excel table)
2. **INCOMING** - Positive transactions (income) with categorization (formatted as Excel table)
3. **OUTGOING** - Negative transactions (expenses) with categorization (formatted as Excel table)

Each sheet includes:
- **Date** - Transaction date (formatted as yyyy-mm-dd)
- **Detail** - Transaction description
- **Amount** - Transaction amount (formatted as currency with proper number formatting)
- **Type** - Auto-categorized transaction type
- **Invoice** - Extracted invoice numbers
- **Counterparty** - Extracted payee/payer names

### Formatting Features
- All data ranges are formatted as proper Excel tables with filters
- Rows are color-coded by month for easy visual grouping (12 distinct colors)
- Date column: yyyy-mm-dd format
- Amount column: Currency format with thousands separators (#,##0.00)
- Auto-adjusted column widths for optimal readability

## Transaction Categories

The script automatically recognizes and categorizes:

- Bank transfers (SCT, instant payments, internal)
- Fees and charges
- Salary and employment payments
- Loan repayments
- Tax and government payments
- Direct debits
- Insurance payments
- Retail and food purchases
- Utility payments
- And more...

## Examples

### Wamo Statement Processing

```bash
python wamo_categorization.py statement_7068982_EUR_2025-06-01_2025-09-30.pdf output.xlsx
```

Output:
```
Processing: statement_7068982_EUR_2025-06-01_2025-09-30.pdf
Extracting transactions from PDF...
Found 59 transactions
Categorizing transactions...
  Incoming: 10 transactions
  Outgoing: 49 transactions
Exporting to Excel...
Excel file created: output.xlsx

Complete! Output saved to: output.xlsx
```

The script successfully:
- Extracted 59 transactions from a 4-page Wamo PDF statement
- Identified 10 incoming transactions (payments received, cashback)
- Identified 49 outgoing transactions (card payments, transfers, fees)
- Categorized all transactions by type
- Extracted counterparty information (merchants, recipients)
- Applied month-based color coding for easy visual analysis

### Bank of Valletta Statement Processing

```bash
python bov_categorization.py statement_BOV.csv bov_output.xlsx
```

Output:
```
Processing: statement_BOV.csv
Extracting transactions from CSV...
Found 338 transactions
Categorizing transactions...
  Incoming: 169 transactions
  Outgoing: 169 transactions
Exporting to Excel...
Excel file created: bov_output.xlsx

Complete! Output saved to: bov_output.xlsx
```

The script successfully:
- Extracted 338 transactions from BoV CSV export
- Identified 169 incoming transactions (cheque deposits, SCT transfers, payments received)
- Identified 169 outgoing transactions (cheques, transfers, fees, 24x7 payments)
- Categorized all transactions by type
- Extracted counterparty information from transaction details
- Created formatted Excel tables with month-based color coding

## Requirements

- Python 3.8 or higher
- pandas
- PyPDF2
- xlsxwriter
- openpyxl

## License

MIT License

## Author

Daniel Zo

