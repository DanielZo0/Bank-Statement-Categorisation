# Bank Statement Categorization - Usage Guide

## Quick start

### Windows (double-click)
1. Double-click `Process Statements.bat`.
2. Select one or more statement files (`.pdf` or `.csv`) in the file dialog.
3. Wait for processing to complete.
4. Press Enter to open the output location.

### Command line
```bash
python batch_statement_processor.py
```

This opens the same file picker and supports mixed PDF/CSV batch processing.

## Direct processing (automation-friendly)

### PDF
```bash
python pdf_statement_processor.py statement.pdf output.xlsx
```

### CSV
```bash
python csv_statement_processor.py statement.csv output.xlsx
```

## Output

For each input statement, the tool writes:
- `categorized_[original_name].xlsx`
- In the same directory as the input file

Each workbook contains:
- `SOURCE` sheet (raw transactions)
- `INCOMING` sheet (categorized positive transactions)
- `OUTGOING` sheet (categorized negative transactions)

## Troubleshooting

### Python not found
Install Python 3.8+ and ensure `python` is available on PATH.

### Missing dependencies
```bash
pip install -r requirements.txt
```

### No transactions found
Confirm the statement format is supported:
- PDF: Wamo statements
- CSV: Bank of Valletta exports
