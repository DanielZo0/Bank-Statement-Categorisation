#!/usr/bin/env python3
"""
Bank of Valletta Statement Categorization Script
Processes BoV CSV statements and categorizes transactions
"""

import re
import pandas as pd
from datetime import datetime
from typing import List, Tuple, Optional
from pathlib import Path
import sys


# Month colors for Excel formatting
MONTH_COLORS = {
    1: "#FFCCCC", 2: "#FFE5CC", 3: "#FFFFCC",
    4: "#E5FFCC", 5: "#CCFFCC", 6: "#CCFFE5",
    7: "#CCFFFF", 8: "#CCE5FF", 9: "#CCCCFF",
    10: "#E5CCFF", 11: "#FFCCFF", 12: "#FFCCE5"
}


def parse_number(val: str) -> float:
    """
    Parse number from various formats (EU/US, with currency symbols, parentheses, etc.)
    Handles: €1,234.56, (123.45), 123-, 1.234,56, "1,234.56"
    """
    if not val or not isinstance(val, str):
        return 0.0
    
    val = str(val).strip().strip('"')
    
    # Detect negative indicators
    has_parens = re.match(r'^\(.*\)$', val)
    has_trailing_minus = val.endswith('-')
    is_negative = val.startswith('-')
    
    # Remove currency symbols and spaces
    val = re.sub(r'[\s€$£]', '', val)
    
    # Remove negative signs temporarily
    val = val.replace('-', '').replace('(', '').replace(')', '')
    
    # Handle comma as thousands separator (BoV format: 1,234.56)
    val = val.replace(',', '')
    
    try:
        num = float(val)
        return -num if (has_parens or has_trailing_minus or is_negative) else num
    except ValueError:
        return 0.0


def parse_date_smart(date_str: str) -> Optional[datetime]:
    """
    Parse date from various formats
    BoV uses: yyyy/mm/dd
    Also supports: dd/mm/yyyy, dd-mm-yyyy, ISO formats
    """
    if not date_str:
        return None
    
    date_str = str(date_str).strip()
    
    # Try yyyy/mm/dd (BoV format)
    match = re.match(r'^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$', date_str)
    if match:
        y, m, d = int(match[1]), int(match[2]), int(match[3])
        return datetime(y, m, d)
    
    # Try dd/mm/yyyy or dd-mm-yyyy
    match = re.match(r'^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$', date_str)
    if match:
        d, m, y = int(match[1]), int(match[2]), int(match[3])
        return datetime(y, m, d)
    
    # Try pandas parsing as fallback
    try:
        return pd.to_datetime(date_str, dayfirst=True)
    except:
        return None


def get_transaction_type(detail: str) -> str:
    """
    Categorize transaction based on detail string
    """
    if not detail:
        return "other"
    
    detail_lower = detail.lower()
    
    # Cheques
    if re.search(r'cheque.*deposit', detail_lower):
        return "cheque deposit"
    if re.search(r'cheque.*returned', detail_lower):
        return "cheque returned fee"
    if re.search(r'cheques returned', detail_lower):
        return "cheque returned"
    if re.search(r'cheque', detail_lower):
        return "cheque payment"
    
    # Transfers
    if re.search(r'account to account', detail_lower):
        return "account transfer"
    if re.search(r'transfer between own accounts', detail_lower):
        return "internal transfer"
    if re.search(r'sct inwards', detail_lower):
        return "incoming sct transfer"
    if re.search(r'sct outwards', detail_lower):
        return "outgoing sct transfer"
    if re.search(r'instant payments inwards', detail_lower):
        return "instant payment in"
    if re.search(r'instant payment', detail_lower):
        return "instant payment"
    
    # Fees & charges
    if re.search(r'fee', detail_lower):
        return "bank fee"
    if re.search(r'charge', detail_lower):
        return "bank charge"
    if re.search(r'administration fee', detail_lower):
        return "administration fee"
    if re.search(r'standing instruction charge', detail_lower):
        return "standing instruction charge"
    if re.search(r'standing instruction', detail_lower):
        return "standing instruction"
    
    # Salaries & employment
    if re.search(r'salary', detail_lower):
        return "salary"
    if re.search(r'employment', detail_lower):
        return "employment payment"
    if re.search(r'stipendio|stipend', detail_lower):
        return "stipend/salary"
    
    # Loans & repayments
    if re.search(r'repayment.*principal', detail_lower):
        return "loan principal repayment"
    if re.search(r'repayment.*interest', detail_lower):
        return "loan interest repayment"
    if re.search(r'loan', detail_lower):
        return "loan"
    
    # Taxes & government
    if re.search(r'tax', detail_lower):
        return "tax payment"
    if re.search(r'vat', detail_lower):
        return "vat payment"
    if re.search(r'customs', detail_lower):
        return "customs payment"
    if re.search(r'government|gov', detail_lower):
        return "government payment"
    
    # ATM deposits
    if re.search(r'atm.*cash.*deposit', detail_lower):
        return "atm cash deposit"
    
    # 24x7 payments
    if re.search(r'24x7 pay', detail_lower):
        return "third party payment"
    if re.search(r'24x7 bill', detail_lower):
        return "bill payment"
    if re.search(r'24x7 mobile pay', detail_lower):
        return "mobile payment"
    
    # Direct debits
    if re.search(r'sdd outwards', detail_lower):
        return "direct debit out"
    
    # Insurance
    if re.search(r'mapfre|msv life|insurance', detail_lower):
        return "insurance payment"
    
    # Retail / food / hospitality
    if re.search(r'hotel', detail_lower):
        return "hotel payment"
    if re.search(r'catering', detail_lower):
        return "catering payment"
    if re.search(r'butcher|food|supermarket|restaurant|eat', detail_lower):
        return "food & retail"
    if re.search(r'retail', detail_lower):
        return "retail payment"
    
    # Utilities
    if re.search(r'electricity|water|gas|utility', detail_lower):
        return "utility payment"
    
    # Misc
    if re.search(r'refund', detail_lower):
        return "refund"
    if re.search(r'deposit', detail_lower):
        return "deposit"
    if re.search(r'withdrawal', detail_lower):
        return "withdrawal"
    
    return "other"


def extract_invoice(detail: str) -> str:
    """
    Extract invoice number from transaction detail
    """
    if not detail:
        return ""
    
    match = re.search(r'(invoice|inv|fatt(?:ura)?\s*nr?)\s*([0-9]+)', detail, re.IGNORECASE)
    if match:
        return f"invoice {match[2]}"
    return ""


def extract_counterparty(detail: str) -> str:
    """
    Extract counterparty name from transaction detail
    """
    if not detail:
        return ""
    
    # Check for tax administration reference
    if re.search(r'administratio', detail, re.IGNORECASE):
        tax_ref = re.search(r'ADMINISTRATIO\s+([0-9]+)', detail, re.IGNORECASE)
        if tax_ref:
            return tax_ref[1]
    
    # Clean common transaction prefixes/suffixes
    cleaned = detail
    patterns_to_remove = [
        r'24x7\s*pay\s*third\s*parties',
        r'24x7\s*pay',
        r'third\s*parties',
        r'payment order outwards same day',
        r'payment order outwards',
        r'account to account transfer express deposits',
        r'account to account transfer',
        r'transfer between own accounts',
        r'sct instant payments inwards',
        r'sct inwards',
        r'sct outwards',
        r'standing instruction charge',
        r'standing instruction',
        r'administration fee',
        r'unprocessed standing instruction charge',
        r'sdd outwards fee',
        r'atm cash deposit',
        r'cheque deposit.*$',
        r'cheque returned fee.*$',
        r'cheque book order fee.*$',
        r'cheque\s+\d+.*',
        r'relation:\s*[^,]+',
        r'reason:\s*[^,]+',
        r'value date\s*-\s*[0-9/]+',
        r'ref\s*:\s*[-0-9A-Za-z.]+.*$',
        r'\s+eur\s+[0-9.,]+',
    ]
    
    for pattern in patterns_to_remove:
        cleaned = re.sub(pattern, '', cleaned, flags=re.IGNORECASE)
    
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    # Split on common delimiters
    cleaned = re.split(r'ref\s*:|value date|relation:', cleaned, flags=re.IGNORECASE)[0].strip()
    
    # Look for company suffix
    company_match = re.search(r'\b([A-Z][A-Za-z &.\'-]*\s(?:ltd|limited|plc|co|company))\b', cleaned, re.IGNORECASE)
    if company_match:
        return company_match[1]
    
    # Split on EUR amount
    eur_split = re.split(r'\s+eur\s+', cleaned, flags=re.IGNORECASE)[0].strip()
    if eur_split and len(eur_split) >= 3:
        cleaned = eur_split
    
    # Look for person titles
    person_match = re.search(r'\b(Mr|Ms|Mrs|Dr)\.?\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?\b', cleaned)
    if person_match:
        return person_match[0]
    
    # Look for uppercase sequences
    upper_match = re.search(r'\b([A-Z][A-Z &.\'-]{2,})\b', cleaned)
    if upper_match:
        return upper_match[1]
    
    # Look for capitalized sequences
    cap_match = re.search(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,4})\b', cleaned)
    if cap_match:
        return cap_match[1]
    
    # Return first 5 words
    words = cleaned.split()[:5]
    return ' '.join(words).strip()


def capitalize_first(text: str) -> str:
    """Capitalize first letter, lowercase rest"""
    if not text:
        return ""
    return text[0].upper() + text[1:].lower()


def limit_length(text: str, max_len: int = 26) -> str:
    """Limit text length"""
    if not text:
        return ""
    return text[:max_len] if len(text) > max_len else text


def extract_transactions_from_csv(csv_path: str) -> pd.DataFrame:
    """
    Extract transaction data from BoV CSV statement
    """
    try:
        # Read the CSV file
        with open(csv_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Find the transaction history header
        transaction_start = -1
        for i, line in enumerate(lines):
            if 'Transaction History' in line:
                # Next line should be the column headers
                transaction_start = i + 2
                break
        
        if transaction_start == -1:
            print("Error: Could not find Transaction History header")
            return pd.DataFrame()
        
        # Read transactions from the found position
        df = pd.read_csv(csv_path, skiprows=transaction_start - 1, encoding='utf-8')
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Ensure we have the required columns
        if 'Date' not in df.columns or 'Detail' not in df.columns or 'Amount' not in df.columns:
            print(f"Error: Expected columns not found. Found: {df.columns.tolist()}")
            return pd.DataFrame()
        
        # Parse dates
        df['Date'] = df['Date'].apply(parse_date_smart)
        
        # Parse amounts
        df['Amount'] = df['Amount'].apply(parse_number)
        
        # Remove rows with no date
        df = df[df['Date'].notna()]
        
        # Keep only required columns
        df = df[['Date', 'Detail', 'Amount']].copy()
        
        return df
    
    except Exception as e:
        print(f"Error reading CSV: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


def process_transactions(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Process transactions and split into incoming/outgoing
    """
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()
    
    # Split into incoming (amount >= 0) and outgoing (amount < 0)
    incoming = df[df['Amount'] >= 0].copy()
    outgoing = df[df['Amount'] < 0].copy()
    
    # Process both dataframes
    for transactions_df in [incoming, outgoing]:
        if transactions_df.empty:
            continue
        
        # Add derived columns
        transactions_df['Type'] = transactions_df['Detail'].apply(
            lambda x: limit_length(capitalize_first(get_transaction_type(str(x).lower())))
        )
        transactions_df['Invoice'] = transactions_df['Detail'].apply(
            lambda x: limit_length(capitalize_first(extract_invoice(str(x))))
        )
        transactions_df['Counterparty'] = transactions_df['Detail'].apply(
            lambda x: limit_length(capitalize_first(extract_counterparty(str(x))))
        )
        
        # Convert numeric-only counterparties to numbers
        transactions_df['Counterparty'] = transactions_df['Counterparty'].apply(
            lambda x: int(x) if str(x).isdigit() else x
        )
    
    # Sort by date
    incoming = incoming.sort_values('Date').reset_index(drop=True)
    outgoing = outgoing.sort_values('Date').reset_index(drop=True)
    
    return incoming, outgoing


def export_to_excel(source_df: pd.DataFrame, incoming_df: pd.DataFrame, 
                    outgoing_df: pd.DataFrame, output_path: str):
    """
    Export dataframes to Excel with formatting and proper tables
    """
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define formats
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        currency_format = workbook.add_format({'num_format': '#,##0.00'})
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Write SOURCE sheet
        source_df.to_excel(writer, sheet_name='SOURCE', index=False, startrow=0)
        worksheet_source = writer.sheets['SOURCE']
        
        # Format SOURCE sheet as table
        if not source_df.empty:
            last_row = len(source_df)
            worksheet_source.add_table(0, 0, last_row, 2, {
                'name': 'SOURCE_TABLE',
                'style': 'Table Style Medium 2',
                'columns': [
                    {'header': 'Date', 'format': date_format},
                    {'header': 'Detail'},
                    {'header': 'Amount', 'format': currency_format}
                ]
            })
            
            # Set column widths
            worksheet_source.set_column('A:A', 12, date_format)
            worksheet_source.set_column('B:B', 70)
            worksheet_source.set_column('C:C', 15, currency_format)
        
        # Write and format INCOMING sheet
        if not incoming_df.empty:
            incoming_df.to_excel(writer, sheet_name='INCOMING', index=False, startrow=0)
            worksheet_incoming = writer.sheets['INCOMING']
            
            last_row = len(incoming_df)
            worksheet_incoming.add_table(0, 0, last_row, 5, {
                'name': 'INCOMING_TABLE',
                'style': 'Table Style Medium 9',
                'columns': [
                    {'header': 'Date', 'format': date_format},
                    {'header': 'Detail'},
                    {'header': 'Amount', 'format': currency_format},
                    {'header': 'Type'},
                    {'header': 'Invoice'},
                    {'header': 'Counterparty'}
                ]
            })
            
            # Apply month-based row colors
            for idx, row in incoming_df.iterrows():
                if pd.notna(row['Date']):
                    month = row['Date'].month
                    color = MONTH_COLORS.get(month, "#FFFFFF")
                    date_cell_format = workbook.add_format({
                        'bg_color': color,
                        'num_format': 'yyyy-mm-dd'
                    })
                    text_cell_format = workbook.add_format({
                        'bg_color': color
                    })
                    currency_cell_format = workbook.add_format({
                        'bg_color': color,
                        'num_format': '#,##0.00'
                    })
                    
                    # Apply formatting to data rows (skip header)
                    excel_row = idx + 1
                    worksheet_incoming.write(excel_row, 0, row['Date'], date_cell_format)
                    worksheet_incoming.write(excel_row, 1, row['Detail'], text_cell_format)
                    worksheet_incoming.write(excel_row, 2, row['Amount'], currency_cell_format)
                    worksheet_incoming.write(excel_row, 3, row['Type'], text_cell_format)
                    worksheet_incoming.write(excel_row, 4, row['Invoice'], text_cell_format)
                    worksheet_incoming.write(excel_row, 5, row['Counterparty'], text_cell_format)
            
            # Set column widths
            worksheet_incoming.set_column('A:A', 12)
            worksheet_incoming.set_column('B:B', 50)
            worksheet_incoming.set_column('C:C', 15)
            worksheet_incoming.set_column('D:D', 26)
            worksheet_incoming.set_column('E:E', 26)
            worksheet_incoming.set_column('F:F', 26)
        else:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['Date', 'Detail', 'Amount', 'Type', 'Invoice', 'Counterparty'])
            empty_df.to_excel(writer, sheet_name='INCOMING', index=False)
        
        # Write and format OUTGOING sheet
        if not outgoing_df.empty:
            outgoing_df.to_excel(writer, sheet_name='OUTGOING', index=False, startrow=0)
            worksheet_outgoing = writer.sheets['OUTGOING']
            
            last_row = len(outgoing_df)
            worksheet_outgoing.add_table(0, 0, last_row, 5, {
                'name': 'OUTGOING_TABLE',
                'style': 'Table Style Medium 4',
                'columns': [
                    {'header': 'Date', 'format': date_format},
                    {'header': 'Detail'},
                    {'header': 'Amount', 'format': currency_format},
                    {'header': 'Type'},
                    {'header': 'Invoice'},
                    {'header': 'Counterparty'}
                ]
            })
            
            # Apply month-based row colors
            for idx, row in outgoing_df.iterrows():
                if pd.notna(row['Date']):
                    month = row['Date'].month
                    color = MONTH_COLORS.get(month, "#FFFFFF")
                    date_cell_format = workbook.add_format({
                        'bg_color': color,
                        'num_format': 'yyyy-mm-dd'
                    })
                    text_cell_format = workbook.add_format({
                        'bg_color': color
                    })
                    currency_cell_format = workbook.add_format({
                        'bg_color': color,
                        'num_format': '#,##0.00'
                    })
                    
                    # Apply formatting to data rows (skip header)
                    excel_row = idx + 1
                    worksheet_outgoing.write(excel_row, 0, row['Date'], date_cell_format)
                    worksheet_outgoing.write(excel_row, 1, row['Detail'], text_cell_format)
                    worksheet_outgoing.write(excel_row, 2, row['Amount'], currency_cell_format)
                    worksheet_outgoing.write(excel_row, 3, row['Type'], text_cell_format)
                    worksheet_outgoing.write(excel_row, 4, row['Invoice'], text_cell_format)
                    worksheet_outgoing.write(excel_row, 5, row['Counterparty'], text_cell_format)
            
            # Set column widths
            worksheet_outgoing.set_column('A:A', 12)
            worksheet_outgoing.set_column('B:B', 50)
            worksheet_outgoing.set_column('C:C', 15)
            worksheet_outgoing.set_column('D:D', 26)
            worksheet_outgoing.set_column('E:E', 26)
            worksheet_outgoing.set_column('F:F', 26)
        else:
            # Create empty sheet with headers
            empty_df = pd.DataFrame(columns=['Date', 'Detail', 'Amount', 'Type', 'Invoice', 'Counterparty'])
            empty_df.to_excel(writer, sheet_name='OUTGOING', index=False)
    
    print(f"Excel file created: {output_path}")


def main():
    """Main execution function"""
    if len(sys.argv) < 2:
        print("Usage: python bov_categorization.py <path_to_statement.csv> [output.xlsx]")
        sys.exit(1)
    
    csv_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "categorized_bov_statement.xlsx"
    
    if not Path(csv_path).exists():
        print(f"Error: File not found: {csv_path}")
        sys.exit(1)
    
    print(f"Processing: {csv_path}")
    
    # Extract transactions from CSV
    print("Extracting transactions from CSV...")
    df = extract_transactions_from_csv(csv_path)
    
    if df.empty:
        print("Error: No transactions found in CSV")
        sys.exit(1)
    
    print(f"Found {len(df)} transactions")
    
    # Process and categorize
    print("Categorizing transactions...")
    incoming_df, outgoing_df = process_transactions(df)
    
    print(f"  Incoming: {len(incoming_df)} transactions")
    print(f"  Outgoing: {len(outgoing_df)} transactions")
    
    # Export to Excel
    print("Exporting to Excel...")
    export_to_excel(df, incoming_df, outgoing_df, output_path)
    
    print(f"\nComplete! Output saved to: {output_path}")


if __name__ == "__main__":
    main()

