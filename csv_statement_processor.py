#!/usr/bin/env python3
"""
CSV Bank Statement Categorization Script
Processes CSV bank statements and categorizes transactions
"""

import pandas as pd
from pathlib import Path
import sys

from common_categorization import parse_date_smart, parse_number
from shared_statement_pipeline import (
    export_to_excel as export_to_excel_shared,
    process_transactions as process_transactions_shared,
)


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


def process_transactions(df: pd.DataFrame):
    return process_transactions_shared(df)


def export_to_excel(
    source_df: pd.DataFrame, incoming_df: pd.DataFrame, outgoing_df: pd.DataFrame, output_path: str
):
    return export_to_excel_shared(source_df, incoming_df, outgoing_df, output_path)


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
