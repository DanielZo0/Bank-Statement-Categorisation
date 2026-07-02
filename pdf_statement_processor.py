#!/usr/bin/env python3
"""
PDF Bank Statement Categorization Script
Processes PDF bank statements and categorizes transactions
"""

import re
import pandas as pd
import PyPDF2
from pathlib import Path
import sys

from common_categorization import parse_date_smart, parse_number
from shared_statement_pipeline import (
    export_to_excel as export_to_excel_shared,
    process_transactions as process_transactions_shared,
)


def parse_wamo_statement_text(full_text: str) -> pd.DataFrame:
    """Parse extracted Wamo statement text into a transaction DataFrame."""
    transactions = []
    lines = full_text.split("\n")

    in_transactions = False
    current_transaction = None
    previous_line = ""

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if re.search(r"Description\s+Incoming\s+Outgoing\s+Amount", line, re.IGNORECASE):
            in_transactions = True
            previous_line = ""
            continue

        if in_transactions and re.search(
            r"(Opening Balance|Closing Balance|Total|Page \d+)", line, re.IGNORECASE
        ):
            in_transactions = False
            continue

        if not in_transactions:
            continue

        date_match = re.match(
            r"^(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})\s+(.+)",
            line,
        )
        if not date_match:
            previous_line = line
            continue

        if current_transaction and current_transaction.get("description"):
            transactions.append(current_transaction)

        date_str = date_match.group(1)
        rest_of_line = date_match.group(2)
        date_obj = parse_date_smart(date_str)

        clean_line = re.sub(
            r"Transaction:\s*[A-Z_]+-[a-f0-9-]{36}", "Transaction: [ID]", rest_of_line
        )
        clean_line = re.sub(r"Transaction:\s*[A-Z]+-\d{10}", "Transaction: [ID]", clean_line)
        clean_line = re.sub(r"(-\d+)(\d{1,3},\d{3}\.\d{2})", r"\1 \2", clean_line)
        amounts = re.findall(r"[-]?\b(?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2}\b", clean_line)

        incoming = None
        outgoing = None
        if len(amounts) >= 2:
            transaction_amount = parse_number(amounts[-2])
            if amounts[-2].startswith("-"):
                outgoing = abs(transaction_amount)
            else:
                incoming = abs(transaction_amount)

        if previous_line and not re.match(
            r"^\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)",
            previous_line,
        ):
            description = previous_line + " "
        else:
            description = ""

        desc_part = clean_line
        for amount in amounts:
            desc_part = desc_part.replace(amount, "")
        description += desc_part.strip()

        current_transaction = {
            "Date": date_obj,
            "description": description,
            "incoming": incoming,
            "outgoing": outgoing,
        }
        previous_line = ""

    if current_transaction and current_transaction.get("description"):
        transactions.append(current_transaction)

    parsed_rows = []
    for trans in transactions:
        date = trans.get("Date")
        description = trans.get("description", "").strip()
        incoming = trans.get("incoming", 0) or 0
        outgoing = trans.get("outgoing", 0) or 0

        amount = incoming if incoming > 0 else (-outgoing if outgoing > 0 else 0)
        if date and description:
            parsed_rows.append({"Date": date, "Detail": description, "Amount": amount})

    if not parsed_rows:
        return pd.DataFrame(columns=["Date", "Detail", "Amount"])
    return pd.DataFrame(parsed_rows)


def extract_transactions_from_pdf(pdf_path: str) -> pd.DataFrame:
    """
    Extract transaction data from Wamo PDF statement
    Wamo format: Date | Description | Incoming | Outgoing | Balance
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            full_text = "\n".join((page.extract_text() or "") for page in pdf_reader.pages)
            result_df = parse_wamo_statement_text(full_text)
    
    except Exception as e:
        print(f"Error reading PDF: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

    if result_df.empty:
        print("Warning: No transactions found in PDF")
    return result_df


def process_transactions(df: pd.DataFrame):
    return process_transactions_shared(df)


def export_to_excel(
    source_df: pd.DataFrame, incoming_df: pd.DataFrame, outgoing_df: pd.DataFrame, output_path: str
):
    return export_to_excel_shared(source_df, incoming_df, outgoing_df, output_path)


def main():
    """Main execution function"""
    if len(sys.argv) < 2:
        print("Usage: python wamo_categorization.py <path_to_wamo_statement.pdf> [output.xlsx]")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "categorized_statement.xlsx"
    
    if not Path(pdf_path).exists():
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    
    print(f"Processing: {pdf_path}")
    
    # Extract transactions from PDF
    print("Extracting transactions from PDF...")
    df = extract_transactions_from_pdf(pdf_path)
    
    if df.empty:
        print("Error: No transactions found in PDF")
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
