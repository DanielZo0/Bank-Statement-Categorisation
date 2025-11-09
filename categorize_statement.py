#!/usr/bin/env python3
"""
Bank Statement Categorization Tool
Unified interface for processing Wamo PDF and Bank of Valletta CSV statements
"""

import os
import sys
from pathlib import Path
from tkinter import Tk, filedialog
import subprocess


def select_file():
    """
    Open a file selection dialog to choose a statement file
    """
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring dialog to front
    
    print("\n" + "="*60)
    print("  BANK STATEMENT CATEGORIZATION TOOL")
    print("="*60)
    print("\nOpening file selection dialog...")
    print("Please select your bank statement file (PDF or CSV)")
    
    file_path = filedialog.askopenfilename(
        title="Select Bank Statement File",
        filetypes=[
            ("All Supported Files", "*.pdf *.csv"),
            ("PDF Files", "*.pdf"),
            ("CSV Files", "*.csv"),
            ("All Files", "*.*")
        ]
    )
    
    root.destroy()
    return file_path


def get_output_filename(input_path):
    """
    Prompt user for output filename
    """
    input_name = Path(input_path).stem
    default_output = f"categorized_{input_name}.xlsx"
    
    print(f"\nDefault output filename: {default_output}")
    print("Press Enter to use default, or type a custom filename:")
    
    custom_name = input("Output filename: ").strip()
    
    if custom_name:
        # Ensure .xlsx extension
        if not custom_name.lower().endswith('.xlsx'):
            custom_name += '.xlsx'
        return custom_name
    
    return default_output


def detect_file_type(file_path):
    """
    Detect if file is PDF or CSV
    """
    extension = Path(file_path).suffix.lower()
    
    if extension == '.pdf':
        return 'pdf'
    elif extension == '.csv':
        return 'csv'
    else:
        return None


def process_statement(statement_path, output_path):
    """
    Process the statement using the appropriate script
    """
    file_type = detect_file_type(statement_path)
    
    if file_type == 'pdf':
        print("\n" + "-"*60)
        print("  Detected: PDF Statement (Wamo)")
        print("-"*60)
        script = 'wamo_categorization.py'
    elif file_type == 'csv':
        print("\n" + "-"*60)
        print("  Detected: CSV Statement (Bank of Valletta)")
        print("-"*60)
        script = 'bov_categorization.py'
    else:
        print(f"\nError: Unsupported file type: {Path(statement_path).suffix}")
        print("Supported types: PDF (Wamo) or CSV (Bank of Valletta)")
        return False
    
    # Run the appropriate categorization script
    try:
        result = subprocess.run(
            [sys.executable, script, statement_path, output_path],
            check=True,
            capture_output=False
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"\nError processing statement: {e}")
        return False
    except FileNotFoundError:
        print(f"\nError: Could not find {script}")
        print("Please ensure the script is in the same directory.")
        return False


def main():
    """
    Main function with user interaction
    """
    print("\n")
    print("+" + "="*58 + "+")
    print("|" + " "*58 + "|")
    print("|" + "  BANK STATEMENT CATEGORIZATION TOOL".center(58) + "|")
    print("|" + " "*58 + "|")
    print("|" + "  Supports: Wamo PDF & Bank of Valletta CSV".center(58) + "|")
    print("|" + " "*58 + "|")
    print("+" + "="*58 + "+")
    
    print("\nStep 1: Select your statement file")
    print("-" * 60)
    print("\nChoose how to select your file:")
    print("  1. Browse for file (GUI dialog)")
    print("  2. Enter file path manually")
    print("  3. Exit")
    
    while True:
        choice = input("\nYour choice (1-3): ").strip()
        
        if choice == '1':
            # Use file dialog
            statement_path = select_file()
            if not statement_path:
                print("\nNo file selected. Exiting.")
                return
            break
        elif choice == '2':
            # Manual input
            print("\nEnter the full path to your statement file:")
            statement_path = input("File path: ").strip()
            
            # Remove quotes if present
            statement_path = statement_path.strip('"').strip("'")
            
            if not Path(statement_path).exists():
                print(f"\nError: File not found: {statement_path}")
                retry = input("\nTry again? (y/n): ").strip().lower()
                if retry != 'y':
                    return
                continue
            break
        elif choice == '3':
            print("\nGoodbye!")
            return
        else:
            print("Invalid choice. Please enter 1, 2, or 3.")
    
    print(f"\n[OK] Selected file: {Path(statement_path).name}")
    
    # Detect file type
    file_type = detect_file_type(statement_path)
    if file_type == 'pdf':
        print("File type: PDF (Wamo Statement)")
    elif file_type == 'csv':
        print("File type: CSV (Bank of Valletta Statement)")
    else:
        print(f"\nError: Unsupported file type: {Path(statement_path).suffix}")
        print("   Supported: .pdf (Wamo) or .csv (Bank of Valletta)")
        input("\nPress Enter to exit...")
        return
    
    # Get output filename
    print("\n" + "-" * 60)
    print("Step 2: Choose output filename")
    print("-" * 60)
    
    output_path = get_output_filename(statement_path)
    print(f"\n[OK] Output will be saved as: {output_path}")
    
    # Confirm
    print("\n" + "-" * 60)
    print("Summary")
    print("-" * 60)
    print(f"  Input:  {Path(statement_path).name}")
    print(f"  Type:   {file_type.upper()}")
    print(f"  Output: {output_path}")
    print("-" * 60)
    
    confirm = input("\nProceed with processing? (Y/n): ").strip().lower()
    if confirm and confirm not in ['y', 'yes']:
        print("\nProcessing cancelled.")
        return
    
    # Process the statement
    print("\n" + "="*60)
    print("  PROCESSING STATEMENT")
    print("="*60)
    
    success = process_statement(statement_path, output_path)
    
    if success:
        print("\n" + "="*60)
        print("  SUCCESS!")
        print("="*60)
        print(f"\nCategorized statement saved to: {output_path}")
        print("\nThe Excel file contains:")
        print("  - SOURCE sheet - Raw transaction data")
        print("  - INCOMING sheet - Income transactions with categorization")
        print("  - OUTGOING sheet - Expense transactions with categorization")
        print("\nAll sheets are formatted as Excel tables with:")
        print("  - Month-based color coding")
        print("  - Proper date and currency formatting")
        print("  - Sortable/filterable columns")
        print("\n" + "="*60)
        
        # Ask if user wants to open the file
        open_file = input("\nWould you like to open the output file? (Y/n): ").strip().lower()
        if not open_file or open_file in ['y', 'yes']:
            try:
                if sys.platform == 'win32':
                    os.startfile(output_path)
                elif sys.platform == 'darwin':  # macOS
                    subprocess.run(['open', output_path])
                else:  # linux
                    subprocess.run(['xdg-open', output_path])
                print("Opening file...")
            except Exception as e:
                print(f"Could not open file automatically: {e}")
                print(f"Please open manually: {output_path}")
    else:
        print("\n" + "="*60)
        print("  PROCESSING FAILED")
        print("="*60)
        print("\nPlease check the error messages above.")
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrupted by user. Goodbye!")
        sys.exit(0)
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)

