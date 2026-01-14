"""
Project 01: E18-E102 (E18-Vestkorridoren E102)
Extraction script for bunnrensk data.

Data sources in inbox:
- SVV/E18 Vestkorridoren E102/
"""

import sys
from pathlib import Path

# Add parent to path for lib imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from lib.excel_utils import read_lab_report, list_sheets, extract_all_sheets
from lib.pdf_utils import extract_text, tables_to_dataframes
from lib.chemistry import normalize_results_df, classify_contamination
from lib.export import save_to_csv, save_to_excel, generate_extraction_report

# Project configuration
PROJECT_CODE = '01_E18-E102'
PROJECT_NAME = 'E18-Vestkorridoren (E102)'

# Paths
BASE_DIR = Path(__file__).parent.parent.parent
INBOX_DIR = BASE_DIR / '00_inbox' / 'SVV' / 'E18 Vestkorridoren E102'
OUTPUT_DIR = BASE_DIR / '01_projects' / '01_e18-e102' / 'extracted'


def get_source_files():
    """List all source files for this project."""
    files = {
        'excel': list(INBOX_DIR.glob('**/*.xlsx')),
        'pdf': list(INBOX_DIR.glob('**/*.pdf')),
    }
    return files


def extract_excel_files():
    """Extract data from Excel files."""
    results = []
    
    for xlsx_file in INBOX_DIR.glob('**/*.xlsx'):
        print(f"\nProcessing: {xlsx_file.name}")
        
        try:
            # List sheets
            sheets = list_sheets(xlsx_file)
            print(f"  Sheets: {sheets}")
            
            # Extract each sheet
            for sheet in sheets:
                df = read_lab_report(xlsx_file, sheet_name=sheet)
                
                if not df.empty:
                    # Normalize if it looks like chemical data
                    if any(col.lower() in ['resultat', 'result', 'verdi'] 
                           for col in df.columns):
                        df = normalize_results_df(df)
                    
                    # Save
                    output_name = f"{xlsx_file.stem}_{sheet}.csv"
                    output_path = OUTPUT_DIR / output_name
                    save_to_csv(df, output_path)
                    
                    results.append({
                        'filename': xlsx_file.name,
                        'sheet': sheet,
                        'rows': len(df),
                        'output': output_name,
                        'success': True
                    })
                    
        except Exception as e:
            print(f"  Error: {e}")
            results.append({
                'filename': xlsx_file.name,
                'error': str(e),
                'success': False
            })
    
    return results


def extract_pdf_files():
    """Extract text and tables from PDF files."""
    results = []
    
    for pdf_file in INBOX_DIR.glob('**/*.pdf'):
        print(f"\nProcessing: {pdf_file.name}")
        
        try:
            # Extract text
            text = extract_text(pdf_file)
            text_path = OUTPUT_DIR / f"{pdf_file.stem}.txt"
            text_path.write_text(text, encoding='utf-8')
            
            # Extract tables
            tables = tables_to_dataframes(pdf_file)
            for i, df in enumerate(tables, 1):
                if not df.empty:
                    table_path = OUTPUT_DIR / f"{pdf_file.stem}_table{i}.csv"
                    save_to_csv(df, table_path)
            
            results.append({
                'filename': pdf_file.name,
                'tables': len(tables),
                'success': True
            })
            
        except Exception as e:
            print(f"  Error: {e}")
            results.append({
                'filename': pdf_file.name,
                'error': str(e),
                'success': False
            })
    
    return results


def run():
    """Main extraction function."""
    print(f"=" * 60)
    print(f"Extracting: {PROJECT_CODE} - {PROJECT_NAME}")
    print(f"=" * 60)
    
    # Ensure output directory exists
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # List source files
    files = get_source_files()
    print(f"\nSource files:")
    print(f"  Excel: {len(files['excel'])}")
    print(f"  PDF: {len(files['pdf'])}")
    
    # Extract
    all_results = []
    
    if files['excel']:
        print("\n--- Extracting Excel files ---")
        all_results.extend(extract_excel_files())
    
    if files['pdf']:
        print("\n--- Extracting PDF files ---")
        all_results.extend(extract_pdf_files())
    
    # Generate report
    generate_extraction_report(PROJECT_CODE, all_results, OUTPUT_DIR)
    
    print(f"\n{'=' * 60}")
    print(f"Extraction complete. Output in: {OUTPUT_DIR}")
    print(f"{'=' * 60}")


if __name__ == '__main__':
    run()
