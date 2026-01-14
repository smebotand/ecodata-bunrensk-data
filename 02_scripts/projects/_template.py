"""
Project Template - Copy this file to create a new project extraction script.

Usage:
    1. Copy this file to projects/pXX_project_name.py
    2. Update PROJECT_CODE, PROJECT_NAME, and INBOX_DIR
    3. Customize extraction logic as needed
    4. Run with: python run_extraction.py XX
"""

import sys
from pathlib import Path

# Add parent to path for lib imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from lib.excel_utils import read_lab_report, list_sheets
from lib.pdf_utils import extract_text, tables_to_dataframes
from lib.chemistry import normalize_results_df
from lib.export import save_to_csv, generate_extraction_report

# ============================================================
# PROJECT CONFIGURATION - Update these for each project
# ============================================================

PROJECT_CODE = 'XX_ProjectCode'
PROJECT_NAME = 'Project Name'

# Paths - Update INBOX_DIR to match the project's inbox location
BASE_DIR = Path(__file__).parent.parent.parent
INBOX_DIR = BASE_DIR / '00_inbox' / 'ORG' / 'ProjectFolder'  # <-- UPDATE THIS
OUTPUT_DIR = BASE_DIR / '01_projects' / 'xx_project-name' / 'extracted'  # <-- UPDATE THIS

# ============================================================


def get_source_files():
    """List all source files for this project."""
    if not INBOX_DIR.exists():
        print(f"Warning: Inbox directory not found: {INBOX_DIR}")
        return {'excel': [], 'pdf': []}
    
    return {
        'excel': list(INBOX_DIR.glob('**/*.xlsx')) + list(INBOX_DIR.glob('**/*.xls')),
        'pdf': list(INBOX_DIR.glob('**/*.pdf')),
    }


def extract_excel_files():
    """Extract data from Excel files. Customize per project needs."""
    results = []
    
    for xlsx_file in INBOX_DIR.glob('**/*.xlsx'):
        print(f"\nProcessing: {xlsx_file.name}")
        
        try:
            sheets = list_sheets(xlsx_file)
            
            for sheet in sheets:
                df = read_lab_report(xlsx_file, sheet_name=sheet)
                
                if not df.empty:
                    # Project-specific processing can go here
                    # df = normalize_results_df(df)
                    
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
    """Extract text from PDF files."""
    results = []
    
    for pdf_file in INBOX_DIR.glob('**/*.pdf'):
        print(f"\nProcessing: {pdf_file.name}")
        
        try:
            text = extract_text(pdf_file)
            text_path = OUTPUT_DIR / f"{pdf_file.stem}.txt"
            text_path.write_text(text, encoding='utf-8')
            
            results.append({
                'filename': pdf_file.name,
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
    
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    files = get_source_files()
    print(f"\nSource files:")
    print(f"  Excel: {len(files['excel'])}")
    print(f"  PDF: {len(files['pdf'])}")
    
    all_results = []
    
    if files['excel']:
        print("\n--- Extracting Excel files ---")
        all_results.extend(extract_excel_files())
    
    if files['pdf']:
        print("\n--- Extracting PDF files ---")
        all_results.extend(extract_pdf_files())
    
    generate_extraction_report(PROJECT_CODE, all_results, OUTPUT_DIR)
    
    print(f"\n{'=' * 60}")
    print(f"Extraction complete. Output in: {OUTPUT_DIR}")


if __name__ == '__main__':
    run()
