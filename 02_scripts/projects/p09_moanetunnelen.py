"""
Project 09 - Moanetunnelen Extraction Script

Extracts data from the Innsyn PDF to populate the 5-table data schema:
- samples.csv
- results.csv
- classifications.csv
- decisions.csv
- extraction_summary.csv

Source pages:
- Pages 2-5: SVV notat with sample info, classifications, and decisions
- Pages 6-30: ALS lab reports with detailed chemical analysis results
"""

import sys
from pathlib import Path
import pandas as pd
from datetime import datetime

# Add parent to path for lib imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from lib.pdf_utils import extract_text, extract_pages
from lib.export import save_to_csv
from lib.excel_utils import create_wide_table, save_wide_table_xlsx
from lib.parsers import lab_results_from_als_pdf_with_THC
from lib.qa_workbook import create_qa_workbook

# ============================================================
# PROJECT CONFIGURATION
# ============================================================

CONFIG = {
    'project_code': '09_moanetunnelen',
    'project_name': 'Moanetunnelen',
    'tunnel_name': 'Moanetunnelen',
    'sampler': 'SVV Region Sør',
    'lab': 'ALS Laboratory Group Norway AS',
    'lab_reports': 'N1715906, N1716546',
    'sampling_dates': '2017-09-20, 2017-09-28',
    'pdf_file': 'Innsyn 25-339107-1Svar på innsyn - Innsynskrav i - 16171864-58 med vedlegg (1).pdf',
    'notat_pages': list(range(2, 6)),       # Pages 2-5: SVV notat
    'lab_report_pages': list(range(6, 31)), # Pages 6-30: ALS lab reports
}

BASE_DIR = Path(__file__).parent.parent.parent
INBOX_DIR = BASE_DIR / '00_inbox' / 'SVV' / 'Moanetunnelen'
OUTPUT_DIR = BASE_DIR / '01_projects' / CONFIG['project_code'] / 'extracted'

# ============================================================
# SAMPLE METADATA (extracted from pages 2-5 of the notat)
# ============================================================

SAMPLES = [
    # First sampling 20.09.2017
    {'sample_key': '1', 'sample_id': 'p09-MOA-001', 'sample_date': '2017-09-20', 
     'location_type': 'vegbane', 'profile_start': 9850, 'profile_end': 9950,
     'sample_type': 'bunnrensk', 'lab_reference': 'N00529080', 'remark': ''},
    {'sample_key': '2', 'sample_id': 'p09-MOA-002', 'sample_date': '2017-09-20',
     'location_type': 'vegbane', 'profile_start': 9950, 'profile_end': 10100,
     'sample_type': 'bunnrensk', 'lab_reference': 'N00529081', 'remark': ''},
    {'sample_key': '3', 'sample_id': 'p09-MOA-003', 'sample_date': '2017-09-20',
     'location_type': 'grøft', 'profile_start': 9950, 'profile_end': 10100,
     'sample_type': 'bunnrensk', 'lab_reference': 'N00529082', 'remark': 'Benyttet som sedimentasjonsbasseng'},
    {'sample_key': '4', 'sample_id': 'p09-MOA-004', 'sample_date': '2017-09-20',
     'location_type': 'grøft', 'profile_start': 9850, 'profile_end': 9950,
     'sample_type': 'bunnrensk', 'lab_reference': 'N00529083', 'remark': 'Benyttet som sedimentasjonsbasseng'},
    # Second sampling 28.09.2017
    {'sample_key': 'T1', 'sample_id': 'p09-MOA-T1', 'sample_date': '2017-09-28',
     'location_type': 'vegbane', 'profile_start': 9850, 'profile_end': 9900,
     'sample_type': 'bunnrensk', 'lab_reference': 'N00531434', 'remark': ''},
    {'sample_key': 'T2', 'sample_id': 'p09-MOA-T2', 'sample_date': '2017-09-28',
     'location_type': 'vegbane', 'profile_start': 9900, 'profile_end': 9950,
     'sample_type': 'bunnrensk', 'lab_reference': 'N00531435', 'remark': ''},
    {'sample_key': 'D1', 'sample_id': 'p09-MOA-D1', 'sample_date': '2017-09-28',
     'location_type': 'grøft', 'profile_start': None, 'profile_end': None,
     'sample_type': 'blandprøve', 'lab_reference': 'N00531432', 'remark': 'Benyttet som sedimentasjonsbasseng'},
    {'sample_key': 'D2', 'sample_id': 'p09-MOA-D2', 'sample_date': '2017-09-28',
     'location_type': 'grøft', 'profile_start': None, 'profile_end': None,
     'sample_type': 'blandprøve', 'lab_reference': 'N00531433', 'remark': 'Benyttet som sedimentasjonsbasseng'},
]

# Classifications from Tabell 1 and 2 in the notat (pages 2-3)
# limiting_parameters is a list to support multiple limiting factors
CLASSIFICATIONS = [
    {'sample_key': '1', 'tilstandsklasse': 4, 'limiting_parameters': ['As']},
    {'sample_key': '2', 'tilstandsklasse': 3, 'limiting_parameters': ['Cu']},
    {'sample_key': '3', 'tilstandsklasse': 2, 'limiting_parameters': []},
    {'sample_key': '4', 'tilstandsklasse': 4, 'limiting_parameters': ['THC_C12-C35']},
    {'sample_key': 'T1', 'tilstandsklasse': 3, 'limiting_parameters': ['THC', 'Cu']},
    {'sample_key': 'T2', 'tilstandsklasse': 4, 'limiting_parameters': ['THC']},
    {'sample_key': 'D1', 'tilstandsklasse': 3, 'limiting_parameters': ['THC', 'Cu']},
    {'sample_key': 'D2', 'tilstandsklasse': 3, 'limiting_parameters': ['THC']},
]

# Decision remarks from the notat (pages 2-5)
# Note: All samples go to deponi per project decision
DECISION_REMARKS = [
    {'sample_key': '1', 'decision': 'supplerende prøvetaking',
     'destination': '', 'notes': 'Nye prøver tas for å se på arsen-nivået. Tilstandsklasse IV for arsen'},
    {'sample_key': '2', 'decision': 'gjenbruk',
     'destination': '', 'notes': 'Kan disponeres som planlagt. Lett/moderat forurenset'},
    {'sample_key': '3', 'decision': 'supplerende prøvetaking',
     'destination': '', 'notes': 'Blandet med prøve 4, nye prøver tas'},
    {'sample_key': '4', 'decision': 'deponi',
     'destination': '', 'notes': 'Må kjøres til godkjent deponi. Oljeutslipp er entreprenørens ansvar'},
    {'sample_key': 'T1', 'decision': 'gjenbruk',
     'destination': '', 'notes': 'Kan benyttes i anlegget. Profil 9850-9900'},
    {'sample_key': 'T2', 'decision': 'deponi',
     'destination': '', 'notes': 'Må kjøres til godkjent deponi. Oljeutslipp - profil 9900-9950'},
    {'sample_key': 'D1', 'decision': 'deponi',
     'destination': '', 'notes': 'Kan ikke benyttes i anlegget. Grøftemasser på deponi'},
    {'sample_key': 'D2', 'decision': 'deponi',
     'destination': '', 'notes': 'Kan ikke benyttes i anlegget. Grøftemasser på deponi'},
]

# Mapping sample_key to sample_id
SAMPLE_KEY_TO_ID = {s['sample_key']: s['sample_id'] for s in SAMPLES}


def extract():
    """Main extraction function for Project 09."""
    print(f"{'='*60}")
    print(f"Extracting: {CONFIG['project_name']}")
    print(f"{'='*60}")
    
    # Ensure output directory exists
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    pdf_path = INBOX_DIR / CONFIG['pdf_file']
    
    if not pdf_path.exists():
        print(f"ERROR: PDF file not found: {pdf_path}")
        return False
    
    print(f"\nSource: {CONFIG['pdf_file']}")
    
    # ============================================================
    # EXTRACT RAW TEXT
    # ============================================================
    print("\n[1/6] Extracting raw text...")
    
    try:
        full_text = extract_text(pdf_path)
        text_output = OUTPUT_DIR / 'p09_raw_text.txt'
        text_output.write_text(full_text, encoding='utf-8')
        print(f"  Saved: p09_raw_text.txt")
    except Exception as e:
        print(f"ERROR: {e}")
        return False
    
    # ============================================================
    # EXTRACT NOTAT PAGES (2-5)
    # ============================================================
    notat_pages = CONFIG['notat_pages']
    print(f"\n[2/6] Extracting notat (pages {notat_pages[0]}-{notat_pages[-1]})...")
    
    try:
        notat_text = extract_pages(pdf_path, notat_pages)
        notat_output = OUTPUT_DIR / 'p09_notat_pages_2-5.txt'
        notat_output.write_text(notat_text, encoding='utf-8')
        print(f"  Saved: p09_notat_pages_2-5.txt")
    except Exception as e:
        print(f"  Warning: {e}")
    
    # ============================================================
    # EXTRACT LAB REPORT PAGES (6-30)
    # ============================================================
    lab_pages = CONFIG['lab_report_pages']
    print(f"\n[3/6] Extracting lab reports (pages {lab_pages[0]}-{lab_pages[-1]})...")
    
    try:
        lab_text = extract_pages(pdf_path, lab_pages)
        lab_output = OUTPUT_DIR / 'p09_lab_reports_pages_6-30.txt'
        lab_output.write_text(lab_text, encoding='utf-8')
        print(f"  Saved: p09_lab_reports_pages_6-30.txt")
    except Exception as e:
        print(f"  Warning: {e}")
        lab_text = full_text  # Fallback to full text
    
    # ============================================================
    # BUILD SAMPLES TABLE
    # ============================================================
    print("\n[4/6] Building samples.csv...")
    
    samples_data = []
    for s in SAMPLES:
        samples_data.append({
            'sample_id': s['sample_id'],
            'project_code': CONFIG['project_code'],
            'sample_date': s['sample_date'],
            'location_type': s['location_type'],
            'profile_start': s['profile_start'] if s['profile_start'] else '',
            'profile_end': s['profile_end'] if s['profile_end'] else '',
            'tunnel_name': CONFIG['tunnel_name'],
            'sample_type': s['sample_type'],
            'lab_reference': s['lab_reference'],
            'sampler': CONFIG['sampler'],
            'remark': s.get('remark', ''),
        })
    
    samples_df = pd.DataFrame(samples_data)
    save_to_csv(samples_df, OUTPUT_DIR / 'p09_samples.csv')
    print(f"  Saved: p09_samples.csv ({len(samples_df)} samples)")
    
    # ============================================================
    # PARSE AND BUILD RESULTS TABLE
    # ============================================================
    print("\n[5/6] Parsing lab results and building results.csv...")
    
    results = lab_results_from_als_pdf_with_THC(
        text=lab_text,
        sample_key_to_id=SAMPLE_KEY_TO_ID,
        project_code=CONFIG['project_code']
    )
    print(f"  Parsed {len(results)} result entries from lab reports")
    
    if results:
        results_df = pd.DataFrame(results)
        results_df = results_df[['sample_id', 'parameter', 'parameter_raw', 'value', 
                                  'unit', 'uncertainty', 'below_limit', 'loq', 'analysis_type']]
        save_to_csv(results_df, OUTPUT_DIR / 'p09_results.csv')
        print(f"  Saved: p09_results.csv ({len(results_df)} results)")
        
        # Summary by sample
        sample_counts = results_df.groupby('sample_id').size()
        for sid, count in sample_counts.items():
            print(f"    {sid}: {count} parameters")
        
        # Create wide format table (samples as columns, parameters as rows)
        wide_df = create_wide_table(results_df)
        save_to_csv(wide_df, OUTPUT_DIR / 'p09_results_wide.csv')
        print(f"  Saved: p09_results_wide.csv ({len(wide_df)} parameters x {len(wide_df.columns)-2} samples)")
        
        # Export as formatted Excel table for QA
        save_wide_table_xlsx(wide_df, OUTPUT_DIR / 'p09_results_wide.xlsx')
        print(f"  Saved: p09_results_wide.xlsx (formatted table for QA)")
    else:
        print("  WARNING: No results parsed!")
        results_df = pd.DataFrame()
    
    # ============================================================
    # BUILD CLASSIFICATIONS TABLE
    # ============================================================
    print("\n[6/6] Building classifications.csv and decisions.csv...")
    
    class_data = []
    for c in CLASSIFICATIONS:
        sample_id = SAMPLE_KEY_TO_ID.get(c['sample_key'], f"MOA-{c['sample_key']}")
        # Join lists with semicolon for CSV compatibility
        params = c['limiting_parameters']
        class_data.append({
            'sample_id': sample_id,
            'tilstandsklasse': c['tilstandsklasse'],
            'limiting_parameters': ';'.join(params) if params else '',
            'classification_basis': 'TA-2553/2009',
        })
    
    class_df = pd.DataFrame(class_data)
    save_to_csv(class_df, OUTPUT_DIR / 'p09_classifications.csv')
    print(f"  Saved: p09_classifications.csv ({len(class_df)} classifications)")
    
    # ============================================================
    # BUILD DECISIONS TABLE
    # ============================================================
    dec_data = []
    for d in DECISION_REMARKS:
        sample_id = SAMPLE_KEY_TO_ID.get(d['sample_key'], f"MOA-{d['sample_key']}")
        dec_data.append({
            'sample_id': sample_id,
            'decision': d['decision'],
            'destination': d['destination'],
            'notes': d['notes'],
        })
    
    dec_df = pd.DataFrame(dec_data)
    save_to_csv(dec_df, OUTPUT_DIR / 'p09_decisions.csv')
    print(f"  Saved: p09_decisions.csv ({len(dec_df)} decisions)")
    
    # ============================================================
    # GENERATE EXTRACTION SUMMARY
    # ============================================================
    summary = {
        'project': CONFIG['project_name'],
        'project_code': CONFIG['project_code'],
        'extracted_at': datetime.now().isoformat(),
        'source_file': CONFIG['pdf_file'],
        'notat_pages': f"{notat_pages[0]}-{notat_pages[-1]}",
        'lab_report_pages': f"{lab_pages[0]}-{lab_pages[-1]}",
        'samples_count': len(samples_df),
        'results_count': len(results_df) if len(results) > 0 else 0,
        'classifications_count': len(class_df),
        'decisions_count': len(dec_df),
        'sampling_dates': CONFIG['sampling_dates'],
        'lab': CONFIG['lab'],
        'lab_reports': CONFIG['lab_reports'],
    }
    
    summary_df = pd.DataFrame([summary])
    save_to_csv(summary_df, OUTPUT_DIR / 'p09_extraction_summary.csv')
    
    # ============================================================
    # GENERATE QA WORKBOOK (rename existing with timestamp to preserve notes)
    # ============================================================
    qa_workbook_path = OUTPUT_DIR / 'p09_QA_workbook.xlsx'
    if qa_workbook_path.exists():
        # Rename existing file with timestamp to preserve manual notes
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = OUTPUT_DIR / f'p09_QA_workbook_{timestamp}.xlsx'
        qa_workbook_path.rename(backup_path)
        print(f"  Renamed existing QA workbook to: p09_QA_workbook_{timestamp}.xlsx")
    
    create_qa_workbook(
        qa_workbook_path,
        summary,
        samples_df,
        results_df if len(results) > 0 else pd.DataFrame(),
        class_df,
        dec_df
    )
    print(f"  Saved: p09_QA_workbook.xlsx (for manual QA documentation)")
    
    # ============================================================
    # FINAL REPORT
    # ============================================================
    print(f"\n{'='*60}")
    print("EXTRACTION COMPLETE")
    print(f"{'='*60}")
    print(f"\nOutput files in: {OUTPUT_DIR}")
    print(f"  - p09_samples.csv          ({summary['samples_count']} rows)")
    print(f"  - p09_results.csv          ({summary['results_count']} rows)")
    print(f"  - p09_classifications.csv  ({summary['classifications_count']} rows)")
    print(f"  - p09_decisions.csv        ({summary['decisions_count']} rows)")
    print(f"  - p09_extraction_summary.csv")
    print(f"  - p09_raw_text.txt")
    print(f"  - p09_notat_pages_2-5.txt")
    print(f"  - p09_lab_reports_pages_6-30.txt")
    print(f"{'='*60}")
    
    return True


if __name__ == '__main__':
    success = extract()
    sys.exit(0 if success else 1)
