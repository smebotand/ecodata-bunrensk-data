import pandas as pd
from pathlib import Path
import sys
sys.path.append(str((Path(__file__).parent.parent / 'lib').resolve()))
from conversions import convert_units
import schema

# Define the root projects directory
PROJECTS_DIR = Path(__file__).parent.parent.parent / '01_projects'
OUTPUT_DIR = Path(__file__).parent.parent / 'analysis'
OUTPUT_DIR.mkdir(exist_ok=True)
MERGED_FILE = OUTPUT_DIR / 'merged_results.csv'

# Find all *results.csv files recursively in 01_projects
results_files = list(PROJECTS_DIR.glob('**/*results.csv'))

# Read and concatenate all results
frames = []
for file in results_files:
    try:
        df = pd.read_csv(file)
        df['source_file'] = str(file.relative_to(PROJECTS_DIR))
        frames.append(df)
    except Exception as e:
        print(f"Could not read {file}: {e}")

if frames:
    merged_df = pd.concat(frames, ignore_index=True)

    # Normalize units
    def normalize_value(row):
        unit = str(row.get('unit', '')).strip().lower()
        value = row.get('value')
        # Solids
        if unit in {u.lower() for u in schema.UNIT_TYPES_SOLID}:
            return convert_units(value, row['unit'], 'mg/kg')
        # Liquids
        elif unit in {u.lower() for u in schema.UNIT_TYPES_LIQUID}:
            return convert_units(value, row['unit'], 'mg/l')
        # Other: keep as is
        else:
            return value
    def normalize_unit(row):
        unit = str(row.get('unit', '')).strip().lower()
        # Solids
        if unit in {u.lower() for u in schema.UNIT_TYPES_SOLID}:
            return 'mg/kg'
        # Liquids
        elif unit in {u.lower() for u in schema.UNIT_TYPES_LIQUID}:
            return 'mg/l'
        # Other: keep as is
        else:
            return row['unit']
    if 'value' in merged_df.columns and 'unit' in merged_df.columns:
        merged_df['value'] = merged_df.apply(normalize_value, axis=1)
        merged_df['unit'] = merged_df.apply(normalize_unit, axis=1)

    merged_df.to_csv(MERGED_FILE, index=False)
    print(f"Merged {len(results_files)} files into {MERGED_FILE}")

    # Also dump as Excel table
    xlsx_file = OUTPUT_DIR / 'merged_results_table.xlsx'
    with pd.ExcelWriter(xlsx_file, engine='xlsxwriter') as writer:
        merged_df.to_excel(writer, index=False, sheet_name='Results')
        workbook  = writer.book
        worksheet = writer.sheets['Results']
        (max_row, max_col) = merged_df.shape
        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'name': 'ResultsTable',
            'columns': [{'header': col} for col in merged_df.columns]
        })
    print(f"Also saved as Excel table: {xlsx_file}")
else:
    print("No results.csv files found or all failed to read.")
