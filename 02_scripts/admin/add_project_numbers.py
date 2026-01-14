"""
Add project number prefix to all project folders and update projects.csv
"""

import pandas as pd
from pathlib import Path
import csv
import shutil

base = Path(r'c:\GIT\ecodata-bunrensk-data')
projects_dir = base / '01_projects'
metadata_dir = base / 'metadata'

# Read current projects.csv
df = pd.read_csv(metadata_dir / 'projects.csv')

print(f'Found {len(df)} projects\n')

# Mapping of old folder names to new folder names with prefix
folder_mapping = []

for idx, row in df.iterrows():
    project_num = idx + 1  # 1-based numbering
    prefix = f"{project_num:02d}"  # Zero-padded to 2 digits
    
    old_code = row['project_code']
    new_code = f"{prefix}_{old_code}"
    
    # Find old folder (lowercase version of project_code)
    old_folder = old_code.lower().replace(' ', '-')
    new_folder = f"{prefix}_{old_folder}"
    
    folder_mapping.append({
        'old_code': old_code,
        'new_code': new_code,
        'old_folder': old_folder,
        'new_folder': new_folder,
    })
    
    # Update dataframe
    df.at[idx, 'project_code'] = new_code

# Rename folders
print('=== Renaming project folders ===')
for m in folder_mapping:
    old_path = projects_dir / m['old_folder']
    new_path = projects_dir / m['new_folder']
    
    if old_path.exists():
        old_path.rename(new_path)
        print(f"  {m['old_folder']} -> {m['new_folder']}")
    else:
        print(f"  [SKIP] {m['old_folder']} not found")

# Write updated projects.csv
print('\n=== Updating projects.csv ===')
df.to_csv(metadata_dir / 'projects.csv', index=False)
print(f'  Updated {len(df)} project codes')

print('\n=== Done! ===')
