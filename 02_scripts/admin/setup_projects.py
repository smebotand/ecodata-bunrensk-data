"""
Setup projects from the master Excel list.
Creates folder structure and updates projects.csv
"""

import pandas as pd
from pathlib import Path
import csv

# Read the master list
df = pd.read_excel(r'c:\GIT\ecodata-bunrensk-data\00_inbox\eQ-bunnrnsk-tunnellista.xlsx', header=1)

# Filter out empty rows
df = df.dropna(subset=['Intern ref'])

# Base paths
base = Path(r'c:\GIT\ecodata-bunrensk-data')
projects_dir = base / '01_projects'
metadata_dir = base / 'metadata'

# Build project list
projects = []
for _, row in df.iterrows():
    intern_ref = str(row['Intern ref']).replace('eQ-bunnrensk-', '')
    prosjekt = str(row['Prosjekt'])
    tunnel_type = str(row.get('Tunneltype', 'nan'))
    byggherre = str(row.get('Byggherre', 'nan'))
    entreprenor = str(row.get('Entreprenør', 'nan'))
    kontakt = str(row.get('Kontaktperson', 'nan'))
    lead = str(row.get('Lead', 'nan'))
    
    # Create folder name from intern ref (lowercase, clean)
    folder_name = intern_ref.lower().replace(' ', '-')
    
    projects.append({
        'project_code': intern_ref,
        'folder_name': folder_name,
        'project_name': prosjekt,
        'description': '',
        'tunnel_type': tunnel_type if tunnel_type != 'nan' else '',
        'byggherre': byggherre if byggherre != 'nan' else '',
        'entreprenor': entreprenor if entreprenor != 'nan' else '',
        'kontaktperson': kontakt if kontakt != 'nan' else '',
        'lead': lead if lead != 'nan' else '',
        'start_date': '',
        'end_date': '',
        'status': 'active',
        'tags': '',
    })

print(f'Found {len(projects)} projects\n')

# Create folder structure
print('=== Creating project folders ===')
subfolders = ['raw', 'extracted', 'samples', 'correspondence']

for p in projects:
    project_path = projects_dir / p['folder_name']
    for subfolder in subfolders:
        (project_path / subfolder).mkdir(parents=True, exist_ok=True)
    print(f"  Created: {p['folder_name']}/")

# Write projects.csv
print('\n=== Updating projects.csv ===')
csv_path = metadata_dir / 'projects.csv'

fieldnames = ['project_code', 'project_name', 'description', 'tunnel_type', 
              'byggherre', 'entreprenor', 'kontaktperson', 'lead',
              'start_date', 'end_date', 'status', 'tags']

with open(csv_path, 'w', encoding='utf-8', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    for p in projects:
        row = {k: p.get(k, '') for k in fieldnames}
        writer.writerow(row)

print(f'  Written {len(projects)} projects to {csv_path}')

print('\n=== Done! ===')
