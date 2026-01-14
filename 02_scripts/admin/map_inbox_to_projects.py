"""
Map inbox folders to projects without moving files.
Creates a mapping index that links inbox source locations to project codes.
The inbox folder remains untouched (it's a SharePoint mirror).
"""

import pandas as pd
from pathlib import Path
import csv
import os

base = Path(r'c:\GIT\ecodata-bunrensk-data')
inbox_dir = base / '00_inbox'
projects_dir = base / '01_projects'
metadata_dir = base / 'metadata'

# Read current projects.csv for reference
projects_df = pd.read_csv(metadata_dir / 'projects.csv')

# Manual mapping of inbox folder structure to project codes
# Based on the inbox folder names matching project names
INBOX_TO_PROJECT_MAP = {
    # SVV projects
    'SVV/E18 Vestkorridoren E102': '01_E18-E102',
    'SVV/E18 Vestkorridoren/E102': '01_E18-E102',
    'SVV/E18 Vestkorridoren E103': '02_E18-E103',
    'SVV/E18 Vestkorridoren/E103': '02_E18-E103',
    'SVV/Ålesund E136 Lerstadtunnelen': '03_E136-Ålesund',
    'SVV/Sotrasambandet': '04_Sotrasambandet',
    'SVV/E16 Bjørum Skaret': '05_Bjørum-Skaret',
    'SVV/E10 Hålogalandsvegen': '06_E10-Hålogalandsveien',
    'SVV/E16 Bjørnegårdstunnelen': '07_E16-Bjørnegårdtunnelen',
    'SVV/Svegårdsjønn-Rådal': '08_Svegårdsjønn-Rådal',
    'SVV/Moanetunnelen': '09_Moanetunnelen',
    'SVV/ROGFAST': '10_ROGFAST',
    
    # Oslo VAV projects
    'Oslo VAV/E5 Råvannstunnelen': '11_OsloVAV-e5-råvannstunnelen',
    'Oslo VAV/E6 Rentvannstunnel': '12_OsloVAV-e6',
    
    # Bane NOR projects
    'BaneNOR/Arna - Fløen': '13_Arna-Fløen',
    'BaneNOR/Bane NOR Moss': '14_sms-moss',
    'BaneNOR/Jernbanetrase med dobbeltspor Farriseidet - Porsgrunn': '15_Farriseidet-Porsgrunn',
    'BaneNOR/Drammenstunnelen': '16_Drammenstunnelen',
    'BaneNOR/Blixtunnelen': '17_Blixtunnelen',
    'BaneNOR/Hestnestunnelen': '18_Hestnestunnelen',
    'BaneNOR/Bane NOR Minnesund': '18_Hestnestunnelen',  # Check if this is correct
    
    # Sporveien projects
    'Sporveien/Fornebubanen': '19_Fornebu-K2B',  # May need to split K2B/K2C
    'Sporveien/Lørenbanen': '21_Lørenbanen',
    
    # NyeVeier projects
    'NyeVeier/E18 - Langangen - Rugtvedt': '22_E18 - Langangen - Rugtvedt',
    'NyeVeier/E39 Lyngdal - Lyngdal + E39 Kristiansand - Mandal': '23_E39 Lyngdal – Lyngdal',  # Combined folder
    'NyeVeier/E6 - Storhove - Øyer': '25_E6-Storhove-Øyer',
    'NyeVeier/E6 Kvænangsfjellet + E6 Ranheim – Værnes + E6 Kvithammer – Åsen': '26_E6 Kvænangsfjellet',  # Combined
}


def scan_inbox():
    """Scan inbox and build a list of all files with their paths."""
    files = []
    for root, dirs, filenames in os.walk(inbox_dir):
        # Skip hidden files/folders
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        
        rel_root = Path(root).relative_to(inbox_dir)
        
        for filename in filenames:
            if filename.startswith('.'):
                continue
            
            full_path = Path(root) / filename
            rel_path = rel_root / filename
            
            files.append({
                'inbox_path': str(rel_path),
                'filename': filename,
                'extension': Path(filename).suffix.lower(),
                'size_kb': full_path.stat().st_size // 1024,
                'parent_folder': str(rel_root),
            })
    
    return files


def find_project_for_path(path_str):
    """Find the project code that matches an inbox path."""
    # Try to match against our mapping
    for inbox_pattern, project_code in INBOX_TO_PROJECT_MAP.items():
        if path_str.startswith(inbox_pattern):
            return project_code
    
    # Try to match by folder name heuristics
    path_lower = path_str.lower()
    
    if 'e102' in path_lower and ('e18' in path_lower or 'vestkorridoren' in path_lower):
        return '01_E18-E102'
    if 'e103' in path_lower and ('e18' in path_lower or 'vestkorridoren' in path_lower):
        return '02_E18-E103'
    if 'lerstad' in path_lower or 'ålesund' in path_lower:
        return '03_E136-Ålesund'
    if 'sotra' in path_lower:
        return '04_Sotrasambandet'
    if 'bjørum' in path_lower or 'skaret' in path_lower:
        return '05_Bjørum-Skaret'
    if 'hålogaland' in path_lower:
        return '06_E10-Hålogalandsveien'
    if 'bjørnegård' in path_lower:
        return '07_E16-Bjørnegårdtunnelen'
    if 'svegård' in path_lower or 'rådal' in path_lower:
        return '08_Svegårdsjønn-Rådal'
    if 'moane' in path_lower:
        return '09_Moanetunnelen'
    if 'rogfast' in path_lower:
        return '10_ROGFAST'
    if 'råvann' in path_lower:
        return '11_OsloVAV-e5-råvannstunnelen'
    if 'rentvann' in path_lower:
        return '12_OsloVAV-e6'
    if 'arna' in path_lower and 'fløen' in path_lower:
        return '13_Arna-Fløen'
    if 'moss' in path_lower and 'bane' in path_lower.replace('\\', '/').split('/')[0]:
        return '14_sms-moss'
    if 'farriseidet' in path_lower or 'porsgrunn' in path_lower:
        return '15_Farriseidet-Porsgrunn'
    if 'drammen' in path_lower:
        return '16_Drammenstunnelen'
    if 'blix' in path_lower:
        return '17_Blixtunnelen'
    if 'hestnes' in path_lower:
        return '18_Hestnestunnelen'
    if 'fornebu' in path_lower:
        return '19_Fornebu-K2B'
    if 'løren' in path_lower:
        return '21_Lørenbanen'
    if 'langangen' in path_lower or 'rugtvedt' in path_lower:
        return '22_E18 - Langangen - Rugtvedt'
    if 'lyngdal' in path_lower:
        return '23_E39 Lyngdal – Lyngdal'
    if 'kristiansand' in path_lower or 'mandal' in path_lower:
        return '24_E39 Kristiansand – Mandal'
    if 'storhove' in path_lower or 'øyer' in path_lower:
        return '25_E6-Storhove-Øyer'
    if 'kvænang' in path_lower:
        return '26_E6 Kvænangsfjellet'
    if 'ranheim' in path_lower or 'værnes' in path_lower:
        return '27_E6 Ranheim – Værnes'
    if 'kvithammer' in path_lower or 'åsen' in path_lower:
        return '28_E6 Kvithammer – Åsen'
    
    return 'UNMAPPED'


def main():
    print('=== Scanning inbox folder ===')
    files = scan_inbox()
    print(f'Found {len(files)} files\n')
    
    # Map each file to a project
    print('=== Mapping files to projects ===')
    for f in files:
        f['project_code'] = find_project_for_path(f['parent_folder'])
    
    # Count by project
    from collections import Counter
    project_counts = Counter(f['project_code'] for f in files)
    
    print('\nFiles per project:')
    for project, count in sorted(project_counts.items()):
        print(f'  {project}: {count} files')
    
    # Write mapping to CSV
    mapping_path = metadata_dir / 'inbox_mapping.csv'
    
    fieldnames = ['project_code', 'inbox_path', 'filename', 'extension', 'size_kb', 'parent_folder']
    
    with open(mapping_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for file in sorted(files, key=lambda x: (x['project_code'], x['inbox_path'])):
            writer.writerow(file)
    
    print(f'\n=== Written mapping to {mapping_path} ===')
    
    # Summary of unmapped
    unmapped = [f for f in files if f['project_code'] == 'UNMAPPED']
    if unmapped:
        print(f'\nWARNING: {len(unmapped)} files could not be mapped:')
        unmapped_folders = set(f['parent_folder'] for f in unmapped)
        for folder in sorted(unmapped_folders):
            print(f'  - {folder}')


if __name__ == '__main__':
    main()
