import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib
matplotlib.use('Agg')
from lib.chemistry import PARAMETER_ALIASES, normalize_parameter_name

# Load merged results
data_path = Path(__file__).parent / 'merged_results.csv'
df = pd.read_csv(data_path)

import os
# Load project names from metadata/projects.csv (one level above project root)
projects_meta_path = Path(__file__).resolve().parents[2] / 'metadata' / 'projects.csv'
projects_meta = pd.read_csv(projects_meta_path)
project_code_to_name = dict(zip(projects_meta['project_code'].str.lower(), projects_meta['project_name']))

# Extract project code from sample_id (prefix before first '_', '-', or similar)
import re
def extract_project_code(sample_id):
    if pd.isna(sample_id):
        return ''
    # Accept both dash and underscore as separator, and allow for pXX_ or XX_ or XX-
    m = re.match(r'([0-9]{2}_[A-Za-z0-9\-]+|[0-9]{2}-[A-Za-z0-9\-]+|[0-9]{2}_[A-Za-z0-9]+)', str(sample_id))
    if m:
        return m.group(1).replace('-', '_').upper()
    # Fallback: try up to first dash or underscore
    return str(sample_id).split('-', 1)[0].split('_', 1)[0]


df['project_code'] = df['sample_id'].apply(extract_project_code)
df['project_code'] = df['project_code'].str.upper()
project_code_to_name_upper = {k.upper(): v for k, v in project_code_to_name.items()}
df['project_name'] = df['project_code'].map(project_code_to_name_upper)

# Remove 'p18' from project names if present
df['project_name'] = df['project_name'].replace('p18', '', regex=True)

# Assign unique numbers to each project for labeling
project_name_counts = {}
def numbered_label(name):
    if pd.isna(name):
        return ''
    base = str(name).strip()
    count = project_name_counts.get(base, 0) + 1
    project_name_counts[base] = count
    return f"{base} ({count})"

# Generate numbered labels for each project occurrence in the plotting order
df = df.copy()  # Avoid SettingWithCopyWarning
df = df.sort_values(['project_name', 'sample_id'])
df['project_label'] = ''
project_name_counts = {}
for idx, row in df.iterrows():
    label = numbered_label(row['project_name'])
    df.at[idx, 'project_label'] = label

# Extract project from source_file (assumes format: 'XX_project/...')
df['project'] = df['source_file'].str.split('/', n=1).str[0]

# Groupings based on chemistry.py comments
PARAM_GROUPS = {
    'Metals': [
        'As', 'Pb', 'Cd', 'Cu', 'Cr', 'Cr_VI', 'Hg', 'Ni', 'Zn', 'Fe', 'Mn', 'Ba', 'V', 'Co', 'Mo', 'Sb', 'Se', 'Sn', 'Tl', 'Al', 'Be', 'Li', 'Ag', 'W'
    ],
    'PAH': [
        'PAH16', 'Naftalen', 'Acenaftylen', 'Acenaften', 'Fluoren', 'Fenantren', 'Antracen', 'Fluoranten', 'Pyren', 'Benzo(a)antracen', 'Krysen', 'Benzo(b)fluoranten', 'Benzo(k)fluoranten', 'BaP', 'Dibenzo(a,h)antracen', 'Benzo(ghi)perylen', 'Indeno(1,2,3-cd)pyren'
    ],
    'PCB': [
        'PCB7', 'PCB28', 'PCB52', 'PCB101', 'PCB118', 'PCB138', 'PCB153', 'PCB180'
    ],
    'BTEX': [
        'BTEX', 'Benzen', 'Toluen', 'Etylbenzen', 'Xylen'
    ],
    'THC': [
        'THC C5-C6', 'THC C6-C8', 'THC C8-C10', 'THC C10-C12', 'THC C12-C16', 'THC C16-C35', 'THC C12-C35', 'THC C35-C40', 'THC C10-C40', 'THC C5-C35', 'TPH', 'THC'
    ],
    'Alifater': [
        'Alifater C5-C6', 'Alifater C6-C8', 'Alifater C8-C10', 'Alifater C10-C12', 'Alifater C12-C35'
    ],
    'THC+Alifater': [
        'THC', 'TPH', 'THC C5-C6', 'THC C6-C8', 'THC C8-C10', 'THC C10-C12', 'THC C12-C16', 'THC C16-C35', 'THC C12-C35', 'THC C35-C40', 'THC C10-C40',
        'Alifater C5-C6', 'Alifater C6-C8', 'Alifater C8-C10', 'Alifater C10-C12', 'Alifater C12-C35'
    ],
    'PFAS': [
        'PFAS', 'PFOS', 'PFOA'
    ],
    'Physical': [
        'DryMatter', 'TOC', 'DOC', 'LOI', 'pH', 'WaterContent'
    ]
}


# Normalize parameter column
df['parameter_norm'] = df['parameter'].apply(normalize_parameter_name)

# Create output directory for plots
plot_dir = Path(__file__).parent / 'boxplots'
plot_dir.mkdir(exist_ok=True)

# Filter for totalanalyse only
if 'analysis_type' in df.columns:
    totalanalyse_df = df[df['analysis_type'] == 'totalanalyse']
else:
    totalanalyse_df = df.copy()
print(f"totalanalyse_df rows: {len(totalanalyse_df)}")

# Plot boxplots for each group (all parameters in group in one plot, not split by project)
for group, params in PARAM_GROUPS.items():

    group_df = totalanalyse_df[totalanalyse_df['parameter_norm'].isin(params)]
    print(f"{group}: {len(group_df)} rows")
    if group == "BTEX":
        for param in params:
            param_df = group_df[group_df['parameter_norm'] == param]
            print(f"  {param}: {len(param_df)} values, unique: {param_df['value'].unique()}")
    if group_df.empty:
        continue

    # --- Add normverdi/TK1 lines ---
    import importlib.util
    import sys as _sys
    thresholds_path = str((Path(__file__).parent.parent / 'lib' / 'thresholds.py').resolve())
    spec = importlib.util.spec_from_file_location('thresholds', thresholds_path)
    thresholds = importlib.util.module_from_spec(spec)
    _sys.modules['thresholds'] = thresholds
    spec.loader.exec_module(thresholds)
    # Map parameter_norm to TK1/normverdi
    normverdi_map = {}
    for p in params:
        # Try TILSTANDSKLASSER first
        tk = thresholds.TILSTANDSKLASSER.get(p)
        if tk and tk.get('TK1') is not None:
            normverdi_map[p] = tk['TK1']
        else:
            # Fallback to NORMVERDIER
            nv = thresholds.NORMVERDIER.get(p)
            if nv is not None:
                normverdi_map[p] = nv

    # Standard boxplot
    plt.figure(figsize=(12, 6))
    ax = sns.boxplot(x='parameter_norm', y='value', data=group_df, showfliers=True, fliersize=5, flierprops={'marker': 'o', 'color': 'red', 'alpha': 0.7})
    # Draw normverdi/TK1 lines for each parameter using ax.hlines for correct alignment
    xtick_labels = [tick.get_text() for tick in ax.get_xticklabels()]
    for i, pname in enumerate(xtick_labels):
        if pname in normverdi_map and normverdi_map[pname] is not None:
            # hlines: y, xmin, xmax (x in data coordinates, so box centers are at i)
            ax.hlines(normverdi_map[pname], i - 0.4, i + 0.4, color='orange', linestyle='--', label='Normverdi/TK1' if i==0 else None, linewidth=2)
    plt.ylabel('Value')
    plt.xlabel('Parameter')
    plt.title(f'{group} (totalanalyse)')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    # Only show legend once
    handles, labels = ax.get_legend_handles_labels()
    if any(l == 'Normverdi/TK1' for l in labels):
        ax.legend()
    plot_file = plot_dir / f"boxplot_{group}_all.png"
    plt.savefig(plot_file)
    plt.close()
    print(f"Saved: {plot_file}")

    # Log-value plot
    plt.figure(figsize=(12, 6))
    ax = sns.boxplot(x='parameter_norm', y='value', data=group_df, showfliers=True, fliersize=5, flierprops={'marker': 'o', 'color': 'red', 'alpha': 0.7})
    plt.yscale('log')
    # Draw normverdi/TK1 lines for each parameter (log scale)
    for i, param in enumerate(ax.get_xticklabels()):
        pname = param.get_text()
        if pname in normverdi_map and normverdi_map[pname] is not None:
            ax.axhline(normverdi_map[pname], color='orange', linestyle='--', xmin=(i+0.1)/len(params), xmax=(i+0.9)/len(params), label='Normverdi/TK1' if i==0 else None)
    plt.ylabel('Log(Value)')
    plt.xlabel('Parameter')
    plt.title(f'{group} (totalanalyse, log scale)')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    handles, labels = ax.get_legend_handles_labels()
    if any(l == 'Normverdi/TK1' for l in labels):
        ax.legend()
    plot_file_log = plot_dir / f"boxplot_{group}_all_log.png"
    plt.savefig(plot_file_log)
    plt.close()
    print(f"Saved: {plot_file_log}")

    # Distribution plots for BTEX group
    if group == "BTEX":
        for param in params:
            param_df = group_df[group_df['parameter_norm'] == param]
            if param_df.empty:
                continue
            plt.figure(figsize=(8, 4))
            sns.histplot(param_df['value'], bins=30, kde=True, color='blue')
            plt.title(f"Distribution of {param} (totalanalyse)")
            plt.xlabel('Value')
            plt.ylabel('Count')
            plt.tight_layout()
            dist_file = plot_dir / f"distplot_{param}_all.png"
            plt.savefig(dist_file)
            plt.close()
            print(f"Saved: {dist_file}")
            # Log scale version
            plt.figure(figsize=(8, 4))
            sns.histplot(param_df['value'], bins=30, kde=True, color='blue')
            plt.yscale('log')
            plt.title(f"Distribution of {param} (totalanalyse, log scale)")
            plt.xlabel('Value')
            plt.ylabel('Count (log)')
            plt.tight_layout()
            dist_file_log = plot_dir / f"distplot_{param}_all_log.png"
            plt.savefig(dist_file_log)
            plt.close()
            print(f"Saved: {dist_file_log}")

# Ensure all results have the same unit for BTEX parameters
unit_map = {}
for param in PARAM_GROUPS.get('BTEX', []):
    units = group_df[group_df['parameter_norm'] == param]['unit'].unique()
    if len(units) > 1:
        print(f"Warning: Multiple units for {param}: {units}")
    if len(units) > 0:
        unit_map[param] = units[0]
    else:
        unit_map[param] = None
# Optionally, convert units here if needed (add conversion logic if required)

