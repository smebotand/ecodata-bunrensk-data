"""
Project 18 - Hestnestunnelen EDA Script

Exploratory Data Analysis on extracted lab results from p18_hestnestunnelen.py.
Reads the CSV outputs and performs statistical analysis and visualizations.

Input (from p18_hestnestunnelen.py):
- p18_samples.csv
- p18_results.csv
- p18_classifications.csv
- p18_decisions.csv

Output:
- Console summary statistics
- Visualizations (optional)
- EDA report
"""

import sys
from pathlib import Path
import pandas as pd
import numpy as np
from datetime import datetime

# Add parent to path for lib imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from lib.thresholds import TILSTANDSKLASSER, get_tilstandsklasse

# ============================================================
# PROJECT CONFIGURATION
# ============================================================

PROJECT_CODE = '18_hestnestunnelen'
PROJECT_NAME = 'Hestnestunnelen'

BASE_DIR = Path(__file__).parent.parent.parent
DATA_DIR = BASE_DIR / '01_projects' / '18_hestnestunnelen' / 'extracted'


def load_data():
    """Load extracted data from CSV files."""
    print(f"{'='*60}")
    print(f"Loading data for: {PROJECT_NAME}")
    print(f"{'='*60}")
    
    samples_path = DATA_DIR / 'p18_samples.csv'
    results_path = DATA_DIR / 'p18_results.csv'
    classifications_path = DATA_DIR / 'p18_classifications.csv'
    decisions_path = DATA_DIR / 'p18_decisions.csv'
    
    # Check files exist
    for path in [samples_path, results_path, classifications_path, decisions_path]:
        if not path.exists():
            print(f"ERROR: File not found: {path}")
            print("Run p18_hestnestunnelen.py first to extract the data.")
            return None, None, None, None
    
    samples_df = pd.read_csv(samples_path)
    results_df = pd.read_csv(results_path)
    classifications_df = pd.read_csv(classifications_path)
    decisions_df = pd.read_csv(decisions_path)
    
    print(f"\nLoaded:")
    print(f"  - Samples: {len(samples_df)} rows")
    print(f"  - Results: {len(results_df)} rows")
    print(f"  - Classifications: {len(classifications_df)} rows")
    print(f"  - Decisions: {len(decisions_df)} rows")
    
    return samples_df, results_df, classifications_df, decisions_df


def analyze_samples(samples_df):
    """Analyze sample distribution."""
    print(f"\n{'='*60}")
    print("SAMPLE ANALYSIS")
    print(f"{'='*60}")
    
    print(f"\nTotal samples: {len(samples_df)}")
    
    # By location type
    print("\nSamples by location:")
    location_counts = samples_df['location_type'].value_counts()
    for loc, count in location_counts.items():
        print(f"  {loc}: {count}")
    
    # Profile coverage
    if 'profile_start' in samples_df.columns and 'profile_end' in samples_df.columns:
        samples_with_profile = samples_df[
            (samples_df['profile_start'].notna()) & 
            (samples_df['profile_start'] != '')
        ]
        if len(samples_with_profile) > 0:
            print(f"\nProfile coverage:")
            for loc in samples_df['location_type'].unique():
                loc_samples = samples_with_profile[samples_with_profile['location_type'] == loc]
                if len(loc_samples) > 0:
                    try:
                        starts = pd.to_numeric(loc_samples['profile_start'], errors='coerce')
                        ends = pd.to_numeric(loc_samples['profile_end'], errors='coerce')
                        print(f"  {loc}: {starts.min():.0f} - {ends.max():.0f} m")
                    except:
                        pass


def analyze_results(results_df):
    """Analyze chemical results."""
    print(f"\n{'='*60}")
    print("RESULTS ANALYSIS")
    print(f"{'='*60}")
    
    # By analysis type
    print("\nResults by analysis type:")
    analysis_counts = results_df['analysis_type'].value_counts()
    for atype, count in analysis_counts.items():
        print(f"  {atype}: {count}")
    
    # Filter to totalanalyse for main statistics
    total_df = results_df[results_df['analysis_type'] == 'totalanalyse'].copy()
    
    print(f"\n--- Totalanalyse Statistics ---")
    print(f"Total measurements: {len(total_df)}")
    
    # Unique parameters
    params = total_df['parameter'].unique()
    print(f"Unique parameters: {len(params)}")
    
    # Below detection limit
    below_limit = total_df['below_limit'].sum()
    below_pct = (below_limit / len(total_df)) * 100
    print(f"Below detection limit: {below_limit} ({below_pct:.1f}%)")
    
    # Key parameters statistics
    key_params = ['As', 'Pb', 'Cd', 'Cu', 'Cr', 'Hg', 'Ni', 'Zn', 'BaP', 'PAH16']
    
    print(f"\n--- Key Parameter Statistics (totalanalyse, mg/kg) ---")
    print(f"{'Parameter':<12} {'N':>6} {'Min':>10} {'Max':>10} {'Mean':>10} {'Median':>10} {'<LOQ%':>8}")
    print("-" * 70)
    
    for param in key_params:
        param_data = total_df[total_df['parameter'] == param]
        if len(param_data) > 0:
            n = len(param_data)
            values = param_data['value']
            below_pct = (param_data['below_limit'].sum() / n) * 100
            
            # For statistics, use detected values only
            detected = param_data[~param_data['below_limit']]['value']
            if len(detected) > 0:
                print(f"{param:<12} {n:>6} {values.min():>10.4g} {values.max():>10.4g} "
                      f"{detected.mean():>10.4g} {detected.median():>10.4g} {below_pct:>7.1f}%")
            else:
                print(f"{param:<12} {n:>6} {'<LOQ':>10} {'<LOQ':>10} {'N/A':>10} {'N/A':>10} {below_pct:>7.1f}%")


def analyze_against_normverdier(results_df):
    """Compare results against Norwegian environmental guidelines."""
    print(f"\n{'='*60}")
    print("COMPARISON AGAINST TILSTANDSKLASSER (TA-2553/2009)")
    print(f"{'='*60}")
    
    total_df = results_df[results_df['analysis_type'] == 'totalanalyse'].copy()
    
    # Key parameters to analyze (must exist in TILSTANDSKLASSER)
    key_params = ['As', 'Pb', 'Cd', 'Cu', 'Cr_total', 'Hg', 'Ni', 'Zn', 'BaP', 'PAH16']
    
    print(f"\n{'Parameter':<12} {'N':>5} {'TK1':>8} {'TK2':>8} {'TK3':>8} {'TK4':>8} {'TK5':>8}")
    print("-" * 70)
    
    for param in key_params:
        # Map Cr to Cr_total for lookup
        lookup_param = param
        data_param = 'Cr' if param == 'Cr_total' else param
        
        if lookup_param not in TILSTANDSKLASSER:
            continue
            
        limits = TILSTANDSKLASSER[lookup_param]
        param_data = total_df[total_df['parameter'] == data_param]
        
        if len(param_data) > 0:
            n = len(param_data)
            values = param_data['value']
            
            # Count samples in each class
            tk1 = (values <= limits['TK1']).sum()
            tk2 = ((values > limits['TK1']) & (values <= limits['TK2'])).sum()
            tk3 = ((values > limits['TK2']) & (values <= limits['TK3'])).sum()
            tk4 = ((values > limits['TK3']) & (values <= limits['TK4'])).sum()
            tk5 = (values > limits['TK4']).sum()
            
            print(f"{data_param:<12} {n:>5} {tk1:>8} {tk2:>8} {tk3:>8} {tk4:>8} {tk5:>8}")
    
    # Summary
    print("\nNote: TK1 = Meget god, TK2 = God, TK3 = Moderat, TK4 = Dårlig, TK5 = Meget dårlig")


def analyze_classifications(classifications_df, decisions_df):
    """Analyze tilstandsklasse distribution and decisions."""
    print(f"\n{'='*60}")
    print("CLASSIFICATION ANALYSIS")
    print(f"{'='*60}")
    
    # Tilstandsklasse distribution
    print("\nTilstandsklasse distribution:")
    class_counts = classifications_df['tilstandsklasse'].value_counts().sort_index()
    total = len(classifications_df)
    for tk, count in class_counts.items():
        pct = (count / total) * 100
        label = {1: 'Meget god', 2: 'God', 3: 'Moderat', 4: 'Dårlig', 5: 'Meget dårlig'}.get(tk, 'Ukjent')
        print(f"  Klasse {int(tk) if pd.notna(tk) else '?'} ({label}): {count} ({pct:.1f}%)")
    
    # Unclassified
    unclassified = classifications_df['tilstandsklasse'].isna().sum()
    if unclassified > 0:
        print(f"  Ikke klassifisert: {unclassified}")
    
    # Decision distribution
    print("\nDecision distribution:")
    decision_counts = decisions_df['decision'].value_counts()
    for dec, count in decision_counts.items():
        pct = (count / len(decisions_df)) * 100
        print(f"  {dec}: {count} ({pct:.1f}%)")
    
    # Destination distribution
    print("\nDestination distribution:")
    dest_counts = decisions_df['destination'].value_counts()
    for dest, count in dest_counts.items():
        if dest:
            pct = (count / len(decisions_df)) * 100
            print(f"  {dest}: {count} ({pct:.1f}%)")


def analyze_leaching_tests(results_df):
    """Analyze leaching test results."""
    print(f"\n{'='*60}")
    print("LEACHING TEST ANALYSIS")
    print(f"{'='*60}")
    
    ristetest_df = results_df[results_df['analysis_type'] == 'ristetest']
    kolonnetest_df = results_df[results_df['analysis_type'] == 'kolonnetest']
    
    if len(ristetest_df) == 0 and len(kolonnetest_df) == 0:
        print("\nNo leaching test data found.")
        return
    
    if len(ristetest_df) > 0:
        print(f"\n--- Ristetest (L/S=10) ---")
        print(f"Measurements: {len(ristetest_df)}")
        print(f"Parameters: {ristetest_df['parameter'].nunique()}")
        print(f"Samples: {ristetest_df['sample_id'].nunique()}")
        
        print(f"\n{'Parameter':<15} {'Value':>12} {'Unit':>10} {'<LOQ':>6}")
        print("-" * 50)
        for _, row in ristetest_df.iterrows():
            loq_marker = '*' if row['below_limit'] else ''
            print(f"{row['parameter']:<15} {row['value']:>12.4g}{loq_marker} {row['unit']:>10}")
    
    if len(kolonnetest_df) > 0:
        print(f"\n--- Kolonnetest (L/S=0.1) ---")
        print(f"Measurements: {len(kolonnetest_df)}")
        print(f"Parameters: {kolonnetest_df['parameter'].nunique()}")
        
        print(f"\n{'Parameter':<15} {'Value':>12} {'Unit':>10} {'<LOQ':>6}")
        print("-" * 50)
        for _, row in kolonnetest_df.iterrows():
            loq_marker = '*' if row['below_limit'] else ''
            print(f"{row['parameter']:<15} {row['value']:>12.4g}{loq_marker} {row['unit']:>10}")


def analyze_by_location(results_df, samples_df):
    """Analyze results by tunnel location."""
    print(f"\n{'='*60}")
    print("ANALYSIS BY LOCATION")
    print(f"{'='*60}")
    
    # Merge results with samples to get location info
    total_df = results_df[results_df['analysis_type'] == 'totalanalyse'].copy()
    
    # Extract location from sample_id (p18-HL-xxx -> HL)
    total_df['location'] = total_df['sample_id'].str.extract(r'p18-([A-ZØ]+)-')[0]
    
    locations = total_df['location'].unique()
    key_params = ['As', 'Pb', 'Cr', 'Ni', 'Zn']
    
    for param in key_params:
        param_data = total_df[total_df['parameter'] == param]
        if len(param_data) > 0:
            print(f"\n{param} by location (mg/kg):")
            for loc in sorted(locations):
                loc_data = param_data[param_data['location'] == loc]
                if len(loc_data) > 0:
                    detected = loc_data[~loc_data['below_limit']]['value']
                    if len(detected) > 0:
                        print(f"  {loc}: n={len(loc_data)}, mean={detected.mean():.3g}, "
                              f"max={loc_data['value'].max():.3g}")
                    else:
                        print(f"  {loc}: n={len(loc_data)}, all <LOQ")


def run_eda():
    """Main EDA function."""
    # Load data
    samples_df, results_df, classifications_df, decisions_df = load_data()
    
    if samples_df is None:
        return False
    
    # Run analyses
    analyze_samples(samples_df)
    analyze_results(results_df)
    analyze_against_normverdier(results_df)
    analyze_classifications(classifications_df, decisions_df)
    analyze_leaching_tests(results_df)
    analyze_by_location(results_df, samples_df)
    
    # Final summary
    print(f"\n{'='*60}")
    print("EDA COMPLETE")
    print(f"{'='*60}")
    
    return True


if __name__ == '__main__':
    success = run_eda()
    sys.exit(0 if success else 1)
