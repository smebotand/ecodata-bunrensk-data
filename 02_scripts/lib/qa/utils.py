"""
QA utility functions for data validation.

Functions for checking data quality, finding violations, and detecting duplicates.
"""

import pandas as pd

from lib.schema import (
    SAMPLE_TYPES, LOCATION_TYPES, ANALYSIS_TYPES, 
    DECISION_TYPES, UNIT_TYPES
)


def find_schema_violations(samples_df: pd.DataFrame, results_df: pd.DataFrame,
                           dec_df: pd.DataFrame) -> list[dict]:
    """
    Check data against schema constraints and return list of violations.
    
    Args:
        samples_df: DataFrame with sample metadata
        results_df: DataFrame with lab results
        dec_df: DataFrame with decisions
    
    Returns:
        List of dicts with keys: table, sample_id, field, value, valid_values
    """
    violations = []
    
    # Check samples for invalid location_type
    if len(samples_df) > 0 and 'location_type' in samples_df.columns:
        for _, row in samples_df.iterrows():
            val = row.get('location_type')
            if val and val not in LOCATION_TYPES:
                violations.append({
                    'table': 'samples',
                    'sample_id': row.get('sample_id', '?'),
                    'field': 'location_type',
                    'value': val,
                    'valid_values': ', '.join(sorted(LOCATION_TYPES)),
                })
    
    # Check samples for invalid sample_type
    if len(samples_df) > 0 and 'sample_type' in samples_df.columns:
        for _, row in samples_df.iterrows():
            val = row.get('sample_type')
            if val and val not in SAMPLE_TYPES:
                violations.append({
                    'table': 'samples',
                    'sample_id': row.get('sample_id', '?'),
                    'field': 'sample_type',
                    'value': val,
                    'valid_values': ', '.join(sorted(SAMPLE_TYPES)),
                })
    
    # Check results for invalid analysis_type
    if len(results_df) > 0 and 'analysis_type' in results_df.columns:
        for _, row in results_df.iterrows():
            val = row.get('analysis_type')
            if val and val not in ANALYSIS_TYPES:
                violations.append({
                    'table': 'results',
                    'sample_id': row.get('sample_id', '?'),
                    'field': 'analysis_type',
                    'value': val,
                    'valid_values': ', '.join(sorted(ANALYSIS_TYPES)),
                })
    
    # Check results for invalid unit
    if len(results_df) > 0 and 'unit' in results_df.columns:
        # Only report unique violations per unit value
        seen_units = set()
        for _, row in results_df.iterrows():
            val = row.get('unit')
            if val and val not in UNIT_TYPES and val not in seen_units:
                seen_units.add(val)
                violations.append({
                    'table': 'results',
                    'sample_id': '(multiple)',
                    'field': 'unit',
                    'value': val,
                    'valid_values': ', '.join(sorted(UNIT_TYPES)),
                })
    
    # Check decisions for invalid decision
    if len(dec_df) > 0 and 'decision' in dec_df.columns:
        for _, row in dec_df.iterrows():
            val = row.get('decision')
            if val and val not in DECISION_TYPES:
                violations.append({
                    'table': 'decisions',
                    'sample_id': row.get('sample_id', '?'),
                    'field': 'decision',
                    'value': val,
                    'valid_values': ', '.join(sorted(DECISION_TYPES)),
                })
    
    return violations


def find_duplicate_results(results_df: pd.DataFrame) -> list[dict]:
    """
    Check for duplicate results (same sample_id + parameter + analysis_type).
    
    Args:
        results_df: DataFrame with lab results
    
    Returns:
        List of dicts with keys: sample_id, parameter, analysis_type, count, values
    """
    duplicates = []
    
    if len(results_df) == 0:
        return duplicates
    
    # Define key columns for uniqueness check
    key_cols = ['sample_id', 'parameter']
    if 'analysis_type' in results_df.columns:
        key_cols.append('analysis_type')
    
    # Find duplicates
    dup_mask = results_df.duplicated(subset=key_cols, keep=False)
    dup_df = results_df[dup_mask]
    
    if len(dup_df) == 0:
        return duplicates
    
    # Group duplicates and report
    for key, group in dup_df.groupby(key_cols):
        if len(group) > 1:
            if len(key_cols) == 3:
                sample_id, parameter, analysis_type = key
            else:
                sample_id, parameter = key
                analysis_type = ''
            
            # Get the values for each duplicate
            values = group['value'].tolist() if 'value' in group.columns else []
            values_str = ', '.join(str(v) for v in values)
            
            duplicates.append({
                'sample_id': sample_id,
                'parameter': parameter,
                'analysis_type': analysis_type,
                'count': len(group),
                'values': values_str,
            })
    
    return duplicates
