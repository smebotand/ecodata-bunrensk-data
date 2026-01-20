"""
DataFrame creation utilities for bunnrensk data.

Functions for creating standardized DataFrames from raw data,
applying schema columns, and creating derived tables.
"""

import pandas as pd

from lib.schema import (
    SAMPLES_COLUMNS,
    RESULTS_COLUMNS,
    CLASSIFICATIONS_COLUMNS,
    DECISIONS_COLUMNS,
)


# ============================================================
# DATAFRAME CREATION FUNCTIONS
# ============================================================

def create_samples_df(data: list[dict]) -> pd.DataFrame:
    """Create a samples DataFrame with standard schema."""
    df = pd.DataFrame(data)
    # Ensure all columns exist
    for col in SAMPLES_COLUMNS:
        if col not in df.columns:
            df[col] = None
    # Reorder and select only standard columns
    return df[SAMPLES_COLUMNS]


def create_results_df(data: list[dict]) -> pd.DataFrame:
    """Create a results DataFrame with standard schema."""
    df = pd.DataFrame(data)
    # Ensure all columns exist
    for col in RESULTS_COLUMNS:
        if col not in df.columns:
            df[col] = None
    # Reorder and select only standard columns
    return df[RESULTS_COLUMNS]


def create_classifications_df(samples_df: pd.DataFrame, results_df: pd.DataFrame,
                               classification_basis: str = 'TA-2553/2009') -> pd.DataFrame:
    """
    Create classifications DataFrame from samples and results.
    
    Classification is based on highest tilstandsklasse among all parameters
    for each sample. TK1 = gjenbruk, otherwise deponi.
    """
    if len(results_df) == 0 or 'tilstandsklasse' not in results_df.columns:
        # Return empty dataframe with correct schema
        return pd.DataFrame(columns=CLASSIFICATIONS_COLUMNS)
    
    classifications = []
    
    for sample_id in samples_df['sample_id'].unique():
        sample_results = results_df[results_df['sample_id'] == sample_id]
        
        # Get highest tilstandsklasse
        max_tk = sample_results['tilstandsklasse'].max()
        
        # Find limiting parameters (those with max tilstandsklasse)
        if pd.notna(max_tk):
            limiting = sample_results[
                sample_results['tilstandsklasse'] == max_tk
            ]['parameter'].unique()
            limiting_str = ', '.join(sorted(limiting))
        else:
            limiting_str = ''
        
        classifications.append({
            'sample_id': sample_id,
            'tilstandsklasse': max_tk if pd.notna(max_tk) else None,
            'limiting_parameters': limiting_str,
            'classification_basis': classification_basis,
        })
    
    return pd.DataFrame(classifications)[CLASSIFICATIONS_COLUMNS]


def create_decisions_df(samples_df: pd.DataFrame, results_df: pd.DataFrame) -> pd.DataFrame:
    """
    Create decisions DataFrame from samples and results.
    
    Decision logic: TK1 = gjenbruk, otherwise deponi.
    """
    if len(results_df) == 0 or 'tilstandsklasse' not in results_df.columns:
        # Return empty dataframe with correct schema
        return pd.DataFrame(columns=DECISIONS_COLUMNS)
    
    decisions = []
    
    for sample_id in samples_df['sample_id'].unique():
        sample_results = results_df[results_df['sample_id'] == sample_id]
        max_tk = sample_results['tilstandsklasse'].max()
        
        if pd.notna(max_tk):
            decision = 'gjenbruk' if max_tk == 1 else 'deponi'
        else:
            decision = None
        
        decisions.append({
            'sample_id': sample_id,
            'decision': decision,
            'decision_remarks': '',
            'destination': '',
            'notes': '',
        })
    
    return pd.DataFrame(decisions)[DECISIONS_COLUMNS]
