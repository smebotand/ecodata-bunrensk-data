"""
Standard schemas for bunnrensk data extraction.

All projects should use these schemas to ensure consistent output
that can be merged across projects.
"""

import pandas as pd
from typing import Optional

# ============================================================
# VALID VALUES (use these constants to ensure consistency)
# ============================================================

SAMPLE_TYPES = {
    'bunnrensk',       # Tunnel floor cleaning sediment
    'blandprøve',     # Mixed/composite sample
    'blandprøve bunnrensk',  # Composite bunnrensk sample (mixed from tunnel cross-section)
    'tunnelstein',     # Tunnel rock/excavation material, mixture of invert masses and blasted rock
    'sprengstein',     # Blasted rock
    'sediment',        # General sediment
    'vann',            # Water sample
}

LOCATION_TYPES = {
    'vegbane',         # Road surface
    'grøft',           # Ditch/drainage channel
    'pumpesump',       # Pump sump
    'ukjent',          # Unknown
}

ANALYSIS_TYPES = {
    'totalanalyse',    # Total content analysis
    'ristetest',       # Shaking/leaching test (L/S=10)
    'kolonnetest',     # Column leaching test (L/S=0.1)
    'kornfordeling',   # Grain size distribution
}

DECISION_TYPES = {
    'gjenbruk',        # Reuse
    'deponi',          # Disposal/landfill
    'supplerende prøvetaking',          # Additional sampling
    'ukjent'      # Unknown
}

UNIT_TYPES = {
    'mg/kg',      # Milligrams per kilogram (most common)
    'µg/kg',      # Micrograms per kilogram
    '%',          # Percent (for TS, grain size)
    'mg/l',       # Milligrams per liter (leaching tests)
    'mS/m',      # Millisiemens per meter (electrical conductivity)
}

LAB_TYPES = {
    'ALS Laboratory Group Norway AS',
    'Eurofins Environment Testing Norway AS',
}

# ============================================================
# SAMPLES SCHEMA
# ============================================================

SAMPLES_COLUMNS = [
    'sample_id',           # Unique identifier: p{nn}-{tunnel_code}-{profile}
    'project_code',        # Project code: e.g., '01_e18-e102', '18_hestnestunnelen'
    'sample_date',         # Date of sampling (YYYY-MM-DD)
    'location_type',       # Location in tunnel: vegbane, grøft, pumpesump, ukjent
    'profile_start',       # Start chainage (pel)
    'profile_end',         # End chainage (pel)
    'tunnel_name',         # Full tunnel name
    'sample_type',         # Type: bunnrensk, blandeprøve, etc.
    'lab_reference',       # Lab order/reference number
    'sampler',             # Who took the sample
    'remark',              # Additional remarks/notes
]

# ============================================================
# RESULTS SCHEMA
# ============================================================

RESULTS_COLUMNS = [
    'sample_id',           # Foreign key to samples
    'parameter',           # Standardized parameter code: As, Pb, PAH16, etc.
    'parameter_raw',       # Original parameter name from lab report
    'value',               # Numeric value
    'unit',                # Unit: mg/kg, µg/kg, %, etc.
    'uncertainty',         # Measurement uncertainty (if available)
    'below_limit',         # Boolean: True if value is below detection/quantification limit
    'loq',                 # Limit of quantification (if below_limit)
    'analysis_type',       # Type: totalanalyse, ristetest, kolonnetest
]

# ============================================================
# CLASSIFICATIONS SCHEMA
# ============================================================

CLASSIFICATIONS_COLUMNS = [
    'sample_id',           # Foreign key to samples
    'tilstandsklasse',     # Classification 1-5 (highest among all parameters)
    'limiting_parameters', # Parameters that determined the classification
    'classification_basis', # Basis for classification: e.g., 'TA-2553/2009'
]

# ============================================================
# DECISIONS SCHEMA
# ============================================================

DECISIONS_COLUMNS = [
    'sample_id',           # Foreign key to samples
    'decision',            # Decision: gjenbruk, deponi
    'decision_remarks',    # Additional remarks on decision
    'destination',         # Destination location (if known)
    'notes',               # Additional notes
]


def validate_value(value, valid_set: set, field_name: str, allow_empty: bool = True) -> str:
    """
    Validate that a value is in the allowed set.
    
    Args:
        value: The value to validate
        valid_set: Set of allowed values
        field_name: Name of the field (for error messages)
        allow_empty: If True, empty/None values are allowed
        
    Returns:
        The validated value
        
    Raises:
        ValueError: If value is not in valid_set
    """
    if allow_empty and (value is None or value == ''):
        return value
    if value not in valid_set:
        raise ValueError(
            f"Invalid {field_name}: '{value}'. "
            f"Must be one of: {sorted(valid_set)}"
        )
    return value


def validate_samples(data: list[dict]) -> list[dict]:
    """
    Validate samples data against schema constraints.
    
    Raises ValueError if any values are invalid.
    """
    for i, row in enumerate(data):
        sample_id = row.get('sample_id', f'row {i}')
        
        if 'location_type' in row and row['location_type']:
            validate_value(row['location_type'], LOCATION_TYPES, 
                          f"location_type in {sample_id}")
        
        if 'sample_type' in row and row['sample_type']:
            validate_value(row['sample_type'], SAMPLE_TYPES,
                          f"sample_type in {sample_id}")
    
    return data


def validate_results(data: list[dict]) -> list[dict]:
    """
    Validate results data against schema constraints.
    
    Raises ValueError if any values are invalid.
    """
    for i, row in enumerate(data):
        sample_id = row.get('sample_id', f'row {i}')
        param = row.get('parameter', '?')
        
        if 'analysis_type' in row and row['analysis_type']:
            validate_value(row['analysis_type'], ANALYSIS_TYPES,
                          f"analysis_type in {sample_id}/{param}")
        
        if 'unit' in row and row['unit']:
            validate_value(row['unit'], UNIT_TYPES,
                          f"unit in {sample_id}/{param}")
    
    return data


def validate_decisions(data: list[dict]) -> list[dict]:
    """
    Validate decisions data against schema constraints.
    
    Raises ValueError if any values are invalid.
    """
    for i, row in enumerate(data):
        sample_id = row.get('sample_id', f'row {i}')
        
        if 'decision' in row and row['decision']:
            validate_value(row['decision'], DECISION_TYPES,
                          f"decision in {sample_id}")
    
    return data


def create_samples_df(data: list[dict], validate: bool = True) -> pd.DataFrame:
    """Create a samples DataFrame with standard schema."""
    if validate:
        validate_samples(data)
    df = pd.DataFrame(data)
    # Ensure all columns exist
    for col in SAMPLES_COLUMNS:
        if col not in df.columns:
            df[col] = None
    # Reorder and select only standard columns
    return df[SAMPLES_COLUMNS]


def create_results_df(data: list[dict], validate: bool = True) -> pd.DataFrame:
    """Create a results DataFrame with standard schema."""
    if validate:
        validate_results(data)
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
            'destination': '',
            'notes': '',
        })
    
    return pd.DataFrame(decisions)[DECISIONS_COLUMNS]


def create_wide_table(results_df: pd.DataFrame, project_prefix: str) -> pd.DataFrame:
    """
    Create a wide format table with samples as columns.
    Includes analysis_type to distinguish totalanalyse, ristetest, kolonnetest.
    
    Args:
        results_df: Results dataframe with standard schema
        project_prefix: Prefix for sample columns (e.g., 'p01-', 'p18-')
    """
    if len(results_df) == 0:
        return pd.DataFrame()
    
    # Create formatted value strings
    def format_value(row):
        if pd.isna(row['value']):
            return ''
        val = row['value']
        if val == int(val):
            val_str = str(int(val))
        elif val < 0.01:
            val_str = f"{val:.4f}"
        elif val < 1:
            val_str = f"{val:.3f}"
        else:
            val_str = f"{val:.2f}".rstrip('0').rstrip('.')
        
        if row['below_limit']:
            val_str = f"<{val_str}"
        return val_str
    
    results_df = results_df.copy()
    results_df['formatted'] = results_df.apply(format_value, axis=1)
    
    # Get unit per sample
    sample_units = results_df.groupby('sample_id')['unit'].agg(
        lambda x: x.value_counts().index[0]
    ).to_dict()
    
    # Pivot with analysis_type in the index
    wide_df = results_df.pivot_table(
        index=['analysis_type', 'parameter', 'parameter_raw'],
        columns='sample_id',
        values='formatted',
        aggfunc='first'
    ).reset_index()
    
    wide_df.columns.name = None
    wide_df = wide_df.rename(columns={'parameter': 'param_code', 'parameter_raw': 'parameter'})
    
    # Add units to column headers
    sample_cols = [c for c in wide_df.columns if c.startswith(project_prefix)]
    sample_cols.sort()
    col_rename = {col: f"{col} ({sample_units.get(col, 'mg/kg')})" for col in sample_cols}
    wide_df = wide_df.rename(columns=col_rename)
    
    # Reorder columns
    meta_cols = ['analysis_type', 'param_code', 'parameter']
    renamed_cols = [f"{col} ({sample_units.get(col, 'mg/kg')})" for col in sample_cols]
    wide_df = wide_df[meta_cols + renamed_cols]
    
    return wide_df
