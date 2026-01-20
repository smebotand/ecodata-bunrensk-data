"""
Standard schemas for bunnrensk data extraction.

All projects should use these schemas to ensure consistent output
that can be merged across projects.

This module defines only constants and type definitions.
For DataFrame creation functions, see lib/dataframes.py
"""

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

UNIT_TYPES_SOLID = {
    'mg/kg',      # Milligrams per kilogram (most common)
    'µg/kg',      # Micrograms per kilogram
    'g/kg',       # Grams per kilogram
}

UNIT_TYPES_LIQUID = {
    'mg/l',      # Milligrams per liter (most common)
    'µg/l',      # Micrograms per liter 
    'ng/l',       # Nanograms per liter
}

UNIT_TYPES_OTHER = {
    '%',          # Percent (for TS, grain size)
    'mS/m',      # Millisiemens per meter (electrical conductivity)
    "-",
    "% tørrvekt",
    "°C",
    "g",
    "mL",
    "µS/cm",
    "--",
    "mL/h",
    "cm"
}


UNIT_TYPES = {
    'mg/kg',      # Milligrams per kilogram (most common)
    'µg/kg',      # Micrograms per kilogram
    '%',          # Percent (for TS, grain size)
    'mg/l',       # Milligrams per liter (leaching tests)
    'mS/m',      # Millisiemens per meter (electrical conductivity)
    "-",
    "% tørrvekt",
    "°C",
    "g",
    "mL",
    "mg/L",
    "µS/cm",
    "--",
    "mL/h",
    "cm"
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
