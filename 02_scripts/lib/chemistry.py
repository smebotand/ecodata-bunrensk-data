"""
Chemistry utilities for normalizing and processing chemical analysis data.
"""

import pandas as pd
import re
from typing import Dict, List, Optional, Tuple


# =============================================================================
# STANDARD PARAMETER DEFINITIONS
# =============================================================================

# Standard parameter name mappings (Norwegian -> English -> Standard code)
PARAMETER_ALIASES = {
    # Metals
    'arsen': 'As', 'arsenic': 'As', 'as': 'As',
    'bly': 'Pb', 'lead': 'Pb', 'pb': 'Pb',
    'kadmium': 'Cd', 'cadmium': 'Cd', 'cd': 'Cd',
    'kobber': 'Cu', 'copper': 'Cu', 'cu': 'Cu',
    'krom': 'Cr', 'chromium': 'Cr', 'cr': 'Cr', 'krom total': 'Cr',
    'krom (vi)': 'Cr_VI', 'krom vi': 'Cr_VI', 'cr(vi)': 'Cr_VI',
    'kvikksølv': 'Hg', 'mercury': 'Hg', 'hg': 'Hg',
    'nikkel': 'Ni', 'nickel': 'Ni', 'ni': 'Ni',
    'sink': 'Zn', 'zinc': 'Zn', 'zn': 'Zn',
    'jern': 'Fe', 'iron': 'Fe', 'fe': 'Fe',
    'mangan': 'Mn', 'manganese': 'Mn', 'mn': 'Mn',
    'barium': 'Ba', 'ba': 'Ba',
    'vanadium': 'V', 'v': 'V',
    'kobolt': 'Co', 'cobalt': 'Co', 'co': 'Co',
    'molybden': 'Mo', 'molybdenum': 'Mo', 'mo': 'Mo',
    'antimon': 'Sb', 'antimony': 'Sb', 'sb': 'Sb',
    'selen': 'Se', 'selenium': 'Se', 'se': 'Se',
    'tinn': 'Sn', 'tin': 'Sn', 'sn': 'Sn',
    'tallium': 'Tl', 'thallium': 'Tl', 'tl': 'Tl',
    'aluminium': 'Al', 'al': 'Al',
    'beryllium': 'Be', 'be': 'Be',
    'litium': 'Li', 'lithium': 'Li', 'li': 'Li',
    'sølv': 'Ag', 'silver': 'Ag', 'ag': 'Ag',
    'wolfram': 'W', 'tungsten': 'W', 'w': 'W',
    
    # Petroleum hydrocarbons (THC fractions)
    'thc': 'THC', 'total hydrocarbons': 'THC', 'totale hydrokarboner': 'THC',
    'tph': 'TPH', 'total petroleum hydrocarbons': 'TPH',
    'olje': 'THC', 'oil': 'THC', 'mineralsk olje': 'THC',
    'sum >c12-c35': 'THC_C12-C35', 'olje sum >c12-c35': 'THC_C12-C35',
    '>c10-c12': 'THC_C10-C12', 'fraksjon >c10-c12': 'THC_C10-C12',
    '>c12-c16': 'THC_C12-C16', 'fraksjon >c12-c16': 'THC_C12-C16',
    '>c16-c35': 'THC_C16-C35', 'fraksjon >c16-c35': 'THC_C16-C35',
    '>c12-c35': 'THC_C12-C35', 'fraksjon >c12-c35': 'THC_C12-C35',
    '>c35-c40': 'THC_C35-C40', 'fraksjon >c35-c40': 'THC_C35-C40',
    'c10-c40': 'THC_C10-C40', 'thc c10-c40': 'THC_C10-C40',
    
    # PAH (Polycyclic Aromatic Hydrocarbons)
    'pah': 'PAH', 'polycyclic aromatic hydrocarbons': 'PAH',
    'pah16': 'PAH16', 'pah 16': 'PAH16', 'sum pah16': 'PAH16', 'pah-16': 'PAH16',
    'pah 16 epa': 'PAH16', 'sum pah 16 epa': 'PAH16',
    'naftalen': 'Naphthalene', 'naphthalene': 'Naphthalene',
    'acenaftylen': 'Acenaphthylene', 'acenaphthylene': 'Acenaphthylene',
    'acenaften': 'Acenaphthene', 'acenaphthene': 'Acenaphthene',
    'fluoren': 'Fluorene', 'fluorene': 'Fluorene',
    'fenantren': 'Phenanthrene', 'phenanthrene': 'Phenanthrene',
    'antracen': 'Anthracene', 'anthracene': 'Anthracene',
    'fluoranten': 'Fluoranthene', 'fluoranthene': 'Fluoranthene',
    'pyren': 'Pyrene', 'pyrene': 'Pyrene',
    'benzo(a)antracen': 'BaA', 'benzo[a]anthracene': 'BaA',
    'krysen': 'Chrysene', 'chrysene': 'Chrysene',
    'benzo(b)fluoranten': 'BbF', 'benzo[b]fluoranthene': 'BbF',
    'benzo(k)fluoranten': 'BkF', 'benzo[k]fluoranthene': 'BkF',
    'benzo(a)pyren': 'BaP', 'benzo[a]pyrene': 'BaP',
    'indeno(1,2,3-cd)pyren': 'IcdP', 'indeno[1,2,3-cd]pyrene': 'IcdP',
    'dibenzo(a,h)antracen': 'DahA', 'dibenz[a,h]anthracene': 'DahA',
    'benzo(ghi)perylen': 'BghiP', 'benzo[ghi]perylene': 'BghiP',
    
    # PCB (Polychlorinated Biphenyls)
    'pcb': 'PCB', 'polychlorinated biphenyls': 'PCB',
    'pcb7': 'PCB7', 'pcb 7': 'PCB7', 'sum pcb7': 'PCB7', 'pcb-7': 'PCB7',
    'pcb 7 (seven dutch)': 'PCB7',
    'pcb 28': 'PCB28', 'pcb28': 'PCB28',
    'pcb 52': 'PCB52', 'pcb52': 'PCB52',
    'pcb 101': 'PCB101', 'pcb101': 'PCB101',
    'pcb 118': 'PCB118', 'pcb118': 'PCB118',
    'pcb 138': 'PCB138', 'pcb138': 'PCB138',
    'pcb 153': 'PCB153', 'pcb153': 'PCB153',
    'pcb 180': 'PCB180', 'pcb180': 'PCB180',
    
    # BTEX
    'btex': 'BTEX', 'sum btex': 'BTEX',
    'benzen': 'Benzene', 'benzene': 'Benzene',
    'toluen': 'Toluene', 'toluene': 'Toluene',
    'etylbenzen': 'Ethylbenzene', 'ethylbenzene': 'Ethylbenzene',
    'xylen': 'Xylene', 'xylene': 'Xylene', 'xylener': 'Xylene',
    'm,p-xylen': 'Xylene_mp', 'o-xylen': 'Xylene_o',
    
    # Other organics
    'pfas': 'PFAS', 'sum pfas': 'PFAS',
    'pfos': 'PFOS', 'pfoa': 'PFOA',
    
    # Physical parameters
    'tørrstoff': 'DryMatter', 'dry matter': 'DryMatter', 'ts': 'DryMatter',
    'tørrstoff (ts)': 'DryMatter', 'ts%': 'DryMatter',
    'ph': 'pH',
    'toc': 'TOC', 'total organic carbon': 'TOC', 'totalt organisk karbon': 'TOC',
    'glødetap': 'LOI', 'loss on ignition': 'LOI',
    'kornfordeling': 'GrainSize', 'siktekurve': 'GrainSize',
    'leire': 'Clay', 'silt': 'Silt', 'sand': 'Sand', 'grus': 'Gravel',
    
    # Nitrogen compounds
    'nitrogen': 'N', 'total nitrogen': 'N_total', 'total-n': 'N_total',
    'ammonium': 'NH4', 'nh4': 'NH4', 'ammonium-n': 'NH4',
    'nitrat': 'NO3', 'no3': 'NO3', 'nitrat-n': 'NO3',
    'nitritt': 'NO2', 'no2': 'NO2', 'nitritt-n': 'NO2',
    
    # Other
    'cyanid': 'CN', 'cyanide': 'CN', 'cn': 'CN',
    'sulfat': 'SO4', 'sulphate': 'SO4',
    'klorid': 'Cl', 'chloride': 'Cl',
    'fluorid': 'F', 'fluoride': 'F',
}


# Unit conversion factors to mg/kg
UNIT_CONVERSIONS = {
    'mg/kg': 1.0,
    'mg/kg ts': 1.0,
    'mg/kg dw': 1.0,
    'mg/kg tørrstoff': 1.0,
    'µg/kg': 0.001,
    'ug/kg': 0.001,
    'μg/kg': 0.001,
    'µg/kg ts': 0.001,
    'ng/kg': 0.000001,
    'g/kg': 1000.0,
    'ppm': 1.0,
    'ppb': 0.001,
    '%': 10000.0,
    'mg/l': 1.0,  # Approximation for water samples
    'µg/l': 0.001,
}


# Norwegian regulatory limits (Tilstandsklasser) in mg/kg
# Source: TA-2553/2009, Miljødirektoratet (updated with M-1884/2022 for some)
TILSTANDSKLASSER = {
    # Metals
    'As':  {'I': 8,    'II': 8,    'III': 20,   'IV': 50,    'V': 1000},
    'Pb':  {'I': 60,   'II': 60,   'III': 100,  'IV': 300,   'V': 1000},
    'Cd':  {'I': 1.5,  'II': 1.5,  'III': 3,    'IV': 15,    'V': 30},
    'Cu':  {'I': 100,  'II': 100,  'III': 200,  'IV': 1000,  'V': 1000},
    'Cr':  {'I': 50,   'II': 50,   'III': 100,  'IV': 500,   'V': 1000},
    'Cr_VI': {'I': 2,  'II': 2,    'III': 5,    'IV': 20,    'V': 1000},
    'Hg':  {'I': 1,    'II': 1,    'III': 2,    'IV': 10,    'V': 10},
    'Ni':  {'I': 60,   'II': 60,   'III': 100,  'IV': 500,   'V': 1000},
    'Zn':  {'I': 200,  'II': 200,  'III': 500,  'IV': 1000,  'V': 5000},
    'V':   {'I': 100,  'II': 100,  'III': 200,  'IV': 700,   'V': 700},
    'Co':  {'I': 20,   'II': 20,   'III': 50,   'IV': 300,   'V': 1000},
    
    # PAH
    'PAH16':     {'I': 2,    'II': 2,    'III': 8,    'IV': 50,   'V': 1000},
    'BaP':       {'I': 0.1,  'II': 0.1,  'III': 0.5,  'IV': 5,    'V': 100},
    'Naphthalene': {'I': 0.8, 'II': 0.8, 'III': 4,   'IV': 40,   'V': 1000},
    
    # PCB
    'PCB7':      {'I': 0.01, 'II': 0.01, 'III': 0.5,  'IV': 1,    'V': 5},
    
    # Petroleum hydrocarbons
    'THC':       {'I': 50,   'II': 50,   'III': 300,  'IV': 2000,  'V': 5000},
    'THC_C12-C35': {'I': 50, 'II': 50,   'III': 300,  'IV': 2000,  'V': 5000},
    'THC_C16-C35': {'I': 50, 'II': 50,   'III': 300,  'IV': 2000,  'V': 5000},
    
    # BTEX
    'Benzene':   {'I': 0.01, 'II': 0.01, 'III': 0.05, 'IV': 1,    'V': 5},
    
    # Other
    'CN':        {'I': 1,    'II': 1,    'III': 5,    'IV': 100,  'V': 500},
}


# =============================================================================
# LOCATION AND SAMPLE TYPE STANDARDIZATION
# =============================================================================

LOCATION_TYPES = {
    'vegbane': 'vegbane',
    'vegbanen': 'vegbane',
    'kjørebane': 'vegbane',
    'grøft': 'grøft',
    'grøfta': 'grøft',
    'sidegrøft': 'grøft',
    'deponi': 'deponi',
    'mellomlagring': 'deponi',
    'mellomlager': 'deponi',
    'tipp': 'deponi',
    'tunnel': 'tunnel',
    'stuff': 'stuff',
    'påhugg': 'påhugg',
    'portal': 'påhugg',
    'tverrslag': 'tverrslag',
    'sjakt': 'sjakt',
}

SAMPLE_TYPES = {
    'bunnrensk': 'bunnrensk',
    'bunnrenskmasser': 'bunnrensk',
    'sålerensk': 'sålerensk',
    'blandeprøve': 'blandeprøve',
    'blandprøve': 'blandeprøve',
    'stikkprøve': 'stikkprøve',
    'enkeltprøve': 'stikkprøve',
    'referanse': 'referanse',
    'rein stein': 'referanse',
    'ren stein': 'referanse',
}


# =============================================================================
# FUNCTIONS
# =============================================================================


def normalize_parameter_name(name: str) -> str:
    """
    Normalize a parameter name to standard form.
    
    Args:
        name: Raw parameter name from lab report
        
    Returns:
        Standardized parameter code
    """
    if pd.isna(name):
        return ''
    
    # Clean and lowercase
    clean = str(name).strip().lower()
    
    # Remove common prefixes/suffixes
    clean = re.sub(r'^(sum\s+|total\s+)', '', clean)
    clean = re.sub(r'\s*\(.*\)$', '', clean)  # Remove parenthetical notes
    
    # Look up in aliases
    if clean in PARAMETER_ALIASES:
        return PARAMETER_ALIASES[clean]
    
    # Return original with title case if not found
    return str(name).strip()


def normalize_location_type(location: str) -> str:
    """Normalize location type to standard form."""
    if pd.isna(location):
        return ''
    clean = str(location).strip().lower()
    return LOCATION_TYPES.get(clean, clean)


def normalize_sample_type(sample_type: str) -> str:
    """Normalize sample type to standard form."""
    if pd.isna(sample_type):
        return ''
    clean = str(sample_type).strip().lower()
    return SAMPLE_TYPES.get(clean, clean)


def parse_value(value: str) -> Tuple[float, bool]:
    """
    Parse a lab result value, handling '<' (below detection limit).
    
    Args:
        value: Value string from lab report
        
    Returns:
        Tuple of (numeric_value, is_below_detection_limit)
    """
    if pd.isna(value):
        return (None, False)
    
    value_str = str(value).strip()
    
    # Check for below detection limit
    below_limit = False
    if value_str.startswith('<') or value_str.startswith('&lt;'):
        below_limit = True
        value_str = re.sub(r'^[<&lt;]+\s*', '', value_str)
    
    # Also handle "n.d.", "nd", "ikke påvist" etc
    if value_str.lower() in ['n.d.', 'nd', 'ikke påvist', 'not detected', '-', 'i.p.']:
        return (0.0, True)
    
    # Extract numeric value
    try:
        # Handle Norwegian decimal comma
        value_str = value_str.replace(',', '.').replace(' ', '')
        # Remove any trailing text like "mg/kg"
        value_str = re.sub(r'[a-zA-Z/]+$', '', value_str)
        numeric = float(value_str)
        return (numeric, below_limit)
    except ValueError:
        return (None, False)


def parse_uncertainty(value: str) -> Optional[float]:
    """
    Parse uncertainty value from formats like "51.9 +/- 10.4" or "±10.4".
    
    Returns:
        Uncertainty value or None
    """
    if pd.isna(value):
        return None
    
    value_str = str(value)
    
    # Look for +/- or ± pattern
    match = re.search(r'[±+/-]+\s*(\d+[.,]?\d*)', value_str)
    if match:
        try:
            return float(match.group(1).replace(',', '.'))
        except ValueError:
            pass
    
    return None


def convert_units(value: float, from_unit: str, to_unit: str = 'mg/kg') -> float:
    """
    Convert between concentration units.
    
    Args:
        value: Numeric value
        from_unit: Source unit
        to_unit: Target unit (default mg/kg)
        
    Returns:
        Converted value
    """
    if value is None:
        return None
    
    from_unit_lower = from_unit.lower().strip()
    to_unit_lower = to_unit.lower().strip()
    
    # Get conversion factors
    from_factor = UNIT_CONVERSIONS.get(from_unit_lower, 1.0)
    to_factor = UNIT_CONVERSIONS.get(to_unit_lower, 1.0)
    
    # Convert: value * from_factor gives mg/kg, then divide by to_factor
    return value * from_factor / to_factor


def normalize_results_df(df: pd.DataFrame,
                         value_col: str = None,
                         parameter_col: str = None,
                         unit_col: str = None) -> pd.DataFrame:
    """
    Normalize a results dataframe with standard parameter names and units.
    
    Args:
        df: Input dataframe
        value_col: Column containing values (auto-detect if None)
        parameter_col: Column containing parameter names (auto-detect if None)
        unit_col: Column containing units (auto-detect if None)
        
    Returns:
        Normalized dataframe with added columns
    """
    df = df.copy()
    
    # Auto-detect columns if not specified
    cols_lower = {c.lower(): c for c in df.columns}
    
    if parameter_col is None:
        for key in ['parameter', 'analyse', 'komponent', 'stoff']:
            if key in cols_lower:
                parameter_col = cols_lower[key]
                break
    
    if value_col is None:
        for key in ['resultat', 'result', 'verdi', 'value', 'konsentrasjon']:
            if key in cols_lower:
                value_col = cols_lower[key]
                break
    
    if unit_col is None:
        for key in ['enhet', 'unit', 'måleenhet']:
            if key in cols_lower:
                unit_col = cols_lower[key]
                break
    
    # Normalize parameter names
    if parameter_col and parameter_col in df.columns:
        df['parameter_std'] = df[parameter_col].apply(normalize_parameter_name)
    
    # Parse values and detect below-limit
    if value_col and value_col in df.columns:
        parsed = df[value_col].apply(parse_value)
        df['value_numeric'] = parsed.apply(lambda x: x[0])
        df['below_limit'] = parsed.apply(lambda x: x[1])
        
        # Convert to mg/kg if unit column exists
        if unit_col and unit_col in df.columns:
            df['value_mg_kg'] = df.apply(
                lambda row: convert_units(row['value_numeric'], str(row[unit_col])) 
                if pd.notna(row['value_numeric']) else None,
                axis=1
            )
    
    return df


def classify_contamination(parameter: str, value: float) -> str:
    """
    Classify a contamination level according to Norwegian standards.
    
    Args:
        parameter: Standard parameter code (e.g., 'As', 'Pb')
        value: Concentration in mg/kg
        
    Returns:
        Tilstandsklasse ('I', 'II', 'III', 'IV', 'V') or 'Unknown'
    """
    if value is None or parameter not in TILSTANDSKLASSER:
        return 'Unknown'
    
    limits = TILSTANDSKLASSER[parameter]
    
    if value <= limits['I']:
        return 'I'
    elif value <= limits['II']:
        return 'II'
    elif value <= limits['III']:
        return 'III'
    elif value <= limits['IV']:
        return 'IV'
    else:
        return 'V'


def get_limiting_parameter(results: Dict[str, float]) -> Tuple[str, float, str]:
    """
    Find the parameter that determines the overall classification.
    
    Args:
        results: Dict of parameter -> value in mg/kg
        
    Returns:
        Tuple of (parameter, value, tilstandsklasse)
    """
    worst_class = 'I'
    worst_param = None
    worst_value = None
    
    class_order = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'Unknown': 0}
    
    for param, value in results.items():
        if value is None:
            continue
        
        classification = classify_contamination(param, value)
        if class_order.get(classification, 0) > class_order.get(worst_class, 0):
            worst_class = classification
            worst_param = param
            worst_value = value
    
    return (worst_param, worst_value, worst_class)


def classify_sample_results(sample_id: str, results: pd.DataFrame) -> Dict:
    """
    Classify a sample based on all its chemical results.
    
    Args:
        sample_id: Sample identifier
        results: DataFrame with columns 'parameter' and 'value' (in mg/kg)
        
    Returns:
        Dict with classification info matching classifications.csv schema
    """
    # Build dict of parameter -> value
    param_values = {}
    for _, row in results.iterrows():
        param = row.get('parameter', row.get('parameter_std', ''))
        value = row.get('value', row.get('value_numeric', row.get('value_mg_kg', None)))
        if param and value is not None:
            param_values[param] = value
    
    # Get limiting parameter
    limiting_param, limiting_value, tilstandsklasse = get_limiting_parameter(param_values)
    
    return {
        'sample_id': sample_id,
        'tilstandsklasse': tilstandsklasse,
        'limiting_parameter': limiting_param,
        'limiting_value': limiting_value,
        'classification_basis': 'TA-2553/2009'
    }


def format_result_for_csv(
    sample_id: str,
    parameter_raw: str,
    value_raw: str,
    unit: str = 'mg/kg',
    loq: float = None
) -> Dict:
    """
    Format a single result for output to results.csv.
    
    Args:
        sample_id: Sample identifier
        parameter_raw: Original parameter name from source
        value_raw: Original value string from source
        unit: Unit of measurement
        loq: Limit of quantification
        
    Returns:
        Dict matching results.csv schema
    """
    parameter_std = normalize_parameter_name(parameter_raw)
    value, below_limit = parse_value(value_raw)
    uncertainty = parse_uncertainty(str(value_raw))
    
    return {
        'sample_id': sample_id,
        'parameter': parameter_std,
        'parameter_raw': parameter_raw,
        'value': value,
        'unit': unit,
        'uncertainty': uncertainty,
        'below_limit': below_limit,
        'loq': loq
    }
