"""
Tilstandsklasser (environmental quality classes) thresholds for Norwegian soil/sediment.

Based on TA-2553/2009 - Helsebaserte tilstandsklasser for forurenset grunn.

Tilstandsklasse definitions:
- TK1: Meget god tilstand (Very good) - below normverdier
- TK2: God tilstand (Good)
- TK3: Moderat tilstand (Moderate)
- TK4: Dårlig tilstand (Poor)
- TK5: Meget dårlig tilstand (Very poor)

All values in mg/kg (dry weight) unless otherwise noted.
"""
from typing import Optional, Dict, List, Tuple

# ============================================================
# NORMVERDIER (Class 1 upper limits) - TA-2553/2009
# ============================================================
# These are the "normverdier" - values below these are considered
# background/natural levels (Tilstandsklasse 1)

NORMVERDIER = {
    # Heavy metals
    'As': 8,
    'Pb': 60,
    'Cd': 1.5,
    'Hg': 1,
    'Cu': 100,
    'Zn': 200,
    'Cr_total': 50,
    'Cr_VI': 2,
    'Ni': 60,
    'Cyanid_fri': 1,
    
    # PCB
    'PCB7': 0.01,
    
    # Chlorinated pesticides
    'Lindan': 0.001,
    'DDT': 0.04,
    
    # Chlorinated benzenes
    'Monoklorbenzen': 0.03,
    'Diklorbenzen_12': 0.1,
    'Diklorbenzen_14': 0.07,
    'Triklorbenzen_124': 0.05,
    'Triklorbenzen_123': 0.01,
    'Triklorbenzen_135': 0.01,
    'Tetraklorbenzen_1245': 0.05,
    'Pentaklorbenzen': 0.1,
    'Heksaklorbenzen': 0.01,
    
    # Volatile halogenated hydrocarbons
    'Diklormetan': 0.06,
    'Triklormetan': 0.02,
    'Trikloreten': 0.1,
    'Tetraklormetan': 0.02,
    'Tetrakloreten': 0.01,
    'Dikloretan_12': 0.01,
    'Dibrometan_12': 0.004,
    'Trikloretan_111': 0.1,
    'Trikloretan_112': 0.01,
    
    # Phenols and chlorophenols
    'Fenol': 0.1,
    'Klorfenol_sum': 0.06,
    'Pentaklorfenol': 0.006,
    
    # PAH compounds
    'PAH16': 2,
    'Naftalen': 0.8,
    'Fluoren': 0.8,
    'Fluoranten': 1,
    'Pyren': 1,
    'BaP': 0.1,
    
    # BTEX
    'Benzen': 0.01,
    'Toluen': 0.3,
    'Etylbenzen': 0.2,
    'Xylen': 0.2,
    
    # Aliphatic hydrocarbons
    'C5_C6': 7,
    'C6_C8': 7,
    'C8_C10': 10,
    'C10_C12': 50,
    'C12_C35': 100,
    
    # Gasoline/oil additives
    'MTBE': 0.2,
    'Tetraetylbly': 0.001,
    
    # Brominated flame retardants
    'PBDE_99': 0.08,
    'PBDE_209': 0.002,
    
    # PFAS
    'PFOS': 0.1,
    
    # Phthalates
    'DEHP': 2.8,
    
    # Dioxins/furans
    'Dioksiner': 0.00001,
    
    # Organotin compounds
    'TBT': 0.015,
    'TPHT': 0.015,
}

# ============================================================
# TILSTANDSKLASSE THRESHOLDS (TA-2553/2009)
# ============================================================

# Format: parameter_code -> {name, TK1, TK2, TK3, TK4, TK5}
# TK1 = upper limit for class 1 (normverdi), TK5 = upper limit for class 5
# For parameters with only normverdi defined, TK2-TK5 are set to None

TILSTANDSKLASSER = {
    # Heavy metals (full thresholds)
    'As': {'name': 'Arsen', 'TK1': 8, 'TK2': 20, 'TK3': 50, 'TK4': 600, 'TK5': 1000},
    'Pb': {'name': 'Bly', 'TK1': 60, 'TK2': 100, 'TK3': 300, 'TK4': 700, 'TK5': 2500},
    'Cd': {'name': 'Kadmium', 'TK1': 1.5, 'TK2': 10, 'TK3': 15, 'TK4': 30, 'TK5': 1000},
    'Cu': {'name': 'Kobber', 'TK1': 100, 'TK2': 200, 'TK3': 1000, 'TK4': 8500, 'TK5': 25000},
    'Cr_total': {'name': 'Krom (total)', 'TK1': 50, 'TK2': 200, 'TK3': 500, 'TK4': 2800, 'TK5': 25000},
    'Cr_VI': {'name': 'Krom VI', 'TK1': 2, 'TK2': 5, 'TK3': 20, 'TK4': 80, 'TK5': 1000},
    'Hg': {'name': 'Kvikksølv', 'TK1': 1, 'TK2': 2, 'TK3': 4, 'TK4': 10, 'TK5': 1000},
    'Ni': {'name': 'Nikkel', 'TK1': 60, 'TK2': 135, 'TK3': 200, 'TK4': 1200, 'TK5': 2500},
    'Zn': {'name': 'Sink', 'TK1': 200, 'TK2': 500, 'TK3': 1000, 'TK4': 5000, 'TK5': 25000},
    'Cyanid_fri': {'name': 'Cyanid fri', 'TK1': 1, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    
    # PCB (full thresholds)
    'PCB7': {'name': 'Sum PCB7', 'TK1': 0.01, 'TK2': 0.5, 'TK3': 1, 'TK4': 5, 'TK5': 50},
    
    # Chlorinated pesticides
    'Lindan': {'name': 'Lindan', 'TK1': 0.001, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'DDT': {'name': 'Sum DDT', 'TK1': 0.04, 'TK2': 4, 'TK3': 12, 'TK4': 30, 'TK5': 50},
    
    # Chlorinated benzenes (normverdi only)
    'Monoklorbenzen': {'name': 'Monoklorbenzen', 'TK1': 0.03, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Diklorbenzen_12': {'name': '1,2-diklorbenzen', 'TK1': 0.1, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Diklorbenzen_14': {'name': '1,4-diklorbenzen', 'TK1': 0.07, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Triklorbenzen_124': {'name': '1,2,4-triklorbenzen', 'TK1': 0.05, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Triklorbenzen_123': {'name': '1,2,3-triklorbenzen', 'TK1': 0.01, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Triklorbenzen_135': {'name': '1,3,5-triklorbenzen', 'TK1': 0.01, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Tetraklorbenzen_1245': {'name': '1,2,4,5-tetraklorbenzen', 'TK1': 0.05, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Pentaklorbenzen': {'name': 'Pentaklorbenzen', 'TK1': 0.1, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Heksaklorbenzen': {'name': 'Heksaklorbenzen', 'TK1': 0.01, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    
    # Volatile halogenated hydrocarbons
    'Diklormetan': {'name': 'Diklormetan', 'TK1': 0.06, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Triklormetan': {'name': 'Triklormetan', 'TK1': 0.02, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Trikloreten': {'name': 'Trikloreten', 'TK1': 0.1, 'TK2': 0.2, 'TK3': 0.6, 'TK4': 0.8, 'TK5': 1000},
    'Tetraklormetan': {'name': 'Tetraklormetan', 'TK1': 0.02, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Tetrakloreten': {'name': 'Tetrakloreten', 'TK1': 0.01, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Dikloretan_12': {'name': '1,2-dikloretan', 'TK1': 0.01, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Dibrometan_12': {'name': '1,2-dibrometan', 'TK1': 0.004, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Trikloretan_111': {'name': '1,1,1-trikloretan', 'TK1': 0.1, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Trikloretan_112': {'name': '1,1,2-trikloretan', 'TK1': 0.01, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    
    # Phenols and chlorophenols
    'Fenol': {'name': 'Fenol', 'TK1': 0.1, 'TK2': 4, 'TK3': 40, 'TK4': 400, 'TK5': 25000},
    'Klorfenol_sum': {'name': 'Sum klorfenol', 'TK1': 0.06, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Pentaklorfenol': {'name': 'Pentaklorfenol', 'TK1': 0.006, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    
    # PAH compounds
    'PAH16': {'name': 'Sum PAH16', 'TK1': 2, 'TK2': 8, 'TK3': 50, 'TK4': 150, 'TK5': 2500},
    'Naftalen': {'name': 'Naftalen', 'TK1': 0.8, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Fluoren': {'name': 'Fluoren', 'TK1': 0.8, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Fluoranten': {'name': 'Fluoranten', 'TK1': 1, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Pyren': {'name': 'Pyren', 'TK1': 1, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'BaP': {'name': 'Benzo(a)pyren', 'TK1': 0.1, 'TK2': 0.5, 'TK3': 5, 'TK4': 15, 'TK5': 50},
    
    # BTEX
    'Benzen': {'name': 'Benzen', 'TK1': 0.01, 'TK2': 0.015, 'TK3': 0.04, 'TK4': 0.05, 'TK5': 1000},
    'Toluen': {'name': 'Toluen', 'TK1': 0.3, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Etylbenzen': {'name': 'Etylbenzen', 'TK1': 0.2, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Xylen': {'name': 'Xylen', 'TK1': 0.2, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    
    # Aliphatic hydrocarbons
    'C5_C6': {'name': 'Alifater C5-C6', 'TK1': 7, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'C6_C8': {'name': 'Alifater >C6-C8', 'TK1': 7, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'C8_C10': {'name': 'Alifater >C8-C10', 'TK1': 10, 'TK2': 10, 'TK3': 40, 'TK4': 50, 'TK5': 2000},
    'C10_C12': {'name': 'Alifater >C10-C12', 'TK1': 50, 'TK2': 60, 'TK3': 130, 'TK4': 300, 'TK5': 2000},
    'C12_C35': {'name': 'Alifater >C12-C35', 'TK1': 100, 'TK2': 300, 'TK3': 600, 'TK4': 2000, 'TK5': 20000},
    
    # Gasoline/oil additives
    'MTBE': {'name': 'MTBE', 'TK1': 0.2, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'Tetraetylbly': {'name': 'Tetraetylbly', 'TK1': 0.001, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    
    # Brominated flame retardants
    'PBDE_99': {'name': 'PBDE-99', 'TK1': 0.08, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'PBDE_209': {'name': 'PBDE-209', 'TK1': 0.002, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    
    # PFAS
    'PFOS': {'name': 'PFOS', 'TK1': 0.1, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    
    # Phthalates
    'DEHP': {'name': 'DEHP', 'TK1': 2.8, 'TK2': 25, 'TK3': 40, 'TK4': 60, 'TK5': 5000},
    
    # Dioxins/furans
    'Dioksiner': {'name': 'Dioksiner (TEQ)', 'TK1': 0.00001, 'TK2': 0.00002, 'TK3': 0.0001, 'TK4': 0.00036, 'TK5': 0.015},
    
    # Organotin compounds
    'TBT': {'name': 'TBT', 'TK1': 0.015, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
    'TPHT': {'name': 'TPHT', 'TK1': 0.015, 'TK2': None, 'TK3': None, 'TK4': None, 'TK5': None},
}


def exceeds_normverdi(parameter: str, value: float) -> Optional[bool]:
    """
    Check if a value exceeds the normverdi (class 1 threshold).
    
    Args:
        parameter: Parameter code (e.g., 'As', 'Pb', 'PAH16')
        value: Measured value in mg/kg
        
    Returns:
        True if exceeds, False if within normverdi, None if parameter not found
    """
    if parameter not in NORMVERDIER:
        return None
    return value > NORMVERDIER[parameter]


def get_tilstandsklasse(parameter: str, value: float) -> Optional[int]:
    """
    Determine tilstandsklasse for a given parameter value.
    
    Args:
        parameter: Parameter code (e.g., 'As', 'Pb', 'PAH16')
        value: Measured value in mg/kg
        
    Returns:
        Tilstandsklasse (1-5) or None if parameter not found.
        For parameters with only normverdi (TK1), returns 1 if below threshold,
        or None if above (since TK2-TK5 are not defined).
    """
    if parameter not in TILSTANDSKLASSER:
        return None
    
    thresholds = TILSTANDSKLASSER[parameter]
    
    if value <= thresholds['TK1']:
        return 1
    elif thresholds['TK2'] is None:
        # Only normverdi defined - can't classify beyond TK1
        return None
    elif value <= thresholds['TK2']:
        return 2
    elif value <= thresholds['TK3']:
        return 3
    elif value <= thresholds['TK4']:
        return 4
    else:
        return 5


def get_sample_tilstandsklasse(results: Dict[str, float]) -> Tuple[int, List[str]]:
    """
    Determine overall tilstandsklasse for a sample based on all parameters.
    
    The overall class is the worst (highest) class among all parameters.
    
    Args:
        results: Dictionary of parameter -> value
        
    Returns:
        Tuple of (tilstandsklasse, list of limiting parameters)
    """
    worst_class = 1
    limiting_params = []
    
    for param, value in results.items():
        tk = get_tilstandsklasse(param, value)
        if tk is not None:
            if tk > worst_class:
                worst_class = tk
                limiting_params = [param]
            elif tk == worst_class and tk > 1:
                limiting_params.append(param)
    
    return worst_class, limiting_params
