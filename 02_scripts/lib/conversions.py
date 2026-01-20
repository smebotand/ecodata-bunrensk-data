"""
Unit conversion utilities for chemical analysis data.
"""

from typing import Optional


# =============================================================================
# CONCENTRATION UNIT CONVERSIONS
# =============================================================================

# Conversion factors to normalize concentration units to mg/kg (dry weight)
# Usage: value_in_mgkg = value * CONCENTRATION_TO_MGKG[unit]
CONCENTRATION_TO_MGKG = {
    # mg/kg variants (factor = 1.0)
    'mg/kg': 1.0,
    'mg/kg ts': 1.0,
    'mg/kg dw': 1.0,
    'mg/kg tørrstoff': 1.0,
    
    # µg/kg variants (factor = 0.001)
    'µg/kg': 0.001,
    'ug/kg': 0.001,
    'μg/kg': 0.001,
    'µg/kg ts': 0.001,
    
    # ng/kg (factor = 0.000001)
    'ng/kg': 0.000001,
    
    # g/kg (factor = 1000)
    'g/kg': 1000.0,
    
    # ppm/ppb equivalents
    'ppm': 1.0,
    'ppb': 0.001,
    
    # Percentage (factor = 10000)
    '%': 10000.0,
}

# Conversion factors to normalize concentration units to mg/l (for water samples)
# Usage: value_in_mgl = value * CONCENTRATION_TO_MGL[unit]
CONCENTRATION_TO_MGL = {
    # mg/l variants (factor = 1.0)
    'mg/l': 1.0,
    'mg/l ts': 1.0,
    'mg/l dw': 1.0,
    'mg/l tørrstoff': 1.0,

    # µg/l variants (factor = 0.001)
    'µg/l': 0.001,
    'ug/l': 0.001,
    'μg/l': 0.001,
    'µg/l ts': 0.001,

    # ng/l (factor = 0.000001)
    'ng/l': 0.000001,

    # g/l (factor = 1000)
    'g/l': 1000.0,

    # ppm/ppb equivalents (for water, ppm ≈ mg/l, ppb ≈ µg/l)
    'ppm': 1.0,
    'ppb': 0.001,

    # Percentage (factor = 10000)
    '%': 10000.0,
}


def convert_to_mgkg(value: float, from_unit: str) -> Optional[float]:
    """
    Convert a concentration value to mg/kg (dry weight).
    
    Args:
        value: Numeric value in source unit
        from_unit: Source unit (e.g., 'µg/kg', 'ppm', '%')
        
    Returns:
        Value in mg/kg, or None if input is None
    """
    if value is None:
        return None
    
    from_unit_lower = from_unit.lower().strip()
    factor = CONCENTRATION_TO_MGKG.get(from_unit_lower, 1.0)
    
    return value * factor


def convert_units(value: float, from_unit: str, to_unit: str = 'mg/kg') -> Optional[float]:
    """
    Convert between concentration units.
    
    Args:
        value: Numeric value
        from_unit: Source unit
        to_unit: Target unit (default mg/kg)
        
    Returns:
        Converted value, or None if input is None
    """
    if value is None:
        return None
    
    from_unit_lower = from_unit.lower().strip()
    to_unit_lower = to_unit.lower().strip()
    
    # Get conversion factors (to mg/kg)
    from_factor = CONCENTRATION_TO_MGKG.get(from_unit_lower, 1.0)
    to_factor = CONCENTRATION_TO_MGKG.get(to_unit_lower, 1.0)
    
    # Convert: value * from_factor gives mg/kg, then divide by to_factor
    return value * from_factor / to_factor
