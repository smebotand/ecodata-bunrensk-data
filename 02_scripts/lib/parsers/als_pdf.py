"""
ALS Laboratory PDF Parser

Parses chemical analysis results from ALS Laboratory Group Norway AS PDF reports.
These reports use the THC (Total Hydrocarbon Content) method for oil/hydrocarbon
fractions, which differs from Eurofins' Alifater method.

The PDF text extraction often strips spaces and adds footnote markers like "aulev",
so this parser handles those quirks.

Typical ALS PDF structure:
    Deresprøvenavn X
    Sediment
    Labnummer NXXXXXXXX
    Analyse Resultater Usikkerhet (±) Enhet Metode Utført Sign
    ParameterName 51.9 10.4 mg/kgTS ...

Usage:
    from lib.parsers import lab_results_from_als_pdf_with_THC
    
    results = lab_results_from_als_pdf_with_THC(
        text=pdf_text,
        sample_key_to_id={'1': 'p09-MOA-001', '2': 'p09-MOA-002'},
        project_code='09_moanetunnelen',
        default_sample_id_prefix='p09-MOA-'
    )
"""

import re
from typing import Dict, List, Optional, Callable

# Import the parameter mapping from chemistry module
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from chemistry import ALS_THC_PDF_MAP


def lab_results_from_als_pdf_with_THC(
    text: str,
    sample_key_to_id: Dict[str, str],
    project_code: str,
    default_sample_id_prefix: Optional[str] = None,
    sample_pattern: str = r'Deresprøvenavn\s+(\w+)',
    analysis_type: str = 'totalanalyse',
) -> List[dict]:
    """
    Parse ALS lab report PDF text to extract chemical analysis results.
    
    This parser handles ALS Laboratory Group Norway AS reports that use the THC
    (Total Hydrocarbon Content) method for oil fractions. It extracts results
    from PDF text that has been extracted using tools like pdfplumber or PyMuPDF.
    
    The parser:
    1. Splits the text into sample sections using the sample_pattern
    2. Cleans each section (removes footnote markers like "aulev")
    3. Matches parameters using regex patterns from ALS_THC_PDF_MAP
    4. Extracts values, uncertainties, and below-limit flags
    
    Args:
        text: Raw text extracted from the ALS PDF report.
        sample_key_to_id: Mapping from sample keys (as they appear in PDF, e.g., "1", "T1")
            to standardized sample IDs (e.g., "p09-MOA-001").
        project_code: Project code for generating fallback sample IDs.
        default_sample_id_prefix: Prefix for auto-generated sample IDs when a sample key
            is not found in sample_key_to_id. If None, uses f"{project_code.upper()}-".
        sample_pattern: Regex pattern to identify sample sections in the PDF.
            Must contain one capture group for the sample key.
            Default: r'Deresprøvenavn\\s+(\\w+)'
        analysis_type: Type of analysis (e.g., 'totalanalyse', 'ristetest', 'kolonnetest').
            This value is stored in each result record.
    
    Returns:
        List of result dictionaries, each containing:
            - sample_id: Standardized sample identifier
            - parameter: Standard parameter code (e.g., 'As', 'Pb', 'PAH16')
            - parameter_raw: Original parameter name as it appears in the report
            - value: Numeric value (float)
            - unit: Unit of measurement (e.g., 'mg/kg', '%')
            - uncertainty: Measurement uncertainty if available (float or None)
            - below_limit: True if value is below detection limit
            - loq: Limit of quantification if below_limit is True
            - analysis_type: Type of analysis performed
    
    Example:
        >>> text = extract_text(Path('als_report.pdf'))
        >>> sample_map = {'1': 'p09-MOA-001', '2': 'p09-MOA-002'}
        >>> results = lab_results_from_als_pdf_with_THC(
        ...     text=text,
        ...     sample_key_to_id=sample_map,
        ...     project_code='09_moanetunnelen'
        ... )
        >>> len(results)
        156
        >>> results[0]['parameter']
        'As'
    
    Note:
        This parser uses ALS_THC_PDF_MAP from lib/chemistry.py which defines
        regex patterns for matching parameter names in ALS reports. If you need
        to add new parameters, update that mapping.
    """
    if default_sample_id_prefix is None:
        # Extract project number from code like "09_moanetunnelen" -> "p09-MOA-"
        parts = project_code.split('_')
        if len(parts) >= 2:
            default_sample_id_prefix = f"p{parts[0]}-{parts[1][:3].upper()}-"
        else:
            default_sample_id_prefix = f"p{project_code}-"
    
    results = []
    
    # Find all sample sections
    sample_matches = list(re.finditer(sample_pattern, text))
    
    for i, match in enumerate(sample_matches):
        sample_key = match.group(1)
        start_pos = match.end()
        
        # End position is start of next sample or end of text
        if i + 1 < len(sample_matches):
            end_pos = sample_matches[i + 1].start()
        else:
            end_pos = len(text)
        
        section = text[start_pos:end_pos]
        
        # Get sample_id from mapping, or generate a default
        sample_id = sample_key_to_id.get(
            sample_key, 
            f'{default_sample_id_prefix}{sample_key}'
        )
        
        # Parse results from this section
        section_results = _parse_als_section_results(
            section=section,
            sample_id=sample_id,
            analysis_type=analysis_type
        )
        results.extend(section_results)
    
    return results


def _parse_als_section_results(
    section: str,
    sample_id: str,
    analysis_type: str = 'totalanalyse'
) -> List[dict]:
    """
    Parse all parameter results from a single sample section of an ALS PDF.
    
    This is an internal function that processes one sample's worth of data.
    It cleans the text, matches parameters using ALS_THC_PDF_MAP, and extracts
    values with proper handling of below-limit markers.
    
    Args:
        section: Text content for one sample section.
        sample_id: Standardized sample identifier.
        analysis_type: Type of analysis (stored in results).
    
    Returns:
        List of result dictionaries for this sample.
    """
    results = []
    found_params = set()  # Track what we've found to avoid duplicates
    
    # =========================================================================
    # CLEAN THE TEXT
    # Remove footnote references like "aulev", "^aulev", "a ulev"
    # These are common artifacts in ALS PDF extraction
    # =========================================================================
    clean_section = section
    clean_section = re.sub(r'\^?a\s*u\s*l\s*e\s*v', ' ', clean_section)  # Remove aulev variants
    clean_section = re.sub(r'\^', '', clean_section)  # Remove caret symbols
    clean_section = re.sub(r'\s+', ' ', clean_section)  # Normalize whitespace
    
    # =========================================================================
    # PARSE PARAMETERS using ALS_THC_PDF_MAP
    # Generic value pattern: captures value (with optional <) and optional uncertainty
    # =========================================================================
    value_pattern = r'\s+([<n][\d.,n\.d]+|[\d.,]+)\s+(?:(\d+[.,]?\d*)\s+)?(mg/kgTS|mg/kg\s*TS|%)'
    
    for param_regex, (param_code, param_raw, expected_unit) in ALS_THC_PDF_MAP.items():
        if param_code in found_params:
            continue
        
        # Build full pattern: param name + value + unit
        full_pattern = param_regex + value_pattern
        match = re.search(full_pattern, clean_section, re.IGNORECASE)
        
        if match:
            value_str = match.group(1).replace(',', '.').strip()
            uncertainty_str = match.group(2) if match.group(2) else None
            
            result = _build_als_result(
                sample_id=sample_id,
                param_code=param_code,
                param_raw=param_raw,
                value_str=value_str,
                unit=expected_unit,
                uncertainty_str=uncertainty_str,
                analysis_type=analysis_type
            )
            
            if result:
                results.append(result)
                found_params.add(param_code)
    
    return results


def _build_als_result(
    sample_id: str,
    param_code: str,
    param_raw: str,
    value_str: str,
    unit: str,
    uncertainty_str: Optional[str] = None,
    analysis_type: str = 'totalanalyse'
) -> Optional[dict]:
    """
    Build a result dictionary from parsed values.
    
    Handles various value formats:
    - Regular numbers: "51.9"
    - Below-limit values: "<0.5"
    - Not detected: "n.d.", "nd"
    
    Args:
        sample_id: Sample identifier.
        param_code: Standard parameter code.
        param_raw: Original parameter name for display.
        value_str: Raw value string from PDF.
        unit: Unit of measurement.
        uncertainty_str: Uncertainty value string if available.
        analysis_type: Type of analysis.
    
    Returns:
        Result dictionary, or None if value couldn't be parsed.
    """
    below_limit = False
    loq = None
    uncertainty = None
    value = None
    
    # Parse uncertainty
    if uncertainty_str:
        try:
            uncertainty = float(uncertainty_str.replace(',', '.'))
        except (ValueError, TypeError):
            pass
    
    # Parse value
    if value_str.startswith('<'):
        below_limit = True
        value_str = value_str[1:]
        try:
            loq = float(value_str)
            value = loq
        except (ValueError, TypeError):
            value = None
    elif value_str.lower() in ['n.d.', 'nd', 'n.d']:
        below_limit = True
        value = 0.0
    else:
        try:
            value = float(value_str)
        except (ValueError, TypeError):
            return None
    
    return {
        'sample_id': sample_id,
        'parameter': param_code,
        'parameter_raw': param_raw,
        'value': value,
        'unit': unit,
        'uncertainty': uncertainty,
        'below_limit': below_limit,
        'loq': loq,
        'analysis_type': analysis_type,
    }


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def get_lab_number_from_section(section: str) -> Optional[str]:
    """
    Extract the ALS lab number (e.g., "N00529080") from a sample section.
    
    Args:
        section: Text content of a sample section.
    
    Returns:
        Lab number string if found, None otherwise.
    """
    lab_match = re.search(r'Labnummer\s+(\w+)', section)
    return lab_match.group(1) if lab_match else None
