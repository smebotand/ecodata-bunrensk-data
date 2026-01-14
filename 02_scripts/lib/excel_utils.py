"""
Excel utilities for reading and parsing lab reports and data files.
"""
from __future__ import annotations

import pandas as pd
from pathlib import Path
from typing import Optional, List, Dict, Any, Union


def read_excel_file(filepath: str | Path, 
                    sheet_name: str | int = 0,
                    header_row: int = 0,
                    skip_rows: List[int] = None) -> pd.DataFrame:
    """
    Read an Excel file with common options.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Sheet name or index (0-based)
        header_row: Row number to use as column headers (0-based)
        skip_rows: List of row indices to skip
        
    Returns:
        DataFrame with the data
    """
    filepath = Path(filepath)
    
    return pd.read_excel(
        filepath,
        sheet_name=sheet_name,
        header=header_row,
        skiprows=skip_rows
    )


def list_sheets(filepath: str | Path) -> List[str]:
    """List all sheet names in an Excel file."""
    xl = pd.ExcelFile(filepath)
    return xl.sheet_names


def find_header_row(filepath: str | Path, 
                    sheet_name: str | int = 0,
                    keywords: List[str] = None,
                    max_rows: int = 20) -> int:
    """
    Auto-detect the header row by looking for common column keywords.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Sheet to search
        keywords: List of keywords to look for (default: common lab terms)
        max_rows: Maximum rows to search
        
    Returns:
        Row index (0-based) of the header row
    """
    if keywords is None:
        keywords = ['prøve', 'sample', 'parameter', 'enhet', 'unit', 
                    'metall', 'metal', 'resultat', 'result', 'mg/kg', 
                    'analysemetode', 'LOQ', 'LOD']
    
    keywords_lower = [k.lower() for k in keywords]
    
    # Read first N rows without header
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, nrows=max_rows)
    
    for idx, row in df.iterrows():
        row_text = ' '.join(str(v).lower() for v in row.values if pd.notna(v))
        matches = sum(1 for kw in keywords_lower if kw in row_text)
        if matches >= 2:  # At least 2 keyword matches
            return idx
    
    return 0  # Default to first row


def read_lab_report(filepath: str | Path,
                    sheet_name: str | int = 0,
                    header_row: int = None,
                    auto_detect_header: bool = True) -> pd.DataFrame:
    """
    Read a lab report Excel file with smart header detection.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Sheet name or index
        header_row: Explicit header row (overrides auto-detect)
        auto_detect_header: Whether to auto-detect header row
        
    Returns:
        DataFrame with the lab data
    """
    filepath = Path(filepath)
    
    if header_row is None and auto_detect_header:
        header_row = find_header_row(filepath, sheet_name)
    elif header_row is None:
        header_row = 0
    
    df = pd.read_excel(filepath, sheet_name=sheet_name, header=header_row)
    
    # Clean up column names
    df.columns = [str(c).strip() for c in df.columns]
    
    # Drop completely empty rows and columns
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')
    
    return df


def extract_all_sheets(filepath: str | Path) -> Dict[str, pd.DataFrame]:
    """
    Extract all sheets from an Excel file.
    
    Returns:
        Dictionary of sheet_name -> DataFrame
    """
    filepath = Path(filepath)
    sheets = list_sheets(filepath)
    
    result = {}
    for sheet in sheets:
        try:
            df = read_lab_report(filepath, sheet_name=sheet)
            if not df.empty:
                result[sheet] = df
        except Exception as e:
            print(f"Warning: Could not read sheet '{sheet}': {e}")
    
    return result


def merge_sample_data(dataframes: List[pd.DataFrame],
                      sample_col: str = 'Prøve',
                      how: str = 'outer') -> pd.DataFrame:
    """
    Merge multiple dataframes by sample ID.
    
    Args:
        dataframes: List of DataFrames to merge
        sample_col: Column name containing sample IDs
        how: Merge method ('outer', 'inner', 'left', 'right')
        
    Returns:
        Merged DataFrame
    """
    if not dataframes:
        return pd.DataFrame()
    
    result = dataframes[0]
    for df in dataframes[1:]:
        result = result.merge(df, on=sample_col, how=how, suffixes=('', '_dup'))
    
    # Remove duplicate columns
    result = result.loc[:, ~result.columns.str.endswith('_dup')]
    
    return result
