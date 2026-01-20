"""
Export utilities for saving processed data.
"""
from __future__ import annotations

import pandas as pd
from pathlib import Path


def save_to_csv(df: pd.DataFrame, 
                filepath: str | Path,
                encoding: str = 'utf-8-sig') -> Path:
    """
    Save DataFrame to CSV with UTF-8 BOM for Excel compatibility.
    
    Args:
        df: DataFrame to save
        filepath: Output path
        encoding: File encoding (default utf-8-sig for Excel)
        
    Returns:
        Path to saved file
    """
    filepath = Path(filepath)
    filepath.parent.mkdir(parents=True, exist_ok=True)
    
    df.to_csv(filepath, index=False, encoding=encoding)
    print(f"Saved: {filepath} ({len(df)} rows)")
    
    return filepath
