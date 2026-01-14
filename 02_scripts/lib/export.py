"""
Export utilities for saving processed data.
"""
from __future__ import annotations

import pandas as pd
from pathlib import Path
from typing import List, Dict, Any, Optional
from datetime import datetime


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


def save_to_excel(dataframes: Dict[str, pd.DataFrame],
                  filepath: str | Path) -> Path:
    """
    Save multiple DataFrames to Excel with one sheet per DataFrame.
    
    Args:
        dataframes: Dict of sheet_name -> DataFrame
        filepath: Output path
        
    Returns:
        Path to saved file
    """
    filepath = Path(filepath)
    filepath.parent.mkdir(parents=True, exist_ok=True)
    
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        for sheet_name, df in dataframes.items():
            # Excel sheet names max 31 chars
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
    
    print(f"Saved: {filepath} ({len(dataframes)} sheets)")
    
    return filepath


def generate_extraction_report(project_code: str,
                               extracted_files: List[Dict[str, Any]],
                               output_dir: str | Path) -> Path:
    """
    Generate a summary report of extracted data.
    
    Args:
        project_code: Project identifier
        extracted_files: List of dicts with extraction info
        output_dir: Directory for output
        
    Returns:
        Path to report file
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    report_lines = [
        f"# Extraction Report: {project_code}",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "## Files Processed",
        ""
    ]
    
    for f in extracted_files:
        status = "✓" if f.get('success', True) else "✗"
        report_lines.append(f"- {status} {f.get('filename', 'unknown')}")
        if f.get('rows'):
            report_lines.append(f"  - Rows: {f['rows']}")
        if f.get('output'):
            report_lines.append(f"  - Output: {f['output']}")
        if f.get('error'):
            report_lines.append(f"  - Error: {f['error']}")
    
    report_lines.extend([
        "",
        "## Summary",
        f"- Total files: {len(extracted_files)}",
        f"- Successful: {sum(1 for f in extracted_files if f.get('success', True))}",
        f"- Failed: {sum(1 for f in extracted_files if not f.get('success', True))}",
    ])
    
    report_path = output_dir / f"{project_code}_extraction_report.md"
    report_path.write_text('\n'.join(report_lines), encoding='utf-8')
    
    print(f"Report saved: {report_path}")
    return report_path


def combine_project_data(csv_files: List[str | Path],
                         output_path: str | Path,
                         add_source_column: bool = True) -> pd.DataFrame:
    """
    Combine multiple CSV files into one master file.
    
    Args:
        csv_files: List of CSV file paths
        output_path: Path for combined output
        add_source_column: Whether to add a column with source filename
        
    Returns:
        Combined DataFrame
    """
    dfs = []
    
    for csv_file in csv_files:
        csv_file = Path(csv_file)
        if csv_file.exists():
            df = pd.read_csv(csv_file)
            if add_source_column:
                df['_source_file'] = csv_file.name
            dfs.append(df)
    
    if not dfs:
        print("No files to combine")
        return pd.DataFrame()
    
    combined = pd.concat(dfs, ignore_index=True)
    save_to_csv(combined, output_path)
    
    return combined
