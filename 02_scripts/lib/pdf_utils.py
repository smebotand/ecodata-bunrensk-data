"""
PDF utilities for extracting text and tables from PDF documents.
"""
from __future__ import annotations

from pathlib import Path
from typing import List, Dict, Any, Optional

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False
    print("Warning: pdfplumber not installed. PDF functions limited.")

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


def extract_text(filepath: str | Path) -> str:
    """
    Extract all text from a PDF file.
    
    Args:
        filepath: Path to PDF file
        
    Returns:
        Concatenated text from all pages
    """
    if not HAS_PDFPLUMBER:
        raise ImportError("pdfplumber required: pip install pdfplumber")
    
    filepath = Path(filepath)
    text_parts = []
    
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_parts.append(text)
    
    return '\n\n'.join(text_parts)


def extract_text_by_page(filepath: str | Path) -> List[str]:
    """
    Extract text from each page separately.
    
    Returns:
        List of text strings, one per page
    """
    if not HAS_PDFPLUMBER:
        raise ImportError("pdfplumber required: pip install pdfplumber")
    
    filepath = Path(filepath)
    pages = []
    
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            pages.append(text)
    
    return pages


def extract_pages(filepath: str | Path, page_numbers: List[int]) -> str:
    """
    Extract text from specific pages (1-indexed).
    
    Args:
        filepath: Path to PDF file
        page_numbers: List of page numbers to extract (1-indexed)
        
    Returns:
        Concatenated text from specified pages
    """
    if not HAS_PDFPLUMBER:
        raise ImportError("pdfplumber required: pip install pdfplumber")
    
    filepath = Path(filepath)
    text_parts = []
    
    with pdfplumber.open(filepath) as pdf:
        for page_num in page_numbers:
            if 1 <= page_num <= len(pdf.pages):
                page = pdf.pages[page_num - 1]  # Convert to 0-indexed
                text = page.extract_text()
                if text:
                    text_parts.append(f"--- Page {page_num} ---\n{text}")
    
    return '\n\n'.join(text_parts)


def extract_tables(filepath: str | Path) -> List[Dict[str, Any]]:
    """
    Extract all tables from a PDF file.
    
    Returns:
        List of dicts with 'page', 'table_num', and 'data' keys
    """
    if not HAS_PDFPLUMBER:
        raise ImportError("pdfplumber required: pip install pdfplumber")
    
    filepath = Path(filepath)
    all_tables = []
    
    with pdfplumber.open(filepath) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table_num, table in enumerate(tables, 1):
                all_tables.append({
                    'page': page_num,
                    'table_num': table_num,
                    'data': table
                })
    
    return all_tables


def tables_to_dataframes(filepath: str | Path) -> List['pd.DataFrame']:
    """
    Extract all tables from PDF as pandas DataFrames.
    
    Returns:
        List of DataFrames
    """
    if not HAS_PANDAS:
        raise ImportError("pandas required: pip install pandas")
    
    import pandas as pd
    
    tables = extract_tables(filepath)
    dataframes = []
    
    for table_info in tables:
        data = table_info['data']
        if data and len(data) > 1:
            # Use first row as header
            df = pd.DataFrame(data[1:], columns=data[0])
            df.attrs['source_page'] = table_info['page']
            df.attrs['table_num'] = table_info['table_num']
            dataframes.append(df)
    
    return dataframes


def search_text(filepath: str | Path, 
                pattern: str,
                case_sensitive: bool = False) -> List[Dict[str, Any]]:
    """
    Search for text pattern in PDF.
    
    Args:
        filepath: Path to PDF
        pattern: Text to search for
        case_sensitive: Whether search is case-sensitive
        
    Returns:
        List of matches with page number and context
    """
    import re
    
    pages = extract_text_by_page(filepath)
    matches = []
    
    flags = 0 if case_sensitive else re.IGNORECASE
    
    for page_num, text in enumerate(pages, 1):
        for match in re.finditer(pattern, text, flags):
            start = max(0, match.start() - 50)
            end = min(len(text), match.end() + 50)
            context = text[start:end]
            
            matches.append({
                'page': page_num,
                'match': match.group(),
                'context': context.replace('\n', ' ')
            })
    
    return matches


def get_pdf_info(filepath: str | Path) -> Dict[str, Any]:
    """
    Get basic info about a PDF file.
    
    Returns:
        Dict with page count, metadata, etc.
    """
    if not HAS_PDFPLUMBER:
        raise ImportError("pdfplumber required: pip install pdfplumber")
    
    filepath = Path(filepath)
    
    with pdfplumber.open(filepath) as pdf:
        return {
            'path': str(filepath),
            'filename': filepath.name,
            'pages': len(pdf.pages),
            'metadata': pdf.metadata
        }
