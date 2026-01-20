"""
Parsers for various lab report formats.

This module provides parsers and utilities for extracting structured data from 
different lab report formats (PDF, Excel) from various laboratories (ALS, Eurofins, etc.).
"""

from lib.parser.pdf_utils import extract_text, extract_text_by_page, extract_pages, extract_tables
from lib.parser.als_pdf import lab_results_from_als_pdf_with_THC

__all__ = [
    # PDF utilities
    'extract_text',
    'extract_text_by_page', 
    'extract_pages',
    'extract_tables',
    # ALS parser
    'lab_results_from_als_pdf_with_THC',
]
