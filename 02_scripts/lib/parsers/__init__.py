"""
Parsers for various lab report formats.

This module provides parsers for extracting structured data from different
lab report formats (PDF, Excel) from various laboratories (ALS, Eurofins, etc.).
"""

from .als_pdf import lab_results_from_als_pdf_with_THC

__all__ = [
    'lab_results_from_als_pdf_with_THC',
]
