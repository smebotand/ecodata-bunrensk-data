# EcoData Bunrensk - Document Organization System

## Overview
This repository organizes environmental and chemical data from approximately 27 projects.
It handles PDFs, Excel sheets, and emails containing chemical data, sample descriptions, and metadata.

## Folder Structure

```
ecodata-bunrensk-data/
├── 00_inbox/                    # Drop raw files here for processing
├── 01_projects/                 # Organized project data
│   └── {project-code}_{project-name}/
│       ├── raw/                 # Original source files (untouched)
│       ├── extracted/           # Data extracted from source files
│       ├── samples/             # Sample descriptions and metadata
│       └── correspondence/      # Emails and communications
├── 02_scripts/                  # Extraction and processing scripts
├── 03_templates/                # Document templates
├── 04_exports/                  # Final consolidated outputs
└── metadata/                    # Central metadata index
```

## Naming Conventions

### Project Folders
Format: `{project-code}_{project-name}`
- Project code: 3-letter abbreviation + 2-digit number (e.g., `WAT01`, `SOI02`, `AIR03`)
- Project name: lowercase, hyphens for spaces
- Examples:
  - `WAT01_river-sampling-2024`
  - `SOI02_contamination-assessment`
  - `CHM03_groundwater-analysis`

### Files
Format: `{date}_{project-code}_{doctype}_{description}.{ext}`

**Date**: `YYYYMMDD` or `YYYY` if only year known

**Document Types (doctype)**:
| Code | Meaning |
|------|---------|
| `DATA` | Raw data files (Excel with measurements) |
| `SAMP` | Sample information/descriptions |
| `RPRT` | Reports and analysis documents |
| `CORR` | Correspondence (emails, letters) |
| `CERT` | Certificates and lab reports |
| `META` | Metadata and index files |
| `MISC` | Miscellaneous |

**Examples**:
- `20240315_WAT01_DATA_ph-measurements.xlsx`
- `20240320_WAT01_SAMP_collection-sites.pdf`
- `20240401_SOI02_CORR_lab-results-email.pdf`
- `2023_CHM03_RPRT_annual-summary.pdf`

### Extracted Data Files
Format: `{original-filename}_extracted.csv`

## Quick Start

1. Drop new files into `00_inbox/`
2. Run the intake script to register and organize
3. Run extraction scripts on data files
4. Find consolidated data in `04_exports/`

## Project Codes Registry

See `metadata/projects.csv` for the master list of all 27 projects.
