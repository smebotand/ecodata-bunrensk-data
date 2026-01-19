"""
Chemistry utilities for normalizing and processing chemical analysis data.
"""

import pandas as pd
import re
from typing import Dict, List, Optional, Tuple


# =============================================================================
# STANDARD PARAMETER DEFINITIONS
# =============================================================================

# Eurofins Excel column name -> standard parameter code
# This maps the exact column headers from Eurofins lab reports to our standard codes
# Note: Oil/hydrocarbon parameters use ALIFATER method codes (Alifater_*)
EUROFINS_ALI_COLUMN_MAP = {
    # Physical parameters
    'Tørrstoff': 'DryMatter',
    'Tørrstoff (ite)': 'DryMatter',
    'ite Tørrstoff': 'DryMatter',
    'pH': 'pH',
    
    # Metals
    'Arsen (As)': 'As',
    'Bly (Pb)': 'Pb',
    'Kadmium (Cd)': 'Cd',
    'Kobber (Cu)': 'Cu',
    'Krom (Cr)': 'Cr',
    'Krom total (Cr)': 'Cr',
    'Kvikksølv (Hg)': 'Hg',
    'Nikkel (Ni)': 'Ni',
    'Sink (Zn)': 'Zn',
    'Krom (VI)': 'Cr_VI',
    'Krom VI': 'Cr_VI',
    'Jern (Fe)': 'Fe',
    'Mangan (Mn)': 'Mn',
    'Barium (Ba)': 'Ba',
    'Vanadium (V)': 'V',
    'Kobolt (Co)': 'Co',
    'Molybden (Mo)': 'Mo',
    'Antimon (Sb)': 'Sb',
    'Selen (Se)': 'Se',
    'Tinn (Sn)': 'Sn',
    'Tallium (Tl)': 'Tl',
    'Aluminium (Al)': 'Al',
    'Beryllium (Be)': 'Be',
    'Litium (Li)': 'Li',
    'Sølv (Ag)': 'Ag',
    'Wolfram (W)': 'W',
    
    # BTEX
    'Benzen': 'Benzene',
    'Toluen': 'Toluene',
    'Etylbenzen': 'Ethylbenzene',
    'm/p/o-Xylen': 'Xylene',
    'm,p-Xylen': 'Xylene_mp',
    'o-Xylen': 'Xylene_o',
    'Xylener': 'Xylene',
    'Sum BTEX': 'BTEX',
    
    # Aliphatic hydrocarbons - Alifater method (Eurofins)
    # Note: Distinct from THC_* codes used by ALS lab
    'Alifater C5-C6': 'Alifater_C5-C6',
    'Alifater >C6-C8': 'Alifater_C6-C8',
    'Alifater >C8-C10': 'Alifater_C8-C10',
    'Alifater >C10-C12': 'Alifater_C10-C12',
    'Alifater >C12-C16': 'Alifater_C12-C16',
    'Alifater >C16-C35': 'Alifater_C16-C35',
    'Alifater C5-C35': 'Alifater_C5-C35',
    'Alifater >C12-C35': 'Alifater_C12-C35',
    'Alifater >C35-C40': 'Alifater_C35-C40',
    'Alifater C10-C40': 'Alifater_C10-C40',
    'Olje >C10-C12': 'Alifater_C10-C12',
    'Olje >C12-C16': 'Alifater_C12-C16',
    'Olje >C16-C35': 'Alifater_C16-C35',
    'Olje sum >C12-C35': 'Alifater_C12-C35',
    'Olje (sum >C10-C40)': 'Alifater_C10-C40',
    'Sum THC (>C10-C40)': 'Alifater_C10-C40',
    
    # Aromatic hydrocarbons
    'Aromater >C8-C10': 'Aromatics_C8-C10',
    'Aromater >C10-C16': 'Aromatics_C10-C16',
    'Aromater >C16-C35': 'Aromatics_C16-C35',
    
    # PAH (16 EPA)
    'Naftalen': 'Naphthalene',
    'Acenaftylen': 'Acenaphthylene',
    'Acenaften': 'Acenaphthene',
    'Fluoren': 'Fluorene',
    'Fenantren': 'Phenanthrene',
    'Antracen': 'Anthracene',
    'Fluoranten': 'Fluoranthene',
    'Pyren': 'Pyrene',
    'Benzo[a]antracen': 'BaA',
    'Benzo(a)antracen': 'BaA',
    'Krysen': 'Chrysene',
    'Krysen/Trifenylen': 'Chrysene',
    'Benzo[b]fluoranten': 'BbF',
    'Benzo(b)fluoranten': 'BbF',
    'Benzo[k]fluoranten': 'BkF',
    'Benzo(k)fluoranten': 'BkF',
    'Benzo(b,k)fluoranten': 'BbkF',
    'Benzo[a]pyren': 'BaP',
    'Benzo(a)pyren': 'BaP',
    'Indeno[1,2,3-cd]pyren': 'IcdP',
    'Indeno(1,2,3-cd)pyren': 'IcdP',
    'Dibenzo[a,h]antracen': 'DahA',
    'Dibenzo(a,h)antracen': 'DahA',
    'Benzo[ghi]perylen': 'BghiP',
    'Benzo(ghi)perylen': 'BghiP',
    'Sum PAH(16) EPA': 'PAH16',
    'Sum PAH-16': 'PAH16',
    'Sum PAH 16': 'PAH16',
    'PAH 16 EPA': 'PAH16',
    
    # PCB (7 Dutch)
    'PCB 28': 'PCB28',
    'PCB 52': 'PCB52',
    'PCB 101': 'PCB101',
    'PCB 118': 'PCB118',
    'PCB 138': 'PCB138',
    'PCB 153': 'PCB153',
    'PCB 180': 'PCB180',
    'Sum 7 PCB': 'PCB7',
    'Sum PCB-7': 'PCB7',
    'Sum PCB 7': 'PCB7',
    
    # Other
    'Cyanid': 'CN',
    'Cyanid total': 'CN',
    'Cyanid fri': 'CN_free',
    'TOC': 'TOC',
    'Glødetap': 'LOI',
}


# ALS Excel column name -> standard parameter code
# Maps column headers from ALS lab reports (Excel export format) to standard codes
# Note: ALS uses "Element (Norwegian)" format vs Eurofins "Norwegian (Element)"
ALS_EXCEL_COLUMN_MAP = {
    # Metals
    'As (Arsen)': 'As',
    'Cd (Kadmium)': 'Cd',
    'Cr (Krom)': 'Cr',
    'Cu (Kopper)': 'Cu',
    'Hg (Kvikksølv)': 'Hg',
    'Ni (Nikkel)': 'Ni',
    'Pb (Bly)': 'Pb',
    'Zn (Sink)': 'Zn',
    'Ba (Barium)': 'Ba',
    'Mo (Molybden)': 'Mo',
    'Sb (Antimon)': 'Sb',
    'Se (Selen)': 'Se',
    'V (Vanadium)': 'V',
    'Co (Kobolt)': 'Co',
    'Fe (Jern)': 'Fe',
    'Mn (Mangan)': 'Mn',
    
    # PCB
    'Sum PCB-7': 'PCB7',
    'PCB 28': 'PCB28',
    'PCB 52': 'PCB52',
    'PCB 101': 'PCB101',
    'PCB 118': 'PCB118',
    'PCB 138': 'PCB138',
    'PCB 153': 'PCB153',
    'PCB 180': 'PCB180',
    
    # PAH
    'Benso(a)pyren^': 'BaP',
    'Benso(a)pyren': 'BaP',
    'Sum PAH-16': 'PAH16',
    'Sum of 16 PAH (M1)': 'PAH16',
    'Naftalen': 'Naphthalene',
    'Acenaftylen': 'Acenaphthylene',
    'Acenaften': 'Acenaphthene',
    'Fluoren': 'Fluorene',
    'Fenantren': 'Phenanthrene',
    'Antracen': 'Anthracene',
    'Fluoranten': 'Fluoranthene',
    'Pyren': 'Pyrene',
    'Benso(a)antracen^': 'BaA',
    'Krysen^': 'Chrysene',
    'Benso(b+j)fluoranten^': 'BbjF',
    'Sum av benso(b+j)fluoranten': 'BbjF_sum',
    'Benso(k)fluoranten^': 'BkF',
    'Indeno(123cd)pyren^': 'IcdP',
    'Dibenso(ah)antracen^': 'DahA',
    'Benso(ghi)perylen': 'BghiP',
    
    # BTEX
    'Benzen': 'Benzene',
    'Toluen': 'Toluene',
    'Etylbensen': 'Ethylbenzene',
    'Xylener': 'Xylenes',
    'm/p-Xylener': 'Xylenes_mp',
    'o-Xylen': 'Xylene_o',
    'Sum xylener (M1)': 'Xylenes_sum',
    'Sum BTEX (M1)': 'BTEX_sum',
    
    # Aliphatic hydrocarbons (Alifater)
    'Alifater >C5-C6': 'Alifater_C5-C6',
    'Alifater >C6-C8': 'Alifater_C6-C8',
    'Alifater >C8-C10': 'Alifater_C8-C10',
    'Alifater >C10-C12': 'Alifater_C10-C12',
    'Alifater C10-C12': 'Alifater_C10-C12',
    'Alifater >C12-C16': 'Alifater_C12-C16',
    'Alifater >C16-C35': 'Alifater_C16-C35',
    'Sum alifater >C5-C35 (M1)': 'Alifater_C5-C35_sum',
    'Sum alifater >C12-C35 (M1)': 'Alifater_C12-C35_sum',
    
    # Oil fractions
    'Fraksjon >C5-C6': 'Fraksjon_C5-C6',
    'Fraksjon >C6-C8': 'Fraksjon_C6-C8',
    'Fraksjon >C8-C10': 'Fraksjon_C8-C10',
    'Fraksjon >C10-C12': 'Fraksjon_C10-C12',
    'Fraksjon >C12-C16': 'Fraksjon_C12-C16',
    'Fraksjon >C16-C35': 'Fraksjon_C16-C35',
    'Fraksjon >C35-C40': 'Fraksjon_C35-C40',
    'Fraksjon >C12-C35 (sum)': 'Fraksjon_C12-C35_sum',
    'Fraksjon >C12-C35 (sum, M1)': 'Fraksjon_C12-C35_sum',
    'Sum >C10-C40': 'Sum_C10-C40',
    'Olje, Fraksjon >C10-C40': 'Olje_C10-C40',
    
    # Leaching test parameters
    'Cl': 'Cl',
    'F': 'F',
    'Fenolindeks': 'Phenol',
    'DOC': 'DOC',
    'Sulfat': 'SO4',
    'Ledningsevne (konduktivitet)': 'Conductivity',
    'Elektrisk konduktivitet': 'Conductivity',
    'pH': 'pH',
    
    # Physical
    'Tørrstoff ved 105 grader': 'DryMatter',
    'Vanninnhold': 'WaterContent',
    'TOC': 'TOC',
}


# Eurofins leaching test parameter mapping (for ristetest L/S=10 and kolonnetest L/S=0,1)
# Maps Norwegian parameter names from PDF reports to standard codes
EUROFINS_LEACHING_PARAM_MAP = {
    'Arsen (As)': 'As',
    'Barium (Ba)': 'Ba',
    'Kadmium (Cd)': 'Cd',
    'Krom (Cr)': 'Cr',
    'Kobber (Cu)': 'Cu',
    'Kvikksølv (Hg)': 'Hg',
    'Molybden (Mo)': 'Mo',
    'Nikkel (Ni)': 'Ni',
    'Bly (Pb)': 'Pb',
    'Antimon (Sb)': 'Sb',
    'Selen (Se)': 'Se',
    'Vanadium (V)': 'V',
    'Sink (Zn)': 'Zn',
    'Jern (Fe)': 'Fe',
    'Klorid': 'Cl',
    'Fluorid': 'F',
    'Sulfat': 'SO4',
    'Fenolindeks': 'Phenol',
    'DOC': 'DOC',
    'Totalt løst stoff (TDS)': 'TDS',
    'Total tørrstoff': 'DryMatter',
    'pH': 'pH',
    'Konduktivitet': 'Conductivity',
}


# ALS PDF regex patterns -> (param_code, display_name, default_unit)
# Used for parsing ALS lab reports from PDF text extraction
# Note: Oil/hydrocarbon parameters use THC method codes (THC_*)
# Format: regex pattern -> (standard_code, raw_name_for_display, unit)
ALS_THC_PDF_MAP = {
    # Dry matter (%)
    r'Tørrstoff\s*\([ED]\)': ('DryMatter', 'Tørrstoff', '%'),
    
    # METALS
    r'As\s*\(Arsen\)': ('As', 'Arsen', 'mg/kg'),
    r'Cd\s*\(Kadmium\)': ('Cd', 'Kadmium', 'mg/kg'),
    r'Cr\s*\(Krom\)': ('Cr', 'Krom', 'mg/kg'),
    r'Cu\s*\(Kopper\)': ('Cu', 'Kopper', 'mg/kg'),
    r'Hg\s*\(Kvikksølv\)': ('Hg', 'Kvikksølv', 'mg/kg'),
    r'Ni\s*\(Nikkel\)': ('Ni', 'Nikkel', 'mg/kg'),
    r'Pb\s*\(Bly\)': ('Pb', 'Bly', 'mg/kg'),
    r'Zn\s*\(Sink\)': ('Zn', 'Sink', 'mg/kg'),
    r'Cr6\+': ('Cr_VI', 'Cr6+', 'mg/kg'),
    
    # PCB CONGENERS
    r'PCB\s*28(?!\d)': ('PCB28', 'PCB 28', 'mg/kg'),
    r'PCB\s*52(?!\d)': ('PCB52', 'PCB 52', 'mg/kg'),
    r'PCB\s*101(?!\d)': ('PCB101', 'PCB 101', 'mg/kg'),
    r'PCB\s*118(?!\d)': ('PCB118', 'PCB 118', 'mg/kg'),
    r'PCB\s*138(?!\d)': ('PCB138', 'PCB 138', 'mg/kg'),
    r'PCB\s*153(?!\d)': ('PCB153', 'PCB 153', 'mg/kg'),
    r'PCB\s*180(?!\d)': ('PCB180', 'PCB 180', 'mg/kg'),
    r'Sum\s*PCB-7': ('PCB7', 'Sum PCB-7', 'mg/kg'),
    
    # PAH - 16 EPA COMPOUNDS
    r'Naftalen(?!-)': ('Naphthalene', 'Naftalen', 'mg/kg'),
    r'Acenaftylen': ('Acenaphthylene', 'Acenaftylen', 'mg/kg'),
    r'Acenaften(?!-)': ('Acenaphthene', 'Acenaften', 'mg/kg'),
    r'Fluoren(?!-)': ('Fluorene', 'Fluoren', 'mg/kg'),
    r'Fenantren': ('Phenanthrene', 'Fenantren', 'mg/kg'),
    r'Antracen(?!-)': ('Anthracene', 'Antracen', 'mg/kg'),
    r'Fluoranten': ('Fluoranthene', 'Fluoranten', 'mg/kg'),
    r'Pyren(?!-)': ('Pyrene', 'Pyren', 'mg/kg'),
    r'Benso\(a\)antracen': ('BaA', 'Benzo(a)antracen', 'mg/kg'),
    r'Krysen': ('Chrysene', 'Krysen', 'mg/kg'),
    r'Benso\(b\)fluoranten': ('BbF', 'Benzo(b)fluoranten', 'mg/kg'),
    r'Benso\(k\)fluoranten': ('BkF', 'Benzo(k)fluoranten', 'mg/kg'),
    r'Benso\(a\)pyren': ('BaP', 'Benzo(a)pyren', 'mg/kg'),
    r'Dibenso\(ah\)antracen': ('DahA', 'Dibenzo(ah)antracen', 'mg/kg'),
    r'Benso\(ghi\)perylen': ('BghiP', 'Benzo(ghi)perylen', 'mg/kg'),
    r'Indeno\(1,?2,?3-?c,?d\)pyren': ('IcdP', 'Indeno(123cd)pyren', 'mg/kg'),
    r'Sum\s*PAH-16': ('PAH16', 'Sum PAH-16', 'mg/kg'),
    
    # BTEX
    r'Bensen(?!o)': ('Benzene', 'Benzen', 'mg/kg'),
    r'Toluen': ('Toluene', 'Toluen', 'mg/kg'),
    r'Etylbensen': ('Ethylbenzene', 'Etylbensen', 'mg/kg'),
    r'Xylener': ('Xylene', 'Xylener', 'mg/kg'),
    r'Sum\s*BTEX': ('BTEX', 'Sum BTEX', 'mg/kg'),
    
    # THC FRACTIONS
    r'Fraksjon\s*>?\s*C5-C6': ('THC_C5-C6', 'Fraksjon >C5-C6', 'mg/kg'),
    r'Fraksjon\s*>?\s*C6-C8': ('THC_C6-C8', 'Fraksjon >C6-C8', 'mg/kg'),
    r'Fraksjon\s*>?\s*C8-C10': ('THC_C8-C10', 'Fraksjon >C8-C10', 'mg/kg'),
    r'Fraksjon\s*>?\s*C10-C12': ('THC_C10-C12', 'Fraksjon >C10-C12', 'mg/kg'),
    r'Fraksjon\s*>?\s*C12-C16': ('THC_C12-C16', 'Fraksjon >C12-C16', 'mg/kg'),
    r'Fraksjon\s*>?\s*C16-C35': ('THC_C16-C35', 'Fraksjon >C16-C35', 'mg/kg'),
    r'Sum\s*>?\s*C12-C35': ('THC_C12-C35', 'Sum >C12-C35', 'mg/kg'),
    
    # CYANIDE
    r'Cyanid-fri': ('CN', 'Cyanid-fri', 'mg/kg'),
    
    # CHLOROPHENOLS
    r'2-Monoklorfenol': ('CP_2-MCP', '2-Monoklorfenol', 'mg/kg'),
    r'3-Monoklorfenol': ('CP_3-MCP', '3-Monoklorfenol', 'mg/kg'),
    r'4-Monoklorfenol': ('CP_4-MCP', '4-Monoklorfenol', 'mg/kg'),
    r'2,3-Diklorfenol': ('CP_2,3-DCP', '2,3-Diklorfenol', 'mg/kg'),
    r'2,4\+2,5-Diklorfenol': ('CP_2,4+2,5-DCP', '2,4+2,5-Diklorfenol', 'mg/kg'),
    r'2,6-Diklorfenol': ('CP_2,6-DCP', '2,6-Diklorfenol', 'mg/kg'),
    r'3,4-Diklorfenol': ('CP_3,4-DCP', '3,4-Diklorfenol', 'mg/kg'),
    r'3,5-Diklorfenol': ('CP_3,5-DCP', '3,5-Diklorfenol', 'mg/kg'),
    r'2,3,4-Triklorfenol': ('CP_2,3,4-TCP', '2,3,4-Triklorfenol', 'mg/kg'),
    r'2,3,5-Triklorfenol': ('CP_2,3,5-TCP', '2,3,5-Triklorfenol', 'mg/kg'),
    r'2,3,6-Triklorfenol': ('CP_2,3,6-TCP', '2,3,6-Triklorfenol', 'mg/kg'),
    r'2,4,5-Triklorfenol': ('CP_2,4,5-TCP', '2,4,5-Triklorfenol', 'mg/kg'),
    r'2,4,6-Triklorfenol': ('CP_2,4,6-TCP', '2,4,6-Triklorfenol', 'mg/kg'),
    r'3,4,5-Triklorfenol': ('CP_3,4,5-TCP', '3,4,5-Triklorfenol', 'mg/kg'),
    r'2,3,4,5-Tetraklorfenol': ('CP_2,3,4,5-TeCP', '2,3,4,5-Tetraklorfenol', 'mg/kg'),
    r'2,3,4,6-Tetraklorfenol': ('CP_2,3,4,6-TeCP', '2,3,4,6-Tetraklorfenol', 'mg/kg'),
    r'2,3,5,6-Tetraklorfenol': ('CP_2,3,5,6-TeCP', '2,3,5,6-Tetraklorfenol', 'mg/kg'),
    r'Pentaklorfenol': ('CP_PCP', 'Pentaklorfenol', 'mg/kg'),
    
    # CHLOROBENZENES
    r'Monoklorbensen': ('CB_MCB', 'Monoklorbensen', 'mg/kg'),
    r'1,2-Diklorbensen': ('CB_1,2-DCB', '1,2-Diklorbensen', 'mg/kg'),
    r'1,4-Diklorbensen': ('CB_1,4-DCB', '1,4-Diklorbensen', 'mg/kg'),
    r'1,2,3-Triklorbensen': ('CB_1,2,3-TCB', '1,2,3-Triklorbensen', 'mg/kg'),
    r'1,2,4-Triklorbensen': ('CB_1,2,4-TCB', '1,2,4-Triklorbensen', 'mg/kg'),
    r'1,3,5-Triklorbensen': ('CB_1,3,5-TCB', '1,3,5-Triklorbensen', 'mg/kg'),
    r'1,2,3,5\+1,2,4,5-Tetraklorbensen': ('CB_TeCB', '1,2,3,5+1,2,4,5-Tetraklorbensen', 'mg/kg'),
    r'Pentaklorbensen': ('CB_PeCB', 'Pentaklorbensen', 'mg/kg'),
    r'Heksaklorbensen': ('CB_HCB', 'Heksaklorbensen', 'mg/kg'),
    
    # CHLORINATED SOLVENTS
    r'Diklormetana?(?!\s*\()': ('VOC_DCM', 'Diklormetana', 'mg/kg'),
    r'Triklormetan\s*\(kloroform\)': ('VOC_Chloroform', 'Triklormetan (kloroform)', 'mg/kg'),
    r'Trikloreten': ('VOC_TCE', 'Trikloreten', 'mg/kg'),
    r'Tetraklormetana?': ('VOC_CCl4', 'Tetraklormetana', 'mg/kg'),
    r'Tetrakloreten': ('VOC_PCE', 'Tetrakloreten', 'mg/kg'),
    r'1,2-Dikloretan': ('VOC_1,2-DCA', '1,2-Dikloretan', 'mg/kg'),
    r'1,1,1-Trikloretan': ('VOC_1,1,1-TCA', '1,1,1-Trikloretan', 'mg/kg'),
    r'1,2-Dibrometan': ('VOC_EDB', '1,2-Dibrometan', 'mg/kg'),
    r'1,1,2-Trikloretan': ('VOC_1,1,2-TCA', '1,1,2-Trikloretan', 'mg/kg'),
    
    # PESTICIDES
    r'g-HCH\s*\(Lindan\)': ('PEST_Lindane', 'g-HCH (Lindan)', 'mg/kg'),
    r"o,p'-DDT": ('PEST_op-DDT', "o,p'-DDT", 'mg/kg'),
    r"p,p'-DDT": ('PEST_pp-DDT', "p,p'-DDT", 'mg/kg'),
    r"o,p'-DDD": ('PEST_op-DDD', "o,p'-DDD", 'mg/kg'),
    r"p,p'-DDD": ('PEST_pp-DDD', "p,p'-DDD", 'mg/kg'),
    r"o,p'-DDE": ('PEST_op-DDE', "o,p'-DDE", 'mg/kg'),
    r"p,p'-DDE": ('PEST_pp-DDE', "p,p'-DDE", 'mg/kg'),
}


# Standard parameter name mappings (Norwegian -> English -> Standard code)
PARAMETER_ALIASES = {
    # Metals
    'arsen': 'As', 'arsenic': 'As', 'as': 'As',
    'bly': 'Pb', 'lead': 'Pb', 'pb': 'Pb',
    'kadmium': 'Cd', 'cadmium': 'Cd', 'cd': 'Cd',
    'kobber': 'Cu', 'copper': 'Cu', 'cu': 'Cu',
    'krom': 'Cr', 'chromium': 'Cr', 'cr': 'Cr', 'krom total': 'Cr',
    'krom (vi)': 'Cr_VI', 'krom vi': 'Cr_VI', 'cr(vi)': 'Cr_VI',
    'kvikksølv': 'Hg', 'mercury': 'Hg', 'hg': 'Hg',
    'nikkel': 'Ni', 'nickel': 'Ni', 'ni': 'Ni',
    'sink': 'Zn', 'zinc': 'Zn', 'zn': 'Zn',
    'jern': 'Fe', 'iron': 'Fe', 'fe': 'Fe',
    'mangan': 'Mn', 'manganese': 'Mn', 'mn': 'Mn',
    'barium': 'Ba', 'ba': 'Ba',
    'vanadium': 'V', 'v': 'V',
    'kobolt': 'Co', 'cobalt': 'Co', 'co': 'Co',
    'molybden': 'Mo', 'molybdenum': 'Mo', 'mo': 'Mo',
    'antimon': 'Sb', 'antimony': 'Sb', 'sb': 'Sb',
    'selen': 'Se', 'selenium': 'Se', 'se': 'Se',
    'tinn': 'Sn', 'tin': 'Sn', 'sn': 'Sn',
    'tallium': 'Tl', 'thallium': 'Tl', 'tl': 'Tl',
    'aluminium': 'Al', 'al': 'Al',
    'beryllium': 'Be', 'be': 'Be',
    'litium': 'Li', 'lithium': 'Li', 'li': 'Li',
    'sølv': 'Ag', 'silver': 'Ag', 'ag': 'Ag',
    'wolfram': 'W', 'tungsten': 'W', 'w': 'W',
    
    # Petroleum hydrocarbons (THC fractions)
    'thc': 'THC', 'total hydrocarbons': 'THC', 'totale hydrokarboner': 'THC',
    'tph': 'TPH', 'total petroleum hydrocarbons': 'TPH',
    'olje': 'THC', 'oil': 'THC', 'mineralsk olje': 'THC',
    'sum >c12-c35': 'THC_C12-C35', 'olje sum >c12-c35': 'THC_C12-C35',
    '>c10-c12': 'THC_C10-C12', 'fraksjon >c10-c12': 'THC_C10-C12',
    '>c12-c16': 'THC_C12-C16', 'fraksjon >c12-c16': 'THC_C12-C16',
    '>c16-c35': 'THC_C16-C35', 'fraksjon >c16-c35': 'THC_C16-C35',
    '>c12-c35': 'THC_C12-C35', 'fraksjon >c12-c35': 'THC_C12-C35',
    '>c35-c40': 'THC_C35-C40', 'fraksjon >c35-c40': 'THC_C35-C40',
    'c10-c40': 'THC_C10-C40', 'thc c10-c40': 'THC_C10-C40',
    
    # PAH (Polycyclic Aromatic Hydrocarbons)
    'pah': 'PAH', 'polycyclic aromatic hydrocarbons': 'PAH',
    'pah16': 'PAH16', 'pah 16': 'PAH16', 'sum pah16': 'PAH16', 'pah-16': 'PAH16',
    'pah 16 epa': 'PAH16', 'sum pah 16 epa': 'PAH16',
    'naftalen': 'Naphthalene', 'naphthalene': 'Naphthalene',
    'acenaftylen': 'Acenaphthylene', 'acenaphthylene': 'Acenaphthylene',
    'acenaften': 'Acenaphthene', 'acenaphthene': 'Acenaphthene',
    'fluoren': 'Fluorene', 'fluorene': 'Fluorene',
    'fenantren': 'Phenanthrene', 'phenanthrene': 'Phenanthrene',
    'antracen': 'Anthracene', 'anthracene': 'Anthracene',
    'fluoranten': 'Fluoranthene', 'fluoranthene': 'Fluoranthene',
    'pyren': 'Pyrene', 'pyrene': 'Pyrene',
    'benzo(a)antracen': 'BaA', 'benzo[a]anthracene': 'BaA',
    'krysen': 'Chrysene', 'chrysene': 'Chrysene',
    'benzo(b)fluoranten': 'BbF', 'benzo[b]fluoranthene': 'BbF',
    'benzo(k)fluoranten': 'BkF', 'benzo[k]fluoranthene': 'BkF',
    'benzo(a)pyren': 'BaP', 'benzo[a]pyrene': 'BaP',
    'indeno(1,2,3-cd)pyren': 'IcdP', 'indeno[1,2,3-cd]pyrene': 'IcdP',
    'dibenzo(a,h)antracen': 'DahA', 'dibenz[a,h]anthracene': 'DahA',
    'benzo(ghi)perylen': 'BghiP', 'benzo[ghi]perylene': 'BghiP',
    
    # PCB (Polychlorinated Biphenyls)
    'pcb': 'PCB', 'polychlorinated biphenyls': 'PCB',
    'pcb7': 'PCB7', 'pcb 7': 'PCB7', 'sum pcb7': 'PCB7', 'pcb-7': 'PCB7',
    'pcb 7 (seven dutch)': 'PCB7',
    'pcb 28': 'PCB28', 'pcb28': 'PCB28',
    'pcb 52': 'PCB52', 'pcb52': 'PCB52',
    'pcb 101': 'PCB101', 'pcb101': 'PCB101',
    'pcb 118': 'PCB118', 'pcb118': 'PCB118',
    'pcb 138': 'PCB138', 'pcb138': 'PCB138',
    'pcb 153': 'PCB153', 'pcb153': 'PCB153',
    'pcb 180': 'PCB180', 'pcb180': 'PCB180',
    
    # BTEX
    'btex': 'BTEX', 'sum btex': 'BTEX',
    'benzen': 'Benzene', 'benzene': 'Benzene',
    'toluen': 'Toluene', 'toluene': 'Toluene',
    'etylbenzen': 'Ethylbenzene', 'ethylbenzene': 'Ethylbenzene',
    'xylen': 'Xylene', 'xylene': 'Xylene', 'xylener': 'Xylene',
    'm,p-xylen': 'Xylene_mp', 'o-xylen': 'Xylene_o',
    
    # Other organics
    'pfas': 'PFAS', 'sum pfas': 'PFAS',
    'pfos': 'PFOS', 'pfoa': 'PFOA',
    
    # Physical parameters
    'tørrstoff': 'DryMatter', 'dry matter': 'DryMatter', 'ts': 'DryMatter',
    'tørrstoff (ts)': 'DryMatter', 'ts%': 'DryMatter',
    'ph': 'pH',
    'toc': 'TOC', 'total organic carbon': 'TOC', 'totalt organisk karbon': 'TOC',
    'glødetap': 'LOI', 'loss on ignition': 'LOI',
    'kornfordeling': 'GrainSize', 'siktekurve': 'GrainSize',
    'leire': 'Clay', 'silt': 'Silt', 'sand': 'Sand', 'grus': 'Gravel',
    
    # Nitrogen compounds
    'nitrogen': 'N', 'total nitrogen': 'N_total', 'total-n': 'N_total',
    'ammonium': 'NH4', 'nh4': 'NH4', 'ammonium-n': 'NH4',
    'nitrat': 'NO3', 'no3': 'NO3', 'nitrat-n': 'NO3',
    'nitritt': 'NO2', 'no2': 'NO2', 'nitritt-n': 'NO2',
    
    # Other
    'cyanid': 'CN', 'cyanide': 'CN', 'cn': 'CN',
    'sulfat': 'SO4', 'sulphate': 'SO4',
    'klorid': 'Cl', 'chloride': 'Cl',
    'fluorid': 'F', 'fluoride': 'F',
}


# Unit conversion factors to mg/kg
UNIT_CONVERSIONS = {
    'mg/kg': 1.0,
    'mg/kg ts': 1.0,
    'mg/kg dw': 1.0,
    'mg/kg tørrstoff': 1.0,
    'µg/kg': 0.001,
    'ug/kg': 0.001,
    'μg/kg': 0.001,
    'µg/kg ts': 0.001,
    'ng/kg': 0.000001,
    'g/kg': 1000.0,
    'ppm': 1.0,
    'ppb': 0.001,
    '%': 10000.0,
    'mg/l': 1.0,  # Approximation for water samples
    'µg/l': 0.001,
}


# Norwegian regulatory limits (Tilstandsklasser) in mg/kg
# Source: TA-2553/2009, Miljødirektoratet (updated with M-1884/2022 for some)
TILSTANDSKLASSER = {
    # Metals
    'As':  {'I': 8,    'II': 8,    'III': 20,   'IV': 50,    'V': 1000},
    'Pb':  {'I': 60,   'II': 60,   'III': 100,  'IV': 300,   'V': 1000},
    'Cd':  {'I': 1.5,  'II': 1.5,  'III': 3,    'IV': 15,    'V': 30},
    'Cu':  {'I': 100,  'II': 100,  'III': 200,  'IV': 1000,  'V': 1000},
    'Cr':  {'I': 50,   'II': 50,   'III': 100,  'IV': 500,   'V': 1000},
    'Cr_VI': {'I': 2,  'II': 2,    'III': 5,    'IV': 20,    'V': 1000},
    'Hg':  {'I': 1,    'II': 1,    'III': 2,    'IV': 10,    'V': 10},
    'Ni':  {'I': 60,   'II': 60,   'III': 100,  'IV': 500,   'V': 1000},
    'Zn':  {'I': 200,  'II': 200,  'III': 500,  'IV': 1000,  'V': 5000},
    'V':   {'I': 100,  'II': 100,  'III': 200,  'IV': 700,   'V': 700},
    'Co':  {'I': 20,   'II': 20,   'III': 50,   'IV': 300,   'V': 1000},
    
    # PAH
    'PAH16':     {'I': 2,    'II': 2,    'III': 8,    'IV': 50,   'V': 1000},
    'BaP':       {'I': 0.1,  'II': 0.1,  'III': 0.5,  'IV': 5,    'V': 100},
    'Naphthalene': {'I': 0.8, 'II': 0.8, 'III': 4,   'IV': 40,   'V': 1000},
    
    # PCB
    'PCB7':      {'I': 0.01, 'II': 0.01, 'III': 0.5,  'IV': 1,    'V': 5},
    
    # Petroleum hydrocarbons
    'THC':       {'I': 50,   'II': 50,   'III': 300,  'IV': 2000,  'V': 5000},
    'THC_C12-C35': {'I': 50, 'II': 50,   'III': 300,  'IV': 2000,  'V': 5000},
    'THC_C16-C35': {'I': 50, 'II': 50,   'III': 300,  'IV': 2000,  'V': 5000},
    
    # BTEX
    'Benzene':   {'I': 0.01, 'II': 0.01, 'III': 0.05, 'IV': 1,    'V': 5},
    
    # Other
    'CN':        {'I': 1,    'II': 1,    'III': 5,    'IV': 100,  'V': 500},
}


# =============================================================================
# LOCATION AND SAMPLE TYPE STANDARDIZATION
# =============================================================================

LOCATION_TYPES = {
    'vegbane': 'vegbane',
    'vegbanen': 'vegbane',
    'kjørebane': 'vegbane',
    'grøft': 'grøft',
    'grøfta': 'grøft',
    'sidegrøft': 'grøft',
    'deponi': 'deponi',
    'mellomlagring': 'deponi',
    'mellomlager': 'deponi',
    'tipp': 'deponi',
    'tunnel': 'tunnel',
    'stuff': 'stuff',
    'påhugg': 'påhugg',
    'portal': 'påhugg',
    'tverrslag': 'tverrslag',
    'sjakt': 'sjakt',
}

SAMPLE_TYPES = {
    'bunnrensk': 'bunnrensk',
    'bunnrenskmasser': 'bunnrensk',
    'sålerensk': 'sålerensk',
    'blandeprøve': 'blandeprøve',
    'blandprøve': 'blandeprøve',
    'stikkprøve': 'stikkprøve',
    'enkeltprøve': 'stikkprøve',
    'referanse': 'referanse',
    'rein stein': 'referanse',
    'ren stein': 'referanse',
}


# =============================================================================
# FUNCTIONS
# =============================================================================


def normalize_parameter_name(name: str) -> str:
    """
    Normalize a parameter name to standard form.
    
    Args:
        name: Raw parameter name from lab report
        
    Returns:
        Standardized parameter code
    """
    if pd.isna(name):
        return ''
    
    name_str = str(name).strip()
    
    # First check exact match in Eurofins column map
    if name_str in EUROFINS_ALI_COLUMN_MAP:
        return EUROFINS_ALI_COLUMN_MAP[name_str]
    
    # Clean and lowercase for alias lookup
    clean = name_str.lower()
    
    # Remove common prefixes/suffixes
    clean = re.sub(r'^(sum\s+|total\s+)', '', clean)
    clean = re.sub(r'\s*\(.*\)$', '', clean)  # Remove parenthetical notes
    
    # Look up in aliases
    if clean in PARAMETER_ALIASES:
        return PARAMETER_ALIASES[clean]
    
    # Return original with title case if not found
    return name_str


def get_parameter_code(column_name: str) -> str:
    """
    Get standard parameter code from Excel column name.
    
    Args:
        column_name: Column header from lab report Excel
        
    Returns:
        Standard parameter code or original name if not found
    """
    if pd.isna(column_name):
        return ''
    
    name_str = str(column_name).strip()
    
    # Check ALS Excel column map first
    if name_str in ALS_EXCEL_COLUMN_MAP:
        return ALS_EXCEL_COLUMN_MAP[name_str]
    
    # Check Eurofins column map
    if name_str in EUROFINS_ALI_COLUMN_MAP:
        return EUROFINS_ALI_COLUMN_MAP[name_str]
    
    # Fall back to normalize function
    return normalize_parameter_name(name_str)


def normalize_location_type(location: str) -> str:
    """Normalize location type to standard form."""
    if pd.isna(location):
        return ''
    clean = str(location).strip().lower()
    return LOCATION_TYPES.get(clean, clean)


def normalize_sample_type(sample_type: str) -> str:
    """Normalize sample type to standard form."""
    if pd.isna(sample_type):
        return ''
    clean = str(sample_type).strip().lower()
    return SAMPLE_TYPES.get(clean, clean)


def parse_value(value: str) -> Tuple[float, bool]:
    """
    Parse a lab result value, handling '<' (below detection limit).
    
    Args:
        value: Value string from lab report
        
    Returns:
        Tuple of (numeric_value, is_below_detection_limit)
    """
    if pd.isna(value):
        return (None, False)
    
    value_str = str(value).strip()
    
    # Check for below detection limit
    below_limit = False
    if value_str.startswith('<') or value_str.startswith('&lt;'):
        below_limit = True
        value_str = re.sub(r'^[<&lt;]+\s*', '', value_str)
    
    # Also handle "n.d.", "nd", "ikke påvist" etc
    if value_str.lower() in ['n.d.', 'nd', 'ikke påvist', 'not detected', '-', 'i.p.']:
        return (0.0, True)
    
    # Extract numeric value
    try:
        # Handle Norwegian decimal comma
        value_str = value_str.replace(',', '.').replace(' ', '')
        # Remove any trailing text like "mg/kg"
        value_str = re.sub(r'[a-zA-Z/]+$', '', value_str)
        numeric = float(value_str)
        return (numeric, below_limit)
    except ValueError:
        return (None, False)


def parse_uncertainty(value: str) -> Optional[float]:
    """
    Parse uncertainty value from formats like "51.9 +/- 10.4" or "±10.4".
    
    Returns:
        Uncertainty value or None
    """
    if pd.isna(value):
        return None
    
    value_str = str(value)
    
    # Look for +/- or ± pattern
    match = re.search(r'[±+/-]+\s*(\d+[.,]?\d*)', value_str)
    if match:
        try:
            return float(match.group(1).replace(',', '.'))
        except ValueError:
            pass
    
    return None


def convert_units(value: float, from_unit: str, to_unit: str = 'mg/kg') -> float:
    """
    Convert between concentration units.
    
    Args:
        value: Numeric value
        from_unit: Source unit
        to_unit: Target unit (default mg/kg)
        
    Returns:
        Converted value
    """
    if value is None:
        return None
    
    from_unit_lower = from_unit.lower().strip()
    to_unit_lower = to_unit.lower().strip()
    
    # Get conversion factors
    from_factor = UNIT_CONVERSIONS.get(from_unit_lower, 1.0)
    to_factor = UNIT_CONVERSIONS.get(to_unit_lower, 1.0)
    
    # Convert: value * from_factor gives mg/kg, then divide by to_factor
    return value * from_factor / to_factor


def normalize_results_df(df: pd.DataFrame,
                         value_col: str = None,
                         parameter_col: str = None,
                         unit_col: str = None) -> pd.DataFrame:
    """
    Normalize a results dataframe with standard parameter names and units.
    
    Args:
        df: Input dataframe
        value_col: Column containing values (auto-detect if None)
        parameter_col: Column containing parameter names (auto-detect if None)
        unit_col: Column containing units (auto-detect if None)
        
    Returns:
        Normalized dataframe with added columns
    """
    df = df.copy()
    
    # Auto-detect columns if not specified
    cols_lower = {c.lower(): c for c in df.columns}
    
    if parameter_col is None:
        for key in ['parameter', 'analyse', 'komponent', 'stoff']:
            if key in cols_lower:
                parameter_col = cols_lower[key]
                break
    
    if value_col is None:
        for key in ['resultat', 'result', 'verdi', 'value', 'konsentrasjon']:
            if key in cols_lower:
                value_col = cols_lower[key]
                break
    
    if unit_col is None:
        for key in ['enhet', 'unit', 'måleenhet']:
            if key in cols_lower:
                unit_col = cols_lower[key]
                break
    
    # Normalize parameter names
    if parameter_col and parameter_col in df.columns:
        df['parameter_std'] = df[parameter_col].apply(normalize_parameter_name)
    
    # Parse values and detect below-limit
    if value_col and value_col in df.columns:
        parsed = df[value_col].apply(parse_value)
        df['value_numeric'] = parsed.apply(lambda x: x[0])
        df['below_limit'] = parsed.apply(lambda x: x[1])
        
        # Convert to mg/kg if unit column exists
        if unit_col and unit_col in df.columns:
            df['value_mg_kg'] = df.apply(
                lambda row: convert_units(row['value_numeric'], str(row[unit_col])) 
                if pd.notna(row['value_numeric']) else None,
                axis=1
            )
    
    return df


def classify_contamination(parameter: str, value: float) -> str:
    """
    Classify a contamination level according to Norwegian standards.
    
    Args:
        parameter: Standard parameter code (e.g., 'As', 'Pb')
        value: Concentration in mg/kg
        
    Returns:
        Tilstandsklasse ('I', 'II', 'III', 'IV', 'V') or 'Unknown'
    """
    if value is None or parameter not in TILSTANDSKLASSER:
        return 'Unknown'
    
    limits = TILSTANDSKLASSER[parameter]
    
    if value <= limits['I']:
        return 'I'
    elif value <= limits['II']:
        return 'II'
    elif value <= limits['III']:
        return 'III'
    elif value <= limits['IV']:
        return 'IV'
    else:
        return 'V'


def get_limiting_parameter(results: Dict[str, float]) -> Tuple[str, float, str]:
    """
    Find the parameter that determines the overall classification.
    
    Args:
        results: Dict of parameter -> value in mg/kg
        
    Returns:
        Tuple of (parameter, value, tilstandsklasse)
    """
    worst_class = 'I'
    worst_param = None
    worst_value = None
    
    class_order = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'Unknown': 0}
    
    for param, value in results.items():
        if value is None:
            continue
        
        classification = classify_contamination(param, value)
        if class_order.get(classification, 0) > class_order.get(worst_class, 0):
            worst_class = classification
            worst_param = param
            worst_value = value
    
    return (worst_param, worst_value, worst_class)


def classify_sample_results(sample_id: str, results: pd.DataFrame) -> Dict:
    """
    Classify a sample based on all its chemical results.
    
    Args:
        sample_id: Sample identifier
        results: DataFrame with columns 'parameter' and 'value' (in mg/kg)
        
    Returns:
        Dict with classification info matching classifications.csv schema
    """
    # Build dict of parameter -> value
    param_values = {}
    for _, row in results.iterrows():
        param = row.get('parameter', row.get('parameter_std', ''))
        value = row.get('value', row.get('value_numeric', row.get('value_mg_kg', None)))
        if param and value is not None:
            param_values[param] = value
    
    # Get limiting parameter
    limiting_param, limiting_value, tilstandsklasse = get_limiting_parameter(param_values)
    
    return {
        'sample_id': sample_id,
        'tilstandsklasse': tilstandsklasse,
        'limiting_parameter': limiting_param,
        'limiting_value': limiting_value,
        'classification_basis': 'TA-2553/2009'
    }


def format_result_for_csv(
    sample_id: str,
    parameter_raw: str,
    value_raw: str,
    unit: str = 'mg/kg',
    loq: float = None
) -> Dict:
    """
    Format a single result for output to results.csv.
    
    Args:
        sample_id: Sample identifier
        parameter_raw: Original parameter name from source
        value_raw: Original value string from source
        unit: Unit of measurement
        loq: Limit of quantification
        
    Returns:
        Dict matching results.csv schema
    """
    parameter_std = normalize_parameter_name(parameter_raw)
    value, below_limit = parse_value(value_raw)
    uncertainty = parse_uncertainty(str(value_raw))
    
    return {
        'sample_id': sample_id,
        'parameter': parameter_std,
        'parameter_raw': parameter_raw,
        'value': value,
        'unit': unit,
        'uncertainty': uncertainty,
        'below_limit': below_limit,
        'loq': loq
    }
