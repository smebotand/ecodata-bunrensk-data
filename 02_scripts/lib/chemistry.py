"""
Chemistry utilities for normalizing and processing chemical analysis data.
"""

import pandas as pd
import re


# =============================================================================
# STANDARD PARAMETER DEFINITIONS
# =============================================================================

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
    
    # PAH - 16 EPA COMPOUNDS (codes match thresholds.py)
    r'Naftalen(?!-)': ('Naftalen', 'Naftalen', 'mg/kg'),
    r'Acenaftylen': ('Acenaftylen', 'Acenaftylen', 'mg/kg'),
    r'Acenaften(?!-)': ('Acenaften', 'Acenaften', 'mg/kg'),
    r'Fluoren(?!-)': ('Fluoren', 'Fluoren', 'mg/kg'),
    r'Fenantren': ('Fenantren', 'Fenantren', 'mg/kg'),
    r'Antracen(?!-)': ('Antracen', 'Antracen', 'mg/kg'),
    r'Fluoranten': ('Fluoranten', 'Fluoranten', 'mg/kg'),
    r'Pyren(?!-)': ('Pyren', 'Pyren', 'mg/kg'),
    r'Benso\(a\)antracen': ('Benzo(a)antracen', 'Benzo(a)antracen', 'mg/kg'),
    r'Krysen': ('Krysen', 'Krysen', 'mg/kg'),
    r'Benso\(b\)fluoranten': ('Benzo(b)fluoranten', 'Benzo(b)fluoranten', 'mg/kg'),
    r'Benso\(k\)fluoranten': ('Benzo(k)fluoranten', 'Benzo(k)fluoranten', 'mg/kg'),
    r'Benso\(a\)pyren': ('BaP', 'Benzo(a)pyren', 'mg/kg'),
    r'Dibenso\(ah\)antracen': ('Dibenzo(a,h)antracen', 'Dibenzo(a,h)antracen', 'mg/kg'),
    r'Benso\(ghi\)perylen': ('Benzo(ghi)perylen', 'Benzo(ghi)perylen', 'mg/kg'),
    r'Indeno\(1,?2,?3-?c,?d\)pyren': ('Indeno(1,2,3-cd)pyren', 'Indeno(1,2,3-cd)pyren', 'mg/kg'),
    r'Sum\s*PAH-16': ('PAH16', 'Sum PAH-16', 'mg/kg'),
    
    # BTEX (codes match thresholds.py)
    r'Bensen(?!o)': ('Benzen', 'Benzen', 'mg/kg'),
    r'Toluen': ('Toluen', 'Toluen', 'mg/kg'),
    r'Etylbensen': ('Etylbenzen', 'Etylbenzen', 'mg/kg'),
    r'Xylener': ('Xylen', 'Xylener', 'mg/kg'),
    r'Sum\s*BTEX': ('BTEX', 'Sum BTEX', 'mg/kg'),
    
    # THC FRACTIONS
    r'Fraksjon\s*>?\s*C5-C6': ('THC C5-C6', 'Fraksjon >C5-C6', 'mg/kg'),
    r'Fraksjon\s*>?\s*C6-C8': ('THC C6-C8', 'Fraksjon >C6-C8', 'mg/kg'),
    r'Fraksjon\s*>?\s*C8-C10': ('THC C8-C10', 'Fraksjon >C8-C10', 'mg/kg'),
    r'Fraksjon\s*>?\s*C10-C12': ('THC C10-C12', 'Fraksjon >C10-C12', 'mg/kg'),
    r'Fraksjon\s*>?\s*C12-C16': ('THC C12-C16', 'Fraksjon >C12-C16', 'mg/kg'),
    r'Fraksjon\s*>?\s*C16-C35': ('THC C16-C35', 'Fraksjon >C16-C35', 'mg/kg'),
    r'Sum\s*>?\s*C12-C35': ('THC C12-C35', 'Sum >C12-C35', 'mg/kg'),
    
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
    # =========================================================================
    # PETROLEUM HYDROKARBONER - THC-fraksjoner (ALS-metode)
    # =========================================================================
    'fraksjon >c5-c35 (sum)': 'THC C5-C35',
    'fraksjon >c5-c35': 'THC C5-C35',
    'fraksjon >c5-c35 (sum, norm, m1)': 'THC C5-C35',

    # Prosjektspesifikk: C18/fytan
    'c18/fytan': 'C18/Fytan',

    # =========================================================================
    # PAH - Polysykliske aromatiske hydrokarboner
    # =========================================================================
    'sum pah carcinogene': 'Sum PAH carcinogene',

    # =========================================================================
    # FYSISKE PARAMETRE OG ANDRE
    # =========================================================================
    'løst organisk karbon': 'Løst organisk karbon (DOC)',
    'løst organisk karbon (doc)': 'Løst organisk karbon (DOC)',
    'temperatur': 'Temperatur',
    'mengde innveid': 'Mengde innveid',
    'volum eluat l/s = 10': 'Volum eluat (L/S = 10)',
    'volum tilsatt': 'Volum tilsatt',
    'ph-verdi': 'pH',
    'gjennomsnittlig flow': 'Gjennomsnittlig flow',
    'høyde av materiale i kolonnen': 'Høyde av materiale i kolonnen',
    'indre diameter i kolonnen': 'Indre diameter i kolonnen',
    'mengde tørt materiale i kolonne': 'Mengde tørt materiale i kolonne',
    'ph av første 15 ml': 'pH (første 15 mL)',
    'ph av rest l/s=0.1': 'pH (L/S=0.1)',
        
    # =========================================================================
    # METALS (Norwegian, English, Symbol)
    # =========================================================================
    'arsen': 'As', 'arsenic': 'As', 'as': 'As',
    'bly': 'Pb', 'lead': 'Pb', 'pb': 'Pb',
    'kadmium': 'Cd', 'cadmium': 'Cd', 'cd': 'Cd',
    'kobber': 'Cu', 'kopper': 'Cu', 'Kopper': 'Cu', 'copper': 'Cu', 'cu': 'Cu',
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
    
    # =========================================================================
    # PETROLEUM HYDROCARBONS - THC fractions (ALS method)
    # Format: 'THC C5-C6' (space, not underscore) to match thresholds.py
    # =========================================================================
    'thc': 'THC', 'total hydrocarbons': 'THC', 'totale hydrokarboner': 'THC',
    'tph': 'TPH', 'total petroleum hydrocarbons': 'TPH',
    'olje': 'TPH', 'oil': 'TPH', 'mineralsk olje': 'TPH',
    'thc >c10-c12': 'THC C10-C12', 'thc >c12-c16': 'THC C12-C16',
    'thc >c16-c35': 'THC C16-C35', 'thc >c12-c35': 'THC C12-C35',
    'thc >c35-c40': 'THC C35-C40', 'thc >c10-c40': 'THC C10-C40',
    'thc c10-c40': 'THC C10-C40', 'sum thc >c12-c35': 'THC C12-C35',
    # Fraksjon format (ALS PDF)
    'fraksjon >c5-c6': 'THC C5-C6', 'fraksjon >c6-c8': 'THC C6-C8',
    'fraksjon >c8-c10': 'THC C8-C10', 'fraksjon >c10-c12': 'THC C10-C12',
    'fraksjon >c12-c16': 'THC C12-C16', 'fraksjon >c16-c35': 'THC C16-C35',
    'fraksjon >c35-c40': 'THC C35-C40', 'fraksjon >c12-c35': 'THC C12-C35',
    'fraksjon >c12-c35 (sum)': 'THC C12-C35', 'sum >c10-c40': 'THC C10-C40',
    # Olje format aliases
    'olje >c10-c12': 'THC C10-C12', 'olje >c12-c16': 'THC C12-C16',
    'olje >c16-c35': 'THC C16-C35', 'olje sum >c12-c35': 'THC C12-C35',
    'olje (sum >c10-c40)': 'THC C10-C40', 'sum thc (>c10-c40)': 'THC C10-C40',
    'olje, fraksjon >c10-c40': 'THC C10-C40',

    'c17/pristan': 'C17/Pristan',
    'c18/fytan': 'C18/Fytan',
    
    
    # =========================================================================
    # PETROLEUM HYDROCARBONS - Alifater fractions (Eurofins method)
    # Format: 'Alifater C5-C6' (space, not underscore) to match thresholds.py
    # =========================================================================
    'alifater c5-c6': 'Alifater C5-C6', 'alifater >c5-c6': 'Alifater C5-C6',
    'alifater >c6-c8': 'Alifater C6-C8', 'alifater c6-c8': 'Alifater C6-C8',
    'alifater >c8-c10': 'Alifater C8-C10', 'alifater c8-c10': 'Alifater C8-C10',
    'alifater >c10-c12': 'Alifater C10-C12', 'alifater c10-c12': 'Alifater C10-C12',
    'alifater >c12-c16': 'Alifater C12-C16', 'alifater c12-c16': 'Alifater C12-C16',
    'alifater >c16-c35': 'Alifater C16-C35', 'alifater c16-c35': 'Alifater C16-C35',
    'alifater >c35-c40': 'Alifater C35-C40', 'alifater c35-c40': 'Alifater C35-C40',
    'alifater >c12-c35': 'Alifater C12-C35', 'alifater c5-c35': 'Alifater C5-C35',
    'alifater c10-c40': 'Alifater C10-C40',
    'sum alifater > c12-c35': 'Alifater C12-C35',
    # Sum formats
    'sum alifater c5-c35': 'Alifater C5-C35', 'sum alifater >c5-c35': 'Alifater C5-C35',
    'sum alifater c12-c35': 'Alifater C12-C35', 'sum alifater >c12-c35': 'Alifater C12-C35',
    'sum alifater c10-c40': 'Alifater C10-C40',
    'sum alifater >c5-c35 (m1)': 'Alifater C5-C35', 'sum alifater >c12-c35 (m1)': 'Alifater C12-C35',
    
    # =========================================================================
    # AROMATIC HYDROCARBONS
    # =========================================================================
    'aromater >c8-c10': 'Aromatics C8-C10',
    'aromater >c10-c16': 'Aromatics C10-C16',
    'aromater >c16-c35': 'Aromatics C16-C35',
    
    # =========================================================================
    # PAH - Polycyclic Aromatic Hydrocarbons (16 EPA)
    # Codes match thresholds.py (Norwegian names)
    # =========================================================================
    'pah': 'PAH', 'polycyclic aromatic hydrocarbons': 'PAH',
    # Sum PAH16 variations
    'pah16': 'PAH16', 'pah 16': 'PAH16', 'pah-16': 'PAH16',
    'sum pah16': 'PAH16', 'sum pah-16': 'PAH16', 'sum pah 16': 'PAH16',
    'pah 16 epa': 'PAH16', 'sum pah 16 epa': 'PAH16', 'sum pah(16) epa': 'PAH16',
    'sum of 16 pah (m1)': 'PAH16',
    # Individual PAH compounds - Norwegian/English -> Norwegian code
    'naftalen': 'Naftalen', 'naphthalene': 'Naftalen',
    'acenaftylen': 'Acenaftylen', 'acenaphthylene': 'Acenaftylen',
    'acenaften': 'Acenaften', 'acenaphthene': 'Acenaften',
    'fluoren': 'Fluoren', 'fluorene': 'Fluoren',
    'fenantren': 'Fenantren', 'phenanthrene': 'Fenantren',
    'antracen': 'Antracen', 'anthracene': 'Antracen',
    'fluoranten': 'Fluoranten', 'fluoranthene': 'Fluoranten',
    'pyren': 'Pyren', 'pyrene': 'Pyren',
    'krysen': 'Krysen', 'chrysene': 'Krysen', 'krysen/trifenylen': 'Krysen',
    # Benzo compounds - parentheses format -> Norwegian code
    'benzo(a)antracen': 'Benzo(a)antracen', 'benzo(b)fluoranten': 'Benzo(b)fluoranten', 'benzo(k)fluoranten': 'Benzo(k)fluoranten',
    'benzo(a)pyren': 'BaP', 'benzo(ghi)perylen': 'Benzo(ghi)perylen',
    'indeno(1,2,3-cd)pyren': 'Indeno(1,2,3-cd)pyren', 'dibenzo(a,h)antracen': 'Dibenzo(a,h)antracen',
    'benzo(b,k)fluoranten': 'Benzo(b,k)fluoranten',
    # Benzo compounds - bracket format -> Norwegian code
    'benzo[a]antracen': 'Benzo(a)antracen', 'benzo[b]fluoranten': 'Benzo(b)fluoranten', 'benzo[k]fluoranten': 'Benzo(k)fluoranten',
    'benzo[a]pyren': 'BaP', 'benzo[ghi]perylen': 'Benzo(ghi)perylen',
    'indeno[1,2,3-cd]pyren': 'Indeno(1,2,3-cd)pyren', 'dibenzo[a,h]antracen': 'Dibenzo(a,h)antracen',
    # English bracket format -> Norwegian code
    'benzo[a]anthracene': 'Benzo(a)antracen', 'benzo[b]fluoranthene': 'Benzo(b)fluoranten', 'benzo[k]fluoranthene': 'Benzo(k)fluoranten',
    'benzo[a]pyrene': 'BaP', 'benzo[ghi]perylene': 'Benzo(ghi)perylen',
    'methylchrysener/benzo(a)anthracener': 'Benzo(a)antracen/Methylchrysener',
    'indeno[1,2,3-cd]pyrene': 'Indeno(1,2,3-cd)pyren', 'dibenz[a,h]anthracene': 'Dibenzo(a,h)antracen',
    # ALS format (benso instead of benzo) -> Norwegian code
    'benso(a)pyren': 'BaP', 'benso(a)antracen': 'Benzo(a)antracen', 'benso(k)fluoranten': 'Benzo(k)fluoranten',
    'benso(ghi)perylen': 'Benzo(ghi)perylen', 'dibenso(ah)antracen': 'Dibenzo(a,h)antracen',
    'indeno(123cd)pyren': 'Indeno(1,2,3-cd)pyren',
    'benso(b+j)fluoranten': 'Benzo(b+j)fluoranten', 'sum av benso(b+j)fluoranten': 'Benzo(b+j)fluoranten_sum',
    # Custom aliases for unmapped parameters
    'methylpyrene/fluoranthense': 'Methylpyren/Fluoranthen',
    'oljetype < c10': 'Oljetype <C10',
    'oljetype > c10': 'Oljetype >C10',
    'sum karsinogene pah': 'Sum karsinogene PAH',
    'sum karsinogene pah2': 'Sum karsinogene PAH2',
    'sum pah(16) epa3': 'PAH16-EPA3',
    
    # =========================================================================
    # PCB - Polychlorinated Biphenyls (7 Dutch)
    # =========================================================================
    'pcb': 'PCB', 'polychlorinated biphenyls': 'PCB',
    # Sum PCB7 variations
    'pcb7': 'PCB7', 'pcb 7': 'PCB7', 'pcb-7': 'PCB7',
    'sum pcb7': 'PCB7', 'sum pcb-7': 'PCB7', 'sum pcb 7': 'PCB7',
    'sum 7 pcb': 'PCB7', 'pcb 7 (seven dutch)': 'PCB7',
    # Individual congeners
    'pcb 28': 'PCB28', 'pcb28': 'PCB28',
    'pcb 52': 'PCB52', 'pcb52': 'PCB52',
    'pcb 101': 'PCB101', 'pcb101': 'PCB101',
    'pcb 118': 'PCB118', 'pcb118': 'PCB118',
    'pcb 138': 'PCB138', 'pcb138': 'PCB138',
    'pcb 153': 'PCB153', 'pcb153': 'PCB153',
    'pcb 180': 'PCB180', 'pcb180': 'PCB180',
    
    # =========================================================================
    # BTEX - Benzen, Toluen, Etylbenzen, Xylen (Norwegian codes match thresholds.py)
    # =========================================================================
    'btex': 'BTEX', 'sum btex': 'BTEX', 'sum btex (m1)': 'BTEX',
    'benzen': 'Benzen', 'benzene': 'Benzen',
    'toluen': 'Toluen', 'toluene': 'Toluen',
    'etylbenzen': 'Etylbenzen', 'etylbensen': 'Etylbenzen', 'ethylbenzene': 'Etylbenzen',
    # Xylenes - total of all isomers (o-, m-, p-xylene)
    'xylen': 'Xylen', 'xylene': 'Xylen', 'xylener': 'Xylen', 'xylenes': 'Xylen',
    'sum xylener': 'Xylen', 'sum xylener (m1)': 'Xylen', 'm/p/o-xylen': 'Xylen',
    # Xylene isomers - m+p often co-elute in GC, reported together
    'm,p-xylen': 'Xylen_mp', 'm/p-xylener': 'Xylen_mp', 'm,p-xylene': 'Xylen_mp',
    'o-xylen': 'Xylen_o', 'o-xylene': 'Xylen_o',
    
    # =========================================================================
    # PFAS - Per- and Polyfluoroalkyl Substances
    # =========================================================================
    'pfas': 'PFAS', 'sum pfas': 'PFAS',
    # PFOS - Perfluorooctanesulfonic acid (distinct from PFOA)
    'pfos': 'PFOS', 'perfluoroktansulfonsyre': 'PFOS', 'perfluorooctanesulfonic acid': 'PFOS',
    # PFOA - Perfluorooctanoic acid (distinct from PFOS)
    'pfoa': 'PFOA', 'perfluoroktansyre': 'PFOA', 'perfluorooctanoic acid': 'PFOA',
    
    # =========================================================================
    # PHYSICAL PARAMETERS
    # =========================================================================
    # Dry matter
    'tørrstoff': 'DryMatter', 'dry matter': 'DryMatter', 'ts': 'DryMatter',
    'tørrstoff (ts)': 'DryMatter', 'ts%': 'DryMatter',
    'tørrstoff (ite)': 'DryMatter', 'ite tørrstoff': 'DryMatter',
    'tørrstoff ved 105 grader': 'DryMatter',
    # pH
    'ph': 'pH',
    # Organic content
    'toc': 'TOC', 'total organic carbon': 'TOC', 'totalt organisk karbon': 'TOC',
    'doc': 'DOC',
    'glødetap': 'LOI', 'loss on ignition': 'LOI',
    # Grain size
    'kornfordeling': 'Kornfordeling', 'siktekurve': 'Siktekurve',
    'leire': 'Leire', 'silt': 'Silt', 'sand': 'Sand', 'grus': 'Grus',
    # Water content
    'vanninnhold': 'WaterContent', 'water content': 'WaterContent',
    
    # =========================================================================
    # NITROGEN COMPOUNDS
    # =========================================================================
    'nitrogen': 'N', 'total nitrogen': 'N_total', 'total-n': 'N_total',
    'ammonium': 'NH4', 'nh4': 'NH4', 'ammonium-n': 'NH4',
    'nitrat': 'NO3', 'no3': 'NO3', 'nitrat-n': 'NO3',
    'nitritt': 'NO2', 'no2': 'NO2', 'nitritt-n': 'NO2',
    
    # =========================================================================
    # LEACHING TEST & OTHER INORGANIC PARAMETERS
    # =========================================================================
    # Cyanide
    'cyanid': 'CN', 'cyanide': 'CN', 'cn': 'CN',
    'cyanid total': 'CN', 'cyanid fri': 'CN_free',
    # Anions
    'sulfat': 'SO4', 'sulphate': 'SO4',
    'klorid': 'Cl', 'chloride': 'Cl', 'cl': 'Cl',
    'fluorid': 'F', 'fluoride': 'F', 'f': 'F',
    # Conductivity
    'konduktivitet': 'Conductivity', 'conductivity': 'Conductivity',
    'ledningsevne (konduktivitet)': 'Conductivity', 'elektrisk konduktivitet': 'Conductivity',
    # Other
    'totalt løst stoff': 'TDS', 'tds': 'TDS',
    'fenolindeks': 'Phenol',
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
    Raises:
        ValueError: If parameter name cannot be normalized
    """
    if pd.isna(name):
        raise ValueError(f"Parameter name is NaN or missing. Input: {name!r}")

    name_str = str(name).strip()
    clean = name_str.lower()



    # Conservative: print a simple message if Kolonne1 is encountered
    if clean == 'kolonne1':
        print("[INFO] Kolonne1 encountered as parameter name (likely empty Excel column header). Check source file if unexpected.")

    # Direct lookup first (handles most cases)
    if clean in PARAMETER_ALIASES:
        return PARAMETER_ALIASES[clean]

    # Try with trailing markers removed (^, *)
    clean_stripped = re.sub(r'[\*\^]+$', '', clean).strip()
    if clean_stripped in PARAMETER_ALIASES:
        return PARAMETER_ALIASES[clean_stripped]

    # Try with parenthetical notes removed
    clean_no_parens = re.sub(r'\s*\(.*\)$', '', clean_stripped).strip()
    if clean_no_parens in PARAMETER_ALIASES:
        return PARAMETER_ALIASES[clean_no_parens]

    # If not found, raise error with more context, except for Kolonne1
    if clean == 'kolonne1':
        # Already printed info above, return None to signal skip
        return None
    import traceback
    tb = traceback.extract_stack()
    # Print debug info
    print(f"\n[DEBUG] Unknown parameter name encountered!")
    print(f"  Parameter: '{name_str}' (normalized: '{clean}')")
    print(f"  Raw input: {name!r}")
    print(f"  Cleaned: {clean!r}")
    print(f"  Cleaned (stripped): {clean_stripped!r}")
    print(f"  Cleaned (no parens): {clean_no_parens!r}")
    print(f"  Known aliases: {sorted(PARAMETER_ALIASES.keys())[:10]}... (total {len(PARAMETER_ALIASES)})")
    print(f"  Call stack:")
    for entry in tb[-6:-1]:
        print(f"    File '{entry.filename}', line {entry.lineno}, in {entry.name}")
        print(f"      {entry.line}")
    print(f"  (If called from pandas, check the DataFrame column headers and sample rows for 'Kolonne1'.)")
    #raise ValueError(f"Unknown parameter name: '{name_str}' (normalized: '{clean}')")


