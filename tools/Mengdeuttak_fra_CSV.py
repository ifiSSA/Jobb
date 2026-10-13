# -*- coding: utf-8 -*-
"""
Kjør Mengdeuttak_Armering.py mot en kombinert selection-basket CSV
istedenfor to separate Excel-filer.

Splittelse av ARM vs RIB skjer på Path-kolonnen (se csv_til_armrib.py).

Resultat-Excel (Sammendrag, Manglende_Mengder, Armering_Detaljer,
LM_Oversikt, LM_Omfar, LM_PE_Svinn, Dokumentasjon, Forklaring) skrives
til samme katalog som Mengdeuttak_Armering.py:
    resultat_mengdeuttak_<omega>_<timestamp>.xlsx
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from csv_til_armrib import kjor_mengdeuttak

CSV_PATH   = r"C:\Users\ifi\Downloads\selection-basket (33).csv"
OMEGA_PATH = r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\6900-01-010 K7 - Lysaker og Vækerø Stasjon\APM FOB Gjennomgang\Omega_K7_April.xlsx"

if __name__ == "__main__":
    kjor_mengdeuttak(csv_path=CSV_PATH, omega_path=OMEGA_PATH)
