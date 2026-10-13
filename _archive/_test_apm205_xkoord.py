# -*- coding: utf-8 -*-
"""Raskt test av x-koordinat-ekstraksjon og Omega-korreksjon for APM205."""
import sys
import os
import glob

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from ifi.pdf_maalebrev_til_excel import les_pdf_hoved, les_omega

def finn_fil(pattern):
    treff = glob.glob(pattern, recursive=True)
    return treff[0] if treff else ""

APM205 = finn_fil(r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\**\*APM205*.pdf")
OMEGA  = finn_fil(r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen\**\Omega_K7_Mars.xlsx")

if not APM205:
    print("[ERROR] APM205 ikke funnet")
    sys.exit(1)

print("APM205:", os.path.basename(APM205))
print("Omega: ", os.path.basename(OMEGA) if OMEGA else "(ikke funnet)")

print()
print("=== Test 1: uten omega_poster ===")
rader_uten = les_pdf_hoved(APM205)
print(f"Rader: {len(rader_uten)}")
for r in rader_uten:
    post = r["Post"]
    print(f"  {post}")

print()
if OMEGA:
    print("=== Test 2: med omega_data ===")
    omega_data = les_omega(OMEGA)
    omega_poster = set(omega_data.keys())
    print(f"Omega: {len(omega_poster)} poster")
    print()
    rader_med = les_pdf_hoved(APM205, omega_data=omega_data)
    print(f"\nRader: {len(rader_med)}")
    for r in rader_med:
        post = r["Post"]
        flag = "[OK]" if post in omega_poster else "[IKKE I OMEGA]"
        print(f"  {flag}  {post}")
else:
    print("[WARNING] Omega-fil ikke funnet")
