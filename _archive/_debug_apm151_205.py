# -*- coding: utf-8 -*-
import sys, os, glob, re, pdfplumber
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from ifi.pdf_maalebrev_til_excel import les_pdf_hoved

base = r'C:\Users\ifi\Scan Survey\6900 - Fornebubanen'
pdf205_mars = [p for p in glob.glob(base + r'\**\*APM205*.pdf', recursive=True)]

print("=== APM205 etter fix ===")
for sti in pdf205_mars:
    rader = les_pdf_hoved(sti)
    print(f"  {len(rader)} rader: {os.path.basename(sti)}")
    for r in rader:
        post = r['Post']
        mengde = r['Mengde fra APM']
        enhet = r['Enhet fra APM']
        komm = r['Kommentar']
        print(f"  {post:42}  {str(mengde):10}  {enhet:4}  | {komm}")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from ifi.pdf_maalebrev_til_excel import les_pdf_hoved

base = r'C:\Users\ifi\Scan Survey\6900 - Fornebubanen'
pdf205_mars = [p for p in glob.glob(base + r'\**\*APM205*.pdf', recursive=True)]

print("=== APM205 etter fix ===")
for sti in pdf205_mars:
    rader = les_pdf_hoved(sti)
    print(f"  {len(rader)} rader: {os.path.basename(sti)}")
    for r in rader:
        post = r['Post']
        mengde = r['Mengde fra APM']
        enhet = r['Enhet fra APM']
        komm = r['Kommentar']
        print(f"  {post:42}  {str(mengde):10}  {enhet:4}  | {komm}")
