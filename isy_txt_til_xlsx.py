# -*- coding: utf-8 -*-
"""Konverterer ISY Linker-eksport (.txt, tab-separert, utan header) til .xlsx
med dei kolonnenamna mengde_oversikt.les_isy ventar.

Dei seks første felta i .txt-en svarar 1:1 til kolonnene i den eksisterande
"NS3420 K10 … .xlsx" frå 22.06 — kontrollert felt for felt.
"""
import io, os, sys
import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

KOL = ["Post nummer", "Kode", "All tekst", "Enhet", "Mengde", "Pris: Anbud"]

def konverter(txt_sti: str, xlsx_sti: str) -> int:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tabell1"
    ws.append(KOL)
    n = 0
    with io.open(txt_sti, encoding="cp1252", newline="") as f:
        for linje in f:
            felt = linje.rstrip("\r\n").split("\t")
            if not felt or not felt[0].strip():
                continue
            rad = (felt + [""] * 6)[:6]
            # Mengde og pris som tal der det går
            ut = []
            for i, v in enumerate(rad):
                v = ILLEGAL_CHARACTERS_RE.sub("", v.replace("_x000D_", "")).strip()
                if i in (4, 5) and v:
                    try:
                        ut.append(float(v.replace(",", ".")))
                        continue
                    except ValueError:
                        pass
                ut.append(v[:32000])
            ws.append(ut)
            n += 1
    wb.save(xlsx_sti)
    return n

if __name__ == "__main__":
    MAPPE = (r"C:\Users\ifi\Scan Survey\6900 - Fornebubanen"
             r"\6900-01-010 K7 - Lysaker og Vækerø Stasjon\Målebrev\SKY - Maalebrev")
    for navn in ["NS3420 K10 Skøyen Stasjon  Madserud og Bestum",
                 "SVV Prosesskode K10 Skøyen Stasjon  Madserud og Bestum"]:
        txt = os.path.join(MAPPE, navn + ".txt")
        xlsx = os.path.join(MAPPE, navn + ".xlsx")
        n = konverter(txt, xlsx)
        print(f"{n:>7,} rader -> {os.path.basename(xlsx)}")
