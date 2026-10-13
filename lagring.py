# -*- coding: utf-8 -*-
"""Trygg lagring av Excel-filer som kan vere opne i Excel.

Når ei fil er open i Excel held Windows ein skrivelås, og openpyxl/pandas
kastar PermissionError. I staden for å krasje med ein feilmelding ventar vi
og ber brukaren lukke fila — så held køyringa fram når fila er fri.

Bruk:
    from lagring import excel_mappe, vent_til_skrivbar
    utfil = os.path.join(excel_mappe(), "resultat.xlsx")   # <kodemappa>/excel/
    vent_til_skrivbar(utfil)        # rett før wb.save(utfil) / ExcelWriter(utfil)
    wb.save(utfil)
"""

from __future__ import annotations

import os
import time

# Excel-filer som vert lagra i kodemappa, hamnar i undermappa "excel".
# Ho står i .gitignore, så resultatfilene vert ikkje pusha til GitHub.
EXCEL_MAPPE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel")


def excel_mappe() -> str:
    """Stien til excel-mappa. Mappa vert oppretta om ho ikkje finst."""
    os.makedirs(EXCEL_MAPPE, exist_ok=True)
    return EXCEL_MAPPE


def _er_laast(path: str) -> bool:
    """True dersom fila finst og ikkje kan opnast for skriving (låst)."""
    if not os.path.exists(path):
        return False
    try:
        # Append-binær endrar ikkje innhald, men feilar om fila er låst av Excel.
        with open(path, "a+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        # Andre I/O-feil — la den faktiske skrivinga handtere det.
        return False


def vent_til_skrivbar(
    path: str,
    *,
    ventetid: float = 3.0,
    maks_sekund: float = 600.0,
) -> bool:
    """Vent til `path` kan skrivast til (ikkje open i Excel).

    Skriv ei tydeleg melding kvar runde slik at brukaren veit kva fil som må
    lukkast. Returnerer True når fila er skrivbar, eller False dersom
    `maks_sekund` er passert (då held kallaren fram og lar ei evt. feil sprette).
    Nye filer (som ikkje finst enno) returnerer True med ein gong.
    """
    if not _er_laast(path):
        return True

    namn = os.path.basename(path)
    starttid = time.time()
    forsok = 0
    while _er_laast(path):
        forsok += 1
        if time.time() - starttid > maks_sekund:
            print(f"[ADVARSEL] '{namn}' er framleis open etter "
                  f"{maks_sekund:.0f}s — prøver å lagre likevel.")
            return False
        print(f"[VENTAR] Kan ikkje lagre '{namn}' — fila er open (truleg i Excel). "
              f"Lukk ho, så held køyringa fram automatisk. "
              f"Prøver igjen om {ventetid:.0f}s (forsøk {forsok})…")
        time.sleep(ventetid)

    print(f"[OK] '{namn}' er no skrivbar — held fram.")
    return True
