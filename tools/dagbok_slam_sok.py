"""
dagbok_slam_sok.py
------------------
Soker gjennom alle dagbok-PDF-er i G:\\Min disk\\Dagboker og finner
omtaler av slam, bunnrensk, prelltapp og renseanlegg.
Lager en Excel-oversikt med ett ark per sokeord.
"""
import os
import re
import io
import glob
import random
import threading
import tkinter as tk
from tkinter import messagebox
from pdfminer.high_level import extract_text
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

DAGBOK_MAPPE = r"G:\Min disk\Dagbøker"
UT_FIL = os.path.join(DAGBOK_MAPPE, "Dagbok_soek_oversikt.xlsx")

# --- Sokeord-definisjoner ---
# Hvert element: (arknavn, sokeord_label, regex-pattern)
SOKEORD = [
    ("Slam", "slam", re.compile(r"slam", re.IGNORECASE)),
    ("Bunnrensk", "bunnrensk", re.compile(r"bunnrensk", re.IGNORECASE)),
    ("Prelltapp", "prelltapp", re.compile(
        r"prel+[ei]?t+a?p+", re.IGNORECASE)),
    ("Renseanlegg", "renseanlegg", re.compile(
        r"rense\s*an+l?e?g+", re.IGNORECASE)),
    ("Basseng", "basseng", re.compile(
        r"bas+[ea]n?g", re.IGNORECASE)),
    ("Toemming", "toemming", re.compile(
        r"t[øo]m+[ei]ng|t[øo]m+e[sr]|t[øo]m+t", re.IGNORECASE)),
]

# Regex for dato oevrst paa dagboksidene (format "16.01.23" eller "16.01.2023")
# Dagbok-sidene har datoer som "DD.MM.YY" eller lengre datostrenger
dato_re = re.compile(r"\b(\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2}))\b")

# Mer fullstendig dato-linje: "mandag 16. januar 2023 16:51"
lang_dato_re = re.compile(
    r"(?:mandag|tirsdag|onsdag|torsdag|fredag|lørdag|søndag)\s+"
    r"(\d{1,2})\.\s*(januar|februar|mars|april|mai|juni|juli|august|september|oktober|november|desember)\s+(\d{4})",
    re.IGNORECASE,
)

MAANEDER = {
    "januar": 1, "februar": 2, "mars": 3, "april": 4,
    "mai": 5, "juni": 6, "juli": 7, "august": 8,
    "september": 9, "oktober": 10, "november": 11, "desember": 12,
}


def parse_dato(tekst):
    """Proever aa finne dato fra dagbokside-tekst."""
    # Prov lang dato foerst
    m = lang_dato_re.search(tekst)
    if m:
        dag = int(m.group(1))
        mnd = MAANEDER.get(m.group(2).lower(), 0)
        aar = int(m.group(3))
        if mnd:
            try:
                return datetime(aar, mnd, dag).date()
            except ValueError:
                pass

    # Prov kort dato (DD.MM.YY eller DD.MM.YYYY) - ta foerste treff
    datoer = dato_re.findall(tekst[:200])  # Dato er typisk i starten
    for d in datoer:
        for fmt in ("%d.%m.%y", "%d.%m.%Y"):
            try:
                return datetime.strptime(d, fmt).date()
            except ValueError:
                continue
    return None


def finn_treff_avsnitt(tekst, pattern):
    """Finner alle avsnitt (paragraphs) som matcher pattern (regex).

    Et avsnitt er en blokk med tekst adskilt av tomme linjer.
    Returnerer hele avsnittet slik at flerlinjes beskrivelser fanges opp.
    """
    # Del opp i avsnitt: en eller flere tomme linjer skiller avsnitt
    avsnitt_liste = re.split(r"\n\s*\n", tekst)
    treff = []
    for avsnitt in avsnitt_liste:
        if pattern.search(avsnitt):
            # Rens: erstatt enkle linjeskift med mellomrom, trim
            renset = re.sub(r"\s*\n\s*", " ", avsnitt).strip()
            if renset:
                treff.append(renset)
    return treff


def les_fil_med_timeout(sti, timeout=30):
    """Leser fil til bytes med timeout for aa handtere trege Google Drive-filer."""
    result = [None]
    error = [None]

    def _les():
        try:
            with open(sti, "rb") as f:
                result[0] = f.read()
        except Exception as e:
            error[0] = e

    t = threading.Thread(target=_les, daemon=True)
    t.start()
    t.join(timeout)
    if t.is_alive():
        return None  # Timeout
    if error[0]:
        raise error[0]
    return result[0]


def prosesser_pdf(pdf_sti):
    """Returnerer dict {sokeord_label: [(dato, sidenr, tekst, pdf_sti), ...]}."""
    resultat = {s[1]: [] for s in SOKEORD}
    try:
        data = les_fil_med_timeout(pdf_sti, timeout=30)
        if data is None:
            print(f"[TIMEOUT]", end=" ")
            return resultat
        all_tekst = extract_text(io.BytesIO(data))
        sider = all_tekst.split("\x0c")  # formfeed = sideskift

        siste_dato = None  # Arver dato fra forrige side om ny side mangler dato
        for sidenr, tekst in enumerate(sider, 1):
            side_dato = parse_dato(tekst)
            if side_dato is not None:
                siste_dato = side_dato
            bruk_dato = siste_dato

            for arknavn, label, pattern in SOKEORD:
                treff = finn_treff_avsnitt(tekst, pattern)
                for linje in treff:
                    resultat[label].append((bruk_dato, sidenr, linje, pdf_sti))
    except Exception as e:
        print(f"[ERROR] {e}")
    return resultat


def lag_excel(alle_funn_per_sokeord, ut_sti):
    """alle_funn_per_sokeord: dict {label: [(dato, sidenr, tekst, pdf_sti), ...]}"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Fjern standard-ark

    navy = PatternFill("solid", fgColor="1E3070")
    hfont = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    nfont = Font(name="Calibri", size=10)
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    mid = Alignment(horizontal="center", vertical="center", wrap_text=False)
    lft = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_s = Side(style="thin", color="B0A8A0")
    brd = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)

    for arknavn, label, _pattern in SOKEORD:
        funn = alle_funn_per_sokeord.get(label, [])
        ws = wb.create_sheet(title=arknavn)

        headers = ["Dato", "PDF-fil", "Side", f"Tekst med {label}"]
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = hfont
            cell.fill = navy
            cell.alignment = mid
            cell.border = brd

        sortert = sorted(funn, key=lambda x: (x[0] or datetime(1900, 1, 1).date(), x[3], x[1]))

        for idx, (dato, sidenr, tekst, pdf_sti) in enumerate(sortert):
            r = idx + 2

            cell_dato = ws.cell(row=r, column=1, value=dato)
            cell_dato.number_format = "DD.MM.YYYY"
            cell_dato.font = nfont
            cell_dato.alignment = mid
            cell_dato.border = brd

            cell_fil = ws.cell(row=r, column=2, value=pdf_sti)
            cell_fil.font = link_font
            cell_fil.alignment = lft
            cell_fil.border = brd
            cell_fil.hyperlink = pdf_sti

            cell_side = ws.cell(row=r, column=3, value=sidenr)
            cell_side.font = nfont
            cell_side.alignment = mid
            cell_side.border = brd

            cell_tekst = ws.cell(row=r, column=4, value=tekst)
            cell_tekst.font = nfont
            cell_tekst.alignment = lft
            cell_tekst.border = brd

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 80
        ws.column_dimensions["C"].width = 8
        ws.column_dimensions["D"].width = 100

        ws.freeze_panes = "A2"
        from openpyxl.utils import get_column_letter
        ws.auto_filter.ref = f"A1:D{len(sortert)+1}"

        print(f"  [{arknavn}] {len(sortert)} treff")

    wb.save(ut_sti)
    print(f"[SUCCESS] Lagret: {ut_sti}")


def spoer_testmodus():
    """Viser popup for aa velge testmodus (15%) eller fullt soek."""
    root = tk.Tk()
    root.withdraw()
    svar = messagebox.askyesno(
        "Dagbok slam-soek",
        "Vil du kjoere i testmodus?\n\n"
        "Ja = Kun 15% av filene (rask test)\n"
        "Nei = Alle filer (fullstendig soek)"
    )
    root.destroy()
    return svar


def main():
    # Finn PDF-filer kun i undermapper (ikke rot-mappen)
    pdfer = sorted(glob.glob(os.path.join(DAGBOK_MAPPE, "*", "*.pdf")))

    # Spoer om testmodus
    testmodus = spoer_testmodus()
    if testmodus:
        antall_test = max(1, int(len(pdfer) * 0.15))
        pdfer = sorted(random.sample(pdfer, antall_test))
        print(f"[INFO] TESTMODUS: Valgt {antall_test} av {len(pdfer) + (len(pdfer) * 85 // 15)} filer")

    print(f"[INFO] Behandler {len(pdfer)} PDF-filer")

    alle_funn = {s[1]: [] for s in SOKEORD}
    for i, pdf_sti in enumerate(pdfer):
        print(f"  [{i+1}/{len(pdfer)}] {os.path.basename(pdf_sti)}...", end=" ", flush=True)
        resultat = prosesser_pdf(pdf_sti)
        treff_str = ", ".join(
            f"{label}:{len(resultat[label])}"
            for _, label, _ in SOKEORD if resultat[label]
        )
        print(treff_str if treff_str else "0")
        for label in alle_funn:
            alle_funn[label].extend(resultat[label])

    totalt = sum(len(v) for v in alle_funn.values())
    if totalt > 0:
        lag_excel(alle_funn, UT_FIL)
    else:
        print("[INFO] Ingen treff funnet i dagbokene")


if __name__ == "__main__":
    main()
